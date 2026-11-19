# -*- coding: utf-8 -*-
"""
run_amazon_check.py  ·  亚马逊前台关键词「同款占比」批处理验证（自适应版）
================================================================================
作者：Doubao Agent（YDK）      版本：v3  2026-11-18

做什么
-----
读一份卖家精灵导出的关键词表（.xlsx，工作表名/列位都不固定），逐词到亚马逊前台搜索，
统计搜索结果**前 50 个位置**（自然位 + Sponsored 广告位全部计入）里，与基准产品图
**外观同款**的商品数量、广告位数量、同款占比，输出 CSV。

自适应能力
---------
1. 工作表：自动遍历全部 Sheet，找"含关键词表头 + 数据行最多"的那张（优先表名带 ASIN 的）；
2. 基准图：自动扫描工作簿同级目录里的 .jpg/.png/.jpeg/.webp，按文件名意图（基准/目标/
   target/product/主图/main/参考）排序锁定；扫不到时回退到"表名里的 ASIN 主图"在线抓取；
3. 站点：从表名前缀（US-/UK-/DE-…）自动识别，匹配标准邮编；也可 --site 强制指定；
4. 否定词：A 列单元格红底（FFFF0000 / FF0000 / 调色板红）或整行刷红，记为"否定词已跳过"。

用法
----
  python run_amazon_check.py --limit 5                 # 先跑前 5 个有效词（试点）
  python run_amazon_check.py                           # 全量
  python run_amazon_check.py --skip-scrape --judgments work/judgments.json   # 只重算
  python run_amazon_check.py --print-config            # 只看本次生效的配置

依赖：Python 3.10+、openpyxl、Pillow、ego-browser（本机已装）
"""

from __future__ import annotations

import argparse
import csv
import glob
import json
import os
import random
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

# --------------------------------------------------------------------------- #
# 0. 全球站点与标准邮编字典
# --------------------------------------------------------------------------- #
SITES = {
    "US": {"domain": "https://www.amazon.com",   "postcode": "90001",    "name": "美国"},
    "UK": {"domain": "https://www.amazon.co.uk", "postcode": "DN11 0GU", "name": "英国"},
    "DE": {"domain": "https://www.amazon.de",    "postcode": "10115",    "name": "德国"},
    "JP": {"domain": "https://www.amazon.co.jp", "postcode": "001-0024", "name": "日本"},
    "CA": {"domain": "https://www.amazon.ca",    "postcode": "M4Y 1M7",  "name": "加拿大"},
    "AU": {"domain": "https://www.amazon.com.au","postcode": "3175",     "name": "澳大利亚"},
    "FR": {"domain": "https://www.amazon.fr",    "postcode": "75020",    "name": "法国"},
    "ES": {"domain": "https://www.amazon.es",    "postcode": "35600",    "name": "西班牙"},
    "IT": {"domain": "https://www.amazon.it",    "postcode": "29015",    "name": "意大利"},
    "SG": {"domain": "https://www.amazon.sg",    "postcode": "608829",   "name": "新加坡"},
    "SE": {"domain": "https://www.amazon.se",    "postcode": "10272",    "name": "瑞典"},
    "PL": {"domain": "https://www.amazon.pl",    "postcode": "01-541",   "name": "波兰"},
    "TR": {"domain": "https://www.amazon.com.tr","postcode": "06000",    "name": "土耳其"},
    "IE": {"domain": "https://www.amazon.co.uk", "postcode": "D24 E0CX", "name": "爱尔兰"},
    "MX": {"domain": "https://www.amazon.com.mx","postcode": "11000",    "name": "墨西哥"},
}

# 站点别名（表名前缀 / 用户输入 / 域名）→ 站点键
SITE_ALIAS = {
    "US": "US", "USA": "US", "AMZUS": "US", "AMERICAN": "US",
    "UK": "UK", "GB": "UK", "GBR": "UK",
    "DE": "DE", "GER": "DE", "GERMANY": "DE",
    "JP": "JP", "JPN": "JP", "JAPAN": "JP",
    "CA": "CA", "CAN": "CA",
    "AU": "AU", "AUS": "AU",
    "FR": "FR", "FRA": "FR",
    "ES": "ES", "ESP": "ES",
    "IT": "IT", "ITA": "IT",
    "SG": "SG", "SGP": "SG",
    "SE": "SE", "SWE": "SE",
    "PL": "PL", "POL": "PL",
    "TR": "TR", "TUR": "TR",
    "IE": "IE", "IRL": "IE",
    "MX": "MX", "MEX": "MX",
}
for _k, _v in SITES.items():
    SITE_ALIAS.setdefault(_v["domain"].split("//")[-1].replace("www.", "").split(".")[0].upper(), _k)

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")

CSV_HEADER = ["关键词", "翻译", "验证状态", "广告位数(前50名)",
              "前50名同款数量(含广告)", "同款占比 (%)"]

# 同款判定口径：默认只放跨品类通用原则，具体边界由用户按产品逐次确认后写入 _meta.rules
JUDGE_RULES = [
    "按【品类 + 可见外形结构】判定，不看品牌、颜色、尺码、logo、包装。",
    "同一型号的多件装算同款：一个位次只计 1，不因件数放大。",
    "不同型号/不同结构混装不算同款：混装里出现非本款产品即整组不算。",
    "配件、耗材、赠品、周边不算同款。",
    "本产品具体边界（哪些近似品类算、哪些不算）以用户逐次确认的口径为准，写进 _meta.rules 后执行。",
]

IMG_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp")
NAME_HINTS = ("基准", "目标", "target", "product", "主图", "main", "参考", "同款", "sample")
SKIP_DIRS = {"work", "pillow_run", "前台验证图片", "__pycache__", "node_modules"}


# --------------------------------------------------------------------------- #
# 1. 工具函数
# --------------------------------------------------------------------------- #
def log(msg: str = "") -> None:
    print(msg, flush=True)


def find_ego() -> str:
    """定位 ego-browser 可执行文件：环境变量 → PATH → ego 安装目录。"""
    env = os.environ.get("EGO_BROWSER", "").strip()
    if env and Path(env).exists():
        return env
    which = shutil.which("ego-browser")
    if which:
        return which
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    for cand in (home / ".local/share/ego/active_version_dir/ego-browser.exe",
                 Path(os.environ.get("LOCALAPPDATA", "")) / "ego/active_version_dir/ego-browser.exe"):
        if cand.exists():
            return str(cand)
    return ""


def red_fill_hex(cell) -> str | None:
    """返回单元格底色的可识别描述；非填充返回 None。"""
    try:
        f = cell.fill
        if f is None or f.patternType is None:
            return None
        fg = f.fgColor
        if fg is None:
            return None
        if getattr(fg, "type", None) == "rgb" and isinstance(fg.rgb, str):
            return fg.rgb.upper()
        if getattr(fg, "type", None) == "indexed" and fg.indexed is not None:
            return f"INDEXED:{fg.indexed}"
        if getattr(fg, "type", None) == "theme":
            return f"THEME:{fg.theme}"
    except Exception:
        return None


def is_red(cell) -> bool:
    h = red_fill_hex(cell)
    if not h:
        return False
    return ("FF0000" in h) or h in ("INDEXED:2", "INDEXED:10") or h.startswith("THEME:10")


def slugify(text: str) -> str:
    s = re.sub(r"[^0-9A-Za-z]+", "_", text).strip("_").lower()
    return (s or "kw")[:50]


def keyword_stem(keyword: str) -> str:
    import hashlib
    norm = re.sub(r"\s+", " ", keyword.strip().lower())
    return f"{slugify(keyword)}_{hashlib.md5(norm.encode('utf-8')).hexdigest()[:6]}"


# --------------------------------------------------------------------------- #
# 2. 工作簿自适应：定位主数据表 + 读 A/B 列 + 红底
# --------------------------------------------------------------------------- #
def detect_sheet(wb) -> tuple:
    """遍历所有 Sheet，返回 (最优 sheet, kw_col, tr_col, header_row, 诊断表)。"""
    report = []
    best = None
    for ws in wb.worksheets:
        kw_col = tr_col = None
        header_row = None
        for r in range(1, min(10, (ws.max_row or 0) + 1)):
            vals = [str(ws.cell(r, c).value or "").strip()
                    for c in range(1, min(ws.max_column or 1, 15) + 1)]
            hit = None
            for i, v in enumerate(vals):
                if v in ("关键词", "关键词翻译", "Keyword", "keyword", "词"):
                    hit = i
                    break
            if hit is not None:
                header_row = r
                kw_col = hit + 1
                for i, v in enumerate(vals):
                    if "翻译" in v or v.lower() in ("translation", "translate"):
                        tr_col = i + 1
                        break
                if tr_col is None and kw_col + 1 <= (ws.max_column or 1):
                    # 第 2 列不是翻译时也默认按"关键词右边一列"取值
                    tr_col = kw_col + 1
                break
        rows = 0
        if kw_col:
            for r in range(header_row + 1, (ws.max_row or 0) + 1):
                if str(ws.cell(r, kw_col).value or "").strip():
                    rows += 1
        has_asin_name = bool(re.search(r"B0[A-Z0-9]{8}", str(ws.title), re.I))
        score = rows + (200 if has_asin_name else 0)
        report.append((ws.title, kw_col, tr_col, rows, has_asin_name, score))
        if rows and (best is None or score > best[5]):
            best = (ws, kw_col, tr_col, header_row, score)

    log("  Sheet 扫描：")
    for t, k, tr, n, asin, sc in report:
        log(f"    - {t!r:34} 关键词列={k} 翻译列={tr} 数据行={n:<6} 表名含ASIN={asin}")
    if not best:
        raise SystemExit("× 没有找到任何含「关键词」表头的工作表，请检查文件。")
    return best[0], best[1], best[2], best[3]


def read_keywords(workbook: Path, sheet_name: str = "") -> tuple:
    wb = load_workbook(workbook, data_only=True)
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise SystemExit(f"× 工作簿里没有工作表 {sheet_name!r}，现有：{wb.sheetnames}")
        ws = wb[sheet_name]
        kw_col, tr_col, header_row = None, None, None
        for r in range(1, 10):
            for c in range(1, 15):
                if str(ws.cell(r, c).value or "").strip() == "关键词":
                    header_row, kw_col = r, c
                    tr_col = c + 1
                    break
            if header_row:
                break
    else:
        ws, kw_col, tr_col, header_row = detect_sheet(wb)

    rows = []
    red_count = 0
    for r in range(header_row + 1, ws.max_row + 1):
        kw = str(ws.cell(r, kw_col).value or "").strip()
        if not kw:
            continue
        tr = str(ws.cell(r, tr_col).value or "").strip() if tr_col else ""
        cell_red = is_red(ws.cell(r, kw_col))
        # 整行刷红：≥2 个红格视为整行否定（避免单列装饰红底误判整表）
        row_red = sum(1 for c in range(1, min(ws.max_column, 12) + 1) if is_red(ws.cell(r, c))) >= 2
        red = cell_red or row_red
        if red:
            red_count += 1
        rows.append({"row": r, "keyword": kw, "translation": tr, "red": red})

    info = {"sheet": ws.title, "header_row": header_row, "kw_col": kw_col,
            "tr_col": tr_col, "total": len(rows), "red": red_count,
            "valid": len(rows) - red_count}
    log(f"  主表: {ws.title!r}  表头行={header_row}  关键词列={kw_col}  翻译列={tr_col}")
    log(f"  关键词 {info['total']} 个，其中红底否定词 {red_count} 个，有效 {info['valid']} 个")
    return rows, info


def sniff_site(sheet_name: str):
    """从表名/文件名识别站点，如 'US-B09WXSPDQ9(20)__' → US、'UK-B0...' → UK。

    识别不出来时返回 None，绝不默认美国——站点决定域名与邮编，猜错整套数据作废。
    """
    for token in re.split(r"[^A-Za-z]+", sheet_name or ""):
        key = SITE_ALIAS.get(token.upper())
        if key:
            return key
    for key, cfg in SITES.items():
        if cfg["domain"].split("www.")[-1] in (sheet_name or ""):
            return key
    return None


# --------------------------------------------------------------------------- #
# 3. 基准图自适应：目录扫描 → ASIN 主图兜底
# --------------------------------------------------------------------------- #
def scan_reference_image(base_dir: Path) -> Path | None:
    cands = []
    for p in base_dir.iterdir():
        if p.is_file() and p.suffix.lower() in IMG_EXT and p.stat().st_size > 20_000:
            cands.append(p)
    if not cands:
        for pat in ("*" + e for e in IMG_EXT):
            for s in glob.glob(str(base_dir / "**" / pat), recursive=True):
                p = Path(s)
                if any(part in SKIP_DIRS for part in p.parts):
                    continue
                if p.is_file() and p.stat().st_size > 20_000:
                    cands.append(p)
    if not cands:
        return None

    def score(p: Path):
        name = p.name.lower()
        hint = 1 if any(h.lower() in name for h in NAME_HINTS) else 0
        return (hint, p.stat().st_size)

    cands.sort(key=score, reverse=True)
    log("  目录内候选图片：")
    for p in cands[:6]:
        log(f"    - {p.name}  ({p.stat().st_size/1024:.0f} KB)")
    return cands[0]


def fetch_asin_main_image(asin: str, ego: str, work_dir: Path, domain: str) -> Path | None:
    """回退方案：用 ego-browser 打开 ASIN 详情页，抓主图并落到本地。"""
    if not (ego and asin):
        return None
    out_json = work_dir / "browser" / "target.json"
    out_img = work_dir / "browser" / f"target_{asin}.jpg"
    out_json.parent.mkdir(parents=True, exist_ok=True)
    js = f"""
const fs = await import("node:fs/promises");
const OUT = {json.dumps(str(out_json).replace(chr(92), '/'))};
const task = (await listTaskSpaces()).find(s => s.name === "amazon-check-{asin}");
const space = task ? await taskSpace(task.id) : await taskSpace("amazon-check-{asin}");
const page = space.page("p1");
await page.goto({json.dumps(domain)} + "/dp/{asin}", {{ timeout: 60000 }});
await page.waitForLoadState("domcontentloaded");
await page.waitForTimeout(1500);
const info = await page.evaluate(() => {{
  const img = document.querySelector("#landingImage") ||
              document.querySelector("#imgTagWrapperId img");
  const t = document.querySelector("#productTitle");
  return {{ src: img ? (img.getAttribute("data-old-hires") || img.getAttribute("src") || "") : "",
            title: t ? t.innerText.replace(/\\s+/g, " ").trim() : "" }};
}});
await fs.writeFile(OUT, JSON.stringify(info, null, 2), "utf8");
console.log(info.src);
"""
    mjs = work_dir / "browser" / f"fetch_{asin}.mjs"
    mjs.write_text(js, encoding="utf-8")
    subprocess.run([ego, "nodejs", "-e", f"await import('file:///{mjs.as_posix()}')"],
                   capture_output=True, timeout=180)
    if not out_json.exists():
        return None
    info = json.loads(out_json.read_text(encoding="utf-8"))
    url = info.get("src", "")
    if not url:
        return None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
        out_img.write_bytes(raw)
        log(f"  ASIN 主图已抓取：{info.get('title','')[:70]}")
        return out_img
    except Exception as e:
        log(f"  ASIN 主图下载失败：{e}")
        return None


# --------------------------------------------------------------------------- #
# 4. ego-browser：设配送地址 + 批量抓取前 50 个位置
# --------------------------------------------------------------------------- #
SCRAPE_JS = r"""
const fs = await import("node:fs/promises");
const CFG = __CFG__;

async function openTask() {
  const name = CFG.space;
  // 接管已有空间的顺序很重要：agent 空间用 taskSpace(id) 拿不到标签页，必须 claim
  async function claim(id) {
    try { const t = await claimTaskSpace(id); const tabs = await t.tabs(); if (tabs && tabs.length) return t; } catch (e) {}
    try { const t = await takeOverTaskSpace(id); const tabs = await t.tabs(); if (tabs && tabs.length) return t; } catch (e) {}
    try { const t = await taskSpace(id); const tabs = await t.tabs(); if (tabs && tabs.length) return t; } catch (e) {}
    try { return await claimTaskSpace(id); } catch (e) {}
    try { return await takeOverTaskSpace(id); } catch (e) {}
    try { return await taskSpace(id); } catch (e) {}
    return null;
  }
  try {
    const spaces = await listTaskSpaces();
    // 名字可能在被接管后变空，按“已有空间里第一个可用的”兜底查找
    const found = spaces.find((s) => s.name === name) ||
                  spaces.find((s) => s.name && s.name.startsWith(name.slice(0, 16))) ||
                  spaces.find((s) => !s.name);
    if (found) {
      const t = await claim(found.id);
      if (t) return t;
    }
  } catch (e) {}
  try { return await taskSpace(name, CFG.profileId ? { profileId: CFG.profileId } : undefined); }
  catch (e) { return await taskSpace(name + " " + Date.now()); }
}

const task = await openTask();
// 取页面：空间里的标签页已被跟踪，直接用 tabs[i].page（adopt 只接受未跟踪页面，会抛错）；
// 空间有页时绝不 newPage——标签页预算可能已满（Page budget reached）会直接失败。
async function ensurePage() {
  for (let i = 0; i < 3; i++) {
    let tabs = [];
    try { tabs = await task.tabs(); } catch (e) {}
    if (tabs && tabs.length) {
      const hit = tabs.find((x) => x.page);
      if (hit && hit.page) return hit.page;
    }
    try { const np = await task.newPage(); if (np) return np; } catch (e) {}
    await new Promise((r) => setTimeout(r, 1200));
  }
  throw new Error("ego 空间未就绪：没有可用标签页，且新建失败（可能标签页预算已满）");
}
const page = await ensurePage();
const OUT = { location: { ok: false, text: "", attempts: [] }, results: [], blocked: false };

async function visible(sel) {
  return await page.evaluate((s) => {
    const e = document.querySelector(s);
    if (!e) return false;
    const r = e.getBoundingClientRect();
    const cs = getComputedStyle(e);
    return r.width > 0 && r.height > 0 && cs.display !== "none" && cs.visibility !== "hidden";
  }, sel);
}

const norm = (x) => (x || "").toUpperCase().replace(/[^A-Z0-9]/g, "");

async function setLocation() {
  const attempt = { postcode: CFG.postcode, ok: false, note: "" };
  try {
    await page.goto(CFG.marketplace, { timeout: 60000 });
    await page.waitForLoadState("domcontentloaded");
    await page.waitForTimeout(800);
    // 地址已生效就直接复用，避免每次重复点击弹窗（也降低被 ego 判定为用户接管的风险）
    const pre = await page.evaluate(() => {
      const e = document.querySelector("#glow-ingress-line2");
      return e ? e.textContent.replace(/\s+/g, " ").trim() : "";
    });
    const pcPre = norm(CFG.postcode);
    if (pcPre && norm(pre).includes(pcPre.slice(0, 5))) {
      OUT.location.text = pre;
      attempt.ok = true;
      attempt.note = pre + " (already set)";
      OUT.location.attempts.push(attempt);
      OUT.location.ok = true;
      return true;
    }
    const openers = ["#nav-global-location-popover-link", "#glow-ingress-block", "#nav-global-location-slot"];
    let opened = false;
    for (const s of openers) {
      if (await visible(s)) { await page.click(s); opened = true; break; }
    }
    if (!opened) throw new Error("location opener not found");
    const inputs = ["#GLUXZipUpdateInput", "input[data-action='GLUXPostalInputAction']", "input[name='glowDoneButton']"];
    let input = "";
    for (const s of inputs) if (await visible(s)) { input = s; break; }
    if (!input) { await page.waitForTimeout(900); for (const s of inputs) if (await visible(s)) { input = s; break; } }
    if (!input) throw new Error("postcode input not found");
    await page.fill(input, CFG.postcode);
    const submits = ["#GLUXZipUpdate", "#GLUXZipUpdate input", "input[aria-labelledby='GLUXZipUpdate-announce']", "button[name='glowDoneButton']"];
    let submitted = false;
    for (const s of submits) if (await visible(s)) { await page.click(s); submitted = true; break; }
    if (!submitted) throw new Error("postcode submit not found");
    await page.waitForTimeout(1400);
    for (const s of ["#GLUXConfirmClose", "button[name='glowDoneButton']"]) {
      if (await visible(s)) { try { await page.click(s); } catch {} }
    }
    await page.reload({ timeout: 60000 });
    await page.waitForLoadState("domcontentloaded");
    await page.waitForTimeout(600);
    const txt = await page.evaluate(() => {
      const e = document.querySelector("#glow-ingress-line2") || document.querySelector("#glow-ingress-line1");
      return (e ? e.textContent : "").replace(/\s+/g, " ").trim();
    });
    OUT.location.text = txt;
    const pc = norm(CFG.postcode);
    const seen = norm(txt);
    attempt.ok = pc.length >= 3 && (seen.includes(pc) || seen.includes(pc.slice(0, 5)));
    attempt.note = txt;
  } catch (e) {
    attempt.note = String((e && e.message) || e);
  }
  OUT.location.attempts.push(attempt);
  OUT.location.ok = attempt.ok;
  return attempt.ok;
}

async function harvest() {
  return await page.evaluate(() => {
    const bodyText = document.body ? document.body.innerText : "";
    const href = location.href || "";
    const checkpoint = /\/checkpoint\/|\/errors\/validateCaptcha|\/ap\/signin/i.test(href);
    const captcha =
      checkpoint ||
      /unusual browser activity|confirm you.?re human|Enter the characters you see below|Robot Check|Sorry, we just need to make sure/i.test(bodyText) ||
      /Robot Check|Sorry! Something went wrong/i.test(document.title || "");

    function inAdWidget(el) {
      let p = el, d = 0;
      while (p && d <= 8) {
        const cls = (p.className || "").toString();
        if (/(a-carousel|a-carousel-card|a-carousel-viewport|_c2Itd_)/.test(cls)) return true;
        if (p.getAttribute && p.getAttribute("data-a-carousel-type")) return true;
        p = p.parentElement; d++;
      }
      return false;
    }

    // 品牌广告条（brand shelf）：整条带 Sponsored 头 + 横排多个商品，条内每个商品都算广告位
    function adShelf(el) {
      let p = el, d = 0;
      while (p && d <= 6) {
        const dtype = p.getAttribute && p.getAttribute("data-component-type");
        if (dtype === "s-search-result") return false;
        const n = p.querySelectorAll ? p.querySelectorAll("[data-asin]").length : 0;
        const t = (p.innerText || "").replace(/\s+/g, " ").slice(0, 200);
        const al = (p.getAttribute && p.getAttribute("aria-label")) || "";
        if (n >= 2 && n <= 15 && (/Sponsored/i.test(t) || /Sponsor/i.test(al) ||
            !!p.querySelector('[aria-label*="Sponsored" i]'))) return true;
        p = p.parentElement; d++;
      }
      return false;
    }

    // 广告货架自带标记：<span class="_c2Itd_ad-feedback-text-desktop" aria-label="Leave feedback on Sponsored ad">Sponsored</span>
    // 以及品牌条头部的 a[aria-label^="Sponsored ad from"]；瓦片本身没有 aria 时靠祖先命中
    const SHELF_MARK = '._c2Itd_ad-feedback-text-desktop, ._c2Itd_ctaSponsoredContainer, [aria-label="Leave feedback on Sponsored ad"], [aria-label*="Sponsored ad from" i]';
    function inAdShelfMark(el) {
      let p = el, d = 0;
      while (p && d <= 8) {
        const dtype = p.getAttribute && p.getAttribute("data-component-type");
        if (dtype === "s-search-result") return false;
        if (p.querySelector && p.querySelector(SHELF_MARK)) return true;
        p = p.parentElement; d++;
      }
      return false;
    }

    // 轮播货架里被 overflow 裁掉的隐藏瓦片不算展示位：只计与 .a-carousel-viewport 横向重叠 >50% 的瓦片
    function inCarousel(el) {
      let p = el, d = 0;
      while (p && d <= 10) { if (/a-carousel/.test(String(p.className || ""))) return true; p = p.parentElement; d++; }
      return false;
    }
    function carouselVisible(el) {
      if (!inCarousel(el)) return true;
      const vp = el.closest ? el.closest(".a-carousel-viewport") : null;
      if (!vp) return true;
      const a = el.getBoundingClientRect(), b = vp.getBoundingClientRect();
      const ov = Math.min(a.right, b.right) - Math.max(a.left, b.left);
      return ov > Math.min(a.width, b.width) * 0.5;
    }

    const nodes = [...document.querySelectorAll("[data-asin]")];
    const items = [], pushed = [];
    for (const card of nodes) {
      if (pushed.some((el) => el !== card && el.contains(card))) continue;
      const asin = (card.getAttribute("data-asin") || "").trim();
      if (asin.length < 10) continue;
      if (!card.getClientRects || card.getClientRects().length === 0) continue;
      if (!carouselVisible(card)) continue;
      const cs = window.getComputedStyle ? window.getComputedStyle(card) : null;
      if (cs && (cs.display === "none" || cs.visibility === "hidden")) continue;
      const imgEl = card.querySelector("img.s-image") || card.querySelector("img");
      if (!imgEl) continue;
      const label = card.querySelector('[data-component-type="s-sponsored-label-info"], .puis-sponsored-label-text, .s-sponsored-label-info');
      // 主判据：亚马逊广告信息图标 aria-label="View Sponsored information or leave ad feedback"
      const ariaSp = !!card.querySelector('[aria-label*="Sponsored" i]') ||
                     !!(card.getAttribute && /Sponsor/i.test(card.getAttribute("aria-label") || ""));
      const text = card.innerText || "";
      const flat = text.replace(/\s+/g, " ").trim();
      const sponsored = ariaSp || !!label || /(^|\n)\s*Sponsored\s*(\n|$)/i.test(text) || /^Sponsored\b/i.test(flat)
        || inAdWidget(card) || adShelf(card) || inAdShelfMark(card);
      const titleEl = card.querySelector('[data-cy="title-recipe"] h2 span') || card.querySelector("h2 span") ||
                      card.querySelector("h2") || card.querySelector("img[alt]");
      // 品牌广告条的横排商品常无 h2，回退到图片 alt / 卡片首行文字
      const title = (titleEl ? (titleEl.innerText || titleEl.getAttribute("alt") || "") : flat)
        .replace(/\s+/g, " ").trim().slice(0, 300);
      if (!title) continue;
      items.push({ asin, title, img: imgEl.getAttribute("src") || "", sponsored, slot: sponsored ? "ad" : "organic" });
      pushed.push(card);
    }
    return { captcha, checkpoint, items };
  });
}

async function scrollPage() {
  for (let i = 0; i < 4; i++) {
    await page.evaluate(() => window.scrollBy(0, 1400));
    await page.waitForTimeout(600);
  }
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.waitForTimeout(300);
}

await setLocation();
console.log(JSON.stringify({ stage: "location", ok: OUT.location.ok, text: OUT.location.text }));

async function searchByBox(kw) {
  // 模拟真实用户：在顶部搜索框里清空 → 逐字键入 → 回车
  const box = "#twotabsearchtextbox";
  await page.waitForSelector(box, { state: "visible", timeout: 15000 });
  try {
    await page.click(box);
    await page.press(box, "ControlOrMeta+A");
    await page.keyboard.type(kw, { delay: 80 });
    await page.press(box, "Enter");
  } catch (e) {
    await page.fill(box, kw);
    await page.press(box, "Enter");
  }
  await page.waitForURL(/\/s\?/, { timeout: 45000 });
  await page.waitForTimeout(1200);
}

let consecutiveCaptcha = 0;
for (const job of CFG.jobs) {
  const res = { keyword: job.kw, row: job.row, status: "ok", pages: 0, items: [], notes: [] };
  const base = CFG.marketplace + "/s?k=" + encodeURIComponent(job.kw);
  try {
    for (let p = 1; p <= CFG.maxPages && res.items.length < CFG.topN; p++) {
      try {
        if (p === 1) {
          await searchByBox(job.kw);
        } else {
          let clicked = false;
          try { await page.click("a.s-pagination-next"); clicked = true; } catch (e) {}
          if (!clicked) await page.goto(base + "&page=" + p, { timeout: 60000 });
          await page.waitForTimeout(1300);
        }
      } catch (e) {
        res.notes.push("nav fail p" + p + ": " + String((e && e.message) || e).slice(0, 160));
        try { await page.goto(base + (p === 1 ? "" : "&page=" + p), { timeout: 60000 }); } catch (e2) {}
        await page.waitForTimeout(2000);
      }
      try {
        await page.waitForSelector('div[data-component-type="s-search-result"], div[data-asin]',
                                    { state: "visible", timeout: 15000 });
      } catch (e) {}
      await scrollPage();
      const data = await harvest();
      res.pages = p;
      if (data.captcha) {
        res.status = "captcha"; res.notes.push("captcha p" + p);
        if (data.checkpoint) { OUT.blocked = true; res.notes.push("amazon checkpoint"); }
        break;
      }
      for (const it of data.items) {
        res.items.push(it);
        if (res.items.length >= CFG.topN) break;
      }
      if (res.items.length >= CFG.topN) break;
      await page.waitForTimeout(900);
    }
  } catch (e) {
    res.status = res.status === "ok" ? "error" : res.status;
    res.notes.push("loop: " + String((e && e.message) || e).slice(0, 200));
    try { await page.waitForTimeout(3000); } catch (e2) {}
  }
  res.items = res.items.slice(0, CFG.topN).map((it, i) => ({ ...it, rank: i + 1 }));
  if (res.status === "ok" && res.items.length === 0) res.status = "empty";
  await fs.mkdir(job.out.substring(0, job.out.lastIndexOf("/")), { recursive: true });
  await fs.writeFile(job.out, JSON.stringify(res, null, 1), "utf8");
  OUT.results.push({ keyword: job.kw, status: res.status, n: res.items.length, pages: res.pages });
  console.log(JSON.stringify({ keyword: job.kw, status: res.status, n: res.items.length, pages: res.pages }));
  // 命中 checkpoint：立即中止本批，避免持续硬撞风控（剩余词下次再跑）
  if (OUT.blocked) break;
  // 验证码自适应：连续触发时指数退避（最长 10 分钟），避免持续硬撞风控
  if (res.status === "captcha") {
    consecutiveCaptcha += 1;
    const waitMs = Math.min(600000, 45000 * Math.pow(2, consecutiveCaptcha - 1));
    console.log(JSON.stringify({ cooldown_ms: waitMs, consecutive_captcha: consecutiveCaptcha }));
    await page.waitForTimeout(waitMs);
  } else {
    consecutiveCaptcha = 0;
  }
    await page.waitForTimeout(CFG.delayMs + Math.floor(Math.random() * (CFG.jitterMs || 0)));
}
await fs.writeFile(CFG.summary, JSON.stringify(OUT, null, 2), "utf8");
"""


def run_scrape(jobs: list, cfg: dict, work_dir: Path, ego: str) -> dict:
    """把一批关键词交给一个 ego-browser 进程串行抓取（省掉每词重启的开销）。"""
    bdir = work_dir / "browser"
    bdir.mkdir(parents=True, exist_ok=True)
    payload = dict(cfg)
    payload["jobs"] = jobs
    mjs = bdir / f"batch_{int(time.time())}.mjs"
    js = SCRAPE_JS.replace("__CFG__", json.dumps(payload, ensure_ascii=False))
    mjs.write_text(js, encoding="utf-8")
    proc = subprocess.run([ego, "nodejs", "-e", f"await import('file:///{mjs.as_posix()}')"],
                          capture_output=True, text=True, encoding="utf-8", errors="replace",
                          timeout=cfg.get("timeout", 1200))
    out = proc.stdout or ""
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                log("    " + json.dumps(json.loads(line), ensure_ascii=False))
            except Exception:
                pass
    if proc.returncode != 0:
        log(f"    ego-browser 返回码 {proc.returncode}：{(proc.stderr or '')[-400:]}")
    summary_path = Path(payload["summary"])
    if summary_path.exists():
        return json.loads(summary_path.read_text(encoding="utf-8"))
    return {"location": {"ok": False, "text": "", "attempts": [{"note": (proc.stderr or '')[-300:]}]}, "results": []}


# --------------------------------------------------------------------------- #
# 5. 主图下载 + 复核拼图
# --------------------------------------------------------------------------- #
def upgrade_image_url(url: str) -> str:
    return re.sub(r"\._[A-Za-z0-9_,]+_\.(jpg|jpeg|png)", r"._AC_SL500_.\1", url, flags=re.I)


def download_image(url: str, dest: Path, timeout: int = 20) -> bool:
    if dest.exists() and dest.stat().st_size > 2000:
        return True
    for u in (upgrade_image_url(url), url):
        if not u:
            continue
        try:
            req = urllib.request.Request(u, headers={"User-Agent": UA, "Accept": "image/*,*/*"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read()
            im = Image.open(__import__("io").BytesIO(raw)).convert("RGB")
            im.thumbnail((600, 600))
            dest.parent.mkdir(parents=True, exist_ok=True)
            im.save(dest, "JPEG", quality=88)
            return True
        except Exception:
            continue
    return False


def download_all(items: list, kw_dir: Path, workers: int = 8) -> None:
    kw_dir.mkdir(parents=True, exist_ok=True)
    for it in items:
        it["image"] = str(kw_dir / f'{int(it["rank"]):02d}_{it["asin"]}{"_ad" if it["sponsored"] else ""}.jpg')

    def job(it):
        return it, download_image(it["img"], Path(it["image"]))

    with ThreadPoolExecutor(max_workers=workers) as ex:
        for fut in as_completed([ex.submit(job, it) for it in items]):
            it, ok = fut.result()
            it["image_ok"] = bool(ok)


def _font(size: int):
    for f in ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/segoeui.ttf"):
        try:
            return ImageFont.truetype(f, size)
        except Exception:
            continue
    return ImageFont.load_default()


def build_contact_sheets(keyword: str, items: list, out_dir: Path, per: int = 25, cols: int = 5) -> list:
    """把判断单元切成 5×5 编号拼图；每格带 #位次 ASIN [AD] 标签。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    cell_w, cell_h, bar = 230, 265, 26
    font = _font(14)
    sheets = []
    for si in range(0, len(items), per):
        chunk = items[si:si + per]
        rows = (len(chunk) + cols - 1) // cols
        canvas = Image.new("RGB", (cell_w * cols, (cell_h + bar) * rows + 8), "white")
        d = ImageDraw.Draw(canvas)
        for i, it in enumerate(chunk):
            r, c = divmod(i, cols)
            x, y = c * cell_w, r * (cell_h + bar)
            tag = f'#{it["rank"]:02d} {it["asin"]}' + (" [AD]" if it["sponsored"] else "")
            d.text((x + 4, y + 5), tag, fill=(180, 0, 0) if it["sponsored"] else (0, 0, 0), font=font)
            p = Path(it.get("image", ""))
            if it.get("image_ok") and p.exists():
                try:
                    im = Image.open(p).convert("RGB")
                    im.thumbnail((cell_w - 10, cell_h - 8))
                    canvas.paste(im, (x + (cell_w - im.width) // 2, y + bar + (cell_h - im.height) // 2))
                except Exception:
                    d.text((x + 10, y + bar + 60), "IMG ERR", fill="red", font=font)
            else:
                d.rectangle([x + 8, y + bar + 6, x + cell_w - 8, y + bar + cell_h - 8], outline=(200, 200, 200))
                d.text((x + 30, y + bar + 80), "DOWNLOAD FAILED", fill=(200, 0, 0), font=font)
        f = out_dir / f"{keyword_stem(keyword)}_sheet{si // per + 1}.jpg"
        canvas.save(f, "JPEG", quality=88)
        sheets.append(f)
    return sheets


# --------------------------------------------------------------------------- #
# 6. CSV 输出
# --------------------------------------------------------------------------- #
def write_csv(rows: list, path: Path) -> Path:
    try:
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(CSV_HEADER)
            for r in rows:
                w.writerow(r)
        return path
    except PermissionError:
        alt = path.with_suffix(".new.csv")
        with open(alt, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(CSV_HEADER)
            for r in rows:
                w.writerow(r)
        log(f"  ! 目标 CSV 被占用，已改写到 {alt}")
        return alt


def build_rows(rows: list, work: Path, judgments: dict, top_n: int) -> list:
    """把工作簿行 + 抓取结果 + 判定结果拼成 CSV 行（未抓取/待复核也如实写出）。"""
    out_rows = []
    for r in rows:
        kw = r["keyword"]
        base = [kw, r["translation"]]
        if r["red"]:
            out_rows.append(base + ["否定词已跳过", "", "", ""])
            continue
        sp = work / "scrape" / f"{keyword_stem(kw)}.json"
        if not sp.exists():
            out_rows.append(base + ["未抓取", "", "", ""])
            continue
        data = json.loads(sp.read_text(encoding="utf-8"))
        items = data.get("items", [])
        status = data.get("status", "ok")
        ads = sum(1 for it in items if it.get("sponsored"))
        sample = len(items)
        if status == "captcha":
            out_rows.append(base + ["风控未完成", ads, "", ""])
            continue
        if not items:
            out_rows.append(base + ["无搜索结果", 0, "", ""])
            continue
        entry = judgments.get(kw)
        if entry is None:
            out_rows.append(base + ["待复核", ads, "", ""])
            continue
        if entry.get("insufficient"):
            out_rows.append(base + ["图片数据不足", ads, "", ""])
            continue
        same = set(entry.get("same", []))
        cnt = sum(1 for it in items if it["rank"] in same)
        ratio = round(cnt / (sample or 1) * 100, 1)
        st = "已完成" if sample >= top_n else "已完成（样本不足）"
        out_rows.append(base + [st, ads, cnt, ratio])
    return out_rows


# --------------------------------------------------------------------------- #
# 7. 主流程
# --------------------------------------------------------------------------- #
def main() -> int:
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description="亚马逊前台关键词同款占比验证（自适应版）")
    ap.add_argument("--workbook", default="", help="卖家精灵导出 xlsx；默认自动在脚本目录找")
    ap.add_argument("--sheet", default="", help="强制指定工作表名")
    ap.add_argument("--target", nargs="+", default=None,
                    help="基准产品图，1-3 张（多张不同角度/不同包装）；默认自动扫描目录")
    ap.add_argument("--site", default="", help="站点，如 US/UK/DE…；默认从表名识别")
    ap.add_argument("--limit", type=int, default=0, help="只跑前 N 个有效词（0=全部）")
    ap.add_argument("--start", type=int, default=1, help="从第 N 个有效词开始（1 基）")
    ap.add_argument("--keyword", action="append", default=[], help="只跑指定关键词，可重复")
    ap.add_argument("--out", default=str(here / "amazon_analysis_result_top50.csv"))
    ap.add_argument("--work-dir", default=str(here / "work"))
    ap.add_argument("--judgments", default="", help="人工/模型判定 JSON；默认 work/judgments.json")
    ap.add_argument("--skip-scrape", action="store_true", help="复用已抓取的 work/scrape/*.json")
    ap.add_argument("--skip-location", action="store_true", help="跳过配送地址设置（调试用）")
    ap.add_argument("--no-images", action="store_true", help="不下图、不出拼图")
    ap.add_argument("--top-n", type=int, default=50)
    ap.add_argument("--pages", type=int, default=5)
    ap.add_argument("--delay", type=float, default=12.0, help="关键词间隔秒数")
    ap.add_argument("--batch", type=int, default=30, help="每个 ego 进程串行处理多少个词")
    ap.add_argument("--profile", default="", help="ego profileId（如 'Profile 1'），留空用 ego 默认配置")
    ap.add_argument("--print-config", action="store_true", help="只打印配置，不执行")
    args = ap.parse_args()

    work = Path(args.work_dir).resolve()
    work.mkdir(parents=True, exist_ok=True)
    (work / "scrape").mkdir(exist_ok=True)

    # --- 工作簿 ---
    wb_path = Path(args.workbook) if args.workbook else None
    if wb_path is None:
        cands = [p for p in here.glob("*.xlsx") if not p.name.startswith("~$")]
        pref = [p for p in cands if "卖家精灵" in p.name]
        wb_path = (pref or sorted(cands, key=lambda p: p.stat().st_size, reverse=True))[0] if cands else None
    if not wb_path or not wb_path.exists():
        raise SystemExit("× 找不到工作簿，请用 --workbook 指定")
    log(f"工作簿: {wb_path}")

    rows, info = read_keywords(wb_path, args.sheet)
    sniffed = sniff_site(info["sheet"] or wb_path.stem) or ""
    site_key = (args.site or sniffed or "").upper()
    site_key = SITE_ALIAS.get(site_key, site_key)
    if args.site and sniffed and SITE_ALIAS.get(sniffed.upper(), sniffed.upper()) != site_key:
        log(f"! 表名前缀指向 {sniffed}，但本次指定 {site_key}；按指定的 {site_key} 执行")
    if not site_key:
        raise SystemExit(
            "× 无法从表名识别站点，请用 --site 指定（可选："
            + ", ".join(f"{k}={v['name']}" for k, v in SITES.items())
            + f"）。当前表名：{info['sheet']}"
        )
    if site_key not in SITES:
        raise SystemExit(f"× 不支持的站点 {site_key}，可选：{', '.join(SITES)}")
    site = SITES[site_key]
    log(f"站点: {site_key} {site['name']}  {site['domain']}  邮编 {site['postcode']}")

    # --- 基准图（1-3 张，第一张为主参照） ---
    targets = [Path(t) for t in (args.target or [])]
    if not targets:
        auto = scan_reference_image(wb_path.parent)
        targets = [auto] if auto else []
    if len(targets) > 3:
        log(f"! 基准图给了 {len(targets)} 张，只取前 3 张")
        targets = targets[:3]
    for t in targets:
        if not t.exists():
            raise SystemExit(f"× 基准图不存在：{t}")
    target = targets[0] if targets else None
    if targets:
        for i, t in enumerate(targets, 1):
            log(f"基准图{i}: {t}  ({t.stat().st_size/1024:.0f} KB)")
    else:
        log("基准图: 目录内未找到，尝试用表名 ASIN 主图兜底")

    # --- 关键词筛选 ---
    valid = [r for r in rows if not r["red"]]
    if args.keyword:
        kset = {k.strip().lower() for k in args.keyword}
        valid = [r for r in valid if r["keyword"].lower() in kset]
    if args.start > 1:
        valid = valid[args.start - 1:]
    if args.limit:
        valid = valid[:args.limit]

    ego = find_ego()
    log(f"ego-browser: {ego or '未找到'}")
    log(f"待验证关键词 {len(valid)} 个（含广告位，每词取前 {args.top_n} 位）")

    if args.print_config:
        log(json.dumps({"workbook": str(wb_path), "sheet": info["sheet"], "site": site,
                        "targets": [str(t) for t in targets], "count": len(valid),
                        "judge_rules": JUDGE_RULES}, ensure_ascii=False, indent=2))
        return 0
    if not valid:
        log("× 没有待验证的有效关键词（可能全部是红底否定词）")
        return 1

    # --- 抓取 ---
    # --- 判定文件（先读，抓取过程中每批都刷 CSV）---
    jpath = Path(args.judgments) if args.judgments else work / "judgments.json"
    judgments = json.loads(jpath.read_text(encoding="utf-8")) if jpath.exists() else {}

    if not args.skip_scrape:
        if not ego:
            raise SystemExit("× 未找到 ego-browser，无法抓取亚马逊")
        if not target:
            asin = re.search(r"B0[A-Z0-9]{8}", info["sheet"] or "")
            target = fetch_asin_main_image(asin.group(0) if asin else "", ego, work, site["domain"])
            if not target:
                raise SystemExit("× 基准产品图缺失：请用 --target 指定，或确认表名里有 ASIN")
            log(f"基准图（ASIN 兜底）: {target}")
            targets = [target]

        batch_n = max(1, args.batch)
        total = len(valid)
        t0 = time.time()
        bi = 0
        blocked_retry = 0
        while bi < total:
            chunk = valid[bi:bi + batch_n]
            jobs = [{"kw": r["keyword"], "row": r["row"],
                     "out": str(work / "scrape" / f"{keyword_stem(r['keyword'])}.json").replace("\\", "/")}
                    for r in chunk]
            cfg = {
                "space": f"amazon-check-{slugify(info['sheet'])[:24] or 'task'}",
                "profileId": args.profile,
                "marketplace": site["domain"],
                "postcode": site["postcode"],
                "postcodes": [site["postcode"]],
                "topN": args.top_n,
                "maxPages": args.pages,
                "delayMs": int(args.delay * 1000),
                "jitterMs": int(args.delay * 500),  # 0.5x 随机抖动，贴近人工节奏
                "summary": str(work / "browser" / "summary.json").replace("\\", "/"),
                "timeout": 900 + batch_n * 60,
            }
            log(f"\n[{bi + 1}-{min(bi + batch_n, total)}/{total}] 抓取中…")
            summary = run_scrape(jobs, cfg, work, ego)
            loc = summary.get("location", {})
            log(f"  配送地址: {'OK ' if loc.get('ok') else 'FAIL'} {loc.get('text','')}")
            if not loc.get("ok") and not args.skip_location:
                log("  ! 配送地址未验证成功，排名可能不精准；先继续，建议下一轮检查选择器")
            for res in summary.get("results", []):
                if res.get("status") == "captcha":
                    log(f"  ! {res['keyword']} 触发验证码，已跳过")
            log(f"  进度 {min(bi + batch_n, total)}/{total}  用时 {(time.time()-t0)/60:.1f} 分钟")
            write_csv(build_rows(rows, work, judgments, args.top_n), Path(args.out))
            if summary.get("blocked"):
                blocked_retry += 1
                if blocked_retry <= 1:
                    log("\n× 命中亚马逊人机校验，冷却 10 分钟后自动重试本批（第 1 次重试）…")
                    time.sleep(600)
                    continue
                log("\n× 冷却重试后仍被风控拦截，停止本轮。")
                log("  已完成部分已写入 CSV；稍后重跑同一命令可断点续跑（已抓到的词不会重复请求）。")
                return 2
            blocked_retry = 0
            bi += batch_n
            if bi < total:
                rest = random.uniform(60, 90)
                log(f"  批次间休息 {rest:.0f} 秒…")
                time.sleep(rest)

    # --- 下图 + 拼图 ---
    if not args.no_images:
        log("\n下载主图并生成复核拼图…")
        for i, r in enumerate(valid, 1):
            sp = work / "scrape" / f"{keyword_stem(r['keyword'])}.json"
            if not sp.exists():
                continue
            data = json.loads(sp.read_text(encoding="utf-8"))
            items = data.get("items", [])
            if not items:
                continue
            kw_dir = work / "images" / keyword_stem(r["keyword"])
            download_all(items, kw_dir)
            sp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
            build_contact_sheets(r["keyword"], items, work / "review")
            if i % 10 == 0:
                log(f"  已处理 {i}/{len(valid)}")

    # --- 判定 + CSV ---
    if jpath.exists():
        kw_cnt = len([k for k in judgments if not k.startswith("_")])
        log(f"\n判定文件: {jpath}（{kw_cnt} 个词）")
    else:
        log(f"\n判定文件尚不存在：{jpath}（先生成拼图，人工/模型复核后回填）")

    out_rows = build_rows(rows, work, judgments, args.top_n)

    final = write_csv(out_rows, Path(args.out))
    done = sum(1 for r in out_rows if r[2] in ("已完成", "已完成（样本不足）"))
    log(f"\n✔ CSV 已输出：{final}")
    log(f"  已完成 {done} / 总 {len(out_rows)}（含否定词与待复核）")
    return 0


if __name__ == "__main__":
    sys.exit(main())




