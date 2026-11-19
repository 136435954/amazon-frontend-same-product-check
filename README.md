# amazon-frontend-same-product-check

亚马逊前台「同款占比」批量验证 Skill。

给定 **站点 + 1~3 张产品图 + 卖家精灵关键词表**，脚本自动设配送邮编、在真实前台搜索框逐词检索、抓每词前 50 个展示位（自然位 + Sponsored 广告位 + 品牌广告条 / SB 网格广告的可见瓦片）、下载主图拼成编号拼图，再由模型看图判定"外观同款"，最后输出带广告位数与同款占比的 CSV。

适用于：卖家精灵词表调研、前 50 位竞品与广告位盘点、某产品词的市场结构分析、多站点（美国 / 英国 / 德国 / 日本等 15 站）前台真实排名核查。

---

## 1. 它做什么、不做什么

**做**
- 15 个亚马逊站点，按站点自动设标准邮编（美国 90001、英国 DN11 0GU、德国 10115 …）
- 台前搜索框逐字输入（不是 `/s?k=` 直连），滚动加载，每词凑满 50 个**可见**展示位
- 识别广告位：站内广告卡、品牌广告条整条及其条内可见瓦片、SB 网格广告可见瓦片
- 红底（`FFFF0000`）关键词识别为否定词，写「否定词已跳过」，不搜索
- 产出 `amazon_analysis_result_top50.csv`

**不做**
- 不代点、不破解任何人机校验（`I am human` / `unusual browser activity`）
- 不默认美国站点，不猜口径，不沿用别的品类口径
- 不保证每个词都能跑完（亚马逊会按请求动态下发广告，数量会波动）

---

## 2. 依赖

| 项 | 要求 | 说明 |
|---|---|---|
| 系统 | Windows | 依赖 ego-browser 与豆包办公环境 |
| ego-browser | 0.6.3+ | 豆包办公自带；不在 PATH 时脚本会自动查找，也可设环境变量 `EGO_BROWSER` |
| Python | 3.10+ | 需要 `openpyxl`、`Pillow` |
| 模型视觉 | 判同款阶段 | 唯一消耗额度的地方 |

安装 Python 依赖：

```powershell
pip install openpyxl Pillow
```

---

## 3. 下载

**方式 A：Git**

```powershell
git clone https://github.com/<你的用户名>/amazon-frontend-same-product-check.git
```

**方式 B：不用 Git**

打开仓库页 → 绿色 `Code` 按钮 → `Download ZIP` → 解压到任意目录。

---

## 4. 安装到豆包 Skill 目录

Skill 要被豆包识别，必须放进 `workspace\.user_skills\`。本机路径通常是：

```
%LOCALAPPDATA%\Doubao\User Data\Profile 1\.doubao\agent_mode\workspace\.user_skills\
```

**一键安装**（在解压出来的仓库目录里执行）：

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

脚本会把 `skill\` 整个复制成 `<user_skills>\amazon-frontend-same-product-check\`，并打印目标路径；路径不对时可用 `-Dest` 指定：

```powershell
.\install.ps1 -Dest "C:\Users\<你>\AppData\Local\Doubao\User Data\Profile 1\.doubao\agent_mode\workspace\.user_skills"
```

**手工安装**：把 `skill` 文件夹整个复制进 `user_skills`，并重命名为 `amazon-frontend-same-product-check`。

装完**重开一个对话**，Skill 才会被加载。

---

## 5. 首次自检（不联网、不花额度）

```powershell
python skill\scripts\amazon_frontend_check.py `
  --workbook "D:\path\卖家精灵数据.xlsx" `
  --site US `
  --target "D:\path\产品图1.jpg" "D:\path\产品图2.jpg" `
  --print-config
```

应当看到：扫到的 Sheet 列表、主表名、关键词/否定词数量、站点与邮编、基准图（1~3 张）、判定口径。配置不对就在这里改，别往下跑。

---

## 6. 使用流程（四步）

### Step 0 — 准备三样输入

| 输入 | 说明 |
|---|---|
| 站点 | `US` / `UK` / `DE` / `JP` / `CA` / `AU` / `FR` / `ES` / `IT` / `SG` / `SE` / `PL` / `TR` / `IE` / `MX` |
| 产品图 | 1~3 张，正面主图为主，可补侧面、多件装、包装图 |
| 卖家精灵 xlsx | 含关键词列；多 Sheet 时自动选含 ASIN 或有效行最多的那张 |

脚本会自动读 A 列关键词、B 列翻译，并跳过红底否定词。

### Step 1 — 第一关：口径草案

读产品图，产出一版口径草案给用户确认（四项）：

1. 品类核心识别特征 2~4 条（具体到形状 / 结构）
2. 哪些近似品类算、哪些明确不算（各举 1~2 例）
3. 同型号多件装算不算、异款混装算不算
4. 颜色 / 尺寸 / 品牌 / logo / 包装 / 配件赠品怎么处理

**未通过不得抓取任何数据。换产品必须重来一遍。**

### Step 2 — 第二关：前 5 词出图复核

```powershell
python skill\scripts\amazon_frontend_check.py `
  --workbook "D:\path\卖家精灵数据.xlsx" `
  --site US `
  --target "D:\path\产品图1.jpg" `
  --limit 5 --delay 10 `
  --out "D:\path\amazon_analysis_result_top50.csv"
```

每词生成 2 张 5×5 编号拼图（`#01` 即第 1 位，格内左上到右下）。对着图复核口径，改完再确认，确认后把口径原文写进 `work\judgments.json` 的 `_meta.rules`。

### Step 3 — 批量抓取（零额度，可断点续跑）

```powershell
# 抓 1~500 个词，不下图（快）
python skill\scripts\amazon_frontend_check.py --workbook "<xlsx>" --site US `
  --target "<产品图>" --start 1 --limit 500 --no-images --delay 10
```

中断了就从没跑完的位置继续（`--start` 是 1 基的有效词序号），已抓过的词会被 `--skip-scrape` 直接复用。

### Step 4 — 下图 + 拼图 + 判同款

```powershell
# 4.1 只给子集下图、拼图（仍零额度）
python skill\scripts\amazon_frontend_check.py --workbook "<xlsx>" --site US `
  --target "<产品图>" --start 1 --limit 300 --skip-scrape

# 4.2 看图判同款，写入 work\judgments.json，再重跑出最终 CSV
python skill\scripts\amazon_frontend_check.py --workbook "<xlsx>" --site US `
  --target "<产品图>" --skip-scrape
```

> **额度提示**：读表、抓取、下图、拼图、写 CSV 全部零模型成本；只有「看图判同款」花钱。2000 词 ≈ 10 万位 ≈ 4000 张拼图，全量判不现实，默认只判高价值子集（200~300 词），其余词同款列留空待补。

---

## 7. 参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `--workbook` | 自动找 | 卖家精灵 xlsx；缺省在脚本目录搜「卖家精灵」 |
| `--sheet` | 自动 | 强制指定工作表名 |
| `--target` | 自动扫描 | 基准产品图，可给 1~3 张（空格并列） |
| `--site` | 表名前缀 | 站点键，建议显式传入；与表名前缀冲突时以 `--site` 为准并提示 |
| `--limit` | 0（全部） | 只跑前 N 个有效词 |
| `--start` | 1 | 从第 N 个有效词开始（1 基，断点续跑用） |
| `--keyword` | — | 只跑指定词，可重复 |
| `--out` | 脚本目录下 CSV | 输出 CSV 路径 |
| `--work-dir` | 脚本目录下 `work` | 工作区，**必须给绝对路径** |
| `--judgments` | `work\judgments.json` | 判定结果 JSON |
| `--skip-scrape` | 关 | 复用已抓取的 `work\scrape\*.json`，不重新联网 |
| `--no-images` | 关 | 不下图、不出拼图 |
| `--skip-location` | 关 | 跳过配送地址设置（仅调试） |
| `--top-n` | 50 | 每词取多少位 |
| `--pages` | 5 | 不足 50 位时最多翻到第几页 |
| `--delay` | 12 | 词间隔秒数（保守节流，别调太小） |
| `--batch` | 30 | 每个 ego 进程串行处理多少词 |
| `--profile` | 空 | ego profileId（如 `Profile 1`），**只在创建空间时生效** |
| `--print-config` | 关 | 只打印配置不执行 |

---

## 8. 产物

```
amazon_analysis_result_top50.csv       # 关键词|翻译|验证状态|广告位数(前50名)|前50名同款数量(含广告)|同款占比 (%)
work\
  scrape\<关键词>.json                 # 每词 50 条明细（ASIN/标题/主图/广告标记/位次）
  images\<关键词>\<位次>_<ASIN>.jpg    # 主图，位次前缀保证 1 ASIN 多位次也不冲突
  review\<关键词>_<hash>_sheet1.jpg    # 前 1~25 位 5×5 编号拼图
  review\<关键词>_<hash>_sheet2.jpg    # 前 26~50 位
  judgments.json                       # 判定结果 + _meta.rules（口径原文）
  browser\                             # 临时 .mjs 与 summary.json
```

各站点域名 / 邮编表见 `skill\references\marketplaces.md`；广告位判定的 DOM 判据与可见性代码见 `skill\references\dom-recipes.md`。

---

## 9. 计数与判定规则

1. 站点由用户给，不默认美国；站点决定域名与邮编，猜错整份数据作废。
2. 第一个关键词之前必须把配送地址改成该站标准邮编，并确认页面顶部文案（如 `Los Angeles 90001`）。
3. 每词必须凑满 50 个可见展示位，不足就翻页续取；翻到第 5 页仍不足，按实际位数记录并注明。
4. 一个 ASIN 占多个位次就按位次重复计（广告位 + 自然位各算 1）。
5. 广告位计入：站内广告卡、品牌广告条整条及其条内**可见**瓦片、SB 网格广告**可见**瓦片；被轮播 overflow 裁掉的隐藏瓦片既不算广告位也不算名额。
6. 搜索必须在前台搜索框逐字输入 + 回车，禁止 `/s?k=` 直连。
7. 同款 = 同品类 + 同可见外形；品牌 / 颜色 / logo / 尺寸 / 包装不同都算同款；同型号多件装算 1 个位次，不同款混装不算。
8. 红底关键词写「否定词已跳过」，不搜索。
9. 同款占比 = 同款位次数 / 实际样本位次数 × 100，保留 1 位小数。

---

## 10. 风控

- 命中 `/checkpoint/`、`/errors/validateCaptcha`、页面出现 `I am human` / `unusual browser activity` → 该词记 `captcha`，整批早停，冷却 600 秒后重试一次。
- **绝不代点、绝不破解人机校验**。连续两次命中就交给用户处理，CSV 写 `风控未完成`。
- 词间默认 12 秒，同一会话内不要并行搜索。

---

## 11. 常见问题

**Q：报 `page label not found: p1` / `Managed Pages in task space 0: none` / `ego 空间未就绪：没有可用标签页`？**
ego 空间里没有可用标签页，或标签页预算已满（`Page budget reached (8/8)`）。脚本内置 `ensurePage()`：优先复用 `task.tabs()` 里已有的页面，只有完全没有页面时才新建。注意两个 ego 的坑——`task.adopt()` 只接受**未跟踪**页面，对已被空间跟踪的标签页会报 `requires an untracked page`；旧版 `taskSpace(id)` 接管 agent 空间时可能拿不到标签页，需要用 `claimTaskSpace(id)`。若仍失败，先结束残留的 ego 空间再重跑。

**Q：报 `dDefaultResolve` 之类模块解析错误？**
`--work-dir` 用了相对路径。改成绝对路径。

**Q：抓到的广告位数是 0，但页面上明明有 Sponsored？**
先跑证据链：`sp-sponsored-result`、`cel_widget_id` 含 SP、`.AdHolder`、aria 计数，任一项非 0 却判 0 就是漏判。对照 `references\dom-recipes.md` 核对当前版式。

**Q：同一关键词两次跑，广告数不一样？**
正常。亚马逊按请求动态下发广告，跨词比较请用同一批跑出来的数。

**Q：能一次跑完 2000 词并全部判同款吗？**
不建议。抓取零成本，但判同款要花额度（≈4000 张拼图）。默认只判高价值子集；全量判定需显式确认。

**Q：换产品还能用吗？**
可以，但要重新走第一关：读新图、出新口径草案、你确认。不许沿用上一个品类的口径。

---

## 12. 目录结构

```
amazon-frontend-same-product-check/
├── README.md                 # 本文件
├── install.ps1               # 一键安装到 .user_skills
├── .gitignore
└── skill/                    # 技能本体（复制进 user_skills 的内容）
    ├── SKILL.md
    ├── references/
    │   ├── marketplaces.md
    │   └── dom-recipes.md
    └── scripts/
        └── amazon_frontend_check.py
```

---

## 免责

本项目仅用于自建店铺的选品与词表分析。使用前请自行确认符合亚马逊服务条款与当地法律；因抓取频率过高导致的限流、封禁由使用者自行承担。
