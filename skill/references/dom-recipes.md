# 亚马逊搜索结果页 DOM 判据（实测记录）

用途：广告位识别。这些版式在 2026-11 实测通过（US 站，英文界面）。版式随站点/AB 试验变化，拿不到数时按本文件逐条核对。

## 1. 单个自然卡 / 站内广告卡

容器：`div[data-component-type="s-search-result"]`，带 `data-asin`。

广告判据（任一命中即广告）：

| 判据 | 选择器 / 特征 |
|---|---|
| 主判据 | 卡内存在 `[aria-label*="Sponsored" i]`，典型值 `View Sponsored information or leave ad feedback` |
| 标签节点 | `[data-component-type="s-sponsored-label-info"]`、`.puis-sponsored-label-text`、`.s-sponsored-label-info` |
| 文本 | 卡内 `innerText` 有独立成行的 `Sponsored`，或首行以 `Sponsored` 开头 |
| 广告容器 | 祖先（≤8 层）类名含 `a-carousel`、`a-carousel-card`、`a-carousel-viewport`、`_c2Itd_`，或带 `data-a-carousel-type` |

## 2. 品牌广告条（brand shelf）

结构（BOENFU 实例）：

```
div.s-result-item.s-widget
└─ div.celwidget
   └─ div._c2Itd_container_<hash> sb-desktop
      ├─ a._c2Itd_link_*            ← aria-label="Sponsored ad from BOENFU…Shop BOENFU"
      └─ 8 个 [data-asin] 瓦片      ← 无 h2，标题在 img[alt]
```

判据：祖先（≤6 层，遇到 `s-search-result` 即否）同时满足「含 2~15 个 `[data-asin]`」+「文本或 aria 含 Sponsored」。

补充标记（瓦片自身无 aria 时靠它命中）：

```html
<span class="_c2Itd_ad-feedback-text-desktop_…" aria-label="Leave feedback on Sponsored ad">Sponsored</b></span>
```

选择器：`._c2Itd_ad-feedback-text-desktop, ._c2Itd_ctaSponsoredContainer, [aria-label="Leave feedback on Sponsored ad"], [aria-label*="Sponsored ad from" i]`

## 3. SB 网格广告（Sponsored Brands search grid）

容器：`a-carousel-display-sb-search-grid` + `._c2Itd_*`，`ol.a-carousel` 里 1~16 个 `[data-asin]` 瓦片，`div.a-carousel-viewport` 做横向裁剪。

**关键规则**：屏幕上一次只显示约 5 个瓦片，其余被 overflow 裁掉。只把与 `.a-carousel-viewport` 横向重叠 >50% 的瓦片算作位次；隐藏瓦片既不算广告位、也不占前 50 的名额。

```js
function carouselVisible(el) {
  if (!inCarousel(el)) return true;
  const vp = el.closest(".a-carousel-viewport");
  if (!vp) return true;
  const a = el.getBoundingClientRect(), b = vp.getBoundingClientRect();
  const ov = Math.min(a.right, b.right) - Math.max(a.left, b.left);
  return ov > Math.min(a.width, b.width) * 0.5;
}
```

## 4. 广告确实不存在的证据链

判定"0 广告"之前，先确认不是漏判：

| 检查 | 期望（无广告时） |
|---|---|
| `[data-component-type="sp-sponsored-result"]` | 0 |
| `cel_widget_id` 含 SP/sponsored 的节点 | 0 |
| 整页 HTML 含 `sp-sponsored-result` / `MAIN-SP` | false |
| `.AdHolder` | 存在但内部为空 |
| `[aria-label*="Sponsored"]` 节点数 | 0 |

任一项不为 0 而脚本仍判 0，属于漏判，回来核对本文件的版式条目。

## 5. 取标题的兼容顺序

`[data-cy="title-recipe"] h2 span` → `h2 span` → `h2` → `img[alt]`。广告瓦片常无 h2，只能读 `img[alt]`。
