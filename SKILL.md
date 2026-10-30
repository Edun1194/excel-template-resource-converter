---
name: template-resource-converter
description: |
  将不同供应商的 SKU 报价单/库存单(含 color、size、库存、cost、cost+)自动识别并转换为固定模板
  (template1-cn/template1-en)。覆盖 Excel/Word/PDF 来源,内含 37 个内置供应商适配器
  (V1001-V1072 + 期货通道 + category-futures 聚合通道),支持多 Section linesheet / 平铺库存清单 /
  col14 块状 / NuORDER / PDF 等版式族。
  触发:用户要把多供应商 Excel/Word 报价/库存格式统一成目标模板;需要分析新供应商格式并生成
  转换脚本骨架;需要在已有 vendor 笔记基础上 debug 转换输出。
displayNames:
  zh-Hans: "转换标准货单"
---

# Template Resource Converter

执行供应商报价与库存文件到固定模板的自动转换,默认输出中文模板 `assets/templates/template1-cn.xlsx`。
图片字段用 Excel 内嵌图片(embed image)写入,不是链接文本;当单行存在多张图片时会以独立图片对象写入同一行。

## Quick Start

```bash
SKILL=~/.minimax/skills/template-resource-converter

# 1. 装依赖
python3 -m pip install -r "$SKILL/requirements.txt"

# 2. 转换(自动识别 vendor)
python3 "$SKILL/scripts/convert_vendor_quote.py" \
  --primary-file "path/to/vendor-stock.xlsx"

# 3. 指定 vendor + C+ + 仓库 ID
python3 "$SKILL/scripts/convert_vendor_quote.py" \
  --primary-file "path/to/vendor-stock.xlsx" \
  --vendor "V1061" --cplus 10 --warehouse-id "V1061"

# 4. 英文模板
python3 "$SKILL/scripts/convert_vendor_quote.py" \
  --primary-file "path/to/vendor-stock.xlsx" --language en
```

### 完整 CLI 参数

| 参数 | 说明 |
|---|---|
| `--primary-file` | **必填**,供应商源文件(.xlsx / .xls / .pdf) |
| `--vendor` | vendor ID;不传则按文件名/内容自动检测 |
| `--markup-file` | 旧版 markup 文档(.xlsx/.docx),多数 vendor 已不需要 |
| `--language {cn,en}` | 模板语言,默认 `cn` |
| `--template` | 直接指定模板文件,覆盖 `--language` |
| `--warehouse-id` | 覆盖「云仓」列的 vendor ID |
| `--output` | 输出路径;默认与源文件同目录 |
| `--cplus` | **本次订单**的 C+ 整数(如 `10` = 10%,`-18` = 让利 18%) |
| `--cplus-rules-json` | C+ 品类规则 JSON(`[{field,keyword,sign,cplus}]`),前端用 |
| `--cplus-rules-file` | C+ 品类规则 yml 路径(CLI 用) |
| `--cplus-overrides-file` | C+ 行级覆盖 CSV(`sku,cplus`),优先级最高 |
| `--cplus-default` | 规则都没命中时的兜底 C+;`None` = 这些行留空 |
| `--no-compute-supply-price` | 只写 C+ 文本,批发价留空 |
| `--brand` | 覆盖品牌名 |
| `--season` | 覆盖季节(如 `SS26` / `FW26`) |
| `--image-row-height` | 有图数据行的目标行高(pt),默认 `92` |
| `--image-col-width` | 图片列列宽(Excel 字符单位),多图产品需加宽 |
| `--keep-images-per-row` | 每行保留几张图;`None` = 用 vendor 默认值 |

## Use Cases

1. **批量转换多供应商报价单** — 接到 10+ 个品牌(BARBOUR / NUGNES / POMPEU / GUCCI 等)的报价 Excel,要统一成同一目标模板
2. **新供应商接入** — 用户给一个新的品牌文件,需要先分析结构 → 写适配器 → 输出
3. **图片提取** — 报价单里的产品图要从源文件(.xls OLE2 / .xlsx drawing / WPS richValue / PDF 嵌入)抠出来放到目标模板
4. **bug 修复** — 某个 vendor 输出缺图/缺尺码/价格错位,需要按 references 笔记查同类问题修复
5. **多 Section 通用适配** — 新 brandsheet 结构(行 0-4 metadata + 多尺码矩阵),可复用 `convert_vendor_quote.py` 通用入口

## File Layout

- **转换脚本和最终输出**: 默认存放在与源报价单**相同的目录**下
  - 输出文件名: `<VendorID>-template1-cn-converted.xlsx` 或 `converted/<primary-stem>.xlsx`
- **临时文件** (图片缓存): 统一放到 `/tmp/template-resource-converter/`
- **新供应商转换脚本**: 以 `convert_<VENDOR>.py` 命名,放在供应商报价单**所在目录**,主调度器会自动发现并加载

## Required Target Fields

模板 `template1-cn.xlsx`(18 列) / `template1-en.xlsx`(18 列)标黄表头必填:

| # | 中文 | 英文 | 说明 |
|---|---|---|---|
| 1 | `品牌` | `Brand` | 品牌名 |
| 2 | `图片` | `Image` | 嵌入图片(不是链接文本) |
| 3 | `货号` | `International code` | SKU / Style code;**当源文件 STYLE 不含颜色代码时,自动拼接为 `{STYLE}-{颜色代码}`** 保证货号唯一性;**分隔符固定用 `-`** |
| 4 | `描述` | `Described` | 商品描述 / 类型 |
| 5 | `类目` | `Category` | 品类 |
| 6 | `材质` | `Material` | 材质成分 |
| 7 | `性别` | `Gender` | MAN / WOMAN / KIDS / Unisex |
| 8 | `颜色` | `Color` | 颜色描述(用户 2026-07-15 确认,可空) |
| 9 | `色号` | `Color Code` | 源文件没有就留空 |
| 10 | `季节` | `Season` | SS26 / FW26 … |
| 11 | `产地` | `Made in` | 源文件没填(如 `Non indicato`)就留空 |
| 12 | `零售价` | `Retail price` | 建议零售价(数字) |
| 13 | `cost` | `cost` | 供货价(数字) |
| 14 | `C+-` | `C+-` | 加价率,**纯整数**(正数 `18`、负数 `-18`);无 C+ 时**留空** |
| 15 | `批发价` | `Wholesale price` | 供货价 = `cost × (1 + C+-/100)`;**无 C+ 时留空,不填 0** |
| 16 | `尺码` | `Size` | 单个尺码,矩阵展开后每行一个 |
| 17 | `数量` | `Quantity` | 库存数(int,0 也写入) |
| 18 | `云仓` | `Vendor ID` | 供应商 / 仓库 ID(V1021 / V1044 / V1061 …) |

> ⚠️ **列名变更(2026-09 起)**:`仓库ID` → `云仓`,`库存` → `数量`,`供货价` → `批发价`。
> 旧代码里按老列名找列会静默拿到空值,一律按**表头名**定位。

脚本会校验模板是否具备这些列,缺失则报错。

## ⚠️ C+ 是「每单单独给」的(2026-10-28 用户确认)

> 原话:「是我单独给的,不用保存」

- 每次转换以**用户当次给的 C+ 值**为准(`--cplus 10` = 10%,批发价 = cost × 1.10)。
- **不要**把值写进品牌查找表当品牌默认值,**不要**硬编码品牌级常量,**不要**套 `0.18` 之类的默认加价率。
- 用户没给 C+ → C+ 列和批发价**都留空**,并跟用户说明。
- 详见 `references/patterns/cplus-column.md`。

## Image Sources

图片数据支持多种来源,解析策略完全不同:

1. **Excel 内嵌图片** (`xl/drawings/drawingN.xml`)
   - 命名空间 `<xdr:ext>` 或 `<a:ext>`(必须两种都尝试,见 image-handling.md §1)
   - `twoCellAnchor` / `oneCellAnchor` / `vm=N` 三种锚点格式
2. **图片目录** — 供应商提供文件夹,按货号或 SKU 命名匹配
3. **品牌方 PDF 文件** — lookbook / line-sheet,需 `pdftotext` + `pdfimages` + `pdftohtml` 组合解析

PDF 解析工具(系统级,非 Python):

| 工具 | 用途 | 安装 |
|---|---|---|
| `pdftotext -layout` | 保留布局提取结构化文本 | `brew install poppler` |
| `pdfimages -all` | 批量提取嵌入图片 | 同上 |
| `pdftohtml -xml` | 保留文本/图片坐标 | 同上 |
| `pdfplumber` | 精确表格/坐标提取 | `pip install pdfplumber` |
| `PyMuPDF` (`fitz`) | 高效文本+图片 | `pip install PyMuPDF` |

详细 PDF 解析工作流见 `references/patterns/PDF-image-extraction-notes.md`。

## Code Architecture

```
template-resource-converter/
├── SKILL.md                          # 本文件
├── README.md                         # GitHub 首页说明
├── requirements.txt                  # pandas, openpyxl, pillow, PyYAML …
├── assets/templates/
│   ├── template1-cn.xlsx             # 中文目标模板(默认,18 列)
│   └── template1-en.xlsx             # 英文目标模板(18 列)
├── scripts/
│   ├── convert_vendor_quote.py       # 主入口 + dispatcher(37 个 BUILTIN_VENDORS)
│   ├── convert_vendor_quote_shared/  # shared 包(不是单文件!)
│   │   ├── __init__.py               #   写输出、图片加载、模板校验、C+ lookup
│   │   ├── flat_stocklist.py         #   平铺库存清单版式(共用)
│   │   ├── col14_parsers.py          #   col14 块状版式(共用)
│   │   └── cplus_resolver.py         #   C+ 3 层优先级 resolver
│   ├── category_futures/             # 期货聚合通道的 6 个版式解析器
│   │   ├── common.py                 #   open_safe_workbook / 表头定位
│   │   ├── linesheet.py  nuorder.py  boss.py
│   │   └── salomon.py   longchamp.py vw.py
│   ├── analyze_vendor_format.py      # 新供应商结构分析工具
│   ├── convert_batch.py              # 批量转换
│   └── vendor_*.py                   # 各供应商适配器
├── outputs/                          # generated adapter 样例
└── references/
    ├── patterns/                     # 通用技术/踩坑模式
    └── vendors/                      # 少量品牌专属笔记
```

### 共享模块为什么是「包」不是单文件

`flat_stocklist` / `col14_parsers` / `cplus_resolver` 三个模块被多个 vendor 共用
(V1038 / V1061 / V1062 共用平铺解析器;V1062 同时管平铺 + col14 两种版式)。
**多 vendor 共用同一版式 → 放 shared**,不要复制到各 vendor 里。

## Built-in Vendor Adapters (37)

主调度器 `convert_vendor_quote.py` 的 `BUILTIN_VENDORS` 自动按 vendor ID 加载:

| Vendor ID | 文件 | 说明 |
|---|---|---|
| `V1001` | `vendor_v1001.py` | BARBOUR 货单(稀疏表:每产品块 15 行,块内每尺码一行) |
| `V1003` | `vendor_v1003.py` | McQueen SS27 期货(.xls + .pdf preprocess) |
| `V1003-期货` | `vendor_v1003_futures.py` | 期货系总入口:McQueen + Armani PLM Report + LONGCHAMP SS27,输出 vendor_id 统一 `V1003` |
| `category-futures` | `vendor_category_futures.py` | 期货聚合通道,「云仓」列写 `品牌名称（期货）` |
| `V1004` | `vendor_v1004.py` | 现货 4 品牌(ARC'TERYX / PATAGONIA / PRL / STONE ISLAND) |
| `V1004-期货` | `vendor_v1004_futures.py` | SALOMON Foglio 版式,数量留空等客户填 |
| `V1010` | `vendor_v1010.py` | LONGCHAMP(col14 块状版式) |
| `A1001` | `vendor_a1001.py` | CANADA GOOSE hk 现货 .xls |
| `V1011` | `vendor_v1011.py` | MAX MARA 意大利供应商 CATALOGO sheet |
| `V1012` | `vendor_v1012.py` | |
| `V1015` | `vendor_v1015.py` | |
| `V1015-补货单` | `vendor_v1015_restock.py` | richValue 路径,意大利大公司 SharePoint 风格 |
| `V1015-补货单2` | `vendor_v1015_restock2.py` | MIUMIU ATS,3 sheet 多品类 |
| `V1015-补货单3` | `vendor_v1015_restock3.py` | GIORGIO ARMANI C-20,4 sheet 横排尺码 |
| `V1015-补货单4` | `vendor_v1015_restock4.py` | EA7 C-20,1 sheet 纵排 17 列 |
| `V1021` | `vendor_v1021.py` | BRUNA ROSSO |
| `V1025` | `vendor_v1025.py` | POLO RALPH LAUREN 现货单 xlsx 新版式 |
| `V1031` | `vendor_v1031.py` | NUGNES |
| `V1037` | `vendor_v1037.py` | SALOMON 意大利线(WPS 从 .xls 转) |
| `C0000` | `vendor_c0000.py` | GIVENCHY China 库存清单 |
| `V1038` | `vendor_v1038.py` | APC / TORY BURCH(col14)+ GUCCI WAREHOUSE A(平铺) |
| `V1039` | `vendor_v1039.py` | ANTONIOLI |
| `V1041` | `vendor_v1041.py` | BIFFI |
| `V1042` | `vendor_v1042.py` | ALINA / POLO RALPH LAUREN |
| `V1043` | `vendor_v1043.py` | PRITELLI |
| `V1044` | `vendor_v1044.py` | RATTI(需 MU_SS26 markup) |
| `V1046` | `vendor_v1046.py` | LOSCHI |
| `V1048` | `vendor_v1048.py` | MALTEMPI |
| `V1049` | `vendor_v1049.py` | POMPEU |
| `V1050` | `vendor_v1050.py` | FOLLIFOLLIE + POLO RALPH LAUREN ATS(`MATERIAL NUMBER` 版式, richValue 图, 尺码行共享图) |
| `V1061` | `vendor_v1061.py` | GUCCI FW26 - WAREHOUSE B(平铺) |
| `V1062` | `vendor_v1062.py` | GUCCI WAREHOUSE B/C(col14 老单)+ CANADA GOOSE FW26 WAREHOUSE C(平铺) |
| `V1063` | `vendor_v1063.py` | |
| `V1064` | `vendor_v1064.py` | |
| `V1067` | `vendor_v1067.py` | PANICALE 现货单 + 外部图片目录匹配 |
| `V1070` | `vendor_v1070.py` | BRUNELLO CUCINELLI 现货单 |
| `V1070-期货` | `vendor_v1070_nuorder.py` | BC NuORDER SS27/SS28 |
| `V1072` | `vendor_v1072.py` | BURBERRY 现货单 |
| `V1072-期货` | `vendor_v1072_in_arrivo.py` | LISTE FW26 IN ARRIVO 8 个品牌 |

**未内置的 vendor**:按「新供应商 5 步」流程,在报价单同目录放 `convert_<VENDOR>.py`,
主调度器会自动发现并加载;也支持 `scripts/vendor_<vendor>.py` fallback。

## 版式族(接新文件时先对号入座)

| 版式族 | 特征 | 详见 |
|---|---|---|
| **多 Section linesheet** | 行 0-4 metadata + 多尺码矩阵,section 靠代码识别 | `references/patterns/section-detection-and-row-lifecycle.md` |
| **平铺库存清单** | 意大利语表头,`Modello/Variante` + `TOTALE`,尺码铺成横矩阵,一行一个尺码 | `references/patterns/flat-stocklist-layout.md` |
| **col14 块状** | 货号在 col 14,块内纵向排 描述/尺码/数量 | `references/patterns/col14-block-layout.md` |
| **期货通道** | 没有真实库存,数量留空;云仓写工厂代号或 `品牌名称（期货）` | `references/patterns/futures-channel.md` |
| **NuORDER** | sheet `NuORDER Order Data`,同一份文件不同品牌列序完全不同 | `references/patterns/futures-channel.md` |
| **OLE2 .xls** | 老式二进制 Excel,图片是 JPEG 流 | `references/patterns/OLE2-xls-JPEG-extraction-pattern.md` |
| **PDF lookbook** | 品牌方 PDF,需抽文字+抽图 | `references/patterns/PDF-image-extraction-notes.md` |

## Generic Bug Patterns (最致命的那些)

写新适配器时,这些 bug 反复出现。详细修复见对应 references。

### 1. `isinstance(v, (int, float))` 必须 (xlrd vs openpyxl 整数差异)

xlrd 读整数返回 `float`(`32.0`),openpyxl 返回 `int`(`32`)。用 `isinstance(v, float)` 会漏掉 openpyxl 的 int → FENDI V section 整数尺码全部消失 → 所有尺码变 TU。

**修复**: 统一用 `isinstance(v, (int, float))`。
→ `references/patterns/type-and-cell-traps.md §1`

### 2. 数字 `0` 灰格陷阱

xlrd 灰格返回 `0.0`/`0`,如果用 `v > 0` 判断合法尺码,`0` 会被误认为是尺码 "0"。

**修复**: `v == 0` 单独判断返回 `None`。
→ `references/patterns/type-and-cell-traps.md §2`

### 3. `'L'` 和 `'S'` 不应是 section code

A2 section 尺码标签包含 `S/M/L/XL`。如果把 `'L'` 加进 `section_codes`,`_get_size_label` 会跳过这些尺码。

**修复**: `section_codes` 只放真正的 section 标识符,不放单字母尺码。
→ `references/patterns/type-and-cell-traps.md §3`

### 4. EMU 命名空间 `<a:ext>` vs `<xdr:ext>`

- `<xdr:ext cx cy>` — HOGAN, HOKA, Keen Kids, VIVWESTWOOD
- `<a:ext cx cy>` — STONE ISLAND, JILSANDER, FENDI, MIU MIU, MONCLER, MAMMUT

**修复**: 先查 `<xdr:ext>`,fallback `<a:ext>`。
→ `references/patterns/image-handling.md §1`

### 5. `<a:srcRect>` 图片裁剪(百万分比)

FENDI/GUCCI xlsx 用 `<a:srcRect l="41307" t="22650" r="42702" b="36365"/>` 裁剪原图,单位是百万分比(÷ 100000)。**不解析 srcRect 会嵌入完整原图**。

**修复**: PIL `crop((int(w*l), int(h*t), int(w*(1-r)), int(h*(1-b))))` 后重算 display_emu。空 srcRect `<a:srcRect/>` 要保护。
→ `references/patterns/image-handling.md §3`

### 6. C+ 整数 vs 字符串(支持负数,无 C+ 留空)

模板 `C+-` 列必须存**纯整数**,正数 `18`、负数 `-18` 保留负号,不能是字符串 `"18"` 或带 `"C+"`/`"%"` 前缀。`format_cplus_text(0.18)` 返回字符串,不能直接用。

**`supply_price` 配套规则**: 无 C+ 时**留空(None)**,不填 0。
→ `references/patterns/cplus-column.md`

### 7. 大文件 `openpyxl read_only` 超时

>5000 行的 xlsx,`openpyxl.load_workbook(..., read_only=True)` 会 >300s 超时。

**修复**: 用 `zipfile + xml.etree.ElementTree` 直接解析 `xl/worksheets/sheetN.xml` + `xl/sharedStrings.xml`。**不要用 regex 解析 sharedStrings**——`<si>` 含 `<t/>` 自闭合或多 `<t>` 时会 off-by-one。
→ `references/patterns/EXCEL-xlsx-XML-parsing-guide.md`

### 8. 图片加载器选择(按文件版式)

`convert_vendor_quote_shared` 里有 4 个图加载器,**选错就漏图**。按文件版式选:

| 加载器 | 适用场景 | 代表 vendor |
|---|---|---|
| `load_sheet_embedded_images` | 图 descr 里有 SKU,anchor 行号和数据行错位 | V1051 SALOMON |
| `load_sheet_images_by_row` | 单 sheet,图 anchor 行号 = 数据行号 | V1031 FENDI, V1010 LONGCHAMP, **V1049 POMPEU(多图同 cell)** |
| `load_images_by_row_per_sheet` | 多 sheet,图 anchor 行号 = 数据行号 | V1044 RATTI, V1046 LOSCHI, V1048 MALTEMPI |
| `load_images_by_merged_range` | 图片所在 cell 是 N 行合并,同款多尺码共用一张图 | **V1043 PRITELLI** |

**判定方法**:
1. 看 `len(ws.merged_cells.ranges)` → 几十以上且 col A 有大量合并 → `load_images_by_merged_range`
2. 看图数 vs 行数:`len(ws._images) >> max_row` → 多图同 cell,用 `load_sheet_images_by_row` 但**取全部**而不是 [0]
3. 看 `Workbook.sheetnames` 数量 → 多 sheet 用 `load_images_by_row_per_sheet`
4. 看 descr 是否含 SKU → 用 `load_sheet_embedded_images`

**多图同 cell 必须设 `default_image_col_width`**(见 image-handling §10)。
→ `references/patterns/image-handling.md §8-§10`、`adapter-development.md §3.9`

### 9. `openpyxl.Image.width/height` 是**像素**单位(不是 EMU)

写 `img.width = X` 实际写入 ext `cx = X * 9525` EMU。**绝对不要**写 `img.width = X * 9525`(双倍乘,图会变成 79ft 高)。

```python
# ✅ 想要 cx = 92pt = 1,168,400 EMU → 写 img.width = 92 (像素)
img.width = 92
img.height = 92
# ❌ 错误:写 img.width = 92 * 9525 = 876300 → 实际 ext cx = 876300 * 9525 = 巨大
```
→ `references/patterns/image-handling.md §16`

### 10. 保清晰度模式 — OneCellAnchor 改 ext,bytes 不动

**用户场景**:"图片显示 92pt 但保持清晰度不变"。

**错误做法**:`PIL.resize` 缩到 92pt → 重新存 JPEG → 1.4KB/张,糊。

**正确做法**:**原图 bytes 一字不动**,只改 openpyxl 写入的 ext cx/cy。

```python
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D

img = XLImage(path)  # 不动 bytes
cx, cy = compute_display_emu(w_px, h_px, max_pt=92)
img.anchor = OneCellAnchor(
    _from=AnchorMarker(col=image_col - 1, row=r - 1, colOff=0, rowOff=0),
    ext=XDRPositiveSize2D(cx=cx, cy=cy),
)
ws.add_image(img)
```
→ `references/patterns/image-handling.md §17`

### 11. 库存为 0 跳过 — 整行不输出

旧逻辑会输出 1 行 "size=OS qty=0" 占位 row。**新规则**:所有 size 列都 0 → 整行跳过。

```python
if not size_rows:
    continue  # 整行跳过
```

### 12. 阿里云推送 SMTP 附件限额 + Clash fake-IP 拦截

- `smtpdm.aliyun.com:465`:附件 ~10-20MB 上限(7.3MB ✅ / 17.5MB ❌ / 27MB ❌)
- `smtp.qiye.aliyun.com:465`:限额更大,27MB 能发,115MB 需 `split -b 20m` 分卷多封
- **VPN 拦截诊断**:`nslookup smtpdm.aliyun.com` → 真实 IP `47.x`/`106.x` = 通;Clash fake-IP `198.18.0.52` = SSL EOF,必失败
- **唯一有效绕过**:关 Clash 30 秒

### 13. 按表头名定位列,绝不写死列号

同一份 NuORDER,ACNE / BC / Men ss27 / POLO RL 四家的列序都不一样(19 / 21 / 55 / 29 列)。
写死列号 = 下次换文件就全错位。

**修复**: 表头名归一化(去括号、去标点、小写)→ 别名表覆盖全部变体
(`'Wholesale (EUR)'` → `wholesaleeur`,所以别名要同时写 `Wholesale (EUR)` / `Wholesale Price (EUR)` / `Wholesale` …)。
→ `references/patterns/futures-channel.md`

### 14. detect 的表头特征必须**同时含全部关键词**

宽松匹配会让 Linesheet 和 NuORDER 互相误判——两者都有 `Style Number`。
判定条件要写成「同时命中 sku + sizes」「同时命中 sku + qty」这种**合取**。

### 15. 尺码区起点 = 最后一个**已映射固定字段**列 + 1

否则 `Name` 这种 4 字母词会被误判成尺码列。

### 16. 数量列 = 尺码列右边最近的表头含 `Qty` 且**不含 `size price`** 的列

没有就留空。不要猜。

### 17. 半码归一:`'7/'` → `'7½'`

同一份文件里半码可能混用 ASCII 斜杠和标准 `½`(WAREHOUSE B 的 2-11 鞋码段用 `'7/'`,同一文件另一段用 `'36½'`)。
**有源内证据才归一**,不要凭感觉。

### 18. 同一列语义按**内容**判定,不按品牌硬编码

GUCCI 的 `Descrizione` 是材质成分(`70% CO 30% PL`),CANADA GOOSE 的同一列是商品描述。
判据:整份文件看多数数据行该列是否含 `%`。
→ `references/patterns/flat-stocklist-layout.md`

### 19. openpyxl 读不了的 xlsx 要先"修"再读

- `<sheetProtection insertLinks="1">` → TypeError(Alexander Wang)
- `<dataValidation type=非法值>` → ValueError(LONGCHAMP)

**修复**: `category_futures/common.py` 的 `open_safe_workbook()` —— 复制临时文件摘掉标签再 load,还不行退回 `read_only`。
→ `references/patterns/futures-channel.md`

### 20. 期货通道数量留空 ≠ 填 0

`0` 表示「确实没库存」,「留空」表示「还没填」,语义不同。
→ `references/patterns/futures-channel.md`

## Adaptation Workflow (新供应商 5 步)

1. **分析结构** — 用 `scripts/analyze_vendor_format.py` 或自己写 xlrd/openpyxl 探查:文件类型 / sheet / metadata 行 / 数据列 / 图片格式
2. **对号入座版式族** — 先看上面「版式族」表;命中已有版式就直接复用 shared 解析器,别写新 parser
3. **写适配器** — 复制 `vendor_v1021.py` 等内置适配器骨架,按 `references/patterns/adapter-development.md` 改字段映射
4. **跑 + 验证** — 必查:行数/图片数/cplus_text 整数/批发价公式/货号格式/颜色/材质/尺码集合/数量合计
5. **不进 references/** — 这次调试的踩坑如果触发了**新通用模式**,补到 `references/patterns/` 相应文件;否则该供应商专属的列偏移不进 skill

### 适配器最小骨架

```python
"""VendorName 适配器 (2026-XX-XX)."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from convert_vendor_quote_shared import write_output   # 按实际 import 调整

VENDOR_ID = "V10XX"
BRAND = "VENDOR NAME"


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    ...
    return rows


def get_vendor_meta() -> Dict[str, object]:
    return {
        "default_image_col_width": 40.0,
        "default_keep_images_per_row": 1,
    }
```

注册两处才算生效:

1. `scripts/convert_vendor_quote.py` → `BUILTIN_VENDORS` 加 `"V10XX": "vendor_v10xx"`
2. (如果接 Web UI) `app.py` → `VENDOR_CODES` 加 `"V10XX"`

## 大文件性能陷阱 (>5000 行)

任何超过 5000 行的 xlsx,**不要用 `openpyxl.load_workbook(read_only=True)`**,会超时(>300s)。改用 `zipfile + xml.etree.ElementTree` 直接解析。

**MAMMUT ATS 实测**: 7700+ 行,read_only 超时,zipfile+ET 5 秒。

**sharedStrings 解析必须 ET,不要 regex**: 文件含 `<si><t/></si>` 自闭合或多 `<t>` 时,regex 会 off-by-one,索引全乱。

完整代码见 `references/patterns/EXCEL-xlsx-XML-parsing-guide.md`。

## references 索引

**通用模式** (`references/patterns/`):

| 文件 | 内容 |
|---|---|
| `image-handling.md` | 图片全流程(EMU、srcRect、92pt、RGBA、WPS richValue、行号、合并单元格图、多图同 cell、列宽、colOff 累加 bug、全局只保留第一张图、货号列宽 26、openpyxl Image 像素单位、OneCellAnchor 保清晰度、按目录文件名匹配) |
| `image-directory-matching.md` | 按目录文件名匹配图(PANICALE FW26 案例,精确 (sku, color),保清晰度,库存 0 跳过,分卷邮件) |
| `flat-stocklist-layout.md` | **新** 平铺库存清单版式(GUCCI A/B、CANADA GOOSE):段落标题行交替、c4 语义判定、颜色列被描述填满、产地占位值、半码归一 |
| `col14-block-layout.md` | **新** col14 块状版式:块结构、反向定位描述行、欧式/美式价格、尺码标签白名单 |
| `futures-channel.md` | **新** 期货通道:数量留空、云仓写法、category-futures 6 版式、openpyxl 读不了的 xlsx |
| `cplus-column.md` | C+ 整数规则 + supply_price + **每单单独给** + 3 层 resolver |
| `type-and-cell-traps.md` | isinstance、灰格 0、'L'/'S'、xlrd 浮点、行号转换 |
| `section-detection-and-row-lifecycle.md` | 多 section 检测、section_col 动态定位、merged_sku |
| `adapter-development.md` | 新供应商开发流程(Step 3.5 SKU 拼接 … Step 3.17 一次性无 V10XX 转换) |
| `EXCEL-xlsx-XML-parsing-guide.md` | 大文件 zipfile+ET |
| `MULTI-FORMAT-XLS-XLSX-ADAPTER.md` | 多格式适配 |
| `OLE2-xls-JPEG-extraction-pattern.md` | OLE2 JPEG 提取 + F010 anchor |
| `PDF-image-extraction-notes.md` | PDF 解析工作流 |
| `XLS-to-XLSX-image-analysis-2026-07-10.md` | XLS→XLSX 图片结构分析 |
| `generic-multi-section-xls-debugging.md` | 多 section xls 调试 |

**品牌专属笔记** (`references/vendors/`): `V0000-MOORER-notes.md`、`V1042-POLO-RALPH-LAUREN-notes.md`。
其余品牌专属列偏移**故意不放**(2026-07-15 决定),按需现场分析。

## 环境变量

| 变量 | 作用 | 默认 |
|---|---|---|
| `CPLUS_BRAND_MAP` | 「品牌对应C+.xlsx」路径 | `~/Desktop/表格处理/6.29/品牌对应C+.xlsx`(存在才用) |
| `TEST_DATA_DIR` | 本地测试数据目录(诊断脚本用) | 仓库相对 `testdata/` |
| `V1067_BASE` | V1067(PANICALE)数据目录 | 仓库相对 `testdata/V1067` |
| `V1042_SRC` / `V1042_CPLUS_RULES` | V1042 C+ 验证脚本的源文件 / 规则文件 | 仓库相对 `testdata/` |
| `TB_OUTPUT_ROOT` | 一次性 TB 转换的输出根目录 | 源文件同级 `converted/` |

## 子 Skill 与本 Skill 关系(2026-07-15 调整)

按"是否带工厂代码"决定是否升级为子 skill:

| 形式 | 条件 | 当前例子 |
|---|---|---|
| **独立子 skill** | 有 V10XX 工厂代码 + 单独仓库 | (暂无) |
| **本 skill 内 references/** | 按品牌(无工厂代码)/ 按结构类型 | FENDI / MONCLER / MIU MIU / BOTTEGA / CHLOE / ALAIA / GUCCI 等 18+ 个 |

所有 vendor 适配器共享本 skill 的 `scripts/convert_vendor_quote_shared/` 包。
