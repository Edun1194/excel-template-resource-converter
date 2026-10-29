# Image Handling Patterns

处理供应商 Excel/Word/PDF 中嵌入图片时反复踩的坑。所有 vendor 适配器共享。

## 1. EMU 命名空间陷阱 (`<a:ext>` vs `<xdr:ext>`)

Drawing XML 中 EMU 尺寸标签有两种命名空间,不同品牌用不同:

| 命名空间 | 出现品牌 | 写法 |
|---|---|---|
| `<xdr:ext cx="..." cy="..."/>` | HOGAN, HOKA LIFESTYLE, Keen Kids, VIVWESTWOOD | Spreadsheet Drawing |
| `<a:ext cx="..." cy="..."/>` | STONE ISLAND, JILSANDER, FENDI, MIU MIU, MONCLER, MAMMUT | DrawingML (`<a:xfrm>` 内) |

两种都是合法 EMU 值,数值可直接用。**先用前者找不到再 fallback 后者**:

```python
# ✅ 兼容两种格式
ext_m = re.search(r'<xdr:ext\s+cx="(\d+)"\s+cy="(\d+)"', block)
if not ext_m:
    ext_m = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"', block)
orig_cx_emu = int(ext_m.group(1)) if ext_m else None
orig_cy_emu = int(ext_m.group(2)) if ext_m else None
```

## 2. 高度上限 92pt 缩放

所有 vendor 输出统一用 92pt 行高。EMU → pt: 1pt = 12700 EMU。

```python
TARGET_CY_EMU = 92 * 12700  # 1,168,400 EMU

def scale_to_92pt(orig_cx, orig_cy):
    if not orig_cy or orig_cy <= 0:
        return (orig_cx, orig_cy)
    if orig_cy > TARGET_CY_EMU:
        scale = TARGET_CY_EMU / orig_cy
        return (int(orig_cx * scale), TARGET_CY_EMU)
    return (orig_cx, orig_cy)
```

> **注意**: 92pt 是行高,不是 124px(124px = 93pt×96/72 是早期错误推导,见 BC SS WOMAN 笔记)。

## 3. `<a:srcRect>` 图片裁剪(百分比,百万分比单位)

FENDI / GUCCI 等 xlsx 经常用 `<a:srcRect>` 裁剪原图,**必须解析后用 PIL crop**,否则显示的是完整原图而非裁剪后的产品图。

```xml
<a:srcRect l="41307" t="22650" r="42702" b="36365"/>
```

单位是**百万分比**(÷ 100000):

```python
m = re.search(r'<a:srcRect[^/]*l="(\d+)"[^/]*t="(\d+)"[^/]*r="(\d+)"[^/]*b="(\d+)"', block)
if m and all(m.groups()):
    l, t, r, b = (int(x) / 100000.0 for x in m.groups())
    pil = PILImage.open(BytesIO(img_blob))
    w_px, h_px = pil.size
    cropped = pil.crop((
        int(w_px * l),
        int(h_px * t),
        int(w_px * (1 - r)),
        int(h_px * (1 - b)),
    ))
    # 用 cropped 实际像素重新计算 display_emu
    cw, ch = cropped.size
    display_emu = (cw * 9525, ch * 9525)
    if ch * 9525 > TARGET_CY_EMU:
        scale = TARGET_CY_EMU / (ch * 9525)
        display_emu = (int(cw * 9525 * scale), TARGET_CY_EMU)
    img_blob = image_to_bytes(cropped, fmt)
```

**空 srcRect 保护**: `<a:srcRect/>`(无属性)时 `m.groups()` 含 `None`,`all()` 返回 False,直接跳过解析,不能 `int(None)`。

## 4. PIL RGBA → RGB 转换

Excel 不正确渲染带 alpha 通道的 PNG,输出前转 RGB(白底):

```python
from PIL import Image as PILImage

def to_rgb_bytes(img_blob, fmt_hint=None):
    pil = PILImage.open(BytesIO(img_blob))
    pil.load()  # ← 强制全解码,捕获 IDAT 损坏
    if pil.mode == 'RGBA':
        bg = PILImage.new('RGB', pil.size, (255, 255, 255))
        bg.paste(pil, mask=pil.split()[3])
        pil = bg
    elif pil.mode != 'RGB':
        pil = pil.convert('RGB')
    buf = BytesIO()
    fmt = fmt_hint or ('jpeg' if pil.format == 'JPEG' else 'png')
    pil.save(buf, format=fmt.upper(), quality=90)
    return buf.getvalue(), fmt
```

**PIL `img.load()` 必加**: 仅 `img.size` 不解压 IDAT,损坏的 PNG 会静默通过 size 检查后写入 xlsx 报错。

## 5. WPS richValue 图片格式(无 drawing XML)

WPS Office 保存的 xlsx 使用专属格式,**无** `xl/drawings/drawing*.xml`,图片通过 `vm=N` 属性映射到 cell:

```
cell vm="N" → xl/richData/richValueRel.xml 第 N 个 <rel r:id="rIdN"/>
            → xl/richData/_rels/richValueRel.xml.rels 的 Target="../media/imageN.png"
```

**适配品牌**: RV (Roger Vivier), GUCCI 部分, ALAIA EYEWEAR

**解析代码**(从 cell XML 直接抓 vm):

```python
import xml.etree.ElementTree as ET
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# 1. 读 richValueRel.xml 拿到 rId 列表
rv_root = ET.fromstring(rich_value_rel_xml)
rids = [rel.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        for rel in rv_root.findall('{...}rel')]

# 2. 读 _rels/richValueRel.xml.rels 拿到 Target → 媒体路径
rels_root = ET.fromstring(rich_value_rel_rels_xml)
id_to_target = {}
for rel in rels_root:
    id_to_target[rel.get('Id')] = rel.get('Target')

# 3. 从 sheet XML 提取 (row, vm_idx)
for cell in ws.iter(f'{NS}c'):
    vm = cell.get('vm')
    if not vm:
        continue
    col_num = col_letter_to_num(''.join(ch for ch in cell.get('r') if ch.isalpha()))
    if col_num != IMAGE_COL:  # 只看图片列
        continue
    row_num = int(''.join(ch for ch in cell.get('r') if ch.isdigit()))
    rid = rids[int(vm)]
    target = id_to_target[rid]  # "../media/imageN.png"
    img_path = 'xl/richData/' + target.replace('../', '')
    img_blob = zip.read(img_path)
    # ...
```

**⚠️ WPS richValue 不提供 EMU 尺寸**,必须用 PIL 读实际像素 + 手动乘 9525:

```python
pil = PILImage.open(BytesIO(img_blob))
w_px, h_px = pil.size
display_emu = (w_px * 9525, h_px * 9525)
if h_px * 9525 > TARGET_CY_EMU:
    scale = TARGET_CY_EMU / (h_px * 9525)
    display_emu = (int(w_px * 9525 * scale), TARGET_CY_EMU)
```

**⚠️ 多图同单元格 hash 去重会失效**: WPS richValue 图片被 PIL 重新编码后 hash 变化,后一张会覆盖前一张。**直接 append,不要 hash 去重**。

## 6. twoCellAnchor 行号 1-indexed ↔ 0-indexed 转换

Drawing XML 用的 Excel 1-indexed 行号,代码内一般用 0-indexed:

```python
from_row_1idx = int(re.search(r'<xdr:from>.*?row="(\d+)"', anchor).group(1))
from_row = from_row_1idx - 1  # → 0-indexed
```

**regex 必须带 `[^>]*`**: `<xdr:twoCellAnchor[^>]*>(.*?)</xdr:twoCellAnchor>`,anchor 标签含 `editAs="oneCell"` 等属性,漏写会 0 match。

## 7. 图片列位置表(各品牌)

| 品牌/格式 | 图片列 (1-indexed) | 命名空间 |
|---|---|---|
| HOGAN BORSE | col 10 (J) | xdr |
| HOGAN DONNA | col 26 (Z) | xdr |
| HOGAN UOMO | col 29 (AC) | xdr |
| HOKA LIFESTYLE | col A (drawing) | xdr |
| FENDI xlsx | col A (drawing) | a:ext |
| FENDI CALZATURE UOMO | col A (drawing) | a:ext |
| MIU MIU | col A (drawing) | a:ext |
| MONCLER GRENOBLLE | col A (drawing) | a:ext |
| RV Calzatura | col 27 (AA) | WPS richValue |
| RV Pelletteria | col 18 (R) | WPS richValue |
| GUCCI | col A (drawing) | a:ext / xdr |
| ALAIA EYEWEAR | col B | WPS richValue (`vm`) |
| MAMMUT ATS | col A | a:ext |
| Keen Kids / Lifestyle | col A | xdr |
| LONGCHAMP | col B (FOTO) | OLE2 JPEG |
| **V1043 PRITELLI** | col A | xdr,但**图片所在 cell 是几行合并**(A11:A13) |
| **V1044 RATTI** | col A | xdr,直接 anchor |
| **V1046 LOSCHI** | col A | xdr,多 sheet 都有 |
| **V1049 POMPEU** | col A | xdr,**每个产品 3-4 张图同 cell** |

具体列号可能因文件版本不同有 ±1 偏差,首次解析时打印 `ws.max_column` 确认。

## 8. 合并单元格图片(V1043 PRITELLI 模式)

**场景**:图片所在 cell 是 N 行合并(如 A11:A13),所有这些行的数据(同一产品的不同尺码)共用这一张图。

**错误做法**:用 `load_sheet_images_by_row` 按 anchor._from.row 分桶 → 只有第一行拿到图,其他行没图。

**正确做法**:用 `load_images_by_merged_range(primary_file, image_col=N)`,它会扫每个 image 的 anchor,找到它所在的 merged range,把图复制到 range 内每一行。

```python
from convert_vendor_quote_shared import load_images_by_merged_range

# V1043 PRITELLI: col A 合并
sheet_images = load_images_by_merged_range(primary_file, image_col=1)
# 返回: {sheet_name -> {row_0idx -> [images]}}
# row_0idx 在 merge 范围 [11, 13] 内,每行都拿到那张图
row_images = sheet_images.get("Foglio1", {}).get(openpyxl_row, [])
```

**坐标系坑**:
- `ws.merged_cells.ranges` 用 **1-indexed** Excel 坐标
- `img.anchor._from` 用 **0-indexed** openpyxl 坐标
- 转换:`excel_col = anchor_col + 1`,`excel_row = anchor_row + 1`

**判定要不要用**:如果源文件 col A(图片列)有大量 merged cell,而且合并区有 2+ 数据行(可用 `len(list(ws.merged_cells.ranges)) >> 50` 启发式判断),就值得用 `load_images_by_merged_range`。

## 9. 多图同单元格(V1049 POMPEU 模式)

**场景**:每个产品 3-4 张图(不同角度)放在同一个 row 的同一图片列,需要在输出 xlsx 里**全部嵌入**。

**错误做法 1**:用 `load_sheet_embedded_images`(V1051 SALOMON 专用,descr-based SKU 匹配)— 会去重漏图,只返回 1 张。

**错误做法 2**:用 `load_sheet_images_by_row` 但只取 `images[0]` — 同样只拿到 1 张。

**正确做法**:用 `load_sheet_images_by_row` 取**全部**(`images_by_row[row_idx]` 是个 list),把它们都放进 `image_items`:

```python
from convert_vendor_quote_shared import load_sheet_images_by_row

sheet_images = load_sheet_images_by_row(primary_file)
# V1049: openpyxl 0-idx row 15 (Excel 16) → 3-4 张图
# 注意:load_sheet_images_by_row 只读第一个 sheet(单 sheet 用)
row_images = sheet_images.get(openpyxl_row, [])  # list of {bytes, format}
# 然后 image_items = row_images(整个 list 都给 dispatcher)
```

**判定要不要用**:检查 `len(ws._images)` 远大于 `max_row`,或者 row 重复出现多图(可用 `Counter({(img.anchor._from.row): count for img in ws._images})` 看有没有 count > 1 的)。

## 10. 图片列宽设置(多图产品必加)

**默认模板的图片列宽是 35 字符**。多图产品(V1049 每产品 3-4 张)放不下,需要在 vendor adapter 里声明 `default_image_col_width` 让 dispatcher 自动加宽。

```python
# vendor_v1049.py
def get_vendor_meta() -> Dict[str, object]:
    return {
        "default_image_col_width": 70.0,  # 35 → 70 字符,放下 3-4 张图
    }
```

**dispatcher 优先级**(跟 image_row_height 一样):
1. CLI `--image-col-width 80` (最高)
2. vendor `get_vendor_meta()["default_image_col_width"]`
3. 模板默认列宽(通常 35)

**列宽经验值**:
- 单图产品 → 35(默认)
- 2-3 张图 → 50
- 3-4 张图 → 70
- 5+ 张图 → 90+
- 行高默认 92pt 不变

## 11. 多图横向排列 colOff 累加 bug ⚠️(2026-07-17 发现)

**症状**:每个产品 3+ 张图,但只有第 1 张在图片列里,第 2/3 张被推到右边很远(列宽外),用户看到图片"跑到后面去了"。

**根因**:`embed_images_in_row` 里 `running_col_off_emu` 累加用的是 `img.width`(原图像素宽度),不是缩放后的显示宽度 `cx`。

```python
# ❌ 错误(2026-07-17 前的代码)
running_col_off_emu += int((img.width + gap_px) * 9525)
# 假设原图 500x500,缩到 92pt(122px)后,这行算成 508px
# 导致 colOff = 508*9525 = 4,838,700 EMU ≈ 508px
# 但实际图只占 122px,中间留了 386px 空白,被推到下一个数据列

# ✅ 正确(2026-07-17 修复后)
running_col_off_emu += cx + int(gap_px * 9525)
# cx 是缩放后的实际显示宽度(EMU),gap 单独算
```

**验证方法**:
```python
# 看一张 500x500 原图、92pt 行高的 colOff:
# ❌ 错误:colOff = 508 * 9525 = 4,838,700 EMU
# ✅ 正确:colOff = 1168400 + 76200 = 1,244,600 EMU (= 122 + 8 px)
```

**判定方法**:`unzip -p` 或 openpyxl 读出来,看同一行 3 张图的 colOff:
- 如果 colOff ≈ img.width × 9525(原图宽)→ 有 bug
- 如果 colOff ≈ cx + gap(显示宽)→ 正确

## 12. 全局只保留第一张图(KEEP_IMAGES_PER_ROW = 1)

**2026-07-17 用户决定**:所有数据行(不管哪个 vendor)都只保留 1 张主图。

**原因**:V1015 Prada 每个产品 3 张图、V1049 POMPEU 3-7 张图,撑爆了图片列(80 字符还不够),且前端的「计算供货价」开关等 UI 元素让多图显示的复杂度变高。

**实现位置**:`convert_vendor_quote_shared/__init__.py` 顶部全局常量 + `write_output()` 里:

```python
KEEP_IMAGES_PER_ROW = 1

# write_output 里:
image_items = r.get("image_items", [])[:KEEP_IMAGES_PER_ROW]
```

**改回多图**:把 `KEEP_IMAGES_PER_ROW` 改大(2-3),并在 vendor 的 `get_vendor_meta()` 里把 `default_image_col_width` 加到对应宽度(每张图占 ~28 字符)。

**所有 vendor 改动后对照表**:
| Vendor | 多图数 | 改后嵌入 | 减少 |
|---|---|---|---|
| V1015 Prada Italy | 3 张/产品 | 1 张/产品 | 128→46 |
| V1049 POMPEU | 3-7 张/产品 | 1 张/产品 | 4160→1161 |
| V1044 RATTI | 1 张/产品 | 1 张/产品 | 4724 (不变) |
| V1043 PRITELLI | merged 范围共享 | 1 张/产品 | 2407 (不变) |

## 13. 模板列宽覆盖 — 货号列默认 26(2026-07-17)

**问题**:模板默认 货号列宽 = 13 字符,放不下长 SKU(如 V1004 的 `211870245002-NEWPORT NAVY` 25 位)。

**解决**:`convert_vendor_quote_shared/__init__.py` 顶部全局常量 + `write_output()` 里:

```python
SKU_COL_WIDTH = 26.0  # 13 → 26 翻倍,够放 25 位货号 + 颜色后缀

# write_output 里(在 image_col_width 之后):
sku_col_idx = headers.index("货号") + 1 if "货号" in headers else (
    headers.index("International code") + 1 if "International code" in headers else None
)
if sku_col_idx is not None and SKU_COL_WIDTH > 0:
    sku_col_letter = get_column_letter(sku_col_idx)
    ws.column_dimensions[sku_col_letter].width = float(SKU_COL_WIDTH)
```

**适用所有 vendor**:在 write_output 统一改,无需 vendor 自己设置。

**调整列宽(其他列)**:同样的 pattern,在 `write_output` 里加一段设置 col_dimensions 即可。推荐封装为 `set_col_width(ws, headers, col_name, width)`。

## 14. `editAs` 属性让 shared regex 0 match → 124 张图全丢(V1019 MCM 模式,2026-07-27)

**症状**:源文件 xlsx 内有 100+ 张图(`ws._images` 能数到),但 `load_sheet_embedded_images` 返回 0 images,输出 xlsx 一张图都没有。

**根因**:MCM(MONCLER 也部分踩到)生成的 drawing XML 用的是 `<xdr:twoCellAnchor editAs="oneCell">`,而 `convert_vendor_quote_shared._load_drawing_images_via_xml` 里的旧 regex 是:

```python
re.findall(r'<xdr:twoCellAnchor>(.*?)</xdr:twoCellAnchor>', xml, re.DOTALL)
```

这个 regex 不接受 `<xdr:twoCellAnchor>` 后跟任何属性,碰到 `editAs="..."` 直接 0 match。

**修复**(`shared.py:_load_drawing_images_via_xml`):

```python
# 兼容 <xdr:twoCellAnchor editAs="oneCell"> 等带属性的形式
anchors = re.findall(r'<xdr:twoCellAnchor\b[^>]*>(.*?)</xdr:twoCellAnchor>', xml, re.DOTALL)
```

加 `\b` 防误匹配,加 `[^>]*` 吃属性。

**验证方法**(碰到 0 images 时跑一下):

```python
import zipfile, re
with zipfile.ZipFile("vendor.xlsx") as z:
    xml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
print("contains 'twoCellAnchor':", xml.count("twoCellAnchor"))  # 应 > 0
print("with editAs attr:        ", len(re.findall(r'<xdr:twoCellAnchor\b[^>]*>', xml)))
print("without attr:            ", len(re.findall(r'<xdr:twoCellAnchor>', xml)))
# 若含 editAs 的 count > 不含的 count,说明有 anchor 带属性,要用新 regex
```

**影响范围**:任何 anchor 带 `editAs`/`editAs="oneCell"` 等属性的 vendor 都会踩到(实测 V1019 MCM 100% 踩,MONCLER 14/24 踩)。修复后这些 vendor 全部恢复正常。

**适配器定位**:V1019 触发本 bug,已在 `shared.py` 修,无需 vendor 单独 workaround。

## 15. Excel 365 `richValue` 机制 — `drawing1.xml` 的图是错的 (2026-07-29)

**症状**:
- 源文件用 `openpyxl` 读到 `ws._images` 有图,PIL 抽出来是 *A 类图*(如运动鞋/boxer 短裤)
- 但在 Excel/WPS 里打开 *同样的 xlsx*,单元格显示的是 *B 类图*(如靴子/T-shirt)
- 用户说"我转换后表格里的图,我源文件里都找不到一样的"

**根因**(典型场景:Vivienne Westwood 等大公司从 SharePoint 同步下来的 xlsx):

xlsx **有两套图机制**,Excel 365 同时存在:

| 机制 | 文件 | openpyxl 读得到 | Excel 实际显示 |
|---|---|---|---|
| 老的 `drawing1.xml` + 1320x1800 大图 | `xl/drawings/drawing1.xml` + `xl/media/image198.jpg` 等 | ✅ | ❌(被覆盖) |
| 新的 `richValue` linked value + 84x112 缩略图 | `xl/richData/*.xml` + `xl/media/image1.png` 等 | ❌(openpyxl 不支持) | ✅ |

`workbook.xml` 里的 `<x15ac:absPath url="https://*.sharepoint.com/..."/>` 标记文件来自 SharePoint 同步,这种文件 90% 用 richValue。

**检测**(碰到 2 套图大小不一,或者 PIL 读出来的图跟 Excel 截图对不上时,跑一下):

```python
import zipfile
fp = "vendor.xlsx"
z = zipfile.ZipFile(fp)
# 1) 有没有 richData 目录?
has_richdata = any("xl/richData/" in n for n in z.namelist())
print(f"has richData: {has_richdata}")
# 2) drawing 几张图
draw_imgs = sum(1 for n in z.namelist() if n.startswith("xl/media/image") and n.endswith(".jpg"))
png_imgs = sum(1 for n in z.namelist() if n.startswith("xl/media/image") and n.endswith(".png"))
print(f"jpg imgs: {draw_imgs}  png imgs: {png_imgs}")
# 3) sheet1.xml 看 E 列 cell 是否有 vm 属性
sheet1 = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
import re
vm_count = len(re.findall(r'<c r="E\d+"[^>]*?vm="\d+"', sheet1))
print(f"E 列带 vm 的 cell: {vm_count}")
# 全 True → richValue 机制在工作,要用下面这个 load 函数
```

**解决**:`load_sheet_images_by_row` 跳过 richValue,要用专门走 richValue 路径的加载器。

```python
import re
import zipfile
from pathlib import Path

def load_richvalue_images(fp: Path) -> dict[int, list[dict]]:
    """Excel 365 richValue 机制:从 cell.vm → rdrichvalue → richValueRel → image file。

    返回: {row_1idx -> [{"bytes": ..., "format": "png"|"jpeg"}]}

    流程:
    1. sheet1.xml: E 列 cell 的 vm=N           (1-idx 索引到 rdrichvalue)
    2. rdrichvalue.xml: rvs[vm-1]              (vm 是 1-idx,rvs 是 0-idx)
       = (_rvRel:LocalImageIdentifier=M, CalcOrigin=...)
    3. richValueRel.xml: rels_list[M]          (0-idx 索引,直接是 rId 编号-1)
       = "rIdN"
    4. richValueRel.xml.rels: rIdN → "../media/imageK.png"
    5. 读 xl/media/imageK.png
    """
    z = zipfile.ZipFile(fp, "r")
    sheet1 = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    rd = z.read("xl/richData/rdrichvalue.xml").decode("utf-8")
    rels = z.read("xl/richData/richValueRel.xml").decode("utf-8")
    rels_targets = z.read("xl/richData/_rels/richValueRel.xml.rels").decode("utf-8")

    # 1) E 列 cell → vm
    e_vm: dict[int, int] = {}
    for m in re.finditer(r'<c r="E(\d+)"[^>]*?vm="(\d+)"', sheet1):
        e_vm[int(m.group(1))] = int(m.group(2))

    # 2) rdrichvalue 全部 rv
    rvs = re.findall(r'<rv s="\d+"><v>(\d+)</v><v>(\d+)</v></rv>', rd)
    # 3) richValueRel 全部 rel
    rels_list = re.findall(r'<rel r:id="(rId\d+)"\s*/>', rels)
    # 4) rId → media file
    rid_to_target: dict[str, str] = {}
    for m in re.finditer(
        r'<Relationship Id="(rId\d+)"[^>]*Target="\.\./(media/[^"]+)"', rels_targets
    ):
        rid_to_target[m.group(1)] = m.group(2)

    out: dict[int, list[dict]] = {}
    for row_1idx, vm in e_vm.items():
        # ★ 关键:vm 是 1-idx,rvs 是 0-idx,差 1
        rv_idx = vm - 1
        if not (0 <= rv_idx < len(rvs)):
            continue
        lid = int(rvs[rv_idx][0])  # LocalImageIdentifier
        if not (0 <= lid < len(rels_list)):
            continue
        rid = rels_list[lid]
        target = rid_to_target.get(rid)
        if not target:
            continue
        try:
            data = z.read(f"xl/{target}")
        except KeyError:
            continue
        fmt = "png" if target.lower().endswith(".png") else "jpeg"
        out[row_1idx] = [{"bytes": data, "format": fmt}]
    return out
```

**用法**(`load_sheet_images_by_row` 的 richValue 替代):

```python
# 旧(走 drawing,SharePoint 同步文件拿到的是旧 cache 图)
imgs_by_row = load_sheet_images_by_row(fp)  # 64 行,大 jpg 1920x1800

# 新(走 richValue,Excel 实际显示这个)
imgs_by_row = load_richvalue_images(fp)     # 440 行,小 png 84x112

# 后续同 c3 共享代表行图、embed_images_in_row 都一样
```

**验证正确性**(抽 R2 跟源文件比 md5):

```python
import zipfile, hashlib
fp = "vendor.xlsx"
z = zipfile.ZipFile(fp)
print("R2 真图 md5:", hashlib.md5(z.read("xl/media/image1.png")).hexdigest()[:8])
# 比对 converted 文件 image1.png md5,应一致
```

**踩坑记录**:
1. **vm 是 1-idx 不是 0-idx** — E2 cell 的 vm=1 对应 rvs[0],`vm-1` 一定要写。
2. **rvs 是 0-idx 列表** — `re.findall(...)` 拿到的 list 索引从 0 开始。
3. **rels_list 也是 0-idx** — `rels_list[lid]` 而非 `rels_list[lid+1]`。
4. **richValue 的图都是小缩略图**(~2-3KB png 84x112),不是 drawing 那种 100KB+ 大图,所以最终 converted 文件会从 94MB 缩到 3MB。
5. **drawing1.xml 的图是历史/旧 cache**,跟 richValue 不一致时,*以 richValue 为准*(Excel 实际显示 richValue)。

**影响范围**:任何从 Microsoft 365 / SharePoint / OneDrive 同步下来的 xlsx,且 1) 在 Excel 里看着有图 + 2) 描述/货号跟 drawing 抽出来的图对不上时,都触发本机制。代表案例:Vivienne Westwood ATS VW 文件(446 行,440 行 richValue 图)。

**通用策略建议**:以后新 vendor 适配器,**先跑 `load_richvalue_images` 检测一下** — 命中就用 richValue,没命中 fallback 到 `load_sheet_images_by_row`。可考虑在 `convert_vendor_quote_shared/__init__.py` 加个 `load_sheet_images_smart(fp)`,自动选路径。

### 15.1 前置判断:每行一张图 vs 同 SKU 共享一张图(2026-07-29)

**问题**:从 richValue 拿到图后,要不要 *按 (c3 / Generic Prod Code) 共享代表行图*?错案例:ATS VW V1015 补货单,T-shirt 同 c3 (`81060008-J001O-`) 不同色码 (-061/-062/-063/-064/-066/-067) 原表每行都配了不同图,我代码按 c3 共享,导致 6 行全用同一张图,用户投诉"所有颜色都是同一张图"。

**判断方法**(在用图之前跑一下):

```python
# 假设你已经 load_richvalue_images 拿到 rich_imgs: {1-idx row -> [{bytes, format}]}
# 看 *每个 SKU 前缀* 的图 md5 是不是都不同
from collections import defaultdict
import hashlib

sku_prefix_to_md5s = defaultdict(set)
for r_1idx, items in rich_imgs.items():
    if not items: continue
    md5 = hashlib.md5(items[0]["bytes"]).hexdigest()[:8]
    # 假设 c3 是 Generic Prod Code 列 (0-idx 2)
    sku = ws_in.cell(r_1idx, 3).value
    if not sku: continue
    sku_prefix = str(sku).strip()  # 用整个 c3 当 key
    sku_prefix_to_md5s[sku_prefix].add(md5)

# 同一个 SKU 前缀的图 md5 数
duplicated = {k: v for k, v in sku_prefix_to_md5s.items() if len(v) == 1}
print(f"同 SKU 只用 1 张图: {len(duplicated)}")
print(f"同 SKU 配了多张图: {len(sku_prefix_to_md5s) - len(duplicated)}")
```

**结论**:
- 如果 *同 SKU 多张图* → 走 **per-row** 模式,每行用自己的图(不要共享)
- 如果 *同 SKU 1 张图* → 走 **shared** 模式(代表行图给所有同 SKU 用,省 converted 文件大小)

ATS VW V1015 实测:**同 c3 多色码都是不同图**,必须 per-row。错误共享后 R3-R8 全用 image2.png(md5=8f4a19a6),R4-R8 实际该是 image3-image7(900d961d / 4db87412 / 8c92bbba / d71a7335 / 791ef621)。

**正确的 per-row 写法**:

```python
# ★ 不要做 sku_to_rep 共享索引
# 直接每行拿自己的图
for r in range(2, ws.max_row + 1):
    img_items = rich_imgs.get(r, [])  # r 是 1-idx
    rows.append({..., "image_items": img_items})
```

**通用建议**:
- 新 vendor 适配器写完,先跑上面的判断脚本
- 如果是 *shared* 模式(同一 SKU 多行共用一张图)→ 共享省空间,推荐
- 如果是 *per-row* 模式(同一 SKU 多行不同图)→ 不能共享,必须每行自己拿
- 现在 ATS V1015 = per-row,实测

## 16. ⚠️ `openpyxl.Image.width/height` 单位是**像素**(不是 EMU) — 2026-08-28 踩坑

**症状**:明明给 `img.width = new_w` 赋了像素值,Excel 打开后图片巨大(几个屏幕放不下),有时 79ft 高这种"天上一脚地上一脚"。

**根因**:`openpyxl.Image.width/height` 的单位是**像素**(底层自动 × 9525 转 EMU)。

```python
# ❌ 错误(2026-08-28 我犯的):把 new_w 当 EMU 用,又乘 9525
img.width = int(new_w * 9525)  # 实际 EMU = px × 9525²,巨大
# 假设 new_w = 82 → EMU = 82 × 9525 = 780450(看上去对)
# 但 openpyxl 读出 width 居然 = 82(不是 780450)→ 它真把它当 px 存
# 当 openpyxl 写 ext 时,实际写出 cx = 780450 × 9525 = 7,435,072,500 EMU ≈ 779 inches

# ✅ 正确:直接给像素值
img.width = new_w
img.height = new_h
```

**为什么之前"行得通"**:openpyxl 在序列化时,`img.width` 会被当作"显示宽度像素"读出来再乘 9525 写 ext。所以你写 82 像素,实际 ext 是 82 × 9525 = 780450 EMU = 82 px = 61.5 pt。

**验证方法**(调试时跑一下):

```python
import zipfile, re
with zipfile.ZipFile("out.xlsx") as z:
    drawing = z.read("xl/drawings/drawing1.xml").decode()
exts = re.findall(r'<ext\s+cx="(\d+)"\s+cy="(\d+)"', drawing)
for cx, cy in exts[:3]:
    print(f"cx={cx} cy={cy} EMU = {int(cx)/9525:.0f}×{int(cy)/9525:.0f} px = {int(cx)/12700:.1f}×{int(cy)/12700:.1f} pt")
# 如果写 img.width=82 但 EMU 出来是 cx=780450 (= 82 × 9525)= 正确
# 如果 EMU 出来是 cx=7,435,072,500 (= 82 × 9525²)= 你 double-multiplied 了
```

**判定规则**:
- 写 `img.width = X` → 实际 ext `cx = X * 9525` EMU
- 想要 cx = 92pt = 1,168,400 EMU → 写 `img.width = 92` (像素)
- 想要 cx = 200 EMU → 写 `img.width = 200 / 9525` (反推)
- **永远不要写 `img.width = X * 9525`**

**配套的 PIL 输出尺寸**:PIL `img.size` 返回 `(w_px, h_px)`,跟 openpyxl 像素单位一致,直接传。

## 17. OneCellAnchor + XDRPositiveSize2D 改 display 尺寸,保留原图 bytes(保清晰度,2026-08-28)

**场景**:用户要求"图片显示 92pt 但**保持清晰度不变**" — 原图是 1280×1920 高清,不能缩到 92pt 后再压缩(会糊)。

**错误做法**:`PIL.resize` + `img.save(JPEG, quality=80)` 把原图缩到 92pt 边长 → 1.4KB/张,糊(用户 2026-08-28 反馈"完全看不清")。

**正确做法**:**原图 bytes 一字不动**,只改 openpyxl 写入的 ext cx/cy 让 Excel 按 92pt 显示(Excel 内部高质量缩放渲染)。

```python
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D

img = XLImage(str(original_path))  # 不动 bytes
# img.width/height 已经从 PIL 读到了原图 px 尺寸(不需要算 EMU)
# §16 说 img.width = px,openpyxl 内部 × 9525 → 写入 ext = px × 9525 EMU

# 但 §17 模式:不依赖 img.width/height 的隐式 × 9525,显式 OneCellAnchor
# 算出我们想要显示的 cx,cy (EMU,按 aspect 比例 + 限最长边 ≤ 92pt)
PX_TO_EMU = 9525
PT_TO_EMU = 12700
max_emu = int(92 * PT_TO_EMU)  # 92pt = 1,168,400 EMU
orig_cx = w_px * PX_TO_EMU
orig_cy = h_px * PX_TO_EMU
if max(orig_cx, orig_cy) <= max_emu:
    cx, cy = orig_cx, orig_cy
elif orig_cx > orig_cy:
    cx, cy = max_emu, int(orig_cy * max_emu / orig_cx)
else:
    cx, cy = int(orig_cx * max_emu / orig_cy), max_emu

img.anchor = OneCellAnchor(
    _from=AnchorMarker(col=image_col - 1, row=r - 1, colOff=0, rowOff=0),
    ext=XDRPositiveSize2D(cx=cx, cy=cy),
)
ws.add_image(img)
# 关键:不调 img.width/height,OneCellAnchor 的 ext 才是最终控制
```

**对比两种模式**:

| 模式 | 文件大小 | 清晰度 | 适用 |
|---|---|---|---|
| 缩图后嵌入 (PIL + JPEG) | 1-3 MB | 糊(损失) | 邮件发送 / 体积限制 |
| OneCellAnchor 改 ext (保 bytes) | 几百 MB | 100% 原图 | 本地查看 / 不限体积 |

**frame_work 的 `embed_images_in_row` 默认就是 OneCellAnchor 模式**(bytes 不变,只改 ext),符合"清晰度保留"原则。**不要随便改它去做 PIL resize**,除非用户明确说"清晰度无所谓"。

## 18. 按目录文件名的图片匹配(`(sku, color) → file` 索引,2026-08-28)

**场景**:供应商提供图库文件夹,文件名包含 SKU 和颜色信息,要按源数据行的 (国际编码, 色号) 找到对应图。

**文件名规则**(PANICALE FW26 案例,2026-08-28 接入):

```
PANICALE高清图-货号+色号/
├── FW26 MEN/
│   ├── U390010PA 39U36 0200.jpg    ← <sku> <color>.<ext>
│   ├── U390010PA 39U36 0260.jpg
│   └── ...
└── FW26 WOMEN/
    ├── D390010PA 39D49B 0215.jpg
    └── D390010PA-M 39D49B 0935.jpg  ← 同款不同模特带 -M 后缀
```

**通用解析规则**:
- `<sku>` 是第一段(空格前)
- `<color>` 是后面所有段(空格 join 回来)
- 3 段形式最常见:`<sku> <color1> <color2>.<ext>` → `sku = parts[0]`, `color = ' '.join(parts[1:])`
- 2 段形式:`<sku> <color>.<ext>` → `parts[0]` 和 `parts[1]`

**索引构建**(精确匹配,不放宽变体):

```python
SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp"}

def build_image_index(img_dir: Path) -> dict[tuple[str, str], list[Path]]:
    idx = {}
    for f in os.listdir(img_dir):
        ext = os.path.splitext(f)[1].lower()
        if ext not in SUPPORTED_EXTS:
            continue
        base = os.path.splitext(f)[0]
        parts = base.split(" ")
        if len(parts) == 2:
            sku, color = parts
        elif len(parts) == 3:
            sku, color = parts[0], " ".join(parts[1:])
        else:
            continue
        idx.setdefault((sku, color), []).append(img_dir / f)
    return idx
```

**查找逻辑**:

```python
img_idx = build_image_index(img_dir)
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, sku_col).value
    color = ws.cell(r, color_col).value
    if not sku or not color: continue
    key = (sku, str(color))  # 精确匹配
    files = img_idx.get(key, [])
    if not files:
        # miss:图片列留空,行还是输出
        continue
    img = embed_image(files[0], image_col=r, max_pt=92)
    # ...
```

**判定要不要做 miss 兜底**(变体 -M / -NN):

- 如果图库有同款不同模特图(命名带 `-M` / `-NN` / `-NM` / `-X` 后缀),但源数据行的色号没后缀
- **严格按用户要求**:"图片命名就是国际编码和色号" → 严格匹配不兜底
- **如果想兜底**:试 `(sku, color)` → `(sku + "-M", color)` → `(sku + "-NN", color)` 等

**PANICALE FW26 实测命中率**:
- 男 182/272 = 66.9% (252 个 key 覆盖)
- 女 501/863 = 58.1% (652 个 key 覆盖)
- miss 的 SKU 主要是"真没图"(图库没收录对应 SKU+色号),不是匹配逻辑问题

**特殊场景 — 同 SKU 拼 B+C 列**(2026-08-28 Tory Burch WAREHOUSE A 案例):

源文件 B 列 = 国际编码(`U390010PA`), C 列 = 色号(`39U36 0200`)。**用户要求把 B+C 拼成 B 列 = `<sku> <color>`**(对齐图片命名),C 列清空。

```python
for r in range(2, ws.max_row + 1):
    b = ws.cell(r, 2).value
    c = ws.cell(r, 3).value
    if b and c:
        ws.cell(r, 2).value = f"{b} {c}"  # "U390010PA 39U36 0200"
        ws.cell(r, 3).value = None
```

**新文件**:这种"按目录文件名匹配"的模式抽到独立文件 `image-directory-matching.md`,见那里更详细(全 SKU 遍历、跳过 0 inventory 行、嵌入原图保留清晰度等)。

## 19. ⚠️ POLO 等供应商源图是"产品 + swatch 组合图",自动裁剪 swatch (2026-09-12)

**问题**:POLO RALPH LAUREN 的 .xls 文件里,每件产品的图本身就是"产品照 + 下方 1-2 个色卡 (swatch)"组合图(281x281 / 375x375 缩略图版式)。如果直接嵌到 92pt 缩略图,看起来像"产品 + 一条彩色色块",用户觉得是损坏。

**根因**:POLO 在源 .xls 里直接塞了 swatch 组合图,跟图损坏无关,是源数据设计。需要源文件给 ALINA 改导出才能根治,或者在 helper 里自动裁掉 swatch。

**自动 swatch 裁剪算法** (`_strip_bottom_swatch`,在 `convert_vendor_quote_shared/__init__.py` 里全局可用):

```python
def _strip_bottom_swatch(jpeg_bytes: bytes) -> bytes:
    """POLO 源图常见"产品 + 下方 swatch"组合图,自动检测并裁掉 swatch,补白回 1:1。
    """
    from PIL import Image
    import io as _io
    import numpy as _np
    pi = Image.open(_io.BytesIO(jpeg_bytes)).convert("RGB")
    w, h = pi.size
    if w < 80 or h < 80:
        return jpeg_bytes  # 太小不处理

    arr = _np.array(pi).astype(_np.int32)
    quant = (arr // 32).astype(_np.int32).reshape(h, -1, 3)
    row_uniques = _np.array([len(_np.unique(quant[r], axis=0)) for r in range(h)])

    # 从下往上,找连续 ≥8 行 unique ≤ 8 的"纯色条带"(swatch 特征)
    consec = 0
    for r in range(h - 1, h // 3, -1):
        if row_uniques[r] <= 8:
            consec += 1
        else:
            break
    if consec < 8:
        return jpeg_bytes  # 没检测到 swatch,保留原图

    swatch_start = h - consec
    # 白色背景过滤:swatch 区域接近白色 → 是纯产品照的空白底,不裁
    swatch_region = arr[swatch_start:h]
    avg_color = swatch_region.mean(axis=(0, 1))
    is_white_bg = avg_color.min() > 240 and swatch_region.std() < 30
    if is_white_bg:
        return jpeg_bytes

    # 裁 + 补白 1:1
    cropped = pi.crop((0, 0, w, swatch_start))
    if swatch_start < h * 0.4:
        return jpeg_bytes  # 裁过头(< 40% 留产品),放弃
    padded = Image.new("RGB", (w, h), (255, 255, 255))
    padded.paste(cropped, (0, 0))
    buf = _io.BytesIO()
    padded.save(buf, format="JPEG", quality=92)
    return buf.getvalue()
```

**算法 4 步**:
1. 量化 RGB 到 32 阶(降噪声),每行算 unique color 数
2. 从下往上扫,连续 ≥8 行 unique ≤ 8 → "swatch 纯色条带"
3. 该条带主色不是白色(纯产品照的空白底) → 才是 swatch
4. 裁 + 补白回 1:1,质量 92 reencode

**实测效果** (POLO RALPH LAUREN R-75 .xls):
- 紫白条纹衬衫 (281x281): 裁到 78%,留产品照 + 白色 padding
- 国旗毛衣 (375x375): 裁到 59%,留产品照 + 白色 padding
- 纯产品照 500x500: 不裁(白色背景被识别)

**何时调用**:在 `load_xls_ole2_jpeg_images` / `load_sheet_images_by_row` 配对 JPEG[i] ↔ F010[i] 后,return 前过一遍 `_strip_bottom_swatch(jpeg_bytes)`。POLO 等用 .xls + OLE2 JPEG 流的 vendor 自动受益;.xlsx 走 openpyxl 路径同样适用。

**已知限制**:
- 仅裁"底部 1 个 swatch 条带",POLO 还有"产品 + 底部品牌色装饰条 + 1 个色卡 swatch + padding"复杂版式,装饰条裁不到(只能在产品图上方看装饰)
- 裁后 reencode quality=92,文件略增大但清晰度保留
- 281x281 / 375x375 这类小尺寸源图裁后产品会变小(到 92pt 缩略图后细节丢失);用户看到 500x500 大图嵌进 92pt 的清晰度比小尺寸裁后好
