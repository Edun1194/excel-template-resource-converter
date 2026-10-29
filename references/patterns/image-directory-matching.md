# Image Directory Matching(按目录文件名匹配图)

供应商提供图库文件夹,文件名包含 SKU + 色号信息,要按源数据行的 (国际编码, 色号) 找到对应图。

**典型场景**(2026-08-28 PANICALE FW26 接入):

```
供应商给文件夹:
  PANICALE高清图-货号+色号/
  ├── FW26 MEN/
  │   ├── U390010PA 39U36 0200.jpg    ← sku="U390010PA", color="39U36 0200"
  │   ├── U390010PA 39U36 0260.jpg
  │   └── ...
  └── FW26 WOMEN/
      ├── D390010PA 39D49B 0215.jpg
      └── D390010PA-M 39D49B 0935.jpg  ← 同款不同模特 -M 后缀

源数据文件 Sheet1 (颜色级聚合):
  R2: sku="U390010PA" color="39U36 0200" qty=10
  R3: sku="U390010PA" color="39U36 0260" qty=5
  ...
```

源数据每行说"我是哪个 sku + 哪个 color",目录里文件名说"我代表哪个 sku + 哪个 color",**两边都按 (sku, color) 元组精确匹配**。

## 1. 文件名解析规则

**PANICALE / 多数意大利供应商**:`<sku> <color1> <color2>.<ext>` — 3 段,sku 是第 1 段,color 是后 2 段 join 回来。

也兼容 2 段形式:`<sku> <color>.<ext>`(适用于 SKU 后没空格的情况,少见)。

```python
import os
from pathlib import Path

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp"}

def build_image_index(img_dir: Path) -> dict[tuple[str, str], list[Path]]:
    """扫目录,建 (sku, color) → [file_paths] 索引。"""
    idx: dict[tuple[str, str], list[Path]] = {}
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
            continue  # 异常文件名(4+ 段),跳过
        idx.setdefault((sku, color), []).append(img_dir / f)
    return idx
```

**不要**:
- 去掉空格 normalize 后匹配(`U390010PA39U36 0200` 跟 `U390010PA 39U36 0200` 不等价)
- 用正则硬拆(供应商随时换格式,3 段假设可能变)
- 拆 SKU/Color 后用 OR 匹配(可能错配)

## 2. 查找逻辑 — 严格匹配,miss 留空

```python
img_idx = build_image_index(IMG_DIR_MEN)
hit, miss = 0, 0
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, sku_col).value
    color = ws.cell(r, color_col).value
    if not sku or not color:
        continue
    key = (sku, str(color))  # 精确 tuple 匹配
    files = img_idx.get(key, [])
    if not files:
        miss += 1
        continue
    # 嵌入 files[0] (同 key 多张取第 1 张)
    hit += 1
```

**判定**:用户原始需求是"图片命名就是国际编码和色号",**严格匹配不兜底**。miss 行为:
- 图列留空,行还是输出
- 报告 `hit=N miss=M` 统计,让用户知道哪些 SKU 缺图

**可选兜底** — 当图库里有"同款不同模特"变体(命名带 `-M` / `-NN` / `-NM` / `-X` 后缀):

```python
VARIANT_SUFFIXES = ["", "-M", "-NN", "-NM", "-X"]
def find_with_variant(img_idx, sku, color, suffixes=VARIANT_SUFFIXES):
    for suf in suffixes:
        key = (sku + suf, color)
        if key in img_idx:
            return img_idx[key]
    return []
```

PANICALE 案例实测命中率(严格匹配):
- 男 182/272 = 66.9% (252 个 key 覆盖,20 个 SKU 缺图)
- 女 501/863 = 58.1% (652 个 key 覆盖,miss 主要是真没图)

## 3. 嵌入 — 保清晰度模式(OneCellAnchor 改 ext)

详见 `image-handling.md §17`。**核心原则**:**原图 bytes 不动**,只改 openpyxl 写入的 ext cx/cy 让 Excel 按目标 pt 边长显示。

```python
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D

def load_image_keep_bytes(path: Path) -> tuple[XLImage, int, int]:
    """读原图,**不改 bytes**,返回 (XLImage, 原宽 px, 原高 px)。"""
    img = XLImage(str(path))
    return img, img.width, img.height  # openpyxl 已用 PIL 读了 size(像素)

def compute_display_emu(w_px: int, h_px: int, max_pt: float = 92) -> tuple[int, int]:
    """算 (cx, cy) EMU,保持 aspect,最长边 ≤ max_pt。"""
    max_emu = int(max_pt * 12700)
    cx = w_px * 9525
    cy = h_px * 9525
    if max(cx, cy) <= max_emu:
        return cx, cy
    if cx > cy:
        return max_emu, int(cy * max_emu / cx)
    return int(cx * max_emu / cy), max_emu

# 用法
img, w_px, h_px = load_image_keep_bytes(file_path)
cx, cy = compute_display_emu(w_px, h_px, max_pt=92)
img.anchor = OneCellAnchor(
    _from=AnchorMarker(col=image_col - 1, row=r - 1, colOff=0, rowOff=0),
    ext=XDRPositiveSize2D(cx=cx, cy=cy),
)
ws.add_image(img)
ws.row_dimensions[r].height = 92.0  # 行高跟图片最长边一致
```

**绝对不要**:
- `PIL.resize()` 把原图缩到 92pt 边长再存(丢清晰度,文件能小但用户看不清)
- `img.save(JPEG, quality=80)` 重新压缩(同上)

**`img.width/height` 是像素单位**(`image-handling.md §16` 重要 bug 教训),不要 × 9525。

## 4. 输出位置 + 列宽

- **原文件不动** — 用 `wb = openpyxl.load_workbook(src)` 读,在新 sheet (e.g. `wb.create_sheet("Sheet1_with_images")`) 或追加列写图
- **新列 = "图片"** — 写到最后一列后,加 header `图片`,设列宽 92 字符(跟行高 92pt 对齐)
- **行高 92pt** — 所有有图行统一 92pt,无图行沿用源行高
- **不嵌多图** — 同 key 多张图只取第 1 张(`KEEP_IMAGES_PER_ROW = 1` framework 默认)

```python
img_col = ws.max_column + 1
ws.cell(1, img_col, value="图片")
ws.column_dimensions[get_column_letter(img_col)].width = 92.0  # 跟行高对齐
```

## 5. 库存为 0 跳过(2026-08-28 Tory Burch 案例)

**用户偏好**(多次确认):**库存为 0 的行不输出**(不输出占位 row,SKU 整行跳过)。

```python
def _is_positive_qty(v) -> bool:
    if v is None: return False
    if isinstance(v, str):
        s = v.strip()
        if not s or s == "0": return False
    n = _to_float(v)
    return n is not None and n > 0

# 扫 size 列时,所有 size 都 0 → 整个 SKU 跳过
size_rows = []
for c in range(size_col_start, size_col_end + 1):
    qty_raw = ws.cell(r, c).value
    if _is_positive_qty(qty_raw):
        size_rows.append((_size_label(ws.cell(4, c).value), int(_to_float(qty_raw) or 0)))
if not size_rows:
    continue  # 整行跳过,不输出
```

**对比之前**:旧逻辑会输出 1 行 "size=OS qty=0" 占位。**新规则:跳过**。

**与 framework `is_positive_qty` 一致**(`convert_vendor_quote_shared/__init__.py`)。

## 6. 同 SKU 拼 B+C 列(对齐图片命名,2026-08-28)

**用户场景**:源数据 B 列 = 国际编码, C 列 = 色号。用户要求把 B+C 拼成 B 列 = `<sku> <color>`(对齐图片文件命名),C 列清空。

```python
for r in range(2, ws.max_row + 1):
    b = ws.cell(r, 2).value  # 国际编码
    c = ws.cell(r, 3).value  # 色号
    if b and c:
        new_b = f"{b} {c}"  # "U390010PA 39U36 0200"
        if ws.cell(r, 2).value != new_b:
            ws.cell(r, 2).value = new_b
            ws.cell(r, 3).value = None
```

**使用前提**:用户明确要求按"图片命名格式"合并;否则保留 B / C 两列分开的原样。

## 7. 邮件发送大文件 (2026-08-28 实测限制)

阿里云推送 (smtpdm.aliyun.com:465) **附件大小限制 ~10-20MB**:
- 7.3 MB xlsx → ✅ 通过
- 17.5 MB xlsx → ❌ `557 The length of DATA content is achieved the maximum threshold`
- 27 MB 邮件 → ❌ 同上
- 115 MB 保清晰度版 → ❌ 同上

**方案**:
- **小图**(< 5MB) → 直接 `mail_blast.py send` 发单封
- **中大图**(5-50MB) → 用 himalaya 走企业邮箱 `smtp.qiye.aliyun.com:465`,限额更大(测过 27MB 邮件能发)
- **超大图** (>50MB) → `split -b 20m` 分卷,每封 20MB,himalaya 多封发
- **拼清晰度** vs **省体积** → 跟用户确认优先级;默认走"保清晰度"(`OneCellAnchor` 模式)

**VPN 拦截**(常见):macOS Clash fake-IP 模式把 `smtpdm.aliyun.com` DNS 解析到 `198.18.0.x`,SSL 握手直接 EOF。**判定方法**:`nslookup smtpdm.aliyun.com` 看是不是真实 IP(`47.x.x.x`)。如果是 fake-IP,**让用户关 Clash 30 秒再发**。

## 8. 完整脚本骨架

```python
#!/usr/bin/env python3
"""PANICALE FW26 图匹配 — 一次性脚本骨架。"""
import os, sys, uuid
from pathlib import Path
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter

IMG_DIR = Path("/path/to/图库/FW26 MEN")
OUT_DIR = Path("/path/to/output")
SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp"}

def build_image_index(img_dir): ...  # 见 §1
def load_image_keep_bytes(path): ...  # 见 §3
def compute_display_emu(w_px, h_px, max_pt=92): ...  # 见 §3

def process(src_xlsx, img_dir, out_xlsx, sheet_name, sku_col, color_col, img_col):
    wb = openpyxl.load_workbook(src_xlsx)
    ws = wb[sheet_name]
    img_idx = build_image_index(img_dir)
    ws.cell(1, img_col, value="图片")
    ws.column_dimensions[get_column_letter(img_col)].width = 92.0

    hit = miss = 0
    for r in range(2, ws.max_row + 1):
        sku = ws.cell(r, sku_col).value
        color = ws.cell(r, color_col).value
        if not sku or not color: continue
        files = img_idx.get((sku, str(color)), [])
        if not files:
            miss += 1
            continue
        img, w_px, h_px = load_image_keep_bytes(files[0])
        cx, cy = compute_display_emu(w_px, h_px, max_pt=92)
        img.anchor = OneCellAnchor(
            _from=AnchorMarker(col=img_col - 1, row=r - 1, colOff=0, rowOff=0),
            ext=XDRPositiveSize2D(cx=cx, cy=cy),
        )
        ws.add_image(img)
        ws.row_dimensions[r].height = 92.0
        hit += 1
    wb.save(out_xlsx)
    print(f"hit={hit} miss={miss} output={out_xlsx}")
```
