# OLE2 XLS JPEG Thumbnail Extraction Pattern

多个品牌（ALAIA、CHLOE 等）的 .xls 文件将产品图片以 JPEG 缩略图形式嵌入 OLE2 Workbook 流中，
不通过标准 MSODRAWING/OfficeArt 格式，而是连续存储 JPEG 二进制数据。

## 适用文件特征

- 文件后缀: `.xls`（xlrd 可读）
- OLE2 流: 只有一个 `Workbook` 流含图片数据
- 图片特征: 136×181 px JPEG 缩略图（每产品一张）
- **判断方法**: 用 olefile 打开，看 Workbook 流大小是否合理（几百 KB 到几 MB），然后搜索 `b'\xff\xd8\xff'` JPEG signature

## 提取算法（基础版）

```python
import olefile, io, re
from PIL import Image as PILImage

def _extract_jpegs(xls_path: str) -> List[bytes]:
    """扫描 Workbook 流 JPEG SOI→EOI 标记，返回 JPEG blob 列表"""
    ole = olefile.OleFileIO(xls_path)
    wb_data = ole.openstream('Workbook').read()
    ole.close()
    jpegs = []
    pos = 0
    while True:
        start = wb_data.find(b'\xff\xd8\xff', pos)
        if start == -1: break
        end = wb_data.find(b'\xff\xd9', start)
        if end == -1: break
        jpegs.append(wb_data[start:end + 2])
        pos = start + 1
    return jpegs
```

## ⚠️ 顺序陷阱：JPEG 数量 ≠ 产品数量

**问题**：37 个产品但只有 34 张 JPEG，顺序映射会导致错位（无图 SKU 导致后续图片上移）。

**解决**：解析 MSODrawing F010 Client Anchor 获取每张图片的精确行位置。

## ⚠️ JPEG spurious EOI 切短陷阱（2026-09-12 POLO 实战）

**问题**：源 .xls 的 JPEG 数据流里偶尔有 spurious `\xff\xd9` 字节序列当数据（非（不是 marker），基础版 `wb_bytes.find(b'\xff\xd9', start)` 找到第一个 EOI 就切，导致某张图切短，后续所有图错位。

**验证方法**：
```python
soi_count = wb_bytes.count(b'\xff\xd8\xff')
eoi_count = wb_bytes.count(b'\xff\xd9')
# soi_count == 122, eoi_count == 123 → 多 1 个 EOI 是 spurious
```

**定位**：找哪些图区间内有 >1 个 EOI（spurious 都在 SOI[i] 和 SOI[i+1] 之间）。
```python
for i in range(len(soi_pos)):
    eois_between = [x for x in eoi_pos if soi_pos[i] <= x < soi_pos[i+1]]
    if len(eois_between) > 1:
        # JPEG[i] 数据流里有 spurious EOI,基础切法切短
        print(f'JPEG[{i}]: {len(eois_between)} EOI in range')
```

**解决方案**（POLO R-75 用过）：每张图切到下一张 SOI 之前，让 PIL 自动 trim 到真 EOI：
```python
# 不用第一个 EOI 切,改用 [SOI[i], SOI[i+1]) 切,bytes 包含整个区间
# PIL decode 时会自动跳过 spurious EOI,找到真 EOI 并 trim
for i in range(len(soi_pos)):
    start = soi_pos[i]
    end = soi_pos[i + 1] if i + 1 < len(soi_pos) else len(wb_bytes)
    jpeg = wb_bytes[start:end]
    pi = Image.open(io.BytesIO(jpeg))  # PIL 内部 trim 到真 EOI
    # ...
```

**更稳的方案**：维护 JPEG 数据流状态机，正确处理 escape bytes (`\xff\x00`) 和 RST marker (`\xff\xd0` ~ `\xff\xd7`)，但实现复杂。如果切到 next SOI 够用，不必上状态机。

## ⚠️ xlrd MSODrawing 大文件漏解 + raw scan fallback (2026-09-12)

**问题**（POLO R-75 .xls, 1.2MB, 124 个产品）：xlrd `handle_msodrawingetc` monkey-patch 只能解前 ~28 个 F010 Client Anchor，剩下的 96 个 anchor 字节全在 wb stream 里但被 xlrd 跳过。

**根因**：xlrd 解析 MSODrawing record 时，可能因为 BIFF 结构异常 / record 边界错位，中断扫描后续 MSODrawing 块。

**验证**：
```python
# xlrd 能解 28 个 F010
# 但 raw scan F010 signature 能解 124 个
import struct
sig = b"\x10\xf0\x12\x00\x00\x00"  # F010 + cb=18
positions = [i for i in range(len(wb_bytes)) if wb_bytes[i:i+6] == sig]
print(f'raw F010: {len(positions)}, xlrd F010: {len(xlrd_anchors)}')
```

**F010 Client Anchor 实际 layout (BIFF8,POLO 现货单实测)**：
- **6 字节 header** = F010 sig(2 bytes `10 f0`) + cb(4 bytes, 值=18)
- **18 字节 data** = move(2) + col_lo(4) + row_lo(4) + col_hi(4) + row_hi(4)

> ⚠️ 之前 memory / 早期笔记误以为 header 是 8 字节，实际只有 6 字节（sig 2 + cb 4）。raw scan 偏移用 `idx + 6`，不是 `idx + 8`。

**Raw scan fallback 代码** (优先用 raw,raw 解不到 0 个才退回 xlrd):
```python
import struct
sig = b"\x10\xf0\x12\x00\x00\x00"
anchors = []
cursor = 0
while True:
    idx = wb_bytes.find(sig, cursor)
    if idx == -1:
        break
    cursor = idx + 1
    if idx + 6 + 10 > len(wb_bytes):
        continue
    try:
        move, col_lo, row_lo = struct.unpack('<Hii', wb_bytes[idx + 6:idx + 6 + 10])
    except Exception:
        continue
    if 0 <= col_lo <= 15 and 0 <= row_lo <= 3000:  # sanity
        anchors.append((col_lo, row_lo))

if not anchors:
    # 回退到 xlrd monkey-patch（兼容老文件）
    import xlrd, xlrd.sheet as sheet_mod, xlrd.biffh
    # ... 原来的 monkey-patch 代码
```

**调用时机**：在 `load_xls_ole2_jpeg_images` 里，先 raw scan 解 F010；raw 解不到才走 xlrd fallback。这样 .xls > 1MB / 大产品数 (>50) 都能正确配图。

## F010 Client Anchor 解析（精确映射）

**核心发现**：xlrd 2.0.2 的 `handle_msodrawingetc` 内部解析 MSODrawing 子记录，但默认不暴露 API。需 monkey-patch 捕获 F010 数据。

**代码**：
```python
import xlrd, xlrd.sheet, xlrd.biffh

anchors = []  # [(sheet_num, row_lo, col_lo)]

original = xlrd.sheet.Sheet.handle_msodrawingetc

def patched(self, recid, data_len, data):
    pos = 0
    while pos < data_len:
        tmp, fbt, cb = xlrd.biffh.unpack('<HHI', data[pos:pos+8])
        ver = tmp & 0xF
        ndb = cb if ver != 0xF else 0
        if fbt == 0xF010 and ndb == 18:
            # F010 Client Anchor: 18 bytes = unk(2) + col_lo(4) + row_lo(4) + col_hi(4) + row_hi(4)
            vals = xlrd.biffh.unpack('<Hiiii', data[pos+8:pos+8+ndb])
            anchors.append((self.number, vals[2], vals[1]))  # (sheet, col_lo, row_lo)
        pos += ndb + 8
    return original(self, recid, data_len, data)

xlrd.sheet.Sheet.handle_msodrawingetc = patched

wb = xlrd.open_workbook(xls_path, formatting_info=True)
ws = wb.sheet_by_index(0)
_ = ws.nrows  # force sheet load
xlrd.sheet.Sheet.handle_msodrawingetc = original  # restore
```

**F010 unpack 关键**：
- F010 Client Anchor 整体 = **6 字节 header** (`10 f0` + `12 00 00 00` cb=18) + **18 字节 payload**
- payload: `<Hiiii` = move(2) + col_lo(4) + row_lo(4) + col_hi(4) + row_hi(4)
- `vals[2]` = `col_lo`（列），`vals[1]` = `row_lo`（行）
- **不是** `vals[0]` = row，`vals[1]` = col
- Raw scan 偏移用 `idx + 6`（不是 `idx + 8`，之前 memory 写错过）

**构建映射**：
```python
# 过滤图片列（col==2 = B 列）
image_anchors = [(row, col) for (sheet, row, col) in anchors if col == 2]

# JPEG[n] → anchor[n] → product row = anchor_row + 1
jpeg_to_product = {}  # {jpeg_idx: xlrd_product_row}
for jpeg_idx in range(len(all_jpegs)):
    if jpeg_idx < len(image_anchors):
        prod_row = image_anchors[jpeg_idx][0] + 1
        if prod_row in product_rows:
            jpeg_to_product[jpeg_idx] = prod_row
```

**⚠️ 关键陷阱**：
- `range(len(jpegs))` 而非 `range(len(anchors))` — anchors 含空白锚点，比 JPEGs 多 3 个
- 用 `range(len(anchors))` 会导致 jpeg_idx 越界，映射为空

## entry_idx 映射规则（基础版，用于简单顺序映射场景）

| 品牌/文件 | 产品行位置 | entry_idx 计算 | image 数组下标 |
|-----------|-----------|----------------|----------------|
| ALAIA DONNA RTW | xls rows 5,10,15... | `entry_idx = len(entries)`（按 entry 累积计数） | `rtw_row_images.get(entry_idx + 2, [])`（+2 是因为 Excel 1-indexed + header） |
| CHLOE DONNA RTW | xls rows 5,10,15...（每 5 行一个产品） | `entry_idx = xls_row // 5`（直接除法） | `row_images[entry_idx]`（直接用数组下标，无 offset） |

**ALAIA 需要 +2 offset 的原因**: `_extract_rtw_images()` 用 `output_row = xls_product_row // 5 + 1`（Excel 1-indexed）作 key，而 `_load_rtw()` 中 `entry_idx` 从 0 开始，所以要用 `entry_idx + 2` 转换。

**CHLOE 直接用数组下标**: `_extract_chloe_images()` 返回的 list 本身就是按 image 出现顺序排列，`entry_idx = xls_row // 5` 恰好等于 image 数组下标，无需额外 offset。

## ⚠️ embed_images_in_row 顺序陷阱

**`write_output` → `embed_images_in_row` 会把 `image_items` 按输出行顺序依次 placement**：
- 如果 entry[0] 有 3 个尺码（展开 3 行），entry[1] 有 4 个尺码（展开 4 行）
- 那么 2 张图片会被 placement 到 **7 个输出行**（第 1 张在 R2-R4，第 2 张在 R5-R8）
- 同一 product 的多个尺码行**都会收到同一张图片**（img_items 被复制到每行）

**根本原因**：`embed_images_in_row` 没有"每产品只放一张"的逻辑，按输出行顺序依次处理。

**解决方案（2026-07-02 验证）**：
- 如果用户确认顺序对应是正确的（每产品首行一张图，后续尺码行也显示同一张图）→ 当前实现正确
- 如果用户要求同一 product 只在**第一尺码行**显示图片 → 需要在适配器层控制：`entry_idx` 变化时才传 `img_items`，同 product 不同尺码行传 `img_items=[]`

**CHLOE 验证结论（2026-07-02）**：用户确认每行都要图片，9 个产品 × 各自尺码数 = 35 张图片，实际每行都显示该 product 的缩略图，**顺序对应正确**。

## 验证步骤

1. 用 `olefile.OleFileIO(xls_path)` 确认文件可打开
2. 搜索 `b'\xff\xd8\xff'` 确认有 JPEG 数据
3. 提取并用 PIL 打开，验证 size 和内容
4. 统计提取到的缩略图数量，与产品行数对比
5. 输出转换后 xlsx，用 openpyxl 读取 `ws._images` 验证图片数量
6. **关键验证**：解压 xlsx 的 `xl/drawings/drawing1.xml`，用 `ET.fromstring` + `xdr:oneCellAnchor` 解析每张图的 anchor row，对照 SKU 验证映射是否正确

## LONGCHAMP 实例（2026-07-09）

- 文件: `LONGCHAMP FW26 - WAREHOUSE A(1).xls`
- 37 产品，34 JPEG，37 anchors（含 3 个空白/表头锚点）
- 34 JPEG → 34 anchors（含头尾空白锚点）→ 映射后 33 个有图产品 + 4 个无图 SKU
- 无图 SKU：34175089510, L1621089510, L1626089510, L2605089002
- 图片内容：淡粉/浅紫背景色（avg_rgb 166,163,178）是 LONGCHAMP 品牌原图风格，**非损坏**；JPEG 结构正常（SOI/EOI 标记完整）
