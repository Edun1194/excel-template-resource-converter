# Section Detection & Row Lifecycle Patterns

多 section linesheet 解析的通用规则(品牌文件结构多变,核心逻辑要稳)。

## Section header 识别

典型 linesheet 结构(行 0-4 = metadata/headers,行 5+ = 数据):

```
Row 0: 品牌元数据
Row 1: section code (如 "A2", "UKD") + 数字尺码标签
Row 2: section code (如 "FR", "A2") + 文本尺码标签
Row 3: section code (如 "V") + 数字尺码标签
Row 4: 列头 (FOTO, N° FOTO, MADE IN, COMPOSIZIONE...)
Row 5+: 数据行(可能每 N 行一个产品)
```

**Section code 必须在数据行也出现**(在 col 13 或 18 等固定列),数据行的 section code 决定该行使用哪组尺码。

## 自动发现 Section Header Row

```python
def _find_section_header_row(ws, section_code, expected_row=None):
    # 先尝试预期行
    if expected_row is not None:
        for col in range(13, ws.ncols):
            v = ws.cell_value(expected_row, col)
            if str(v).strip() == section_code:
                return expected_row

    # Fallback:扫描 rows 0-4
    for row_idx in range(5):
        for col in range(13, min(ws.ncols, 40)):
            v = ws.cell_value(row_idx, col)
            if str(v).strip() == section_code:
                return row_idx
    return expected_row
```

## build_size_col_map(section_code, section_row)

```python
def build_size_col_map(ws, section_code, section_row, fallback_rows=(0,1,2,3,4)):
    size_col_map = {}

    # 1. 从 section_row 自身找(部分 section code 与尺码同行)
    for col in range(14, ws.ncols):
        label = _get_size_label(ws, section_row, col)
        if label:
            size_col_map[col] = label

    # 2. Fallback:扫描其他行找尺码
    if not size_col_map:
        for row_idx in fallback_rows:
            if row_idx == section_row:
                continue
            for col in range(14, ws.ncols):
                v = ws.cell_value(row_idx, col)
                if str(v).strip() == section_code:
                    continue
                label = _get_size_label(ws, row_idx, col)
                if label:
                    size_col_map[col] = label
            if size_col_map:
                break

    return size_col_map
```

## section_code 列定位陷阱

**section code 列不固定**:
- 标准 linesheet: col 13(N)
- FENDI CALZATURE UOMO: col 18(S)
- MONCLER GRENOBLLE DONNA: col 19(T)
- MONCLER GRENOBLLE UOMO: col 13(N)
- VIVIENNE WESTWOOD: col 14(O)
- GUCCI SHOES: col 17(Q)

**正确做法**: 扫描 rows 5+ 同时检查多个 col,取实际有 recognized section codes 的列。

```python
def detect_section_col(ws, sheet_idx, max_row=30):
    known = {'FF','TU','A2','V','F','FR','UKD','L','Q','TB','UKU'}
    best_col = None
    best_hits = 0
    for col in (13, 14, 17, 18, 19, 20):
        hits = 0
        for r in range(5, min(max_row, ws.nrows)):
            v = ws.cell_value(r, col)
            if v and str(v).strip() in known:
                hits += 1
        if hits > best_hits:
            best_hits = hits
            best_col = col
    return best_col
```

## 数据行 section code 残留 bug

当某行 `n_val` 为空但 `n` 没被重置,下一次循环会用旧值触发 section 误切换:

```python
# ❌ 错:用原始 n_val 判断
for r in range(...):
    n_val = ws.cell_value(r, section_col)
    if not n_val:
        continue
    n = str(n_val).strip()
    # n 仍是旧值!

# ✅ 对:用处理后的 n 判断
for r in range(...):
    n_val = ws.cell_value(r, section_col)
    n = str(n_val).strip() if n_val else ''
    if not n:
        continue
```

## 数据列偏移不一致

**MONCLER GRENOBLLE 案例**:
- DONNA: category=col 8, desc=col 9, sku=col 10, ...
- UOMO:  category=col 4, desc=col 5, sku=col 6, ...

**必须 if/else 硬编码两组列索引**,不要用统一偏移公式。

## HOGAN 多 section 行高合并问题

HOGAN UOMO 有 6 个 section header(行 2/31/46/66/145/149),数据段长度不一。**`merged_sku` 集合必须在每个 section 处理时重建为空 set**,防止跨 section 同名 SKU 被错误合并。

```python
for idx, (sec_row, size_col_map, sec_wholesale_col) in enumerate(section_headers):
    merged_sku = set()  # ← 必须 per-section 独立
    data_start = sec_row + 1
    data_end = section_headers[idx + 1][0] - 1 if idx + 1 < len(section_headers) else ws.max_row
    for r in range(data_start, data_end + 1):
        ...
```

## 第一段数据丢失陷阱

若第一个 detected section header 在 `header_row + 1` 之后,**header_row 和第一个 section header 之间的数据会被静默跳过**。修复:在 `not section_headers` 的 else 分支把 `header_row` 作为第一个 section 插入。

## size 文本识别 fallback

xlrd 读 `"PZ"`(单件)会返回字符串 `"PZ"`,读数字尺码返回 float。先判断 string 排除非尺码文本:

```python
NON_SIZE_TEXTS = {'TOTALE', 'TOT', 'TOTAL', 'CLIENTE', 'PZ'}  # PZ 不是 size matrix,单件默认 1 行

def _get_size_label(ws, row_idx, col_idx):
    v = ws.cell_value(row_idx, col_idx)
    if v is None or v == '':
        return None
    label_str = str(v).strip()
    if not label_str or label_str.upper() in NON_SIZE_TEXTS:
        return None
    if label_str.upper() in SECTION_CODES:
        return None
    ...
```

## TU qty 计算

`TU` (Taglia Unica) 没有尺码矩阵,单行单 SKU。qty 直接读 col 14+(限制 cols 14-31,跳过 > 100 的总计列):

```python
total = 0
for col in range(14, 31):
    qty = ws.cell_value(r, col)
    if qty and float(qty or 0) > 100:
        continue  # 总计列
    total += max(0, int(float(qty or 0)))
```

## 多 sheet 独立输出

SALOMON MAIN SS27 等文件每个 sheet 是不同产品类别(FTW/ADV7/App & Acc/BAGS & PACKS/SOCKS),**每个 sheet 独立输出一个 xlsx 文件**:

```python
for sheet_name, handler in sheet_handlers.items():
    rows = handler(ws, sheet_images, vendor_id)
    if rows:
        output_path = output_dir / f"V1051-{sheet_name}-template1-cn-converted.xlsx"
        write_output(template_path, output_path, rows)
return {}  # 返回空 dict,主调度器不会重复处理
```

调度器不支持多 sheet 输出,需**直接调用适配器**:

```bash
cd "供应商报价单目录"
python3 convert_V1051.py "ORDER FORM SALOMON MAIN SS27.xlsx" "V1051"
```
