# Cell Type & Section Code Traps

适配器开发中最容易踩的几个类型/字符串判断陷阱,跨品牌反复出现。

## 1. `isinstance(v, (int, float))` 必须,不要用 `isinstance(v, float)`

xlrd 和 openpyxl 读整数行为不同:
- **xlrd** 读整数返回 `float`(`32.0`)
- **openpyxl** 读整数返回 `int`(`32`)

```python
# ❌ 错:openpyxl 读整数返回 int,被 isinstance(float) 漏掉 → FENDI V section 整数尺码消失
if isinstance(v, float) and v > 0:
    label = str(int(v))

# ✅ 对:同时覆盖 int 和 float
if isinstance(v, (int, float)) and v > 0:
    if v == int(v):
        label = str(int(v))   # 32.0 → "32"
    else:
        label = str(v)        # 35.5 → "35.5"
```

**影响品牌**: FENDI V section(整数尺码 32, 34, 36...)、GUCCI SHOES FF section、MONCLER TB section、BOTTEGA VENETA V section。

**唯一例外**: 灰格(grey cell)判断里 `v == 0`,`int(0) == 0` 和 `float(0.0) == 0` 都成立,安全。

## 2. 数字 `0` 灰格陷阱

xlrd 灰格返回 `0.0`(或 `0`),`isinstance` 判断会被误认为合法尺码标签。

```python
# ❌ 错:0 会被当作尺码 "0"
if isinstance(v, (int, float)) and v > 0:
    return str(int(v)) if v == int(v) else str(v)

# ✅ 对:0 单独判断返回 None
if isinstance(v, (int, float)):
    if v == 0:
        return None  # 灰格 / unavailable cell
    return str(int(v)) if v == int(v) else str(v)
```

**影响品牌**: MONCLER TB section header row、HOGAN UOMO 4 号段(尺寸 4, 4½ 灰格)。

## 3. `'L'` 和 `'S'` 不应是 section code

A2 section 的尺码标签包含单字母尺码 `S`、`M`、`L`、`XL`。如果把这些加进 `section_codes` 集合,`_get_size_label` 会跳过这些尺码。

```python
# ❌ 错:把 'L' (皮带 section) 和 'S' (A2 尺码) 都当 section code
SECTION_CODES = {'FF', 'TU', 'A2', 'V', 'F', 'FR', 'UKD', 'L', 'Q', 'TB', 'S'}

# ✅ 对:只放真正的 section 标识符
SECTION_CODES = {'FF', 'TU', 'A2', 'V', 'F', 'FR', 'UKD', 'Q', 'TB', 'L'}
# ↑ 'L' 保留(皮带 section),但要确保同时允许 A2 section 把 'L' 当尺码
# ↑ 'S' 移除(是 A2 尺码,不是 section code)
```

**特别测试**:
- A2 section: 3XS / 2XS / XS / **S** / M / **L** / XL / 2XL / 3XL / 4XL 全部要能识别
- L section: 皮带尺码 50, 55, 60...120(数字,不会与字符串 'L' 冲突)

**影响品牌**: BOTTEGA VENETA(2026-07-02)、MONCLER GRENOBLLE(2026-07-13)、GUCCI BAGS+SOFT。

## 4. xlrd 浮点尺码格式化

xlrd 返回的 size label 是 `float`,写进 `size_col_map` 时已转为字符串 `"24.0"`。

```python
label_str = str(size_label).strip()
try:
    fval = float(label_str)
    if fval == int(fval):
        label_str = str(int(fval))  # "24.0" → "24"
    # else: "35.5" 保留
except ValueError:
    pass  # 文本尺码 "3XS", "L" 等
```

## 5. xlrd `Totale` 字符串陷阱

xlrd 首次访问某 cell 时若包含非数字字符串会触发 `ValueError: could not convert string: 'Totale'`。try/except 包装:

```python
def safe_cell_value(ws, row, col):
    try:
        v = ws.cell_value(row, col)
        return v
    except ValueError:
        return None
```

**影响品牌**: ALAIA 行 0 col 36 字符串 `'Totale'`。

## 6. 共享字符串前导空白

xlsx 的 shared string 可能含前导 `\t` / 空格,读出来需 `.strip()`:

```python
sku = str(ws.cell_value(r, sku_col)).strip()  # ALAIA EYEWEAR col3 含前导 tab
```

## 7. 行号 1-indexed ↔ 0-indexed

| 来源 | 行号 |
|---|---|
| Excel UI / openpyxl `.cell(row=)` / `ws.max_row` | 1-indexed |
| xlrd `cell_value(row, col)` | 0-indexed |
| openpyxl `.cell(row=, column=).row` | 1-indexed |
| openpyxl `image.anchor._from.row` | **0-indexed** |
| xlsx XML `<row r="...">` / `drawing <xdr:from row="...">` | 1-indexed |
| xlrd `ws.nrows` | 0-indexed count |

```python
# xlrd row → openpyxl row
openpyxl_row = xlrd_row + 1

# openpyxl row → xlrd row
xlrd_row = openpyxl_row - 1

# drawing XML row → 0-indexed
from_row_0 = int(re.search(r'row="(\d+)"', xml).group(1)) - 1
```

## 8. `ws.max_row` 在 `read_only=True` 模式下返回 None

`openpyxl.load_workbook(..., read_only=True)` 时 `ws.max_row` 是 `None`,必须用其他方式确定行数:

```python
# 大文件(>5000行)不要用 read_only,会超时 → 见 xlsx-large-file-parsing.md
# 小文件可以用普通模式,max_row 正常

# 或用 zipfile + ET 解析 sheetN.xml 数 row 元素
```

**影响品牌**: MAX MARA(small,普通模式)、MAMMUT(7700+ 行,zipfile+ET)。

## 9. `'UKD'` vs `'UKU'` section code

SALOMON 用 `UKD`(混合尺码),FENDI CALZATURE UOMO 用 `UKU`(鞋码)。**不要混淆**——`_get_size_label` 不会自动转换,误判会让尺码表错位。
