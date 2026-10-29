# Multi-Format / Multi-Layout Linesheet Adapter (2026-07-01)

## 背景

6.29 批次处理过程中发现：同一个适配器需要同时支持：
1. `.xls`（OLE2 binary，BOTTEGA VENETA 等大多数品牌）
2. `.xlsx`（GUCCI MAN 系列，xlrd 不支持）
3. 不同的列偏移布局（GUCCI WOMAN SHOES 的 section code 在 col17 而非 col13）

## 关键一：openpyxl 兼容（xlsx 支持）

xlrd 不支持 .xlsx 文件。GUCCI MAN 全系列（RTW/BAGS+SOFT/SHOES）都是 .xlsx。

**解决方案：`_OpenpyxlWorksheet` 统一接口**

```python
from pathlib import Path
import xlrd, openpyxl

def load_xls(path: Path):
    suffix = path.suffix.lower()
    if suffix == '.xlsx':
        wb = openpyxl.load_workbook(path, data_only=True)
        # Handle sheet names like 'SS27 (2)'
        sheet_name = next((sn for sn in wb.sheetnames if 'SS27' in sn), wb.sheetnames[0])
        return _OpenpyxlWorksheet(wb[sheet_name])
    else:
        return xlrd.open_workbook(path, formatting_info=True)

class _OpenpyxlWorksheet:
    """Wrap openpyxl to expose xlrd-compatible interface: cell_value(r, c), nrows, ncols"""
    def __init__(self, ws):
        self._ws = ws
        self.nrows = ws.max_row
        self.ncols = ws.max_column
    def cell_value(self, row_idx: int, col_idx: int):
        # xlrd is 0-indexed; openpyxl is 1-indexed
        v = self._ws.cell(row_idx + 1, col_idx + 1).value
        return v if v is not None else ''
```

parse_rows 中使用：
```python
wb = load_xls(primary_file)
ws = wb if hasattr(wb, 'nrows') else wb.sheet_by_index(0)
```

## 关键二：动态 section code 列检测

不同文件的 section code 列位置不同：

| 文件类型 | Section Code 列 | 说明 |
|----------|----------------|------|
| 大多数文件 | xlrd col13 (openpyxl col14) | 标准位置 |
| GUCCI WOMAN SHOES | xlrd col17 (openpyxl col18) | 列偏移 |

```python
def _get_section_code_col(ws) -> int:
    """检测 section codes 在哪一列"""
    for r in range(min(5, ws.nrows)):
        for col in (13, 17):
            v = ws.cell_value(r, col)
            if v and str(v).strip() in ('A2', 'V', 'F', 'FF', 'TU', 'L', 'UKD', 'Q'):
                return col
    return 13
```

`_find_section_header_row` 也需要同时检测 col13 和 col17：
```python
def _find_section_header_row(ws, section_code: str) -> int:
    for r in range(min(5, ws.nrows)):
        for col in (13, 17):
            n_val = ws.cell_value(r, col)
            if n_val and str(n_val).strip() == section_code:
                return r
    return 1
```

## 关键三：动态列布局检测

GUCCI WOMAN SHOES 的列偏移：
- 标准：desc=col5, sku=col6, cost=col10, retail=col11
- SHOES 偏移：desc=col9, sku=col10, cost=col14, retail=col15
- category：标准 col4，偏移 col8

**检测逻辑：**
```python
desc_col5 = str(ws.cell_value(r, 5) or '').strip()
sku_col6 = str(ws.cell_value(r, 6) or '').strip()
desc_col9 = str(ws.cell_value(r, 9) or '').strip()
sku_col10 = str(ws.cell_value(r, 10) or '').strip()

if desc_col5 or sku_col6:
    # 标准布局
    description = desc_col5
    sku = sku_col6
    cost = to_float(ws.cell_value(r, 10))
    retail = to_float(ws.cell_value(r, 11))
    category_col4 = str(ws.cell_value(r, 4) or '').strip()
    category_col8 = str(ws.cell_value(r, 8) or '').strip()
else:
    # 偏移布局（GUCCI SHOES）
    description = desc_col9
    sku = sku_col10
    cost = to_float(ws.cell_value(r, 14))
    retail = to_float(ws.cell_value(r, 15))
    category_col8 = str(ws.cell_value(r, 8) or '').strip()
    category_col4 = ''
```

category fallback 链：col1 → col4 → col8（GUCCI SHOES 偏移）

## 6.29 批次验证结果

| 品牌/文件 | 格式 | 行数 | 状态 |
|-----------|------|------|------|
| BOTTEGA VENETA DONNA | .xls | 17 | ✅ |
| BOTTEGA VENETA UOMO | .xls | 33 | ✅ |
| SALOMON MAIN SS27 | .xls | 192 | ✅ |
| STONE ISLAND MAIN SS27 | .xls | 48 | ✅ |
| CHLOE DONNA RTW | .xls | 28 | ✅ |
| FENDI ABBIGLIAMENTO DONNA | .xlsx | V(17尺码:32起) + A2(10尺码) + F | ✅ 38行 |
| FENDI ABBIGLIAMENTO UOMO | .xlsx | Q/F/V/A2 混合section | ✅ 42行 |
| FENDI BORSE+ACCESSORI DONNA | .xlsx | TU + L(皮带11尺码) | ✅ 61行 |
| FENDI BORSE+ACCESSORI UOMO | .xlsx | TU | ✅ 11行 |
| FENDI CALZATURE DONNA | .xlsx | FF(数字34起) | ✅ 175行 |
| FENDI CALZATURE UOMO | .xlsx | UKU(鞋码) | ✅ 82行 |
| Gucci WOMAN BAGS+ACCESSORIES | .xls | 91 | ✅ |
| Gucci WOMAN SHOES | .xls | 160 | ✅ |
| Gucci MAN RTW | **.xlsx** | 49 | ✅ |
| Gucci MAN BAGS+SOFT | **.xlsx** | 32 | ✅ |
| Gucci MAN SHOES | **.xlsx** | 19 | ✅ |
| MIU MIU RTW | .xls | 95 | ✅ |
| Torry Burch DONNA RTW | .xls | 31 | ✅ |
| Torry Burch DONNA SHOES | .xls | 133 | ✅ |

**总计：19 文件，1307 行**

## 待处理（2026-07-09）

- ✅ FENDI CALZATURE UOMO — 已解决（UKU section）
- MONCLER GRENOBLLE DONNA/UOMO — 列偏移+4，需单独分析

## 适配器文件位置

各品牌文件夹内独立适配器：
```
/Users/superlist/Desktop/表格处理/6.29/BOTTEGA VENETA/convert_BOTTEGA_VENETA.py
/Users/superlist/Desktop/表格处理/6.29/SALOMON/convert_SALOMON.py
/Users/superlist/Desktop/表格处理/6.29/Stone Island/convert_STONE_ISLAND.py
/Users/superlist/Desktop/表格处理/6.29/CHLOE/convert_CHLOE.py
/Users/superlist/Desktop/表格处理/6.29/FENDI/convert_FENDI.py
/Users/superlist/Desktop/表格处理/6.29/Gucci/convert_GUCCI.py
/Users/superlist/Desktop/表格处理/6.29/MIU MIU/convert_MIU_MIU.py
/Users/superlist/Desktop/表格处理/6.29/Torry Burch/convert_TORRY_BURCH.py
```

所有适配器均已同步 xlsx 支持 + Q section code + 动态列布局检测。