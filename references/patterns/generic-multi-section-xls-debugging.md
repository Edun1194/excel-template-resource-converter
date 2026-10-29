# 6.29 批次 Linesheet 调试笔记（2026-06-30）

## 调试技巧

### 用 zipfile 直接读取 xlsx XML 验证输出数据

当 openpyxl 读取转换后 xlsx 出现异常（如描述/类目字段混淆）时，直接读取内部 XML 绕过 openpyxl：

```python
import zipfile, re

files = glob.glob("/path/to/converted/*.xlsx")
with zipfile.ZipFile(files[-1]) as z:
    # 直接读 sheet1 XML
    sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    # 用正则快速找某行某列的值
    # <c r="D2" ...><v>...</v></c> = D列 row2
```

比 openpyxl 更可靠——不走 Python object 映射。

### xlrd 输出验证

xlrd 的 `ws.cell_value(r, c)` 是 **xlrd 0-indexed 列**。xlrd col 4 = Excel 列 E（1-indexed）。

关键字段 xlrd 索引（通用，适用于所有 6.29 xls）：
- `xlrd col 1` = Excel B = 类别
- `xlrd col 4` = Excel E = 品类（CATEGORIE）
- `xlrd col 5` = Excel F = 产品名（DESCRIZIONE）
- `xlrd col 6` = Excel G = STYLE（SKU）
- `xlrd col 10` = Excel K = COSTO FORNITORE（cost）
- `xlrd col 11` = Excel L = PREZZO RETAIL（零售价）
- `xlrd col 13` = Excel N = Section code（FF/F/FR/A2/V/TU/L/USD 等）

## category/description 字段映射（已解决）

### 原始问题
GUCCI RTW 转换后 `描述` 列全是产品名（正确），`类目` 列全是品类名（正确），但中间某次调试时误以为描述/类目反了——实际上是工具误读。

### 根因
误读是因为调试工具的输出格式问题。正确的数据流：
```
description = xlrd.col5 (DESCRIZIONE) → 写入 描述 列（Excel D列）
category    = xlrd.col4 (CATEGORIE)   → 写入 类目 列（Excel E列）
```
模板的 `描述` 在 D 列，`类目` 在 E 列（中文表头时）。

### generic_labels 修复

BOTTEGA VENETA 的 B 列（col1）存品类名，但 FENDI、CHLOE、GUCCI 的 B 列有时是通用组名（如 `RTW NOV`、`LOOK EXTRA`），不是真实类目。

修复逻辑：
```python
generic_labels = {
    'LOOK EXTRA', 'LOOK EXTRA (UNISEX)', 'RTW NOV', 'RTW',
    'LOOK 80', 'LOOK 77', 'RTW EXTRA'
}
if category_col1 and category_col1.upper() not in generic_labels:
    category = category_col1        # BOTTEGA: B列即为类目
elif category_col4:
    category = category_col4        # CHLOE/GUCCI: 类目在 col4
else:
    category = category_col1        # fallback
```

## TORRY BURCH 特殊结构

### RTW：A2 + P2 混合 section

TB RTW 的 section code 在 col13（xlrd）：
- A2 section：尺寸在 row 3 cols 14-23（文本尺码 3XS~4XL），数据行 col1='RTW NOV' 或 'SWIM NOV'
- P2 section：P2 数据行在 col14 开始有 qty 数据（col16=1 即 size 37），尺寸为 row 2 col14+ 的 P2 行

P2 数据行特征：col13='P2'，col14+ 有 qty 数据。

**通用 BOTTEGA 适配器只能处理 A2 section**。P2 需要单独适配器或扩展。

### SHOES：USD section + 混合文本/数字尺码

TB SHOES 的 section code 是 'USD'（货币单位，不是尺码），实际尺码在 row 2 cols 14-32，格式混合：
- 整数：35, 36, 37...（xlrd float: 35.0）
- 半码：3.5, 4.5, 5.5...（xlrd string: '3.5'）
- 总计 19 个尺码列

尺寸标签来自 row 2：`str(35.0)` → `'35.0'` → `'35'`（小数点后为 0 时去掉）

## MONCLER 列偏移警告

MONCLER GRENOBLLE 文件（DONNA/UOMO）的列结构比其他品牌偏移 +4：
- Section code 在 col17（其他品牌在 col13）
- COSTO 在 col14（其他品牌在 col10）

不能用通用适配器，需单独分析 `xlrd ncols + 列内容扫描`。
