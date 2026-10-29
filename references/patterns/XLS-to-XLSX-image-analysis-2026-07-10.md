# XLS→XLSX 转换后图片分析（2026-07-10）

## 背景

用户将多个品牌的 `.xls` 文件转换为 `.xlsx` 后测试转换器。转换器之前在 `.xls` 模式下正常运行，但 `.xlsx` 源文件结构与预期不同。

## 关键发现

### 1. 转换后 xlsx 的图片结构可能变化

LibreOffice/Excel 将 `.xls` 转为 `.xlsx` 时：
- 嵌入的 OLE2 图片（xls 原生格式）→ 丢失或损坏
- 原 `.xlsx` 的 drawing XML 保持 twoCellAnchor 格式

**调试方法**：转换前先用 zipfile 检查源文件是否有 `xl/media/` 和 `xl/drawings/`：
```python
import zipfile, re
with zipfile.ZipFile(path) as z:
    media = [n for n in z.namelist() if re.match(r'xl/media/.*\.(png|jpg|jpeg)', n)]
    drawings = [n for n in z.namelist() if re.match(r'xl/drawings/drawing\d+\.xml$', n)]
    print(f'媒体文件: {len(media)}, Drawing文件: {len(drawings)}')
```

### 2. 有图片 vs 无图片的品牌（2026-07-10 检测结果）

| 品牌 | 文件 | 有图片 |
|------|------|--------|
| MIU MIU | XLSX | ✅ 16张 |
| SALOMON | XLSX | ✅ 21张 |
| Stone Island | XLSX | ❌ 0张 |
| Torry Burch | XLSX | ❌ 0张 |
| VIVIENNE WESTWOOD | XLSX | ✅ 6张 |
| MONCLER | XLSX | ✅ 17张 |

**结论**：Stone Island 和 Torry Burch 的 xlsx 文件完全没有内嵌图片（LibreOffice 转换时丢失）。这两个品牌的转换器不需要 `_load_xlsx_images()` 函数。

### 3. twoCellAnchor row range 分发：一个锚点 → 多行

FENDI xlsx 的 twoCellAnchor 跨越多行（如 `from_row=6, to_row=12`），一个锚点覆盖一个 LOOK 的所有尺码行。

`_load_xlsx_images()` 将图片分发到 row range 内所有产品行：
```python
for row in range(from_row_0idx, to_row_0idx + 1):
    if row not in row_images:
        row_images[row] = []
    row_images[row].append(item)
```

**结果**：源文件图片数 ≠ 输出锚点数。
- MIU MIU: 16张源图 → 137个锚点（每张图分发到多行）
- MONCLER: 17张源图 → 97个锚点

### 4. MONCLER 非标准结构

MONCLER 的 section code 在 **col 20**（而非标准的 col 13），尺码 label header 在 Row 3/4 的 **col 21+**，完全独立于标准 Linesheet 格式。

**关键字段（1-indexed）**:
- `col 9` = 类目 (CATEGORIE)
- `col 10` = 描述 (DESCRIZIONE)
- `col 11` = SKU (STYLE)
- `col 15` = COSTO FORNITORE
- `col 16` = PREZZO RETAIL
- `col 20` = section code (A2/TB)
- `col 21+` = 尺码 qty

### 5. VIVIENNE WESTWOOD 列偏移

VIVIENNE WESTWOOD 的 section code 在 **col 14**（0-indexed），尺码从 **col 15+** 开始，与标准 col 13 不同。

## 调试清单

当接手新的品牌转换器时：
1. 先用 zipfile 检查源文件是否有图片
2. 用 openpyxl 打印前 6 行确认列布局
3. 找到 section code 所在列（通常 col 13 或 col 14/20）
4. 确认尺码 label header 行（通常 row 1/2/3）
5. 对照 GUCCI/FENDI 的 `_load_xlsx_images()` 确认图片锚点类型（twoCellAnchor vs oneCellAnchor）
