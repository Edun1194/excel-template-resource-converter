# Adapter Development Workflow

新供应商报价单 → 转换脚本的标准流程。

## Step 0: 探明文件结构

```bash
# 检查文件类型和 sheet
python3 -c "
import xlrd
wb = xlrd.open_workbook('vendor.xlsx', formatting_info=True)
for s in wb.sheets():
    print(f'sheet={s.name!r} rows={s.nrows} cols={s.ncols}')
    for r in range(min(5, s.nrows)):
        print(f'  row {r}: {[s.cell_value(r, c) for c in range(min(15, s.ncols))]}')
"
```

**判断要点**:
- `.xls` → xlrd 读取(可能含 OLE2 JPEG)
- `.xlsx` → openpyxl 读取(可能含 WPS richValue 图片)
- sheet 数量 / 每个 sheet 的 metadata 行结构
- 数据行布局(每 N 行一个产品,还是单行单 SKU)

## Step 1: 探明图片结构

```python
# xlsx
import zipfile
with zipfile.ZipFile('vendor.xlsx') as z:
    for name in z.namelist():
        if 'media' in name or 'drawing' in name or 'richData' in name:
            print(name)

# xls
import olefile
ole = olefile.OleFileIO('vendor.xls')
print(ole.listdir())
if ole.exists('Workbook'):
    wb_data = ole.openstream('Workbook').read()
    print(f'Workbook stream size: {len(wb_data)}')
    print(f'JPEG count: {wb_data.count(b"\\xff\\xd8\\xff")}')
```

**判断图片格式**:
- 标准 `xl/drawings/drawingN.xml` + `xl/media/image*.png/jpeg` → `twoCellAnchor` 或 `oneCellAnchor`
- WPS `xl/richData/richValueRel.xml` + `xl/richData/_rels/...` → richValue `vm=N` 模式
- OLE2 `Workbook` 流里 `\xFF\xD8\xFF` 连续存储 → 顺序提取(用 F010 anchor 校正)
- 无图片文件 → `image_items = {}`,直接跳过图片处理

## Step 2: 验证模板必填字段

主模板 `assets/templates/template1-cn.xlsx` 标黄表头必填(18 列):
- 中文: `品牌` `图片` `货号` `描述` `类目` `材质` `性别` `颜色` `色号` `季节` `产地` `零售价` `cost` `C+-` `批发价` `尺码` `数量` `云仓`
- 英文: `Brand` `Image` `International code` `Described` `Category` `Material` `Gender` `Color` `Color Code` `Season` `Made in` `Retail price` `cost` `C+-` `Wholesale price` `Size` `Quantity` `Vendor ID`

> ⚠️ 2026-09 起列名是 `云仓` / `数量` / `批发价`,不是旧的 `仓库ID` / `库存` / `供货价`。

## Step 3: 写适配器

### dispatcher 要求的接口

```python
def parse_rows(primary_file: Path, markup_file: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    """返回行列表(一行一个尺码)"""

def get_vendor_meta() -> Dict[str, object]:
    """可选:vendor 级设置,例如图片列宽、每行保留几张图"""
```

**不要**自己写 `main()` / argparse —— `scripts/convert_vendor_quote.py` 是唯一入口,
它负责选模板、算 C+、写图片、输出文件。

### row 字典的字段名是**英文 key**

`map_row_to_template()` 把 row 映射到中英文模板列,所以适配器返回的 dict 用英文 key:

```python
{
    "brand": ..., "sku": ..., "image": ..., "description": ...,
    "category": ..., "material": ..., "color": ..., "color_code": ...,
    "gender": ..., "season": ..., "origin": ...,
    "cost": ..., "cplus_text": ..., "supply_price": ...,
    "retail_price": ..., "size": ..., "quantity": ...,
    "vendor_id": ..., "image_items": [...],
}
```

- `color_code` / `gender` 可选(缺了写 `""`)
- `image_items` 是图片列表,key 用 **数据行号(1-idx)**

### 最小骨架

```python
#!/usr/bin/env python3
"""Vendor V10XX <BRAND> 适配器 (2026-XX-XX)."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    compute_supply_price,
    load_images_by_row_per_sheet,
    to_float,
)

VENDOR_ID = "V10XX"
BRAND = "BRAND NAME"


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """主入口:解析源文件,返回行列表。"""
    rows: List[Dict[str, object]] = []

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    ws = wb.active  # 或 wb['SheetName']

    # 1. 加载图片(4 个加载器,按版式选,见 image-handling.md §8-§10)
    images_by_row = load_images_by_row_per_sheet(primary_file).get(ws.title, {})

    # 2. 解析数据行(按表头名定位列,不要写死列号)
    for r in range(DATA_START_ROW, ws.max_row + 1):
        sku = str(ws.cell(r, SKU_COL).value or "").strip()
        if not sku:
            continue

        # 3. 尺码矩阵展开成一行一个尺码
        for size, qty in expand_sizes(ws, r):
            if qty is None or qty <= 0:
                continue  # 数量 0 → 整行跳过,不要输出占位行
            rows.append({
                "brand": BRAND,
                "sku": sku,
                "image": "",
                "description": ws.cell(r, DESC_COL).value,
                "category": ws.cell(r, CAT_COL).value,
                "material": ws.cell(r, MAT_COL).value,
                "color": ws.cell(r, COLOR_COL).value,
                "color_code": "",
                "gender": "",
                "season": ws.cell(r, SEASON_COL).value or "",
                "origin": ws.cell(r, ORIGIN_COL).value or "",
                "cost": to_float(ws.cell(r, COST_COL).value),
                # ⚠️ C+ 不在这里写死 —— 由 dispatcher 的 --cplus / 规则注入
                "cplus_text": "",
                "supply_price": None,
                "retail_price": to_float(ws.cell(r, RETAIL_COL).value),
                "size": size,
                "quantity": int(qty),
                "vendor_id": vendor_id,
                "image_items": images_by_row.get(r - 1, []),  # anchor 行号是 0-idx
            })

    return rows


def get_vendor_meta() -> Dict[str, object]:
    """vendor 级设置。"""
    return {
        "default_image_col_width": 40.0,
        "default_keep_images_per_row": 1,
    }
```

### ⚠️ C+ 不要在适配器里写死

C+ 是**每单单独给**的(2026-10-28 用户确认):
- 适配器里 `"cplus_text": ""` + `"supply_price": None` 即可
- dispatcher 会用 `--cplus` / `--cplus-rules-json` 覆盖
- **不要**写 `CPLUS_INT = 18` 或 `round(cost * 1.18, 2)` 这种硬编码
- 源文件里真的带 C+ 列(如 BOTTEGA 的 `DESTINAZIONE`)时,才在解析里读出来填 `cplus_text`

详见 `cplus-column.md`。


if __name__ == '__main__':
    main()
```

## Step 3.5: SKU 拼接规范 ⚠️

**当源文件的 STYLE/MODEL 不含颜色代码,需要拼接颜色时,统一格式:`{STYLE}-{颜色代码}`**

| ❌ 错误 | ✓ 正确 | 原因 |
|---|---|---|
| `B90886AUZ` | `B90886-AUZ` | 颜色代码紧贴款号,人眼难分辨边界 |
| `211870245002NEWPORTNAVY` | `211870245002-NEWPORT NAVY` | 同上,且颜色名含空格会混乱 |
| `MODEL+COLOR` 直接拼 | `f"{model}-{color}"` | 必须用 `-` 分隔 |

**为什么用 `-`**:
- 一眼能区分款号 vs 颜色(批发商按货号查库存、运营按颜色配图都依赖这个)
- 全局约定:SKILL.md 货号定义已写 `{STYLE}-{颜色代码}`(TORRY BURCH SHOES 案例)
- Excel 排序时,`B90886-A` 自然排在 `B90886-B` 前面,直接拼会插队

**适用情况**:
- 源文件 `STYLE/MODEL` 字段只到款号,不含颜色代码(同一款多色靠 `COLOR` 字段区分)
- 直接拼款号会冲突:同 MODEL 不同 COLOR 的尺码混在一起,合并时互相覆盖
- V1044 RATTI、V1043 PRITELLI、V1021 BRUNA ROSSO family 等多色同款都属此类

**反例(不要拼)**:
- 源文件 `STYLE` 已含颜色代码(如 `B90886-AUZ` 整体作为款号) → 不拼,直接用
- 源文件 `COLOR` 为空 → 不拼(用款号本身即可)

```python
# 正确的拼接写法
color_code = normalize_text(row.get("COLOR"))
sku = f"{model}-{color_code}" if color_code else model
```

## Step 3.5.1: 颜色字段 vs 色号短码 — 颜色优先 ⚠️

**场景**:源文件同时有「颜色字段」(c8, 长描述如 `VINTAGE BLACK`/`NERO/ARGENTO`/`Fucsia`)和「色号短码」(从 SKU 末段抽出的 `097`/`FC368`/`4040`),拼接货号时**优先用颜色字段,其次用色号短码,都没有不拼**。

**用户期望**(2026-09-03 反馈):
> "货号后面需要拼接颜色。如果有颜色字段优先颜色,其次是色号,都没有则不拼。"

| ❌ 错误 | ✓ 正确 | 真实样本 |
|---|---|---|
| `AMBMJE1119-097` | `AMBMJE1119-VINTAGE BLACK` | 末段 `097` 是色号,但 `VINTAGE BLACK` 才是人眼能识别的颜色 |
| `EM006736 AF29202-FC368` | `EM006736 AF29202-BLACK EA` | 末段 `FC368` 是色号,但 `BLACK EA` 才是真颜色 |
| `PAMYC001DN001-4040` | `PAMYC001DN001-LIGHT BLUE LIGHT BLUE` | 同上 |

**3 档优先级实现**:

```python
import re

def format_sku_with_color_priority(style: str, color: str, color_code: str) -> str:
    """拼接规则:颜色 > 色号 > 不拼。"""
    color_norm = re.sub(r"\s+", " ", color).strip() if color else ""
    if color_norm:
        # 1) 颜色字段优先 (e.g. 'VINTAGE BLACK', 'NERO/ARGENTO')
        return f"{style}-{color_norm}"
    if color_code:
        # 2) 色号短码其次 (e.g. '097', 'FC368')
        return f"{style}-{color_code}"
    # 3) 都没有 → 不拼
    return style
```

**为什么颜色优先**:
- 颜色字段含完整描述(`NERO/ARGENTO`、`VINTAGE BLACK`),人眼和下游系统(配图、按颜色查库存)能直接识别
- 色号短码是 `3-4 位` 数字/字母,跟款号混在一起人眼难分辨边界
- 同款多色靠颜色字段区分(色号短码可能跟其他款号撞,例如 `097` 太通用)

**COLOR 归一**:
- 多空白/换行压成单空格:`'BLACK  EA'` → `'BLACK EA'`、`'BROWN - BROWN -\nBROWN'` → `'BROWN - BROWN - BROWN'`
- 保留 `/`、`-`、`+` 等分隔符(多色描述如 `NERO/ARGENTO`、`BLACK/ SILVER` 整体保留)
- 保留原大小写(品牌命名习惯不同,`Fucsia` ≠ `FUCSIA`)

**反例**(不要用颜色拼接):
- 源 `STYLE` 已含颜色代码(如 `B90886-AUZ`、`241925BLACK`) → 不拼,直接用(Step 3.5 上一条规则)
- `COLOR` 字段为空且 `STYLE` 末段无色号 → 不拼,SKU 用 `STYLE` 原样

## Step 3.6: 自定义尺码表识别(V1044 RATTI 模式)

**场景**:源文件每个数据段前有一行"尺码表头",这一行在 size 列写的是**实际尺码标签**(44/46/48/50),不是默认的列头(XXXS/XXS/XS)。直接用列头会得到错误的尺码(如 B90886 JACKET 应该是 46/48/50,不是 XXS/XS/S)。

**典型版式**(V1044 RATTI,每个段 6 行):
```
Row N+0:  TYPE="TYPE"  DESCRIPTION="..."  MODEL="..."  size列: 44 46 48 50 52 54 ...
Row N+1:  TYPE="AVAIL" desc="JACKET" model="B90886"   size列: 0  1  1  0  0  0 ...
Row N+2:  TYPE="ORDER" (空)
Row N+3-N+5: 空行
```

**错误做法**:用列头(XXXS/XXS/XS)当尺码标签 → B90886 写成 XXS/XS/S

**正确做法**:扫整张表,记录每个 AVAIL 行的"自定义尺码表",fallback 才用列头:

```python
def _build_size_label_map(df, size_cols):
    """建立 df_idx -> {size_col: actual_label} 映射。"""
    result = {}
    current_labels = {}
    for idx, row in df.iterrows():
        type_v = normalize_text(row.get("TYPE")).upper()
        if type_v == "TYPE":
            # header_row: 记录 size 列的实际标签
            current_labels = {c: normalize_text(row.get(c))
                             for c in size_cols
                             if normalize_text(row.get(c))}
        elif type_v == "AVAIL":
            # 数据行:继承最近 header_row 的 labels
            if current_labels:
                result[int(idx)] = dict(current_labels)
    return result

# 使用
size_label_map = _build_size_label_map(df, size_cols)
for idx, row in filtered.iterrows():
    custom_labels = size_label_map.get(int(idx), {})
    for sc in size_cols:
        ...
        size_label = custom_labels.get(sc) or normalize_text(sc)  # 优先 custom
```

**判定要不要用**:如果文件有"header_row (TYPE='TYPE') + data_row (TYPE='AVAIL')"模式,且 header_row 的 size 列出现 44/56/23 等数字(不是 XXXS/XXS/XS 字母),就需要用这个模式。

## Step 3.7: 货号/类目列重映射(V1046 LOSCHI 模式)

**场景**:源文件有多个候选列(ARTICOLO、MODELLO、RIGA),用户实际想要的是 RIGA(每色独立货号),不是 ARTICOLO(同款多色共用)。或者类目字段在不同供应商命名不同(VOCE / CATEGORY / PRODUCT TYPE)。

**错误做法**:用第一匹配的列(`find_col(df, ["ARTICOLO", "ITEM", "MODELLO"])` → 拿到 ARTICOLO,但用户要 RIGA)

**正确做法**:把用户指定的列名放在 find_col 的**第一优先级**:

```python
# 货号用 RIGA 列(LOSCHI 版式:ARTICOLO 是款号,MODELLO 是颜色名,RIGA 才是真货号)
item_col = find_col(df, ["RIGA", "ARTICOLO", "ITEM", "MODELLO"])
# 类目用 VOCE 列(原 DESCRIPTION 字段实际是分类,例如 "JEANS -U")
voce_col = find_col(df, ["VOCE"])
desc_col = find_col(df, ["DESCRIPTION"])  # 真正的描述(没有就空)
```

**判定方法**:让用户确认每个字段的"真名",或者查源文件第 1-2 行的列头(看哪些是数字、哪些是分类文本)。

## Step 3.8: 兼容多套 cost/retail 列名(V1049 POMPEU 模式)

**场景**:源文件的 cost/retail 列在不同供应商命名不同:
- V1049 POMPEU:`WHS` + `RETAIL PRICE` (Hoja1 的 col 35/36)
- V1041 BIFFI:`SELL IN` + `SELL OUT`
- V1042 ALINA:从描述里抠(`<cost> <retail>` 嵌在文本中)
- 早期老供应商:`cost` + `retail` 直接叫

**错误做法**:硬编码列名 → 换一个供应商就漏

**正确做法**:用 `find_col` 列多个候选 + iloc 精确取(因为列号固定比列名稳):

```python
# 候选列名数组(按优先级)
WHS_COL_CANDIDATES = ["WHS", "COST", "SELL IN", "PRZ.ACQUISTO"]
RETAIL_COL_CANDIDATES = ["RETAIL PRICE", "RETAIL", "SELL OUT", "PRZ.VENDITA"]

cost_col = find_col(df, WHS_COL_CANDIDATES)
retail_col = find_col(df, RETAIL_COL_CANDIDATES)

# 拿到 col 后用 df.iloc 取值(比 df.get(col) 稳)
cost = to_float(df.iloc[idx][cost_col]) if cost_col else None
retail = to_float(df.iloc[idx][retail_col]) if retail_col else None
```

**如果 cost/retail 嵌在描述文本里**(V1042 模式):写一个 `parse_desc_with_prices(desc) -> (desc, cost, retail, ...)` 解析器,正则抠出价格。

## Step 3.9: 多图产品要"全收" + 加宽图片列(V1049 POMPEU 模式)

见 `image-handling.md §9 多图同单元格`、`§10 图片列宽设置`。

## Step 3.10: 意大利线/欧式版式(V1015 PRADA 模式)

**场景**:Prada 意大利线的库存清单(鞋类为主),意大利语 + 欧式十进制 + 10 行/产品。

**典型版式**(V1015 ITA_Prada.xlsx):
```
Row N+0  (header_row):  col 1 = 总数量, col 3 = 'SKU' 文本, col 8+ = 各尺码 qty
Row N+2  (sku_row):     col 24 = SKU 编码(20 位 PRADA 内部码)
Row N+3  (desc_row):    col 24 = 描述, col 35 = 颜色
Row N+5  (price_label): col 26 = 'Acquisto'(cost), col 32 = 'Vendita'(retail)
Row N+6  (price_val):   col 26 = cost 数字, col 32 = retail 数字
```
每个产品占 10 行;文件顶部某行是尺码标签行(`SKU` 文本带 `\r\n`),col 4-27 写意大利鞋码 5/5.5/6/.../14。

**关键解析点**:

1. **尺码格式归一化**:
   - 整数(5, 10)  → 浮点 `5.0`/`10.0` → 显示 `5`/`10`
   - 半码(5.5)    → 字符串 `5,5`(欧式十进制逗号) → `5.5`
   - 标签带 `\r\n` 尾巴要先 strip()
   ```python
   if isinstance(v, (int, float)):
       f = float(v)
       out[c] = str(int(f)) if f.is_integer() else str(f).replace(",", ".")
   else:
       sv = str(v).strip()
       if not sv or sv.lower() == "nan":
           continue
       if re.fullmatch(r"\d+(?:[.,]\d+)?", sv):
           sv = sv.replace(",", ".")
       out[c] = sv
   ```

2. **空 cell 是 NaN 不是 None**:pandas `read_excel` 把空 cell 转成 `NaN`,`str(NaN) = 'nan'` 会污染数据
   ```python
   if v is None or (isinstance(v, float) and pd.isna(v)):
       continue
   ```

3. **价格列是意大利语**:`Acquisto` = cost(进价),`Vendita` = retail(售价)
   - 不要硬编码英文 `cost` / `retail` 列名
   - 用 `find_col(df, ["ACQUISTO", "COST"])` 兼容

4. **图片行 = SKU 行**(不是 header 行)
   - openpyxl 0-idx:图片在 `header_row + 2`(因为 SKU 在 header_row + 2)
   - 常见错误:以为是 `header_row + 1`,会导致图片全部丢空

5. **每产品多张图(V1015 是 3 张)**:在 vendor meta 里设 `default_image_col_width`
   ```python
   def get_vendor_meta():
       return {"default_image_col_width": 80.0}  # 3 张 500x500 = ~56 字符,加缓冲
   ```

6. **同 SKU 重复出现**:源文件里同一 SKU 在两处不同 header 行下,各带不同 qty(疑似两批次入库)
   - 不要去重,保留多行(避免漏掉其中一批次的库存)
   - 在 `_missing_fields.md` 的「已知问题」里说明这个情况

**判定要不要用这个模式**:看到 `Acquisto` / `Vendita` 列名,或者文件有大量 10 行一组的固定版式,大概率是意大利线。

## Step 4: 验证

```bash
# 1. 跑转换
python3 convert_V1052.py "/path/to/vendor.xlsx" "V1052"

# 2. 验证输出
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('/path/to/converted/output.xlsx', data_only=True)
ws = wb.active
print(f'rows={ws.max_row} cols={ws.max_column}')
for r in range(1, min(10, ws.max_row + 1)):
    print([ws.cell(r, c).value for c in range(1, 9)])
print(f'images={len(ws._images)}')
"

# 3. 必查项 checklist
# - [ ] 行数 = 期望(每 SKU × 每 size)
# - [ ] 图片数 > 0(如果有源图)
# - [ ] cplus_text 列是整数 18 / 26 / 30
# - [ ] supply_price 列不为空(或符合 C+ 公式)
# - [ ] 货号格式正确(无重复颜色代码拼接)
# - [ ] 颜色列有值(SHOES / EYEWEAR 等非标准品牌)
# - [ ] 材质列有值(或源文件本身为空,记录在 vendor note 里)
# - [ ] qty=0 也写入(不要 if qty > 0 过滤)
# - [ ] cost / 零售价 是数字(没带上 "EUR" / "€" 符号)
```

## Step 5: 记录到 references/vendors/

每个新供应商的调试笔记写到 `references/vendors/<BRAND>-adapter-notes.md`,内容:
- 主文件路径 + sheet 结构
- 关键字段映射(列号表)
- size matrix 结构
- 图片提取方式(xls OLE2 / xlsx drawing / WPS richValue)
- 踩过的 bug + 修复
- 验证结果(行数 / 图片数)

## 适配器存放位置

- **内置 vendor** (见 `SKILL.md` 的「Built-in Vendor Adapters」表,当前 37 个): `scripts/vendor_vXXXX.py`,主调度器自动发现
- **新供应商**: 放在供应商报价单**同目录**下的 `convert_<VendorID>.py`,主调度器会自动加载
- **不要**写到 `~/.minimax/skills/template-resource-converter/scripts/`(那是 skill 内部)

## Step 3.11: 多尺码表自动选用(V1003 GUCCI Sirio 模式)

**场景**:源文件里有 N 套不同的尺码表(衣服欧码 / 美/英鞋码 / 半码 / 单码 TU 等),每个产品通过一个"Scalarino code"列指明用哪套。

**典型版式**(V1003 GUCCI `Sirio DB extract`):
```
Row 1-9:    9 个 Scalarino 尺码表(每行一套)
            col 22 = Scalarino code, col 23+ = 尺码标签
              - Row 1 'A' (ITA/FRA 上装): 32, 34, 36, ..., 68
              - Row 2 'C' (牛仔裤):       0, 1, 23, 24, ..., 48
              - Row 3 'F' (USA 鞋):       1, 1+, 2, 2+, ..., 14+
              - Row 4 'G' (UK 鞋):         2, 2+, 3, 3+, ..., 16
              - Row 5 'H' (XXXS-XXXL 衣服): 3XS, XXS, XS, ..., 6XL, XS/S
              - Row 6 'I' (腰带):          60, 65, 70, ..., 160
              - Row 7 'M' (衬衫 14+):     14, 14+, 14++, 15, ..., 23
              - Row 8 'T' (单码):         TU
              - Row 9 'U' (Unisex 鞋):    34, 34.5, 35, ..., 47.5
Row 10:    空行
Row 11:    HEADER
Row 12+:   DATA
            col 22 = Scalarino code (key 找尺码表)
            col 23+ = qty per size (用对应 Scalarino 表的尺码标签)
            col 68-74: Totale 公式、qty 1676、日期等 summary 列(不是尺码)
```

**关键算法**:
```python
def _build_scalarino_size_maps(df):
    """rows 0-8(0-idx)扫一遍,建 {scalarino_code: {col_idx: size_label}}"""
    maps = {}
    for r in range(0, 9):  # 0-idx 对应 Excel row 1-9
        scalarino = df.iat[r, 22]
        if not scalarino:
            continue
        col_map = {}
        # ⚠️ size 列范围限到 col 67(0-idx 67),跳过 col 68+ 的 summary/date
        for c in range(23, 68):
            v = df.iat[r, c]
            if pd.notna(v) and str(v).strip():
                col_map[c] = str(v).strip()
        if col_map:
            maps[str(scalarino).strip()] = col_map
    return maps

# 数据行处理:
for r in range(11, len(df)):
    scalarino = df.iat[r, 22]
    if scalarino not in maps:
        continue
    for c, size_label in maps[scalarino].items():
        qty = df.iat[r, c]
        if is_positive_qty(qty):
            rows.append(...)
```

**判定方法**:
- 文件第 1-10 行有 col 22-50 的尺码标签(且第 1-9 行的 col 22 不一样)
- 后面数据行的 col 22 是 'A'/'C'/'F' 等短代码
- col 68+ 是公式合计或日期(Totale、SUM、日期等),不是尺码

**坑**:
- col 68-74 容易误判成尺码(因为数据可能是 1676 这种大数字) — 一定要限 size_col_end
- 同一个产品 col 22 决定用哪套尺码表(可能跨 Scalarino,如 DAMA 上装 + 同款不同色)
- 图片用 `load_sheet_images_by_row`(单 sheet 简单 anchor 匹配),每行 1 张图

**常见 bug**(2026-07-17 V1003 踩过):
- `SCALARINO_HEADER_ROWS = list(range(1, 10))` 漏了 0-idx 0(= Excel row 1 = 'A' 尺码表)
- **正确**:`list(range(0, 9))` 覆盖 0-idx 0-8,对应 Excel row 1-9
- 验证方法:parse 后 unique SKU 数应该 = 源 unique SKU 数(否则有尺码表被漏)
- 调试:`_build_scalarino_size_maps` 返回 dict.keys() 打印看是否少了某个 Scalarino

## Step 3.12: 货号使用 Model 而非 ID(V1003 GUCCI 模式)

**场景**:源文件有两套"产品编号":
- `IDArticolo` (col 8):内部数据库 ID,整数(如 1070893)
- `Model` (col 15):正式款号 + 颜色代码,字符串(如 `865077ZATRT 1000`)

用户要的是对外的 **Model 编号**(更像真实货号),不要内部 ID。

**算法**:
```python
# 1) 预扫:找哪些 Model 有多个 Color
model_to_colors: Dict[str, set] = {}
for r in data_rows:
    model = df.iat[r, MODEL_COL]
    color = df.iat[r, COLOR_COL]
    model_to_colors.setdefault(model, set()).add(color)

# 2) 决定 SKU
def make_sku(model, color):
    if not model:
        return ""
    if len(model_to_colors.get(model, set())) > 1:
        # 同 Model 多 Color → 拼,避免 SKU 冲突
        return f"{model}-{color}" if color else model
    return model
```

**典型效果**(V1003 GUCCI):
- 99% 的产品 Model 唯一 → SKU 就是 Model,如 `865077ZATRT 1000`
- 2 个 Model 跟多个 Color 重复 → SKU = `Model-Color`,如 `863136AAGDJ 1000-1000` / `863136AAGDJ 1000-BLACKBLACK`

**判定方法**:
- 源文件有两个候选"产品编号"列,一个是数字 ID,另一个是字符串 Model/Code
- 用户明确说"用 Model" / "用款号" / "用 Code"
- Model 通常是 13-16 位字符(含字母+数字+空格),ID 通常是 5-8 位纯数字

**坑**:
- 614/616 Model 唯一 → 不需要拼 color;但要测,不能默认不拼
- 拼 color 时用 `-` 分隔符(per skill §3.5)
- 如果 Model 本身就含 color code(如 `865077ZATRT 1000` 末尾 1000 是 color),不要重复拼

## Step 3.13: 用"结构判断"代替 `_is_sku` 内容判断(全 vendor 通用)

**核心思想 — 工厂模板一致,结构判断比内容判断更稳**:
- 内容判断(`_is_sku`):看某个字符串"像不像"SKU(长度、字符集、数字位数)
  - 缺点:不同品牌 SKU 命名规则不同,得为每个品牌写一份 _is_sku
  - 缺点:稍严格的校验会毙掉不熟悉的 SKU 长度(如 PATAGONIA 9 字符)
- 结构判断:看 R 周围 3 行的版式特征
  - 优点:版式一致就吃,**SKU 长度随便变**
  - 优点:不再需要 `_is_sku` 这种内容判断函数

**适用范围**:
- 2026-07-17 audit:17 个 vendor 全部纯结构判断,无任何 `_is_*` / `is_*` 内容判断函数
- V1004 / V1010 / V1012 / V1015 / V1021 / V1031 / V1038 / V1039 / V1041 / V1042 /
  V1043 / V1044 / V1046 / V1048 / V1049 / V1050 / V1003

**结构判断的两条铁律**(V1004 / V1021 / V1042 / 任何工厂都适用):
1. **描述行存在**: R+1..R+3 某行同列有字符串,且**含欧洲价格**(`\d+[.,]\d{2}`)
2. **尺码行存在**: 描述行之后(R+2..R+12)能用 `_find_size_and_qty` 扫到尺码+qty

任何"长得像内容"的行(描述、尺码、段头)都不可能同时通过这两条,
所以 R 那个单元格只要是非空字符串就行 — **不要做 _is_sku 内容判断**。

**算法 V1 — V1004 (col 13 固定,行数严格 10-row)**:

```python
_EURO_PRICE_RE = re.compile(r"\d+[.,]\d{2}")
_HEADER_WORDS = frozenset({"DONNA","UOMO","ABBIGLIAMENTO","ACCESSORI",...})

def _find_product_blocks(df):
    blocks = []
    for r in range(len(df)):
        # 排除段头行
        for c in (3, 4):
            v = df.iat[r, c]
            if isinstance(v, str) and v.strip().upper() in _HEADER_WORDS:
                break
        else:
            # 铁律 1: R+2 col 13 必须含欧洲价格
            if r + 2 >= len(df):
                continue
            desc = df.iat[r + 2, 12]
            if not isinstance(desc, str) or not _EURO_PRICE_RE.search(desc):
                continue
            # 铁律 2: R+2 之后能搜到尺码+qty
            size_map, qty_map = _find_size_and_qty(df, r + 2)
            if not size_map or not qty_map:
                continue
            # R col 13 = SKU(任何非空字符串)
            sku = df.iat[r, 12]
            if not isinstance(sku, str) or not sku.strip():
                continue
            blocks.append({"sku_row": r, "desc_row": r + 2, "sku": sku.strip(), ...})
    return blocks
```

**算法 V2 — V1021 (多列扫描,描述行偏移 1-3 行)**:

```python
def _find_product_blocks(df):
    blocks = []
    for r in range(len(df)):
        for c in range(df.shape[1]):  # 扫所有列
            val = df.iloc[r, c]
            if not (pd.notna(val) and isinstance(val, str)):
                continue
            sku = val.strip()
            if not sku:
                continue
            # 铁律 1: R+1..R+3 同列有描述(且含欧洲价格)
            desc_val = None
            desc_row = None
            for off in (1, 2, 3):
                if r + off >= len(df): break
                nv = df.iloc[r + off, c]
                if isinstance(nv, str) and nv.strip() and _EURO_PRICE_RE.search(nv):
                    desc_val = nv.strip()
                    desc_row = r + off
                    break
            if desc_val is None:
                continue
            # 铁律 2: 描述行之后能搜到尺码+qty
            size_map, qty_map = _find_size_and_qty(df, desc_row)
            if not size_map or not qty_map:
                continue
            blocks.append({"sku_row": r, "desc_row": desc_row, "sku": sku, ...})
            break
    return blocks
```

**关键点**:
- **不要再写 `_is_sku`** — 直接靠周围版式判断,SKU 长度怎么变都不影响
- **不要写工厂校验**(`_validate_v1004_format` 之类)— parse_rows 返回 0 行
  就是隐式校验,用户换 V-code 即可,不需要单独的校验函数
- **V1004 和 V1021 互相独立** — V1004 显式 import V1021 的 helper(`_find_size_and_qty`、
  `_parse_description` 等),V1021 的全局变量(包括任何 SKU 校验)不被修改

**复用 vs 独立 — 什么时候该 import,什么时候该复制**:
| 复用(import helper) | 独立(自己写) |
|---|---|
| 描述行内嵌价格解析(`_parse_description`) | 产品块定位(工厂相关) |
| 尺码行 + qty 搜索(`_find_size_and_qty`) | SKU 校验(本例干脆不要) |
| 图片按行索引(`_load_images_by_row`) | 段头排除(工厂相关) |
| 颜色拆分(POLO 风格 `_extract_color_from_sku`) | 块判断的版式细节 |
| 段头类目查找(`_find_category_for_row`) | |

**坑**:
- 描述行必须含**欧洲价格**(`\d+[.,]\d{2}`)— 这是过滤"非产品行"的最强信号
- 段头行(Donna/Uomo/Abbigliamento 等意大利词)必须主动排除,否则会当成 SKU
- POLO 有些产品的尺码标签是 UK 00/02/04/06,不是 XXS/XS/S — **白名单尺码词会漏**
  → 直接用 V1021 的灵活扫(只看 cols 8-24 里有没有短字符串),不做白名单

**反面教材**(我之前 3 个错的设计):
1. 写"自动检测 + 双解析器" — 过度设计,版式一致只差 SKU 长度
2. monkey-patch V1021._is_sku — hacky,V1004 跟 V1021 强耦合,不能独立运行
3. 写 `_validate_xxx_format()` 工厂校验函数 — 过度设计,parse_rows 返回 0 行就是隐式校验
4. 留个 `is_likely_sku()` 内容判断 — 跟 V1021 一样会拦掉不熟悉的 SKU 格式(V1042 之前就有这个问题,已删)

## Step 3.14: 多品牌同版式 — 扫源文件顶部 brand/gender(V1038 TORY BURCH 模式,2026-08-28)

**场景**:V1038 / V1010 / V1062 / V1063 等 col 14 块状版式,同套结构但出**多品牌**(A.P.C. / TORY BURCH / GUCCI 等)。vendor adapter 不能写死 `default_brand = "A.P.C."`,要扫源文件。

**扫描规则**:
- 扫 R1-R20(顶部 metadata 区)
- col 4 (D) 找 brand:全大写、多字、长度 ≥ 2
- col 5 (E) 找 gender:白名单 `WOMAN / MAN / UNISEX / UOMO / DONNA / KIDS / BAMBINO / BAMBINA`

```python
GENDER_KEYS = {"WOMAN", "MAN", "UNISEX", "UOMO", "DONNA", "KIDS", "BAMBINO", "BAMBINA"}

def _scan_top_meta(primary_file: Path) -> tuple[str, str]:
    """扫源文件顶部 R1-R20, 在 col 4 (D) / col 5 (E) 找 brand / gender 字符串。"""
    brand = gender = ""
    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True, read_only=True)
    except Exception:
        return brand, gender
    if not wb.worksheets:
        return brand, gender
    ws = wb.worksheets[0]
    for r in range(1, min(21, ws.max_row + 1)):
        v4 = ws.cell(r, 4).value
        v5 = ws.cell(r, 5).value
        if not brand and isinstance(v4, str):
            cand = v4.strip()
            if cand and cand == cand.upper() and len(cand) >= 2 and not cand.isdigit():
                brand = cand
        if not gender and isinstance(v5, str):
            cand = v5.strip().upper()
            if cand in GENDER_KEYS:
                gender = cand
    return brand, gender

def parse_rows(primary_file, conditions_file, vendor_id):
    rows = _col14_parse(primary_file, VENDOR_ID, parser="v1038")
    src_brand, src_gender = _scan_top_meta(primary_file)
    if src_brand or src_gender:
        for r in rows:
            if src_brand and not r.get("brand"):
                r["brand"] = src_brand
            if src_gender and not r.get("gender"):
                r["gender"] = src_gender
    return rows
```

**要点**:
- `brand` 和 `gender` 都是**可选**字段,扫不到时 vendor `default_brand` 兜底
- 扫到空就别覆盖 — 让 dispatcher 走 5 级 brand 优先级
- 跟 `v1038.py` 的 A.P.C. 默认共存 — 源扫到 `TORY BURCH` 优先,扫不到走 `A.P.C.`

**判断要不要做**:
- 源文件 R1-R20 看到 D 列大写多字字符串(如 `TORY BURCH` / `GUCCI` / `BRUNELLO CUCINELLI`)+ E 列 `WOMAN` / `MAN` → 启用
- 看不到这种 metadata → 不用,走 `default_brand`

## Step 3.15: Sirio DB 多语言 HEADER 兼容(V1003 GANNI 模式,2026-09-01)

**场景**:V1003 之前只认 `CATALOGO` sheet + 英文 HEADER labels。GANNI 导出是 `Foglio1` sheet + 意大利语 HEADER(`Costo` / `Modello` / `Variante` / `Marca` / `Sesso` 等)。

**两个兼容改动**:

### 3.15.1 多 sheet name fallback

```python
def parse_rows(primary_file, _, vendor_id):
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    # 找数据 sheet — 优先级:CATALOGO → Foglio1 → 第一个非 Parametri
    target_sheet = None
    for cand in ("CATALOGO", "Foglio1", "Sheet1"):
        if cand in wb.sheetnames:
            target_sheet = cand
            break
    if target_sheet is None:
        # 兜底:跳过 Parametri 这种配置 sheet, 取第一个数据 sheet
        for sn in wb.sheetnames:
            if sn.lower() not in ("parametri", "parameters", "config"):
                target_sheet = sn
                break
    if target_sheet is None:
        return []
    ws = wb[target_sheet]
    df = pd.read_excel(primary_file, sheet_name=target_sheet, header=None)
    # ... 后续正常解析
```

**判断规则**:
- 优先尝试意大利 Sirio DB 默认 sheet name(`Foglio1`)
- 再尝试英文版默认(`Sheet1` / `CATALOGO`)
- 兜底:跳过配置类 sheet(`Parametri` / `Parameters` / `Config`),取第一个剩余 sheet

### 3.15.2 LABEL_ALIASES 扩展(意大利语字段)

GANNI Sirio DB 用 `Marca` 作 origin 字段(不是 `Made in` / `Origine` / `Paese`):

```python
LABEL_ALIASES = {
    # ... 原有别名
    LABEL_MADE_IN: ["made in", "origine", "paese", "marca"],  # ← 加 marca
}
```

**判断规则**:
- 看 HEADER 行 R1-R20,识别 label 文本
- 跟源文件实际 label 文本对照(不是英文 lable 就是意大利语 lable)
- 缺的别名加进 `LABEL_ALIASES`

**GANNI 完整映射**:
| 模板字段 | GANNI HEADER label | LABEL 常量 | 已支持? |
|---|---|---|---|
| cost | `Costo` | LABEL_COST | ✅ alias `costo` |
| retail | `Retail` | LABEL_RETAIL | ✅ |
| brand | `Linea` | LABEL_BRAND | ✅ alias `linea` |
| gender | `Sesso` | LABEL_GENDER | ✅ alias `sesso` |
| description | `Descrizione` | LABEL_DESCRIPTION | ✅ alias `descrizione` |
| model | `Modello` | LABEL_MODEL | ✅ alias `modello` |
| variant/color | `Variante` | LABEL_VARIANTE | ✅ alias `variante` |
| **origin** | **`Marca`** | LABEL_MADE_IN | **✅ 加 alias `marca`** |
| material | `Materiale` | LABEL_MATERIAL | ✅ alias `materiale` |

**注意 `marca` 在 LABEL_BRAND 也有**!不要两个都加 brand alias,会冲突。Sirio DB 里 `Linea` 才是 brand,`Marca` 才是 origin(产地)。

## Step 3.16: 货号里抽色号(V1003 GANNI Modello 模式,2026-09-01)

**场景**:源数据 Modello 字段格式 `<字母数字> <3位色号>`(e.g. `A6976 099`),色号藏在 Modello 末尾。要抽出来写到模板的「色号」列。

**抽取代码**:

```python
import re

# 6.1 SKU
sku_base = normalize_text(df.iat[r, sku_col]) if sku_col < df.shape[1] else ""
if not sku_base:
    continue
color = normalize_text(df.iat[r, color_col]) if color_col is not None and color_col < df.shape[1] else ""
sku = make_sku(sku_base, color)
# 2026-09-01 GANNI 接入:Modello 格式 "<字母数字> <3位色号>" (e.g. "A6976 099"),
# 末尾 3 位数字是 color_code,写到模板的「色号」列
color_code = ""
if sku_base:
    m = re.search(r"(\d{3})$", sku_base.strip())
    if m:
        color_code = m.group(1)
# ... 在 rows.append() 里加 "color_code": color_code
```

**判定规则**:
- Modello 末尾连续 3 位数字 → 抽出来当 color_code
- 用 `re.search(r"(\d{3})$", sku_base)` 而不是 `endswith` 切字符串 — 兼容 4+ 位 ID 变体
- 找不到就留空(空字符串),不要 fallback

**模板输出**:`color_code` 字段在 row dict 里,framework 的 `map_row_to_template` 会写到中文模板的「色号」列(英文模板对应 color code 位置)。

**和 SKU 的关系**:
- `sku = "<Modello>-<Variante>"`(e.g. `"A6976 099-BLACK"`)— 完整标识
- `color_code = "099"`(从 Modello 末尾抽)— 单独色号
- `color = "BLACK"`(从 Variante 列读)— 颜色名

三者独立,在模板三列分别填。

## Step 3.17: 一次性无 V10XX 工厂代码的转换(2026-07-21 + 2026-08-26 案例)

**场景**:用户给一份临时现货单(没有 V10XX 工厂代码),要求"一次性的"转换输出 18 列模板。**不**进 BUILTIN_VENDORS,写成 `scripts/oneoff_xxx.py` 独立脚本。

**判断条件**:
- 用户明确说"没有工厂代码" / "一次性的" / "不要存到 BUILTIN"
- 文件结构跟现有 vendor 都对不上(或者用户不想归类)
- 但模板字段映射可以手工写

**脚本骨架**(基于 framework 工具复用):

```python
#!/usr/bin/env python3
"""Tory Burch Stock List 4199 — 一次性转换(2026-07-21)。"""
import sys
from pathlib import Path
import openpyxl
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from convert_vendor_quote_shared import (
    KEEP_IMAGES_PER_ROW, SKU_COL_WIDTH, TMP_DIR,
    embed_images_in_row, load_images_by_row_per_sheet,
    load_sheet_row_heights, map_row_to_template,
)

# 硬编码(因为没有 V10XX 工厂代码)
BRAND = "TORY BURCH"
GENDER = "WOMAN"
CPLUS = 8  # 用户口头 "c+=+8" 直接转整数 8
WAREHOUSE_ID = ""  # 没工厂代码,空

CN_HEADERS = [
    "品牌", "图片", "货号", "描述", "类目", "材质", "性别",
    "颜色", "色号", "季节", "产地", "零售价", "cost", "C+-",
    "批发价", "尺码", "数量", "云仓",
]

def _to_float(v):
    if v is None: return None
    try: return float(v)
    except: return None

def _is_positive_qty(v):
    if v is None: return False
    if isinstance(v, str):
        s = v.strip()
        if not s or s == "0": return False
    n = _to_float(v)
    return n is not None and n > 0

def _parse_sheet(ws, sheet_images, sheet_heights):
    out = []
    # ... 自定义字段映射
    return out

def convert(source_file, template_file, output_file):
    src_wb = openpyxl.load_workbook(source_file, data_only=True)
    # ... 5 sheet 各跑一次 _parse_sheet
    # 写到新文件,每 sheet 1:1 对应
    pass

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--template", default=str(SCRIPTS_DIR.parent / "assets" / "templates" / "template1-cn.xlsx"))
    ap.add_argument("--output", default=None)
    args = ap.parse_args()
    # ... 自动 output = 配置器测试文件/<VENDOR>/converted/<clean_stem>.xlsx
```

**关键决策**:
- 复用 framework `load_images_by_row_per_sheet` / `load_sheet_row_heights` / `embed_images_in_row` / `map_row_to_template`
- **不**复用 `convert_vendor_quote.py` 的 dispatcher(没 vendor 走 dispatcher)
- 输出目录按 framework 惯例: `<配置器测试文件>/<品牌或代号>/converted/<clean_stem>.xlsx`

**不要**:
- 进 BUILTIN_VENDORS(避免污染 dispatcher 状态)
- 写到 `scripts/vendor_v<notV10XX>.py` — 命名约定保留给工厂代码 vendor
- 在 framework `convert_vendor_quote.py` 加特殊路径(framework 不该知道一次性脚本)

**Clean name helper**(去掉 user 上传到 `.minimax/v2/assets/` 时被加的 hash 前缀):

```python
import re
def _clean_name(stem: str) -> str:
    m = re.match(
        r"^\d{2}-\d{2}-\d{2}-\d{3}-asset_\d{8}-\d{6}-\d{3}_[0-9a-f]+_[0-9a-f]+-(.+)$",
        stem,
    )
    return m.group(1) if m else stem
```

例:`'17-29-29-693-asset_20260828-172929-693_eea453a5b0c1_88138a71-女士FW26'` → `'女士FW26'`



