# V1042 POLO RALPH LAUREN 适配笔记

**Vendor code**: V1042
**品牌**: POLO RALPH LAUREN (Ralph Lauren 子品牌)
**来源工厂**: ALINA (意大利供应商)
**适配器路径**: `scripts/vendor_v1042.py`
**最后更新**: 2026-09-12

## 文件特征

- 后缀: `.xls` (OLE2 老格式,带图)
- 大小: ~1.2 MB (POLO R-75 现货单)
- 产品数: 124
- 唯一 SKU: 通常 ≤ 产品数 (某些产品多色)
- 总行数: ~1300 行 (xlrd 0-idx),包括 SKU + 描述 + size_row + qty_row 重复块
- Sheet: 单 sheet

## 版式 (6 行/product 块)

```
R+0: SKU 在 col 9 (0-idx)         e.g. "211870245---003"
R+1: 描述 col 9,含两个欧洲价格    e.g. "Camicia Donna Bianco       257,00       720,00 ITALIA Natural (Veg)->Cotton72 %"
R+2: 空
R+3: 尺码行 (header 风格)         INT  XXXS  XXS  XS  S  M  L  XL  XXL  XXXL  4XL  5XL
R+4: 空
R+5: qty 行                      0   0   0   1   7   5   0   0   0   0
```

## 关键解析细节

### 1. SKU 拼接 (POLO 风格 `<style>---<color_code>`)

源 SKU 是 `<style>---<color_code>` 格式(注意是 `---` 三个连字符,不是 `-`):

```python
import re
m = re.match(r"^(.+?)---+(.+)$", sku)
if m:
    sku = f"{m.group(1)}-{m.group(2)}"  # 重组为 style-color_code (单 - 分隔符)
    color_code = m.group(2)
```

**为什么用单 `-` 重组**:`---` 是源文件格式,框架模板用单 `-` 作为最终分隔符(V1042 输出模板用单 `-`,跟其他 vendor 一致)。

### 2. 描述行解析 (`<类型> <性别> <颜色> <cost> <retail> <产地> <材质>`)

POLO 描述行嵌入了所有信息,需要拆:

```
"Camicia Donna Bianco       257,00       720,00 ITALIA Natural (Veg)->Cotton72 %"
       ↓ 拆 ↓
type=Camicia, gender=Donna, color=Bianco, cost=257.0, retail=720.0, origin=ITALIA, material="Natural (Veg)->Cotton72 %"
```

`_parse_desc_with_prices(desc_line)` 实现:
1. 抓末尾所有 `\d+[.,]\d+` 价格(可能有 1-2 个, cost/retail)
2. 抠掉价格,找最后 1 个性别词(Uomo/Donna/Unisex/Unica/Bambino/Bambina)
3. 性别词之前是 type+gender,之后是 color+origin+material
4. 产地 = 第一个全大写单词(或已知多词产地 "Stati Uniti" / "Repubblica Popolare Cinese")
5. gender 翻译:`Uomo→Man / Donna→Woman / Unisex→Unisex / Bambino→Kid`

### 3. 尺码行 size_code 识别

POLO 尺码行第一个是 size_code (如 `INT` / `STN` / `ITD` / `FRD` / `UNI` / `PZ` / `CMU` / `UKU`),之后是 size_label (XXXS/XXS/XS/S/M/L/XL/XXL/XXXL/4XL/5XL)。

**坑**: 早期 `row_values()` 收集所有 size/qty 然后按位置 zip,会错位 5-13 列。

**修复** (2026-09-12): 用列对齐
```python
SIZE_CODE_TOKENS = {"ITD", "FRD", "UNI", "PZ", "CMU", "UKU", "INT", "STN",
                    "FIT", "G", "FIN", "TU", "OS", "M", "L", "S", "XL", "XXL", "XXXL", "4XL", "5XL"}

def _find_size_code_col(row_idx):  # 找 size row 第一个 size_code col
    for c in range(df.shape[1]):
        v = df.iat[row_idx, c]
        sv = str(v).strip()
        if sv.upper() in SIZE_CODE_TOKENS:
            return c
    return None

def _find_qty_start_col(row_idx):  # 找 qty row 第一个数字 (>0) col
    for c in range(df.shape[1]):
        v = df.iat[row_idx, c]
        f = to_float(v)
        if f is not None and f > 0:
            return c
    return None

# 配对: offset = qty_起点 - size_label 起点, qty 前补 0
size_label_start = size_code_col + 1
qty_nums = [...row 5 从 qty_start_col 起...]
qty_padded = [0.0] * offset + qty_nums
size_labels = [...row 3 从 size_label_start 起...]
pairs = list(zip(size_labels, qty_padded))[:min_len]
```

### 4. 图片抽取 (F010 raw scan fallback)

**问题**: POLO R-75 .xls (1.2MB, 124 产品) 用 xlrd monkey-patch 只解 28 个 F010 anchor,后 96 个产品的图被 xlrd 跳过。

**修复**: 优先 raw scan F010 signature (`\x10\xf0\x12\x00\x00\x00`),解 (col_lo, row_lo),raw 解不到才退回 xlrd。

详细 raw scan 算法见 `OLE2-xls-JPEG-extraction-pattern.md` § "xlrd MSODrawing 大文件漏解"。

### 5. swatch 裁剪 (POLO 源图是"产品 + 色卡"组合)

POLO .xls 里每件产品的图就是"产品照 + 下方 1 个色卡 (swatch)"组合图(281x281 / 375x375)。直接嵌到 92pt 缩略图看起来像"产品 + 彩色色块"。

**修复**: 用 `_strip_bottom_swatch` 全局 helper 自动检测并裁掉 swatch 区域,补白回 1:1。详见 `image-handling.md` § 19。

### 6. qty=0 行过滤 (2026-09-12 fix)

V1042 早期 commit 注释写"保留全部 size 含 qty=0 写 0 行,让 framework 知道有货 size",**与 framework 主流规则 (`is_positive_qty` 过滤) 冲突**。

**用户偏好** (2026-08-28 多次确认, 见 `image-directory-matching.md` § 5):
> 库存为 0 的行不输出(不输出占位 row, SKU 整行跳过)

**修复**: 在 `(size_label, qty) in pairs` 循环里加 `if not is_positive_qty(qty): continue`。

### 7. 季节从文件名提取

POLO 现货单文件名常含 `SS26` / `FW26`,用正则 `SS\d{2}|FW\d{2}` 提取作为 season 字段(目前 vendor_v1042 没填 season,可后续补)。

## 实测数据

POLO RALPH LAUREN 2026 现货单 (`POLO RALPH LAUREN.xls`, 1.2MB):
- 124 个产品 / 124 unique SKU
- 修复前 (qty=0 保留 + xlrd F010 漏解): 1191 rows / 28 SKU 有图 / 297 imgs
- 修复后 (qty=0 过滤 + raw scan F010): 327 rows / 122 SKU 有图 / 319 imgs (98.4%)

最后 2 个 SKU (`710B30157-001`, `710B30186-001`) 源 .xls 没图(122 张 JPEG vs 124 个 F010 anchor),不可避免。

## 相关文件

- `scripts/vendor_v1042.py` — V1042 适配器
- `scripts/convert_vendor_quote_shared/__init__.py` — `_strip_bottom_swatch` / `load_xls_ole2_jpeg_images`
- `references/patterns/OLE2-xls-JPEG-extraction-pattern.md` — F010 raw scan 算法
- `references/patterns/image-handling.md` § 19 — swatch 裁剪
- `references/patterns/image-directory-matching.md` § 5 — qty=0 跳过用户偏好

## 经验教训

1. **硬编码 size 列位置会翻车**,必须从 size_code token (`INT` / `STN`) 动态找列(用户发的两个 PDF spacing 不一样)
2. **POLO .xls 大文件 xlrd MSODrawing 解析不可靠**,必须 raw scan F010 fallback
3. **POLO .xls JPEG 数据流有 spurious `\xff\xd9`**,不能简单用第一个 EOI 切
4. **POLO 源图是 swatch 组合图**,必须自动裁剪或让 ALINA 重导出
5. **qty=0 行必须过滤**,跟 framework `is_positive_qty` 一致(不是"保留 0 行让 framework 知道有货 size",那是早期错误 commit 注释)