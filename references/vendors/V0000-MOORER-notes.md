# V0000 MOORER 适配笔记 (临时)

**Vendor code**: V0000 (临时, 等待正式 V10XX 分配)
**品牌**: MOORER (意大利高端外套品牌)
**适配器路径**: `scripts/vendor_v0000.py` (走 scripts/ fallback, 不入 BUILTIN_VENDORS)
**最后更新**: 2026-09-15

## 文件特征

- 后缀: `.xlsx` (OOXML)
- 大小: ~3 MB (MOORER R-75 special list)
- 产品数: 776 unique SKU
- 总行数: 777 (header + 776 数据行)
- Sheet: 1 sheet "FILE"
- 列数: 59 (col 1-9 metadata, col 10-58 各尺码 qty, col 59 总数)

## 版式 (单 header + 数据行)

```
R1 (header):
  col 1   性别         "Men" / "Women"
  col 2   类目         "BOMBER" / "GIACCHE" / "MAGLIERIA" / "SCIARPE" / ...
  col 3   图片         (None 或 "Immagine non trovata" 文本)
  col 4   描述         e.g. "ANDRAS-MAM"
  col 5   颜色         e.g. "CASTANO" / "PISTACCHIO/BEAVER"
  col 6   国际编码     e.g. "MOUBO100118TEPA663U0237" (23 字符)
  col 7   Size grid    "MO" 男 / "MD" 女 / "UNI"
  col 8   欧洲零售价   e.g. 3390
  col 9   供货价       e.g. 847.5
  col 10-58 各尺码 qty (尺码 label 在 R1, 数字如 26/27/28/.../62 + 小数 41.5/42.5 + M/L/XL/UNI)
  col 59  总数         (不要当尺码!)
```

## 关键解析细节

### 1. 货号拼接规则 (2026-09-15 user 要求)

**规则**:
1. 先抽色号: SKU 末尾 5 位 `字母+4位数字`(`U\d{4}` / `V\d{4}`)
2. 有颜色名(col 5) → 拼 `<原SKU>---<颜色名>`(V1042 POLO 风格)
3. 没颜色名但有抽出的色号 → 拼 `<原SKU>---<色号>`
4. 都没 → 原 SKU 不动

```python
COLOR_CODE_RE = re.compile(r"[A-Z]\d{4}$")
m = COLOR_CODE_RE.search(sku_raw)
color_code = m.group() if m else ""

if color_name:
    sku = f"{sku_raw}---{color_name}"
elif color_code:
    sku = f"{sku_raw}---{color_code}"
else:
    sku = sku_raw
```

**实测**:
- `MOUBO100118TEPA663U0237` + `CASTANO` → `MOUBO100118TEPA663U0237---CASTANO`
- color_code 自动填入 col 9 = `U0237`

### 2. 尺码列读取 + 排除"总数"

尺码在 R1 col 10-58,但 col 59 header 是 "总数" 这种 metadata,不能当尺码:

```python
SKIP_HEADERS = {"总数", "total", "合计", "sum"}
size_labels: List[str] = []
size_cols: List[int] = []
for c in range(10, ws.max_column + 1):
    hdr = normalize_text(ws.cell(1, c).value)
    if hdr.lower() in SKIP_HEADERS:
        continue
    size_labels.append(hdr)
    size_cols.append(c)
```

### 3. 类目 (意大利语 → 英文) 映射

MOORER 用意大利语类目 (`BOMBER` / `GIACCHE` / `MAGLIERIA` / `SCIARPE` / `CAPPOTTI` / `PARKA` / `POLO` / `TSHIRTS` / `CARGO` / `CHINO` / `DENIM` / `BOOTS` / `SNEAKERS` / `MOCASSINI` / `STRINGATE` / `BORSE E ZAINI` / `BERRETTI` / `PICCOLA PELLETTERIA` / `SWIMWEAR` / `FELPA` / `TUTA` / `5 TASCHE` / ...)

`CATEGORY_MAP` 在 vendor_v0000.py 里维护 (40+ 映射)。**未映射的留原意大利语**。

### 4. 性别翻译

```python
gender_it = ws.cell(r, 1).value.upper()  # "MEN" / "WOMEN"
gender = "Man" if gender_it == "MEN" else "Woman" if gender_it == "WOMEN" else gender_it
```

### 5. 图片 (从 ws._images 拿, 不读 col 3 文本)

MOORER 的 col 3 "图片" 都是 None 或 `Immagine non trovata` 文本(不是真的 Image 对象)。但 openpyxl 能从 drawing 里读出真实图:

```python
image_items: List[Dict[str, object]] = []
for img in ws._images:
    anchor = img.anchor
    if hasattr(anchor, "_from") and anchor._from and anchor._from.row + 1 == r:
        try:
            img_bytes = img._data() if hasattr(img, "_data") else None
        except Exception:
            img_bytes = None
        if img_bytes:
            image_items.append({"bytes": img_bytes, "format": img.format or "jpeg"})
```

**注意**: 必须传 `dict` (`{"bytes": ..., "format": ...}`),不能传 openpyxl Image 对象 — `write_output` 期望 dict 格式。

### 6. is_positive_qty 过滤 (framework 主流)

跟所有 vendor 一样, qty <= 0 不输出。

## 实测数据

MOORER R-75 special list.xlsx (3MB):
- 776 unique SKU
- 修复前 (未排除"总数"列): 1891 rows / 776 SKU / 0 imgs / 4662 件
- 修复后 (排除总数 + 拼接 color_name + 从 ws._images 抽图): 1115 rows / 776 SKU / 862 imgs / 4662 件

注意 1891 → 1115 是因为很多 size 行在总数列的 qty 都是 0(从 col 58 跳到 col 59 总数时数据没了),实际 qty>0 的行只有 1115。

## 临时性质 & 后续

V0000 是占位 vendor code,因为:
- MOORER 还没正式分配 V10XX code (类似 V1070 BC / V1072 BURBERRY 的命名)
- adapter 是 inline 简化版,40+ 意大利语类目映射手写维护

**升级路径**:
1. 等运维/产品给正式 vendor code (如 V1073 / V1075 等)
2. 改 `scripts/vendor_v0000.py` → `scripts/vendor_v1073.py`
3. 改 `VENDOR_ID = "V0000"` → `VENDOR_ID = "V1073"`
4. 改 `app.py` 的 `VENDOR_CODES` 条目
5. 旧文件可以保留作 fallback

## 相关文件

- `scripts/vendor_v0000.py` — V0000 适配器 (inline)
- `app.py` line 52 — `VENDOR_CODES` 注册 "V0000"
- `references/patterns/image-handling.md` — swatch 裁剪 / OneCellAnchor 嵌入 (framework 通用)
- `references/patterns/OLE2-xls-JPEG-extraction-pattern.md` — .xls JPEG 抽取 (MOORER 是 .xlsx 不适用, 但 pattern 通用)