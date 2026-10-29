# col14 块状版式(Column-14 Block Layout)

> 实现:`scripts/convert_vendor_quote_shared/col14_parsers.py`
> 适用:V1010 LONGCHAMP(.xlsx)、V1038 APC / TORY BURCH / GUCCI、GUCCI WAREHOUSE B/C 老单

## 版式结构

货号在 **col 14 (N)**,一个货号占**一个块**,块内纵向排布:

```
col 14 (N)
  R0 : 货号            ← 块头
  R1 : 描述            ← 含 2 个 XX,XX 格式的价格 = cost + retail
  R2 : (空)
  R3 : 尺码行          ← col 13+
  R4 : (空)
  R5 : 数量行          ← col 17+
```

真实样本:

| 文件 | 货号数 | 描述样例 |
|---|---|---|
| V1010 LONGCHAMP (.xlsx) | 36 | `made in china, leather Yellow 37,50 90,00` |
| V1038 APC (.xlsx) | 24 | `T-SHIRTS White 44,00 110,00 TURKEY 100% CO` |

## 关键实现要点

### 1. 反向定位货号行,不依赖 SKU 格式

不要用「像不像 SKU」判断块头。用**描述特征**判断:

```python
def _looks_like_desc(s: str) -> bool:
    """判断字符串是不是描述行内容(用于反向定位货号行)。"""
    if not s:
        return False
    s_low = s.lower()
    if 'made in' in s_low or 'madei' in s_low:      # 强信号 1: 产地短语
        return True
    if PCT_RE.search(s):                            # 强信号 2: 材质百分比
        return True
    if re.search(r'\d+[.,]\d{2}\s+\d+[.,]\d{2}', s): # 强信号 3: 成对价格
        return True
    return False
```

`'madei'` 是为了容忍 OCR 错字。

### 2. 价格正则要同时吃欧式和美式

```python
PRICE_RE = re.compile(r"\d{1,5}\s*,\s*\d{2}")   # 欧式 37,50 / 90,00
```

描述里还有美式 `100.50 250.00`,所以 `_looks_like_desc` 用的是
`\d+[.,]\d{2}\s+\d+[.,]\d{2}`。

### 3. 尺码标签白名单

块里有一行是尺码标签,不是货号。用白名单排除:

```python
SIZE_LABELS = {
    "UNI", "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
    "TU", "TU.", "ONE", "ONE SIZE", "UNICA",
}
```

### 4. 材质百分比要单独抽

```python
PCT_RE = re.compile(r"\d+\s*%")   # '100% Cotton' / '5% Elastane'
```

描述里可能混着产地 / 材质 / 颜色 / 价格,拆的时候按特征分段,不要按位置硬切。

### 5. 图片位置

`.xlsx` 时图片在 **col 5 (F) 附近**,**行号 ≈ 货号行 - 1**(openpyxl anchor 0-idx)。

## 与平铺版式的关系

同一个 vendor 可能同时管两种版式,按内容分流:

| 版式 | 判据 | 解析器 |
|---|---|---|
| col14 块状 | 货号在 col 14 | `col14_parsers` |
| 平铺库存清单 | 表头含 `Modello/Variante` + `TOTALE` | `flat_stocklist` |

分流顺序:**先试平铺(表头特征明确),不命中再走 col14**。
反过来会让 col14 文件被平铺解析器接走。

## 自检清单

- [ ] 块数 = 输出货号数
- [ ] 每个货号的尺码行 / 数量行对齐(数量行 ≠ 尺码行的 SKU 数量 → 错位)
- [ ] cost / retail 从描述里拆出来了,顺序没反
- [ ] 图片数 ≈ 块数(有些文件确实没图)
