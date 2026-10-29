# 平铺库存清单版式(Flat Stocklist)

> 实现:`scripts/convert_vendor_quote_shared/flat_stocklist.py`
> 适用:GUCCI `FW26 - WAREHOUSE A/B.xlsx`、CANADA GOOSE `FW26 - WAREHOUSE C.xlsx`

同一种「意大利语库存清单」版式,**一行一个尺码**、尺码铺成横向矩阵。
V1038 / V1061 / V1062 共用这个解析器,只差仓库 ID。

## 版式结构

```
R1  : 标题 "GUCCI — Stocklist FW26"        → 品牌 + 季节
R2  : 空
R3  : 表头行
       c1 Foto | c2 Modello/Variante | c3 Gruppo | c4 Descrizione | c5 Colore
       | c6 MADE IN | c7 Acquisto € | c8 Vendita € | c9 TOTALE | c10..= 尺码列
R4+ : 数据行 与 段落标题行 交替
```

**段落标题行**(只有 c2 有值,c3 空):

| 源值 | 含义 |
|---|---|
| `Donna (38 articoli/varianti)` | 性别段 → `WOMAN` |
| `Uomo (38 articoli/varianti)` | 性别段 → `MAN` |
| `Abbigliamento` / `Calzature` / `Accessori` | 类目段(挂在当前性别下) |
| `TOTALE PEZZI GUCCI` | 合计行 → 跳过 |

**数据行判据**:`c2 货号` 和 `c3 Gruppo` **都有值**。
这比「看起来像 SKU」稳。

## 字段映射(源 → 模板,一一对应不造值)

| 源 | 模板 | 说明 |
|---|---|---|
| `Gruppo` (c3) | 描述 | `DOWN JACKETS` / `SNEAKERS` / `TIES` … |
| 段落大类 | 类目 | `Abbigliamento` / `Calzature` / `Accessori` |
| `Descrizione` (c4) | 材质 **或** 描述 | 取决于下面的语义判定 |
| 标题里的 `FW26` | 季节 | 正则 `(FW|SS|AW|CW)\s*\d{2}` |
| 源无色号列 | 色号 | 留空 |
| `MADE IN` (c6) | 产地 | 占位值 → 留空 |
| `Acquisto €` (c7) | cost | |
| `Vendita €` (c8) | 零售价 | |
| 尺码列表头 | 尺码 | **用原值**(见下方归一化) |
| 尺码格的值 | 数量 | 一行一个尺码展开成多行 |

## ⚠️ 三条内容驱动规则(不按品牌硬编码)

### 1. c4 `Descrizione` 的语义判定

同一版式,不同品牌往这列填的东西不一样:

- **材质成分**(GUCCI):`70% CO 30% PL`
- **商品描述**(CANADA GOOSE):`CHILLIWACK PUFFER - TD Black - Noir`

判据:**整份文件看多数数据行的 c4 是否含 `%`**。

```python
# 多数数据行 c4 含 '%' → 材质模式
材质模式  : 描述 = Gruppo(c3), 类目 = 段落大类, 材质 = c4
描述模式  : 描述 = c4,            类目 = Gruppo(c3), 材质留空
```

实测:GUCCI A 76/76 行含 %,GUCCI B 45/45 行含 %,CANADA GOOSE 0/7 行含 %。

### 2. 颜色列被源填成描述 → 留空

`c5 == c4` 时,说明源把整段描述复制进了颜色列,**该行没有独立颜色值** → 留空。
(CANADA GOOSE 7 行里有 3 行是这种情况。)

### 3. 产地占位值 → 留空

`Non indicato`(意大利语「未填写」)等 → 产地空。
(CANADA GOOSE 全部 7 行 `MADE IN` 都是 `Non indicato`;GUCCI 是 `ITALIA`/`ITALY`,不受影响。)

占位值白名单:`non indicato` / `non indicata` / `non specificato` / `n/d` / `nd` / `n.a.` / `-` / `--` / `/` / `null` / `none` 等。

## 尺码标签归一:半码

**同一份文件里半码有两种写法**:

- WAREHOUSE B 的 2-11 鞋码段用 ASCII 斜杠:`'7/'` `'8/'`
- 同一文件另一段用标准字符:`'36½'` `'37½'` `'38½'`
- WAREHOUSE A 同位置是 `'7½'`

→ `'/'` 在这个版式里就是半码,**有源内证据,不是猜**:

```python
_HALF_SLASH_RE = re.compile(r"^(\d+(?:\.\d+)?)/$")

def normalize_size_label(label: str) -> str:
    m = _HALF_SLASH_RE.match((label or "").strip())
    return f"{m.group(1)}½" if m else (label or "").strip()
```

⚠️ **尺码标签必须用原值**,表头归一化(`_norm_hdr`)只用于匹配固定字段名。
否则 `'37½'` 会被抹成 `'37'`。

```python
def _norm_hdr(value) -> str:
    """只用于匹配固定字段名;'Acquisto €'→'acquisto', 'MADE IN'→'madein'"""
    return re.sub(r"[^0-9a-z]+", "", str(value).strip().lower())
```

## detect 判据

必须**同时**命中 `'Modello/Variante'` 和 `'TOTALE'` 才认得是这个版式:

```python
if "modellovariante" not in cols or "totale" not in cols:
    continue
```

尺码列 = `TOTALE` 列右边所有有值的表头。

## 图片

用 `load_images_by_row_per_sheet`,anchor 在 c1 `Foto` 列。
**0-idx 行号 = 数据行号 - 1**。同一货号的多个尺码行共用同一张图
(所以图片数会少于输出行数,这是正常的)。

## V1038 / V1061 / V1062 的区别

| Vendor | 文件 | 版式 | 云仓 |
|---|---|---|---|
| `V1038` | GUCCI FW26 - WAREHOUSE A | 平铺 | `V1038` |
| `V1061` | GUCCI FW26 - WAREHOUSE B | 平铺 | `V1061` |
| `V1062` | GUCCI WAREHOUSE B/C(老单) | **col14 块状** | `V1062` |
| `V1062` | CANADA GOOSE FW26 - WAREHOUSE C | 平铺 | `V1062` |

⚠️ **V1062 同时管两种版式**,按内容自动分流:
平铺 → `flat_stocklist.parse_rows`;col14 → `col14_parsers` 的 v1038 parser。

⚠️ **V1061 vs V1062 都是 GUCCI WAREHOUSE B**,但 V1061 是平铺库存清单(新单),
V1062 是 col14 块状版式(老单)。选错 vendor 会输出完全不同的结果。

## 版式不匹配时直接抛错

V1061 在文件不是平铺版式时**直接抛 ValueError**(中文提示 + 指路
「GUCCI WAREHOUSE A 的平铺表用 V1038」),不静默输出空表——
选错 vendor 能立刻看出来。

## 自检清单

- [ ] 输出行数 = 各尺码数量 > 0 的格数之和
- [ ] 数量合计 = 源文件 `TOTALE` 合计
- [ ] 尺码集合与源表逐格一致(归一后)
- [ ] 品牌 / 季节从标题抽对了
- [ ] 性别段、类目段切分正确
