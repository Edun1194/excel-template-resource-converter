# 期货通道(Futures Channel)

> 适用:供应商给的**期货 / 订货单**——还没有真实库存,数量列要留空等客户填。

## 什么是「期货通道」

期货单和现货单走**同一个解析器**,区别只在两处输出约定:

| | 现货通道 | 期货通道 |
|---|---|---|
| 数量列 | 从源文件读真实库存 | **留空**(等客户填) |
| 「云仓」列 | 工厂代号(如 `V1061`) | 工厂代号(`V1004-期货` → `V1004`)**或**品牌占位(`category-futures` → `品牌名称（期货）`) |

「云仓」列映射的是 `row["vendor_id"]`,见 `convert_vendor_quote_shared/__init__.py` 的 `map_row_to_template`。

### 已注册的期货通道

| Vendor ID | 适配器 | 说明 |
|---|---|---|
| `V1003-期货` | `vendor_v1003_futures.py` | 期货系总入口:McQueen SS27 + Armani PLM Report + LONGCHAMP SS27。输出 `vendor_id` 统一 `V1003` |
| `V1004-期货` | `vendor_v1004_futures.py` | SALOMON Foglio 版式数量留空;现货 4 品牌(ARC'TERYX / PATAGONIA / PRL / STONE ISLAND)委托 `vendor_v1004.parse_rows`,数量照旧 |
| `V1070-期货` | `vendor_v1070_nuorder.py` | BC NuORDER SS27/SS28 |
| `V1072-期货` | `vendor_v1072_in_arrivo.py` | LISTE FW26 IN ARRIVO 8 个品牌 |
| `category-futures` | `vendor_category_futures.py` | 聚合通道:一个适配器覆盖多品牌多版式,「云仓」写 `品牌名称（期货）` |

## category-futures:6 种版式(2026-10-27)

一个适配器覆盖 6 个品牌、6 种完全不同的表结构,按文件特征自动路由。
实现见 `scripts/category_futures/`。

**铁律:一律按表头名定位列,绝不写死列号。**
同一份 NuORDER,ACNE / BC / Men ss27 / POLO RL 四家的列序都不一样(19 / 21 / 55 / 29 列)。

| 模块 | 版式 | 表头特征(SIGNATURE) | 尺码/数量 |
|---|---|---|---|
| `linesheet.py` | LEMAIRE / Alexander Wang / GANNI | `Style Number` + `Available Sizes` + `Style Image` | 尺码在 **'Available Sizes' 文本列**,按逗号拆行;数量留空 |
| `nuorder.py` | ACNE / BC / Men ss27 / POLO RL | sheet 名含 `NuORDER` + `Season`+`Style Number`+`Wholesale` | **分两种**: 行级尺码(ACNE/BC/POLO RL,一行一个尺码)/ 网格型(Men ss27,列=尺码) |
| `boss.py` | BOSS | `Style No.` + `WHS` + `SKU` | 尺码取 **'Sizes' 文本列**(每个 SKU 自带尺码表),按逗号拆行;数量留空 |
| `salomon.py` | SALOMON | `Article code` + `Color` + `WHLS` | sheet `FTW`,R1 表头但 c14-c32 是**订货尺码矩阵**(R1 欧码 36/36.67…, R2 英码 3.5/4/4.5…);**期货单: 忽略数量, 每个 SKU × 全部尺码展开, 数量留空** |
| `longchamp.py` | LONGCHAMP | `Code Style` + `Tarif Gros` + `Qté` | sheet `QTE BDC`,尺码在固定字段区之后(c34);数量 `Qté` 空则取 `Qty`;**该文件无行级图片** |
| `vw.py` | VIVIENNE WESTWOOD | `Generic Prod Code` + `Available Qty` + `WHL` | sheet `Foglio1`,每行一个尺码(`Prod Size`);`---` 是一片装占位符 → **留空**;`Prod Brand` 是内部码不是品牌名;`Product Image` 是 `#VALUE!` 无图 |

### 通用规律(反复验证)

1. **按表头名定位列**,别名表要覆盖变体列:
   `'Wholesale (EUR)'` → 归一化后 `wholesaleeur`,所以别名要同时写
   `("Wholesale (EUR)", "Wholesale (USD)", "Wholesale Price (EUR)", "Wholesale Price", "Wholesale")`。
   归一化会**去括号**:去括号后才能对上 `'Wholesale (EUR)'`。
   ⚠️ **`header_col_map` 是「归一化后精确相等」匹配, 不是前缀匹配 —— 别名必须写全。**
   写成截断版会让字段**静默变空**(不会报错, 只会少值)。踩过的坑(SALOMON FTW):
   | 别名(错) | 源表头(实际) | 后果 |
   |---|---|---|
   | `Category BU P` | `Category BU Perimeter` | 类目全空 |
   | `Model Euro Re` | `Model Euro Retail Price €` | 零售价全空 |
   | `Unit total` | `Unit total (10 pcs min)` | 数量取不到 |
   | `Value total` | `Value total (€)` | Value 取不到 |

   改完别名**一定要重跑看字段是不是真填上了**, 别只看行数对不对。
2. **detect 的表头特征必须同时含全部关键词**才算命中。
   宽松匹配会让 Linesheet 和 NuORDER 互相误判——两者都有 `Style Number`。
3. **尺码区起点 = 最后一个已映射固定字段列 + 1**。
   否则 `Name` 这种 4 字母词会被误判成尺码。
   **尺码区右边界不要依赖别名匹配**: 表头变体多, 匹配不稳 —— 直接扫表头找
   `unittotal` / `valuetotal` 开头的列(汇总列)。
4. **数量列 = 尺码列右边最近的表头含 `Qty` 且不含 `size price` 的列**;没有就留空。
5. **尺码矩阵型版式: 尺码区是「每个尺码的订货格」, 要按尺码展开**(2026-10-30 SALOMON FTW 踩坑)。
   特征: 表头行下面还有一行尺码标签, 数据行里对应位置是**订货格子**。
   - 尺码标签取**第二行**(R2)—— 第一行常是换算参考值(SALOMON: R1 欧码 36/36.67… / R2 英码 3.5/4/4.5…)
   - **期货单: 忽略源里的数量**, 每个 SKU × **全部**尺码各出一行, 数量留空等客户填
     (源里只有部分 SKU 填了量 —— 152 个里 17 个填了, 填了���**尺码合计精确等于汇总列**,
      75=75/130=130/40=40 —— 但期货单的语义是全尺码供应, 客户自己填量)
   - 结果: 152 SKU × 19 尺码 = **2888 行**, 数量全空
   - ⚠️ 别按"源里哪些尺码有量"来筛 —— 那是现货/已下单单的逻辑, 期货单不适用

6. **期货统一口径(2026-10-30 用户拍板,适用全部 category-futures 版式)**:
   **忽略源里的数量,每个 SKU × 全部尺码各出一行,数量一律留空等客户填。**
   不要按"源里哪些尺码有量"筛 —— 期货单的语义是全尺码供应。
7. **尺码列的两种版式必须分开走(2026-10-30 Men ss27 踩坑)**:
   - **行级尺码**(尺码列只有 1 个): 尺码值**就在那一列里**, 数量取右边的 `Qty` 列。
     ACNE / BC / POLO RL 都是这种。
   - **网格型**(尺码列 > 1): 列 = 尺码, 行 = SKU, 单元格里的数字是**订货数量不是尺码**!
     Men ss27 是这种(c18=XXS … c34=4XL, 中间夹 `XXS size price` 这种单价列)。
     ⚠️ 网格型**绝不能**走"读尺码列的值当尺码"的老路径 —— 会把 c22 的 12 / c24 的 30
     当成尺码, 出来一堆 `12`/`30`/`38` 这种假尺码。
   - **网格型的尺码组怎么分(靠标签类型, 不能只看列间距)**:
     服装尺码组和鞋码组**都是**"尺码列隔一个 size price 列"排列, 列间距(间隔 ≤2)分不开。
     规则: 从**有数据的尺码段**向两侧扩展, 直到遇到**标签类型变化**(纯数字 ↔ 含字母)。
     Men ss27 实测: live = c22(S)/c24(M)/c26(L)/c28(XL)/c30(XXL)
     → 向左扩到 c18(XXS)、向右扩到 c34(4XL), 遇 c36('26' 纯数字) 停
     → 组 = c18-c34 共 9 个; 鞋码组 c36-c48 与 OS/S-M/L-XL 组整列无数据 → 不展开。
     结果: **205 数据行 × 9 尺码 = 1845 行**, 数量全空。
8. **尺码来源优先用「每个 SKU 自带尺码表」的文本列, 不是全表通用的网格**(BOSS 踩坑):
   BOSS 的 `Sizes` 列每个 SKU 自带尺码表(`S, M, L, XL, XXL, XXXL, 4XL, 5XL, 6XL` /
   复合码 `29/32, 29/34…` / 一片装 `PZ.`), 而那 82 列网格是全表通用的 ——
   只按"网格里有数"出尺码, 会让绝大多数 SKU 一个尺码都没有(实测 426 行里只有 23 行有量)。
   用 `Sizes` 文本列展开: **6037 行**, 与源 `Sizes` 的尺码集合 100% 一致。
9. **尺码占位符要清成留空**(VW 踩坑): `Prod Size` 的 `---` 是一片装商品的"没有尺码",
   原样输出 `---` 会让客户以为是个尺码。留空即可(`_SIZE_PLACEHOLDERS`)。
10. **`is_size_like` 要认双字母数字尺码**: `^\d+[A-Za-z]$` 漏了 `3XL`/`4XL`/`5XL`/`6XL`,
    要放宽成 `^\d+[A-Za-z]{1,2}$`(BOSS 有 4XL/5XL/6XL)。

## openpyxl 读不了的 xlsx(2 个已知)

| 文件 | 报错 | 原因 |
|---|---|---|
| Alexander Wang | `TypeError` | `<sheetProtection insertLinks="1">` |
| LONGCHAMP | `ValueError` | `<dataValidation type=非法值>` |

**解法**:`category_futures/common.py` 的 `open_safe_workbook()` —— 复制一份临时文件,
把 `<sheetProtection>` / `<dataValidations>` 标签摘掉再 load;还不行退回 `read_only` 模式
(read_only 能读单元格值,但拿不到 `ws._images`,图片再走 zip 解析)。

```python
from category_futures.common import open_safe_workbook

wb = open_safe_workbook(path)   # 别直接 openpyxl.load_workbook
```

## 数量列留空的正确姿势

期货单**没有真实库存**,不要从源文件里硬凑一个数出来:

```python
rows.append({
    ...
    "quantity": None,      # 留空,等客户填
    "vendor_id": vendor_id,
})
```

⚠️ 不要填 `0`。`0` 在模板里表示「确实没有库存」,语义和「还没填」不同。

## 「云仓」列写法

```python
# V1004-期货 → 写 "V1004"(不带 "-期货" 后缀)
rows[i]["vendor_id"] = "V1004"

# category-futures → 写固定占位
rows[i]["vendor_id"] = "品牌名称（期货）"   # 见 vendor_category_futures.CATEGORY_FUTURES_WAREHOUSE
```

## 自检清单

- [ ] 数量列**整列为空**(不是 0)
- [ ] 「云仓」列 = 工厂代号 / `品牌名称（期货）`
- [ ] 尺码已展开成多行(一行一个尺码)
- [ ] **尺码值里没有「数量」冒充的**(网格型: 不能出现 `12`/`30`/`38` 这种纯数字假尺码)
- [ ] 尺码占位符(`---` 之类)已清成留空
- [ ] 图片数与源文件一致(没有的版式如 LONGCHAMP/VW 确实就是 0)
- [ ] detect 没有误判(同一份文件只命中一个版式)

## 更新记录

| 日期 | 变更 |
|---|---|
| 2026-10-27 | category-futures 通道建立(6 种版式) |
| 2026-10-30 | SALOMON FTW 尺码矩阵修复(2888 行, 数量留空) |
| 2026-10-30 | **期货统一口径**: 忽略数量, 全部尺码展开, 数量留空 |
| 2026-10-30 | **Men ss27 网格型**修「把数量当尺码」(205×9=1845 行); 尺码组按标签类型划分 |
| 2026-10-30 | **BOSS** 改用 `Sizes` 文本列展开(474 → 6037 行); **VW** 清 `---` 占位符 |
