# C+- Column: Integer-Only Rule (支持负数)

模板 `C+-` 列必须存**纯整数**,正数如 `18`、`26`、`30`,负数如 `-18` 保留 `-` 号。**不能是字符串 `"18"` 或带 `"C+"`/`"%"` 前缀**。

## ⚠️ 最高优先级:C+ 是「每单单独给」的(2026-10-28 用户确认)

> 原话:「是我单独给的,不用保存」

规则:

1. **每次转换以用户当次给的 C+ 值为准**(例:`--cplus 10` 表示 10%,批发价 = cost × 1.10)。
2. **不要**把某个 C+ 值写进「品牌对应C+.xlsx」之类的品牌查找表当品牌默认值。
3. **不要**在代码里硬编码品牌级 C+ 常量(`CPLUS_INT = 18`)当默认。
4. **不要**套用 `0.18` 之类的兜底加价率 —— 用户没给 C+ 时,C+ 列和批发价**都留空**。

因此:`get_cplus_pct()` / `get_cplus_int()` 的 `fallback` 默认值是 **None**(找不到就返回 None = 留空),
而不是历史上的 `0.18`。`compute_supply_price(cost, None)` 返回 None,两列一起留空。

## 用户确认规则 (2026-07-15)

| C+ 数值 | 写法 |
|---|---|
| 正数(常规) | `18`(整数,无 + 号) |
| 负数(让利/折扣) | `-18`(整数,保留 `-` 号) |
| 无 C+ 数据 | 留空(None/空字符串) |

## 错误 vs 正确

```python
# ❌ 错
"cplus_text": "C+18%"        # 带前缀
"cplus_text": "18"            # 字符串
"cplus_text": "+18"           # 不必要的 +
"cplus_text": format_cplus_text(0.18)  # shared 函数返回字符串 '18'

# ✅ 对
"cplus_text": 18              # 正数
"cplus_text": -18             # 负数(让利)
"cplus_text": int(cplus_pct * 100)  # → 18
"cplus_text": -int(abs(cplus_pct * 100))  # → -18 负数
```

⚠️ **shared 模块的 `format_cplus_text()` 返回字符串**,不能直接用作 `cplus_text` 字段。需要 `int(float(format_cplus_text(0.18)))` → `18`。

## 三种 Pattern(写适配器前先确认属于哪种)

### Pattern A — 品牌级整数常量

C+ 在品牌级别固定,所有行相同(Gucci、CHLOE 等):

```python
CPLUS_INT = 18  # 正数
# 或负数(让利):
# CPLUS_INT = -18

def _build_row(cost, ...):
    return {
        ...
        "cplus_text": CPLUS_INT,
        # 有 C+ 时算供货价,没 C+ 时留空
        "supply_price": round(cost * (1 + CPLUS_INT / 100), 2) if cost is not None else None,
    }
```

### Pattern B — 从品牌 C+ 表查询

C+ 按品牌/品类/性别从另一个文件查(FENDI、ARMANI):

```python
def _get_cplus_pct(brand: str) -> float | None:
    """从 品牌对应C+.xlsx 查 (品牌, 性别, 品类) → 折扣率 (e.g. 0.18 或 None)"""
    ...
    return pct  # float (e.g. 0.18) 或 None

cplus_pct = _get_cplus_pct(BRAND)  # 0.18 或 None
if cplus_pct is not None:
    cplus_int = int(round(cplus_pct * 100))  # 18
else:
    cplus_int = None  # 无 C+ 留空
```

### Pattern C — 每行 DESTINAZIONE 列

C+ 在源文件 DESTINAZIONE 列逐行写(BOTTEGA VENETA、ACNE、PANICALE col22):

```python
def parse_data_rows(ws, ...):
    for r in range(...):
        markup_val = ws.cell_value(r, dest_col)  # 'C+50' / 213 / '' / None
        if markup_val is None or str(markup_val).strip() == '':
            cplus_int = None
            cplus_num = 0
        else:
            s = str(markup_val).strip()
            if s.startswith('C+'):
                cplus_num = float(s[2:])  # 'C+50' → 50
                cplus_int = int(cplus_num)
            else:
                cplus_num = float(s)      # 213 → 213
                cplus_int = int(cplus_num)
        ...
```

⚠️ **`is_new_product` 必须用 `r != prev_r`**(按行号判断),不能用 `sku not in seen_skus`。BOTTEGA 同一 SKU 出现在多个独立行时,两行都应有 C+。

## supply_price 计算(无 C+ 留空,不要 0)

```python
def compute_supply_price(cost, cplus_int):
    """无 C+ 时留空(None),不填 0"""
    if cost is None or cplus_int is None or cplus_int == '':
        return None
    return round(cost * (1 + cplus_int / 100), 2)
```

| Pattern | supply_price |
|---|---|
| A | `compute_supply_price(cost, CPLUS_INT)` — CPLUS_INT 为 None 时返回 None |
| B | `compute_supply_price(cost, cplus_int)` |
| C | `compute_supply_price(cost, cplus_int)`,行级判断 |

⚠️ **原模板 Sheet1 的旧规则说"没有 C+ 默认 0"** → 用户 2026-07-15 确认**改为留空**,新规则优先。后续如果有 vendor 适配器把 supply_price 硬编码为 `cost` 或 `0`,需要修。

## 历史品牌 C+ 值(2026-07-15 记录,**仅供追溯,不要自动套用**)

> ⚠️ 下表是当时的品牌加价率快照。按 2026-10-28 规则,C+ **每单单独给**,
> 写适配器时**不要**从这张表里取默认值;用户当次给了什么就用什么。

| 品牌 | C+ | supply 公式 | 备注 |
|---|---|---|---|
| CHLOE | 16 | cost × 1.16 | 正 |
| ARMANI | (查 MU_SS26) | cost × (1 + pct/100) | 可正可负 |
| FENDI | (查 品牌对应C+.xlsx) | cost × (1 + pct/100) | 可正可负 |
| Gucci | 18 | cost × 1.18 | 正 |
| ALAIA | 18 | cost × 1.18 | 正 |
| BOTTEGA VENETA | 查品牌 C+ 表 | cost × (1 + pct/100) | 可正可负 |
| ACNE | 查品牌 C+ 表 | cost × (1 + pct/100) | 可正可负 |
| MAMMUT ATS | 15 | cost × 1.15 | 正 |
| MONCLER GRENOBLLE | 30 | cost × 1.30 | 正 |
| MIU MIU | 26 | cost × 1.26 | 正 |
| SALOMON | (查品牌 C+ 表) | cost × (1 + pct/100) | 可正可负 |
| VIVIENNE WESTWOOD | 18 | cost × 1.18 | 正 |
| PANICALE | (查 col22) | cost × (1 + num/100) | 可正可负 |

## 通用方法:用 shared 模块的 C+ lookup

`scripts/convert_vendor_quote_shared/__init__.py` 已实现 C+ lookup 函数,所有 vendor 适配器**直接调用**即可,不要每个适配器重复写。

```python
from convert_vendor_quote_shared import (
    get_cplus_pct,    # 0.18 / 0.23 / 0.26 浮点; 找不到 → None
    get_cplus_int,    # 18 / 23 / 26 整数(模板 C+- 列用); 找不到 → None
    compute_supply_price,  # cost × (1 + pct/100),无 cost/C+ 留空 None
)

# 品牌适配器里(注意:不传 fallback = 不套默认加价率)
cplus_pct = get_cplus_pct(BRAND)        # 0.18 或 None
cplus_int = get_cplus_int(BRAND)        # 18 或 None
supply_price = compute_supply_price(cost, cplus_pct)  # None-safe
```

**数据源查找顺序**:

1. 显式传入的 `cplus_file=`
2. 环境变量 `CPLUS_BRAND_MAP` 指向的 xlsx
3. 历史本地路径 `~/Desktop/表格处理/6.29/品牌对应C+.xlsx`(存在才用,不存在不报错)

```bash
export CPLUS_BRAND_MAP="/path/to/品牌对应C+.xlsx"
```

**表结构**:
| col 1 | col 2 |
|---|---|
| 品牌名(可大小写混乱,如 'Gucci' / 'ALAIA' / 'BOTTEGA VENETA ') | C+ 文本(可大小写混乱,如 'C+18%' / 'c+18%' / 'C+30%') |

**关键行为**:
- 品牌名做 `.upper()` 规范化后再查(`gucci` / `GUCCI` 都能命中)
- C+ 文本用 `re.search(r'C\+(\d+)%', text, re.IGNORECASE)` parse,大小写不敏感
- 查不到 → 返回 `fallback`(默认 **None = 留空**,不再自动套 0.18)
- `compute_supply_price` 在 cost 或 cplus_pct 为 None 时返回 None,不返回 0(用户 2026-07-15 规则:无 C+ 供货价留空)

**避免重复**:不要在适配器里写 `_load_cplus_lookup()` / `_get_cplus_pct()` 等私有函数,直接 import shared。

## 3 层优先级 Resolver(可选,前端传规则时用)

`scripts/convert_vendor_quote_shared/cplus_resolver.py` 提供 `CplusResolver`,按 **L1 行级覆盖 → L2 品类规则 → L3 品牌默认** 匹配:

- L1:overrides CSV(`sku,cplus` 两列)
- L2:YAML 规则列表,每条 3 个并列 keyword(`brand` / `gender` / `category`),任一为空 = 该维度不筛选;**列表顺序即优先级**
- L3:品牌默认值(调用方显式传入)
- 兜底:全不命中 → `None`(C+ 留空)

```python
from convert_vendor_quote_shared.cplus_resolver import CplusResolver

resolver = CplusResolver.from_files(rules_file=..., overrides_file=...)
cplus_int, label = resolver.resolve({"sku": "740237-V2X40", "brand": "BURBERRY", "gender": "DONNA", "category": "BORSE"})
# → (10, "L2:BURBERRY|DONNA|BORSE=+10")
```

`label` 是审计标签,建议写进日志便于回溯命中来源。
