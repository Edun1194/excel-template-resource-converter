# template-resource-converter

把不同供应商的 SKU 报价单 / 库存单,**自动识别**并转换成固定模板。

- 输入:Excel(.xlsx / .xls)、PDF(lookbook)、Word(markup)
- 输出:统一模板 `template1-cn.xlsx`(18 列,中文)/ `template1-en.xlsx`(18 列,英文)
- 图片:源文件里的产品图**以内嵌图片对象**写进模板(不是链接文本),同一行多图会写多个图片对象

内含 **37 个内置供应商适配器**,覆盖多 Section linesheet / 平铺库存清单 / col14 块状 /
NuORDER / PDF 等版式族。

## 安装

```bash
git clone <this-repo> template-resource-converter
cd template-resource-converter
python3 -m pip install -r requirements.txt
```

也可以直接当 Mavis / Claude Code 的 skill 用 —— 目录结构就是 skill 标准布局
(`SKILL.md` + `references/` + `scripts/` + `assets/`)。

## 快速开始

```bash
python3 scripts/convert_vendor_quote.py \
  --primary-file "path/to/vendor-stock.xlsx"
```

按文件名 / 内容自动识别 vendor,输出文件默认落在源文件同目录。

### 指定 vendor + 本次订单的 C+

```bash
python3 scripts/convert_vendor_quote.py \
  --primary-file "GUCCI FW26 - WAREHOUSE B.xlsx" \
  --vendor "V1061" \
  --cplus 10 \
  --output "converted/GUCCI_B.xlsx"
```

`--cplus 10` = 加价 10%,批发价 = `cost × 1.10`。

### 英文模板

```bash
python3 scripts/convert_vendor_quote.py \
  --primary-file "path/to/vendor-stock.xlsx" --language en
```

完整参数见 `python3 scripts/convert_vendor_quote.py --help`。

## 模板字段

中文模板 18 列(英文模板列名一一对应):

```
品牌 | 图片 | 货号 | 描述 | 类目 | 材质 | 性别 | 颜色 | 色号
季节 | 产地 | 零售价 | cost | C+- | 批发价 | 尺码 | 数量 | 云仓
```

- `C+-` 存**纯整数**(正数 `18`、负数 `-18`),没有 C+ 就留空
- `批发价` = `cost × (1 + C+-/100)`,没有 C+ 就留空(不填 0)
- `云仓` 是供应商 / 仓库 ID

## 目录结构

```
.
├── SKILL.md                      # skill 主文档(完整规则、踩坑模式)
├── README.md
├── requirements.txt
├── assets/templates/             # template1-cn.xlsx / template1-en.xlsx
├── scripts/
│   ├── convert_vendor_quote.py   # 主入口 + dispatcher
│   ├── convert_vendor_quote_shared/   # 共享包(输出/图片/模板校验/C+)
│   │   ├── flat_stocklist.py     #   平铺库存清单版式
│   │   ├── col14_parsers.py      #   col14 块状版式
│   │   └── cplus_resolver.py     #   C+ 3 层优先级 resolver
│   ├── category_futures/         # 期货通道的 6 个版式解析器
│   ├── analyze_vendor_format.py  # 新供应商结构分析
│   └── vendor_*.py               # 各供应商适配器
├── outputs/                      # 新适配器骨架模板
└── references/
    ├── patterns/                 # 通用技术/踩坑模式
    └── vendors/                  # 少量品牌专属笔记
```

## 加一个新供应商

1. **分析结构** — `python3 scripts/analyze_vendor_format.py <file>`
2. **对号入座版式族** — 先看 `SKILL.md` 的「版式族」表;命中已有版式就直接复用 shared 解析器
3. **写适配器** — 复制 `outputs/generated_adapter_vxxxx.py`,实现 `parse_rows()` / `get_vendor_meta()`
4. **注册** — `scripts/convert_vendor_quote.py` 的 `BUILTIN_VENDORS` 加一条
5. **验证** — 行数 / 图片数 / 尺码集合 / 数量合计 / 货号格式 / C+ 是否留空

细节见 `SKILL.md` 的「Adaptation Workflow(新供应商 5 步)」。

## 几条容易踩的规则

- **按表头名定位列,绝不写死列号** —— 同一份 NuORDER,不同品牌的列序完全不同
- **C+ 是每单单独给的** —— 不要写品牌默认值,不要套默认加价率,没给就留空
- **缺数据就留空,不要推断** —— 源文件没有的字段不要自己补
- **期货单数量留空 ≠ 填 0** —— `0` 是「没库存」,留空是「还没填」
- **图片按 anchor 行号取,行号是 0-idx** —— 数据行号 r(1-idx) → anchor r-1

更多见 `SKILL.md` 的「Generic Bug Patterns」和 `references/patterns/`。
