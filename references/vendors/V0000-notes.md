# V0000 适配笔记 —— 多品牌占位通道 (MOORER / GUCCI STOCKLIST / HERNO)

**Vendor code**: V0000 (临时, 等待正式 V10XX 分配)
**适配器路径**: `scripts/vendor_v0000.py`(路由入口) + `scripts/vendor_v0000_herno.py`(HERNO 版式)
**注册方式**: 走 `scripts/` fallback, 不入 `BUILTIN_VENDORS`
  (dispatcher 找不到 `BUILTIN_VENDORS["V0000"]` 时, 会找 `scripts/vendor_v0000.py`)
**最后更新**: 2026-10-30

> ⚠️ **V0000 是个"占位通道"**: 三家品牌版式完全不同, 按文件特征自动分流。
> 将来拿到正式 vendor code 后, 建议拆成独立的 `vendor_v10XX.py`。

## 三个版式的分流顺序

`vendor_v0000.parse_rows()` 按顺序判断, **先到先得**:

| 顺序 | 判据 | 走哪个解析器 |
|---|---|---|
| 1 | 任意行 col 6 == `ARTICLE_COLOR` | `_parse_gucci_stocklist` (GUCCI STOCKLIST) |
| 2 | 表头行同时含 `SKU` + `BRAND` | `vendor_v0000_herno.parse_rows` (HERNO) |
| 3 | 其余 | `_parse_moorer` (MOORER) |

> 判据 1 必须排在前面 —— GUCCI 的老版本式表头里 col 6 也叫 `SKU`。

---

## 一、HERNO (2026-10-30 新增)

源: `HERNO FW 29-10.xlsx` — HERNO, 意大利高端成衣/配饰品牌
单 sheet `Worksheet`, R2 表头, R3+ 数据。实测 300 行 × 33 列 / 298 数据行 / 298 张内嵌图。

### 版式

```
R1:  只有 c14 = 2560 (残留值, 正好是全表 TOT.QTY 合计 —— 供应商放错位置了)
R2:  表头
  c1   IMAGE          空列(真正的图在 drawing 里, 不读这一列)
  c2   SKU            货号, 例 'GI000023X12732_9018'
  c3   BRAND          品牌 (全是 'Herno')
  c4   PRODUCT        产品名 → 描述 ('Asymmetric jacket with collar')
  c5   GROUP          类目 ('Clothing' 291 / 'Accessories' 7)
  c6   SEASON         季节 (全是 'FW')
  c7   GENDER         性别 (Female 162 / Male 128 / W 6 / M 1 / UNI 1)
  c8   COLOR          颜色 (意大利语色名, 31 种)
  c9   COMPOSITION    材质 ('100% PA')
  c10  DESC_EN        英文长描述(营销文案, 太长, 不用)
  c11  MADE_IN_CODE   产地 (IT 184 / RO 64 / BG 14 / MD 14 / AM 12 / TR 4 / BA 4 / CN 2)
  c12  RETAIL PRICE   零售价
  c13  TOTAL RETAIL   = 零售价 × 总件数 (冗余列, 不用)
  c14  TOT.QTY        总件数 (= 尺码格合计)
  c15  TU             总件数, **只有 8 行填了** (见「8 行特例」)
  c16-c33  尺码数量矩阵  表头是尺码 (XS/S/M/L/XL + 36…60 数字码)
```

### 关键点

1. **源文件没有 cost / 供货价列** → `cost` 和 `批发价` **一律留空**。
   不凭空造, 也不套默认值。C+ 由用户每单单独给 —— 但没有 cost, 给了也算不出批发价。
2. **源文件没有色号列** → `色号` 留空。
   ⚠️ SKU 末尾 4 位数字(`9018`)**不是色号**: 102 个末尾码里有 35 个同时对应多个颜色名
   (`1000` → {bianco, beige, panna}), 它是款式内部编号。
3. **表头行要扫, 不能写死 R2** —— R1 有个残留的 2560。
4. **性别要归一**: 源混用全称和缩写(`Female` 和 `W` 都有) →
   Female/W→Woman, Male/M→Man, UNI→Unisex。
   ⚠️ `_norm()` 返回**小写**, 映射表 key 是**大写** → 查表前要 `.upper()`(踩过: 性别整列空)。
5. **产地**: 源是 ISO 3166-1 alpha-2 国家码 → 英文国名
   (IT→ITALY / RO→ROMANIA / BG→BULGARIA / MD→MOLDOVA / AM→ARMENIA /
   TR→TURKEY / BA→BOSNIA AND HERZEGOVINA / CN→CHINA)。
   映射表里没有的码原样输出, 不丢。
6. **尺码区 = 表头不在固定字段里的非空列**。
   ⚠️ 比对要用**归一化后**的表头名 —— 直接拿原始表头比对会漏,
   导致 `RETAIL PRICE` / `TOTAL RETAIL` / `TOT.QTY` 被当成尺码列,
   凭空多出 900 行(踩过)。

### 「8 行特例」(有总件数但没填尺码明细)

有 8 个 SKU 的 `TOT.QTY > 0` 但尺码矩阵整行全空, 且这 8 行的 `TU` 值与 `TOT.QTY`
**完全相等**(1/3/3/5/6/6/10/74, 合计 108 件)。

跟框架惯例(qty<=0 的行跳过)不一样, 这里**保留**: 输出一行「**尺码留空 + 数量=TOT.QTY**」。
理由: 数量是源里明写的, 尺码源里没写 —— 留空比整行丢掉 108 件真实库存更忠实。
尺码为空的行在模板里是合法的。

### 实测

| 指标 | 值 |
|---|---|
| 输出行数 | **617** = 609 个 (尺码, 数量) 对 + 8 个无明细行 |
| 唯一货号 | 298 (= 源全部 SKU, 一个不丢) |
| 数量合计 | **2560** = 源 TOT.QTY 合计 ✓ |
| 零售价 × 数量 | **1978161** = 源 TOTAL RETAIL 合计 ✓ |
| 图片 | 617 行**全部带图**; media 298 个(去重), **字节与源 md5 100% 一致** |
| 行↔图绑定 | 617/617 **100% 一致** |
| cost / 批发价 / 色号 | 全空(源没有) |

> 注: 源文件 media 是 `.png` 扩展名但内容其实是 JPEG, 产物里会被写成 `.jpeg`
> —— 字节没变(md5 一致), 只是扩展名更正了。

### ⚠️ CLI 在 `配置器测试文件/V0000/` 里跑会被外部脚本抢先

dispatcher 找 adapter 的顺序是:
`BUILTIN_VENDORS` → `primary_file.parent/convert_<vendor>.py` → `scripts/vendor_<vendor>.py`。

`配置器测试文件/V0000/convert_V0000.py`(2026-09-23 的一次性 GUCCI 脚本)会**抢先于**
app 端模块被加载 —— 而它只认 GUCCI 的 ARTICLE 表头, 不处理 HERNO 的话
**在这个目录里跑 CLI 会静默出 0 行**。

已在该脚本里加了一段: 检测到 HERNO 就委托 app 端 `vendor_v0000_herno.py`。
(2026-10-30 修)

---

## 二、GUCCI STOCKLIST 220926 (2026-09-23)

源: `GUCCI STOCKLIST - 220926.xlsx`, sheet `Foglio1`, 7486 行 × 31 列。

- 文件按 **section** 切分, 每个 section 顶部是 `ARTICLE` 表头, 下面 1 行数据, 然后 ~13 行空白。
  一共 507 个 section, **11 种不同的尺码表**(`U` 单码 / XXS…XXL / 数字 EU / 童鞋带 `+` / 腰带 65-120cm …)。
- 每个 section 的尺码表**各自不同**, 不能用一个全局默认尺码表。
- col 6 `ARTICLE_COLOR` 是 SKU (article + "_" + color_code), col 9 零售价, col 10 cost。
- 数量: 各尺码格的数量, **qty <= 0 的行跳过**(framework `is_positive_qty` 惯例)。
- 仓库 ID = V0000, C+ = 18%。

> 另有一份独立脚本 `配置器测试文件/V0000/convert_V0000.py`(2026-09-23),
> 带 C+18 / SEASON=FW26 / 用户 18 列模板 writer, 与 app 端 `vendor_v0000.py` 的 GUCCI 分支
> 行为**不完全一样**(app 端 C+ 留空、season 默认 FW24)。CLI 跑 GUCCI 时外部脚本优先。

## 三、MOORER (2026-09-15)

源: `MooRer R-75 special list.xlsx` + `Moorer R-75 Special list-update 1809 (2).xlsx`。
单 sheet `FILE`, R1 中文表头, 59-61 列。

- 列顺序两个版本不同(R-75 是 性别|类目|**图片**|描述|… ; 1809 版是 性别|类目|描述|颜色|**图片**|…)
  → **按 header 名找列**, 不写死列号。
- col「国际编码」是 SKU (23 字符 `MOUBO100118TEPA663U0237`), **末尾 5 位是色号**
  (`COLOR_CODE_RE = [A-Z]\d{4}$`)。
- **货号拼接**(2026-09-15 user 要求): 有颜色名 → `<原SKU>---<颜色名>`;
  没颜色名但有色号 → `<原SKU>---<色号>`; 都没 → 原样。
- 尺码列 = R1 中 cost 列之后的所有非空表头, **但要排除 `总数`/`total`/`合计`/`sum`**
  (最后一列是总数, 不是尺码)。
- 类目是意大利语 (`BOMBER`/`GIACCHE`/`MAGLIERIA`/…), `CATEGORY_MAP` 维护 40+ 映射,
  **未映射的留原意大利语**。
- 性别: Men/Man/Mens→Man, Women/Woman/Womens→Woman。
- 图片: col「图片」是 None 或 `Immagine non trovata` 文本, 真正的图从 `ws._images` 取。

---

## 相关文件

- `scripts/vendor_v0000.py` — 路由入口 (GUCCI / MOORER)
- `scripts/vendor_v0000_herno.py` — HERNO 版式 (2026-10-30)
- `app.py` `VENDOR_CODES` — 里有 `"V0000"`
- `references/patterns/image-handling.md` — 图片嵌入 / 行高 (framework 通用)
