# V1015 适配笔记 —— 意大利线多品牌通道

**Vendor code**: V1015(工厂代码, 一个码覆盖多个品牌/版式)
**适配器路径**: `scripts/vendor_v1015.py`(主适配器, 内含 4 个子版式)
**最后更新**: 2026-11-06

> V1015 是个**多版式自动分流**的工厂代码 —— 下拉框里只选 "V1015",
> 适配器按文件特征自己认版式。

## 五个子版式(按 detect 顺序)

| # | 版式 | 判据 | 结构 |
|---|---|---|---|
| 0 | **APC ATS REORDER(补货单)** | 表头行含 `REFERENZA` + `VARIANT` | 款 → 尺码行 → 颜色行(+材质/产地行), **每段重复表头** |
| 1 | Ready to be shipped | R2 c24 含 `\n` 且 len>5 | per-row, c24 尺码码段 |
| 2 | BOS-MONCLER | R1 c3 == 'ModelloVariante' | 意大利线 |
| 3 | 怀远(HK Huaiyuan) | R1 c12 是短大写 token | R1 起是尺码码矩阵 |
| 4 | 老 V1015 (PRADA ITALY) | 其余 | 每产品 10 行, 尺码标签行在顶部 |

⚠️ **APC 的 detect 必须排在最前面** —— 它的表头在 R16,
而默认的 PRADA 路径会把 R16 当数据行 → 出一堆垃圾行。

---

## APC ATS REORDER(补货单) (2026-11-06 新增)

源: `A.P.C._ ATS FW26 {ACCESSORIES|MAN RTW|WOMAN RTW}.xlsx`
(A.P.C. 法国极简品牌, 意大利代理发的 ATS REORDER 表)

### 版式

```
单 sheet 'Worksheet', 19 列, **每段重复 R15(EUR)+R16(header)**(共 16 个表头行)

  R10  c8  = 'FW26 / REORDER MEN RTW'    ← 标题(季节/单据类型/性别)
  R13  c2  = 'OUTERWEAR'                 ← 段落大类
  R15  c5/c18/c19 = 'EUR'
  R16  header: STYLE(c2) REFERENZA(c3) VARIANT(c4) SRP(c5) TAGLIE(c6)
         QTà(c17) COSTO UNITARIO(c18) TOTALE(c19)
  尺码行: c2 = 款名('MAC AUSTER'), c6-c16 = 该款的尺码标签
  颜色行: c3 = 货号('WOZAA-H01299'), c4 = '<色码> <色名>', c5 = SRP,
          c6-c16 = 各尺码数量, c17 = QTà, c18 = 成本, c19 = TOTALE
  材质行: c3 = '70% RECYCLED WOOL, ...'
  产地行: c3 = 'Origin: BULGARIA'
```

### 关键坑

1. **这是补货单, 数量留给客户填**:
   实测三个文件里 `QTà`(c17) 和 `TOTALE`(c19) **全是 0**, 各尺码数量格(c6-c16)全空。
   → 尺码照常展开(每个尺码一行), **数量列留空**(源里有正数才填)。

2. **每段都重复表头**: 一个文件里有 16 个含 `REFERENZA` 的行 → 必须跳过。

3. **段落行 vs 尺码行**: 都是 "c2 有值, c3/c4 空", 靠 **c6-c16 有没有尺码标签**区分:
   - 有尺码 → 尺码行(c2 = 款名 → 描述)
   - 无尺码 → 段落行(c2 = 大类 → 类目: OUTERWEAR / DENIM / BELTS / SHOES / HATS / HOME ...)

4. **材质格有时是数字 0**(空值):
   往后找材质/产地行时, 遇到 0 要**跳过继续往下找** ——
   不能把它当材质, 也不能在那里停下(下一行才是产地行)。
   (踩过: 材质写成 '0'、产地漏掉)

5. **c4 = '<色码> <色名>'**, 首 token 是 3 个大写字母 = 色码:
   `'PIA HEATHER MARINE'` → 色码 `PIA` + 色名 `HEATHER MARINE`
   实测 120 个首 token 全是 3 字母码, 没有一个是颜色英文缩写 ✓
   ⚠️ **双色组合拆不动**(`'ECRU/KAKI'` / `'ROSE CHINE'`) → 色号留空, 颜色保留原值。

6. **货号要带色码**: `REFERENZA` 同款多色(241 个款 → 412 个颜色行),
   单靠 REFERENZA 不唯一 → 货号 = `<REFERENZA>-<色码>`
   (同 V1015-补货单2 MIUMIU ATS 的口径: 款码-色码)。
   双色组合行(无色码)用 `-<色值>` 兜底, 保证唯一。

7. **标题行给性别/季节**: `'FW26 / REORDER MEN RTW'` → Man / FW26;
   `'WOMEN RTW'` → Woman; `'ACC HOMME & FEMME'` → Unisex。
   ⚠️ 判断顺序: WOMEN → HOMME&FEMME → MEN。

8. **图片锚在 B 列的颜色行上**(不是尺码行), 410 张里 409 张在颜色行。

### 字段映射

| 模板列 | 源 |
|---|---|
| 品牌 | 固定 `A.P.C.`(源无品牌列; 文件名兜底会是 "A.P.C._ ATS FW26 (MAN RTW)" 太长) |
| 货号 | `<REFERENZA>-<色码>` |
| 描述 | 款名(尺码行的 c2, 'MAC AUSTER') |
| 类目 | 段落大类(R13 那种行) |
| 材质 | 材质行 c3 |
| 性别 | 标题行 |
| 颜色 | c4 拆出的色名 |
| 色号 | c4 拆出的色码(拆不开留空) |
| 季节 | 标题行 'FW26' |
| 产地 | 产地行 'Origin: BULGARIA' → 'BULGARIA' |
| 零售价 | c5 SRP |
| cost | c18 COSTO UNITARIO |
| C+- / 批发价 | 留空(C+ 由用户每单单独给) |
| 尺码 | 尺码行标签 |
| 数量 | 各尺码数量格(源里全空 → 留空) |
| 云仓 | V1015 |

### 实测

| 文件 | 行数 | 货号 | 性别 | 数量非空 | 带图行 |
|---|---|---|---|---|---|
| ACCESSORIES | **596** | 412 | Unisex | 0 | 589 |
| MAN RTW | **2056** | 411 | Man | 0 | 2043 |
| WOMAN RTW | **1814** | 426 | Woman | 0 | 1793 |

校验口径: 从源独立重算 (货号, 色, 尺码) 三元组 → 缺失 0 / 多出 0 / 字段不符 0;
图字节 100% 来自源; 行↔图绑定 0 不一致。
回归: V1015 原有 8 个文件逐行完全一致。

---

## 相关文件

- `scripts/vendor_v1015.py` — 主适配器(含 APC ATS REORDER 子版式)
- `scripts/vendor_v1015_restock.py` ~ `_restock4.py` — 独立的补货单2/3/4(不同工厂, 不同版式)
- `references/patterns/image-handling.md` — 图片嵌入 / 行高 (framework 通用)
