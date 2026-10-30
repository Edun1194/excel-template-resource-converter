#!/usr/bin/env python3
"""为每个 vendor 重新生成 _missing_fields.md 报告。

区分 3 类:
- ✗ 完全空 - 可能需要修
- · 部分空 - 源数据部分缺失,通常可接受
- ✓ 全部填 - OK
"""
import re
import sys
import subprocess
from pathlib import Path
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
TEST_DIR = Path('/Users/superlist/Desktop/表格处理/配置器测试文件')

VENDORS = [
    ('V1003', ['GUCCI new ready goods.xlsx']),
    ('V1004', ['POLO RALPH LAUREN 15.07.xlsx', 'PATAGONIA 17.07.xlsx', 'STONE ISLAND 17.07.xlsx']),
    ('V1015', ['ITA_Prada.xlsx']),
    ('V1021', ['Gucci.xlsx', 'Hogan.xlsx']),
    ('V1039', ['ADIDAS.xlsx', 'BURBERRY.xlsx']),
    ('V1041', ['MEN SS26.xls', 'WOMEN SS26.xls']),
    ('V1042', ['BRUNELLO CUCINELLI.xls', 'PRADA.xls']),
    ('V1043', ['FW25 MAN 09.02.26.xlsx', 'FW25 WOMAN 09.02.26.xlsx']),
    ('V1044', ['SS2026-10_02_26.xlsx']),
    ('V1046', ['SS26 110226.xlsx']),
    ('V1048', ['BOTTEGA VENETA.xlsx', 'BURBERRY.xlsx']),
    ('V1049', ['STOCKLIST SS26 23.02.xlsx']),  # WPS 把 .xls 转成了 .xlsx
    ('V1050', ['MAISON MARGIELA PELLETTERIA U+D SS26_18.02.26.xlsx']),
]

# 字段列定义（按模板列号）
COLS = [(1, '品牌'), (2, '货号'), (4, '描述'), (5, '类目'), (6, '材质'),
        (7, '颜色'), (8, '季节'), (9, '产地'), (10, 'cost'), (13, '零售价'),
        (14, '尺码'), (15, '库存')]

# 每个 vendor 的已知问题(从源结构分析)
KNOWN_ISSUES = {
    'V1041': """### 已知问题（数据源限制,非 adapter bug）
- **描述/颜色/产地**: 源文件无这些列(只有 Image/Brand/Style number/Season/Sex/Composition/供货价/Retail/COST/Size/Qty)
- **类目**: 源无专门 category 列,已用 Sex (Man/Donna) 填充
- **图片**: 源是 .xls,本程序不抽图(WPS/Excel 转 .xlsx 后能抽)
""",
    'V1042': """### 已知问题（数据源限制,非 adapter bug）
- **颜色**: 源数据是 col 10 单行描述 "Camicia Donna Bianco 257,00",颜色(Bianco)在描述里,未单独拆
- **产地**: 部分行源数据就没填
- **图片**: 源是 .xls,本程序不抽图
- **备注**: 实际版式(SKU col 10 + 描述 col 10 单价格 + 尺码 col 9-20 + qty col 13+)跟当前 adapter 假设的"双价格"版式不一致,部分行解析可能错位
""",
    'V1015': """### 已知问题（数据源限制,非 adapter bug）
- **材质**: 源无材质列(意大利线只标描述,材质从描述"Nylon+scamosciato"等推断未做)
- **季节**: 源无季节列(意大利线按 SS26 推断未做)
- **产地**: 源是 ITA_Prada 意大利线,统一填 ITALY(实际可能不是意大利产,源未标)
- **货号**: PRADA 内部 20 位 SKU 编码,无款号/颜色拼接需要(已唯一)
- **图片**: 每产品 3 张(同角度不同裁剪?),全嵌入
- **尺码**: 意大利鞋码 5-14(整数 + 半码),已归一化为 "5"/"5.5" 格式
- **同 SKU 重复出现**: 源文件中 `2EE443FG0003LF5F0632` 在两处不同 header 行下,各带不同 qty(疑似两批次入库),已保留多行
""",
    'V1003': """### 已知问题（数据源限制,非 adapter bug）
- **品牌**: 源 1 个品牌 GUCCI(从源 col 11 读);模板层 5 级优先级保留源 BRAND
- **季节**: 源无 season 列(只有 Department=WOMAN/MAN)
- **类目**: 用源 col 20 (Category, 英) 如 DRESSES / SNEAKERS;col 14 Description 作为描述列
- **货号**: 7 位 IDArticolo(整数型,如 1070893),已唯一(同款不同色靠 Variante 列区分)
- **图片**: 1248/1348 行有图(100 行源无图)
- **尺码**: 9 种 Scalarino 尺码表(A/C/F/G/H/I/M/T/U)由 col 22 自动选用 — A 衣服欧码、C 牛仔、F/G 美/英鞋码、H XXXS-XXXL、I 腰带、M 衬衫、T 单码 TU、U 鞋
- **col 68-74**: Totale 公式、qty 1676、日期等 summary 列(已排除,不算尺码)
- **Parametri sheet**: Sirio DB 连接参数,本 adapter 不读
""",
    'V1043': """### 已知问题（数据源限制,非 adapter bug）
- **材质**: 部分行源数据 MATERIAL 字段就是空的
- **产地**: 源无 MADE IN / ORIGIN 列
- **类目**: 改用源 COLLECTION 列(COL)填充
- **图片合并单元格**: 同一产品多尺码行的图片在合并单元格(A11:A13 等),用 load_images_by_merged_range 自动扩展到所有行
""",
    'V1044': """### 已知问题（数据源限制,非 adapter bug）
- **颜色**: 部分行源 COLOR DESC 字段就是空(供应商没填)
- **产地**: 源无 MADE IN / ORIGIN 列(整列缺失)
- **货号**: `MODEL-{COLOR}` 拼接(每色一码,如 `B90886-AUZ`),用 `-` 分隔 — 解决了"同 MODEL 多色导致尺码错乱"的问题
- **尺码**: 部分段用自定义尺码表(header_row 写 44/46/48/50... 替代默认 XXXS/XXS/XS),已用 `_build_size_label_map` 自动识别继承
""",
    'V1046': """### 已知问题（数据源限制,非 adapter bug）
- **材质/季节/产地**: 源数据列定义是 FOTO/MARCHIO/VOCE/ARTICOLO/MODELLO/RIGA + 尺码 + 成本价 + 零售价,无这些字段
- **描述**: 源无单独的 description 列(VOCE 已作为类目使用,如 "JEANS -U")
- **货号**: 用 RIGA 列(每色一码,如 4030465/4030475/4030485),不再是 ARTICOLO
- **图片**: 之前 0 张(因用了 V1051 专用 load_sheet_embedded_images,没 SKU in descr),改用 load_images_by_row_per_sheet 后每行都嵌入
""",
    'V1048': """### 已知问题（数据源限制,非 adapter bug）
- **描述**: 部分行源数据描述为空
- **材质/季节/产地**: 源 layout 复杂(header 在 R12,数据从 R13 开始),adapter 用固定列偏移读不到
- **图片修复**: 之前 0 张,改用 load_images_by_row_per_sheet 后 499/546(BV)/964/1025(BUR) 嵌入
""",
    'V1049': """### 已知问题
- **产地/材质**: 源无这些列
- **C+/供货价**: `assets/resources/V1049 POMPEU/MARKUP POMPEU.docx` 缺失 — 用户输入(CLI `--cplus`)后正常填充
- **图片**: 每个产品 3-4 张图,都嵌入;图片列加宽到 70 字符(默认 35)便于多图显示
""",
    'V1050': """### 已知问题（数据源限制,非 adapter bug）
- **产地**: 源无 MADE IN / ORIGIN 列
- **尺码**: 源是 TOT PZ 总量字段,无 size 列(本产品类目本身是 UNICA 单码,不需要尺码)
""",
}

for vendor, files in VENDORS:
    vdir = TEST_DIR / vendor
    out_dir = vdir / 'converted'
    out_dir.mkdir(exist_ok=True)
    lines = [
        f"# {vendor} 字段缺失检查",
        "",
        "**说明**: 品牌/C+-/供货价 由用户输入,不在检查范围。",
        "",
    ]
    for fn in files:
        src = vdir / fn
        if not src.exists():
            # V1049 旧版记录的是 .xls,WPS 转 .xlsx 后,fallback 试一下
            alt = vdir / fn.replace('.xls', '.xlsx')
            if alt.exists():
                src = alt
            else:
                continue
        stem = src.stem
        out = out_dir / f'{stem}_converted.xlsx'
        if not out.exists():
            lines.append(f"## {stem}")
            lines.append("")
            lines.append("**❌ 转换失败** — 文件可能未转换,或依赖文件缺失")
            lines.append("")
            continue
        wb = load_workbook(out)
        ws = wb.active
        total = ws.max_row - 1
        lines.append(f"## {stem}")
        lines.append("")
        lines.append(f"行数: {total}")
        lines.append(f"嵌入图片: {len(ws._images)} (按 92pt 行高缩放,字节保留)")
        lines.append("")
        for c, name in COLS:
            filled = sum(1 for r in range(2, ws.max_row + 1)
                         if ws.cell(r, c).value not in (None, "", 0))
            mark = '✓' if filled == total else ('·' if filled > 0 else '✗')
            lines.append(f"  {mark} {name}: {filled}/{total}")
        lines.append("")
    if vendor in KNOWN_ISSUES:
        lines.append(KNOWN_ISSUES[vendor])
    (vdir / 'converted' / '_missing_fields.md').write_text('\n'.join(lines), encoding='utf-8')
    print(f'  [OK] {vendor}/converted/_missing_fields.md 重新生成')
