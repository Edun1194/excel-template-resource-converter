#!/usr/bin/env python3
"""批量诊断所有 vendor adapter 的字段填充情况。

检查每个 vendor 的"应有字段"实际填了多少(行级填充率),并区分:
- 源数据缺失(无法修)
- 代码 bug(可修)
- 报告数错了(图片是嵌入对象,cell value 是 "")

输出到 stdout,可重定向到文件。
"""
import sys
import os
import subprocess
from pathlib import Path
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(__file__))
SCRIPTS = Path(__file__).parent

# 本地测试数据目录。默认用仓库相对路径 testdata/,
# 也可用环境变量 TEST_DATA_DIR 指向本机的「配置器测试文件」目录。
TEST_DIR = Path(os.environ.get("TEST_DATA_DIR") or (SCRIPTS.parent / "testdata"))

VENDORS = [
    ('V1039', ['ADIDAS.xlsx', 'BURBERRY.xlsx']),
    ('V1041', ['MEN SS26.xlsx', 'WOMEN SS26.xlsx']),
    ('V1042', ['BRUNELLO CUCINELLI.xlsx', 'PRADA.xlsx']),
    ('V1043', ['FW25 MAN 09.02.26.xlsx', 'FW25 WOMAN 09.02.26.xlsx']),
    ('V1044', ['SS2026-10_02_26.xlsx']),
    ('V1046', ['SS26 110226.xlsx']),
    ('V1048', ['BOTTEGA VENETA.xlsx', 'BURBERRY.xlsx']),
    ('V1049', ['STOCKLIST SS26 23.02.xlsx']),
    ('V1050', ['MAISON MARGIELA PELLETTERIA U+D SS26_18.02.26.xlsx']),
]

# 18 列模板的列号(2026-09-09 跟 template1-cn.xlsx 新表头同步)
# 品牌(1) 图片(2) 货号(3) 描述(4) 类目(5) 材质(6) 性别(7) 颜色(8) 色号(9)
# 季节(10) 产地(11) 零售价(12) cost(13) C+-(14) 批发价(15) 尺码(16) 数量(17) 云仓(18)
COLS = [(1, '品牌'), (3, '货号'), (4, '描述'), (5, '类目'), (6, '材质'),
        (7, '性别'), (8, '颜色'), (9, '色号'), (10, '季节'), (11, '产地'),
        (12, '零售价'), (13, 'cost'), (14, 'C+-'), (15, '批发价'),
        (16, '尺码'), (17, '数量'), (18, '云仓')]

for vendor, files in VENDORS:
    print(f'\n========== {vendor} ==========')
    vdir = TEST_DIR / vendor
    out_dir = vdir / 'converted'
    out_dir.mkdir(exist_ok=True)
    for fn in files:
        src = vdir / fn
        if not src.exists():
            print(f'  [SKIP] {fn} (不存在)')
            continue
        stem = src.stem
        out = out_dir / f'{stem}_converted.xlsx'
        # Run conversion
        result = subprocess.run(
            ['python3', str(SCRIPTS / 'convert_vendor_quote.py'),
             '--primary-file', str(src),
             '--vendor', vendor,
             '--output', str(out)],
            capture_output=True, text=True, timeout=300,
        )
        # Read output and count fills
        if not out.exists():
            print(f'  [FAIL] {fn}: 转换失败')
            print(f'    {result.stderr[-200:]}')
            continue
        wb = load_workbook(out)
        ws = wb.active
        total = ws.max_row - 1
        if total <= 0:
            print(f'  [SKIP] {fn}: 0 行')
            continue
        embedded_imgs = len(ws._images)
        print(f'  {fn}: {total} 行, 嵌入图 {embedded_imgs}')
        for c, name in COLS:
            filled = sum(1 for r in range(2, ws.max_row + 1)
                         if ws.cell(r, c).value not in (None, "", 0))
            mark = '✓' if filled == total else ('·' if filled > 0 else '✗')
            print(f'    {mark} {name}: {filled}/{total}')
