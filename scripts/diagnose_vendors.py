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
TEST_DIR = Path('/Users/superlist/Desktop/表格处理/配置器测试文件')

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

COLS = [(1, '品牌'), (2, '货号'), (4, '描述'), (5, '类目'), (6, '材质'),
        (7, '颜色'), (8, '季节'), (9, '产地'), (10, 'cost'), (13, '零售价'),
        (14, '尺码'), (15, '库存')]

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
