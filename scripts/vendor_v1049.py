"""V1049 POMPEU vendor adapter — SPECIAL OFFER C-40 / C-45 (2026-09-03 重写)。

**版式 (2026-09-03 新)**:
- sheet 名: 'Sheet1' (1 sheet, 块状)
- R1: blank
- R2 ~ R(N-1): size code 矩阵 — 1-idx col 9 = size code (D/O/Y/X/U/M/OS/PL/W1/T/UK/38 等),
  1-idx col 10+ = size labels (e.g. 'D' 行 = 28/30/.../64, 'O' 行 = XXXS/XXS/.../5XL)
- R(N): header (1-idx col 1 = 'Item')
- R(N+1)+: data
  - c1 = Item (SKU)
  - c2 = Group, c3 = Sector, c4 = Super, c5 = Season
  - c6 = Description, c7 = Brand, c8 = Colour
  - c9 = size code (匹配 size code 矩阵)
  - c10+ = qty per size (按 size code 矩阵的 size labels 位置)
  - c35 = total Qty, c36 = Whole Sale (cost), c37 = Retail

**两个 C-40/C-45 文件差异**:
- C-40 (133 行): size code 矩阵 9 行 (R2-R10), header R11, data R12+
- C-45 (182 行): size code 矩阵 11 行 (R2-R12), header R13, data R14+

C+ 永远来自用户输入(前端),本 adapter 不读 markup 文档,也不计算 supply_price。
这些字段在 dispatcher 里统一重置 + 覆盖(参见 convert_vendor_quote.py)。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    load_sheet_images_by_row,
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def _build_size_maps_from_xlsx(ws) -> Dict[str, Dict[int, str]]:
    """从 ws 扫描 size code 矩阵,返回 {size_code: {col_idx: size_label}}。

    算法: 找 header row (c1 = 'Item'), 从 header_row-1 向上扫到 R1,
    每行 c9 是 size code, c10+ 是 size labels。
    """
    # 1) 找 header row
    header_row = None
    for r in range(1, min(20, ws.max_row + 1)):
        v = ws.cell(r, 1).value
        if v and isinstance(v, str) and v.strip().lower() == "item":
            header_row = r
            break
    if header_row is None:
        return {}

    # 2) 向上扫 size code 矩阵 (从 header_row - 1 往下到 R1)
    size_maps: Dict[str, Dict[int, str]] = {}
    for r in range(header_row - 1, 0, -1):
        code_raw = ws.cell(r, 9).value  # 1-idx col 9
        if code_raw is None:
            continue
        code = str(code_raw).strip().upper()
        if not code:
            continue
        col_map: Dict[int, str] = {}
        for c in range(10, ws.max_column + 1):
            label = ws.cell(r, c).value
            if label is not None:
                label_str = str(label).strip()
                if label_str:
                    col_map[c] = label_str
        if col_map:
            size_maps[code] = col_map
    return size_maps


def _format_sku_with_color(sku: str, color: str) -> str:
    """V1049 货号拼接:sku + '-' + color 归一化。

    V1049 SKU 本身不含颜色代码 (e.g. '041001140215', '777491T01571081'),
    颜色独立存在 c8 (e.g. 'Fucsia', 'NERO/ARGENTO', 'Black/white')。
    按 Step 3.5 规则 + user 反馈 (2026-09-03):
    - color 有值 → 拼 `<SKU>-<COLOR>`, 多空白/换行压成单空格, 保留 '/' 等
    - color 为空 → 不拼, 用 SKU 原样
    """
    color_norm = re.sub(r"\s+", " ", color).strip() if color else ""
    if color_norm:
        return f"{sku}-{color_norm}"
    return sku


def get_vendor_meta() -> Dict[str, object]:
    """V1049 vendor 元数据 — 给 dispatcher 用。

    default_image_col_width=70:每个产品 3-4 张图,默认 35 字符列宽放不下,加宽。
    """
    return {
        "default_image_col_width": 70.0,
    }


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    """V1049 POMPEU 解析入口 (2026-09-03 重写适配 SPECIAL OFFER C-40/C-45)。

    注:第二个参数原来是 markup_docx(V1049 的 C+ 加价规则文档)。
    现在 C+ 永远来自用户输入(前端表单),markup 文档不再使用,签名保留但忽略。
    """
    rows: List[Dict[str, object]] = []
    # 用 load_sheet_images_by_row 拿全所有图(每个产品 3-4 张,不要走 SALOMON-style 去重)
    sheet_images = load_sheet_images_by_row(primary_file)
    sheet_row_heights = load_sheet_row_heights(primary_file)

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    if not wb.worksheets:
        return rows
    ws = wb.worksheets[0]

    # 1) 找 header row + 扫 size code 矩阵
    header_row = None
    for r in range(1, min(20, ws.max_row + 1)):
        v = ws.cell(r, 1).value
        if v and isinstance(v, str) and v.strip().lower() == "item":
            header_row = r
            break
    if header_row is None:
        return rows

    size_maps = _build_size_maps_from_xlsx(ws)
    if not size_maps:
        return rows

    # 2) 遍历 data rows (header_row + 1 ~ max_row)
    for r in range(header_row + 1, ws.max_row + 1):
        sku_raw = ws.cell(r, 1).value
        if sku_raw is None or not str(sku_raw).strip():
            continue
        sku = str(sku_raw).strip()

        size_code_raw = ws.cell(r, 9).value
        if size_code_raw is None:
            continue
        size_code = str(size_code_raw).strip().upper()
        if not size_code:
            continue

        col_map = size_maps.get(size_code, {})
        if not col_map:
            # 未识别的 size code,跳过
            continue

        season = normalize_text(ws.cell(r, 5).value)
        desc = normalize_text(ws.cell(r, 6).value)
        brand = normalize_text(ws.cell(r, 7).value)
        color = normalize_text(ws.cell(r, 8).value)
        # cost = c36 (Whole Sale), retail = c37 (Retail)
        cost = to_float(ws.cell(r, 36).value) if ws.max_column >= 36 else None
        retail = to_float(ws.cell(r, 37).value) if ws.max_column >= 37 else None

        # 货号拼接:sku + '-' + color 归一化 (2026-09-03 user 反馈)
        sku_final = _format_sku_with_color(sku, color)

        # 图:openpyxl 0-idx = 1-idx row - 1
        row_images = sheet_images.get(r - 1, [])
        row_height = sheet_row_heights.get(ws.title, {}).get(r)

        used_size = False
        for c_idx, size_label in col_map.items():
            if c_idx > ws.max_column:
                continue
            qty_raw = ws.cell(r, c_idx).value
            if not is_positive_qty(qty_raw):
                continue
            used_size = True
            qty = int(round(float(to_float(qty_raw) or 0)))
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": desc,
                "category": "",  # V1049 没独立 category 字段;留空
                "material": "",
                "color": color,
                "season": season,
                "origin": "",
                "cost": cost,
                "cplus_text": "",  # 由 dispatcher 覆盖
                "supply_price": None,  # 由 dispatcher 覆盖
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": row_height,
            })

        if not used_size:
            continue
    return rows
