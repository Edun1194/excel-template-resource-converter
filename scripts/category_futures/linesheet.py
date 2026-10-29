"""category-futures: Linesheet 版式 (LEMAIRE / Alexander Wang / GANNI).

结构 (表头行位置随文件而变, 所以动态找):
  表头行: 含 'Style Number' + 'Available Sizes' 的那一行
    LEMAIRE          R4
    Alexander Wang   R1
    GANNI            R7
  列 (按表头名定位, 不写死列号):
    Style Image | Style Name | Style Number | Color | Color Code | Color ID
    | Fabrication | Fab. Code | Materials | Description | Category | Subcategory
    | Available Sizes | BULK | <尺码标签列…> | Wholesale (EUR) | Sugg. Retail
    | Units | Style Minimum | Color Minimum | …

尺码处理:
  源文件的尺码在 **'Available Sizes' 文本列** (如 "XXS, XS, S, M, L, XL"),
  按逗号拆成多行; 源文件**没有按尺码的数量** (只有整单的 Units), 所以数量留空
  (与 BURBERRY B&H 的处理一致, 不凭空造数)。

季节:
  表头行之上有 'Season Year' 标签行, 取其右侧单元格的值; 没有则留空。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List

from convert_vendor_quote_shared import normalize_text, to_float

from .common import (
    cellv,
    extract_images,
    find_header_row,
    header_col_map,
    norm_header,
    open_safe_workbook,
)


ALIASES: Dict[str, tuple] = {
    "sku":          ("Style Number",),
    "name":         ("Style Name",),
    "color":        ("Color",),
    "color_code":   ("Color Code",),
    "color_id":     ("Color ID",),
    "fabrication":  ("Fabrication",),
    "fab_code":     ("Fab. Code", "Fab Code"),
    "material":     ("Materials", "Material"),
    "description":  ("Description",),
    "category":     ("Category",),
    "subcategory":  ("Subcategory",),
    "sizes":        ("Available Sizes",),
    "cost":         ("Wholesale (EUR)", "Wholesale (USD)", "Wholesale Price (EUR)",
                     "Wholesale Price", "Wholesale"),
    "retail":       ("Sugg. Retail (EUR)", "Sugg. Retail", "Suggested Retail"),
    "units":        ("Units",),
}

# 表头特征 (用来认这个版式)
SIGNATURE = ("Style Number", "Available Sizes", "Style Image")


def looks_like_linesheet(ws) -> bool:
    """是不是 Linesheet 版式."""
    hdr = find_header_row(ws, SIGNATURE, max_scan=15)
    if hdr is None:
        return False
    cols = header_col_map(ws, hdr, ALIASES)
    return "sku" in cols and "sizes" in cols


def _find_season(ws, header_row: int) -> str:
    """表头行之上的 'Season Year' 标签行 → 右侧单元格."""
    for r in range(1, header_row):
        if norm_header(ws.cell(r, 1).value) == "seasonyear":
            return normalize_text(ws.cell(r, 2).value)
    return ""


def _split_sizes(raw) -> List[str]:
    """'XXS, XS, S' → ['XXS','XS','S']; 空 → []."""
    if raw in (None, ""):
        return []
    parts = [p.strip() for p in re.split(r"[,;/]", str(raw)) if p.strip()]
    return parts


def parse(primary_file: Path, vendor_id: str) -> List[Dict[str, object]]:
    wb = open_safe_workbook(primary_file)
    try:
        ws = None
        for sn in wb.sheetnames:
            cand = wb[sn]
            if cand.max_row and cand.max_column and looks_like_linesheet(cand):
                ws = cand
                break
        if ws is None:
            return []

        hdr = find_header_row(ws, SIGNATURE, max_scan=15)
        if hdr is None:
            return []
        COL = header_col_map(ws, hdr, ALIASES)
        if "sku" not in COL:
            return []

        season = _find_season(ws, hdr)
        images = extract_images(ws)
        sku_col = COL["sku"]

        rows: List[Dict[str, object]] = []
        for r in range(hdr + 1, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, sku_col))
            if not sku:
                continue

            color = normalize_text(cellv(ws, r, COL.get("color")))
            color_code = normalize_text(cellv(ws, r, COL.get("color_code")))
            description = normalize_text(cellv(ws, r, COL.get("description")))
            material = normalize_text(cellv(ws, r, COL.get("material")))
            category = normalize_text(cellv(ws, r, COL.get("category")))
            cost = to_float(cellv(ws, r, COL.get("cost")))
            retail = to_float(cellv(ws, r, COL.get("retail")))
            img_items = images.get(r, [])

            sizes = _split_sizes(cellv(ws, r, COL.get("sizes")))
            if not sizes:
                sizes = [""]        # 源文件没给尺码 → 一行, 尺码留空

            for size in sizes:
                rows.append({
                    "brand": "",          # 源文件无品牌列 → 交给框架用文件名兜底
                    "sku": sku,
                    "image": "",
                    "description": description,
                    "category": category,
                    "material": material,
                    "color": color,
                    "color_code": color_code,
                    "gender": "",
                    "season": season,
                    "origin": "",
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size,
                    "quantity": None,     # 源文件无按尺码数量 → 留空
                    "vendor_id": vendor_id,
                    "image_items": img_items,
                })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
