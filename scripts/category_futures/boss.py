"""category-futures: BOSS RTW 版式.

sheet 'RTW', R1 表头, R2+ 数据。**每行一个尺码, 数量就在尺码列本身**。

列 (按表头名定位):
  sp_id | POS | MPG | SPG | Picture | Style No. | Color | Form Name | WHS | RRP
  | Total Units | Total WHS | Core range | Keylook | Fit | Material Comp.
  | Quality | Responsible | Sizes | Total RRP | SKU | Name | <尺码列…>

尺码:
  表头行里 'Sizes' 之后那一段全是尺码标签列 (XS/S/M/L/XL/XXL/…/44/46/…),
  **同一数据行的同列就是该尺码的数量** → 按尺码展开成多行。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from convert_vendor_quote_shared import normalize_text, to_float

from .common import (
    cellv,
    extract_images,
    find_header_row,
    header_col_map,
    is_size_like,
    norm_header,
    open_safe_workbook,
)


ALIASES: Dict[str, tuple] = {
    "sku":         ("SKU",),
    "style_no":    ("Style No.",),
    "color":       ("Color",),
    "form_name":   ("Form Name",),
    "name":        ("Name",),
    "cost":        ("WHS",),
    "retail":      ("RRP",),
    "total_units": ("Total Units",),
    "material":    ("Material Comp.", "Material"),
    "sizes_text":  ("Sizes",),
}

SIGNATURE = ("Style No.", "WHS", "SKU")


def looks_like_boss(ws) -> bool:
    hdr = find_header_row(ws, SIGNATURE, max_scan=6)
    if hdr is None:
        return False
    cols = header_col_map(ws, hdr, ALIASES)
    return "sku" in cols and "cost" in cols


def parse(primary_file: Path, vendor_id: str) -> List[Dict[str, object]]:
    wb = open_safe_workbook(primary_file)
    try:
        ws = None
        for sn in wb.sheetnames:
            cand = wb[sn]
            if find_header_row(cand, SIGNATURE, max_scan=6) is not None:
                ws = cand
                break
        if ws is None:
            return []

        hdr = find_header_row(ws, SIGNATURE, max_scan=6)
        if hdr is None:
            return []
        COL = header_col_map(ws, hdr, ALIASES)
        if "sku" not in COL:
            return []

        # 尺码列: 'Sizes'/'Total RRP'/'SKU'/'Name' 之后, 表头是尺码名的列
        tail = max([COL.get(k, 0) for k in ("sizes_text", "sku", "name")] or [0])
        size_cols: List[int] = []
        for c in range(tail + 1, ws.max_column + 1):
            raw = ws.cell(hdr, c).value
            if raw in (None, ""):
                continue
            if is_size_like(str(raw)):
                size_cols.append(c)

        images = extract_images(ws)
        sku_col = COL["sku"]

        rows: List[Dict[str, object]] = []
        for r in range(hdr + 1, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, sku_col))
            if not sku:
                continue
            color = normalize_text(cellv(ws, r, COL.get("color")))
            desc = (normalize_text(cellv(ws, r, COL.get("name")))
                    or normalize_text(cellv(ws, r, COL.get("form_name"))))
            cost = to_float(cellv(ws, r, COL.get("cost")))
            retail = to_float(cellv(ws, r, COL.get("retail")))
            material = normalize_text(cellv(ws, r, COL.get("material")))
            img_items = images.get(r, [])

            emitted = False
            for c in size_cols:
                qv = ws.cell(r, c).value
                if not isinstance(qv, (int, float)):
                    continue
                size = normalize_text(ws.cell(hdr, c).value)
                rows.append({
                    "brand": "",
                    "sku": sku,
                    "image": "",
                    "description": desc,
                    "category": "",
                    "material": material,
                    "color": color,
                    "color_code": "",
                    "gender": "",
                    "season": "",
                    "origin": "",
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size,
                    "quantity": int(round(qv)),
                    "vendor_id": vendor_id,
                    "image_items": img_items,
                })
                emitted = True

            if not emitted:
                # 源文件这行没填任何尺码数量 → 一行, 尺码留空, 数量取 Total Units
                rows.append({
                    "brand": "",
                    "sku": sku,
                    "image": "",
                    "description": desc,
                    "category": "",
                    "material": material,
                    "color": color,
                    "color_code": "",
                    "gender": "",
                    "season": "",
                    "origin": "",
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": "",
                    "quantity": to_float(cellv(ws, r, COL.get("total_units"))),
                    "vendor_id": vendor_id,
                    "image_items": img_items,
                })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
