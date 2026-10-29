"""category-futures: SALOMON FTW 版式.

sheet 'FTW', 结构比较怪:
  R1 = 表头行, 但 c14 起是**尺码换算表** (两行):
        R1: 36 / 36.6666666666 / 37.3333333333 / 38 …   (欧码)
        R2: 3.5 / 4 / 4.5 / 5 …                      (英码)
  R2 = 换算值行
  R3 = 空
  R4+ = 数据行 (c1='x' 标记, c2=Category BU P, c5=Article code, c9=Color,
                 c12=WHLS €, c13=Model Euro Re, c33=Unit total, c34=Value total)

数据行里 c14-c32 (尺码区) **全是空的** —— 这批期货单没填按尺码的量,
只有 c33 'Unit total' 整单数量。所以每个 SKU 输出一行, 尺码留空, 数量取 Unit total。

图片: anchor 在 c10 'Article image' 列。
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
    norm_header,
    open_safe_workbook,
)


ALIASES: Dict[str, tuple] = {
    "sku":       ("Article code",),
    "name":      ("Article name",),
    "category":  ("Category BU P", "Category BU"),
    "color":     ("Color",),
    "cost":      ("WHLS €", "WHLS"),
    "retail":    ("Model Euro Re", "Model Euro Retail"),
    "units":     ("Unit total",),
    "value":     ("Value total",),
    "image":     ("Article image",),
}

SIGNATURE = ("Article code", "Color", "WHLS")


def looks_like_salomon(ws) -> bool:
    hdr = find_header_row(ws, SIGNATURE, max_scan=6)
    if hdr is None:
        return False
    return "sku" in header_col_map(ws, hdr, ALIASES)


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

        # 数据行起点: 表头之后第一行非空的行 (R2 是换算行, R3 空)
        data_start = hdr + 1
        for r in range(hdr + 1, min(hdr + 6, ws.max_row) + 1):
            if normalize_text(cellv(ws, r, COL["sku"])):
                data_start = r
                break

        images = extract_images(ws)
        rows: List[Dict[str, object]] = []
        for r in range(data_start, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, COL.get("sku")))
            if not sku:
                continue
            rows.append({
                "brand": "",
                "sku": sku,
                "image": "",
                "description": normalize_text(cellv(ws, r, COL.get("name"))),
                "category": normalize_text(cellv(ws, r, COL.get("category"))),
                "material": "",
                "color": normalize_text(cellv(ws, r, COL.get("color"))),
                "color_code": "",
                "gender": "",
                "season": "",
                "origin": "",
                "cost": to_float(cellv(ws, r, COL.get("cost"))),
                "cplus_text": "",
                "supply_price": None,
                "retail_price": to_float(cellv(ws, r, COL.get("retail"))),
                "size": "",                       # 源文件没填该 SKU 的尺码
                "quantity": to_float(cellv(ws, r, COL.get("units"))),
                "vendor_id": vendor_id,
                "image_items": images.get(r, []),
            })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
