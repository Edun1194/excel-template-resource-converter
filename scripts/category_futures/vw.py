"""category-futures: VIVIENNE WESTWOOD 'Foglio1' 版式 (VW_ATS.xlsx).

sheet 'Foglio1', R1 表头, R2+ 数据。**每行一个 SKU+尺码**。

列 (按表头名定位):
  Prod Selling Season | Collection | Product Department | Prod Class | Prod Brand
  | Generic Product Code | Material Code | Product Image | Style Description
  | Prod Fabric Type | Prod Fabric Description | Colour Description | Colour Code
  | Prod Size | Available Qty | WHL €

注意:
  - 'Product Image' 列的值是 '#VALUE!' (公式错误), **没有内嵌图片** → 图片留空
  - 'Prod Size' 的值可能是 '---' (无尺码), 原样保留
  - 'Prod Brand' 的值是 'IA IT ACCESSORI' 这种**内部品牌码**, 不是品牌名 → 不当品牌用
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
    open_safe_workbook,
)


ALIASES: Dict[str, tuple] = {
    "season":     ("Prod Selling Season", "Prod Selling Se"),
    "collection": ("Collection",),
    "department": ("Product Department", "Product Departm"),
    "category":   ("Prod Class ", "Prod Class"),
    "brand_code": ("Prod Brand ", "Prod Brand"),
    "sku":        ("Generic Prod Code", "Generic Product Code"),
    "material_code": ("Material Code ", "Material Code"),
    "description": ("Style Description", "Style Des"),
    "fabric_type": ("Prod Fabric Type", "Prod Fabric Typ"),
    "material":   ("Prod Fabric Description", "Prod Fabric Des"),
    "color":      ("Colour Description", "Colour Des"),
    "color_code": ("Colour Code ", "Colour Code"),
    "size":       ("Prod Size ", "Prod Size"),
    "qty":        ("Available Qty ", "Available Qty"),
    "cost":       ("WHL €", "WHL"),
}

SIGNATURE = ("Generic Prod Code", "Available Qty", "WHL")

GENDER_MAP = {
    "UNI-UNISEX": "Unisex", "UNISEX": "Unisex",
    "WOMAN": "Woman", "WOMEN": "Woman", "DONNA": "Woman",
    "MAN": "Man", "MEN": "Man", "UOMO": "Man",
}

# 源 'Prod Size' 列里的"没有尺码"占位符(一片装商品写 '---')→ 留空
_SIZE_PLACEHOLDERS = {"---", "--", "-", "–", "—", "/"}


def _clean_size(value: str) -> str:
    """尺码占位符('---' 一片装) → 留空; 其余原样。"""
    s = (value or "").strip()
    return "" if s in _SIZE_PLACEHOLDERS else s


def looks_like_vw(ws) -> bool:
    hdr = find_header_row(ws, SIGNATURE, max_scan=6)
    if hdr is None:
        return False
    cols = header_col_map(ws, hdr, ALIASES)
    return "sku" in cols and "qty" in cols


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

        images = extract_images(ws)
        rows: List[Dict[str, object]] = []
        for r in range(hdr + 1, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, COL.get("sku")))
            if not sku:
                continue
            dept = normalize_text(cellv(ws, r, COL.get("department")))
            gender = GENDER_MAP.get(dept.upper(), "")
            rows.append({
                "brand": "",                    # 源文件的 'Prod Brand' 是内部码, 不是品牌名
                "sku": sku,
                "image": "",
                "description": normalize_text(cellv(ws, r, COL.get("description"))),
                "category": normalize_text(cellv(ws, r, COL.get("category"))),
                "material": normalize_text(cellv(ws, r, COL.get("material"))),
                "color": normalize_text(cellv(ws, r, COL.get("color"))),
                "color_code": normalize_text(cellv(ws, r, COL.get("color_code"))),
                "gender": gender,
                "season": normalize_text(cellv(ws, r, COL.get("season"))),
                "origin": "",
                "cost": to_float(cellv(ws, r, COL.get("cost"))),
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,          # 源文件没有零售价
                "size": _clean_size(normalize_text(cellv(ws, r, COL.get("size")))),
                "quantity": None,                  # 期货单: 数量留空等客户填
                "vendor_id": vendor_id,
                "image_items": images.get(r, []),
            })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
