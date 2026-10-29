"""category-futures: LONGCHAMP 'QTE BDC' 版式.

sheet 'QTE BDC', R1 表头, R2+ 数据。**每行一个 SKU+尺码**。

列 (按表头名定位):
  Code Style | BDC | Thème | … | Genre Produit | Catégorie Matière | Catégorie Produit
  | Code fonction SAP | Fonction SAP | CL/NT | Code Modèle | Code Matière
  | Désignation BDC | Code Couleur | Nom couleur | Tarif Gros | Tarif Détail
  | Qté | Val Gros | Val PDC* | … | Door Code | Custom PO Label | UPC | Qty
  | SKU | Article | <尺码列>

数量: 'Qté' (订购量) 优先, 空则取 'Qty'。
尺码: 表头末尾那列 (本例 c34='844', 与 Code Couleur 同值)。
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
    "sku":         ("Code Style", "SKU"),
    "description": ("Désignation BDC", "Designation BDC"),
    "color":       ("Nom couleur",),
    "color_code":  ("Code Couleur",),
    "category":    ("Catégorie Produit", "Categorie Produit"),
    "material":    ("Catégorie Matière", "Categorie Matiere"),
    "gender":      ("Genre Produit",),
    "cost":        ("Tarif Gros",),
    "retail":      ("Tarif Détail", "Tarif Detail"),
    "qty_order":   ("Qté", "Qte"),
    "qty":         ("Qty",),
    # 以下都是固定字段, 映射进来是为了把尺码列定位到正确位置
    "article":     ("Article",),
    "upc":         ("UPC",),
    "door_code":   ("Door Code",),
    "custom_po":   ("Custom PO Label",),
    "val_gross":   ("Val Gros",),
    "val_pdc":     ("Val PDC*",),
}

SIGNATURE = ("Code Style", "Tarif Gros", "Qté")

GENDER_MAP = {
    "FEMININ": "Woman", "FEMME": "Woman", "WOMAN": "Woman", "F": "Woman",
    "MASCULIN": "Man", "HOMME": "Man", "MAN": "Man", "M": "Man",
    "UNISEXE": "Unisex", "UNISEX": "Unisex",
}


def looks_like_longchamp(ws) -> bool:
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

        # 尺码列: 固定字段区之后, 第一个非空表头的列 (本例 c34='844')
        fixed_max = max(COL.values())
        size_col = None
        for c in range(fixed_max + 1, ws.max_column + 1):
            raw = ws.cell(hdr, c).value
            if raw in (None, ""):
                continue
            size_col = c
            break

        images = extract_images(ws)
        rows: List[Dict[str, object]] = []
        for r in range(hdr + 1, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, COL.get("sku")))
            if not sku:
                continue
            qty = cellv(ws, r, COL.get("qty_order"))
            if qty in (None, ""):
                qty = cellv(ws, r, COL.get("qty"))
            gender_raw = normalize_text(cellv(ws, r, COL.get("gender")))
            rows.append({
                "brand": "",
                "sku": sku,
                "image": "",
                "description": normalize_text(cellv(ws, r, COL.get("description"))),
                "category": normalize_text(cellv(ws, r, COL.get("category"))),
                "material": normalize_text(cellv(ws, r, COL.get("material"))),
                "color": normalize_text(cellv(ws, r, COL.get("color"))),
                "color_code": normalize_text(cellv(ws, r, COL.get("color_code"))),
                "gender": GENDER_MAP.get(gender_raw.upper(), gender_raw),
                "season": "",
                "origin": "",
                "cost": to_float(cellv(ws, r, COL.get("cost"))),
                "cplus_text": "",
                "supply_price": None,
                "retail_price": to_float(cellv(ws, r, COL.get("retail"))),
                "size": normalize_text(ws.cell(r, size_col).value) if size_col else "",
                "quantity": to_float(qty),
                "vendor_id": vendor_id,
                "image_items": images.get(r, []),
            })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
