"""category-futures: NuORDER 版式 (ACNE / BC / Men ss27 / POLO RL).

sheet 'NuORDER Order Data', 表头在 R1, 数据 R2+。**每行一个尺码**。

四家列序完全不同 (所以一律按表头名定位列):
  ACNE     19 列: ... c17=Size  c18=Qty  c19=Units per package
  BC       21 列: ... c19=Size  c20=Qty  c21=Units per package
  POLO RL  29 列: ... c24=Size c25=Inventory Qty c26=EAN c27=UPC c28=Qty c29=Units per package
  Men ss27 55 列: ... c18=XXS c19=XXS size price c20=XS c21=XS size price ... (只有尺码单价, 没有数量)

尺码区识别:
  1. 先按表头名定位所有**固定字段**列, 尺码区 = 最后一个固定字段列 + 1
     (避开 'Name' 这种 4 字母词被误判成尺码)
  2. 尺码区里, 表头是尺码名(XXS/XS/32/44Sizes)或恰好是 'Size' 的列 → 尺码列
  3. 数量列 = 表头含 'Qty' 且**不含 'size price'** 的列;
     每个尺码列取它右边最近的那个数量列。
     没有数量列 (Men ss27 只有尺码单价) → 数量留空, 不凭空造。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple

from convert_vendor_quote_shared import normalize_text, to_float

from .common import (
    cellv,
    find_header_row,
    header_col_map,
    is_size_like,
    load_images_by_row_zip,
    norm_header,
    open_safe_workbook,
)


ALIASES: Dict[str, tuple] = {
    "season":       ("Season",),
    "color":        ("Color",),
    "color_code":   ("Color Code",),
    "sku":          ("Style Number",),
    "name":         ("Name",),
    "cost":         ("Wholesale Price (EUR)", "Wholesale Price", "Wholesale (EUR)",
                     "Wholesale"),
    "retail":       ("RRP (EUR)", "M.S.R.P. (EUR)", "M.S.R.P (EUR)", "M.S.R.P."),
    "division":     ("Division",),
    "department":   ("Department",),
    "category":     ("Category",),
    "subcategory":  ("Subcategory",),
    "notes":        ("Product Notes",),
    "material":     ("Fabric Description", "Materials", "Material"),
    "total_units":  ("Total Units",),
    "image":        ("Image",),
}

SIGNATURE = ("Season", "Style Number", "Wholesale")


def looks_like_nuorder(ws) -> bool:
    """是不是 NuORDER 版式."""
    if "nuorder" in str(ws.title).lower():
        hdr = find_header_row(ws, SIGNATURE, max_scan=6)
        return hdr is not None
    return False


def _gender_from_division(division: str) -> str:
    """Division 推断性别 (同 V1070-期货 的规则)."""
    u = (division or "").upper()
    if "LADIES" in u or "WOMAN" in u or "WOMEN" in u:
        return "Woman"
    if "MEN" in u or "UOMO" in u or " MAN" in u or u.startswith("MAN"):
        return "Man"
    return ""


def _find_size_cols(ws, header_row: int, COL: Dict[str, int]
                    ) -> Tuple[List[Tuple[int, Optional[int]]], List[int]]:
    """返回 (尺码列+数量列 列表, 数量列候选列表)."""
    # 固定字段区最大列 → 尺码区起点
    fixed_max = max(COL.values()) if COL else 0
    start = fixed_max + 1

    qty_cols: List[int] = []
    size_cols: List[int] = []
    for c in range(start, ws.max_column + 1):
        h = norm_header(ws.cell(header_row, c).value)
        if not h:
            continue
        if "sizeprice" in h:          # 尺码单价 → 不是数量
            continue
        if "qty" in h or h == "qty":  # 数量列 (Qty / Inventory Qty)
            qty_cols.append(c)
            continue
        if h in ("ean", "upc", "unitsperpackage", "perpackage"):
            continue
        raw = ws.cell(header_row, c).value
        if h == "size" or is_size_like(str(raw)):
            size_cols.append(c)

    pairs: List[Tuple[int, Optional[int]]] = []
    for c in size_cols:
        near = [q for q in qty_cols if q > c]
        pairs.append((c, min(near) if near else None))
    return pairs, qty_cols


def parse(primary_file: Path, vendor_id: str) -> List[Dict[str, object]]:
    wb = open_safe_workbook(primary_file)
    try:
        target = None
        for sn in wb.sheetnames:
            if "nuorder" in sn.lower():
                target = wb[sn]
                break
        if target is None:
            for sn in wb.sheetnames:
                cand = wb[sn]
                if find_header_row(cand, SIGNATURE, max_scan=6) is not None:
                    target = cand
                    break
        if target is None:
            return []
        ws = target

        hdr = find_header_row(ws, SIGNATURE, max_scan=6)
        if hdr is None:
            return []
        COL = header_col_map(ws, hdr, ALIASES)
        if "sku" not in COL:
            return []

        size_pairs, _qty_cols = _find_size_cols(ws, hdr, COL)

        # 图片: openpyxl 解析不了就退回 zip 解析
        images_0 = load_images_by_row_zip(primary_file)
        images_1 = {r + 1: v for r, v in images_0.items()}   # → 1-idx 行号

        rows: List[Dict[str, object]] = []
        for r in range(hdr + 1, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, COL.get("sku")))
            if not sku:
                continue

            color = normalize_text(cellv(ws, r, COL.get("color")))
            color_code = normalize_text(cellv(ws, r, COL.get("color_code")))
            name = normalize_text(cellv(ws, r, COL.get("name")))
            cost = to_float(cellv(ws, r, COL.get("cost")))
            retail = to_float(cellv(ws, r, COL.get("retail")))
            category_v = normalize_text(cellv(ws, r, COL.get("category")))
            subcategory = normalize_text(cellv(ws, r, COL.get("subcategory")))
            category = f"{category_v} / {subcategory}" if subcategory else category_v
            season = normalize_text(cellv(ws, r, COL.get("season")))
            material = normalize_text(cellv(ws, r, COL.get("material")))
            division = normalize_text(cellv(ws, r, COL.get("division")))
            gender = _gender_from_division(division)
            img_items = images_1.get(r, [])

            # SKU = <style>-<color_code> (同 V1070-期货 规范)
            sku_out = f"{sku}-{color_code}" if color_code else sku

            emitted = False
            for size_col, qty_col in size_pairs:
                size_raw = ws.cell(r, size_col).value
                if size_raw in (None, ""):
                    continue          # 这一列这行没填尺码 → 跳过
                size = normalize_text(size_raw)
                qty = None
                if qty_col:
                    qv = ws.cell(r, qty_col).value
                    if isinstance(qv, (int, float)):
                        qty = int(round(qv))
                rows.append({
                    "brand": "",
                    "sku": sku_out,
                    "image": "",
                    "description": name,
                    "category": category,
                    "material": material,
                    "color": color,
                    "color_code": color_code,
                    "gender": gender,
                    "season": season,
                    "origin": "",
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "image_items": img_items,
                })
                emitted = True

            if not emitted:
                # 这行没填任何尺码 (如 Men ss27 只有尺码单价没有尺码) → 一行, 尺码留空
                rows.append({
                    "brand": "",
                    "sku": sku_out,
                    "image": "",
                    "description": name,
                    "category": category,
                    "material": material,
                    "color": color,
                    "color_code": color_code,
                    "gender": gender,
                    "season": season,
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
