"""category-futures: NuORDER 版式 (ACNE / BC / Men ss27 / POLO RL).

sheet 'NuORDER Order Data', 表头在 R1, 数据 R2+。**每行一个尺码**。

四家列序完全不同 (所以一律按表头名定位列):
  ACNE     19 列: ... c17=Size  c18=Qty  c19=Units per package
  BC       21 列: ... c19=Size  c20=Qty  c21=Units per package
  POLO RL  29 列: ... c24=Size c25=Inventory Qty c26=EAN c27=UPC c28=Qty c29=Units per package
  Men ss27 55 列: c18=XXS c19=XXS size price c20=XS … c55=L/XL size price
                  (尺码列交错"尺码单价"列; 尺码列里的数字是**数量**不是尺码)

尺码区识别:
  1. 先按表头名定位所有**固定字段**列, 尺码区 = 最后一个固定字段列 + 1
     (避开 'Name' 这种 4 字母词被误判成尺码)
  2. 尺码区里, 表头是尺码名(XXS/XS/32/44Sizes)或恰好是 'Size' 的列 → 尺码列
  3. 数量列 = 表头含 'Qty' 且**不含 'size price'** 的列;
     每个尺码列取它右边最近的那个数量列。
     没有数量列 (Men ss27 只有尺码单价) → 数量留空, 不凭空造。

两种版式(2026-10-30 期货口径):
  A. 行级尺码(尺码列只有 1 个): ACNE / BC / POLO RL
     → 一行一个尺码, 尺码值就在那列里, 数量取右边的 Qty 列。
  B. 网格型(尺码列 > 1): Men ss27
     → 列 = 尺码, 行 = SKU, 单元格里的数字是**订货数量**不是尺码;
       期货单忽略数量, 按尺码组**全部展开**, 数量留空。
       尺码组分法见 _build_grid_groups (按标签类型, 不能只看列间距)。
"""

from __future__ import annotations

import re
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


def _label_kind(label: str) -> str:
    """尺码标签类型: 'num'(纯数字, 如 26/38) / 'alpha'(含字母, 如 XXS / 3XL / OS)。"""
    s = str(label or "").strip()
    return "num" if re.fullmatch(r"\d+(\.\d+)?", s) else "alpha"


def _build_grid_groups(ws, header_row: int, size_cols: List[int]
                       ) -> List[List[int]]:
    """尺码网格分组 —— 只在**网格型**版式(尺码列 > 1)用。

    NuORDER 的尺码列和 'XXS size price' 这种**尺码单价列交错**:
        c18=XXS | c19=XXS size price | c20=XS | c21=XS size price | …
    而且**服装尺码组和鞋码组都是这么排的**, 光看列间距(都是隔 1 列)分不开。

    所以分组规则是: 从「有数据的尺码段」向两侧扩展, 直到遇到**标签类型不同**的尺码列
    (纯数字 ↔ 含字母) 或非尺码列。
    例(Men ss27): live = c22(S) c24(M) c26(L) c28(XL) c30(XXL)
                   → 向左扩到 c18(XXS), 向右扩到 c34(4XL), 遇到 c36('26' 纯数字) 停
                   → 组 = c18..c34 共 9 个(XXS/XS/S/M/L/XL/XXL/3XL/4XL)
    鞋码组 c36-c48 和 OS/S-M/L-XL 组整列无数据 → 不展开(不该给服装 SKU 加鞋码)。

    Returns:
        所有分组(每个分组是一串列号)
    """
    if len(size_cols) <= 1:
        return []
    labels = {c: str(ws.cell(header_row, c).value or "").strip() for c in size_cols}

    live_idx = [
        i for i, c in enumerate(size_cols)
        if any(ws.cell(r, c).value not in (None, "")
               for r in range(header_row + 1, ws.max_row + 1))
    ]
    if not live_idx:
        return []

    # live 列按在 size_cols 里的位置切成连续段
    runs: List[List[int]] = []
    for i in live_idx:
        if runs and i - runs[-1][-1] <= 1:
            runs[-1].append(i)
        else:
            runs.append([i])

    groups: List[List[int]] = []
    for run in runs:
        lo, hi = run[0], run[-1]
        k0 = _label_kind(labels[size_cols[lo]])
        i = lo
        while i - 1 >= 0 and _label_kind(labels[size_cols[i - 1]]) == k0:
            i -= 1
        j = hi
        while j + 1 < len(size_cols) and _label_kind(labels[size_cols[j + 1]]) == k0:
            j += 1
        groups.append(size_cols[i:j + 1])
    return groups


def _row_grid_groups(ws, groups: List[List[int]], row: int) -> List[List[int]]:
    """这一行该展开哪些尺码组: 本行有数据的组; 没有就用全部有数据的组。"""
    hit = [g for g in groups
           if any(ws.cell(row, c).value not in (None, "") for c in g)]
    return hit or groups


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
        # 网格型(尺码列 > 1)才做"全部尺码展开"(2026-10-30 期货口径);
        # 单尺码列的文件(ACNE/BC/POLO RL 那种一行一尺码)保持原逻辑。
        grid_groups = _build_grid_groups(ws, hdr, [c for c, _ in size_pairs])
        is_grid = bool(grid_groups)

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

            if is_grid:
                # ---- 网格型(尺码列 > 1): 期货口径 ----
                # ⚠️ 源里这些列是「每个尺码的订货格」(Men ss27: c18-c34 服装组 XXS…4XL),
                #    单元格里的数字是**数量**不是尺码 —— 绝不能当尺码读。
                #    (老路径会把 c22 的 12 / c24 的 30 当成尺码, 出来一堆假尺码)
                # 期货单口径: **忽略数量, 按该 SKU 所属的尺码组全部展开, 数量留空**。
                # 大多数行整行全空 → 只按"这行哪列有值"出尺码会让 180/205 行没尺码。
                for g in _row_grid_groups(ws, grid_groups, r):
                    for size_col in g:
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
                            "size": normalize_text(ws.cell(hdr, size_col).value),
                            "quantity": None,     # 期货单: 数量留空等客户填
                            "vendor_id": vendor_id,
                            "image_items": img_items,
                        })
                    emitted = True
            else:
                # ---- 行级尺码(一行一个尺码): ACNE / BC / POLO RL ----
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
