"""V1046 LOSCHI vendor adapter."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from convert_vendor_quote_shared import (
    find_col,
    guess_product_group,
    is_positive_qty,
    load_images_by_row_per_sheet,  # 替代 load_sheet_embedded_images(多 sheet + 简单行匹配)
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    sheet_images = load_images_by_row_per_sheet(primary_file)
    sheet_row_heights = load_sheet_row_heights(primary_file)
    workbook = pd.ExcelFile(primary_file)

    for sheet in workbook.sheet_names:
        df = pd.read_excel(primary_file, sheet_name=sheet)
        if df.empty:
            continue
        brand_col = find_col(df, ["MARCHIO", "BRAND"])
        # 货号用 RIGA 列(LOSCHI 版式:ARTICOLO 是款号,MODELLO 是颜色名,RIGA 才是真货号)
        item_col = find_col(df, ["RIGA", "ARTICOLO", "ITEM", "MODELLO"])
        model_col = find_col(df, ["MODELLO", "COLOR", "COLORE"])
        # 类目用 VOCE 列(原 DESCRIPTION 字段实际是分类,例如 "JEANS -U")
        voce_col = find_col(df, ["VOCE"])
        desc_col = find_col(df, ["DESCRIPTION"])  # 没有就用空
        cost_col = find_col(df, ["PRZ.ACQUISTO", "COST"])
        retail_col = find_col(df, ["PRZ.VENDITA", "RETAIL"])
        total_col = find_col(df, ["TOT.", "TOTAL", "QTY"])
        if brand_col is None or item_col is None:
            continue

        cols = list(df.columns)
        start = cols.index(item_col) + 3 if model_col else 6
        end = cols.index(total_col) - 1 if total_col in cols else len(cols) - 3
        if start >= len(cols):
            continue
        size_cols = cols[max(start, 6): max(start, 6) + max(end - max(start, 6) + 1, 0)]
        size_labels: Dict[Tuple[str, str, str], Dict[str, str]] = {}

        for idx, row in df.iterrows():
            brand = normalize_text(row.get(brand_col))
            item = normalize_text(row.get(item_col))
            model = normalize_text(row.get(model_col)) if model_col else ""
            if not brand or not item:
                continue
            key = (brand, item, model)
            label_map = size_labels.get(key, {})

            has_label_row = False
            candidate_map: Dict[str, str] = {}
            for c in size_cols:
                val = normalize_text(row.get(c))
                if not val:
                    continue
                if to_float(val) is not None or re.fullmatch(r"[A-Z]{1,5}", val.upper()):
                    has_label_row = True
                    candidate_map[c] = val
            if has_label_row and not is_positive_qty(row.get(total_col) if total_col else None):
                size_labels[key] = candidate_map
                continue

            effective_labels = label_map if label_map else {c: normalize_text(c) for c in size_cols}
            emitted = 0
            for c in size_cols:
                qty_raw = row.get(c)
                if not is_positive_qty(qty_raw):
                    continue
                size = normalize_text(effective_labels.get(c))
                if not size:
                    continue
                emitted += 1
                excel_row = int(idx) + 2
                # 类目 = VOCE 列(产品类型);描述 = 真正的 DESCRIPTION 列(没有就空)
                category = normalize_text(row.get(voce_col)) if voce_col else ""
                desc = normalize_text(row.get(desc_col)) if desc_col else ""
                cost = to_float(row.get(cost_col)) if cost_col else None
                retail = to_float(row.get(retail_col)) if retail_col else None
                rows.append(
                    {
                        "brand": brand,
                        "sku": item,
                        "image": "",
                        "description": desc,
                        "category": category,
                        "material": "",
                        "color": model,
                        "season": "",
                        "origin": "",
                        "cost": cost,
                        "cplus_text": "",
                        "supply_price": cost,
                        "retail_price": retail,
                        "size": size,
                        "quantity": int(round(float(to_float(qty_raw) or 0))),
                        "vendor_id": vendor_id,
                        "image_items": sheet_images.get(sheet, {}).get(int(idx), []),
                        "source_row_height": sheet_row_heights.get(sheet, {}).get(excel_row),
                    }
                )
            if emitted == 0:
                continue
    return rows
