"""V1043 PRITELLI vendor adapter."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from convert_vendor_quote_shared import (
    find_col,
    guess_product_group,
    is_positive_qty,
    load_images_by_merged_range,  # 图片所在单元格是几行合并的,要用 merge 范围匹配
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    # V1043 的图片在 col A(第 1 列),用合并单元格范围匹配:同一 merge 内的所有
    # 数据行共享同一张图(如 A11:A13 合并 → 行 11/12/13 都是同一张图)。
    sheet_images = load_images_by_merged_range(primary_file, image_col=1)
    sheet_row_heights = load_sheet_row_heights(primary_file)
    df = pd.read_excel(primary_file, sheet_name="Foglio1")

    brand_col = find_col(df, ["BRAND"])
    sku_col = find_col(df, ["ITEM", "MODELLO", "SKU"])
    desc_col = find_col(df, ["DESCRIPTION", "DESCRIZIONE"])
    material_col = find_col(df, ["MATERIAL", "COMPOSIZIONE"])
    color_col = find_col(df, ["COLOR", "COLORE"])
    season_col = find_col(df, ["SEASON", "STAGIONE"])
    collection_col = find_col(df, ["COLLECTION"])  # COL (clothing/accessories)
    cost_col = find_col(df, ["SELL IN", "COST"])
    retail_col = find_col(df, ["RETAIL", "SELL OUT"])
    size_col = find_col(df, ["SIZE", "TAGLIA"])
    qty_col = find_col(df, ["QTY", "QUANTITY", "QUANTITÀ", "QUANTITA"])

    required_cols = [brand_col, sku_col, size_col, qty_col]
    if any(c is None for c in required_cols):
        return rows

    for idx, row in df.iterrows():
        brand = normalize_text(row.get(brand_col))
        sku = normalize_text(row.get(sku_col))
        size = normalize_text(row.get(size_col))
        qty_raw = row.get(qty_col)
        if not brand or not sku or not size or not is_positive_qty(qty_raw):
            continue
        # load_images_by_merged_range 的 key 是 openpyxl 的 0-idx row
        # df.iterrows() 的 idx 是 0-idx = Excel 行号 - 2(标题占 1 行 + pandas 0-idx -1)。
        # 所以 df idx 0 对应 openpyxl 0-idx row 1,要 +1 才能匹配上。
        openpyxl_row = int(idx) + 1
        desc = normalize_text(row.get(desc_col)) if desc_col else ""
        cost = to_float(row.get(cost_col)) if cost_col else None
        retail = to_float(row.get(retail_col)) if retail_col else None
        rows.append(
            {
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": desc,
                "category": normalize_text(row.get(collection_col)) if collection_col else guess_product_group(desc),
                "material": normalize_text(row.get(material_col)) if material_col else "",
                "color": normalize_text(row.get(color_col)) if color_col else "",
                "season": normalize_text(row.get(season_col)) if season_col else "",
                "origin": "",  # 源无产地列
                "cost": cost,
                "cplus_text": "",
                "supply_price": cost,
                "retail_price": retail,
                "size": size,
                "quantity": int(round(float(to_float(qty_raw) or 0))),
                "vendor_id": vendor_id,
                "image_items": sheet_images.get("Foglio1", {}).get(openpyxl_row, []),
                "source_row_height": sheet_row_heights.get("Foglio1", {}).get(openpyxl_row + 1),
            }
        )
    return rows
