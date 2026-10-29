"""V1041 BIFFI vendor adapter."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from convert_vendor_quote_shared import (
    find_col,
    is_positive_qty,
    load_sheet_images_by_row,  # 抽图(.xls 走空,.xlsx 走 openpyxl)
    normalize_text,
    to_float,
)


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    # 抽图(.xls 走空,.xlsx 走 openpyxl)
    images_by_row = load_sheet_images_by_row(primary_file)
    df = pd.read_excel(primary_file, sheet_name=0)
    brand_col = find_col(df, ["Brand"])
    sku_col = find_col(df, ["Style number", "Item", "SKU"])
    season_col = find_col(df, ["Season"])
    material_col = find_col(df, ["Composition", "Material"])
    sex_col = find_col(df, ["Sex", "Gender", "Genere"])
    cost_col = find_col(df, ["COST", "cost"])
    retail_col = find_col(df, ["Retail", "RETAIL"])
    size_col = find_col(df, ["Size", "SIZE"])
    qty_col = find_col(df, ["Qty", "QTY", "Quantity"])

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
        cost = to_float(row.get(cost_col)) if cost_col else None
        retail = to_float(row.get(retail_col)) if retail_col else None
        # 性别(2026-08-14 加):V1041 源表 c5 'Sex' 列 = Man / Woman / Unisex,直接写到 gender
        gender_raw = normalize_text(row.get(sex_col)) if sex_col else ""
        # sex_val 是性别(不是 7 类目之一), 跟 7 类目体系不一致 → 不写到 category
        #   类别优先级: 数据源中直接表明 > 类目规则分类 > 空
        #   V1041 源表没类目列, 也没 description, 类目规则 guess 用不了 → 留空
        # ⚠️ image lookup:df idx 0 → openpyxl 0-idx 1(标题行 + pandas 0-idx -1)
        image_items = images_by_row.get(idx + 1, [])
        rows.append(
            {
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": "",
                "category": "",
                "material": normalize_text(row.get(material_col)) if material_col else "",
                "color": "",
                "gender": gender_raw,
                "season": normalize_text(row.get(season_col)) if season_col else "",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": cost,
                "retail_price": retail,
                "size": size,
                "quantity": int(round(float(to_float(qty_raw) or 0))),
                "vendor_id": vendor_id,
                "image_items": image_items,
            }
        )
    return rows
