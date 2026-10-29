"""V1048 MALTEMPI vendor adapter."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from convert_vendor_quote_shared import (
    build_scalar_size_map,
    is_positive_qty,
    load_images_by_row_per_sheet,  # 替代 load_sheet_embedded_images(简单行匹配)
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    sheet_images = load_images_by_row_per_sheet(primary_file)
    sheet_row_heights = load_sheet_row_heights(primary_file)
    sheet_name = "Foglio1"
    df = pd.read_excel(primary_file, sheet_name=sheet_name, header=None)
    if df.empty:
        return rows

    header_row = 11
    scalar_col = 11
    qty_start_col = 12
    cost_col = 37
    retail_col = 38
    qty_end_col = cost_col - 1

    scalar_map = build_scalar_size_map(df, header_row, scalar_col, qty_start_col, qty_end_col)

    for r in range(header_row + 1, df.shape[0]):
        sku_raw = normalize_text(df.iat[r, 1] if df.shape[1] > 1 else "")
        # 2026-08-14:SKU 拼接符统一 `-`(源可能用 `.` 或 `_`,改 `-`)
        # 例: "8011693.A7026" → sku="8011693-A7026", color_code="A7026"
        sku = sku_raw.replace(".", "-").replace("_", "-")
        color_code = ""
        m = re.match(r"^(.+?)-([A-Za-z0-9]+)$", sku)
        if m and len(m.group(1)) >= 6:  # style 通常 ≥ 6 位
            # 看是否像是 `<style>-<color_code>` 结构(style 末尾是数字,color_code 是字母+数字)
            style_part, tail_part = m.group(1), m.group(2)
            if style_part[-1].isdigit() and tail_part[0].isalpha():
                color_code = tail_part
        brand = normalize_text(df.iat[r, 8] if df.shape[1] > 8 else "")
        scalar = normalize_text(df.iat[r, scalar_col] if scalar_col < df.shape[1] else "").upper()
        if not sku or not brand or not scalar:
            continue
        excel_row = r + 1  # pandas index to Excel row (1-indexed)
        # ⚠️ load_images_by_row_per_sheet 的 key 是 openpyxl 的 0-idx anchor row。
        # 这里 r 是 0-idx = Excel 行号 - 1(标题占 1 行,pandas 0-idx -1)。
        # 所以 r 对应 openpyxl anchor row r,直接用 r 作 key。
        row_images = sheet_images.get(sheet_name, {}).get(r, [])
        size_map = scalar_map.get(scalar, {})
        cost = to_float(df.iat[r, cost_col] if cost_col < df.shape[1] else None)
        retail = to_float(df.iat[r, retail_col] if retail_col < df.shape[1] else None)
        for c in range(qty_start_col, min(qty_end_col + 1, df.shape[1])):
            qty_raw = df.iat[r, c]
            if not is_positive_qty(qty_raw):
                continue
            size = normalize_text(size_map.get(c))
            if not size:
                continue
            rows.append(
                {
                    "brand": brand,
                    "sku": sku,
                    "image": "",
                    "description": normalize_text(df.iat[r, 7] if df.shape[1] > 7 else ""),
                    "category": normalize_text(df.iat[r, 5] if df.shape[1] > 5 else ""),
                    "material": normalize_text(df.iat[r, 9] if df.shape[1] > 9 else ""),
                    "color": normalize_text(df.iat[r, 10] if df.shape[1] > 10 else ""),
                    "color_code": color_code,
                    "season": "",
                    "origin": "",
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,  # 供货价留空，如有源数据供货价字段则取该值
                    "retail_price": retail,
                    "size": size,
                    "quantity": int(round(float(to_float(qty_raw) or 0))),
                    "vendor_id": vendor_id,
                    "image_items": row_images,
                    "source_row_height": sheet_row_heights.get(sheet_name, {}).get(excel_row),
                }
            )
    return rows
