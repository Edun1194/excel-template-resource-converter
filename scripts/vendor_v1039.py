"""V1039 ANTONIOLI vendor adapter."""

from __future__ import annotations

import openpyxl
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from convert_vendor_quote_shared import (
    build_scalar_size_map,
    format_cplus_text,
    is_positive_qty,
    load_sheet_embedded_images,
    normalize_text,
    to_float,
)


def _load_cplus_map(markup_file: Optional[Path]) -> Dict[str, Dict[str, float]]:
    """Load C+ mapping from CONDITIONS file.
    
    Returns: {brand_lower: {category_upper: cplus_percentage}}
    e.g., {"gucci": {"BAGS": 0.20, "SHOES": 0.15}}
    """
    if markup_file is None or not markup_file.exists():
        return {}
    
    try:
        wb = openpyxl.load_workbook(markup_file)
        ws = wb.active
        
        # Row 5 has category headers (Accessories, Bags, Clothing, Shoes)
        # Column 1 has brand names
        cat_map = {}  # col_index -> category_name
        for c in range(2, ws.max_column + 1):
            v = ws.cell(5, c).value
            if v:
                cat_map[c] = normalize_text(v).upper()
        
        cplus_map: Dict[str, Dict[str, float]] = {}
        for r in range(6, ws.max_row + 1):
            brand = normalize_text(ws.cell(r, 1).value or "").lower()
            if not brand:
                continue
            cplus_map[brand] = {}
            for c, cat in cat_map.items():
                v = ws.cell(r, c).value
                if v is not None:
                    cplus_map[brand][cat] = float(v)
        
        return cplus_map
    except Exception:
        return {}


def _normalize_brand(s: str) -> str:
    """Normalize brand name for matching - remove accents, spaces, hyphens."""
    import unicodedata
    # Remove accents
    normalized = unicodedata.normalize('NFKD', s)
    # Remove accents (喇 -> l)
    cleaned = ''.join(c for c in normalized if not unicodedata.combining(c))
    # Lowercase and remove spaces/hyphens/underscores
    return cleaned.lower().replace(' ', '').replace('-', '').replace('_', '')


def _match_brand(brand: str, cplus_map: Dict[str, Dict[str, float]]) -> Optional[Dict[str, float]]:
    """Match brand name to C+ map using fuzzy matching."""
    brand_norm = _normalize_brand(brand)
    # Exact match on normalized
    for key in cplus_map:
        if _normalize_brand(key) == brand_norm:
            return cplus_map[key]
    # Prefix match
    for key in cplus_map:
        key_norm = _normalize_brand(key)
        if key_norm.startswith(brand_norm) or brand_norm.startswith(key_norm):
            return cplus_map[key]
    # Partial match
    for key in cplus_map:
        key_norm = _normalize_brand(key)
        if brand_norm in key_norm or key_norm in brand_norm:
            return cplus_map[key]
    return None


def parse_rows(primary_file: Path, markup_file: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    df = pd.read_excel(primary_file, sheet_name="CATALOGO", header=None)
    if df.empty or df.shape[0] < 3:
        return rows

    # Load C+ mapping
    cplus_map = _load_cplus_map(markup_file)

    header_row = None
    for r in range(min(20, df.shape[0])):
        row_text = "|".join(normalize_text(v).upper() for v in df.iloc[r].tolist())
        if "IDARTICOLO" in row_text and "SCALARINO" in row_text:
            header_row = r
            break
    if header_row is None:
        return rows

    # Load embedded images
    sheet_images = load_sheet_embedded_images(primary_file)
    cat_images = sheet_images.get("CATALOGO", {})

    header_map: Dict[str, int] = {}
    for c in range(df.shape[1]):
        key = normalize_text(df.iat[header_row, c]).upper()
        if key:
            header_map[key] = c

    def cidx(*names: str, default: Optional[int] = None) -> int:
        for n in names:
            idx = header_map.get(n.upper())
            if idx is not None:
                return idx
        if default is not None:
            return default
        raise KeyError(names[0])

    scalar_col = cidx("SCALARINO", default=24)
    qty_start_col = scalar_col + 1
    total_col = cidx("TOTALE", default=70)
    cost_col = cidx("COSTO", default=8)
    retail_col = cidx("RETAIL", default=9)
    brand_col = cidx("LINEA", default=10)
    category_col = cidx("REPARTO", default=12)
    desc_col = cidx("DESCRIZIONE", default=13)
    sku_col = cidx("MODELLO", default=14)
    season_col = cidx("STAGIONE", default=15)
    color_col = cidx("VARIANTE", default=18)
    origin_col = cidx("MARCA", default=19)
    material_col = cidx("MATERIALE", default=20)

    scalar_map = build_scalar_size_map(df, header_row, scalar_col, qty_start_col, total_col - 1)

    for r in range(header_row + 1, df.shape[0]):
        scalar = normalize_text(df.iat[r, scalar_col] if scalar_col < df.shape[1] else "").upper()
        sku = normalize_text(df.iat[r, sku_col] if sku_col < df.shape[1] else "")
        if not scalar or not sku:
            continue
        size_map = scalar_map.get(scalar, {})
        cost = to_float(df.iat[r, cost_col] if cost_col < df.shape[1] else None)
        retail = to_float(df.iat[r, retail_col] if retail_col < df.shape[1] else None)
        brand = normalize_text(df.iat[r, brand_col] if brand_col < df.shape[1] else "")
        desc = normalize_text(df.iat[r, desc_col] if desc_col < df.shape[1] else "")
        category = normalize_text(df.iat[r, category_col] if category_col < df.shape[1] else "")
        
        # Look up C+ value
        brand_cplus = _match_brand(brand, cplus_map)
        category_upper = category.upper()
        cplus_val = None
        if brand_cplus and category_upper in brand_cplus:
            cplus_val = brand_cplus[category_upper]
        
        for c in range(qty_start_col, min(total_col, df.shape[1])):
            qty_raw = df.iat[r, c]
            if not is_positive_qty(qty_raw):
                continue
            size = normalize_text(size_map.get(c))
            if not size:
                continue
            # Excel row is 1-indexed, our header_row is 0-indexed
            excel_row = r + 1
            row_images = cat_images.get(excel_row, [])
            
            # Format C+ using standardized function
            cplus_text = format_cplus_text(cplus_val)
            
            rows.append(
                {
                    "brand": brand,
                    "sku": sku,
                    "image": "",
                    "description": desc,
                    "category": category,
                    "material": normalize_text(df.iat[r, material_col] if material_col < df.shape[1] else ""),
                    "color": normalize_text(df.iat[r, color_col] if color_col < df.shape[1] else ""),
                    "season": normalize_text(df.iat[r, season_col] if season_col < df.shape[1] else ""),
                    "origin": normalize_text(df.iat[r, origin_col] if origin_col < df.shape[1] else ""),
                    "cost": cost,
                    "cplus_text": cplus_text,
                    "supply_price": cost,
                    "retail_price": retail,
                    "size": size,
                    "quantity": int(round(float(to_float(qty_raw) or 0))),
                    "vendor_id": vendor_id,
                    "image_items": row_images,
                }
            )
    return rows
