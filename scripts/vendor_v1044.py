"""V1044 RATTI vendor adapter."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from convert_vendor_quote_shared import (
    COMMON_SIZE_HEADERS,
    calc_supply_price,
    first_non_missing,
    format_cplus_text,
    guess_product_group,
    is_positive_qty,
    load_images_by_row_per_sheet,  # 替代 load_sheet_embedded_images(多 sheet + 简单行匹配)
    load_sheet_row_heights,
    normalize_brand,
    normalize_text,
    parse_cplus,
    to_float,
)


def _extract_size_columns(columns) -> List[str]:
    size_cols: List[str] = []
    for c in columns:
        c_norm = normalize_text(c).upper()
        if not c_norm:
            continue
        if c_norm in {
            "TYPE", "QTY", "TOTAL", "FOTO", "BRAND", "DESCRIPTION",
            "MODEL", "TEXTYLE", "COLOR", "COLOR DESC", "COST EUR", "RETAIL",
        }:
            continue
        if c_norm.startswith("UNNAMED"):
            continue
        if c_norm in COMMON_SIZE_HEADERS or re.fullmatch(r"\d+(?:[.,]\d+)?", c_norm):
            size_cols.append(c)
    return size_cols


def _build_markup_map(markup_file: Optional[Path]) -> Dict[Tuple[str, str, str], float]:
    result: Dict[Tuple[str, str, str], float] = {}
    if markup_file is None or not Path(markup_file).exists():
        return result

    women_men = pd.read_excel(markup_file, sheet_name="Foglio1", header=None)
    kids = pd.read_excel(markup_file, sheet_name="kids", header=None)

    for _, row in women_men.iterrows():
        brand = normalize_brand(row.iloc[0] if len(row) > 0 else None)
        if not brand or brand in {"BRAND", "A", "B", "C", "D", "E", "F", "G", "H"}:
            continue
        women_cols = {"RTW": 1, "SHOES": 2, "LEATHER": 3, "ACC.RIES": 4}
        men_cols = {"RTW": 5, "SHOES": 6, "LEATHER": 7, "ACC.RIES": 8}
        for group, idx in women_cols.items():
            c = parse_cplus(row.iloc[idx] if idx < len(row) else None)
            if c is not None:
                result[(brand, "W", group)] = c
        for group, idx in men_cols.items():
            c = parse_cplus(row.iloc[idx] if idx < len(row) else None)
            if c is not None:
                result[(brand, "M", group)] = c

    for _, row in kids.iterrows():
        brand = normalize_brand(row.iloc[0] if len(row) > 0 else None)
        if not brand or brand == "BRAND":
            continue
        kids_cols = {"RTW": 1, "SHOES": 2, "ACC.RIES": 3}
        for group, idx in kids_cols.items():
            c = parse_cplus(row.iloc[idx] if idx < len(row) else None)
            if c is not None:
                result[(brand, "K", group)] = c

    return result


def _pick_cplus(
    markup_map: Dict[Tuple[str, str, str], float], brand: str, segment: str, description: str,
) -> Optional[float]:
    key_brand = normalize_brand(brand)
    group = guess_product_group(description)
    for seg in [segment, "W" if segment == "w" else "M" if segment == "m" else "K"]:
        c = markup_map.get((key_brand, str(seg).upper(), group))
        if c is not None:
            return c
    for seg in [segment, "W" if segment == "w" else "M" if segment == "m" else "K"]:
        c = markup_map.get((key_brand, str(seg).upper(), "RTW"))
        if c is not None:
            return c
    return None


def _build_size_label_map(
    df: pd.DataFrame,
    size_cols: List[str],
) -> Dict[int, Dict[str, str]]:
    """扫整张表,记录每个数据行(AVAIL)对应的"自定义尺码表"。

    V1044 文件结构:
      - 每 6 行一段:header_row (TYPE='TYPE') → data_row (TYPE='AVAIL') → order_row → ... → 空行
      - header_row 里 size 列写的是这一段数据的**实际尺码标签**
        (有些段写 44/46/48/50/52... 替代默认的 XXXS/XXS/XS...)
      - 数据行(AVAIL)的 size 位置对应 header_row 的尺码标签,不是列头的尺码名

    Returns: {df_idx -> {size_col_name: actual_size_label}}
    """
    result: Dict[int, Dict[str, str]] = {}
    current_labels: Dict[str, str] = {}  # 最新一个 header_row 的 size labels
    for idx, row in df.iterrows():
        type_v = normalize_text(row.get("TYPE")).upper()
        if type_v == "TYPE":
            # 这是 header_row — 记录 size labels
            current_labels = {
                c: normalize_text(row.get(c))
                for c in size_cols
                if normalize_text(row.get(c))
            }
        elif type_v == "AVAIL":
            # 数据行 — 继承最近的 header_row 的 labels
            if current_labels:
                result[int(idx)] = dict(current_labels)
    return result


def parse_rows(primary_file: Path, markup_file: Path, vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    markup_map = _build_markup_map(markup_file)
    sheet_images = load_images_by_row_per_sheet(primary_file)
    sheet_row_heights = load_sheet_row_heights(primary_file)

    workbook = pd.ExcelFile(primary_file)
    for sheet in workbook.sheet_names:
        segment = sheet[:1].upper()
        df = pd.read_excel(primary_file, sheet_name=sheet)
        if "TYPE" not in df.columns:
            continue
        size_cols = _extract_size_columns(df.columns)
        # 先扫整张表,建立"AVAIL 行 → 自定义尺码标签"映射(处理 B90886 那种
        # header 写 44/46/48/50 替代 XXXS/XXS/XS 的情况)
        size_label_map = _build_size_label_map(df, size_cols)
        filtered = df[df["TYPE"].astype(str).str.upper().eq("AVAIL")]
        for idx, row in filtered.iterrows():
            brand = normalize_text(row.get("BRAND"))
            if not brand or brand.upper() == "BRAND":
                continue
            model = normalize_text(row.get("MODEL"))
            if not model:
                continue
            # 货号 = MODEL + "-" + COLOR(同一 MODEL 在不同 COLOR 下是不同产品,
            # 不拼接会冲突导致尺码错乱。COLOR 是 3 位代码如 AAB / 001 / AUZ)。
            # 格式用 "-" 分隔(参考 SKILL.md 货号定义:{STYLE}-{颜色代码})。
            color_code = normalize_text(row.get("COLOR"))
            sku = f"{model}-{color_code}" if color_code else model
            excel_row = int(idx) + 2
            # ⚠️ load_images_by_row_per_sheet 的 key 是 openpyxl 的 0-idx anchor row。
            # df.iterrows() 的 idx 是 0-idx = Excel 行号 - 2(标题占 1 行 + pandas 0-idx -1)。
            # 所以 df idx 0 对应 openpyxl anchor row 1,要 +1 才能匹配上。
            row_images = sheet_images.get(sheet, {}).get(int(idx) + 1, [])
            cost = to_float(row.get("COST EUR"))
            retail = to_float(row.get("RETAIL"))
            description = normalize_text(row.get("DESCRIPTION"))
            color = first_non_missing(
                row.get("COLOR DESC"), row.get("COLOR"), ignore_values={"."},
            )
            material = first_non_missing(row.get("TEXTYLE"), ignore_values={""})
            cplus = _pick_cplus(markup_map, brand, segment, description)
            supply = calc_supply_price(cost, cplus)

            # 优先用 header_row 自定义尺码标签(44/46/48/50 等),没有就 fallback 到列头
            custom_labels = size_label_map.get(int(idx), {})
            for sc in size_cols:
                qty_raw = row.get(sc)
                if not is_positive_qty(qty_raw):
                    continue
                qty = int(round(float(to_float(qty_raw) or 0)))
                size_label = custom_labels.get(sc) or normalize_text(sc)
                rows.append(
                    {
                        "brand": brand,
                        "sku": sku,
                        "image": "",
                        "description": description,
                        "category": guess_product_group(description),
                        "material": material,
                        "color": color,
                        "season": "",
                        "origin": "",
                        "cost": cost,
                        "cplus_text": format_cplus_text(cplus),
                        "supply_price": supply,
                        "retail_price": retail,
                        "size": size_label,
                        "quantity": qty,
                        "vendor_id": vendor_id,
                        "image_items": row_images,
                        "source_row_height": sheet_row_heights.get(sheet, {}).get(excel_row),
                    }
                )
    return rows
