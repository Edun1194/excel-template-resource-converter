"""V1031 vendor adapter — FENDI / NUGNES 系(意大利时尚品牌)。

来源:.xlsx 表结构化版式(列名: BRAND / MODELLO VARIANTE / DESCRIZIONE / COLORE / ...)

特殊点(与 V1021 / V1010 / V1038 的差异):
- 仓库 ID 默认 ``V1031``(由 vendor_id 决定,可在 dispatcher 覆盖)
- 品牌/C+/供货价 走 dispatcher 默认策略(源文件若有 BRAND 字段,parse_rows 读出来;
  用户在 CLI 传 --brand 可覆盖)
- 图加载:每行 1 张图,锚点行 = 数据行,用 ``load_sheet_images_by_row``(简单版)
- 列名都是意大利语(MODELLO VARIANTE / TAGLIA / QUANTITÀ 等),通过 ``find_col`` 容错匹配
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
import pandas as pd

from convert_vendor_quote_shared import (
    find_col,
    guess_product_group,
    is_positive_qty,
    load_sheet_images_by_row,  # ← 简单版,适配 FENDI 这种"1 行 1 图"版式
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    # 抽图(.xlsx 才有,.xls 走空 dict)— 用简单版(每行 1 图,锚点 = 数据行)
    images_by_row = load_sheet_images_by_row(primary_file)
    sheet_row_heights = load_sheet_row_heights(primary_file)
    df = pd.read_excel(primary_file)

    brand_col = find_col(df, ["BRAND"])
    sku_col = find_col(df, ["MODELLO VARIANTE", "MODELLO", "ITEM", "SKU"])
    desc_col = find_col(df, ["DESCRIZIONE", "DESCRIZIONE ARTICOLO"])
    color_col = find_col(df, ["COLORE", "COLOR"])
    category_col = find_col(df, ["CATEGORIA", "CATEGORY"])
    material_col = find_col(df, ["TESSUTO", "MATERIAL", "COMPOSIZIONE"])
    made_in_col = find_col(df, ["MADE IN", "ORIGIN"])
    season_col = find_col(df, ["STAGIONE", "SEASON"])
    cost_col = find_col(df, ["PREZZO DI COSTO", "COST", "COST EUR"])
    retail_col = find_col(df, ["PREZZO VENDITA", "PREZZO CONSIGLIATO", "RETAIL"])
    size_col = find_col(df, ["TAGLIA", "SIZE"])
    # TIPOLOGIA TAGLIE 列常常充当 size 占位 (e.g. "UNI" 表示 one-size),在 TAGLIA 为空时 fallback
    tipologia_col = find_col(df, ["TIPOLOGIA TAGLIE", "TIPOLOGIA"])
    qty_col = find_col(df, ["QUANTITÀ", "QUANTITA", "QTY", "QUANTITÀ"])

    # 鲁棒 fallback:部分 NUGNES 系供应商(LONGCHAMP PARIS 实测)表头规整但数据列填错位:
    #   - TAGLIA 全空,真正尺码在 TIPOLOGIA TAGLIE
    #   - QUANTITà 全 'UNI' 占位,真正 qty 在表头外的 Unnamed: <N> 数字列
    # 自动检测后切换列,避免整文件 0 行。
    if size_col and tipologia_col and size_col != tipologia_col:
        size_has_value = df[size_col].notna().any() and (
            df[size_col].dropna().astype(str).str.strip() != ""
        ).any()
        if not size_has_value:
            size_col = tipologia_col

    unnamed_qty_col: Optional[str] = None
    if qty_col:
        qty_has_numeric = df[qty_col].apply(is_positive_qty).any()
        if not qty_has_numeric:
            # 找第一列全是数字(>0)的 Unnamed: N 列
            for col in df.columns:
                if str(col).startswith("Unnamed:"):
                    series = df[col].dropna()
                    if len(series) and series.apply(is_positive_qty).all():
                        unnamed_qty_col = col
                        break

    # 必需列:sku / size / qty(brand 可缺失,留空让 dispatcher 决定)
    if not all([sku_col, size_col, qty_col]):
        return rows

    for idx, row in df.iterrows():
        sku = normalize_text(row.get(sku_col))
        size = normalize_text(row.get(size_col))
        if not size and tipologia_col and tipologia_col != qty_col:
            size = normalize_text(row.get(tipologia_col))
        qty_raw = row.get(unnamed_qty_col) if unnamed_qty_col else row.get(qty_col)
        if not sku or not size or not is_positive_qty(qty_raw):
            continue
        # 读源 BRAND 字段(若存在)。dispatcher 会:
        #   1. 清掉所有 brand
        #   2. 如果源有 brand 但用户没传 --brand,会把"清空"再"填回源 brand"
        # 所以 parse_rows 这层只管"读源 brand",不管"用不用"
        brand = normalize_text(row.get(brand_col)) if brand_col else ""
        excel_row = int(idx) + 2  # pandas iloc 是 0-idx,Excel 行号 +1,标题行 +1
        qty = int(round(float(to_float(qty_raw) or 0)))
        desc = normalize_text(row.get(desc_col)) if desc_col else ""
        cost = to_float(row.get(cost_col)) if cost_col else None
        retail = to_float(row.get(retail_col)) if retail_col else None
        rows.append(
            {
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": desc,
                "category": normalize_text(row.get(category_col)) if category_col else guess_product_group(desc),
                "material": normalize_text(row.get(material_col)) if material_col else "",
                "color": normalize_text(row.get(color_col)) if color_col else "",
                "season": normalize_text(row.get(season_col)) if season_col else "",
                "origin": normalize_text(row.get(made_in_col)) if made_in_col else "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": cost,
                "retail_price": retail,
                "size": size,
                "quantity": qty,
                "vendor_id": vendor_id,
                # ⚠️ 关键:df.iterrows() 的 idx 是 0-idx = Excel 行号 - 2(表头占 1 行,
                # pandas 0-idx 再 -1)。openpyxl 的 anchor._from.row 是 0-idx = Excel 行号 - 1。
                # 所以 df idx 0 对应 anchor._from.row 1,要 +1 才能匹配上。
                "image_items": images_by_row.get(idx + 1, []),
                "source_row_height": sheet_row_heights.get(idx + 1),
            }
        )
    return rows
