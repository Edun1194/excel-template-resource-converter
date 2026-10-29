"""V1001 BARBOUR vendor adapter (2026-10-22).

适用文件:
  BARBOUR C+25.xlsx (Barbour 品牌货单, 1 sheet 'Worksheet')

源文件结构 (Worksheet, 392 行 x 22 列, **稀疏表**):
  R1: header (c2-c22)
  数据行非常稀疏 — 每个产品块间隔 15 行, 块内每个尺码一行 (每行都带完整产品信息)

  列映射 (1-idx):
    c1  = (空, 图片 anchor 在这列)
    c2  = designer (BARBOUR)
    c3  = id (内部数字 ID)
    c4  = variant_id
    c5  = designer (重复)
    c6  = gender (Woman / Man — **已经是模板规范值, 不用翻译**)
    c7  = macro-category (Clothing / Accessories — 太粗, 只 2 个值)
    c8  = category (Scarves, foulards / Down jacket, Parka... — 用于类目)
    c9  = product ('sku: LGS0054LGSTN11' — **去掉 'sku: ' 前缀就是货号**)
    c10 = product name (描述)
    c11 = description (长文本营销描述, 不进模板)
    c12 = season ('Fall Winter 2026/2027')
    c13 = ean (含尺码后缀, 不用)
    c14 = made in (Moldova / China / United Kingdom ...)
    c15 = composition (材质: '53% viscose 26% polyester 21% polyamide')
    c16 = internalId
    c17 = retail price (不含税, 带小数: 90.163934)
    c18 = retail price vat included (**含税整数: 110**)
    c19 = wholesale price (cost: 44)
    c20 = size (U / 8 UK / 10 UK ...)
    c21 = qty
    c22 = order qty (全空, 不用)

  图片: c1 列, anchor 在**产品块首行** (3, 18, 33, 48, ... 每 15 行一个),
        块内所有尺码行共享该图。

映射决策 (2026-10-22 用户确认):
  - 货号 = c9 去掉 'sku: ' 前缀 (不含尺码)
  - 零售价 = c18 (含税整数)
  - 季节 = c12 原值 ('Fall Winter 2026/2027')
  - 类目 = c8 (category), c7 macro-category 太粗不用
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    load_sheet_images_by_row,
    normalize_text,
    to_float,
)


VENDOR_ID = "V1001"

# 源文件表头特征 (用于 detect)
_SHEET_NAME = "Worksheet"
_REQUIRED_HEADERS = ("designer", "macro-category", "product name", "wholesale price")


def _detect_barbour_format(ws) -> bool:
    """detect BARBOUR 版式: R1 c2='designer' + c8='category' + c19='wholesale price'."""
    if ws is None or ws.max_row < 2 or ws.max_column < 19:
        return False
    h2 = str(ws.cell(1, 2).value or "").strip()
    h8 = str(ws.cell(1, 8).value or "").strip()
    h19 = str(ws.cell(1, 19).value or "").strip()
    return h2 == "designer" and h8 == "category" and h19 == "wholesale price"


def _strip_sku_prefix(raw: str) -> str:
    """'sku: LGS0054LGSTN11' → 'LGS0054LGSTN11'."""
    s = normalize_text(raw)
    if s.lower().startswith("sku:"):
        s = s[4:].strip()
    return s


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """Parse BARBOUR xlsx → output List[Dict]."""
    rows: List[Dict[str, object]] = []
    suffix = primary_file.suffix.lower()
    if suffix != ".xlsx":
        return rows

    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True)
    except Exception:
        return rows

    ws = wb[_SHEET_NAME] if _SHEET_NAME in wb.sheetnames else wb.active
    if not _detect_barbour_format(ws):
        return rows

    # 图片: c1 列 anchor 在产品块首行
    images_by_row = load_sheet_images_by_row(primary_file)

    # anchor 行 → 向下覆盖到下一个 anchor 之前 (块内所有尺码共享该图)
    anchor_rows = sorted(images_by_row.keys())          # 0-idx

    def _image_for(row_0idx: int) -> List[Dict[str, object]]:
        """给 0-idx 数据行找所属块的图 (取 <= row 的最近 anchor)."""
        best = None
        for ar in anchor_rows:
            if ar <= row_0idx:
                best = ar
            else:
                break
        return images_by_row.get(best, []) if best is not None else []

    for r in range(2, ws.max_row + 1):
        # 数据行判定: c3 (id) 有值。源文件很稀疏, 大量空行要跳过。
        if ws.cell(r, 3).value in (None, ""):
            continue

        sku = _strip_sku_prefix(ws.cell(r, 9).value)          # c9 product
        if not sku:
            continue

        brand = normalize_text(ws.cell(r, 2).value) or "BARBOUR"
        gender = normalize_text(ws.cell(r, 6).value)          # Woman / Man (已是规范值)
        category = normalize_text(ws.cell(r, 8).value)        # c8 category
        description = normalize_text(ws.cell(r, 10).value)    # c10 product name
        season = normalize_text(ws.cell(r, 12).value)         # c12 'Fall Winter 2026/2027'
        origin = normalize_text(ws.cell(r, 14).value)         # c14 made in
        material = normalize_text(ws.cell(r, 15).value)       # c15 composition
        size = normalize_text(ws.cell(r, 20).value)           # c20 size

        cost = to_float(ws.cell(r, 19).value)                 # c19 wholesale price
        retail = to_float(ws.cell(r, 18).value)               # c18 retail price VAT included
        quantity = to_float(ws.cell(r, 21).value)             # c21 qty

        row_images = _image_for(r - 1)

        rows.append({
            "brand": brand,
            "sku": sku,
            "image": "",
            "description": description,
            "category": category,
            "material": material,
            "color": "",           # 源文件无颜色列
            "color_code": "",
            "gender": gender,      # 源文件已是 Woman/Man, 直接用
            "season": season,      # 2026-10-22 用户确认: 存原值
            "origin": origin,
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,   # 2026-10-22 用户确认: 用 c18 含税价
            "size": size,
            "quantity": int(quantity) if quantity is not None else None,
            "vendor_id": vendor_id,
            "image_items": row_images,
            "source_row_height": 92.0 if row_images else None,
            "_row_num": r,
        })

    return rows


def get_vendor_meta() -> Dict[str, object]:
    return {
        "default_image_col_width": 40.0,
    }