"""A1001 CANADA GOOSE HK 现货 .xls 适配器 (2026-10-12).

适配文件: CANADA GOOSE hk 现货.xls (OLE2 老格式, 1 sheet)

文件结构 (Sheet1, 30 行 18 列):
  R1: header (c1=示意图 / c2=爱莎码 / c3=商品名称 / c4=国际码 /
            c5-c14=size labels (XXXS/XXS/XS/S/M/L/XL/XXL/XXXL/XXXXL) /
            c15=总数 / c16=海外零售价 / c17=出厂价 / c18=批发价加币)
  R3, R5, R7, ...: data rows (商品 + size labels + 总数 + 价格)
  R4, R6, R8, ...: qty 分布行 (c5-c14 是 per-size qty, 数字)

字段映射 (1-idx):
  c2=爱莎码 (国内 SKU) / c3=商品名称 (description) / c4=国际码 (modello) /
  c5-c14=size labels (10 个) / c15=总数 / c16=海外零售价 /
  c17=出厂价 (cost) / c18=批发价加币 (wholesale CAD)

类目/性别/颜色从 c3 商品名称推断 (中文字符串):
  - 羽绒服 / POLO / T恤 / 休闲裤 / 外套 / 大衣 / 针织衫
  - 男士/女士
  - 黑色/白色/深蓝色/浅灰色/...

图片: framework load_xls_ole2_jpeg_images (OLE2 流), anchor row 是 0-idx,
       实际图在商品行 (0-idx 2, 4, 14, 20, 22 等)
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
import xlrd

from convert_vendor_quote_shared import (
    load_sheet_images_by_row,
    load_xls_ole2_jpeg_images,
    normalize_text,
    to_float,
)


VENDOR_ID = "A1001"

# 1-idx size 列范围
SIZE_LABEL_COL_START = 5   # XXXS
SIZE_LABEL_COL_END = 15    # XXXXL (inclusive)


_CATEGORY_KEYWORDS = [
    ("羽绒服", "Down Jacket"),
    ("POLO", "Polo"),
    ("T恤", "Tee"),
    ("休闲裤", "Pants"),
    ("外套", "Jacket"),
    ("大衣", "Coat"),
    ("针织衫", "Knitwear"),
    ("卫衣", "Sweatshirt"),
    ("衬衫", "Shirt"),
    ("风衣", "Trench"),
    ("马甲", "Vest"),
]

_GENDER_KEYWORDS = [
    ("男士", "Man"),
    ("男", "Man"),
    ("女士", "Woman"),
    ("女", "Woman"),
    ("中性", "Unisex"),
    ("儿童", "Kid"),
    ("童", "Kid"),
]

_COLOR_KEYWORDS = [
    "黑色", "白色", "深蓝色", "浅灰色", "深灰色", "灰色",
    "蓝色", "红色", "绿色", "米色", "驼色", "紫色", "粉色",
    "黄色", "橘色", "棕色", "卡其", "藏青", "墨绿", "酒红",
    "象牙", "杏色", "军绿",
]


def _infer_category(desc: str) -> str:
    for kw, cat in _CATEGORY_KEYWORDS:
        if kw in desc:
            return cat
    return ""


def _infer_gender(desc: str) -> str:
    # 男士/女士 优先于 男/女
    for kw, g in _GENDER_KEYWORDS:
        if kw in desc:
            return g
    return ""


def _infer_color(desc: str) -> str:
    for c in _COLOR_KEYWORDS:
        if c in desc:
            return c
    return ""


def parse_rows(
    primary_file: Path,
    _: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    suffix = primary_file.suffix.lower()

    # 2026-10-12:同时支持 .xls (OLE2 老格式, xlrd) 和 .xlsx (openpyxl)
    if suffix == ".xlsx":
        try:
            wb = openpyxl.load_workbook(str(primary_file), data_only=True)
        except Exception:
            return rows
        if "Sheet1" not in wb.sheetnames:
            return rows
        ws = wb["Sheet1"]
        if ws.max_row < 3:
            return rows

        def cell(r_1idx: int, c_1idx: int):
            return ws.cell(r_1idx, c_1idx).value

        # 图 (openpyxl 0-idx row = 1-idx - 1)
        images_by_row = load_sheet_images_by_row(primary_file)
    else:
        # .xls
        try:
            wb = xlrd.open_workbook(str(primary_file))
        except Exception:
            return rows
        if "Sheet1" not in wb.sheet_names():
            return rows
        ws = wb.sheet_by_name("Sheet1")
        if ws.nrows < 3:
            return rows

        def cell(r_1idx: int, c_1idx: int):
            return ws.cell_value(r_1idx - 1, c_1idx - 1)

        # 图 (0-idx row)
        images_by_row = load_xls_ole2_jpeg_images(primary_file)

    # 2026-10-14: detect 文件版式
    # 老版式 (CANADA GOOSE hk 现货, 18 列): R3 起每 2 行 data
    # 新版式 (加币 GC-A1001-3, 15 列): R4 起每 3 行 data
    #   - 商品名(c4) vs 国际码(c3) 互换
    #   - size labels c5-c12 (8 个, XXS-XXXL) vs 老 c5-c14 (10 个, XXXS-XXXXL)
    #   - 字段: c2=爱莎码, c3=国际码, c4=商品名称, c5-c12=size labels, c13=总数, c14=出厂价, c15=批发价
    max_col = ws.max_column if suffix == ".xlsx" else ws.ncols
    if max_col == 15:
        return _parse_a1001_cad_gc(ws, suffix, cell, images_by_row, vendor_id)

    # 2026-10-15: ASA 格式 (13 列): CANADA GOOSE ASA -2.xls
    #   字段 (1-idx): c1=示意图 / c2=爱莎码(SKU) / c3=CN NAME(desc) /
    #                  c4=ARTICLE(modello) /
    #                  c5=XS / c6=S / c7=M / c8=L / c9=XL / c10=XXL /
    #                  c11=PCS(total qty) / c12=出厂价(cost) / c13=批发价(wholesale CAD)
    #   结构: R1=header / R5,R9,R13,...=data 行 / R6,R10,R14,...=qty 行
    #   特点: size labels 6 个(XS-XXL), qty 行 c6-c9 有 S/M/L/XL, XS/XXL 可能为空
    if max_col == 13:
        return _parse_a1001_asa(ws, suffix, cell, images_by_row, vendor_id)

    # 老版式 (18 列)
    if max_col < 18:
        return rows
    return _parse_a1001_legacy(ws, suffix, cell, images_by_row, vendor_id)


def _parse_a1001_asa(ws, suffix, cell, images_by_row, vendor_id):
    """ASA 格式 (13 列, 2026-10-15): CANADA GOOSE ASA -2.xls.

    字段 (1-idx):
      c1=示意图 / c2=爱莎码(SKU) / c3=CN NAME(description) /
      c4=ARTICLE(modello) /
      c5=XS / c6=S / c7=M / c8=L / c9=XL / c10=XXL /
      c11=PCS(total qty) / c12=出厂价(cost) / c13=批发价(wholesale CAD)

    结构: R1=header / R5,R9,R13,...=data 行 / R6,R10,R14,...=qty 行
    qty 行 c6-c9: S/M/L/XL 的数量 (XS c5 和 XXL c10 可能为空)
    """
    rows: List[Dict[str, object]] = []
    max_data_row = ws.max_row if suffix == ".xlsx" else ws.nrows
    # data 行: R5, R9, R13, ... (每 4 行一个 block, data 在第 1 行)
    for r in range(5, max_data_row + 1, 4):
        sku_aisa = str(cell(r, 2) or "").strip()
        if not sku_aisa:
            continue

        description = normalize_text(cell(r, 3))
        modello = normalize_text(cell(r, 4))
        total_qty = to_float(cell(r, 11)) or 0
        cost = to_float(cell(r, 12))
        wholesale_cad = to_float(cell(r, 13))
        retail = None  # ASA 格式没有零售价字段

        category = _infer_category(description)
        gender = _infer_gender(description)
        color = _infer_color(description)

        row_images = images_by_row.get(r - 1, [])
        sku_final = modello or sku_aisa

        # qty 行: r+1
        qty_row = r + 1
        emitted = False
        # size labels c5-c10 (6 个: XS/S/M/L/XL/XXL)
        for c_idx in range(5, 11):
            size_label = normalize_text(cell(r, c_idx))
            if not size_label:
                continue
            qty_v = cell(qty_row, c_idx)
            try:
                qty = int(round(float(qty_v or 0)))
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue

            rows.append({
                "brand": "CANADA GOOSE",
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
                "material": "",
                "color": color,
                "color_code": "",
                "gender": gender,
                "season": "",
                "origin": "Canada",
                "cost": cost,
                "cplus_text": "",
                "supply_price": wholesale_cad,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
            emitted = True

    return rows


def _parse_a1001_legacy(ws, suffix, cell, images_by_row, vendor_id):
    """老版式: 18 列, R3 R5 R7... data 行, R+1 qty 行."""
    rows: List[Dict[str, object]] = []
    # data 行: R3, R5, R7, ... (1-idx, 奇数从 3 开始)
    max_data_row = ws.max_row if suffix == ".xlsx" else ws.nrows
    for r in range(3, max_data_row + 1, 2):
        sku_aisa = str(cell(r, 2) or "").strip()
        if not sku_aisa:
            continue

        description = normalize_text(cell(r, 3))
        modello = normalize_text(cell(r, 4))
        total_qty = to_float(cell(r, 15)) or 0
        retail = to_float(cell(r, 16))
        cost = to_float(cell(r, 17))
        wholesale_cad = to_float(cell(r, 18))

        category = _infer_category(description)
        gender = _infer_gender(description)
        color = _infer_color(description)

        # 图 (0-idx row = r - 1)
        row_images = images_by_row.get(r - 1, [])

        # 货号: 国际码 (modello),不用爱莎码 (2026-10-12 用户改)
        sku_final = modello or sku_aisa

        # qty 行: r+1 (下一行)
        qty_row = r + 1
        emitted = False
        for c_idx in range(SIZE_LABEL_COL_START, SIZE_LABEL_COL_END):
            size_label = normalize_text(cell(r, c_idx))
            if not size_label:
                continue
            qty_v = cell(qty_row, c_idx)
            try:
                qty = int(round(float(qty_v or 0)))
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue

            rows.append({
                "brand": "CANADA GOOSE",
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
                "material": "",
                "color": color,
                "color_code": "",
                "gender": gender,
                "season": "",
                "origin": "Canada",
                "cost": cost,
                "cplus_text": "",
                "supply_price": wholesale_cad,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
            emitted = True

        # Fallback: c5-c14 全 0 但 c15 total > 0
        if not emitted and total_qty > 0:
            rows.append({
                "brand": "CANADA GOOSE",
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
                "material": "",
                "color": color,
                "color_code": "",
                "gender": gender,
                "season": "",
                "origin": "Canada",
                "cost": cost,
                "cplus_text": "",
                "supply_price": wholesale_cad,
                "retail_price": retail,
                "size": "TOTAL",
                "quantity": int(round(total_qty)),
                "vendor_id": vendor_id,
                "image_items": row_images,
            })

    return rows


def _parse_a1001_cad_gc(ws, suffix, cell, images_by_row, vendor_id):
    """新版式 (加币 GC-A1001-3, 2026-10-14): 15 列, R4 R7 R10... data, R+1 qty.

    字段 (1-idx):
      c2=爱莎码 / c3=国际码 (modello) / c4=商品名称 (description) /
      c5-c12=size labels (XXS/XS/S/M/L/XL/XXL/XXXL, 8 个) /
      c13=总数 / c14=出厂价 (cost) / c15=批发价加币 (wholesale CAD)

    与老版式 (18 列) 差异:
      - 18→15 列, 砍了"海外零售价"列
      - 商品名(c4) 与国际码(c3) 互换位置
      - size labels 10→8 (c5-c12 取代 c5-c14)
      - data 行从 R3 移到 R4, 块间距 2→3
    """
    rows: List[Dict[str, object]] = []
    max_data_row = ws.max_row if suffix == ".xlsx" else ws.nrows
    # data 行: R4, R7, R10, R13, ... (每 3 行)
    for r in range(4, max_data_row + 1, 3):
        sku_aisa = str(cell(r, 2) or "").strip()
        if not sku_aisa:
            continue

        modello = normalize_text(cell(r, 3))           # 新版式 c3 = 国际码
        description = normalize_text(cell(r, 4))        # 新版式 c4 = 商品名称
        total_qty = to_float(cell(r, 13)) or 0          # c13 = 总数
        cost = to_float(cell(r, 14))                    # c14 = 出厂价
        wholesale_cad = to_float(cell(r, 15))           # c15 = 批发价加币
        retail = None  # 新版式没有海外零售价字段

        category = _infer_category(description)
        gender = _infer_gender(description)
        color = _infer_color(description)

        row_images = images_by_row.get(r - 1, [])

        # 货号: 国际码 (modello),不用爱莎码 (2026-10-12 用户改)
        sku_final = modello or sku_aisa

        # qty 行: r+1
        qty_row = r + 1
        emitted = False
        # size labels c5-c12 (8 个)
        for c_idx in range(5, 13):
            size_label = normalize_text(cell(r, c_idx))
            if not size_label:
                continue
            qty_v = cell(qty_row, c_idx)
            try:
                qty = int(round(float(qty_v or 0)))
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue

            rows.append({
                "brand": "CANADA GOOSE",
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
                "material": "",
                "color": color,
                "color_code": "",
                "gender": gender,
                "season": "",
                "origin": "Canada",
                "cost": cost,
                "cplus_text": "",
                "supply_price": wholesale_cad,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
            emitted = True

        # Fallback: c5-c12 全 0 但 c13 total > 0
        if not emitted and total_qty > 0:
            rows.append({
                "brand": "CANADA GOOSE",
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
                "material": "",
                "color": color,
                "color_code": "",
                "gender": gender,
                "season": "",
                "origin": "Canada",
                "cost": cost,
                "cplus_text": "",
                "supply_price": wholesale_cad,
                "retail_price": retail,
                "size": "TOTAL",
                "quantity": int(round(total_qty)),
                "vendor_id": vendor_id,
                "image_items": row_images,
            })

    return rows


def get_vendor_meta() -> Dict[str, object]:
    return {"default_image_col_width": 40.0}
