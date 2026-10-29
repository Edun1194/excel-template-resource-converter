"""V1015-补货单4 EA7 C-20 vendor adapter (2026-08-03 新建)

适配文件:Emporio Armani 7 (EA7) 工厂的 Inventory Details 导出 (EA C-20)。
跟 V1015-补货单 (PRADA) 同样是 per-row 模式,但字段映射不同:

  - V1015-补货单 (PRADA):   1 sheet, header R5, 11 列 (Periodo/Cod/Abbinamento/...)
  - V1015-补货单2 (MIUMIU): 3 sheet, header R6, 7/22 列
  - V1015-补货单3 (GA):     4 sheet, header R4, **横排**尺码 (c15+)
  - V1015-补货单4 (本):     1 sheet, header R8, 17 列 (per-row 纵排尺码)

结构(基于 1-idx):
  R1-R7 = 元数据 (EA7 / Inventory Details / Report Generated / Filter By / Price Type)
  R8    = header
  R9+   = 数据

header 列(0-idx):
  c1 = Image           ← 空 col,没内嵌图
  c2 = Linesheet       ← 品牌 + 性别 + 季节 + 类别 (e.g. "EA7 Man Apparel SS26")
                          多 linesheet 用 ", " 拼接
  c3 = Style Number    ← 完整 SKU (e.g. "245055_3R910")
  c4 = Style Name      ← 描述 (e.g. "TRAIN ZIP AROUND WALLET")
  c5 = Fabrication Name ← 材质描述
  c6 = Fabrication Code ← 材质代码
  c7 = Color Name      ← e.g. "MC426-Black/ Green Bay"
  c8 = Color Code      ← e.g. "MC426"
  c9 = Size            ← e.g. "TU" / "S/M" / "L/XL" / "11" / "11.5"
  c10= Units           ← 库存数量
  c11= Wholesale Price ← cost
  c12= Season Name     ← e.g. "Spring-Summer" (多 sheet 用 ", " 拼接)
  c13= Color Group
  c14= Event
  c15= Product Type    ← 类目 (e.g. "Piccola Pelletteria" / "Maglieria" / "Tessuto")
  c16= Range
  c17= Theme

字段映射:
  SKU        ← c3 Style Number (已是完整 SKU,e.g. "245055_3R910")
  描述       ← c4 Style Name
  颜色       ← c7 Color Name
  色号       ← c8 Color Code
  材质       ← c5 Fabrication Name
  季节       ← c12 Season Name (取第一个)
  cost       ← c11 Wholesale Price
  类目       ← c15 Product Type
  品牌       ← "" (留空,main 走文件名兜底 → "EA7" / "EMPORIO ARMANI")
  性别       ← 从 c2 Linesheet 提取 "Man"/"Woman"/"Unisex"/"Kids"
  云仓        ← VENDOR_ID = "V1015" (工厂=V1015,补货单4 是子类型)

Linesheet 解析 (e.g. "EA7 Leather & Shoes SS26"):
  - 移除 "EA7 " 前缀 → "Leather & Shoes SS26"
  - 移除 " SS26" / " Carry-Over SS26" 等 season 后缀
  - 移除 " Carry-Over" 字样
  - 性别关键词: Man / Woman / Unisex / Kids / Leather & Shoes / Accessories / Swimwear / Apparel
  - 多 Linesheet (", " 拼接) 任一含关键词就命中

⚠️ EA C-20 导出文件**有嵌入图**(c1 Image 列是空 col,但 ws._images 实际有 2361+ 张,
  per-row 模式 — 14531 数据行中 2361 行有图,占 16.3%)。
  用 openpyxl ws._images 按 anchor._from.row (0-idx) 分桶,每行 1 张图。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 V1015 (2026-07-31 用户要求),即使 vendor code 是 V1015-补货单4
VENDOR_ID = "V1015"

# R8 是 header (1-idx)
HEADER_ROW = 8


def get_vendor_meta() -> Dict[str, object]:
    """V1015-补货单4 vendor 元数据。"""
    return {
        "default_image_col_width": 40.0,
    }


# 0-idx 关键列位置(基于 EA C-20 17 列版式)
COL_LINESHEET = 1        # c2 Linesheet (含品牌+性别+季节+类别)
COL_STYLE_NUMBER = 2     # c3 Style Number (完整 SKU)
COL_STYLE_NAME = 3       # c4 Style Name
COL_FABRICATION_NAME = 4 # c5 Fabrication Name (材质)
COL_COLOR_NAME = 6       # c7 Color Name
COL_COLOR_CODE = 7       # c8 Color Code
COL_SIZE = 8             # c9 Size
COL_UNITS = 9            # c10 Units
COL_WHOLESALE = 10       # c11 Wholesale Price
COL_SEASON = 11          # c12 Season Name
COL_PRODUCT_TYPE = 14    # c15 Product Type (类目)


# 性别关键词(从 Linesheet 提取,优先级顺序)
_GENDER_KEYWORDS = ["Man", "Woman", "Unisex", "Kids"]


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 openpyxl 0-idx row 分桶图(每行 1 张)。

    实测 EA C-20 14531 数据行中 ~2361 行有图(占 16.3%),anchor row 0-idx
    从 8 起(R9 = 数据起始,header 在 R8 = 0-idx 7)。
    """
    out: Dict[int, List[Dict]] = {}
    for img in ws._images:
        anchor = img.anchor
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        try:
            data = img._data() if hasattr(img, "_data") else None
        except Exception:
            data = None
        if data is None:
            continue
        row = anchor._from.row
        out.setdefault(row, []).append({
            "bytes": data,
            "format": img.format or "jpeg",
        })
    return out


def _pick_gender_from_linesheet(linesheet_raw: str) -> str:
    """从 Linesheet 提取性别。

    e.g. "EA7 Man Apparel SS26" → "Man"
         "EA7 Leather & Shoes SS26, EA7 Leather & Shoes Carry-Over SS26" → ""
         "EA7 Unisex Apparel SS26" → "Unisex"
         "EA7 Kids SS26" → "Kids"
    """
    if not linesheet_raw:
        return ""
    sv = str(linesheet_raw)
    # 多 Linesheet 用 ", " 拼接,任一含关键词就命中
    for kw in _GENDER_KEYWORDS:
        # 用 word boundary,避免 "Manpower" 误匹配
        if re.search(rf"\b{kw}\b", sv):
            return kw
    return ""


def _pick_first_season(season_raw: str) -> str:
    """从 "Spring-Summer, Spring-Summer" 取第一个非空值。"""
    if not season_raw:
        return ""
    sv = str(season_raw)
    parts = [p.strip() for p in sv.split(",") if p.strip()]
    return parts[0] if parts else ""


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 EA C-20 .xlsx,1 sheet 全部转,返回 row dict 列表。

    仓库 ID 统一用 VENDOR_ID = "V1015" (2026-07-31),忽略传入 vendor_id。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    # 用模块级 VENDOR_ID 覆盖 vendor_id,保证「云仓」列稳定是 "V1015"
    vendor_id = VENDOR_ID

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    for ws in wb.worksheets:
        # 防御性检查:R8 c2 (0-idx 1) 应是 "Linesheet" 才是 EA C-20 版式
        h_linesheet = normalize_text(ws.cell(row=HEADER_ROW, column=COL_LINESHEET + 1).value).upper()
        if h_linesheet != "LINESHEET":
            continue  # 跳过不认识的 sheet

        # 加载图片(per-row,按 anchor._from.row 分桶)
        images_by_row = _load_images_by_row(ws)

        for r in range(HEADER_ROW + 1, ws.max_row + 1):
            # 0-idx row(给 images_by_row 用)
            zero_idx = r - 1
            # 1) SKU (Style Number)
            # 2026-08-14:SKU 拼接符统一 `-`(源用 `_` 拼 style+color_code,改 `-`)
            # 例: "245055_3R910" → "245055-3R910"
            sku_raw = normalize_text(ws.cell(row=r, column=COL_STYLE_NUMBER + 1).value)
            sku = sku_raw.replace(".", "-").replace("_", "-")
            if not sku:
                continue

            # 2) Size
            size = normalize_text(ws.cell(row=r, column=COL_SIZE + 1).value)
            if not size:
                continue

            # 3) Units (库存)
            qty_v = ws.cell(row=r, column=COL_UNITS + 1).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            if qty <= 0:
                continue

            # 4) 描述
            desc = normalize_text(ws.cell(row=r, column=COL_STYLE_NAME + 1).value)

            # 5) 颜色 + 色号
            color = normalize_text(ws.cell(row=r, column=COL_COLOR_NAME + 1).value)
            color_code = normalize_text(ws.cell(row=r, column=COL_COLOR_CODE + 1).value)

            # 6) 材质
            material = normalize_text(ws.cell(row=r, column=COL_FABRICATION_NAME + 1).value)

            # 7) 季节 (取第一个)
            season = _pick_first_season(ws.cell(row=r, column=COL_SEASON + 1).value)

            # 8) cost
            cost_v = ws.cell(row=r, column=COL_WHOLESALE + 1).value
            cost = to_float(cost_v)

            # 9) 类目
            category = normalize_text(ws.cell(row=r, column=COL_PRODUCT_TYPE + 1).value)

            # 10) 性别 (从 Linesheet 提取)
            linesheet_raw = ws.cell(row=r, column=COL_LINESHEET + 1).value
            gender = _pick_gender_from_linesheet(linesheet_raw)

            # 11) 图片 (0-idx row = zero_idx):per-row 模式,每行 1 张图
            img_items = images_by_row.get(zero_idx, [])

            rows.append({
                "brand": "",  # 留空,main 走文件名兜底
                "sku": sku,
                "image": "",
                "description": desc,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code,
                "gender": gender,
                "season": season,
                "origin": "ITALY",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,
                "size": size,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": img_items,
                "source_row_height": None,
            })
        break  # 只处理第一个 sheet(EA C-20 只有 1 个 Worksheet)
    return rows
