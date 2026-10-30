"""VC002 CANADA GOOSE WOMAN - FW26 vendor adapter (2026-08-07 新建)

适配文件:CANADA GOOSE 加拿大鹅女装 FW26 现货单 (.xlsx, GRIGLIA sheet)。

跟 VC001 (CANADA GOOSE FW26 ASA) 完全不同的版式:
  - VC001:横向尺码 + 每产品 2 行(产品行+库存行),无内嵌图
  - VC002 (本):**横排尺码 + per-row 模式 + 42 张内嵌图**

文件结构(GRIGLIA sheet,18 列):
  R1 header (A 组尺码):
    c1=X (标志) | c2=PHOTOS (图列) | c3=NOTES
    c4=Style # (完整货号,含色号) | c5=Style Number (款号) | c6=Style Name
    c7=Color Code (色号) | c8=Color (色名)
    c9=WHLS EUR (cost) | c10=CLIENTE (retail)
    c11='A' (尺码组 A 标记) | c12=XS | c13=S | c14=M | c15=L | c16=XL
    c17=TOTAL PCS | c18=TOTAL WHS

  R2 子尺码组 (B 组):
    c11='B' | c12='OS' (One Size 单码组)

  R3 空白(可能 R1 header 副行 / 图片起始行)

  R4+ 数据:
    每行 1 个产品 + 1 个 size 标记 (A/B)
    A 组:c12-c16 数字 = XS/S/M/L/XL 各尺码 qty
    B 组:c12 数字 = OS 单码 qty(其他 c13-c16 全空)

字段映射:
  SKU        ← c4 Style # (e.g. "1487WT.9846")
  描述       ← c6 Style Name (e.g. "WANDER CREW WITH TONAL DI")
  颜色       ← c8 Color (e.g. "Bloom Pink - Rose bourgeo")
  色号       ← c7 Color Code (e.g. "9846")
  季节       ← "" 留空(main 走文件名兜底 → "FW26")
  cost       ← c9 WHLS EUR
  retail     ← c10 CLIENTE
  品牌       ← "" 留空(main 走文件名兜底 → "CANADA GOOSE")
  性别       ← "Woman" (从文件名"WOMAN")
  仓库ID     ← VENDOR_ID = "VC002"

图片:openpyxl ws._images 按 anchor._from.row (0-idx) 分桶,
     42 张图,R3-R49 范围(实测 R4-R48 产品行 1:1 配图)。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 VC002(2026-08-07 用户说"没有工厂代码",暂定)
VENDOR_ID = "VC002"

# R1 是 A 组 header (1-idx)
HEADER_ROW_A = 1
# R2 是 B 组子 header (1-idx)
HEADER_ROW_B = 2


def get_vendor_meta() -> Dict[str, object]:
    """VC002 vendor 元数据。"""
    return {
        "default_image_col_width": 40.0,
    }


# 0-idx 关键列位置(基于 18 列版式)
COL_STYLE_FULL = 3    # c4 Style # (完整货号含色号)
COL_STYLE_NUM = 4     # c5 Style Number (款号)
COL_STYLE_NAME = 5    # c6 Style Name
COL_COLOR_CODE = 6    # c7 Color Code
COL_COLOR = 7         # c8 Color
COL_COST = 8          # c9 WHLS EUR
COL_RETAIL = 9        # c10 CLIENTE
COL_SIZE_GROUP = 10   # c11 尺码组标记 (A/B)
COL_SIZE_START = 11   # c12 起尺码 label 列 (0-idx 11+)
COL_TOTAL_PCS = 16    # c17 TOTAL PCS
COL_TOTAL_WHS = 17    # c18 TOTAL WHS


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 openpyxl 0-idx row 分桶图(每行 1 张)。"""
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


def _build_size_labels(ws) -> Tuple[Dict[int, str], Dict[int, str]]:
    """从 R1 + R2 双行 header 提取尺码 label,分 A/B 组(2026-08-07)。

    A 组:R1 c12-c16 (XS/S/M/L/XL)
    B 组:R2 c12-c16 (OS 等单码)
    c17/c18 (0-idx 16/17) 是 TOTAL PCS / TOTAL WHS,不算尺码。

    ⚠️ R1 和 R2 用的是同一组 col (0-idx 11-15),但值不同 — R2 是 B 组覆盖
    (c12 R1='XS',R2='OS' 冲突)。所以必须分开两个字典,数据行按 size_group 选择。

    返回: (a_size_labels, b_size_labels)
    """
    a_labels: Dict[int, str] = {}
    b_labels: Dict[int, str] = {}

    # 尺码 col 范围:0-idx [COL_SIZE_START, COL_TOTAL_PCS) = [11, 16) = c12-c16
    size_col_range = range(COL_SIZE_START, COL_TOTAL_PCS)

    # A 组:R1 header
    for c in size_col_range:
        v = ws.cell(row=HEADER_ROW_A, column=c + 1).value
        if v is None:
            continue
        sv = str(v).strip()
        if sv:
            a_labels[c] = sv

    # B 组:R2 header(可能只有 1 个尺码,其他空)
    for c in size_col_range:
        v = ws.cell(row=HEADER_ROW_B, column=c + 1).value
        if v is None:
            continue
        sv = str(v).strip()
        if sv:
            b_labels[c] = sv

    return a_labels, b_labels


def _pick_gender_from_filename(primary_file: Path) -> str:
    """从文件名猜性别(默认 Woman,MAN 出现则 Man)。"""
    name = primary_file.stem.upper()
    if " MAN" in f" {name} " or name.endswith(" MAN") or "-MAN" in name:
        return "Man"
    return "Woman"


def _parse_sheet(ws, vendor_id: str, gender: str, primary_file: Path) -> List[Dict[str, object]]:
    """解析 GRIGLIA sheet,返回 row dict 列表。"""
    rows: List[Dict[str, object]] = []
    a_size_labels, b_size_labels = _build_size_labels(ws)
    if not a_size_labels and not b_size_labels:
        return rows

    images_by_row = _load_images_by_row(ws)
    # 季节从文件名提取("FW26" / "SS26" 等)
    season = ""
    name = primary_file.stem.upper()
    m_season = re.search(r"\b(FW|SS|FA|SP|HOL|RESORT)\s*(\d{2})\b", name)
    if m_season:
        season = f"{m_season.group(1)}{m_season.group(2)}"
    else:
        season = ""  # 留空,main 走文件名兜底

    for r in range(HEADER_ROW_A + 1, ws.max_row + 1):
        # 1) 识别产品行: c4 (0-idx 3) 非空
        style_full = normalize_text(ws.cell(row=r, column=COL_STYLE_FULL + 1).value)
        if not style_full:
            continue

        # 2) 描述 / 颜色 / 色号
        style_name = normalize_text(ws.cell(row=r, column=COL_STYLE_NAME + 1).value)
        colour_code = normalize_text(ws.cell(row=r, column=COL_COLOR_CODE + 1).value)
        colour = normalize_text(ws.cell(row=r, column=COL_COLOR + 1).value)

        # 3) cost / retail
        cost_v = ws.cell(row=r, column=COL_COST + 1).value
        cost = to_float(cost_v)
        retail_v = ws.cell(row=r, column=COL_RETAIL + 1).value
        retail = to_float(retail_v)

        # 4) 尺码组判断(A/B)
        size_group = ws.cell(row=r, column=COL_SIZE_GROUP + 1).value
        size_group_norm = str(size_group).strip().upper() if size_group else ""

        # 5) 图片 (0-idx row = r - 1)
        zero_idx = r - 1
        row_images = images_by_row.get(zero_idx, [])

        # 6) 按尺码展开
        if size_group_norm == "B":
            size_labels = b_size_labels
        else:
            # A 组 或 未标记,默认用 A 组
            size_labels = a_size_labels

        # 只在尺码 col 范围内展开 (0-idx [COL_SIZE_START, COL_TOTAL_PCS) = c12-c16)
        for c in range(COL_SIZE_START, COL_TOTAL_PCS):
            if c not in size_labels:
                continue
            size_label = size_labels[c]
            qty_v = ws.cell(row=r, column=c + 1).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            if qty <= 0:
                continue
            rows.append({
                "brand": "",  # 留空,main 走文件名兜底 → "CANADA GOOSE"
                "sku": style_full,  # 完整货号 (含色号),不用拼接
                "image": "",
                "description": style_name,
                "category": "",  # 源没类目字段
                "material": "",
                "color": colour,
                "color_code": colour_code,
                "gender": gender,
                "season": season,
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": None,
            })

    return rows


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 CANADA GOOSE FW26 WOMAN .xlsx,返回 row dict 列表。

    仓库 ID 统一用 VENDOR_ID = "VC002" (2026-08-07 用户要求),忽略传入 vendor_id。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    vendor_id = VENDOR_ID

    gender = _pick_gender_from_filename(primary_file)

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    for ws in wb.worksheets:
        # 防御性检查:R1 c4 (0-idx 3) 应是 "Style #" 才是 CANADA GOOSE GRIGLIA 版式
        h_marker = normalize_text(ws.cell(row=HEADER_ROW_A, column=COL_STYLE_FULL + 1).value)
        if h_marker.upper() != "STYLE #":
            continue  # 跳过不认识的 sheet
        rows.extend(_parse_sheet(ws, vendor_id, gender, primary_file))
        break  # 只处理第一个 sheet
    return rows
