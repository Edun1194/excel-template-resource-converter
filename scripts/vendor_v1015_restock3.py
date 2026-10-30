"""V1015-补货单3 GIORGIO ARMANI C-20 vendor adapter (2026-08-03 新建)

适配文件:Giorgio Armani 工厂的库存清单 (GIRGIO C-20),来自 SharePoint 同步。
跟 V1015-补货单 (PRADA 现货单) / V1015-补货单2 (MIUMIU ATS) 完全不同:

  - V1015-补货单 (PRADA): 1 sheet,header R5,per-row size+qty (11 列)
  - V1015-补货单2 (MIUMIU ATS): 3 sheet,header R6,per-row size+qty (7/22 列)
  - V1015-补货单3 (本):  4 sheet,header R4,**横排尺码** (c15+)  →  按列扫描

结构(基于 1-idx):
  R4 = header
  c1 = Sector                 ← WOMAN / MAN
  c2 = Category               ← READY TO WEAR / SHOES / LEATHER GOODS / SOFT ACCESSORIES / BIJOUX
  c3 = Product                ← BLAZER / BOOT / BALLET FLAT / 5 POCKETS PANT
  c4 = Style Code             ← e.g. "3WHGG11D" / "GW000142" / "X1D346"
  c5 = Fabric Code
  c6 = Color Code             ← e.g. "PZ01" / "UC001" / "00002"
  c7 = Color                  ← e.g. "Printed" / "NIGHT" / "NERO"
  c8 = Article Composition    ← "FABRIC 1: 100% VIRGIN WOOL - ..."
  c9 = Style-Fabric Photo     ← 空 cell
  c10= Style-Fabric Sketch    ← 空 cell
  c11= Fabric Color Photo     ← 空 cell
  c12= Active Year Season Code ← e.g. "202560" (6 位数字, 季节编码)
  c13= Storage Location Code  ← 仓库, e.g. "RMB1" / "GA1R" / "SL1S" / "XXOP"
  c14= Seasonality Flag
  c15= Amount Eur001           ← 单件 cost
  c16+ = 尺码列 (每个 sheet 不同: 衣服 XS-S-M-L-..., 鞋 35-36-37-...)

4 个 sheet(都转):
  - GIACENZA GAD RTW: 女装成衣
  - GIACENZA GAD ACC: 女装配饰
  - GIACENZA GAU RTW: 男装成衣
  - GIACENZA GAU ACC: 男装配饰

字段映射:
  SKU        ← Style Code (c4) + "-" + Color Code (c6) (用户期望 `<数字>-<颜色>` 风格)
  描述       ← c3 Product
  颜色       ← c7 Color
  色号       ← c6 Color Code
  季节       ← c12 Active Year Season Code
  cost       ← c15 Amount Eur001
  仓库       ← c13 Storage Location Code (用来推 warehouse 名字)
  类目       ← c2 Category (READY TO WEAR / SHOES / ...)
  性别       ← c1 Sector: WOMAN→Donna / MAN→Uomo
  品牌       ← ""  (留空,main 走文件名兜底 → "GIORGIO ARMANI")
  仓库ID     ← VENDOR_ID = "V1015" (工厂 = V1015, 补货单 3 是子类型)

图片:openpyxl ws._images 按 anchor._from.row (0-idx) 分桶。
     每个产品行有 2-3 张图 (Style-Fabric / Sketch / Fabric-Color),
     KEEP_IMAGES_PER_ROW=1 默认只取第一张。

每行展开:每行产品展开为 (size, qty) pairs — c16+ 尺码列 > 0 才生成 1 行。

性能优化(2026-08-03):
  4 sheet 用 ThreadPoolExecutor 并行处理(openpyxl lxml 后端会释放 GIL,
  lxml XML 解析 + Pillow JPEG 解码都是 native code,Python GIL 释放)。
  4 sheet 顺序处理 272MB GIRGIO 要 4 分钟,4 sheet 并行后实测降到 ~1 分钟。
"""

from __future__ import annotations

import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 V1015 (2026-07-31 用户要求),即使 vendor code 是 V1015-补货单3
VENDOR_ID = "V1015"

# R4 是 header (1-idx)
HEADER_ROW = 4


def get_vendor_meta() -> Dict[str, object]:
    """V1015-补货单3 vendor 元数据。"""
    return {
        "default_image_col_width": 40.0,
    }


# 0-idx 关键列位置(基于 GIRGIO C-20 共享版式)
COL_SECTOR = 0          # c1 Sector (WOMAN/MAN)
COL_CATEGORY = 1        # c2 Category
COL_PRODUCT = 2         # c3 Product (描述)
COL_STYLE_CODE = 3      # c4 Style Code
COL_COLOR_CODE = 5      # c6 Color Code
COL_COLOR = 6           # c7 Color
COL_COMPOSITION = 7     # c8 Article Composition (材质)
COL_SEASON = 11         # c12 Active Year Season Code
COL_STORAGE = 12        # c13 Storage Location Code
COL_COST = 14           # c15 Amount Eur001
COL_SIZE_START = 15     # c16 起尺码列 (0-idx 15+)


# 性别映射(Sector 文本 → 意大利/英语标签)
_SECTOR_TO_GENDER = {
    "WOMAN": "Donna",
    "MAN": "Uomo",
    "UNISEX": "Unisex",
    "BAMBINO": "Bambino",
    "BAMBINA": "Bambina",
}


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 openpyxl 0-idx row 分桶图(每行 1 张,KEEP_IMAGES_PER_ROW)。"""
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


def _build_size_labels(ws) -> Dict[int, str]:
    """从 header 行 (R4 1-idx) c16+ 起提取尺码 label,跳过 None/空 cell。

    label 归一化:
      - '   34'  → '34'  (strip 空格)
      - ' 35,5'  → '35.5'  (欧洲逗号 → 句点)
      - '35/36'  → '35/36'  (保留斜杠双码)
      - '5M'     → '5M'    (童鞋双码)
    """
    out: Dict[int, str] = {}
    for c in range(COL_SIZE_START, ws.max_column):
        v = ws.cell(row=HEADER_ROW, column=c + 1).value
        if v is None:
            if out:
                break  # 已有 token 后面空 → 本行结束
            continue
        sv = str(v).strip()
        if not sv:
            if out:
                break
            continue
        # 纯数字(整数或小数)做欧码归一化
        if re.fullmatch(r"\d+(?:[.,]\d+)?", sv):
            sv = sv.replace(",", ".")
        out[c] = sv
    return out


def _parse_sheet(ws, sheet_name: str, vendor_id: str) -> List[Dict[str, object]]:
    """解析单个 sheet,返回 row dict 列表。

    横向尺码模式:R4 header,R5+ 数据行,c15+ 尺码列(每列一个尺码)。
    """
    rows: List[Dict[str, object]] = []
    size_labels = _build_size_labels(ws)
    if not size_labels:
        return rows

    images_by_row = _load_images_by_row(ws)

    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        zero_idx = r - 1

        # 1) Style Code (SKU 前缀)
        style_code = normalize_text(ws.cell(row=r, column=COL_STYLE_CODE + 1).value)
        if not style_code:
            continue
        # 2) Color Code (SKU 后缀)
        color_code = normalize_text(ws.cell(row=r, column=COL_COLOR_CODE + 1).value)
        # 用户期望 `<数字>-<颜色>` 风格 SKU, e.g. "3WHGG11D-PZ01"
        if color_code:
            sku = f"{style_code}-{color_code}"
        else:
            sku = style_code

        # 3) 描述 (Product)
        product = normalize_text(ws.cell(row=r, column=COL_PRODUCT + 1).value)
        desc = product  # c3 就是描述,没有更详细的描述列

        # 4) 颜色
        color = normalize_text(ws.cell(row=r, column=COL_COLOR + 1).value)

        # 5) 材质
        material = normalize_text(ws.cell(row=r, column=COL_COMPOSITION + 1).value)

        # 6) 季节
        season = normalize_text(ws.cell(row=r, column=COL_SEASON + 1).value)

        # 7) 仓库 (Storage Location Code)
        storage = normalize_text(ws.cell(row=r, column=COL_STORAGE + 1).value)

        # 8) cost
        cost_v = ws.cell(row=r, column=COL_COST + 1).value
        cost = to_float(cost_v)

        # 9) 性别 (Sector → Donna/Uomo)
        sector = normalize_text(ws.cell(row=r, column=COL_SECTOR + 1).value).upper()
        gender = _SECTOR_TO_GENDER.get(sector, sector.title() if sector else "")

        # 10) 类目
        category = normalize_text(ws.cell(row=r, column=COL_CATEGORY + 1).value)

        # 11) 图片 (0-idx row = zero_idx)
        row_images = images_by_row.get(zero_idx, [])

        # 12) 按尺码展开(横排:每列 1 个尺码,值 > 0 → 1 行)
        size_count = 0
        for c, size_label in size_labels.items():
            qty_v = ws.cell(row=r, column=c + 1).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            if qty <= 0:
                continue
            size_count += 1
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
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "warehouse": storage,  # Storage Location Code (e.g. RMB1/SL1S/XXOP/GA1R) — 给前端展示
                "image_items": row_images,
                "source_row_height": None,
            })

    return rows


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 GIRGIO C-20 .xlsx,4 sheet 全部转,合并返回 row dict 列表。

    仓库 ID 统一用 VENDOR_ID = "V1015" (2026-07-31),忽略传入 vendor_id。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    # 用模块级 VENDOR_ID 覆盖 vendor_id,保证仓库ID 列稳定是 "V1015"
    vendor_id = VENDOR_ID

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    # 1) 先筛选要处理的 sheet(R4 c1 == 'Sector' 才是 GIRGIO 版式)
    valid_sheets = []
    for ws in wb.worksheets:
        h0 = normalize_text(ws.cell(row=HEADER_ROW, column=1).value).upper()
        if h0 == "SECTOR":
            valid_sheets.append(ws)
    if not valid_sheets:
        return rows

    # 2) 多 sheet 并行处理(2026-08-03 性能优化)
    # 4 sheet 顺序处理 272MB GIRGIO 要 4 分钟,并行后 ~1 分钟。
    # openpyxl lxml 后端 XML 解析 + Pillow JPEG 解码都是 native code,
    # 会释放 GIL,ThreadPoolExecutor 能有效并行。
    # max_workers 限制为 min(sheet 数, 2),留点内存余量(7.3G 总内存,
    # 4 thread 4 sheet 6000+ 图同时吃内存可能 OOM;2 thread 安全又快 1.7 倍)。
    max_workers = min(len(valid_sheets), 2)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = [
            pool.submit(_parse_sheet, ws, ws.title, vendor_id)
            for ws in valid_sheets
        ]
        for f in futures:
            rows.extend(f.result())
    return rows
