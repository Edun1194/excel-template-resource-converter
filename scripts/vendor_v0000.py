"""V0000 placeholder vendor —— 多品牌占位通道 (MOORER / GUCCI STOCKLIST / HERNO)。

来源: MooRer R-75 special list.xlsx (MOORER 意大利高端外套品牌, 2026-09-15)
      + Moorer R-75 Special list-update 1809 (2).xlsx (2026-09-19, 列顺序有变化)
      + GUCCI STOCKLIST 220926 (2026-09-23, section 版式)
      + HERNO FW 29-10.xlsx (2026-10-30, 单表头 + 尺码数量矩阵)
仓库 ID 临时用 V0000 (尚未分配正式 vendor code)

**这是个占位通道**: 三家品牌版式完全不同, 按文件特征自动分流 ——
  GUCCI STOCKLIST → 表头含 "ARTICLE_COLOR"
  HERNO          → 表头行含 "SKU" + "BRAND" (实现见 vendor_v0000_herno.py)
  MOORER         → 其余 (中文表头 R1)

结构 (MOORER 单 sheet "FILE", 2026-09-19 版 860 行 61 列 / R-75 版 777 行 59 列):
  R1 header (中文): 性别 | 类目 | 描述 | 颜色 | 图片 | 国际编码 | Size grid
                   | 欧洲零售价 | 供货价 | 26 | 27 | 28 | ... | M | L | XL | UNI | 总数
  注: R-75 版列顺序是 性别 | 类目 | 图片 | 描述 | 颜色 | 国际编码 | ... (图片在 col 3)
      新版 (1809 update) 列顺序是 性别 | 类目 | 描述 | 颜色 | 图片 | 国际编码 | ... (图片在 col 5)

设计:不硬编码列号,按 header 名查找列索引 — 兼容两个版式

R2+: data — col_gender=性别 (Mens/Womens), col_category=类目 (意大利语),
       col_description=描述, col_color=颜色, col_sku=国际编码 (23 字符 `MOUBO100118TEPA663U0237`,
         末尾 5 位是色号), col_size_grid=Size grid (MO/MD/UNI), col_retail=欧洲零售价, col_cost=供货价,
       col_size_labels 后各尺码 qty (尺码 label 在 R1), 最后一列是"总数"不要当尺码

性别翻译: Men/Man/Mens → Man, Women/Woman/Womens → Woman

季节: 参数化 (默认 FW24)

排除 qty=0 行 (跟 framework is_positive_qty 一致)
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


VENDOR_ID = "V0000"


# GUCCI STOCKLIST size grid 默认尺码 (col 12-28 共 17 列)
# 文件 header 不会标具体尺码,按 col 顺序填
# 默认 GUCCI 女鞋 35-51 (连续 17 个数字尺码, 适用于大部分 GUCCI 现货单)
# 如果 size grid 列数不对, 按 col 索引取对应尺码
GUCCI_SIZE_GRID_DEFAULT = [
    "35", "36", "37", "38", "39", "40", "41", "42", "43",
    "44", "45", "46", "47", "48", "49", "50", "51",
]


# 意大利语类目 → 中文/英文 category (简单映射)
CATEGORY_MAP = {
    "BOMBER": "Bomber Jacket",
    "GIACCHE": "Jacket",
    "MAGLIERIA": "Knitwear",
    "SCIARPE": "Scarf",
    "CAMICIE": "Shirt",
    "CAPPELLI": "Hat",
    "GILET": "Vest",
    "PELLICCERIA": "Fur",
    "FELPE": "Sweatshirt",
    "FELPA": "Sweatshirt",
    "TUTE": "Suit",
    "TUTA": "Suit",
    "PIUMINI": "Down Jacket",
    "CINTURE": "Belt",
    "GUANTI": "Gloves",
    "CAPPOTTI": "Coat",
    "PARKA": "Parka",
    "SOPRABITI E IMPERMEABILI": "Outerwear",
    "POLO": "Polo",
    "TSHIRTS": "T-Shirt",
    "T-SHIRTS": "T-Shirt",
    "CARGO": "Cargo Pants",
    "CHINO": "Chino Pants",
    "DENIM": "Denim",
    "JOGGING": "Joggers",
    "PANTALONI": "Pants",
    "5 TASCHE": "5-Pocket Pants",
    "CARDIGAN": "Cardigan",
    "TOP": "Top",
    "SWEAT PANTS": "Sweat Pants",
    "TESSUTO": "Fabric",
    "GONNE": "Skirt",
    "VESTITI": "Dress",
    "BOOTS": "Boots",
    "SNEAKERS": "Sneakers",
    "MOCASSINI": "Loafers",
    "STRINGATE": "Lace-ups",
    "ACCESSORI": "Accessories",
    "BERRETTI": "Beanie",
    "BORSE E ZAINI": "Bags & Backpacks",
    "BORSE": "Bags",
    "ZAINI": "Backpacks",
    "PICCOLA PELLETTERIA": "Small Leather Goods",
    "SWIMWEAR": "Swimwear",
}


# 性别翻译 (兼容 Men/Man/Mens, Women/Woman/Womens)
GENDER_MAP = {
    "MEN": "Man", "MAN": "Man", "MENS": "Man",
    "WOMEN": "Woman", "WOMAN": "Woman", "WOMENS": "Woman",
}


def _find_col(headers: List[str], *candidates: str) -> Optional[int]:
    """在 headers 里找候选列名,返回 1-idx 列号。"""
    norm_headers = [normalize_text(h).lower() for h in headers]
    for cand in candidates:
        cand_norm = cand.lower()
        for i, h in enumerate(norm_headers):
            if h == cand_norm or cand_norm in h:
                return i + 1  # 1-idx
    return None


def parse_rows(
    primary_file: Path,
    _: Path,
    vendor_id: str,
    season: str = "FW24",
) -> List[Dict[str, object]]:
    """读 MOORER xlsx,返回 18 列 row dict 列表。

    货号拼接规则 (2026-09-15 user):
      1) 先抽色号 (SKU 末尾 5 位 `U\\d{4}` / `V\\d{4}` 这种 `字母+4位数字`)
      2) 有色名就拼 `<原SKU>---<color_name>` (参考 V1042 POLO 拼接风格)
      3) 没色号但有 color_name → 拼 `<原SKU>---<color_name>`
      4) 都没 → 原 SKU 不动

    2026-09-23 加 GUCCI STOCKLIST 版式适配:
      - header 含 "ARTICLE_COLOR" → GUCCI STOCKLIST 路径
      - col 6 ARTICLE_COLOR 是 SKU (article + "_" + color_code)
      - col 11 U = 单 size 库存
      - col 9 RETAIL PRICE, col 10 COST PRICE

    2026-10-30 加 HERNO 版式适配 (实现见 vendor_v0000_herno.py):
      - 表头行含 "SKU" + "BRAND" → HERNO 路径
      - 单 sheet 'Worksheet', R2 表头, 尺码矩阵在最后一段(c16-c33)
      - 源文件没有 cost / 供货价 / 色号列 → 那三列留空
    """
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    ws = wb.active
    # GUCCI STOCKLIST 的 header (ARTICLE_COLOR) 不一定在 row 1,可能在中间 — 扫整个文件
    is_gucci = False
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 6).value
        if v and normalize_text(v) == "ARTICLE_COLOR":
            is_gucci = True
            break
    if is_gucci:
        return _parse_gucci_stocklist(ws, vendor_id, season)

    # HERNO: 表头行含 SKU + BRAND (表头固定 R2, 上方可能压着一行残留值)
    if _looks_like_herno(ws):
        from vendor_v0000_herno import parse_rows as _herno_parse
        return _herno_parse(primary_file, None, vendor_id)

    return _parse_moorer(ws, vendor_id, season)


def _looks_like_herno(ws) -> bool:
    """是不是 HERNO 版式(表头行同时含 'SKU' 和 'BRAND')。"""
    try:
        from vendor_v0000_herno import looks_like_herno
        return looks_like_herno(ws)
    except Exception:
        return False


def _parse_gucci_stocklist(ws, vendor_id: str, season: str) -> List[Dict[str, object]]:
    """GUCCI STOCKLIST 版式 (2026-09-23 新增):

    header (新版本): ARTICLE | _ | COLOR | _ | COLOR | ARTICLE_COLOR | DESCRIPTION | _
                    | RETAIL PRICE | COST PRICE | U | [size grid 12-28] | QTY | TOT RETAIL | TOT COST
    header (老版本): 同上但 col 6 = "SKU" 不是 "ARTICLE_COLOR"
    data: 每行一个 SKU, col 6 是完整 SKU
          col 11 (U) 单 size 库存 OR col 12-28 (size grid) 各尺码数量
    图片: 在每个 SKU 行的下一行 (anchor col=0)

    注:
      - header 行可能不在 row 1(文件前面有空行/分页符)
      - 第一个 SKU (row 4) 在第一个 header (row 18) 之前,需要单独扫描文件最前段
      - 2026-09-23 user: 多 size 行展开成多行(每 size 一行)
      - 2026-09-23 user: 货号不拼接,直接用 col 6 原值
    """
    rows: List[Dict[str, object]] = []

    # 1) 收集图: anchor._from.row+1 → SKU 行 (img_row - 1)
    images_by_sku_row: Dict[int, List[Dict[str, object]]] = {}
    try:
        for img in ws._images:
            anchor = img.anchor
            if hasattr(anchor, "_from") and anchor._from:
                img_row = anchor._from.row + 1
                sku_row = img_row - 1
                try:
                    img_bytes = img._data() if hasattr(img, "_data") else None
                except Exception:
                    img_bytes = None
                if img_bytes:
                    fmt = (img.format or "jpeg").lower()
                    if fmt == "jpg":
                        fmt = "jpeg"
                    images_by_sku_row.setdefault(sku_row, []).append({
                        "bytes": img_bytes,
                        "format": fmt,
                    })
    except Exception:
        pass

    # 2) 找所有 header 行 (col 6 == "ARTICLE_COLOR" 或 "SKU")
    header_rows = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 6).value
        if v and normalize_text(v) in ("ARTICLE_COLOR", "SKU"):
            header_rows.append(r)
    if not header_rows:
        return rows

    # 2.5) 每个 section header (R3, R18, ...) 的 col 11-28 是该 section 的尺码标签
    #    例: R3 single-size:  C11='U'
    #        R2988 multi-size: C11=70, C12=75, C13=80, ..., C22=65
    #    每个 section 自己的尺码表,不能用一个全局 GUCCI_SIZE_GRID_DEFAULT
    header_to_size_labels: Dict[int, Dict[int, str]] = {}
    for hr in header_rows:
        labels: Dict[int, str] = {}
        for c in range(11, 29):
            v = ws.cell(hr, c).value
            if v is None or v == "":
                continue
            # size label 是数字 (75) 或字符串 ('U')
            label = str(v).strip()
            if label:
                labels[c] = label
        if labels:
            header_to_size_labels[hr] = labels

    # 3) 段扫描: [file_start, first_header) + [header_r, next_header_r)
    #    2026-09-23 user: 第一个 SKU 在 row 4 (row 18 之前) 之前缺失,加 file_start 分段
    segments = []
    if header_rows[0] > 1:
        segments.append((1, header_rows[0] - 1))
    for i, hr in enumerate(header_rows):
        next_hr = header_rows[i + 1] if i + 1 < len(header_rows) else ws.max_row + 1
        segments.append((hr, next_hr - 1))   # include header_r itself so we can look it up

    for start_r, end_r in segments:
        if end_r < start_r:
            continue
        # 这个段的 header 行 (= start_r)
        hr = start_r
        size_labels = header_to_size_labels.get(hr, {})

        for r in range(start_r + 1, end_r + 1):   # skip header itself
            sku = normalize_text(ws.cell(r, 6).value)
            if not sku:
                continue
            article = normalize_text(ws.cell(r, 1).value)
            color_desc = normalize_text(ws.cell(r, 3).value)
            color_code = ws.cell(r, 5).value
            description = normalize_text(ws.cell(r, 7).value)
            retail = to_float(ws.cell(r, 9).value)
            cost = to_float(ws.cell(r, 10).value)
            color_code_str = str(color_code) if color_code else ""

            # 2026-09-23 user: 货号不拼接,直接用 col 6 原 SKU
            full_sku = sku

            base_row = {
                "brand": "GUCCI",
                "sku": full_sku,
                "image": "",
                "description": description,
                "category": "",
                "material": "",
                "color": color_desc,
                "color_code": color_code_str,
                "season": season,
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "vendor_id": vendor_id,
                "gender": "",
                "image_items": images_by_sku_row.get(r, []),
            }

            # 用本 section header 定义的尺码表扫 col 11-28
            size_grid_qty = []
            for c, size_label in size_labels.items():
                v = ws.cell(r, c).value
                if is_positive_qty(v):
                    size_grid_qty.append((c, size_label, int(v)))

            if size_grid_qty:
                # 每个非零尺码一行
                for _c, size_label, qty_v in size_grid_qty:
                    rows.append({**base_row,
                                "size": size_label,
                                "quantity": qty_v,
                                })
            else:
                # 全 0 → 跳过 (跟 framework is_positive_qty 一致)
                continue
    return rows


def _parse_moorer(ws, vendor_id: str, season: str) -> List[Dict[str, object]]:
    import re

    COLOR_CODE_RE = re.compile(r"[A-Z]\d{4}$")  # 末尾 5 位字母+4位数字

    rows: List[Dict[str, object]] = []

    # 1) 按 header 名找各列 (兼容 R-75 / 1809-update 两个版式)
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_gender = _find_col(headers, "性别", "gender", "sex")
    col_category = _find_col(headers, "类目", "category")
    col_description = _find_col(headers, "描述", "description")
    col_color = _find_col(headers, "颜色", "color", "colour")
    col_sku = _find_col(headers, "国际编码", "sku", "international code", "material")
    col_size_grid = _find_col(headers, "size grid", "sizegrid")
    col_retail = _find_col(headers, "欧洲零售价", "retail", "retail price")
    col_cost = _find_col(headers, "供货价", "cost", "cost price")

    if not col_sku:
        raise ValueError(f"未找到国际编码列 (header row: {headers})")

    # 2) 尺码列: R1 header 中数字/字母尺码 (排除 metadata 列)
    size_cols: List[int] = []
    size_labels: List[str] = []
    SKIP_HEADERS = {"总数", "total", "tot", "合计", "sum", "图片", "image", "picture"}
    last_data_col = col_cost  # 尺码从 cost 列之后开始 (R-75 版)
    for c in range(last_data_col + 1, ws.max_column + 1):
        hdr = normalize_text(ws.cell(1, c).value)
        if not hdr:
            continue
        if hdr.lower() in SKIP_HEADERS:
            continue
        size_labels.append(hdr)
        size_cols.append(c)

    # 3) 图片: 从 ws._images 拿(不读 col '图片' 文本字段)
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    try:
        for img in ws._images:
            anchor = img.anchor
            if hasattr(anchor, "_from") and anchor._from:
                try:
                    img_bytes = img._data() if hasattr(img, "_data") else None
                except Exception:
                    img_bytes = None
                if img_bytes:
                    images_by_row.setdefault(anchor._from.row + 1, []).append({
                        "bytes": img_bytes,
                        "format": img.format or "jpeg",
                    })
    except Exception:
        pass

    # 4) 扫 R2.. 数据行
    for r in range(2, ws.max_row + 1):
        sku_raw = normalize_text(ws.cell(r, col_sku).value)
        if not sku_raw:
            continue

        gender_raw = normalize_text(
            ws.cell(r, col_gender).value if col_gender else ""
        ).upper()
        gender = GENDER_MAP.get(gender_raw, gender_raw.title())

        category_raw = (
            normalize_text(ws.cell(r, col_category).value) if col_category else ""
        )
        category = CATEGORY_MAP.get(category_raw.upper(), category_raw)

        description = (
            normalize_text(ws.cell(r, col_description).value) if col_description else ""
        )
        color_name = (
            normalize_text(ws.cell(r, col_color).value) if col_color else ""
        )

        # 抽色号: SKU 末尾 5 位字母+数字
        m = COLOR_CODE_RE.search(sku_raw)
        color_code = m.group() if m else ""
        # 拼接 SKU: 优先 color_name, 其次 color_code, 都没原样保留
        if color_name:
            sku = f"{sku_raw}---{color_name}"
        elif color_code:
            sku = f"{sku_raw}---{color_code}"
        else:
            sku = sku_raw

        cost = to_float(ws.cell(r, col_cost).value) if col_cost else None
        retail = to_float(ws.cell(r, col_retail).value) if col_retail else None

        image_items = images_by_row.get(r, [])

        # 5) 展开各尺码 qty
        for c_idx, size_label in enumerate(size_labels):
            if not size_label:
                continue
            qty_v = ws.cell(r, size_cols[c_idx]).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            rows.append(
                {
                    "brand": "MOORER",
                    "sku": sku,
                    "image": "",
                    "description": description,
                    "category": category,
                    "material": "",
                    "color": color_name,
                    "color_code": color_code,
                    "season": season,
                    "origin": "",
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": cost,
                    "retail_price": retail,
                    "size": size_label,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "gender": gender,
                    "image_items": image_items,
                }
            )
    return rows