"""V1004 vendor adapter — 工厂 V1004 的多品牌适配器(独立运行)。

来源工厂 V1004 发出的文件,**版式完全一致**(10-row/product):
- R+0: SKU (col 13)
- R+2: 描述 (col 13,内嵌欧洲十进制价格)
- R+4: 尺码标签 (cols 12-24,A/XXS/.../3XS 或 UK 00/02/...)
- R+6: qty (col 15+) + cost (col 36) + total (col 39)
- 段头在 col 4-5 (Donna/Uomo/Abbigliamento/Accessori)

SKU 命名因品牌而异(24 字符 / 17 字符 / 8 字符 都可能),
但因为版式**结构性强**,根本不需要 _is_sku() 内容判断 —

**结构判断的两条铁律**:
1. R+2 col 13 必须含欧洲价格(`\d+[.,]\d{2}`) → 这是描述行
2. 从 R+2 往后搜能找到一个尺码+qty 行 → 这是尺码行

任何"长得像内容"的行(描述、尺码、段头)都不可能同时通过这两条,
所以**不需要 _is_sku 这种内容判断函数**。

复用策略(显式 import,不是 monkey-patch):
- `vendor_v1021._parse_description`  → 描述行内嵌欧洲价格解析
- `vendor_v1021._find_size_and_qty` → 尺码行 + qty 行搜索(灵活扫 cols 8-24)
- `vendor_v1021._load_images_by_row` → 图片按行索引
- `vendor_v1021._find_image_for_block` → 图片归属(±2 行匹配)
- `vendor_v1021._extract_color_from_sku` / `_strip_color_from_sku` → POLO 风格色卡码拆分
- `vendor_v1021._find_category_for_row` → 段头类目查找

V1004 自己做(独立运行):
- `parse_rows()`: 扫行 + 结构判断 → 找产品块
- 没有任何 monkey-patch V1021 全局变量,两个 adapter 完全独立

特殊点:
- 仓库 ID 固定为 ``V1004``
- 不设 default_brand,让 dispatcher 走"文件名兜底"分支
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

import vendor_v1021
from convert_vendor_quote_shared import to_float


VENDOR_ID = "V1004"

# 欧洲十进制价格正则(描述行 R+2 必须匹配)
_EURO_PRICE_RE = re.compile(r"\d+[.,]\d{2}")

# 段头跳过词(不当作产品行,免得把 "Donna"/"Uomo"/"Abbigliamento" 当 SKU)
_HEADER_WORDS = frozenset({
    "DONNA", "UOMO", "ABBIGLIAMENTO", "ACCESSORI", "CALZATURE",
    "MAGLIERIA", "CAMICIE", "TOTALE", "TOTALI", "PAIA", "STOCKLIST",
})

# 性别段头(用作"段位继承"的边界)
_GENDER_HEADERS = frozenset({"DONNA", "UOMO", "UOMINI", "WOMAN", "MAN"})

# 段头类目词(2026-08-04 ARC'TERYX 段位继承)
# 段头行 col 3/4 出现这些词,代表从这一行开始到下一个性别段头(Donna/Uomo)之前,所有产品都属该类目
# ARC'TERYX 用的:Abbigliamento/Accessori/Scarpe/Borse(意大利语),也兼容英语版
_CATEGORY_HEADERS = frozenset({
    "ABBIGLIAMENTO", "ACCESSORI", "SCARPE", "BORSE", "MAGLIERIA",
    "CALZATURE", "CAMICIE", "PANTALONI", "ZAINI", "FELPE", "CAPPELLI",
    "SNEAKER", "GIACCHE", "CAPPOTTI", "JACKET", "PANTS", "BAGS", "SHOES",
    "BELT", "HAT", "FOULARD", "SCIARPE", "LIFESTYLE", "SHORTS", "T-SHIRT",
})


def _find_category_v1004(df: pd.DataFrame, sku_row: int) -> str:
    """V1004 专属类目查找 — 段位继承(2026-08-04)。

    表格里"Donna"/"Uomo" 等性别段头下面,有"Abbigliamento"/"Accessori"/"Scarpe"/"Borse"
    等类目段头,这些类目覆盖到下一个性别段头为止。
    中间所有产品行都继承最近一个类目段头。

    V1021._find_category_for_row 只搜 30 行,远距离段头(ARC'TERYX R8 类目盖到 R580)
    搜不到。本函数:
      1) 往上找最近的性别段头(Donna/Uomo),作为搜索上界
      2) 在该段头之后到 sku_row 之间,找最近的"类目词"行
      3) 找不到回退到 30 行(短文件兜底)
    """
    if 12 >= df.shape[1]:
        return ""

    # 1) 找上一个性别段头(Donna/Uomo)行
    prev_gender_row = -1
    for r in range(sku_row - 1, max(-1, sku_row - 5000), -1):
        for c in (3, 4):
            if c >= df.shape[1]:
                continue
            v = df.iloc[r, c]
            if isinstance(v, str) and v.strip().upper() in _GENDER_HEADERS:
                prev_gender_row = r
                break
        if prev_gender_row >= 0:
            break

    # 2) 搜索范围:从性别段头后一行 到 sku_row(也覆盖没有性别段头的短文件)
    if prev_gender_row >= 0:
        search_start = prev_gender_row + 1
    else:
        search_start = max(0, sku_row - 30)

    # 3) 在搜索范围内找最近的类目词行
    candidates = []
    for r in range(search_start, sku_row):
        for c in (3, 4, 5):
            if c >= df.shape[1]:
                continue
            v = df.iloc[r, c]
            if isinstance(v, str):
                s = v.strip()
                if s and s.upper() in _CATEGORY_HEADERS:
                    # 距离越小越近,sort 后取最近
                    candidates.append((sku_row - r, s))
    if not candidates:
        return ""
    candidates.sort()
    return candidates[0][1]


def _find_product_blocks_v1004(df: pd.DataFrame) -> List[Dict]:
    """扫整张表,找所有产品块(结构判断,不看 SKU 内容)。

    结构判断:
    1) R+2 col 13 (0-idx col 12) 必须含欧洲价格 → 描述行
    2) R+2 之后能搜到尺码+qty(用 V1021 的灵活扫描)→ 尺码行存在

    只要这两条满足,R 就是 SKU 行。
    """
    blocks: List[Dict] = []
    max_col = df.shape[1]
    if 12 >= max_col:
        return blocks

    for r in range(len(df)):
        # 1) 排除段头行(直接 skip 整行)
        for c in (3, 4):  # 0-idx col 3/4 = 1-idx col 4/5
            if c >= max_col:
                continue
            v = df.iat[r, c]
            if isinstance(v, str) and v.strip().upper() in _HEADER_WORDS:
                break
        else:
            # 2) R+2 col 13 必须含欧洲价格
            if r + 2 >= len(df):
                continue
            desc_val = df.iat[r + 2, 12]
            if not isinstance(desc_val, str) or not _EURO_PRICE_RE.search(desc_val):
                continue

            # 3) R+2 之后能搜到尺码 + qty(用 V1021 的灵活扫描)
            size_map, qty_map = vendor_v1021._find_size_and_qty(df, r + 2)
            if not size_map or not qty_map:
                continue

            # 4) R col 13 = SKU(任何非空字符串,不再做 _is_sku 内容判断)
            sku_val = df.iat[r, 12]
            if not isinstance(sku_val, str) or not sku_val.strip():
                continue
            sku = sku_val.strip()
            # 兜底:把段头词当 SKU 排除(虽然 col 4/5 段头已排除,但 col 13 防御性检查)
            if sku.upper() in _HEADER_WORDS:
                continue

            blocks.append({
                "sku_row": r,
                "desc_row": r + 2,
                "sku": sku,
                "desc_text": desc_val.strip(),
                "size_map": size_map,
                "qty_map": qty_map,
            })
    return blocks


def _parse_v1004_ats(
    df: "pd.DataFrame",
    images_by_row: Dict[int, List[Dict[str, object]]],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """C.P. Company / NuORDER ATS 1-row/size 版式 (2026-10-13 接入).

    文件结构 (.xlsx, 1 sheet 'Worksheet'):
      R1: 空
      R2: 品牌 (e.g. 'C.P. Company')
      R3: 空
      R4: header (c1=Image / c2=Style Number / c3=Style Name / c4=Fabrication Name /
                c5=Fabrication Code / c6=Color / c7=Color Code / c8=Size /
                c9=Units / c10=Date Available / c11=Warehouse Name / c12=Currency /
                c13=Wholesale Price / c14=Extended Wholesale / c15=Retail Price /
                c16=Extended Retail / c17=Linesheet / c18=Season Name / c19=Collection)
      R5+: data (每行一个 size, 1-row/size)

    字段映射 (1-idx):
      c2=Style Number (SKU) / c3=Style Name (desc) / c6=Color / c7=Color Code /
      c8=Size / c9=Units (qty) / c13=Wholesale (cost) / c15=Retail /
      c17=Linesheet (gender+collection) / c18=Season Name /
      c19=Collection

    类目: c17 Linesheet (e.g. 'FW26 RE-ORDERS MAN') 或 c19 Collection (e.g. 'PRE COLLECTION')
    性别: c17 末段 (MAN / WOMAN / UNISEX 等)
    """
    rows: List[Dict[str, object]] = []

    # 0-idx row 1 = 1-idx R2 是品牌
    brand_raw = df.iloc[1, 0] if df.shape[0] > 1 else ""
    brand = vendor_v1021._normalize(str(brand_raw)) if brand_raw else ""

    for r in range(4, df.shape[0]):  # 0-idx row 4 = 1-idx R5
        style_number = df.iloc[r, 1] if df.shape[1] > 1 else None
        if not style_number or not isinstance(style_number, str):
            continue
        sn = vendor_v1021._normalize(style_number)
        if not sn:
            continue

        style_name = df.iloc[r, 2] if df.shape[1] > 2 else None
        description = vendor_v1021._normalize(str(style_name)) if style_name else ""

        color = df.iloc[r, 5] if df.shape[1] > 5 else None
        color = vendor_v1021._normalize(str(color)) if color else ""

        color_code = df.iloc[r, 6] if df.shape[1] > 6 else None
        color_code = vendor_v1021._normalize(str(color_code)) if color_code else ""

        size = df.iloc[r, 7] if df.shape[1] > 7 else None
        size = vendor_v1021._normalize(str(size)) if size else ""
        if not size:
            continue

        units_v = df.iloc[r, 8] if df.shape[1] > 8 else None
        try:
            qty = int(round(float(units_v or 0)))
        except (TypeError, ValueError):
            qty = 0
        if qty <= 0:
            continue

        wholesale = df.iloc[r, 12] if df.shape[1] > 12 else None
        wholesale = to_float(wholesale)

        retail = df.iloc[r, 14] if df.shape[1] > 14 else None
        retail = to_float(retail)

        linesheet = df.iloc[r, 16] if df.shape[1] > 16 else None
        linesheet = vendor_v1021._normalize(str(linesheet)) if linesheet else ""

        season_name = df.iloc[r, 17] if df.shape[1] > 17 else None
        season_name = vendor_v1021._normalize(str(season_name)) if season_name else ""

        collection = df.iloc[r, 18] if df.shape[1] > 18 else None
        # NaN / None / 非字符串 都视为空
        if collection is None or (isinstance(collection, float) and pd.isna(collection)):
            collection = ""
        else:
            collection = vendor_v1021._normalize(str(collection))

        # 货号: c2 Style Number 单独 (不拼 Color, 因为 Style Number 可能含 '-' 跟 Color 拼接会双横杠, 如 '21CMSS175A-006453S-RE' + 'RE-COLOR KAKY' → '21CMSS175A-006453S-RE-RE-COLOR KAKY')
        sku_final = sn

        # 类目: c19 Collection (e.g. 'PRE COLLECTION') — 没字段就空
        category = collection

        # 性别: c17 Linesheet 末段 (MAN / WOMAN / UNISEX / KID)
        gender = ""
        if linesheet:
            for kw, g in [
                ("WOMAN", "Woman"),
                ("WOMEN", "Woman"),
                ("MAN", "Man"),
                ("MEN", "Man"),
                ("KID", "Kid"),
                ("UNISEX", "Unisex"),
            ]:
                if kw in linesheet.upper():
                    gender = g
                    break

        # 图: 0-idx row = r, 但 df 用 0-idx; images_by_row 也是 0-idx key
        row_images = images_by_row.get(r, [])

        rows.append({
            "brand": brand or "C.P. COMPANY",
            "sku": sku_final,
            "image": "",
            "description": description,
            "category": category,
            "material": "",
            "color": color,
            "color_code": color_code,
            "gender": gender,
            "season": season_name,
            "origin": "",
            "cost": wholesale,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            "size": size,
            "quantity": qty,
            "vendor_id": vendor_id,
            "image_items": row_images,
        })

    return rows


# ================= SALOMON 'Foglio1' 版式 (2026-10-27 新增) =================
#
# SALOMON FW27 / AW27 订货单, 结构与 V1004 原有的 10-row/product 块**完全不同**,
# 所以单独一个分支, 走 openpyxl (要拿内嵌图, pandas 拿不到)。
#
# 文件结构 (sheet 'Foglio1', 885 行 x 49 列):
#   R1-R3  规则区 (RULES / ANY OTHER FAMILY / MIN 10 PIECES …)
#   R4     段标题 'SNEAKER SIZE US:' + US 码标签 c22-c40
#   R5     段标题 'ACCESSORY (ONE SIZE)' + 标签 c22
#   R6     段标题 'APPAREL & SOCKS' + 标签 c22-c26 (XS/S/M/L/XL)
#   R7     **字段表头行** (BU | SEGMENT | … | Article code | … | WHLS | …)
#          同时 c22-c40 还是**尺码标签行** (3.5 / 4 / 4.5 / 5 …)
#   R8+    数据行: 每行一个 Article code, c22+ 各尺码列是数量
#          c42='Unit total' / c43='Value total' 是汇总列, 不算尺码
#
# 尺码分段: 同一列区间在不同产品大类下含义不同, 所以按 BU(c1) 选段:
#   FOOTWEAR               → c22-c40, 标签取表头行 R7 (与 Foglio3 的 Size 列一致)
#   BAGS / ACCESSORIES    → c22,       标签取 R5 (ONE SIZE)
#   APPAREL / SOCKS / HEADWEAR → c22-c26, 标签取 R6 (XS/S/M/L/XL)
#
# 另有一个 sheet 'Foglio3' (每行 颜色×尺码明细), 但它的 Quantity Required 全是 0,
# 而 Foglio1 里有真实的 cost/retail/gender/图片, 所以以 Foglio1 为主表。

_SALOMON_SEGMENTS: Tuple[Tuple[str, str], ...] = (
    # (BU 关键词, 段 key)
    ("FOOTWEAR", "sneaker"),
    ("BAG", "accessory"),
    ("ACCESSOR", "accessory"),
    ("CATALOG", "accessory"),
    ("APPAREL", "apparel"),
    ("SOCK", "apparel"),
    ("HEADWEAR", "apparel"),
)


def _salomon_segment_of(bu: str) -> str:
    """按 BU(c1) 决定该行属于哪个尺码段."""
    u = (bu or "").upper()
    for kw, key in _SALOMON_SEGMENTS:
        if kw in u:
            return key
    return "sneaker"


def _salomon_size_cols(ws, header_row: int, seg_row: int, seg_key: str
                        ) -> List[Tuple[int, str]]:
    """返回该段的 [(列号, 尺码标签)].

    sneaker 段: 标签在字段表头行 header_row (c22-c40), 取**数字型**标签
                 (c41='CHECK ACS' / c42='Unit total' / c43='Value total' 不是尺码)
    其它段:     标签在段标题行 seg_row, 取该行 c22 起的连续非空标签
    """
    out: List[Tuple[int, str]] = []
    if seg_key == "sneaker":
        for c in range(22, ws.max_column + 1):
            v = ws.cell(header_row, c).value
            if v is None:
                break
            s = str(v).strip()
            if not s:
                break
            try:
                float(s)               # 只要数字型 (3.5 / 4 / 4.5 …)
            except ValueError:
                break
            out.append((c, s))
        return out

    for c in range(22, ws.max_column + 1):
        v = ws.cell(seg_row, c).value
        if v is None:
            break
        s = str(v).strip()
        if not s:
            break
        out.append((c, s))
    return out


def _is_salomon_foglio(primary_file: Path) -> bool:
    """是不是 SALOMON 'Foglio1' 版式."""
    try:
        wb = openpyxl.load_workbook(primary_file, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        if "Foglio1" not in wb.sheetnames:
            return False
        ws = wb["Foglio1"]
        for r in range(1, 12):
            for c in range(1, 12):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip() == "Article code":
                    return True
        return False
    except Exception:
        return False
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _parse_salomon_foglio(primary_file: Path, vendor_id: str
                           ) -> List[Dict[str, object]]:
    """解析 SALOMON 'Foglio1' 版式.

    两个 sheet 配合用:
      - 'Foglio3' = **明细表** (每行一个 颜色 × 尺码): 款名 / 颜色名 / 货号 /
        尺码 / UPC / 数量 / 品类  → 作为**输出行的来源**(有颜色!)
      - 'Foglio1' = 款号汇总表 (每行一个 Article code): 成本 / 零售价 / 性别 /
        品类 / 内嵌图片 → 按货号关联过来补字段

    关联键: 'Foglio3' c3 (Color Code) = 'Foglio1' c6 (Article code)。
    实测 4209 行里 3049 行能关联上(新款 Foglio1 里还没有), 关联不上的字段留空。
    """
    rows: List[Dict[str, object]] = []
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    try:
        if "Foglio3" not in wb.sheetnames:
            return rows
        w3 = wb["Foglio3"]
        w1 = wb["Foglio1"] if "Foglio1" in wb.sheetnames else None

        # ---------- Foglio1: 按货号建索引 (补字段 + 图片) ----------
        info_by_code: Dict[str, Dict[str, object]] = {}
        images_by_row: Dict[int, List[Dict[str, object]]] = {}
        if w1 is not None:
            images_0 = vendor_v1021._load_images_by_row(primary_file)
            images_by_row = {r + 1: v for r, v in images_0.items()}
            for r in range(1, min(w1.max_row, 15) + 1):
                hit = False
                for c in range(1, 12):
                    v = w1.cell(r, c).value
                    if isinstance(v, str) and v.strip() == "Article code":
                        hit = True
                        break
                if hit:
                    header_row = r
                    break
            else:
                header_row = None
            if header_row:
                COL: Dict[str, int] = {}
                for c in range(1, 22):
                    v = w1.cell(header_row, c).value
                    if not isinstance(v, str):
                        continue
                    k = v.strip().lower()
                    if k == "bu" and "bu" not in COL:
                        COL["bu"] = c
                    elif k == "segment" and "segment" not in COL:
                        COL["segment"] = c
                    elif k == "gender" and "gender" not in COL:
                        COL["gender"] = c
                    elif k == "article code" and "sku" not in COL:
                        COL["sku"] = c
                    elif k == "article name" and "name" not in COL:
                        COL["name"] = c
                    elif k == "whls" and "cost" not in COL:
                        COL["cost"] = c
                    elif k.startswith("model euro") and "retail" not in COL:
                        COL["retail"] = c
                if "sku" in COL:
                    for r in range(header_row + 1, w1.max_row + 1):
                        code = w1.cell(r, COL["sku"]).value
                        if code is None:
                            continue
                        code = str(code).strip()
                        if not code or code.lower() == "material":
                            continue
                        bu = w1.cell(r, COL["bu"]).value if "bu" in COL else ""
                        segment = (w1.cell(r, COL["segment"]).value
                                    if "segment" in COL else "")
                        cat_parts = [str(x).strip() for x in (bu, segment)
                                     if x and str(x).strip()]
                        info_by_code[code] = {
                            "description": str(w1.cell(r, COL["name"]).value or "").strip()
                            if "name" in COL else "",
                            "cost": to_float(w1.cell(r, COL["cost"]).value)
                            if "cost" in COL else None,
                            "retail": to_float(w1.cell(r, COL["retail"]).value)
                            if "retail" in COL else None,
                            "gender": str(w1.cell(r, COL["gender"]).value or "").strip()
                            if "gender" in COL else "",
                            "category": " / ".join(cat_parts),
                            "images": images_by_row.get(r, []),
                        }

        # ---------- Foglio3: 输出行 ----------
        # 表头: Style Name | Color Name | Color Code | Size | Alt Size
        #        | UPC/EAN | Quantity Required | Section
        H = {}
        for c in range(1, min(w3.max_column, 12) + 1):
            v = w3.cell(1, c).value
            if isinstance(v, str):
                H[v.strip().lower()] = c
        col_style = H.get("style name")
        col_color = H.get("color name")
        col_code = H.get("color code")
        col_size = H.get("size")
        col_section = H.get("section")
        col_qty = None
        for k, c in H.items():          # 'Quantity Requested' / 'Quantity Required' 变体
            if k.startswith("quantity"):
                col_qty = c
                break
        if not col_code:
            return rows

        for r in range(2, w3.max_row + 1):
            sku = w3.cell(r, col_code).value
            if sku is None or not str(sku).strip():
                continue
            sku = str(sku).strip()

            style = str(w3.cell(r, col_style).value or "").strip() if col_style else ""
            color = str(w3.cell(r, col_color).value or "").strip() if col_color else ""
            size = str(w3.cell(r, col_size).value or "").strip() if col_size else ""
            section = str(w3.cell(r, col_section).value or "").strip() if col_section else ""
            qty = to_float(w3.cell(r, col_qty).value) if col_qty else None

            info = info_by_code.get(sku, {})
            rows.append({
                # 品牌: 源文件是 SALOMON 专用订货单 (BU 段名里就有 'Salomon Apparel' /
                # 'Salomon Headwear'), 品牌名取 SALOMON; 不用文件名兜底(文件名是订货单名)
                "brand": "SALOMON",
                "sku": sku,
                "image": "",
                "description": style or info.get("description", ""),
                "category": info.get("category") or section,
                "material": "",
                "color": color,
                "color_code": "",
                "gender": info.get("gender", ""),
                "season": "",
                "origin": "",
                "cost": info.get("cost"),
                "cplus_text": "",
                "supply_price": None,
                "retail_price": info.get("retail"),
                "size": size,
                "quantity": int(round(qty)) if isinstance(qty, (int, float)) else None,
                "vendor_id": vendor_id,
                "image_items": info.get("images", []),
            })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
    brand_override: str = "",
) -> List[Dict[str, object]]:
    """V1004 解析入口:结构判断找块 + 复用 V1021 helper 解析细节。

    与 V1021 的区别:
    - V1021 用 `_is_sku()` 内容判断找块(13-40 字符 + 6+ 数字)
    - V1004 用"周围结构"判断找块(描述行有欧洲价 + 尺码行存在)
    - 强结构 → 强健壮:任何品牌的 SKU 长度都能吃,不会误判描述/段头为 SKU

    独立性:本函数不修改 V1021 任何全局变量,只 import 用到的 helper 函数。
    """
    rows: List[Dict[str, object]] = []

    # 0) 2026-10-27: SALOMON 'Foglio1' 版式 (FW27 / AW27)
    #    结构与本适配器原有的 10-row/product 块完全不同, 必须在这里先分流,
    #    否则后面 pd.read_excel 会把 45MB 的文件整个读一遍再白跑。
    if primary_file.suffix.lower() == ".xlsx" and _is_salomon_foglio(primary_file):
        return _parse_salomon_foglio(primary_file, vendor_id)

    # 1) 加载图片(.xlsx; .xls 走兜底)
    images_by_row = vendor_v1021._load_images_by_row(primary_file)

    # 2) 读 sheet
    df = pd.read_excel(primary_file, sheet_name=0, header=None)
    if df.empty:
        return rows

    # 2026-10-13: detect C.P. Company / NuORDER ATS 1-row/size 版式
    # R2 是品牌, R4 header c2='Style Number' → 走 ATS 单行版式
    if df.shape[1] >= 19 and df.shape[0] >= 5:
        try:
            r4_c2 = df.iloc[3, 1]  # 0-idx row 3, col 1
            if isinstance(r4_c2, str) and r4_c2.strip() == "Style Number":
                return _parse_v1004_ats(df, images_by_row, vendor_id)
        except (IndexError, KeyError):
            pass

    # 3) 找品牌 / season(header 区域)
    brand_name_raw = vendor_v1021._find_brand_name(df)
    brand_name = vendor_v1021._normalize(brand_name_raw) if brand_name_raw else ""
    header_meta = vendor_v1021._extract_meta_from_header(df)
    season_value = header_meta.get("season", "")

    # 4) C+ map(仅 POLO 用,STONE ISLAND/PATAGONIA 没有 conditions 文件)
    cplus_map: Dict[str, int] = {}
    if conditions_file:
        cplus_map = vendor_v1021._load_cplus_map(conditions_file)
    cplus = cplus_map.get(brand_name)
    if cplus is None:
        for cond_brand, cond_cplus in cplus_map.items():
            if brand_name.startswith(cond_brand) or cond_brand.startswith(brand_name):
                cplus = cond_cplus
                break

    # 5) 找产品块(用 V1004 自己的结构判断,不看 SKU 内容)
    blocks = _find_product_blocks_v1004(df)

    for block in blocks:
        sku = block["sku"]
        desc_text = block["desc_text"]
        description, cost, retail, origin, material = vendor_v1021._parse_description(desc_text)

        # 拆颜色:POLO 模式优先(SKU 后缀含 C####),其他品牌 description 拆
        sku_color = vendor_v1021._extract_color_from_sku(sku)
        if sku_color:
            color = sku_color
            # 货号用 - 拼接(per skill §3.5)
            number_part = vendor_v1021._strip_color_from_sku(sku)
            sku = f"{number_part}-{color}"
        else:
            description, color = vendor_v1021._split_color_from_description(description)

        # 类目(段位继承,2026-08-04:用 V1004 专属查找代替 V1021 的 30 行搜)
        # 找上一个性别段头(Donna/Uomo),在该范围内找最近类目词(Abbigliamento/Accessori/Scarpe/Borse/...)
        category = _find_category_v1004(df, block["sku_row"])

        # 尺码 + qty(已经在 _find_product_blocks_v1004 阶段验证过非空)
        size_map = block["size_map"]
        qty_map = block["qty_map"]

        # 图片(SKU 行附近 ±2 行)
        image_item = vendor_v1021._find_image_for_block(images_by_row, block["sku_row"])

        for col, size in size_map.items():
            qty = qty_map.get(col, 0)
            if qty > 0:
                rows.append({
                    "brand": brand_name,
                    "sku": sku,
                    "image": "",
                    "description": description,
                    "category": category,
                    "material": material,
                    "color": color,
                    "season": season_value,
                    "origin": origin,
                    "cost": cost,
                    "cplus_text": cplus if cplus else "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "image_items": [image_item] if image_item else [],
                })

    return rows


def get_vendor_meta() -> Dict[str, object]:
    """返回该 vendor 的默认元数据。

    ⚠️ 不设 default_brand — V1004 工厂覆盖多品牌(POLO / STONE ISLAND / PATAGONIA),
    让 dispatcher 走"文件名兜底"分支(从 primary_file.stem 拿品牌)。
    用户也可以在 CLI 显式 --brand "POLO RALPH LAUREN" 覆盖。
    """
    return {}
