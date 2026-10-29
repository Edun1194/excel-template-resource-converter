"""V_CK CALVIN KLEIN (CK) Take-All-Stock vendor adapter (2026-08-06 新建)

适配文件:CALVIN KLEIN 泰国清仓库存清单 (CK - TH TAKE ALL STOCK CLIENT.xlsx)。
没有工厂代码,暂定 V_CK(跟 VC001 一样不污染 BUILTIN_VENDORS,只在测试目录用)。

文件结构 — 3 个 sheet:
  - 'STOCK '  (418 rows × 34 cols, 0 images)  主数据, R4 header, R5+ 数据
  - 'Foglio1' (7 rows × 3 cols, 0 images)     统计汇总(总库存/总成本/总零售)
  - 'Images'  (256 rows × 4 cols, 0 images)   图片索引表, R1 header, R2+ 图 + GA
                                              c1=图 / c2=Generic Article / c3=Brand / c4=Offer

⚠️ openpyxl 不识别 Images sheet 的图(用 <xdr:sp>+<a:blipFill> 嵌入,不是 <xdr:pic>)。
   本 vendor 自己解析 drawing1.xml,提取 row → image bytes 映射。

字段映射(STOCK sheet R5+):
  SKU        ← c4 Article (e.g. "000NP1929O1GB002")
  描述       ← c5 Description (e.g. "2 PK LOW RISE TRUNK, 1GB, XS")
  颜色       ← c7 Article Color (e.g. "1GB" / "100" / "B2B" / "NYZ")
  色号       ← "" 留空(色号是 c7 内部编码,没有独立色名字段)
  尺码       ← c9 Article Size 1 (XS/S/M/L/...),c10 是 secondary size 一般为 ##
  库存       ← c23 PCS
  cost       ← c24 WHS
  retail     ← c25 RETAIL
  季节       ← f"{c19}{c20}"  (e.g. "999900" = NOS Current)
  品牌       ← "" 留空(main 走文件名兜底 → "CALVIN KLEIN")
  性别       ← c14 Gender:Men→"Man" / Women→"Woman" / 其他→原文
  类目       ← c17 Sub Category (e.g. "Underwear Bottoms Multi")
  云仓        ← VENDOR_ID = "V_CK"
  图         ← 跨 sheet 匹配:STOCK c3 (Generic Article) → Images c2 → drawing1.xml 拿 bytes

图片匹配流程:
  1) 解析 xl/drawings/drawing1.xml → 每个 <xdr:twoCellAnchor> 的 <xdr:row> + <a:blip r:embed>
  2) 解析 xl/drawings/_rels/drawing1.xml.rels → rId → image file
  3) 读 Images sheet R2+ c2 (Generic Article),得到 {row_0idx: GA}
  4) 反向: {GA: [image_bytes]}
  5) STOCK sheet R5+ c3 (Generic Article) 查 {GA} 拿图,KEEP_IMAGES_PER_ROW=1 取第一张
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 V_CK(2026-08-06 用户说"没有工厂代码",暂定)
VENDOR_ID = "V_CK"

# R4 是 STOCK sheet header (1-idx)
STOCK_HEADER_ROW = 4
# R1 是 Images sheet header (1-idx)
IMAGES_HEADER_ROW = 1


def get_vendor_meta() -> Dict[str, object]:
    """V_CK vendor 元数据。"""
    return {
        "default_image_col_width": 40.0,
    }


# STOCK sheet 列(0-idx)
STOCK_COL_BRAND = 0          # c1 Article Brand
STOCK_COL_GENERIC_ARTICLE = 2  # c3 Generic Article (e.g. "000NP1929O1GB")
STOCK_COL_ARTICLE = 3         # c4 Article (完整 SKU, e.g. "000NP1929O1GB002")
STOCK_COL_DESCRIPTION = 4     # c5 Description
STOCK_COL_COLOR = 6           # c7 Article Color
STOCK_COL_SIZE1 = 8           # c9 Article Size 1
STOCK_COL_SIZE2 = 9           # c10 Article Size 2 (usually "##")
STOCK_COL_GENDER = 13         # c14 Gender
STOCK_COL_SUB_CATEGORY = 16   # c17 Sub Category
STOCK_COL_PRODUCT_CLASS = 17  # c18 Product Class
STOCK_COL_SEASON_YEAR = 18    # c19 Season Year
STOCK_COL_SEASON_SUB = 19     # c20 Season Sub
STOCK_COL_PCS = 22            # c23 PCS (库存)
STOCK_COL_WHS = 23            # c24 WHS (cost)
STOCK_COL_RETAIL = 24         # c25 RETAIL

# Images sheet 列(0-idx)
IMAGES_COL_GENERIC_ARTICLE = 1  # c2 Generic Article


def _parse_drawing_images(xlsx_path: Path) -> Dict[int, List[Dict]]:
    """直接解析 xlsx drawing1.xml,获取 row → image bytes 映射。

    openpyxl 不识别 <xdr:sp>+<a:blipFill> 类型的图,本函数用 zipfile + 正则
    手动解 drawing1.xml + drawing1.xml.rels。

    返回: {0-idx row: [{'bytes': bytes, 'format': str}, ...]}
    """
    images_by_row: Dict[int, List[Dict]] = {}
    with zipfile.ZipFile(xlsx_path) as z:
        # 1) rels:rId → image file
        try:
            with z.open('xl/drawings/_rels/drawing1.xml.rels') as f:
                rels_content = f.read().decode('utf-8')
        except KeyError:
            return images_by_row
        rels = {}
        for m in re.finditer(r'<Relationship Id="([^"]+)"[^>]+Target="([^"]+)"', rels_content):
            rels[m.group(1)] = m.group(2)

        # 2) drawing1.xml:每个 <xdr:twoCellAnchor> 的 row + rId
        try:
            with z.open('xl/drawings/drawing1.xml') as f:
                drawing = f.read().decode('utf-8')
        except KeyError:
            return images_by_row
        anchor_re = re.compile(r'<xdr:twoCellAnchor>(.*?)</xdr:twoCellAnchor>', re.DOTALL)
        for m in anchor_re.finditer(drawing):
            anchor_xml = m.group(1)
            row_match = re.search(r'<xdr:from>.*?<xdr:row>(\d+)</xdr:row>', anchor_xml, re.DOTALL)
            rid_match = re.search(r'<a:blip[^>]*r:embed="([^"]+)"', anchor_xml)
            if not (row_match and rid_match):
                continue
            row = int(row_match.group(1))
            rid = rid_match.group(1)
            target = rels.get(rid, '')
            # target 通常是 "../media/image2.jpeg" — 相对 xl/drawings/
            if target.startswith('../'):
                img_path = 'xl/' + target[3:]
            elif target.startswith('/'):
                img_path = target.lstrip('/')
            else:
                img_path = 'xl/drawings/' + target
            try:
                with z.open(img_path) as f:
                    img_data = f.read()
            except KeyError:
                continue
            ext = img_path.rsplit('.', 1)[-1].lower()
            fmt = 'jpeg' if ext in ('jpg', 'jpeg') else ext
            images_by_row.setdefault(row, []).append({
                'bytes': img_data,
                'format': fmt,
            })
    return images_by_row


def _load_images_by_generic_article(xlsx_path: Path, images_ws) -> Dict[str, List[Dict]]:
    """从 Images sheet + drawing1.xml 构造 {Generic Article: [image_bytes]} 映射。"""
    # 1) 解析 drawing → row → image bytes
    images_by_row = _parse_drawing_images(xlsx_path)

    # 2) Images sheet R2+ c2 = Generic Article
    ga_by_row: Dict[int, str] = {}
    for r in range(IMAGES_HEADER_ROW + 1, images_ws.max_row + 1):
        ga = images_ws.cell(row=r, column=IMAGES_COL_GENERIC_ARTICLE + 1).value
        if ga and isinstance(ga, str) and ga.strip():
            ga_by_row[r - 1] = ga.strip()  # 0-idx row

    # 3) 反向: {GA: [image_bytes]}
    ga_to_images: Dict[str, List[Dict]] = {}
    for row, ga in ga_by_row.items():
        if row in images_by_row:
            ga_to_images[ga] = images_by_row[row]
    return ga_to_images


def _pick_gender(gender_raw: str) -> str:
    """归一化 Gender 字段。

    "Men" → "Man", "Women" → "Woman", 其他原文保留
    """
    if not gender_raw:
        return ""
    sv = str(gender_raw).strip()
    if sv.upper() == "MEN":
        return "Man"
    if sv.upper() == "WOMEN":
        return "Woman"
    return sv


def _pick_size(size1_raw, size2_raw) -> str:
    """选尺码:优先 c9 Article Size 1,c10 是 secondary(常为 "##" 或空)。"""
    if size1_raw and str(size1_raw).strip() and str(size1_raw).strip() != "##":
        return str(size1_raw).strip()
    if size2_raw and str(size2_raw).strip() and str(size2_raw).strip() != "##":
        return str(size2_raw).strip()
    return ""


def _parse_stock_sheet(
    stock_ws,
    ga_to_images: Dict[str, List[Dict]],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 STOCK sheet,返回 row dict 列表。

    每行 1 个 size+qty(STOCK sheet 是 per-row 模式,不是横排尺码)。
    """
    rows: List[Dict[str, object]] = []

    for r in range(STOCK_HEADER_ROW + 1, stock_ws.max_row + 1):
        # 1) SKU (Article, 完整带色号)
        sku = normalize_text(stock_ws.cell(row=r, column=STOCK_COL_ARTICLE + 1).value)
        if not sku:
            continue

        # 2) 尺码
        size1 = stock_ws.cell(row=r, column=STOCK_COL_SIZE1 + 1).value
        size2 = stock_ws.cell(row=r, column=STOCK_COL_SIZE2 + 1).value
        size = _pick_size(size1, size2)
        if not size:
            continue

        # 3) 库存
        pcs_v = stock_ws.cell(row=r, column=STOCK_COL_PCS + 1).value
        if not is_positive_qty(pcs_v):
            continue
        qty = int(round(float(to_float(pcs_v) or 0)))
        if qty <= 0:
            continue

        # 4) 描述
        desc = normalize_text(stock_ws.cell(row=r, column=STOCK_COL_DESCRIPTION + 1).value)

        # 5) 颜色 (Article Color, e.g. "1GB" / "100" / "B2B" / "NYZ")
        color = normalize_text(stock_ws.cell(row=r, column=STOCK_COL_COLOR + 1).value)

        # 6) 季节 (Season Year + Season Sub)
        season_year = stock_ws.cell(row=r, column=STOCK_COL_SEASON_YEAR + 1).value
        season_sub = stock_ws.cell(row=r, column=STOCK_COL_SEASON_SUB + 1).value
        season = ""
        if season_year is not None and season_sub is not None:
            season = f"{season_year}{int(season_sub):02d}" if isinstance(season_sub, (int, float)) else f"{season_year}{season_sub}"
        elif season_year is not None:
            season = str(season_year)

        # 7) cost / retail
        cost_v = stock_ws.cell(row=r, column=STOCK_COL_WHS + 1).value
        cost = to_float(cost_v)
        retail_v = stock_ws.cell(row=r, column=STOCK_COL_RETAIL + 1).value
        retail = to_float(retail_v)

        # 8) 性别
        gender_raw = stock_ws.cell(row=r, column=STOCK_COL_GENDER + 1).value
        gender = _pick_gender(gender_raw)

        # 9) 类目 (优先 Sub Category,空则 Product Class)
        category = normalize_text(stock_ws.cell(row=r, column=STOCK_COL_SUB_CATEGORY + 1).value)
        if not category:
            category = normalize_text(stock_ws.cell(row=r, column=STOCK_COL_PRODUCT_CLASS + 1).value)

        # 10) 图 (跨 sheet 匹配:STOCK c3 Generic Article → Images GA → image bytes)
        ga = stock_ws.cell(row=r, column=STOCK_COL_GENERIC_ARTICLE + 1).value
        ga_key = ga.strip() if isinstance(ga, str) else ""
        img_items = ga_to_images.get(ga_key, [])

        # 11) 品牌 — 留空(main 走文件名兜底 → "CALVIN KLEIN")

        rows.append({
            "brand": "",  # 留空,main 走文件名兜底
            "sku": sku,
            "image": "",
            "description": desc,
            "category": category,
            "material": "",
            "color": color,
            "color_code": "",  # 色号就是 c7 内部编码
            "gender": gender,
            "season": season,
            "origin": "",
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            "size": size,
            "quantity": qty,
            "vendor_id": vendor_id,
            "image_items": img_items,
            "source_row_height": None,
        })

    return rows


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 CALVIN KLEIN 泰国清仓 .xlsx,返回 row dict 列表。

    仓库 ID 统一用 VENDOR_ID = "V_CK",忽略传入 vendor_id。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    vendor_id = VENDOR_ID

    # 1) 解析 drawing1.xml(给 Images sheet 用)
    ga_to_images = _parse_drawing_images_to_ga(primary_file)

    # 2) 解析 STOCK + Images sheet
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    stock_ws = None
    images_ws = None
    for ws in wb.worksheets:
        title_norm = (ws.title or "").strip().upper()
        if title_norm.startswith("STOCK"):
            stock_ws = ws
        elif title_norm == "IMAGES":
            images_ws = ws

    if not stock_ws:
        return rows

    # 3) 重新构造 ga_to_images(用 images_ws 的 GA 列表,跟 drawing 配对)
    if images_ws is not None:
        ga_to_images = _load_images_by_generic_article(primary_file, images_ws)

    # 4) 解析 STOCK sheet
    rows = _parse_stock_sheet(stock_ws, ga_to_images, vendor_id)
    return rows


def _parse_drawing_images_to_ga(xlsx_path: Path) -> Dict[str, List[Dict]]:
    """Dummy 函数,实际 parse_rows 走 _load_images_by_generic_article 重新构造。"""
    return {}
