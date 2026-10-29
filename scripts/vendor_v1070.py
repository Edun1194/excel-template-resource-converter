"""V1070 vendor adapter — BRUNELLO CUCINELLI (BC) 现货单。

来源:BC FW25 / BC Woman new charge .xlsx(单 sheet "Worksheet")。

版式(标准,MAN 和 WOMAN 通用,新旧版通用):
- R1 (1-idx): 标题(c1="BC FW25 MAN"/"BC Woman new charge", c10/c11/c12 = 总零售/总数量,跳过)
- R2 (1-idx): header
  - 必含列(任何位置):SKU / BRAND / PRODUCT / RETAIL PRICE / TOT.QTY
  - 可能含列:GROUP(新版 "BC Woman new charge(1).xlsx" 在 PRODUCT 后多了 GROUP 列)
  - 尺码列:header 中其他非空非识别列
- R3+ (1-idx): 产品行
  - c*_SKU = 完整 SKU, c*_BRAND = 品牌, c*_PRODUCT = 类型,
    c*_SEASON = 季节(如 "FW"), c*_GENDER = 性别, c*_COLOR = 颜色,
    c*_COMPOSITION = 材质, c*_RETAIL = 零售价, c*_TOT_QTY = 总库存
  - 尺码列 = 各尺码 qty(OS/XXS/XS/S/M/L/XL/XXL/3XL/4XL/5XL/6XL + 鞋码 34-115)

特殊点:
- 源文件**无 cost 列**(只有 retail,无供货价/成本)→ cost = None
- 季节字段在 header SEASON 列(e.g. "FW"),V1070 从源读
- 货号 = SKU 列(完整 SKU,带颜色后缀)
- 多品牌(目前只 BC 一家,但 BRAND 列是通用支持)
- 新版 "BC Woman new charge" 文件 header 多了一列 GROUP,**所有后续列右移 1 位**,
  V1070 必须按 header 文本动态找列,不能用硬编码列偏移(2026-07-31 改)

识别策略(纯结构,无内容校验):
- 找 header 行:c2 == "SKU" 且 c3 == "BRAND" 且 c4 == "PRODUCT"(三个锚点列固定不动)
- 找各数据列:扫 header 行所有 col,用关键词匹配:
  RETAIL → retail_col, SEASON → season_col, GENDER → gender_col,
  COLOR → color_col, COMPOSITION → material_col,
  TOT.QTY / TOT QTY / TOTAL QTY / QTY / QUANTITY → tot_qty_col
- 尺码列:header 行中,跳过 IMAGE + 三个锚点 + 上述所有识别列后的非空 cell
- TOT.QTY > 0 验证产品行有效(防止空 SKU 行被误收)
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


# 0-idx 列偏移(版式固定)
# 注:R1 是 title 行(跳过),R2 header(c2=SKU / c3=BRAND / c4=PRODUCT),
# R3+ 数据。SKU 在 0-idx col 1 (1-idx c2),以下用 0-idx。
SKU_COL = 1            # c2 完整 SKU(锚点列,固定不动)
BRAND_COL = 2          # c3 BRAND(锚点列,固定不动)
PRODUCT_COL = 3        # c4 PRODUCT(锚点列,固定不动)
# 以下列**动态**识别(header 文本匹配),不用硬编码:
# SEASON / GENDER / COLOR / COMPOSITION / RETAIL / TOT_QTY
# 因为新版"BC Woman new charge"在 PRODUCT 后多了 GROUP 列,所有这些列右移 1 位

# 尺码列结束位置:header 行中跳过所有识别列后,直到第一个空 col
SIZE_COL_START = 4     # c5 起扫(header 中 c4 之后)


def _is_brand_header(headers_row: List[str]) -> bool:
    """判断是否是 V1070 header 行:c2='SKU' c3='BRAND' c4='PRODUCT'"""
    if len(headers_row) < 4:
        return False
    return (
        normalize_text(headers_row[1]).upper() == "SKU"
        and normalize_text(headers_row[2]).upper() == "BRAND"
        and normalize_text(headers_row[3]).upper() == "PRODUCT"
    )


def _find_header_row(ws) -> Optional[int]:
    """找 header 行(0-idx):扫前 5 行找 c2=SKU c3=BRAND c4=PRODUCT"""
    for r in range(min(5, ws.max_row)):
        row_vals = [
            normalize_text(ws.cell(row=r + 1, column=c + 1).value)
            for c in range(min(5, ws.max_column))
        ]
        if _is_brand_header(row_vals):
            return r
    return None


def _build_size_labels(ws, header_row: int, recognized_cols: set) -> List[int]:
    """从 header 行 c5+ 提取尺码 col 列表(0-idx),直到第一个空 cell

    尺码 label 可能是字符串(OS/XXS/XS/S/M/L/XL/XXL/3XL 等)或数字(鞋码 39/39.5/40)。
    数字需保留为原始 label 不归一化(39.5 鞋码意义不同)。

    recognized_cols: 已知列索引集合(锚点 + 动态识别列),这些列**不**算尺码列
    """
    cols: List[int] = []
    for c in range(SIZE_COL_START, ws.max_column):
        if c in recognized_cols:
            continue
        v = ws.cell(row=header_row + 1, column=c + 1).value
        if v is None or (isinstance(v, str) and not v.strip()):
            if cols:
                break  # 已有 token 后面空 → 本行结束
            continue
        cols.append(c)
    return cols


# 识别列关键词(uppercase / 大小写不敏感)→ 用于在 header 行中找各列位置
# 顺序:先匹配更长的关键词(RETAIL PRICE 在 RETAIL 之前检查)
_RECOGNIZED_HEADERS = {
    "RETAIL_PRICE": ["RETAIL PRICE", "RETAIL"],
    "SEASON": ["SEASON", "STAGIONE"],
    "GENDER": ["GENDER", "SETTORE"],
    "COLOR": ["COLOR", "COLOUR", "COLORE"],
    "COMPOSITION": ["COMPOSITION", "MATERIAL", "TESSUTO"],
    "TOT_QTY": ["TOT.QTY", "TOT QTY", "TOTAL QTY", "TOT. QTY", "QUANTITY", "QTY", "TOTALE PZ"],
}


def _find_column(header_vals: List[str], candidates: List[str]) -> Optional[int]:
    """在 header 行值列表中找第一个匹配任一关键词的列索引(0-idx)

    header_vals 已经 normalize_text + upper 化。candidate 也 upper 后 substring 匹配。
    """
    for idx, hv in enumerate(header_vals):
        if not hv:
            continue
        for cand in candidates:
            if cand.upper() in hv:
                return idx
    return None


def _resolve_columns(ws, header_row: int) -> Dict[str, Optional[int]]:
    """从 header 行动态解析各数据列的 0-idx 位置

    返回 dict:
      - 必有:'sku' / 'brand' / 'product' (锚点)
      - 可有:'retail' / 'season' / 'gender' / 'color' / 'composition' / 'tot_qty'
      - 还有:'recognized_cols' (set[int]):所有应跳过、不算尺码的列索引
             (含锚点列 + 上述所有识别列 + 任何 cell 文本含 _RECOGNIZED_HEADERS 关键词的列)
    """
    header_vals: List[str] = [
        normalize_text(ws.cell(row=header_row + 1, column=c + 1).value).upper()
        for c in range(ws.max_column)
    ]
    # 所有 candidate 关键词(展平 + 去重) — 用来在 header 中找出"非尺码"列
    all_keywords: set = set()
    for candidates in _RECOGNIZED_HEADERS.values():
        for c in candidates:
            all_keywords.add(c.upper())

    # 找各数据列
    retail_col = _find_column(header_vals, _RECOGNIZED_HEADERS["RETAIL_PRICE"])
    season_col = _find_column(header_vals, _RECOGNIZED_HEADERS["SEASON"])
    gender_col = _find_column(header_vals, _RECOGNIZED_HEADERS["GENDER"])
    color_col = _find_column(header_vals, _RECOGNIZED_HEADERS["COLOR"])
    composition_col = _find_column(header_vals, _RECOGNIZED_HEADERS["COMPOSITION"])
    tot_qty_col = _find_column(header_vals, _RECOGNIZED_HEADERS["TOT_QTY"])

    # recognized_cols:锚点 + 上述所有识别列 + 任何 cell 文本含关键词的列
    # (后一步是为了把 c11=TOTAL RETAIL 这类"含 RETAIL 但不是 retail 主列"的列也算进去,
    #  避免被 _build_size_labels 误当尺码列)
    recognized_cols: set = {SKU_COL, BRAND_COL, PRODUCT_COL}
    for c in [retail_col, season_col, gender_col, color_col, composition_col, tot_qty_col]:
        if c is not None:
            recognized_cols.add(c)
    for c, hv in enumerate(header_vals):
        if not hv:
            continue
        for kw in all_keywords:
            if kw in hv:
                recognized_cols.add(c)
                break
    # 锚点 c1=IMAGE 也要跳过(虽然不在 _RECOGNIZED_HEADERS 里,这里硬加)
    recognized_cols.add(0)  # IMAGE 列固定在 c1(0-idx 0)

    out: Dict[str, object] = {
        "sku": SKU_COL,
        "brand": BRAND_COL,
        "product": PRODUCT_COL,
        "retail": retail_col,
        "season": season_col,
        "gender": gender_col,
        "color": color_col,
        "composition": composition_col,
        "tot_qty": tot_qty_col,
        "recognized_cols": recognized_cols,
    }
    return out


def _format_size_label(v) -> str:
    """尺码 label 归一化:数字 → int(无小数) / str;接受欧洲逗号"""
    if isinstance(v, (int, float)):
        f = float(v)
        if f.is_integer():
            return str(int(f))
        return str(f).replace(",", ".")
    sv = str(v).strip()
    return sv.replace(",", ".")


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 0-idx row 分桶图(图片 anchor _from.row 是 0-idx)"""
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


def get_vendor_meta() -> Dict[str, object]:
    """V1070 vendor 元数据 — 给 dispatcher 用。

    V1070 是多品牌工厂代码(目前只 BC 一家,但 c3 字段是 BRAND 列通用支持),
    源文件有 BRAND 字段从源读,空时让 main 走 4 级文件名兜底,不设 default_brand。
    default_image_col_width=40:每行多图时(BC 数据行通常 1-3 张)需要更宽。
    """
    return {
        "default_image_col_width": 40.0,
    }


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    """V1070 — BRUNELLO CUCINELLI 现货单(FW25 / Woman new charge)。

    注:2026-09-02 起,NuORDER 期货版拆出去独立成 `V1070-期货`(vendor_v1070_nuorder.py),
    本 adapter 只处理 FW25 现货 / Woman new charge 现货版式。
    """
    rows: List[Dict[str, object]] = []

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    if not wb.worksheets:
        return rows
    ws = wb.worksheets[0]

    # 1. header 行
    header_row = _find_header_row(ws)
    if header_row is None:
        return rows

    # 2. 动态解析各列位置(应对 GROUP 列插入导致后续列右移的情况)
    col_map = _resolve_columns(ws, header_row)
    tot_qty_col = col_map["tot_qty"]
    retail_col = col_map["retail"]
    season_col = col_map["season"]
    gender_col = col_map["gender"]
    color_col = col_map["color"]
    material_col = col_map["composition"]
    recognized_cols = col_map["recognized_cols"]
    # 至少要能识别 TOT.QTY(产品行核心信号,识别不了就提前返回)
    if tot_qty_col is None:
        return rows

    # 3. 尺码 col 列表(0-idx):跳过所有识别列
    size_cols = _build_size_labels(ws, header_row, recognized_cols)
    if not size_cols:
        return rows

    # 4. 资源
    images_by_row = _load_images_by_row(ws)
    sheet_row_heights = load_sheet_row_heights(primary_file)

    # 5. 尺码 label 缓存 {0-idx col: label_str}
    size_label_map: Dict[int, str] = {}
    for c in size_cols:
        size_label_map[c] = _format_size_label(ws.cell(row=header_row + 1, column=c + 1).value)

    # 6. 遍历产品行(0-idx):header_row 下方
    for r in range(header_row + 1, ws.max_row):
        # 6.1 SKU = col 1(0-idx)非空字符串
        sku_v = ws.cell(row=r + 1, column=SKU_COL + 1).value
        if not isinstance(sku_v, str) or not sku_v.strip():
            continue
        # 2026-08-14:SKU 拼接符统一 `-`(源用 `_` 拼 style+color_code+color_name,改 `-`)
        # 例: "MM8433103G_TM843_C210" → "MM8433103G-TM843-C210"
        sku = sku_v.strip().replace(".", "-").replace("_", "-")

        # 6.2 TOT.QTY 必须 > 0(产品行核心信号,防止空 SKU 行/分隔行被误收)
        tot_qty_v = ws.cell(row=r + 1, column=tot_qty_col + 1).value
        if not isinstance(tot_qty_v, (int, float)) or tot_qty_v <= 0:
            continue

        # 6.3 retail(RETAIL PRICE)可空(没成本价,retail 缺失时 retail_price=None)
        retail: Optional[float] = None
        if retail_col is not None:
            rv = ws.cell(row=r + 1, column=retail_col + 1).value
            if isinstance(rv, (int, float)) and rv > 0:
                retail = float(rv)

        # 6.4 brand 从源读(可能为空,让 main 走文件名兜底)
        brand = normalize_text(ws.cell(row=r + 1, column=BRAND_COL + 1).value)
        # 6.5 描述/类型/季节/性别/颜色/材质(动态列)
        product = normalize_text(ws.cell(row=r + 1, column=PRODUCT_COL + 1).value)
        season_raw = (
            normalize_text(ws.cell(row=r + 1, column=season_col + 1).value)
            if season_col is not None else ""
        )
        gender = (
            normalize_text(ws.cell(row=r + 1, column=gender_col + 1).value)
            if gender_col is not None else ""
        )
        color = (
            normalize_text(ws.cell(row=r + 1, column=color_col + 1).value)
            if color_col is not None else ""
        )
        material = (
            normalize_text(ws.cell(row=r + 1, column=material_col + 1).value)
            if material_col is not None else ""
        )
        # 季节字段: 源 SEASON 列 (e.g. "FW"),app 端填写会在 main 覆盖

        # 6.6 描述 = product + gender (e.g. "BERMUDA MAN"),便于用户看
        desc = " ".join(filter(None, [product, gender]))

        # 6.6b gender 单独存到 row dict,供中文模板"性别"列输出
        # 源 GENDER 列通常是 "MAN" / "WOMAN",直接保留原值(2026-07-31)

        # 6.7 类目:直接用源 PRODUCT 字段(原值保留,如 "ESPADRILLAS" / "BERMUDA" / "JACKET")
        # 不走 guess_product_group 推断 — 用户希望类目列直接显示源文件的 PRODUCT 名
        category = product

        # 6.8 图片(0-idx row = r)
        row_images = images_by_row.get(r, [])

        # 6.9 源行高(1-idx)
        source_row_h = sheet_row_heights.get(ws.title, {}).get(r + 1)

        # 6.10 按尺码展开
        for c in size_cols:
            label = size_label_map[c]
            qty_v = ws.cell(row=r + 1, column=c + 1).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            if qty <= 0:
                continue

            rows.append({
                "brand": brand,  # 源有 brand 从源读,空时让 main 走文件名兜底
                "sku": sku,
                "image": "",
                "description": desc,
                "category": category,
                "material": material,
                "color": color,
                "gender": gender,  # 源 GENDER 列(动态识别,新版"BC Woman new charge"也支持)
                "season": season_raw,  # 源 SEASON 列 (e.g. "FW"),app 端填写会在 main 覆盖
                "origin": "",
                "cost": None,  # 源无 cost 列
                "cplus_text": "",
                "supply_price": None,  # 无 cost → 无 supply_price
                "retail_price": retail,
                "size": label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": source_row_h,
            })

    return rows
