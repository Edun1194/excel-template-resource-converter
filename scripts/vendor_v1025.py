"""V1025 POLO RALPH LAUREN 现货单 (2026-09-16 新增) — xlsx 新版式。

来源:POLO 工厂新版 .xlsx 现货单 (Ralph Lauren 系),不再是 .xls 老 OLE2 格式。
        之前 V1042 处理的是 .xls 老版式(6 行/product 块,F010 Client Anchor),
        V1025 处理的是 .xlsx 新版式(3 行/product 块,ws._images 拿图)。

结构 (单 sheet "Worksheet", 328 行 15 列):
  R1: REPORT DATE label
  R2: datetime
  R3: 空
  R4: header (1-idx): IMAGE | SKU | BRAND | PRODUCT | SEX | COMPOSITION | MADE IN
                   | RETAIL PRICE | COST PRICE | Sizes / Available quantitites / Desired quantities
  R5+: 每件产品 3 行块:
    R+0: metadata 行 (col 2=SKU, col 3=BRAND, col 4=PRODUCT, col 5=SEX,
                       col 6=COMPOSITION, col 7=MADE IN, col 8=RETAIL, col 9=COST,
                       col 10..N=size labels, col N+1='TOTAL')
    R+1: Available qty (col 10..N=numbers, col N+1=totale int)
    R+2: Desired qty (col 10..N=0, col N+1='0' 字符串) — 不输出到模板

SKU 格式: `<style>_<color_code>` (单 `_` 分隔符)
  例: "211870245_012" → style="211870245", color_code="012"
  重组: 保留单 `_` 分隔符(V1025 风格,跟 V1042 POLO 的 `style-color_code` 单 `-` 不同)

性别翻译: Female→Woman, Male→Man, Man→Man, Woman→Woman

产地归一: VN→Vietnam, IN→India, KH→Cambodia, BD→Bangladesh,
         ID→Indonesia, TR→Turkey, "Made in Italy"→Italy, "Made in Turkey"→Turkey

排除 qty=0 行 (跟 framework `is_positive_qty` 一致)。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


VENDOR_ID = "V1025"


# 性别翻译
GENDER_MAP = {
    "FEMALE": "Woman", "WOMAN": "Woman",
    "MALE": "Man", "MAN": "Man",
}

# 产地归一 (POLO 工厂写法五花八门,统一到标准名)
ORIGIN_MAP = {
    "VN": "Vietnam", "VIETNAM": "Vietnam",
    "IN": "India", "INDIA": "India",
    "KH": "Cambodia", "CAMBODIA": "Cambodia",
    "BD": "Bangladesh", "BANGLADESH": "Bangladesh",
    "ID": "Indonesia", "INDONESIA": "Indonesia",
    "CN": "China", "CHINA": "China",
    "TR": "Turkey", "TÜRKIYE": "Turkey", "TURKIYE": "Turkey", "TURKEY": "Turkey", "MADE IN TURKEY": "Turkey",
    "IT": "Italy", "ITALY": "Italy", "MADE IN ITALY": "Italy",
    "PH": "Philippines", "PHILIPPINES": "Philippines",
    "PT": "Portugal", "PORTUGAL": "Portugal",
    "LK": "Sri Lanka", "SRI LANKA": "Sri Lanka",
    "MA": "Morocco", "MOROCCO": "Morocco",
    "RO": "Romania", "ROMANIA": "Romania",
    "EG": "Egypt", "EGYPT": "Egypt",
    "TN": "Tunisia", "TUNISIA": "Tunisia",
    "PK": "Pakistan", "PAKISTAN": "Pakistan",
    "GE": "Georgia", "GEORGIA": "Georgia",
    "PE": "Peru", "PERU": "Peru",
    "NL": "Netherlands", "NETHERLANDS": "Netherlands",
    "FR": "France", "FRANCE": "France",
    "DE": "Germany", "GERMANY": "Germany",
    "GR": "Greece", "GREECE": "Greece",
    "ES": "Spain", "SPAIN": "Spain",
    "BG": "Bulgaria", "BULGARIA": "Bulgaria",
    "MD": "Moldova", "MOLDOVA": "Moldova",
    "MM": "Myanmar", "MYANMAR": "Myanmar",
    "NP": "Nepal", "NEPAL": "Nepal",
    "RS": "Serbia", "SERBIA": "Serbia",
    "TR": "Türkiye",
    # TR 已合并到上面(Turkey),保留此行仅为注释:TR 在 Python dict 里重复键会导致后者覆盖前者,
    # 所以统一只保留一份 Turkey 映射
}


def _norm_origin(raw: str) -> str:
    """产地归一,POLO 工厂写 'VN' / 'Vietnam' / 'Made in Vietnam' 都映射到 'Vietnam'。"""
    if not raw:
        return ""
    s = normalize_text(raw).upper().replace("MADE IN ", "").strip()
    return ORIGIN_MAP.get(s, normalize_text(raw).title())


def parse_rows(primary_file: Path, _: Path, vendor_id: str) -> List[Dict[str, object]]:
    """读 V1025 POLO .xlsx,返回 18 列 row dict 列表。"""
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    ws = wb.active
    rows: List[Dict[str, object]] = []

    # 找所有 SKU 行 (col 2 非空) — R5+ 范围
    sku_row_indices: List[int] = []
    for r in range(5, ws.max_row + 1):
        sku_v = ws.cell(r, 2).value
        if sku_v and str(sku_v).strip():
            sku_row_indices.append(r)

    # 抽图:openpyxl drawing 拿所有图(后面按 row index 匹配)
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

    for r in sku_row_indices:
        sku_raw = str(ws.cell(r, 2).value).strip()
        if not sku_raw:
            continue

        # 拆 SKU: <style>_<color_code> (V1025 单 `_` 分隔符)
        # 例: "211870245_012" → style="211870245", color_code="012"
        # 例: "211941509C_002" → style="211941509C", color_code="002"
        color_code = ""
        if "_" in sku_raw:
            parts = sku_raw.split("_")
            if len(parts) == 2:
                color_code = parts[1]
                # sku 保持原 `<style>_<color_code>` (V1025 风格不重组)

        # metadata
        brand = normalize_text(ws.cell(r, 3).value)
        product = normalize_text(ws.cell(r, 4).value)
        gender_raw = normalize_text(ws.cell(r, 5).value).upper()
        gender = GENDER_MAP.get(gender_raw, gender_raw.title())
        composition = normalize_text(ws.cell(r, 6).value)  # 当 material
        origin_raw = normalize_text(ws.cell(r, 7).value)
        origin = _norm_origin(origin_raw)
        retail = to_float(ws.cell(r, 8).value)
        cost = to_float(ws.cell(r, 9).value)

        # size label: R+0 col 10..N (最后一列是 'TOTAL',跳过)
        # 找 R+0 第一个空 col 作为 size 结束 (TOTAL 之前的最后一个 size)
        size_labels: List[str] = []
        size_cols: List[int] = []
        for c in range(10, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None or (isinstance(v, str) and not v.strip()):
                break  # 空 col 结束
            if isinstance(v, str) and v.strip().upper() == "TOTAL":
                break
            size_labels.append(normalize_text(v))
            size_cols.append(c)

        # R+1 = available qty
        qty_row = r + 1
        image_items = images_by_row.get(r, [])

        for c_idx, size_label in enumerate(size_labels):
            if not size_label:
                continue
            qty_v = ws.cell(qty_row, size_cols[c_idx]).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            rows.append(
                {
                    "brand": brand,
                    "sku": sku_raw,
                    "image": "",
                    "description": product,
                    "category": "",
                    "material": composition,
                    "color": "",
                    "color_code": color_code,
                    "season": "",
                    "origin": origin,
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