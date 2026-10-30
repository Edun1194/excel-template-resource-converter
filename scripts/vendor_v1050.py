"""V1050 FOLLIFOLLIE vendor adapter."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

from convert_vendor_quote_shared import (
    COMMON_SIZE_HEADERS,
    calc_supply_price,
    format_cplus_text,
    guess_product_group,
    is_positive_qty,
    load_richvalue_images_by_col,
    load_sheet_embedded_images,
    load_sheet_images_by_row,
    load_sheet_row_heights,
    normalize_brand,
    normalize_text,
    to_float,
)


# ---------------------------------------------------------------------------
# Markup helpers
# ---------------------------------------------------------------------------


def _parse_griglia(ws, vendor_id: str, images_by_row: Dict[int, List[Dict]]) -> List[Dict[str, object]]:
    """V1050 GRIGLIA sub-format (2026-09-03 接入:CP COMPANY ATS03.09.xlsx)。

    结构:
    - R1: header (c2='Style Number', c3='Style Name', c4='Fabrication Name',
                 c5='Fabrication Code', c6='Color', c7='Color Code',
                 c10='Currency', c11='Wholesale Price', c12='Retail Price',
                 c13='Size', c14='Units', c15='ORDER QTY', c16='TTL ORDER WHS')
    - R2: 空
    - R3+: data — 一行一个 size, c2=SKU, c6=color, c7=color_code,
                   c11=cost, c12=retail, c13=size, c14=qty

    2026-09-03 fix:同一 Style Number 跨多 color/size 行时,图按 Style Number 共享
    (源文件 ~514 张图对应 1388 行, 很多 size/color 行没 anchor 图, 用同 style 共享图补齐)
    """
    rows: List[Dict[str, object]] = []
    if ws.max_row < 3 or ws.max_column < 14:
        return rows

    # 1) 第一遍:扫所有 data 行,按 style number group 收集 image_items
    style_imgs: Dict[str, List[Dict]] = {}
    style_imgs_order: List[str] = []  # 保留 style 首次出现顺序,避免 dict 顺序乱
    data_rows_raw: List[Tuple[int, str, str, str]] = []  # (1-idx r, sku, color, size)
    for r in range(3, ws.max_row + 1):
        sku_v = ws.cell(r, 2).value
        if sku_v is None or not str(sku_v).strip():
            continue
        sku = str(sku_v).strip()

        size = normalize_text(ws.cell(r, 13).value)
        if not size:
            continue

        qty_v = ws.cell(r, 14).value
        try:
            qty = int(round(float(qty_v or 0)))
        except (TypeError, ValueError):
            qty = 0
        if qty <= 0:
            continue

        color = normalize_text(ws.cell(r, 6).value)
        data_rows_raw.append((r, sku, color, size))

        # 收集此 style 的图(去重:同 style 内多 color 行可能 anchor 不同图, 全部保留)
        row_imgs = images_by_row.get(r - 1, [])
        if row_imgs:
            if sku not in style_imgs:
                style_imgs[sku] = []
                style_imgs_order.append(sku)
            for img in row_imgs:
                # 简单去重:按 bytes hash
                img_id = hash(img.get("bytes", b""))
                existing_ids = {hash(x.get("bytes", b"")) for x in style_imgs[sku]}
                if img_id not in existing_ids:
                    style_imgs[sku].append(img)

    # 2) 第二遍:出 row, image_items 用 style 级 imgs
    for r, sku, color, size in data_rows_raw:
        desc = normalize_text(ws.cell(r, 3).value)
        fabrication = normalize_text(ws.cell(r, 4).value)
        color_code = normalize_text(ws.cell(r, 7).value)
        cost = to_float(ws.cell(r, 11).value)
        retail = to_float(ws.cell(r, 12).value)

        # 货号拼接:color 优先(Step 3.5.1)
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        if color_normalized:
            sku_final = f"{sku}-{color_normalized}"
            color_code_field = color_normalized.split()[0]
        else:
            sku_final = sku
            color_code_field = color_code  # 兜底用源 c7 color code

        # CP COMPANY 文件无显式 brand, 推断
        brand = "C.P. COMPANY"

        # 图:同 style 共享
        row_images = style_imgs.get(sku, [])

        rows.append({
            "brand": brand,
            "sku": sku_final,
            "image": "",
            "description": desc,
            "category": "",
            "material": fabrication,
            "color": color,
            "color_code": color_code_field,
            "gender": "",
            "season": "",
            "origin": "",
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            "size": size,
            "quantity": qty,
            "vendor_id": vendor_id,
            "image_items": row_images,
        })
    return rows


def _build_markup_map(cost_plus_path: Path) -> Dict[str, float]:
    result: Dict[str, float] = {}
    if not cost_plus_path.exists():
        return result
    for line in cost_plus_path.read_text(encoding="utf-8").strip().splitlines():
        line = line.strip()
        if not line:
            continue
        fixed = re.sub(r"(\d)O", r"\g<1>0", line)
        fixed = re.sub(r"O(\d)", r"0\1", fixed)
        m = re.search(r"C\+\s*(\d+(?:[.,]\d+)?)\s*%?", fixed, re.IGNORECASE)
        if not m:
            continue
        pct = float(m.group(1).replace(",", ".")) / 100.0
        brand_text = fixed[: m.start()].strip()
        brand_text = re.sub(r"\s+[aà]\s*$", "", brand_text, flags=re.IGNORECASE)
        brand = normalize_brand(brand_text)
        if brand:
            result[brand] = pct
    return result


def _lookup_cplus(brand: str, markup_map: Dict[str, float]) -> Optional[float]:
    key = normalize_brand(brand)
    if not key:
        return None
    if key in markup_map:
        return markup_map[key]
    key_alpha = re.sub(r"[^A-Z0-9]", "", key)
    for mk, val in markup_map.items():
        if re.sub(r"[^A-Z0-9]", "", mk) == key_alpha:
            return val
    first_word = re.sub(r"[^A-Z0-9]", "", key.split()[0]) if key else ""
    if first_word:
        for mk, val in markup_map.items():
            mk_first = re.sub(r"[^A-Z0-9]", "", mk.split()[0]) if mk else ""
            if mk_first == first_word:
                return val
    return None


# ---------------------------------------------------------------------------
# Column detection helpers
# ---------------------------------------------------------------------------


def _find_col(headers: List[str], *candidates: str) -> Optional[int]:
    for c in candidates:
        cup = c.strip().upper()
        for i, h in enumerate(headers):
            if h == cup:
                return i
    return None


def _col_text(df: pd.DataFrame, r: int, idx: Optional[int]) -> str:
    if idx is None or idx >= df.shape[1]:
        return ""
    return normalize_text(df.iat[r, idx])


def _color_value(
    df: pd.DataFrame, r: int, color_idx: Optional[int], desc_idx: Optional[int],
) -> str:
    desc = _col_text(df, r, desc_idx)
    if desc:
        return desc
    return _col_text(df, r, color_idx)


# ---------------------------------------------------------------------------
# v1 parser: 产品行 + TOTALE 标记行配对
# 通用版式:某行 cX 位置 = "TOTALE" 字符串 → 该行是合计行,其上方最近的产品行
# 是数据行。尺码 label 在产品行的某个 NUMERATE 区段,数量在 TOTALE 行同 col。
# 适用于 BOTTEGA VENETA 等意大利/法式品牌的"每产品 2-3 行"版式。
# 列偏移通过结构探测自动定位(不依赖具体列名是 FOTO/ART)。
# ---------------------------------------------------------------------------


# 已知的"产品标识列"候选(找 ART 货号列用,按优先级)
_ART_HEADER_CANDIDATES = ("ART", "ARTICOLO", "ARTICLE", "ARTICLE NR", "STYLE", "SKU", "MODELLO", "CODICE", "REF", "CODE")
# 已知的"成本列"候选
_COST_HEADER_CANDIDATES = ("COSTO", "COST PRICE", "COST", "P. ACQUISTO", "PR. ACQUISTO", "PRIX ACHAT", "P.COST")
# 已知的"零售列"候选
_RETAIL_HEADER_CANDIDATES = ("VENDITA", "RETAIL PRICE", "RETAIL", "P. VENDITA", "PR. VENDITA", "PRIX VENTE", "P.VENDITA")
# 已知的"描述列"候选
_DESCR_HEADER_CANDIDATES = ("DESCR", "DESCRIZIONE", "DESCRIPTION", "PRODUCT NAME", "NAME", "DESCR. COLORE")
# 已知的"材质列"候选
_MATERIAL_HEADER_CANDIDATES = ("MATERIALE", "MATERIAL", "MATIERE")
# 已知的"色码列"候选
_COLOR_HEADER_CANDIDATES = ("COLORE", "COLOR", "COULEUR")
# 已知的"色名列"候选(色名优先于色码)
_COLOR_DESC_HEADER_CANDIDATES = ("DESCR. COLORE", "COLOR DESCRIPTION", "COLOR NAME", "NOM COULEUR")
# 已知的"尺码段起始列"候选(意大利语 NUMERATE / 法式 TAILLES)
_SIZE_HEADER_CANDIDATES = ("NUMERATE", "TAILLES", "TAGLIE", "MISURE", "SIZES")


def _scan_header_candidates(df: pd.DataFrame, max_rows: int = 12) -> Optional[int]:
    """找 header 行:含任意 COSTO/COST/VENDITA/RETAIL 关键字的行(0-idx)"""
    for r in range(min(max_rows, df.shape[0])):
        for c in range(df.shape[1]):
            v = normalize_text(df.iat[r, c]).upper()
            for cand in _COST_HEADER_CANDIDATES + _RETAIL_HEADER_CANDIDATES:
                if v == cand:
                    return r
    return None


def _find_col_by_header(df: pd.DataFrame, header_row: int, candidates: Tuple[str, ...], uppercase: bool = True) -> Optional[int]:
    """在 header_row 找匹配 candidates 的列(精确匹配优先,大小写不敏感)"""
    upper_cands = tuple(c.upper() for c in candidates)
    # 精确匹配
    for c in range(df.shape[1]):
        v = normalize_text(df.iat[header_row, c])
        if uppercase:
            v = v.upper()
        if v in upper_cands:
            return c
    # 子串匹配
    for c in range(df.shape[1]):
        v = normalize_text(df.iat[header_row, c])
        if uppercase:
            v = v.upper()
        for cand in upper_cands:
            if cand in v:
                return c
    return None


def _detect_v1_total_marker(df: pd.DataFrame) -> Optional[Dict[str, object]]:
    """v1 检测:扫描文件找 TOTALE 标记列 + header 行映射 col。

    返回 None 表示不是 v1 版式。返回 dict 含:
      header_row, art_col, descr_col, material_col, color_col, color_desc_col,
      cost_col, retail_col, size_cols(list of col idx), totale_marker_col
    """
    header_row = _scan_header_candidates(df)
    if header_row is None:
        return None

    # 找 TOTALE 标记列:扫前 30 行,找某 col 的 cell 是 "TOTALE" 字符串。
    # 注意:header 里 'TOTALE' / 'TOTALE COSTO' 是列名(c25/c26 是合计数量/合计成本),
    #      不是行标记。真正的 TOTALE 行标记在数据区中部(c10 之类),要排除 header 已标 TOTALE 的 col。
    totale_marker_col: Optional[int] = None
    for c in range(df.shape[1]):
        hdr = normalize_text(df.iat[header_row, c]).upper() if header_row < df.shape[0] else ""
        if hdr in ("TOTALE", "TOTALE COSTO"):
            continue  # 跳过列名是 TOTALE 的 col(c25/c26 合计列)
        for r in range(header_row + 1, min(header_row + 30, df.shape[0])):
            v = df.iat[r, c] if c < df.shape[1] else None
            if normalize_text(v).upper() == "TOTALE":
                totale_marker_col = c
                break
        if totale_marker_col is not None:
            break
    if totale_marker_col is None:
        return None

    # 找 ART 货号列(在 TOTALE 标记列左侧,header_row 行找)
    art_col: Optional[int] = None
    for cand in _ART_HEADER_CANDIDATES:
        col = _find_col_by_header(df, header_row, (cand,))
        if col is not None and col < totale_marker_col:
            art_col = col
            break
    if art_col is None:
        return None  # 没 ART 列 → 不是 v1

    # 找其他字段列
    descr_col = _find_col_by_header(df, header_row, _DESCR_HEADER_CANDIDATES)
    material_col = _find_col_by_header(df, header_row, _MATERIAL_HEADER_CANDIDATES)
    cost_col = _find_col_by_header(df, header_row, _COST_HEADER_CANDIDATES)
    retail_col = _find_col_by_header(df, header_row, _RETAIL_HEADER_CANDIDATES)
    color_col = _find_col_by_header(df, header_row, _COLOR_HEADER_CANDIDATES)
    color_desc_col = _find_col_by_header(df, header_row, _COLOR_DESC_HEADER_CANDIDATES)

    # 找 NUMERATE/TAILLES 尺码区段起始列(header_row 行)
    size_start_col: Optional[int] = None
    for cand in _SIZE_HEADER_CANDIDATES:
        col = _find_col_by_header(df, header_row, (cand,))
        if col is not None and col < totale_marker_col:
            size_start_col = col
            break
    # 兜底:用 TOTALE 标记列 +1 (c10='TOTALE' 则 c11 起是尺码)
    if size_start_col is None:
        size_start_col = totale_marker_col + 1
    size_end_col = totale_marker_col  # NUMERATE 段止于 TOTALE 标记列(不含)

    size_cols = list(range(size_start_col, min(size_end_col, df.shape[1])))

    # 2026-07-30:v5 兜底 — size_cols 空时,说明 TOTALE 标记列在尺码区左侧
    # (e.g. FENDI 补货单 c10='TOTALE',c11..c21=NUMERATE 尺码,c32='TOTALE' 合计)
    # 这种情况:找 NUMERATE 列(在 TOTALE 标记列右侧,header_row 行) +
    #          下一个 TOTALE / TOTALE COSTO 列(也是合计列)
    if not size_cols:
        size_start_v5: Optional[int] = None
        for cand in _SIZE_HEADER_CANDIDATES:
            col = _find_col_by_header(df, header_row, (cand,))
            if col is not None and col > totale_marker_col:
                size_start_v5 = col
                break
        if size_start_v5 is not None:
            size_end_v5: Optional[int] = None
            for c in range(size_start_v5 + 1, df.shape[1]):
                hdr = (
                    normalize_text(df.iat[header_row, c]).upper()
                    if header_row < df.shape[0]
                    else ""
                )
                if hdr in ("TOTALE", "TOTALE COSTO"):
                    size_end_v5 = c
                    break
            if size_end_v5 is not None:
                size_cols = list(range(size_start_v5, size_end_v5))

    return {
        "header_row": header_row,
        "art_col": art_col,
        "descr_col": descr_col,
        "material_col": material_col,
        "color_col": color_col,
        "color_desc_col": color_desc_col,
        "cost_col": cost_col,
        "retail_col": retail_col,
        "size_cols": size_cols,
        "totale_marker_col": totale_marker_col,
    }


def _parse_v1_total_marker(
    df: pd.DataFrame,
    sheet_name: str,
    meta: Dict[str, object],
    sheet_images: Dict[str, Dict[int, List[Dict[str, object]]]],
    sheet_row_heights: Dict[str, Dict[int, float]],
    markup_map: Dict[str, float],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """v1 parser:产品行 + TOTALE 标记行配对。

    扫描所有 TOTALE 标记行 → 向上找最近产品行(art_col 非空字符串)→ 读产品行 + TOTALE 行。
    """
    rows: List[Dict[str, object]] = []
    header_row: int = meta["header_row"]  # type: ignore[assignment]
    art_col: int = meta["art_col"]  # type: ignore[assignment]
    descr_col: Optional[int] = meta["descr_col"]  # type: ignore[assignment]
    material_col: Optional[int] = meta["material_col"]  # type: ignore[assignment]
    color_col: Optional[int] = meta["color_col"]  # type: ignore[assignment]
    color_desc_col: Optional[int] = meta["color_desc_col"]  # type: ignore[assignment]
    cost_col: Optional[int] = meta["cost_col"]  # type: ignore[assignment]
    retail_col: Optional[int] = meta["retail_col"]  # type: ignore[assignment]
    size_cols: List[int] = meta["size_cols"]  # type: ignore[assignment]
    totale_marker_col: int = meta["totale_marker_col"]  # type: ignore[assignment]

    sheet_brand = normalize_text(sheet_name).upper() or "VENDOR"

    r = header_row + 1
    while r < df.shape[0]:
        # 找下一个 TOTALE 标记行(从 r 起)
        totale_row: Optional[int] = None
        for r2 in range(r, min(r + 30, df.shape[0])):
            v = df.iat[r2, totale_marker_col] if totale_marker_col < df.shape[1] else None
            if normalize_text(v).upper() == "TOTALE":
                totale_row = r2
                break
        if totale_row is None:
            break  # 后续没有更多 TOTALE 行

        # 向上找最近产品行(art_col 非空字符串)
        product_row: Optional[int] = None
        for r3 in range(totale_row - 1, max(header_row, totale_row - 10), -1):
            if art_col < df.shape[1]:
                art_v = df.iat[r3, art_col]
                if isinstance(art_v, str) and art_v.strip():
                    product_row = r3
                    break

        if product_row is None:
            r = totale_row + 1
            continue

        # 读产品行
        art_raw = normalize_text(df.iat[product_row, art_col])
        # 2026-08-14:SKU 拼接符统一 `-`(源可能用 `.` 或 `_`,改 `-`)
        # 例: "1487WT.9846" → "1487WT-9846"
        art = art_raw.replace(".", "-").replace("_", "-")
        if not art:
            r = totale_row + 1
            continue
        desc = _col_text(df, product_row, descr_col) if descr_col is not None else ""
        material = _col_text(df, product_row, material_col) if material_col is not None else ""
        color_code = _col_text(df, product_row, color_col) if color_col is not None else ""
        color_name = _col_text(df, product_row, color_desc_col) if color_desc_col is not None else ""
        color = color_name or color_code
        cost = to_float(df.iat[product_row, cost_col]) if cost_col is not None and cost_col < df.shape[1] else None
        retail = to_float(df.iat[product_row, retail_col]) if retail_col is not None and retail_col < df.shape[1] else None

        cplus = _lookup_cplus(sheet_brand, markup_map)
        supply = calc_supply_price(cost, cplus)
        excel_row = product_row + 1
        # BOTTEGA VENETA 等模板图锚在产品行下方 1 行(空行),SALOMON/GUCCI 等在产品行本身。
        # 兼容两种:查 product_row (1-idx) 和 product_row+1 都收集。
        row_images = (
            sheet_images.get(sheet_name, {}).get(excel_row, [])
            + sheet_images.get(sheet_name, {}).get(excel_row + 1, [])
        )

        # 按尺码展开:产品行 size_cols = 尺码 label,TOTALE 行同 col = 数量
        any_size_emitted = False
        for c in size_cols:
            if c >= df.shape[1]:
                break
            label = _col_text(df, product_row, c)
            if not label:
                continue
            qty_raw = df.iat[totale_row, c] if c < df.shape[1] else None
            if not is_positive_qty(qty_raw):
                continue
            qty = int(round(float(to_float(qty_raw) or 0)))
            if qty <= 0:
                continue
            any_size_emitted = True
            rows.append({
                "brand": sheet_brand,
                "sku": art,
                "image": "",
                "description": desc,
                "category": guess_product_group(desc),
                "material": material,
                "color": color,
                "color_code": color_code,  # 2026-07-30:色号(原 c6 COLORE,独立保留供中文模板"色号"列)
                "season": "",
                "origin": "",
                "cost": cost,
                "cplus_text": format_cplus_text(cplus),
                "supply_price": supply,
                "retail_price": retail,
                "size": label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": sheet_row_heights.get(sheet_name, {}).get(excel_row),
            })

        # 兜底:产品行 size_cols 全空但 TOTALE 行 c25 合计数 > 0(很多 'U' 通用码产品)
        if not any_size_emitted:
            # 找 TOTALE 之后的合计列(右侧第一个正数)
            for c in range(totale_marker_col + 1, df.shape[1]):
                qty_raw = df.iat[totale_row, c]
                if is_positive_qty(qty_raw):
                    qty = int(round(float(to_float(qty_raw) or 0)))
                    if qty > 0:
                        size_label = "U"
                        for sc in size_cols:
                            if sc < df.shape[1]:
                                lbl = _col_text(df, product_row, sc)
                                if lbl:
                                    size_label = lbl
                                    break
                        rows.append({
                            "brand": sheet_brand,
                            "sku": art,
                            "image": "",
                            "description": desc,
                            "category": guess_product_group(desc),
                            "material": material,
                            "color": color,
                            "season": "",
                            "origin": "",
                            "cost": cost,
                            "cplus_text": format_cplus_text(cplus),
                            "supply_price": supply,
                            "retail_price": retail,
                            "size": size_label,
                            "quantity": qty,
                            "vendor_id": vendor_id,
                            "image_items": row_images,
                            "source_row_height": sheet_row_heights.get(sheet_name, {}).get(excel_row),
                        })
                        break

        r = totale_row + 1

    return rows


# ---------------------------------------------------------------------------
# Flat sheet parser (CHLOE, MAISON MARGIELA, YSL)
# ---------------------------------------------------------------------------


def _parse_flat(
    df: pd.DataFrame,
    headers: List[str],
    sheet_name: str,
    sheet_images: Dict[str, Dict[int, List[Dict[str, object]]]],
    sheet_row_heights: Dict[str, Dict[int, float]],
    markup_map: Dict[str, float],
    vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    brand_idx = _find_col(headers, "BRAND")
    sku_idx = _find_col(headers, "ARTICOLO", "ARTICLE NR", "STYLE MATERIAL COLOR")
    desc_idx = _find_col(headers, "DESCRIZIONE", "DESCRIPTION", "PRODUCT NAME", "NAME")
    material_idx = _find_col(headers, "MATERIALE", "MATERIAL")
    color_idx = _find_col(headers, "COLORE", "COLOR")
    cost_idx = _find_col(headers, "COSTO", "COST PRICE", "COST")
    retail_idx = _find_col(headers, "VENDITA", "RETAIL PRICE", "RETAIL")
    qty_idx = _find_col(headers, "Q.T\u00c0", "Q.TA", "PEZZI", "TOT PZ", "QTY", "TOT PEZZI", "QTY PEZZI")
    # 2026-07-28:加 size 列识别 — 部分模板(如 PRADA DONNA BORSE)c13 = 'Size/Qty' 复合 'U/5',
    # 拆分出 size='U' 单独写 size 列;空的话默认 ONE SIZE。
    size_idx = _find_col(headers, "SIZE/QTY", "TAGLIA", "TAGLIE", "SIZE", "MISURA")
    season_idx = _find_col(headers, "STAGIONE", "SEASON")

    color_desc_idx = _find_col(headers, "DESCR.COLORE", "DESCRIZIONE COLORE", "COLOR DESCRIPTION")
    if color_desc_idx is None and color_idx is not None and color_idx + 1 < len(headers):
        if not headers[color_idx + 1]:
            color_desc_idx = color_idx + 1

    material_desc_idx: Optional[int] = None
    if material_idx is not None and material_idx + 1 < len(headers):
        if not headers[material_idx + 1]:
            material_desc_idx = material_idx + 1

    sheet_brand = normalize_text(sheet_name)

    for r in range(1, df.shape[0]):
        brand = _col_text(df, r, brand_idx) if brand_idx is not None else sheet_brand
        if not brand or brand.upper() in {"BRAND", ""}:
            continue
        sku = _col_text(df, r, sku_idx)
        if not sku:
            continue
        qty_raw = df.iat[r, qty_idx] if qty_idx is not None and qty_idx < df.shape[1] else None
        if not is_positive_qty(qty_raw):
            continue
        qty = int(round(float(to_float(qty_raw) or 0)))
        cost = to_float(df.iat[r, cost_idx]) if cost_idx is not None and cost_idx < df.shape[1] else None
        retail = to_float(df.iat[r, retail_idx]) if retail_idx is not None and retail_idx < df.shape[1] else None
        desc = _col_text(df, r, desc_idx)
        color = _color_value(df, r, color_idx, color_desc_idx)
        material = _col_text(df, r, material_idx)
        if material_desc_idx is not None:
            md = _col_text(df, r, material_desc_idx)
            if md:
                material = md if not material else f"{material} {md}"
        season = _col_text(df, r, season_idx)
        # 2026-07-28:从 size_idx 拆 size(PRADA DONNA 'U/5' → 'U')fallback 'ONE SIZE'
        if size_idx is not None and size_idx < df.shape[1]:
            size_raw = _col_text(df, r, size_idx)
            size = size_raw.split("/")[0].strip() if size_raw else ""
        else:
            size = ""
        if not size:
            size = "ONE SIZE"
        cplus = _lookup_cplus(brand, markup_map)
        supply = calc_supply_price(cost, cplus)
        excel_row = r + 1
        row_images = sheet_images.get(sheet_name, {}).get(excel_row, [])
        rows.append(
            {
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": desc,
                "category": guess_product_group(desc),
                "material": material,
                "color": color,
                "season": season,
                "origin": "",
                "cost": cost,
                "cplus_text": format_cplus_text(cplus),
                "supply_price": supply,
                "retail_price": retail,
                "size": size,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": sheet_row_heights.get(sheet_name, {}).get(excel_row),
            }
        )
    return rows


# ---------------------------------------------------------------------------
# BURBERRY-style: TAGLIE size-code column, single size per row
# ---------------------------------------------------------------------------


def _parse_size_single(
    df: pd.DataFrame,
    headers: List[str],
    sheet_name: str,
    sheet_images: Dict[str, Dict[int, List[Dict[str, object]]]],
    sheet_row_heights: Dict[str, Dict[int, float]],
    markup_map: Dict[str, float],
    vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    taglie_idx = _find_col(headers, "TAGLIE")
    if taglie_idx is None:
        return rows
    size_label_col = taglie_idx + 1

    brand_idx = _find_col(headers, "BRAND")
    sku_idx = _find_col(headers, "ARTICOLO", "ARTICLE NR")
    desc_idx = _find_col(headers, "NAME", "DESCRIZIONE", "DESCRIPTION")
    material_idx = _find_col(headers, "MATERIALE", "MATERIAL")
    color_idx = _find_col(headers, "COLORE", "COLOR")
    cost_idx = _find_col(headers, "COSTO", "COST PRICE")
    retail_idx = _find_col(headers, "VENDITA", "RETAIL PRICE")
    season_idx = _find_col(headers, "STAGIONE", "SEASON")

    color_desc_idx: Optional[int] = None
    if color_idx is not None and color_idx + 1 < len(headers):
        if not headers[color_idx + 1]:
            color_desc_idx = color_idx + 1

    size_code_map: Dict[str, str] = {}
    data_start = 1
    for r in range(1, min(20, df.shape[0])):
        brand_val = _col_text(df, r, brand_idx) if brand_idx is not None else ""
        if brand_val and brand_val.upper() not in {"", "BRAND", "STAGIONE"}:
            data_start = r
            break
        code = _col_text(df, r, taglie_idx)
        if len(code) == 1 and code.isalpha():
            label = _col_text(df, r, size_label_col)
            if label:
                size_code_map[code.upper()] = label
            data_start = r + 1

    for r in range(data_start, df.shape[0]):
        brand = _col_text(df, r, brand_idx) if brand_idx is not None else ""
        if not brand:
            continue
        sku = _col_text(df, r, sku_idx)
        if not sku:
            continue
        code = _col_text(df, r, taglie_idx).upper()
        size = size_code_map.get(code, code)
        qty_raw = df.iat[r, size_label_col] if size_label_col < df.shape[1] else None
        if not is_positive_qty(qty_raw):
            continue
        qty = int(round(float(to_float(qty_raw) or 0)))
        cost = to_float(df.iat[r, cost_idx]) if cost_idx is not None else None
        retail = to_float(df.iat[r, retail_idx]) if retail_idx is not None else None
        color = _color_value(df, r, color_idx, color_desc_idx)
        desc = _col_text(df, r, desc_idx)
        material = _col_text(df, r, material_idx)
        season = _col_text(df, r, season_idx)
        # 2026-07-28:从 size_idx 拆 size(PRADA DONNA 'U/5' → 'U')fallback 'ONE SIZE'
        if size_idx is not None and size_idx < df.shape[1]:
            size_raw = _col_text(df, r, size_idx)
            size = size_raw.split("/")[0].strip() if size_raw else ""
        else:
            size = ""
        if not size:
            size = "ONE SIZE"
        cplus = _lookup_cplus(brand, markup_map)
        supply = calc_supply_price(cost, cplus)
        excel_row = r + 1
        row_images = sheet_images.get(sheet_name, {}).get(excel_row, [])
        rows.append(
            {
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": desc,
                "category": guess_product_group(desc),
                "material": material,
                "color": color,
                "season": season,
                "origin": "",
                "cost": cost,
                "cplus_text": format_cplus_text(cplus),
                "supply_price": supply,
                "retail_price": retail,
                "size": size,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": sheet_row_heights.get(sheet_name, {}).get(excel_row),
            }
        )
    return rows


# ---------------------------------------------------------------------------
# MAX MARA THE CUBE-style: size-code with matrix of quantities
# ---------------------------------------------------------------------------


def _parse_size_matrix(
    df: pd.DataFrame,
    headers: List[str],
    code_col: int,
    sheet_name: str,
    sheet_images: Dict[str, Dict[int, List[Dict[str, object]]]],
    sheet_row_heights: Dict[str, Dict[int, float]],
    markup_map: Dict[str, float],
    vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    size_start = code_col + 1
    qty_total_col = _find_col(headers, "QTY", "TOTAL", "TOTALE")
    size_end = (qty_total_col - 1) if qty_total_col is not None else (df.shape[1] - 3)

    brand_idx = _find_col(headers, "BRAND")
    sku_idx = _find_col(headers, "ARTICLE NR", "ARTICOLO", "STYLE #", "STYLE")
    color_idx = _find_col(headers, "COLOR", "COLORE")
    color_desc_idx = _find_col(headers, "COLOR DESCRIPTION", "DESCR.COLORE")
    cost_idx = _find_col(headers, "COST PRICE", "COSTO")
    retail_idx = _find_col(headers, "RETAIL PRICE", "VENDITA")

    data_start = None
    for r in range(1, df.shape[0]):
        bv = _col_text(df, r, brand_idx) if brand_idx is not None else ""
        if bv and bv.upper() not in {"", "BRAND"}:
            data_start = r
            break
    if data_start is None:
        return rows

    size_maps: Dict[str, Dict[int, str]] = {}
    for r in range(0, data_start):
        code = normalize_text(df.iat[r, code_col]).upper() if code_col < df.shape[1] else ""
        if len(code) != 1 or not code.isalpha():
            continue
        col_map: Dict[int, str] = {}
        for c in range(size_start, min(size_end + 1, df.shape[1])):
            label = normalize_text(df.iat[r, c])
            if label:
                col_map[c] = label
        if col_map:
            size_maps[code] = col_map

    for r in range(data_start, df.shape[0]):
        brand = _col_text(df, r, brand_idx) if brand_idx is not None else ""
        if not brand:
            continue
        sku = _col_text(df, r, sku_idx)
        if not sku:
            continue
        # 2026-08-14:SKU 拼接符统一 `-`(源可能用 `.` 或 `_`,改 `-`)
        sku = sku.replace(".", "-").replace("_", "-")
        code = normalize_text(df.iat[r, code_col]).upper() if code_col < df.shape[1] else ""
        col_map = size_maps.get(code, {})
        cost = to_float(df.iat[r, cost_idx]) if cost_idx is not None else None
        retail = to_float(df.iat[r, retail_idx]) if retail_idx is not None else None
        color = _color_value(df, r, color_idx, color_desc_idx)
        cplus = _lookup_cplus(brand, markup_map)
        supply = calc_supply_price(cost, cplus)
        excel_row = r + 1
        row_images = sheet_images.get(sheet_name, {}).get(excel_row, [])

        for c, size_label in col_map.items():
            if c >= df.shape[1]:
                continue
            qty_raw = df.iat[r, c]
            if not is_positive_qty(qty_raw):
                continue
            rows.append(
                {
                    "brand": brand,
                    "sku": sku,
                    "image": "",
                    "description": "",
                    "category": "",
                    "material": "",
                    "color": color,
                    "season": "",
                    "origin": "",
                    "cost": cost,
                    "cplus_text": format_cplus_text(cplus),
                    "supply_price": supply,
                    "retail_price": retail,
                    "size": size_label,
                    "quantity": int(round(float(to_float(qty_raw) or 0))),
                    "vendor_id": vendor_id,
                    "image_items": row_images,
                    "source_row_height": sheet_row_heights.get(sheet_name, {}).get(excel_row),
                }
            )
    return rows


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# POLO RALPH LAUREN ATS 版式 (2026-10-30 新增)
# ---------------------------------------------------------------------------
# 源文件: POLO RALPH LAUREN MAN FW27.xlsx (单 sheet 'GRIGLIA', 18671 行 × 41 列)
#
# 结构:
#   R1  : 表头(41 列) —— MATERIAL NUMBER | IMAGE | SEASON CODE | MERCH DIVISION
#         | SHORT DIV | GLOBAL L4 | BOARD | SHORT BOARD | MODEL NAME
#         | MODEL LONG DESCRIPTION | Merch Fabrication | FABRIC DESCRIPTION
#         | COLOR DESCRIPTION | PRODUCT CLASS | PRODUCT SUBCLASS | ... | LABEL
#         | SEASONALITY | SENSIBILITY | Utility Icons ×6 | MSRP (EUR)
#         | WHSL (EUR) | SIZE | Q.TY | TOTAL ORDER
#   R2  : 空
#   R3+ : 数据行 —— **一行一个尺码**, 同一货号(MATERIAL NUMBER)连续多行
#
# 图片(richValue):
#   IMAGE 列(c2)是 Excel 365 的**单元格内嵌图**(richValue 机制, 没有 drawing anchor),
#   而且**一个货号只有第一个尺码行有图**, 其余尺码行的 c2 是 '-' 文本。
#   → 转换时把块首行的图**共享给该货号的所有尺码行**(用户 2026-10-30 明确要求)。
#
# 字段映射:
#   MATERIAL NUMBER     → 货号
#   IMAGE               → 图片(同货号块内共享)
#   MODEL LONG DESCRIPTION / MODEL NAME → 描述
#        ⚠️ 这两列在源里**内容互换**: 有时 c9 是款号、c10 是描述, 有时反过来。
#           规则: c10 含空格 → 取 c10, 否则取 c9。
#   GLOBAL L4           → 类目
#   FABRIC DESCRIPTION  → 材质
#   COLOR DESCRIPTION   → 颜色
#   MSRP (EUR)          → 零售价
#   WHSL (EUR)          → cost
#   SIZE                → 尺码(原样, 含 '30  32' / '14/ 23' 这种复合码)
#   Q.TY                → 数量(源文件整列全空 → 留空等客户填)
#   (源没有色号列 / 没有产地列 → 留空)
#
# 品牌 / 性别 / 季节:
#   源文件没有品牌列 → 从文件名拆('POLO RALPH LAUREN MAN FW27')
#   性别: MERCH DIVISION 前缀 'M PRL' → MAN / 'F PRL' → WOMAN; '-' 的行用文件级性别
#   季节: 从文件名拆('FW27'); 文件名没有就退回源 c3 SEASON CODE 原值

_SEASON_TOKEN_RE = re.compile(r"\b(FW|SS|AW|CW)\s*[-_]?\s*(\d{2})\b", re.IGNORECASE)
_GENDER_TOKEN_RE = re.compile(
    r"\b(MAN|MEN|MENS|WOMAN|WOMEN|WOMENS|WOM)\b", re.IGNORECASE
)


def _split_filename_meta(stem: str) -> Tuple[str, str, str]:
    """从文件名拆 (品牌, 性别, 季节)。

    'POLO RALPH LAUREN MAN FW27' → ('POLO RALPH LAUREN', 'MAN', 'FW27')

    ⚠️ App 上传时走 ``werkzeug.secure_filename``, 空格会变成下划线
       ('POLO_RALPH_LAUREN_MAN_FW27'), 所以先统一把 '_' 还原成空格。
    """
    s = stem.replace("_", " ")
    season = ""
    m = _SEASON_TOKEN_RE.search(s)
    if m:
        season = f"{m.group(1).upper()}{m.group(2)}"
        s = s[:m.start()] + " " + s[m.end():]
    gender = ""
    m = _GENDER_TOKEN_RE.search(s)
    if m:
        tok = m.group(1).upper()
        gender = "WOMAN" if tok.startswith("WOM") else "MAN"
        s = s[:m.start()] + " " + s[m.end():]
    brand = re.sub(r"\s+", " ", s).strip(" -_")
    return brand, gender, season


def _parse_prl_material_number(
    ws, primary_path: Path, vendor_id: str
) -> List[Dict[str, object]]:
    """POLO RL ATS 'MATERIAL NUMBER' 版式(一行一个尺码 + richValue 块内共享图)。"""
    rows: List[Dict[str, object]] = []
    images_by_row = load_richvalue_images_by_col(primary_path, "B")
    brand, file_gender, season = _split_filename_meta(primary_path.stem)

    cur_sku = ""
    cur_images: List[Dict[str, object]] = []
    for r, row in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
        if not row or row[0] is None:
            continue
        sku = str(row[0]).strip()
        if not sku:
            continue

        # ---- 图片: 同一个货号连续多行, 只有块首行有图 → 块内共享 ----
        if sku != cur_sku:
            cur_sku = sku
            cur_images = images_by_row.get(r, [])
        if not cur_images:
            cand = images_by_row.get(r, [])
            if cand:
                cur_images = cand

        def _t(idx: int) -> str:
            v = row[idx] if idx < len(row) else None
            return str(v).strip() if v is not None else ""

        model_name = _t(8)      # MODEL NAME
        model_desc = _t(9)      # MODEL LONG DESCRIPTION
        # 这两列内容在源里互换 → c10 含空格才算描述
        description = model_desc if " " in model_desc else model_name

        merch = _t(3).upper()    # MERCH DIVISION
        if merch.startswith("M PRL"):
            gender = "MAN"
        elif merch.startswith("F PRL"):
            gender = "WOMAN"
        else:
            gender = file_gender

        qty = to_float(row[39] if len(row) > 39 else None)   # Q.TY
        rows.append({
            "brand": brand,
            "sku": sku,
            "image": "",
            "description": description,
            "category": _t(5),            # GLOBAL L4
            "material": _t(11),           # FABRIC DESCRIPTION
            "color": _t(12),              # COLOR DESCRIPTION
            "color_code": "",             # 源文件没有色号列
            "gender": gender,
            "season": season or _t(2),    # SEASON CODE(文件名没季节时兜底)
            "origin": "",                 # 源文件没有产地列
            "cost": to_float(row[37] if len(row) > 37 else None),   # WHSL (EUR)
            "cplus_text": "",             # 由 dispatcher 的 --cplus / 规则填
            "supply_price": None,         # 同上
            "retail_price": to_float(row[36] if len(row) > 36 else None),  # MSRP (EUR)
            "size": _t(38),               # SIZE
            "quantity": int(round(qty)) if qty is not None else None,
            "vendor_id": vendor_id,
            "image_items": cur_images,
        })
    return rows


def parse_rows(primary_path: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    base_dir = primary_path if primary_path.is_dir() else primary_path.parent
    markup_map = _build_markup_map(base_dir / "cost+.txt")
    all_rows: List[Dict[str, object]] = []
    xlsx_files = sorted(base_dir.glob("*.xlsx")) if primary_path.is_dir() else [primary_path]

    for xlsx_file in xlsx_files:
        sheet_row_heights = load_sheet_row_heights(xlsx_file)
        # 2026-09-03 接入:GRIGLIA sub-format detect (CP COMPANY ATS 现货)
        # R1 c2 = 'Style Number' 字符串 → 走新 GRIGLIA parser
        # 用 load_sheet_images_by_row (简单 anchor row 索引,不用 SALOMON-style 慢解析)
        try:
            wb_probe = openpyxl.load_workbook(xlsx_file, read_only=True, data_only=True)
            handled = False
            for sname in wb_probe.sheetnames:
                ws_probe = wb_probe[sname]
                r1c2 = ws_probe.cell(1, 2).value
                if (
                    r1c2 is not None
                    and isinstance(r1c2, str)
                    and r1c2.strip().lower() == "style number"
                ):
                    # load_sheet_images_by_row 是**非 read_only** 加载(整份进内存),
                    # 只在真的命中 GRIGLIA 版式时才付这个代价
                    # (POLO RL ATS 那个 30MB / 18671 行的文件扛不住)
                    griglia_images = load_sheet_images_by_row(xlsx_file)
                    rows = _parse_griglia(ws_probe, vendor_id, griglia_images)
                    if rows:
                        all_rows.extend(rows)
                    handled = True
                    break
            if not handled:
                # 2026-10-30 新增: POLO RL ATS 'MATERIAL NUMBER' 版式(richValue 内嵌图)
                for sname in wb_probe.sheetnames:
                    ws_probe = wb_probe[sname]
                    r1 = next(
                        ws_probe.iter_rows(min_row=1, max_row=1, values_only=True), ()
                    )
                    if not r1:
                        continue
                    h1 = str(r1[0]).strip().upper() if r1[0] is not None else ""
                    h2 = (
                        str(r1[1]).strip().upper()
                        if len(r1) > 1 and r1[1] is not None
                        else ""
                    )
                    if h1 == "MATERIAL NUMBER" and h2 == "IMAGE":
                        sheet_rows = _parse_prl_material_number(
                            ws_probe, xlsx_file, vendor_id
                        )
                        if sheet_rows:
                            all_rows.extend(sheet_rows)
                        handled = True
                        break
            wb_probe.close()
            if handled:
                continue
        except Exception:
            pass
        # 老 FOLLIFOLLIE 版式走原路径
        sheet_images = load_sheet_embedded_images(xlsx_file)
        for sheet_name in pd.ExcelFile(xlsx_file).sheet_names:
            df = pd.read_excel(xlsx_file, sheet_name=sheet_name, header=None)
            if df.empty or df.shape[0] < 2:
                continue

            # v1: 产品行 + TOTALE 标记行配对(意大利/法式品牌常见版式,如 BOTTEGA VENETA)
            #     优先尝试,失败才走 v2(原 _parse_flat / _parse_size_matrix / _parse_size_single)
            v1_meta = _detect_v1_total_marker(df)
            if v1_meta is not None:
                sheet_rows = _parse_v1_total_marker(
                    df, sheet_name, v1_meta,
                    sheet_images, sheet_row_heights, markup_map, vendor_id,
                )
                if sheet_rows:
                    all_rows.extend(sheet_rows)
                    continue
                # v1 检测成功但解析出 0 行 → 走 v2 兜底

            headers = [normalize_text(df.iat[0, c]).upper() for c in range(df.shape[1])]

            size_matrix_col: Optional[int] = None
            for i, h in enumerate(headers):
                if len(h) == 1 and h in "ABCDEF" and i >= 7:
                    if i + 1 < len(headers) and headers[i + 1]:
                        nxt = headers[i + 1]
                        if to_float(nxt) is not None or nxt in COMMON_SIZE_HEADERS:
                            size_matrix_col = i
                            break

            has_taglie = any("TAGLIE" in h for h in headers)

            if size_matrix_col is not None:
                sheet_rows = _parse_size_matrix(
                    df, headers, size_matrix_col, sheet_name,
                    sheet_images, sheet_row_heights, markup_map, vendor_id,
                )
            elif has_taglie:
                sheet_rows = _parse_size_single(
                    df, headers, sheet_name,
                    sheet_images, sheet_row_heights, markup_map, vendor_id,
                )
            else:
                sheet_rows = _parse_flat(
                    df, headers, sheet_name,
                    sheet_images, sheet_row_heights, markup_map, vendor_id,
                )
            all_rows.extend(sheet_rows)

    return all_rows
