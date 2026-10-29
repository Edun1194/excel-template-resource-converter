"""V1072-期货 vendor adapter — 8 个品牌 "FW26 IN ARRIVO" (2026-09-03 接入)。

**独立 vendor** — 不跟 V1072 现货 (BURBERRY) 混:
- V1072 现货 → `scripts/vendor_v1072.py`(BURBERRY 块状,4 行一组)
- V1072-期货 → `scripts/vendor_v1072_in_arrivo.py`(本文件,8 个品牌多版式)

适用:LISTE FW26 IN ARRIVO 目录下的 8 个 xlsx
- D8_EMPORIO ARMANI fw26.xlsx
- I6 / L8 MAGDA BUTRYN / MAGDA BUTRYM.xlsx
- O8_AMIRI / O8_CASABLANCA / O8_OFFWHITE / O8_PALM ANGELS / O8_RHUDE FW26.xlsx

**两种版式**(本 adapter 都支持):

版式 A (6 个,块状):size header 行 + data 行(每行展开 size 矩阵)
- 找 anchor:某行 c22='Size'
- size header 行 = anchor+1 (c22='L'/'N1'/'G'/'U' + c23+ 尺码 labels)
- data 行 = size header 之后 c3 是字符串 (Code) 的所有行
- 字段: c3=Code, c8=Category, c11=Descriptions, c17=Color
- 价格: c45=Unit Rtl, c48=Unit Whls
- size qty: c23+ (与 size header 同列,下一个 qty 行)

版式 B (2 个 MAGDA BUTRYN,单行单尺码):R1 header + R2+ data
- 字段: c2=Brand, c3=Descr_Article, c4=Season, c5=Article_Code, c6=Gender,
       c7=Color, c8=MadeIn, c9=Composizione, c10=CodiceDoganale
- size qty: c11=Size, c12=Qty
- 价格: c13=Cost, c14=Retail (I6) / 无 (L8)

**自动 detect**:
- 扫 R1-R8,c1='Brand' + c3='Descr_Article' 字符串匹配 → B 类
- 否则扫 c22='Size' → A 类
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    normalize_text,
    to_float,
)


def _to_float_eur(v) -> Optional[float]:
    """欧式价格 '668,82' → 668.82。空/非数字 → None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if v > 0 else None
    s = str(v).strip().replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        f = float(s)
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


def _is_size_token(v) -> bool:
    """size header label:数字 / 小数 / 短字符串(XXS/S/M 等)"""
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        sv = v.strip()
        if not sv:
            return False
        if re.fullmatch(r"\d+(?:[.,]\d+)?", sv):
            return True
        if sv in ("UNI", "TU", "Unica", "ONE SIZE", "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
                  "3XL", "4XL", "5XL", "6XL"):
            return True
    return False


def _format_size_label(v) -> str:
    """size label 归一:数字 → int 字符串,字符串 trim"""
    if isinstance(v, (int, float)):
        f = float(v)
        if f.is_integer():
            return str(int(f))
        return str(f).replace(",", ".")
    return str(v).strip()


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 0-idx row 分桶图。

    2026-09-03 修正:**只保留 A 列(0-idx col 0 = 1-idx c1)的图**。
    源文件 B 类/MAGDA 把图片放 c2 附近(col 1),A 类 D8_EMPORIO 等是 A 列。
    但 user 反馈 "图片有问题,采用 A 列的图" → 强制 A 列。
    """
    out: Dict[int, List[Dict]] = {}
    for img in ws._images:
        anchor = img.anchor
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        # 强制 A 列:col_idx == 0
        if anchor._from.col != 0:
            continue
        try:
            data = img._data() if hasattr(img, "_data") else None
        except Exception:
            data = None
        if data is None:
            continue
        out.setdefault(anchor._from.row, []).append({
            "bytes": data,
            "format": img.format or "jpeg",
        })
    return out


def _find_image_in_block(
    images_by_row: Dict[int, List[Dict]],
    prev_data_row: Optional[int],
    current_data_row: int,
) -> List[Dict]:
    """"块"匹配:图必须落在 [prev_data_row, current_data_row + 1) 范围内。

    为什么不用 ±N 容差:源文件里"缺图"的 data 行真没自己的图
    (e.g. D8_EMPORIO row 9 'EM001717 AF28047 U8073' 源文件就没图),
    用 ±2 容差会把下一个商品的图(EM004541 AF10013 M7316 的图 row 11)
    错借给 EM001717。

    块匹配语义:图 anchor row 在"上一个 data row 之后 ~ 当前 data 行(含)"
    范围,才属于当前 data row。半开区间 `current_data_row + 1` 排除下一块。

    对 D8_EMPORIO 各 data row 验证:
    - row 6 'EM006736':prev=None, 范围 (-∞, 7)→img row 5 ✓
    - row 9 'EM001717':prev=6, [6, 10)→空 ✓ (源没图)
    - row 11 'EM004541':prev=9, [9, 12)→img row 11 ✓
    """
    if not images_by_row:
        return []
    lo = (prev_data_row + 1) if prev_data_row is not None else -10**9
    hi = current_data_row + 1  # 含 current_data_row 自身(0-idx 同行图)
    # 在范围内收集所有图,按 row 升序
    matched = []
    for r in sorted(images_by_row.keys()):
        if lo <= r < hi and images_by_row[r]:
            matched.extend(images_by_row[r])
    return matched


def _find_image_near_row(
    images_by_row: Dict[int, List[Dict]],
    row: int,
) -> List[Dict]:
    """B 类(单行单尺码)用 ±1 容差:B 类是 1 行 1 商品,图基本在同 row。

    容差 ±1 防止源文件图行和 data 行有微小偏移。
    """
    if not images_by_row:
        return []
    for offset in (0, -1, 1, -2, 2):
        target = row + offset
        if target in images_by_row and images_by_row[target]:
            return images_by_row[target]
    return []


def _format_sku_split(sku_raw: str) -> Tuple[str, str]:
    """V1072-期货 SKU 拆分: 末段 = color, 前面 = style。
    按空格 split(最多 1 次取末段, 前面是 style) → `<STYLE>-<COLOR>` 格式。
    与 V1072 现货的 `---` 分隔不同(那是 BURBERRY 风格)。

    例:
      "AMBMJE1119  097"    → ("AMBMJE1119-097",    "097")
      "44MAA120F26J00D  001" → ("44MAA120F26J00D-001", "001")
      "EM006736 AF29202 FC368" → ("EM006736 AF29202-FC368", "FC368")
      "241925BLACK"          → ("241925BLACK", "")  # 无空格不分
      "1327260000069RED"     → ("1327260000069RED", "")
    """
    if not sku_raw:
        return "", ""
    sku = str(sku_raw).strip()
    if not sku:
        return "", ""
    if " " in sku:
        parts = sku.split(" ")
        # 末段是 color
        color_code = parts[-1].strip()
        style = " ".join(parts[:-1]).strip()
        if color_code and style:
            return f"{style}-{color_code}", color_code
    # 无空格或末段空,保留原样
    return sku, ""


def _parse_format_b(ws, images_by_row) -> List[Dict[str, object]]:
    """版式 B (MAGDA BUTRYN):R1 header + R2+ data(单行单尺码)"""
    rows: List[Dict[str, object]] = []
    brand_default = normalize_text(ws.cell(2, 2).value)  # MAGDA BUTRYM
    for r in range(2, ws.max_row + 1):  # R2+ (R1 is header)
        # 1-idx 列: 2=Brand, 3=Descr, 4=Season, 5=Code, 6=Gender, 7=Color, 8=MadeIn,
        # 9=Composizione, 10=CodiceDoganale, 11=Size, 12=Qty, 13=Cost, 14=Retail
        code = normalize_text(ws.cell(r, 5).value)
        if not code:
            continue
        size_label = _format_size_label(ws.cell(r, 11).value)
        if not size_label:
            continue
        qty_v = ws.cell(r, 12).value
        qty = 0
        if qty_v is not None:
            try:
                qty = int(round(float(qty_v)))
            except (TypeError, ValueError):
                qty = 0
        if qty <= 0:
            continue  # B 类也不输出 0 库存(单行单尺码,0=无货)

        brand = normalize_text(ws.cell(r, 2).value) or brand_default
        desc = normalize_text(ws.cell(r, 3).value)
        season = normalize_text(ws.cell(r, 4).value)
        gender = normalize_text(ws.cell(r, 6).value)
        color = normalize_text(ws.cell(r, 7).value)
        origin = normalize_text(ws.cell(r, 8).value)
        material = normalize_text(ws.cell(r, 9).value)
        cost = _to_float_eur(ws.cell(r, 13).value)
        retail = _to_float_eur(ws.cell(r, 14).value)  # I6 有, L8 没有

        # B 类 SKU 拼接(2026-09-03 user 规则):color 优先,都没有不拼
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        if color_normalized:
            sku_final = f"{code}-{color_normalized}"
        else:
            sku_final = code  # color 空 → 不拼(用原 code)

        rows.append({
            "brand": brand,
            "sku": sku_final,
            "image": "",
            "description": desc,
            "category": desc,  # 描述当类目(B 类没有独立类目字段)
            "material": material,
            "color": color,
            "color_code": "",  # 源没短码色号
            "gender": gender,  # Donna → WOMAN (framework normalize)
            "season": season,
            "origin": origin,
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            "size": size_label,
            "quantity": qty,
            "vendor_id": "V1072",
            "image_items": _find_image_near_row(images_by_row, r - 1),
        })
    return rows


def _parse_format_a(ws, images_by_row) -> List[Dict[str, object]]:
    """版式 A (块状,6 个品牌):
    - 找 c22='Size' 标记行 → size header 行 = anchor+1
    - data 行 = size header 之后 c3 是字符串的所有行
    - 字段: c3=Code, c8=Category, c11=Descriptions, c17=Color
    - 价格: c45=Unit Rtl, c48=Unit Whls
    - size qty: c23+ (与 size header 同列)
    """
    rows: List[Dict[str, object]] = []

    # 找 c22='Size' 标记行
    size_anchor_row = None  # 0-idx
    for r in range(min(10, ws.max_row)):
        v = ws.cell(r + 1, 22).value  # 1-idx c22
        if isinstance(v, str) and v.strip().upper() == "SIZE":
            size_anchor_row = r
            break
    if size_anchor_row is None:
        return rows

    # size header 行 = anchor+1 (0-idx)
    size_header_row = size_anchor_row + 1
    # 收集 size labels:从 c22 开始(0-idx 21), 跳过 size type 字母 (N1/G/L/U), 直到第一个
    # 非空非 size token 的 col (这是 price label / 其他字段,如 "Unit Rtl" / "Unit Whls" / "Mu/Disc")
    size_cols: List[int] = []  # 0-idx col
    size_label_map: Dict[int, str] = {}
    SIZE_TYPE_TOKENS = {"N1", "G", "L", "U", "M", "F", "K"}  # 跳过 size type 字母
    for c in range(21, min(ws.max_column, 40)):  # 0-idx 21 = 1-idx c22
        v = ws.cell(size_header_row + 1, c + 1).value
        if isinstance(v, str) and v.strip() in SIZE_TYPE_TOKENS:
            # size type 字母 (e.g. N1/G/L/U) → 跳过,继续找 size labels
            continue
        if _is_size_token(v):
            size_cols.append(c)
            size_label_map[c] = _format_size_label(v)
        elif v is None or (isinstance(v, str) and not v.strip()):
            # 空 col 跳过(允许中间空 col 28/33 等)
            continue
        else:
            # 非空非 size token → 终止(price label 或其他字段)
            break

    if not size_cols:
        return rows

    # 先扫一遍 data rows (r, code), 收集所有 data row 位置 — 用于"块"图匹配
    data_rows_raw: List[Tuple[int, str]] = []  # (0-idx row, code)
    for r in range(size_header_row + 1, ws.max_row):
        code = normalize_text(ws.cell(r + 1, 3).value)
        if not code or len(code) < 5:
            continue
        if code.upper() in ("SIZE",):
            continue
        # 检查这行是不是 data (size 列至少有一个 qty > 0)
        has_data = False
        for c in size_cols:
            v = ws.cell(r + 1, c + 1).value
            if v is not None and isinstance(v, (int, float)) and v > 0:
                has_data = True
                break
        if not has_data:
            continue
        data_rows_raw.append((r, code))

    # 第二遍:对每个 data row,用「块匹配」找图 —
    # 图必须落在 [prev_data_row, current_data_row + 1) 范围内,避免把下一个商品的图
    # 错借给"源文件没图"的 data row(2026-09-03 bug 修复)。
    for idx, (r, code) in enumerate(data_rows_raw):
        prev_r = data_rows_raw[idx - 1][0] if idx > 0 else None
        row_images = _find_image_in_block(images_by_row, prev_r, r)

        desc = normalize_text(ws.cell(r + 1, 11).value)
        category = normalize_text(ws.cell(r + 1, 8).value)
        color = normalize_text(ws.cell(r + 1, 17).value)
        retail = _to_float_eur(ws.cell(r + 1, 45).value)  # Unit Rtl
        cost = _to_float_eur(ws.cell(r + 1, 48).value)    # Unit Whls

        # A 类 SKU 拼接规则(2026-09-03 user 反馈):
        #   1. 如果 color 字段 (c17) 有值 → `<STYLE>-<COLOR>`(优先颜色)
        #   2. 否则如果 code 末段抽出 color_code → `<STYLE>-<COLOR_CODE>`(其次色号)
        #   3. 都没有 → 不拼,用 code 原样
        # COLOR 归一:多空格压成单空格,保留原大小写。
        # 真实样本:
        #   'AMBMJE1119  097' + 'VINTAGE BLACK' → 'AMBMJE1119-VINTAGE BLACK' (color 优先)
        #   'EM006736 AF29202 FC368' + 'BLACK  EA' → 'EM006736 AF29202-BLACK EA'
        #   '44MAA120F26J00D  001' + 'BLACK' → '44MAA120F26J00D-BLACK'
        #   'PAMYC001DN001  4040' + 'LIGHT BLUE LIGHT BLUE' → 'PAMYC001DN001-LIGHT BLUE LIGHT BLUE'
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        # _format_sku_split 已经返回了 `STYLE-COLOR_CODE` 拼好的字符串,这里要拆出纯 STYLE
        style_with_suffix, color_code_from_code = _format_sku_split(code)
        style_pure = style_with_suffix.split("-", 1)[0] if "-" in style_with_suffix else style_with_suffix
        if color_normalized:
            sku_formatted = f"{style_pure}-{color_normalized}"
            color_code = color_normalized.split()[0]
        elif color_code_from_code:
            sku_formatted = f"{style_pure}-{color_code_from_code}"
            color_code = color_code_from_code
        else:
            sku_formatted = style_pure
            color_code = ""

        for c in size_cols:
            v = ws.cell(r + 1, c + 1).value
            if v is None or not isinstance(v, (int, float)) or v <= 0:
                continue
            qty = int(round(v))
            rows.append({
                "brand": "",  # A 类源没 brand 字段,framework 走文件名兜底
                "sku": sku_formatted,
                "image": "",
                "description": desc,
                "category": category,
                "material": "",
                "color": color,
                "color_code": color_code,
                "gender": "",
                "season": "",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label_map[c],
                "quantity": qty,
                "vendor_id": "V1072",
                "image_items": row_images,
            })
    return rows


def _parse_format_c(ws, images_by_row) -> List[Dict[str, object]]:
    """Foglio1 A 类 (AMIRI 7 / BURBERRY 7 / CASABLANCA 7 / RHUDE 7 / THOM BROWNE 7)。

    结构(2026-09-03 接入):
    - R1: c2='Brand', c3=品牌名
    - R2: c2='Category', c3=类目 (file-level)
    - R3: c2='Sector', c3=性别 (Uomo/Donna)
    - R4: 空
    - R5: header (c2='Code', c4='Category', c5='Descriptions', c6='Color', c7='Size')
          + Qty/Rtl/Whls 列(位置因文件而异,动态扫)
    - R6-R(N): size matrix(可能多行,每行 c7=size type letter + c8+ size labels,
                       跳 col 13;col 14+ 又是 size type matrix 第二段)
    - R(N+1)+: data(单行单商品, c2=Code, c4=Category, c5=Description, c6=Color,
                     c8+ = qty per size)

    图匹配:复用 _find_image_in_block 块匹配(data 之间空 1-2 行,所以块边界工作)
    """
    rows: List[Dict[str, object]] = []

    # 1) 拿 R1-R3 文件级 brand / category / sector
    file_brand = normalize_text(ws.cell(1, 3).value)  # R1 c3=品牌
    file_category = normalize_text(ws.cell(2, 3).value)  # R2 c3=类目
    file_sector = normalize_text(ws.cell(3, 3).value)  # R3 c3=性别

    # 2) 找 header 行 (c2='Code')
    header_row = None
    for r in range(1, min(15, ws.max_row + 1)):
        v = normalize_text(ws.cell(r, 2).value).upper()
        if v == "CODE":
            header_row = r
            break
    if header_row is None:
        return rows

    # 3) 从 header 找 Qty / Unit Rtl / Unit Whls 列(动态)
    qty_col: Optional[int] = None
    rtl_col: Optional[int] = None
    whls_col: Optional[int] = None
    for c in range(1, ws.max_column + 1):
        v = normalize_text(ws.cell(header_row, c).value).upper()
        if v == "QTY":
            qty_col = c
        elif v == "UNIT RTL":
            rtl_col = c
        elif v == "UNIT WHLS":
            whls_col = c

    # 4) 扫 size matrix(header_row+1 起, 直到第一个非 size type / 非空 size label 行)
    SIZE_TYPE_TOKENS = {"L", "G", "U", "M", "H", "D", "O", "Y", "X", "W1", "T", "UK", "OS", "N1", "F", "K", "38"}
    size_cols: List[int] = []  # 0-idx
    size_label_map: Dict[int, str] = {}

    def _is_size_label(v) -> bool:
        if v is None:
            return False
        if isinstance(v, (int, float)):
            return True
        if isinstance(v, str):
            sv = v.strip()
            if not sv:
                return False
            if re.fullmatch(r"\d+(?:[.,]\d+)?", sv):
                return True
            if sv in ("XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "3XL", "4XL", "5XL", "6XL",
                      "OS", "TU", "UNI", "ONE SIZE"):
                return True
        return False

    r = header_row + 1
    while r <= min(header_row + 6, ws.max_row):
        v = ws.cell(r, 7).value  # 1-idx c7=Size type
        if v is None:
            r += 1
            continue
        v_str = str(v).strip()
        if v_str.upper() not in SIZE_TYPE_TOKENS and not _is_size_label(v):
            # 既不是 size type 也不是 size label → size matrix 结束
            break
        # 收 c8+ 的 size labels(可能跨 col 13/18 等空 col)
        for c in range(8, ws.max_column + 1):
            label_v = ws.cell(r, c).value
            if label_v is None:
                continue
            label_str = str(label_v).strip()
            if not label_str:
                continue
            if _is_size_label(label_v):
                size_cols.append(c)
                size_label_map[c] = label_str
        r += 1

    if not size_cols:
        return rows

    # 5) data 行扫描:从 size matrix 结束后开始, c2 非空字符串且不是 'Code'/'Brand' header label
    DATA_HDR_TOKENS = {"CODE", "BRAND", "CATEGORY", "SECTOR", "DESCRIPTIONS", "COLOR", "SIZE", "QTY",
                        "UNIT RTL", "UNIT WHLS"}
    data_rows_raw: List[Tuple[int, str]] = []  # (1-idx row, code)
    for r in range(r, ws.max_row + 1):
        code_v = ws.cell(r, 2).value
        if code_v is None or not str(code_v).strip():
            continue
        code_str = str(code_v).strip()
        if code_str.upper() in DATA_HDR_TOKENS:
            continue
        # 至少有一个 size col 有 qty > 0
        has_data = False
        for c in size_cols:
            v = ws.cell(r, c).value
            if v is not None and isinstance(v, (int, float)) and v > 0:
                has_data = True
                break
        if has_data:
            data_rows_raw.append((r, code_str))

    # 6) 对每个 data row 出 row
    for idx, (r, code) in enumerate(data_rows_raw):
        prev_r = data_rows_raw[idx - 1][0] if idx > 0 else None
        row_images = _find_image_in_block(images_by_row,
                                          (prev_r - 1) if prev_r is not None else None,
                                          r - 1)
        item_category = normalize_text(ws.cell(r, 4).value)
        desc = normalize_text(ws.cell(r, 5).value)
        color = normalize_text(ws.cell(r, 6).value)
        retail = _to_float_eur(ws.cell(r, rtl_col).value) if rtl_col else None
        cost = _to_float_eur(ws.cell(r, whls_col).value) if whls_col else None

        # color 优先拼接 SKU
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        # Foglio1 版式的 Code 本身可能含空格('AMBMJE1002  097','8119013  A1189')也可能有 '-'
        # (如 CASABLANCA 'M-SS26-JTR-003-02' — '-' 是款号内部结构,不是分隔符)
        # 安全提取:只按空格分末段当色号兜底
        code_clean = code.strip()
        if " " in code_clean:
            parts = code_clean.split()
            style_pure = " ".join(parts[:-1])
            color_code_from_code = parts[-1]
        else:
            style_pure = code_clean
            color_code_from_code = ""
        if color_normalized:
            sku_final = f"{style_pure}-{color_normalized}"
            color_code_field = color_normalized.split()[0]
        elif color_code_from_code:
            sku_final = f"{style_pure}-{color_code_from_code}"
            color_code_field = color_code_from_code
        else:
            sku_final = style_pure
            color_code_field = ""

        for c in size_cols:
            v = ws.cell(r, c).value
            if v is None or not isinstance(v, (int, float)) or v <= 0:
                continue
            qty = int(round(v))
            rows.append({
                "brand": file_brand,
                "sku": sku_final,
                "image": "",
                "description": desc,
                "category": item_category,
                "material": "",
                "color": color,
                "color_code": color_code_field,
                "gender": file_sector,
                "season": "",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label_map[c],
                "quantity": qty,
                "vendor_id": "V1072",
                "image_items": row_images,
            })
    return rows


def _parse_format_d(ws, images_by_row) -> List[Dict[str, object]]:
    """Foglio1 B 类无 header (FENDI 7)。

    R1+ 直接是 data(无 R0 header 行),字段跟 B 类 (Sheet1 + Descr_Article) 一样:
    - c2=Brand, c3=Category, c4=Season, c5=Code, c6=Gender, c7=Color, c8=MadeIn,
      c9=Composizione, c10=CodiceDoganale
    - c11=Size, c12=Qty
    - c13=Cost, c14=Retail
    """
    rows: List[Dict[str, object]] = []
    for r in range(1, ws.max_row + 1):
        code = normalize_text(ws.cell(r, 5).value)
        if not code:
            continue
        size_label = ws.cell(r, 11).value
        if size_label is None:
            continue
        size_str = str(size_label).strip()
        if not size_str:
            continue

        qty_v = ws.cell(r, 12).value
        qty = 0
        if qty_v is not None:
            try:
                qty = int(round(float(qty_v)))
            except (TypeError, ValueError):
                qty = 0
        if qty <= 0:
            continue

        brand = normalize_text(ws.cell(r, 2).value)
        desc = normalize_text(ws.cell(r, 3).value)
        season = normalize_text(ws.cell(r, 4).value)
        gender = normalize_text(ws.cell(r, 6).value)
        color = normalize_text(ws.cell(r, 7).value)
        origin = normalize_text(ws.cell(r, 8).value)
        material = normalize_text(ws.cell(r, 9).value)
        cost = _to_float_eur(ws.cell(r, 13).value)
        retail = _to_float_eur(ws.cell(r, 14).value)

        # B 类 Code '8R8136AQY2F1MXM' 无空格 → 不拼(按 Step 3.5 反例)
        sku_final = code

        row_images = _find_image_near_row(images_by_row, r - 1)
        rows.append({
            "brand": brand,
            "sku": sku_final,
            "image": "",
            "description": desc,
            "category": desc,  # B 类没独立 category
            "material": material,
            "color": color,
            "color_code": "",
            "gender": gender,
            "season": season,
            "origin": origin,
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            "size": size_str,
            "quantity": qty,
            "vendor_id": "V1072",
            "image_items": row_images,
        })
    return rows


def get_vendor_meta() -> Dict[str, object]:
    return {}


def parse_rows(
    primary_file: Path,
    _: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """V1072-期货 (FW26 IN ARRIVO 8 个品牌) 解析入口。

    自动 detect:
    - B 类 (MAGDA BUTRYN/MAGDA BUTRYM): R1 c1='Brand' + c3='Descr_Article'
    - A 类 (其他 6 个): c22='Size' 锚点

    仓库 ID 固定为 V1072(同 V1072 现货,BURBERRY);虽然现货是 BURBERRY 块状,
    本 adapter 是 8 个品牌的期货,共享 V1072 仓库 ID。
    """
    # 强制覆盖为 V1072 — 即使调用方传 'V1072-期货' (dispatch key) 也忽略,
    # 输出行的 vendor_id 字段统一写 V1072,跟 V1072 现货(BURBERRY)共享仓库 ID。
    actual_vendor_id = "V1072"
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    if not wb.worksheets:
        return []
    ws = wb.worksheets[0]

    images_by_row = _load_images_by_row(ws)

    # Detect 版式 — 4 档优先级:
    # 1. B 类 (Sheet1 + Descr_Article header): R1 c3 = 'Descr_Article'
    # 2. Foglio1 B 类无 header (FENDI 7): R1 c2 是品牌名 + c3 是描述(不是 header label)
    # 3. Foglio1 A 类: R1 c2 = 'Brand' + c3 = 品牌名(不是 'Descr_Article')
    # 4. 默认 A 类 (LISTE FW26 IN ARRIVO 6 个品牌)
    c2 = normalize_text(ws.cell(1, 2).value)
    c3 = normalize_text(ws.cell(1, 3).value).upper()
    if c3 == "DESCR_ARTICLE":
        rows = _parse_format_b(ws, images_by_row)
    elif c2.upper() != "BRAND" and c2 and c3 and c3 not in ("DESCR_ARTICLE", "CODE", "ITEM"):
        # Foglio1 B 类无 header: R1 c2 是品牌名(e.g. 'FENDI'),c3 是描述
        rows = _parse_format_d(ws, images_by_row)
    elif c2.upper() == "BRAND" and c3 and c3 not in ("DESCR_ARTICLE",):
        # Foglio1 A 类: R1 c2='Brand' + c3=品牌名(e.g. 'AMIRI','BURBERRY')
        rows = _parse_format_c(ws, images_by_row)
    else:
        rows = _parse_format_a(ws, images_by_row)

    # 强制 vendor_id
    for r in rows:
        r["vendor_id"] = actual_vendor_id
    return rows
