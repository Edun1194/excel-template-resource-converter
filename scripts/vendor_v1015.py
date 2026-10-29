"""V1015 PRADA ITALY vendor adapter.

适配文件:Prada 意大利线的库存清单(鞋类为主,意大利尺码 5-14 / 35-44)。

文件版式(.xlsx):
- 每个产品占 10 行(7-8 个产品 / sheet)
- 第 N 行    (header_row):  col 1 = 总数量, col 3 = 'SKU' 文本, col 8+ = 各尺码 qty
- 第 N+2 行 (sku_row):      col 24 = SKU 编码(20 位)
- 第 N+3 行 (desc_row):     col 24 = 描述(col 35 = 颜色)
- 第 N+5 行 (price_label):  col 26 = 'Acquisto'(cost), col 32 = 'Vendita'(retail)
- 第 N+6 行 (price_val):    col 26 = cost 数字, col 32 = retail 数字

- 尺码标签行:在文件顶部某行(本测试文件是 row 12),col 2 = 'SKU' 文本(带 \\r\\n),
              col 4-27 写的是尺码值(意大利鞋码 5/5.5/6/6.5/7/7.5/8/8.5/9/9.5/10/.../14)
- 图片:每个产品 3 张,在图片行(openpyxl 0-idx = header_row + 1)的 col 0/7/14

参考文档:references/patterns/adapter-development.md §3.6 自定义尺码表 + image-handling §9
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

from convert_vendor_quote_shared import (
    embed_images_in_row,  # noqa: F401  (保留供未来扩展)
    guess_product_group,
    has_richvalue_images,
    is_positive_qty,
    load_richvalue_images,
    load_sheet_images_by_row,
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def get_vendor_meta() -> Dict[str, object]:
    """V1015 PRADA ITALY vendor 元数据 — 给 dispatcher 用。

    default_image_col_width=40:现在 KEEP_IMAGES_PER_ROW=1(只保留第一张),
    92pt 高 ≈ 122px ≈ 17 字符,加一些缓冲 40 字符(2026-07-17 减半,原 80)。
    """
    return {
        "default_image_col_width": 40.0,
    }


# 0-indexed 关键列位置
HEADER_COUNT_COL = 0        # 总数量
SKU_LABEL_COL = 1           # 尺码标签行的 'SKU' 文本(带 \r\n)
HEADER_SKU_LABEL_COL = 2    # 产品 header 行的 'SKU' 文本
SKU_COL = 23                # SKU 编码 (col 24 1-idx)
DESC_COL = 23               # 描述 (col 24 1-idx)
COLOR_COL = 34              # 颜色 (col 35 1-idx)
COST_LABEL_COL = 25         # 'Acquisto' (col 26 1-idx)
RETAIL_LABEL_COL = 31       # 'Vendita' (col 32 1-idx)
COST_VAL_COL = 25           # cost 数字
RETAIL_VAL_COL = 31         # retail 数字

# 尺码标签列范围 (0-idx):col 3-26,中间有空列(col 14/15/17/21/28/30/31 都是空)
SIZE_COL_START = 3
SIZE_COL_END = 27


def _looks_like_size_token(s: str) -> bool:
    """判断字符串是否可能是尺码 token(纯数字 / 欧码格式 / 短大写代码)。

    不硬编码 "SKU" / "SID" 等具体值 — 任何符合"短/不含特殊字符/不含中文"的都算候选。

    注:对 `"34\n5"` (欧+UK 双码)会先 split 再判断第一段。
    """
    if not s:
        return False
    s = str(s).strip()
    if not s:
        return False
    # 双码(欧+UK 用 \n 分隔):只看欧码段(第一段)
    if "\n" in s:
        s = s.split("\n")[0].strip()
    if not s:
        return False
    # 数字(5 / 5,5 / 5.5 / 10.5)
    if re.fullmatch(r"\d+(?:[.,]\d+)?", s):
        return True
    # 短大写 token(2-6 字符,字母/数字,无中文/特殊符号)
    if re.fullmatch(r"[A-Z0-9]{2,6}", s):
        return True
    return False


def _find_size_label_row(df: pd.DataFrame) -> Optional[int]:
    """找尺码标签行 — 行内有多个尺码 token 的就是。

    不硬编码 "SKU" / "SID" — 用"行内尺码 token 数量"判断。
    """
    best_row: Optional[int] = None
    best_count = 0
    for idx, row in df.iterrows():
        size_token_count = 0
        for c in range(SIZE_COL_START, min(SIZE_COL_END, len(row))):
            v = row.iloc[c]
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            if isinstance(v, (int, float)):
                size_token_count += 1
            elif isinstance(v, str) and _looks_like_size_token(v.strip()):
                size_token_count += 1
        if size_token_count >= 5 and size_token_count > best_count:
            best_count = size_token_count
            best_row = int(idx)
    return best_row


def _build_size_labels(df: pd.DataFrame, size_row: int, seg_idx: int = 0) -> Dict[int, str]:
    """从尺码标签行提取 {col_idx: size_label}。空 col 不计入。

    seg_idx: 尺码行每 col 的 cell 可能多段(用 \\n 分隔),选第几段。
      - 0 = 第一段(默认:欧码 / 裤码)
      - 1 = 第二段(UK 码 / 均码)
      - 2 = 第三段(鞋码 — MIU MIU E_485 风格)

    源文件尺码格式:
      - 整数(5, 6, 10)  → 浮点 5.0/6.0/10.0  → 显示 "5"/"6"/"10"
      - 半码(5.5, 7.5)    → 字符串 "5,5"/"7,5" (欧洲十进制逗号) → "5.5"/"7.5"
      - 双码(欧+UK):"34\\n5"  → seg_idx=0 取 "34" (欧码),seg_idx=1 取 "5" (UK)
      - 三码(裤+均+鞋):"XD\\nUNI\\nSID" / "0P\\nUNI\\n34"  → seg_idx=0/1/2 各取一段
    """
    out: Dict[int, str] = {}
    if size_row is None or size_row >= len(df):
        return out
    row = df.iloc[size_row]
    for c in range(SIZE_COL_START, min(SIZE_COL_END, len(row))):
        v = row.iloc[c]
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        # 数值(5.0 / 10.0):显示为整数
        if isinstance(v, (int, float)):
            f = float(v)
            if f.is_integer():
                out[c] = str(int(f))
            else:
                out[c] = str(f).replace(",", ".")
            continue
        # 字符串("5,5" / "10" / "34\n5" / "XD\nUNI\nSID" / "")
        sv = str(v).strip()
        if not sv or sv.lower() == "nan":
            continue
        # 多段(用 \n 分隔):按 seg_idx 取一段
        if "\n" in sv:
            parts = sv.split("\n")
            if seg_idx >= len(parts):
                continue  # 选的段超出范围,跳过这个 col
            sv = parts[seg_idx].strip()
            if not sv:
                continue
        # 数字字符串(可能有 \r\n 尾巴,或欧洲逗号):归一化
        if re.fullmatch(r"\d+(?:[.,]\d+)?", sv):
            sv = sv.replace(",", ".")
        out[c] = sv
    return out


def _pick_seg_idx(cat_token: str, c2_full: str) -> int:
    """根据产品类别(cat_token 单段)和尺码行 c2 cell(c2_full 可能多段)选 seg_idx。

    例子:
      - cat_token='SID', c2_full='SID\\nSKU'   → 0 (SID 欧码段)
      - cat_token='SKU', c2_full='SID\\nSKU'   → 1 (SKU UK 码段)
      - cat_token='XD',  c2_full='XD\\nUNI\\nSID' → 0 (裤码段)
      - cat_token='UNI', c2_full='XD\\nUNI\\nSID' → 1 (均码段)
      - cat_token='SID', c2_full='XD\\nUNI\\nSID' → 2 (鞋码段)
      - cat_token='UNI', c2_full='UNI'         → 0 (单段,只能选段 0)
    """
    if not cat_token or not c2_full:
        return 0
    if "\n" not in c2_full:
        return 0
    segs = c2_full.split("\n")
    for i, s in enumerate(segs):
        if cat_token and cat_token in s:
            return i
    return 0


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """从 ws 抽图,按 openpyxl 0-idx row 分桶。"""
    out: Dict[int, List[Dict]] = {}
    for img in ws._images:
        anchor = img.anchor
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        row = anchor._from.row
        try:
            img_bytes = img._data() if hasattr(img, "_data") else None
        except Exception:
            img_bytes = None
        if img_bytes is None:
            continue
        out.setdefault(row, []).append({
            "bytes": img_bytes,
            "format": img.format or "jpeg",
        })
    return out


# 性别标题行的关键字(意大利语 / 英语)
# PRADA 现货单版式:c1 出现 "Donna - PRADA" / "Uomo - PRADA" 这类标题,标识下面产品是男/女
# 提取出 "Donna" / "Uomo" 作为性别原值,再翻译成英文 (Woman/Man) 跟其他 vendor 对齐
_GENDER_KEYWORDS = ["Donna", "Uomo", "Woman", "Man", "Bambina", "Bambino", "Unisex", "Junior"]
_GENDER_TRANSLATE = {
    "Uomo": "Man", "Man": "Man",
    "Donna": "Woman", "Woman": "Woman",
    "Unisex": "Unisex", "Junior": "Junior",
    "Bambino": "Kid", "Bambina": "Kid",
}


def _find_gender_for_header(df: pd.DataFrame, hdr_row: int) -> str:
    """在 hdr_row 之前向上找最近的性别标题行,提取性别原值。

    V1015 PRADA 现货单版式:
      - R16: c1='Donna - PRADA'  → 后续 R18-R56 的产品都属于 Donna
      - R58: c1='Uomo - PRADA'  → 后续 R60+ 的产品都属于 Uomo
    在 parse_rows 中,每个 hdr_row 用此函数找它的性别来源。

    返回值已翻译成英文 (Woman/Man),跟其他 vendor 的 gender 字段对齐。
    """
    for r in range(hdr_row - 1, -1, -1):
        v = df.iat[r, 0] if df.shape[1] > 0 else None
        if not isinstance(v, str):
            continue
        sv = v.strip()
        if not sv:
            continue
        for kw in _GENDER_KEYWORDS:
            if kw in sv:
                return _GENDER_TRANSLATE.get(kw, kw)
    return ""


def _parse_v1015_ready_to_ship(ws, vendor_id: str, images_by_row: Optional[Dict[int, List[Dict]]] = None) -> List[Dict[str, object]]:
    """V1015 'Ready to be shipped' sub-format (2026-09-08 接入)。

    结构 (Sheet1, 多品牌块):
    - R2: size matrix — 1-idx c24 = size code 段 (用 \n 分隔, e.g. 'XD\\nAIU\\nAID\\nUNI\\nCU\\nSID\\nSIU\\nXU'),
                  c26+ = size label 段 (每 col 用 \n 分多段对应不同 size code)
    - R(N): brand header (c1='Donna - BURBERRY' / 'Uomo - CANADA GOOSE' 等)
    - data 每 2 行一组:
        - SKU 行 (奇数行, 跟在 brand header 之后): c2=SKU, c6=Price (cost), c11=Tot Price,
                c13=Amount, c19=Retail, c23=qty, c25=size code (XD/AIU/...),
                c26+ = 各 size qty (按 c25 size code 选 seg_idx)
        - desc 行 (偶数行): c2=description, c4=color
    - 品牌从当前 brand header 取(Donna - BURBERRY 拆成 'BURBERRY' + gender 'Donna')
    """
    rows: List[Dict[str, object]] = []
    if ws.max_row < 10 or ws.max_column < 30:
        return rows

    # 1) 解析 R2 size matrix: c24=size code 段, c26+ = 各 col 的 size label 段
    size_code_cell = ws.cell(2, 24).value
    if not isinstance(size_code_cell, str) or "\n" not in size_code_cell:
        return rows
    size_codes = [s.strip() for s in size_code_cell.split("\n") if s.strip()]
    if not size_codes:
        return rows

    # 收集每 col 的 size label 段(每个 col 可能 \n 分多段对应不同 size code)
    # 然后用 R2 已有 label 的 col 锚定, 推断 R2 空 col 的 size label (取最近 col 的 label)
    col_label_segs_raw: Dict[int, List[str]] = {}
    SIZE_COL_START = 26  # 1-idx
    SIZE_COL_END = min(ws.max_column + 1, 56)
    for c in range(SIZE_COL_START, SIZE_COL_END):
        v = ws.cell(2, c).value
        if v is None:
            continue
        segs = [s.strip() for s in str(v).split("\n")]
        # 至少有一段非空(可能头尾 \n 全部空, 中间有内容)
        if not any(s for s in segs):
            continue
        col_label_segs_raw[c] = segs

    # 推断空 col 的 size label: 用最近的左/右 R2 label col
    col_label_segs: Dict[int, List[str]] = {}
    sorted_cols = sorted(col_label_segs_raw.keys())
    for c in range(SIZE_COL_START, SIZE_COL_END):
        if c in col_label_segs_raw:
            col_label_segs[c] = col_label_segs_raw[c]
        else:
            # 找最近的 R2 label col (左边优先)
            nearest = None
            for lc in sorted_cols:
                if lc <= c:
                    nearest = lc
                else:
                    break
            if nearest is not None:
                col_label_segs[c] = col_label_segs_raw[nearest]

    # 2) 扫描 data: brand header 行 → 2 行一组 (SKU + desc)
    #    从 R3 开始
    current_brand = ""
    current_gender = ""
    for r in range(3, ws.max_row + 1):
        # 检查是否是 brand header (c1 是 'Donna - XXX' / 'Uomo - XXX' 模式)
        c1_v = ws.cell(r, 1).value
        if c1_v and isinstance(c1_v, str) and " - " in c1_v:
            # brand header
            parts = [p.strip() for p in c1_v.split(" - ", 1)]
            if len(parts) == 2:
                current_gender = parts[0]  # 'Donna' / 'Uomo'
                current_brand = parts[1]
            continue

        # 检查是否是 SKU 行:c2 是 SKU (短数字/字母组合), c25 是 size code 字符
        sku_v = ws.cell(r, 2).value
        size_code_v = ws.cell(r, 25).value  # 1-idx c25 = size code
        if not sku_v or not isinstance(sku_v, str) or not sku_v.strip():
            continue
        if not size_code_v or not isinstance(size_code_v, str):
            continue
        size_code = size_code_v.strip().upper()
        if size_code not in size_codes:
            continue
        sku = sku_v.strip()
        # 跳过纯 size code 标签 (如 'XD' / 'SID' / 'AIU' 单字符)
        if len(sku) <= 3 and sku.upper() in size_codes:
            continue
        # 跳过 'Item' / 'SKU' 等 header 残留
        if sku.upper() in ("ITEM", "SKU", "BRAND", "TOTAL", "TOTALE", "QTY", "PRICE"):
            continue

        # 下一行是 desc 行
        desc = ""
        color = ""
        if r + 1 <= ws.max_row:
            desc = normalize_text(ws.cell(r + 1, 2).value)
            color = normalize_text(ws.cell(r + 1, 4).value)

        # 字段
        cost = to_float(ws.cell(r, 6).value) or to_float(ws.cell(r, 11).value)
        retail = to_float(ws.cell(r, 19).value)
        total_qty = to_float(ws.cell(r, 23).value) or 0
        total_qty = int(round(total_qty))

        # 货号拼接:color 优先
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        if color_normalized:
            sku_final = f"{sku}-{color_normalized}"
            color_code_field = color_normalized.split()[0]
        else:
            sku_final = sku
            color_code_field = ""

        # 找该 size_code 对应的 seg_idx
        seg_idx = size_codes.index(size_code) if size_code in size_codes else 0

        # 图:openpyxl 0-idx anchor row 相对 SKU 1-idx row 偏移有两种风格:
        #   - 怀远 MAX MARA Group 等:图跟 SKU 同行 (0-idx = r - 1)
        #   - 老 PRADA 现货单等:图在 SKU 上一行 (0-idx = r - 2)
        # 先查 r-1,fallback r-2,覆盖两种 anchor
        row_images = (images_by_row or {}).get(r - 1) or (images_by_row or {}).get(r - 2) or []

        # 出 size 展开:每 col 用 col_label_segs 取 seg_idx 段
        emitted = False
        for c, segs in col_label_segs.items():
            if seg_idx >= len(segs):
                continue
            size_label = segs[seg_idx].strip()
            if not size_label:
                continue
            qty_v = ws.cell(r, c).value
            if qty_v is None or not isinstance(qty_v, (int, float)) or qty_v <= 0:
                continue
            qty = int(round(qty_v))
            emitted = True
            rows.append({
                "brand": current_brand,
                "sku": sku_final,
                "image": "",
                "description": desc,
                "category": "",
                "material": "",
                "color": color,
                "color_code": color_code_field,
                "gender": current_gender,
                "season": "",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })

        # 2026-09-08 fix:如果 c26+ size 分布全 0 (emitted=False) 但 c23 total qty > 0,
        # 出一个 TOTAL 行带全量
        if not emitted and total_qty > 0:
            rows.append({
                "brand": current_brand,
                "sku": sku_final,
                "image": "",
                "description": desc,
                "category": "",
                "material": "",
                "color": color,
                "color_code": color_code_field,
                "gender": current_gender,
                "season": "",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": "TOTAL",
                "quantity": total_qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
    return rows


def _parse_v1015_bos(ws, vendor_id: str, images_by_row: Optional[Dict[int, List[Dict]]] = None) -> List[Dict[str, object]]:
    """V1015 BOS-MONCLER sub-format (2026-09-04 接入)。

    结构 (Foglio1 sheet):
    - R1: header (c1='Brand', c2='Settore', c3='ModelloVariante', c5='Taglia',
                 c6='Descrizione Modello', c7='Colore', c8='Prezzo Acquisto',
                 c9='Q.tà')
    - R2+: data — 一行一个 size, c3=SKU (e.g. 'L10911A00115539ZD255'),
                   c5=size (e.g. '2/II'), c6=description, c7=color,
                   c8=cost, c9=qty
    - 图在 c3 (D 列) anchor, openpyxl 0-idx = 1-idx row - 1
    """
    rows: List[Dict[str, object]] = []
    if ws.max_row < 2 or ws.max_column < 9:
        return rows

    for r in range(2, ws.max_row + 1):
        sku_v = ws.cell(r, 3).value
        if sku_v is None or not str(sku_v).strip():
            continue
        sku = str(sku_v).strip()

        size = normalize_text(ws.cell(r, 5).value)
        if not size:
            continue

        qty_v = ws.cell(r, 9).value
        try:
            qty = int(round(float(qty_v or 0)))
        except (TypeError, ValueError):
            qty = 0
        if qty <= 0:
            continue

        brand = normalize_text(ws.cell(r, 1).value)
        gender = normalize_text(ws.cell(r, 2).value)  # Settore
        desc = normalize_text(ws.cell(r, 6).value)
        color = normalize_text(ws.cell(r, 7).value)
        cost = to_float(ws.cell(r, 8).value)

        # 货号拼接:color 优先(Step 3.5.1)
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        if color_normalized:
            sku_final = f"{sku}-{color_normalized}"
            color_code_field = color_normalized.split()[0]
        else:
            sku_final = sku
            color_code_field = ""

        # 图:openpyxl 0-idx anchor row 相对 SKU 1-idx row 偏移有两种风格:
        #   - 怀远 MAX MARA Group 等:图跟 SKU 同行 (0-idx = r - 1)
        #   - 老 PRADA 现货单等:图在 SKU 上一行 (0-idx = r - 2)
        # 先查 r-1,fallback r-2,覆盖两种 anchor
        row_images = (images_by_row or {}).get(r - 1) or (images_by_row or {}).get(r - 2) or []

        rows.append({
            "brand": brand,
            "sku": sku_final,
            "image": "",
            "description": desc,
            "category": "",
            "material": "",
            "color": color,
            "color_code": color_code_field,
            "gender": gender,
            "season": "",
            "origin": "",
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": None,
            "size": size,
            "quantity": qty,
            "vendor_id": vendor_id,
            "image_items": row_images,
        })
    return rows


def _parse_hk_huaiyuan(ws, vendor_id: str, images_by_row: Optional[Dict[int, List[Dict]]] = None) -> List[Dict[str, object]]:
    """V1015 怀远 (HK) 新版式 (2026-09-03 接入)。

    结构:
    - R1-R15+: size code 矩阵 — 1-idx c12 = size code (e.g. 'SID'/'XN'/'JU'/'AIU'/'UNI'),
                 1-idx c14+ = size labels (跳 c13 空白列)
    - data 行: 1-idx c2 = SKU (20 位 alphanumeric, e.g. 'L10911A00042539ZDM93')
                 c3 = 'Abbigliamento' (category 1), c4 = 'Main' (super),
                 c5 = 'I'/'E' (season marker), c6 = brand (MONCLER/STONE ISLAND/C.P. COMPANY),
                 c7 = category 2 (Giubbotti e bomber / Jeans), c8 = gender (Uomo),
                 c9 = color, c10 = cost (欧式 '443,00'), c11 = retail,
                 c12 = size code (匹配 size matrix), c13 = total qty,
                 c14+ = qty per size
    """
    rows: List[Dict[str, object]] = []
    if ws.max_row < 20 or ws.max_column < 14:
        return rows

    # 1) 扫 size code 矩阵:R1 起, c12 是 size code 字符 + c14+ 是字符串 size label
    #    (不是数字 qty) 才当 matrix;遇到 c14 是数字 (0 或正整数) 立即停止
    # 2026-09-03 fix:之前错误地把 data 行 (c14=0) 也当 matrix,导致 XU 矩阵 size labels 被覆盖
    size_maps: Dict[str, Dict[int, str]] = {}  # size_code -> {col_idx: size_label}
    SIZE_CODE_MAX_LEN = 6  # 'AIDD3'/'AFD3'/'JU12' 等可能最长 ~6 字符
    for r in range(1, min(40, ws.max_row + 1)):
        code_raw = ws.cell(r, 12).value
        if code_raw is None:
            continue
        code = str(code_raw).strip()
        if not code:
            continue
        if len(code) > SIZE_CODE_MAX_LEN:
            continue
        if not re.match(r"^[A-Z0-9]+$", code):
            continue
        # 收集 c14+ 的 size labels(允许部分为空),但 c14 必须是字符串才算 matrix
        first_label = ws.cell(r, 14).value
        if first_label is None or isinstance(first_label, (int, float)):
            # c14 是数字 → 这是 data 行,不再是 matrix
            break
        col_map: Dict[int, str] = {}
        for c in range(14, ws.max_column + 1):
            label_v = ws.cell(r, c).value
            if label_v is None:
                continue
            if isinstance(label_v, (int, float)):
                # size label 段里出现数字 (如 '34') — 不对,size label 是字符串
                # 这种情况算 matrix 自然结束
                break
            label_str = str(label_v).strip()
            if not label_str:
                continue
            col_map[c] = label_str
        if col_map:
            size_maps[code] = col_map

    if not size_maps:
        return rows

    # 2) 扫 data 行:c2 是 20 位 alphanumeric 字符串(老 V1015 的 c3='SKU' 文本不同)
    DATA_HDR_TOKENS = {"SKU", "BRAND", "CATEGORY", "SEASON", "ITEM", "TOTAL"}
    data_rows_raw: List[Tuple[int, str]] = []  # (1-idx row, sku)
    for r in range(1, ws.max_row + 1):
        sku_v = ws.cell(r, 2).value
        if sku_v is None:
            continue
        sku_str = str(sku_v).strip()
        if not sku_str or sku_str.upper() in DATA_HDR_TOKENS:
            continue
        # 20 位 alphanumeric(怀远 SKU 模式: 'L10911A00042539ZDM93' / '21CMPA207A006252SR90')
        if not re.match(r"^[A-Z0-9]{9,24}$", sku_str):
            continue
        # size code 必须有(否则跳过)
        size_code_raw = ws.cell(r, 12).value
        if size_code_raw is None or str(size_code_raw).strip() == "":
            continue
        # 至少一个 size col 有 qty > 0
        size_code = str(size_code_raw).strip().upper()
        col_map = size_maps.get(size_code, {})
        if not col_map:
            continue
        has_data = False
        for c in col_map.keys():
            v = ws.cell(r, c).value
            if v is not None and isinstance(v, (int, float)) and v > 0:
                has_data = True
                break
        # 2026-09-04 fix:即使 c14+ size 分布全 0,只要 c13 total qty > 0 也出
        # (源数据有些行只填了 total 没填 size 分布,如 BARBOUR/CANADA GOOSE)
        total_qty = ws.cell(r, 13).value
        try:
            total_qty = int(round(float(total_qty or 0)))
        except (TypeError, ValueError):
            total_qty = 0
        if has_data or total_qty > 0:
            data_rows_raw.append((r, sku_str))

    if not data_rows_raw:
        return rows

    # 3) 出 row
    for r, sku in data_rows_raw:
        size_code = str(ws.cell(r, 12).value).strip().upper()
        col_map = size_maps[size_code]
        category = normalize_text(ws.cell(r, 3).value)
        super_ = normalize_text(ws.cell(r, 4).value)
        season_marker = normalize_text(ws.cell(r, 5).value)
        brand = normalize_text(ws.cell(r, 6).value)
        category2 = normalize_text(ws.cell(r, 7).value)
        gender = normalize_text(ws.cell(r, 8).value)
        color = normalize_text(ws.cell(r, 9).value)
        cost = to_float(ws.cell(r, 10).value)
        retail = to_float(ws.cell(r, 11).value)
        # 读 c13 total qty(2026-09-04 补:即使 c14+ 全 0, c13>0 也出)
        total_qty_v = ws.cell(r, 13).value
        try:
            total_qty = int(round(float(total_qty_v or 0)))
        except (TypeError, ValueError):
            total_qty = 0

        # 货号拼接:color 优先(Step 3.5.1 规则)
        color_normalized = re.sub(r"\s+", " ", color).strip() if color else ""
        if color_normalized:
            sku_final = f"{sku}-{color_normalized}"
            color_code_field = color_normalized.split()[0]
        else:
            sku_final = sku
            color_code_field = ""

        # 图:openpyxl 0-idx anchor row 相对 SKU 1-idx row 偏移有两种风格:
        #   - 怀远 MAX MARA Group 等:图跟 SKU 同行 (0-idx = r - 1)
        #   - 老 PRADA 现货单等:图在 SKU 上一行 (0-idx = r - 2)
        # 先查 r-1,fallback r-2,覆盖两种 anchor
        row_images = (images_by_row or {}).get(r - 1) or (images_by_row or {}).get(r - 2) or []

        emitted_size = False
        for c, size_label in col_map.items():
            v = ws.cell(r, c).value
            if v is None or not isinstance(v, (int, float)) or v <= 0:
                continue
            qty = int(round(v))
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": category2,  # 用 category2 (Giubbotti e bomber) 当描述
                "category": category2,
                "material": "",
                "color": color,
                "color_code": color_code_field,
                "gender": gender,
                "season": season_marker,
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
            emitted_size = True

        # 2026-09-04 补:如果 c14+ size 分布全 0 (emitted_size=False) 但 c13 total > 0,
        # 出一个 "TOTAL" 行带全量
        if not emitted_size and total_qty > 0:
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": category2,
                "category": category2,
                "material": "",
                "color": color,
                "color_code": color_code_field,
                "gender": gender,
                "season": season_marker,
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": "TOTAL",
                "quantity": total_qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
    return rows


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []

    # 直接用 openpyxl 读(避开 pandas 0-idx 偏移坑)
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    if not wb.worksheets:
        return rows
    ws = wb.worksheets[0]

    # 2026-09-08 'Ready to be shipped' sub-format detect:R2 c24 是 size code 多段 (\n 分)
    # 例 'XD\nAIU\nAID\nUNI\nCU\nSID\nSIU\nXU'
    r2_c24 = ws.cell(2, 24).value
    if r2_c24 is not None and isinstance(r2_c24, str) and "\n" in r2_c24 and len(r2_c24) > 5:
        rts_images = load_sheet_images_by_row(primary_file)
        return _parse_v1015_ready_to_ship(ws, vendor_id, images_by_row=rts_images)

    # 2026-09-04 BOS-MONCLER sub-format detect:R1 c3='ModelloVariante' 字符串
    r1_c3 = ws.cell(1, 3).value
    if r1_c3 is not None and isinstance(r1_c3, str) and r1_c3.strip().lower() == "modellovariante":
        bos_images = load_sheet_images_by_row(primary_file)
        return _parse_v1015_bos(ws, vendor_id, images_by_row=bos_images)

    # 2026-09-03 怀远 sub-format detect:R1 c12 是不是 size code 字符(短大写)
    # 老 V1015 (PRADA ITALY) R1 c12 大概率空;怀远的 R1 c12='SID'/'UNI' 等
    r1_c12 = ws.cell(1, 12).value
    if r1_c12 is not None and isinstance(r1_c12, str):
        code_candidate = r1_c12.strip()
        if code_candidate and len(code_candidate) <= 6 and re.match(r"^[A-Z0-9]+$", code_candidate):
            # 怀远 sub-format:用 load_sheet_images_by_row 拿全所有图
            hk_images = load_sheet_images_by_row(primary_file)
            return _parse_hk_huaiyuan(ws, vendor_id, images_by_row=hk_images)

    sheet_row_heights = load_sheet_row_heights(primary_file)
    # 2026-07-29:有 richValue 文件优先走 richValue(SharePoint 同步),否则用 drawing
    if has_richvalue_images(primary_file):
        images_by_row_1idx = load_richvalue_images(primary_file)
        # richValue 返回 1-idx row,转成 0-idx 跟 drawing 一致
        images_by_row = {r_1idx - 1: v for r_1idx, v in images_by_row_1idx.items() if r_1idx >= 1}
    else:
        images_by_row = _load_images_by_row(ws)

    # 转成 pandas 一次,方便按列名访问
    df = pd.read_excel(primary_file, sheet_name=ws.title, header=None)

    # 找尺码标签行
    size_label_row = _find_size_label_row(df)
    if size_label_row is None:
        return rows
    # 尺码行 c2 (0-idx 1) 的原始 cell:可能是 "SID\nSKU" (PRADA) 或 "XD\nUNI\nSID" (MIU MIU)
    size_row_c2 = (
        df.iat[size_label_row, 1] if 1 < df.shape[1] else ""
    )
    if not isinstance(size_row_c2, str):
        size_row_c2 = ""

    # 找所有产品 header 行:
    # - 0-idx col 0 (1-idx col 1) = total qty 正整数
    # - 0-idx col 2 (1-idx col 3) = 类别 token(短大写,如 SID / SKU / XD / UNI,
    #   也可能是多段 'SID\nSKU' —— 取第一段判断)
    # - 至少有一个 size 列有 qty (>0)
    # 不硬编码具体 token — 用"正整数 + 短大写字符串"的结构判断
    product_header_rows: List[int] = []
    for idx, row in df.iterrows():
        # 0-idx col 0 = total qty 整数
        total_qty = row.iloc[0] if len(row) > 0 else None
        # 0-idx col 2 = 类别 token(可能多段,如 'SID\nSKU',取第一段)
        cat_token_raw = row.iloc[2] if len(row) > 2 else None
        cat_token = (
            cat_token_raw.split("\n")[0].strip()
            if isinstance(cat_token_raw, str) and cat_token_raw.strip() else ""
        )
        # col 4+ 至少有一个 qty 数字(>=1)
        has_size_qty = any(
            isinstance(row.iloc[c], (int, float)) and row.iloc[c] > 0
            for c in _build_size_labels(df, size_label_row).keys() if c < len(row)
        )
        if (isinstance(total_qty, (int, float)) and total_qty > 0
                and cat_token
                and _looks_like_size_token(cat_token)
                and has_size_qty):
            product_header_rows.append(int(idx))

    # 逐产品解析
    for hdr_row in product_header_rows:
        # 1) 验证 header 行(应有 qty 数字)
        hdr = df.iloc[hdr_row]
        # 1b) 读产品类别 token(可能是多段 'SID\nSKU',取第一段)
        cat_token_raw = hdr.iloc[2] if len(hdr) > 2 else None
        cat_token = (
            cat_token_raw.split("\n")[0].strip()
            if isinstance(cat_token_raw, str) and cat_token_raw.strip() else ""
        )
        # 1c) 按 cat_token 选尺码段(欧码 / UK 码 / 裤码 / 均码 / 鞋码)
        seg_idx = _pick_seg_idx(cat_token, size_row_c2)
        size_labels = _build_size_labels(df, size_label_row, seg_idx=seg_idx)
        if not size_labels:
            continue
        has_qty = any(
            isinstance(hdr.iloc[c], (int, float)) and hdr.iloc[c] > 0
            for c in size_labels.keys() if c < len(hdr)
        )
        if not has_qty:
            continue

        # 2) SKU:header_row + 2
        sku_row_idx = hdr_row + 2
        if sku_row_idx >= len(df):
            continue
        sku = normalize_text(df.iat[sku_row_idx, SKU_COL]) if SKU_COL < df.shape[1] else ""
        if not sku:
            continue

        # 3) 描述 + 颜色:header_row + 3
        desc_row_idx = hdr_row + 3
        if desc_row_idx >= len(df):
            continue
        description = normalize_text(df.iat[desc_row_idx, DESC_COL]) if DESC_COL < df.shape[1] else ""
        color = normalize_text(df.iat[desc_row_idx, COLOR_COL]) if COLOR_COL < df.shape[1] else ""

        # 4) Cost / Retail:header_row + 6
        price_row_idx = hdr_row + 6
        if price_row_idx >= len(df):
            continue
        cost = to_float(df.iat[price_row_idx, COST_VAL_COL]) if COST_VAL_COL < df.shape[1] else None
        retail = to_float(df.iat[price_row_idx, RETAIL_VAL_COL]) if RETAIL_VAL_COL < df.shape[1] else None

        # 5) 图片行:openpyxl 0-idx = hdr_row + 2(图片锚在 SKU 行,
        #    SKU 行 = header_row + 2 in 0-idx)
        image_row_idx = hdr_row + 2
        row_images = images_by_row.get(image_row_idx, [])

        # 6) 源行高(取 header 行的,做参考)
        source_row_h = sheet_row_heights.get(ws.title, {}).get(hdr_row + 1)  # 1-idx

        # 6b) 性别(2026-07-31 新增):找 hdr_row 上方最近的 "Donna/Uomo/Man/Woman" 标题
        gender = _find_gender_for_header(df, hdr_row)

        # 7) 按尺码展开
        for col_idx, size_label in size_labels.items():
            if col_idx >= len(hdr):
                continue
            qty_raw = hdr.iloc[col_idx]
            if not is_positive_qty(qty_raw):
                continue
            qty = int(round(float(to_float(qty_raw) or 0)))
            if qty <= 0:
                continue
            rows.append({
                "brand": "",  # 不写死:V1015 是多品牌工厂代码(PRADA / STONE ISLAND 等),让 main 走文件名兜底
                "sku": sku,
                "image": "",
                "description": description,
                "category": guess_product_group(description),
                "material": "",
                "color": color,
                "gender": gender,  # 2026-07-31:从 c1 标题行 "Donna - PRADA" / "Uomo - PRADA" 提取
                "season": "",
                "origin": "ITALY",  # 源是意大利线,默认 ITALY
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": source_row_h,
            })

    return rows
