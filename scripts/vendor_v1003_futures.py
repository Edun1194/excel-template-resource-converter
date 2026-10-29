"""V1003 McQueen SS27 期货 vendor adapter (2026-09-29, 2026-10-21 扩展).

适用文件:
  - McQueen SS27 .xls / .xlsx (MAIN MEN/WOMEN ALL CAT 等)
  - McQueen SS27 .pdf (line sheet,需先用 scripts/preprocess_mcqueen_pdf.py 转 .xlsx)
  - Armani 系 PLM Report (2026-10-21): sheet 'PLM - Report', 19 列
  - LONGCHAMP SS27 期货 (2026-10-21 从 vendor_v1003_longchamp.py 合并):
    R1 header, C2='COMPOSIZIONE', 19 列

  期货模版约定: qty 留空等客户填 (LONGCHAMP 例外 — TU 列有真实数字,直接写)

源文件结构 (.xls):
  R14: header (PIC, PIC NUMBER, ..., SKU, MATERIAL, ..., SIZES, WHL PRICE, RTL PRICE, SIZE, [qty cols 18-33], TOT.QTY, TOTAL AMOUNT, ...)
  R15+: data rows,每行一个 SKU
  关键列:
    C4  = GENDER (WOMEN/MEN)
    C5  = CATEGORY (ACC/LG/OTHER/RTW/SHOES/EYEWEAR)
    C9  = DESCRIPTION
    C10 = SKU
    C12 = MATERIAL DESCRIPTION
    C13 = COLOR DESCRIPTION
    C14 = SIZES (semicolon-separated: "U;" / "S;M;L;" / "36;38;40;...")
    C15 = WHL PRICE (cost)
    C16 = RTL PRICE (retail)
    C17 = SIZE code (A/H/T/U/F/C/I - scalarino letter)

  注: 本文件是 line sheet/catalog,**没有实际库存数据** (TOT.QTY 都是 0)。
      每 SKU 输出一行,size 字段 = C17 scalarino letter,qty 占位为 1。

输出字段:
  brand / sku / description / category / material / gender / color
  season / origin / cost / cplus_text / supply_price / retail_price
  size / quantity / vendor_id / image_items / source_row_height
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl  # for .xlsx
import pandas as pd  # for .xls (xlrd backend)

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


VENDOR_ID = "V1003"


# ---------- workbook 缓存 (2026-10-26 性能优化) ----------
# detect 阶段有 6 个版式判断, 每个都重新 load_workbook 一次**同一个文件**,
# 命中版式后对应的 parse 函数还要再 load 一次 —— 同一份 xlsx 被完整解析 3~4 次。
# 大 xlsx 解析极贵 (实测 Boss&Hugo 28 个文件: 97 次 load ≈ 9.4s, 而真正的解析逻辑
# 只有 0.4s —— 几乎全是 load 开销)。改成 load 一次, detect 与 parse 共用。
_WB_CACHE: Dict[str, Any] = {}


def _load_wb(path: Path, data_only: bool = True):
    """加载 workbook (带缓存). 一次解析里同一个文件只真正 load 一次."""
    key = f"{path}|{data_only}"
    wb = _WB_CACHE.get(key)
    if wb is None:
        wb = openpyxl.load_workbook(path, data_only=data_only)
        _WB_CACHE[key] = wb
    return wb


def _clear_wb_cache() -> None:
    """关掉并清空缓存 (一次解析结束后调用, 避免文件句柄 / 内存泄漏)."""
    for wb in _WB_CACHE.values():
        try:
            wb.close()
        except Exception:
            pass
    _WB_CACHE.clear()

# Scalarino letter → English category mapping (McQueen .xls C17)
SIZE_CODE_MAP: Dict[str, str] = {
    "A": "A",   # clothing sizes (XXS-XXXL / 36-50 / etc.)
    "H": "H",   # clothing size grid
    "T": "T",   # UNICA (accessories)
    "U": "U",   # CALZATURE UNISEX
    "F": "F",   # CALZATURE (USA)
    "G": "G",   # CALZATURE (UK)
    "C": "C",   # XXXS-XXXL
    "I": "I",   # 服饰大码 / special
}

# C5 (NOTE) → category 映射 (Italian → English)
CATEGORY_MAP: Dict[str, str] = {
    "ACC": "Accessories",
    "LG": "Leather Goods",
    "OTHER": "Other",
    "RTW": "Ready to Wear",
    "SHOES": "Shoes",
    "EYEWEAR": "Eyewear",
}

# GENDER translation (col C4)
GENDER_MAP: Dict[str, str] = {
    "WOMEN": "Woman",
    "MEN": "Man",
    "UNISEX": "Unisex",
    "BAMBINO": "Kid",
    "BAMBINA": "Kid",
}


def _translate_gender(raw: str) -> str:
    return GENDER_MAP.get(raw.strip().upper(), raw.strip())


def _is_longchamp_format(df) -> bool:
    """Detect LONGCHAMP file format (header at R1, COMPOSIZIONE at C2)."""
    if df.shape[0] < 1 or df.shape[1] < 3:
        return False
    # LONGCHAMP: header is at R1 (not R14 like McQueen), C2 = 'COMPOSIZIONE'
    val_c2 = str(df.iat[0, 2] if df.shape[1] > 2 else "").strip().upper()
    return val_c2 == "COMPOSIZIONE"


def _parse_mcqueen_xls(xls_path: Path, sku_image_map: Optional[Dict[str, bytes]] = None) -> List[Dict[str, object]]:
    """Parse McQueen .xls (xlrd backend via pandas) → list of row dicts.

    sku_image_map: optional {sku: image_bytes} from companion PDFs.
    """
    df = pd.read_excel(xls_path, header=None)
    rows: List[Dict[str, object]] = []

    # Find header row (has 'SKU' / 'SIZES' / 'SIZE' labels)
    header_row = None
    for r in range(min(40, len(df))):
        v = str(df.iat[r, 10]) if df.shape[1] > 10 else ""
        if v.strip() == "SKU":
            header_row = r
            break
    if header_row is None:
        return rows

    # Detect gender from filename or data
    fname_upper = xls_path.name.upper()
    if "WOMEN" in fname_upper or "DONNA" in fname_upper:
        file_gender = "Woman"
    elif "MEN" in fname_upper or "UOMO" in fname_upper:
        file_gender = "Man"
    else:
        file_gender = ""

    for r in range(header_row + 1, len(df)):
        sku = normalize_text(df.iat[r, 10] if df.shape[1] > 10 else "")
        if not sku:
            continue

        # Gender: prefer data row, fall back to filename
        raw_gender = normalize_text(df.iat[r, 4] if df.shape[1] > 4 else "")
        gender = _translate_gender(raw_gender)  # 没显式性别就留空,不按文件名兜底

        # Category from C5 (NOTE)
        cat_raw = normalize_text(df.iat[r, 5] if df.shape[1] > 5 else "")
        category = CATEGORY_MAP.get(cat_raw.upper(), cat_raw)

        description = normalize_text(df.iat[r, 9] if df.shape[1] > 9 else "")
        material = normalize_text(df.iat[r, 12] if df.shape[1] > 12 else "")
        color = normalize_text(df.iat[r, 13] if df.shape[1] > 13 else "")

        # SIZES (C14) - semicolon-separated actual size labels
        # C17 scalarino letter 只是 size group identifier(I=cinture 70-120, A=RTW 36-50 等)
        # size 字段必须用 C14 的实际尺码,不是 scalarino 字母
        sizes_raw = normalize_text(df.iat[r, 14] if df.shape[1] > 14 else "")
        sizes_list = [s.strip() for s in sizes_raw.split(";") if s.strip()]
        # 兼容 "U;" 这种单 size 行(UNI)
        if not sizes_list:
            sizes_list = ["U"]

        cost = to_float(df.iat[r, 15] if df.shape[1] > 15 else "")
        retail = to_float(df.iat[r, 16] if df.shape[1] > 16 else "")

        # 期货模版:每行 size 展开,qty 留空(待客户填)
        # 不应用 is_positive_qty 过滤 — McQueen 期货文件所有 qty=0,但仍要展开
        # 从 companion PDF 抽的图(同 SKU 共享一张图,所有 size 行都挂同一张)
        sku_img = sku_image_map.get(sku) if sku_image_map else None
        row_image_items = [{"bytes": sku_img, "format": "jpeg"}] if sku_img else []

        for size_label in sizes_list:
            rows.append({
                "brand": "MC QUEEN",
                "sku": sku,
                "image": "",
                "gender": gender,
                "description": description,
                "category": category,
                "material": material,
                "color": color,
                "season": "SS27",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": "",
                "retail_price": retail,
                "size": size_label,
                "quantity": None,   # 期货模版 qty 留空,等客户填
                "vendor_id": VENDOR_ID,
                "_sizes_list": sizes_list,  # debug only
                "image_items": row_image_items,
                "source_row_height": 92.0 if row_image_items else None,
                "_row_num": r + 1,
            })

    return rows


def _parse_mcqueen_xlsx(xlsx_path: Path) -> List[Dict[str, object]]:
    """Parse preprocessed McQueen PDF→xlsx (same schema as .xls)."""
    wb = _load_wb(xlsx_path)
    ws = wb.active

    # Find header row (has 'SKU' in col 11 = openpyxl 1-indexed)
    header_row = None
    for r in range(1, min(40, ws.max_row + 1)):
        v = ws.cell(r, 11).value  # col 10 in 0-idx = col 11 in 1-idx
        if v and str(v).strip() == "SKU":
            header_row = r
            break
    if header_row is None:
        return []

    # Load images: group by anchor row → data row mapping
    # (image anchor row is 0-idx, data rows start at header_row + 1)
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    try:
        for img in ws._images:
            anchor = img.anchor
            # openpyxl TwoCellAnchor: _from / _to with .row (0-idx)
            try:
                from_row_0idx = anchor._from.row
            except Exception:
                continue
            try:
                img_bytes = img._data()
            except Exception:
                continue
            fmt = (img.format or "jpeg").lower()
            if fmt == "jpg":
                fmt = "jpeg"
            # Map 0-idx row to 1-idx (Excel) row
            excel_row = from_row_0idx + 1
            images_by_row.setdefault(excel_row, []).append({
                "bytes": img_bytes,
                "format": fmt,
            })
    except Exception:
        pass

    # Detect gender from filename
    fname_upper = xlsx_path.name.upper()
    if "WOMEN" in fname_upper or "DONNA" in fname_upper:
        file_gender = "Woman"
    elif "MEN" in fname_upper or "UOMO" in fname_upper:
        file_gender = "Man"
    else:
        file_gender = ""

    rows: List[Dict[str, object]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        sku = normalize_text(ws.cell(r, 11).value)
        if not sku:
            continue

        raw_gender = normalize_text(ws.cell(r, 5).value)
        gender = _translate_gender(raw_gender)  # 没显式性别就留空,不按文件名兜底

        cat_raw = normalize_text(ws.cell(r, 6).value)
        category = CATEGORY_MAP.get(cat_raw.upper(), cat_raw)

        description = normalize_text(ws.cell(r, 10).value)
        material = normalize_text(ws.cell(r, 13).value)
        color = normalize_text(ws.cell(r, 14).value)

        sizes_raw = normalize_text(ws.cell(r, 15).value)
        sizes_list = [s.strip() for s in sizes_raw.split(";") if s.strip()]

        cost = to_float(ws.cell(r, 16).value)
        retail = to_float(ws.cell(r, 17).value)

        size_code = normalize_text(ws.cell(r, 18).value)

        # SIZES (C15) - semicolon-separated actual size labels
        # C18 scalarino letter 是 size group identifier (I=cinture 70-120, A=RTW 36-50, etc.)
        # size 字段必须用 C15 实际尺码;如果 C15 空(仅 PDF 预处理后),
        # fallback 到 C18 scalarino code
        sizes_raw = normalize_text(ws.cell(r, 15).value)
        sizes_list = [s.strip() for s in sizes_raw.split(";") if s.strip()]
        if not sizes_list:
            # PDF 预处理没写 C15,fallback 到 C18 scalarino 字母
            fallback = normalize_text(ws.cell(r, 18).value)
            sizes_list = [fallback] if fallback else ["U"]

        # Get images for this row
        row_image_items = images_by_row.get(r, [])

        # 期货模版:每行 size 展开,qty 留空(待客户填)
        for sz in sizes_list:
            rows.append({
                "brand": "MC QUEEN",
                "sku": sku,
                "image": "",
                "gender": gender,
                "description": description,
                "category": category,
                "material": material,
                "color": color,
                "season": "SS27",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": "",
                "retail_price": retail,
                "size": sz,
                "quantity": None,   # 期货模版 qty 留空,等客户填
                "vendor_id": VENDOR_ID,
                "_sizes_list": sizes_list,
                "image_items": row_image_items,
                "source_row_height": 92.0 if row_image_items else None,
                "_row_num": r,
            })

    return rows


def _build_sku_image_map(folder: Path) -> Dict[str, bytes]:
    """Extract product images from companion McQueen PDFs in the same folder.

    McQueen SS27 .xls files don't have embedded images, but the companion PDFs
    (RTW.pdf / SHOES.pdf / EYEWEAR.pdf / LG.pdf) contain product images.
    Build a {sku: image_bytes} mapping so .xls conversion can pick them up.

    Strategy:
      1. For each .pdf in folder matching "MC QUEEN SS27 MAIN {gender} *.pdf"
      2. Run preprocess_mcqueen_pdf.pdf_to_xlsx to a tmp file (gets cards + images)
      3. Walk the resulting xlsx rows: SKU at C11, image embedded at B{row}
      4. Collect into sku → image_bytes mapping
    """
    import subprocess
    sku_to_img: Dict[str, bytes] = {}
    pdf_files = sorted(folder.glob("MC QUEEN SS27 MAIN *.pdf"))
    if not pdf_files:
        return sku_to_img

    # Use the preprocessor module
    from preprocess_mcqueen_pdf import pdf_to_xlsx
    import tempfile

    for pdf in pdf_files:
        tmp_xlsx = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
                tmp_xlsx = Path(f.name)
            n = pdf_to_xlsx(pdf, tmp_xlsx)
            if n == 0:
                continue
            # Read embedded images
            import openpyxl as _opx
            wb = _opx.load_workbook(tmp_xlsx, data_only=True)
            ws = wb.active
            images_by_row: Dict[int, List[Dict]] = {}
            for img in ws._images:
                anchor = img.anchor
                try:
                    from_row_0idx = anchor._from.row
                except Exception:
                    continue
                try:
                    img_bytes = img._data()
                except Exception:
                    continue
                fmt = (img.format or "jpeg").lower()
                if fmt == "jpg":
                    fmt = "jpeg"
                excel_row = from_row_0idx + 1
                images_by_row.setdefault(excel_row, []).append({
                    "bytes": img_bytes,
                    "format": fmt,
                })
            # Walk data rows starting at row 15
            for r in range(15, ws.max_row + 1):
                sku = ws.cell(r, 11).value
                if not sku:
                    continue
                sku = str(sku).strip()
                if not sku:
                    continue
                imgs = images_by_row.get(r, [])
                if imgs:
                    # Use first image (col 3 in PDF preprocessor is the product image)
                    sku_to_img[sku] = imgs[0]["bytes"]
        except Exception:
            continue
        finally:
            if tmp_xlsx:
                try:
                    tmp_xlsx.unlink(missing_ok=True)
                except Exception:
                    pass

    return sku_to_img


def _parse_plm_report(
    xlsx_path: Path,
    vendor_id: str,
) -> List[Dict[str, object]]:
    """PLM Report 版式 (2026-10-21): LINESHEET EA FW27 - HOLIDAY.xlsx 等 Armani 系 PLM 导出。

    源文件结构 (sheet='PLM - Report', 19 列, R1 header):
      c1=YEAR / c2=SEASON / c3=SEGMENT / c4=COLLECTION / c5=PRODUCER /
      c6=MODEL / c7=PART / c8=COLOR / c9=FOTO(全空) / c10=CODICE_FGH /
      c11=NO / c12=CATEGORY / c13=PRODUCT / c14=SEASONALITY /
      c15=COD_CONS / c16=PRICE / c17=MADE IN / c18=SIZES / c19=COMPOSITION

    其他 sheet: 'File Mancanti' (缺图清单, 不解析), 'Foglio2'/'Foglio3' (空)

    映射决策 (2026-10-21 用户确认):
      - 货号 = MODEL-PART-COLOR (如 EM000030-TE29097-F1040)
      - 类目 = CATEGORY 原值 (如 '10-READY TO WEAR' / '14-SHOES')
      - 数量 = 0 (源文件无库存列)
      - SIZES 逗号分隔 → 每尺码一行; 'L 44,46...' / 'TU' 格式原样保留单行

    注: 放在 futures 模块 (V1003-期货) 而不是现货 V1003 — PLM Report 是期货货单。
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    if "PLM - Report" not in wb.sheetnames:
        return rows
    ws = wb["PLM - Report"]
    if ws.max_row < 2 or ws.max_column < 19:
        return rows

    # 图片: embedded media 挂在 'PLM - Report' sheet 上, 按 anchor row 取
    from convert_vendor_quote_shared import load_sheet_images_by_row
    images_by_row = load_sheet_images_by_row(xlsx_path)

    def cell(r_1idx: int, c_1idx: int):
        return ws.cell(r_1idx, c_1idx).value

    for r in range(2, ws.max_row + 1):
        model = normalize_text(cell(r, 6))       # MODEL
        part = normalize_text(cell(r, 7))        # PART
        color_code = normalize_text(cell(r, 8))  # COLOR
        if not part:
            continue

        # 货号 = MODEL-PART-COLOR
        sku_final = "-".join([x for x in (model, part, color_code) if x])

        segment = normalize_text(cell(r, 3))    # '302130-EMPORIO ARMANI EA - MAN'
        collection = normalize_text(cell(r, 4)) # 'EA0953-EMPORIO ARMANI MAN FW27'
        category = normalize_text(cell(r, 12))   # CATEGORY 原值
        product = normalize_text(cell(r, 13))    # PRODUCT (如 '021-SUIT')
        made_in = normalize_text(cell(r, 17))    # 'ITALY' / 'CHINA NOT'
        sizes_raw = normalize_text(cell(r, 18))  # 'XXS , XS , S , M'
        composition = normalize_text(cell(r, 19))
        retail = to_float(cell(r, 16))          # PRICE

        # 品牌: 从 COLLECTION 拆 (第 1 个 '-' 之后, 去掉尾部季节 + 性别/系列词)
        brand = ""
        if "-" in collection:
            brand = collection.split("-", 1)[1].strip()
            # 去掉尾部季节 (FW27 / FW 2027 / FW 27 / SS26 / AW27 ...)
            brand = re.sub(
                r"\s+(FW|SS|AW)\s*20?\d{2,4}$", "", brand, flags=re.IGNORECASE
            ).strip()
            # 去掉尾部性别词 + 系列描述 (MAN / WOMAN / MAN CAPSULE / WOMAN ESSENTIAL ...)
            brand = re.sub(
                r"\s+(MAN|WOMAN|WOMEN)\b.*$", "", brand, flags=re.IGNORECASE
            ).strip()
            # EA 简写 → EMPORIO ARMANI
            if brand in ("EA", "EA MAN", "EA WOMAN"):
                brand = "EMPORIO ARMANI"

        # 性别: 从 SEGMENT 取 (源文件 c3 有明确 WOMAN / MAN / DONNA 字样)
        gender = ""
        seg_upper = segment.upper()
        if "WOMAN" in seg_upper or "WOMEN" in seg_upper or "DONNA" in seg_upper:
            gender = "Woman"
        elif "MAN" in seg_upper:
            gender = "Man"

        # 产地: 去掉 ' NOT' 后缀 (如 'BULGARIA NOT' → 'BULGARIA')
        origin = made_in
        if origin.endswith(" NOT"):
            origin = origin[:-4].strip()

        # 材质: COMPOSITION 里取 Fabric 部分 (意大利语优先, 没有就取英文)
        material = ""
        if composition:
            if " - Fabric:" in composition:
                material = composition.split(" - Fabric:", 1)[1].strip()
            elif "Fabric:" in composition:
                material = composition.split("Fabric:", 1)[1].strip()
            elif composition.startswith("Tessuto:"):
                material = composition[len("Tessuto:"):].strip()

        description = product or ""

        # 季节: 从 COLLECTION 抽 (支持 'FW27' / 'FW 2027' / 'FW 27' / 'SS26' / 无季节)
        season = ""
        season_m = re.search(r"\b(FW|SS|AW)\s*(20)?(\d{2}|\d{4})\b", collection, re.IGNORECASE)
        if season_m:
            prefix = season_m.group(1).upper()
            year = season_m.group(3)
            season = f"{prefix}{year}"

        row_images = images_by_row.get(r - 1, [])

        # SIZES 拆行: 'XXS , XS , S , M' → 4 行
        # 多组格式 ('L 44,46...\nR 44,46...') 和 'TU' 保持单行
        size_list: List[str] = []
        if sizes_raw:
            if ";" in sizes_raw:
                size_list = [s.strip() for s in sizes_raw.split(";") if s.strip()]
            elif "," in sizes_raw and "\n" not in sizes_raw:
                size_list = [s.strip() for s in sizes_raw.split(",") if s.strip()]
            elif "\n" in sizes_raw:
                # 多行多组 (L/R/S 三段), 无法干净拆 → 整个当一个 size
                size_list = [sizes_raw.replace("\n", " / ")]
            else:
                size_list = [sizes_raw]

        if not size_list:
            size_list = [""]

        for size in size_list:
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
                "material": material,
                "color": "",           # 源文件只有 COLOR code, 无颜色名 → 留空
                "color_code": color_code,
                "gender": gender,
                "season": season,
                "origin": origin,
                "cost": None,           # 源文件无 cost 列
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size,
                "quantity": 0,          # 2026-10-21 用户确认: 源无库存 → 填 0
                "vendor_id": "V1003",   # 2026-10-21: 期货系统一写 V1003 (与 McQueen / LONGCHAMP 一致)
                "image_items": row_images,
                "source_row_height": 92.0 if row_images else None,
            })

    return rows


def _is_longchamp_format(ws) -> bool:
    """Detect LONGCHAMP SS27 期货文件 (R1 header, C2='COMPOSIZIONE')."""
    if ws is None or ws.max_row < 2 or ws.max_column < 3:
        return False
    return str(ws.cell(1, 3).value or "").strip().upper() == "COMPOSIZIONE"


def _format_date(v) -> str:
    """Format datetime to YYYY-MM-DD, or return string as-is."""
    if v is None:
        return ""
    import datetime
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def _parse_longchamp(xlsx_path: Path) -> List[Dict[str, object]]:
    """Parse LONGCHAMP SS27 期货 .xlsx (R1 header, 19 列).

    源文件结构 (1-idx):
      c1  = PHOTO (image anchor)
      c2  = NOTE COLORE
      c3  = COMPOSIZIONE (material)
      c4  = CO/SEASONAL (collection → category 原值)
      c5  = DESCRIZIONE
      c7  = STYLE (parent SKU)
      c8  = CODICE COLORE
      c9  = DESCRIZIONE COLORE (color)
      c10 = WHL PRICE (cost)
      c11 = PREZZO RETAIL
      c12 = LOCATION (origin)
      c13 = 1° CONSEGNA (first delivery date)
      c14 = 3° CONSEGNA (second delivery date)
      c15 = sku (full SKU)
      c16 = TU (per-size qty — 直接写 quantity, 为 0 也写 0)

    特殊规则 (2026-09-29 user):
      1. quantity 字段 = TU 列 (C16), 即使为 0 也写 0 (不是 None)
      2. category = CO/SEASONAL 列 (C4) 原值
      3. 单 size 商品 → size = "U"
      4. vendor_id 统一写 "V1003" (2026-10-21 用户决定, 与 McQueen / Armani 一致)
    """
    wb = _load_wb(xlsx_path)
    # Pick first sheet with enough rows/cols
    ws = None
    for sn in wb.sheetnames:
        candidate = wb[sn]
        if candidate.max_row > 5 and candidate.max_column > 5:
            ws = candidate
            break
    if ws is None:
        return []

    rows: List[Dict[str, object]] = []

    # 加载 embedded 图片 (LONGCHAMP .xlsx 自带 product 图)
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    try:
        for img in ws._images:
            anchor = img.anchor
            try:
                from_row_0idx = anchor._from.row
            except Exception:
                continue
            try:
                img_bytes = img._data()
            except Exception:
                continue
            fmt = (img.format or "jpeg").lower()
            if fmt == "jpg":
                fmt = "jpeg"
            excel_row = from_row_0idx + 1
            images_by_row.setdefault(excel_row, []).append({
                "bytes": img_bytes,
                "format": fmt,
            })
    except Exception:
        pass

    for r in range(2, ws.max_row + 1):
        # C15 = full SKU, 跳过空行
        full_sku = normalize_text(ws.cell(r, 15).value)
        if not full_sku:
            continue

        material = normalize_text(ws.cell(r, 3).value)    # COMPOSIZIONE
        category = normalize_text(ws.cell(r, 4).value)    # CO/SEASONAL
        description = normalize_text(ws.cell(r, 5).value)
        style = normalize_text(ws.cell(r, 7).value)        # STYLE
        color_code = normalize_text(ws.cell(r, 8).value)   # CODICE COLORE
        color = normalize_text(ws.cell(r, 9).value)        # DESCRIZIONE COLORE
        cost = to_float(ws.cell(r, 10).value)              # WHL
        retail = to_float(ws.cell(r, 11).value)            # PREZZO RETAIL
        origin = normalize_text(ws.cell(r, 12).value)       # LOCATION

        delivery_date_1_str = _format_date(ws.cell(r, 13).value)
        delivery_date_2_str = _format_date(ws.cell(r, 14).value)

        # TU 列 (C16) = 数量, 即使为 0 也写 0
        qty_raw = ws.cell(r, 16).value
        try:
            quantity = int(float(qty_raw or 0))
        except (TypeError, ValueError):
            quantity = 0

        # LONGCHAMP single-size 商品 → size = "U"
        size_label = "U"

        row_image_items = images_by_row.get(r, [])

        rows.append({
            "brand": "LONGCHAMP",
            "sku": full_sku,
            "image": "",
            "gender": "",
            "description": description,
            "category": category,
            "material": material,
            "color": color,
            "color_code": color_code,
            "season": "SS27",
            "origin": origin,
            "cost": cost,
            "cplus_text": "",
            "supply_price": "",
            "retail_price": retail,
            "size": size_label,
            "quantity": quantity,
            "vendor_id": "V1003",   # 2026-10-21: 统一写 V1003
            "_style": style,
            "delivery_date_1": delivery_date_1_str,
            "delivery_date_2": delivery_date_2_str,
            "image_items": row_image_items,
            "source_row_height": 92.0 if row_image_items else None,
            "_row_num": r,
        })

    return rows


def _parse_mammut_linesheet(xlsx_path: Path) -> List[Dict[str, object]]:
    """MAMMUT 德语 B2B 货单 (2026-10-21): LINESHEET MAMMUT MAIN FW27.xlsx。

    源文件结构: 多 sheet 按上市日期分 ('02-08-2027' / '02-09-2027' / '02-10-2027'),
    每个 sheet 20 列 R1 header, **每行 = 一个 SKU + 一个尺码**。

    列映射 (1-idx):
      c1  = Picture (image anchor)
      c2  = SKU (style-color-size, 直接用)
      c3  = EAN Code (条码, 不进模板)
      c4  = Produktname (描述)
      c5  = Farbe (颜色, 格式 '19781 brunsina' → 色号+颜色名)
      c6  = Aktivität (Mountaineering / Hiking / Climbing, 不进模板)
      c7  = Produkthauptgruppe (Apparel, 太粗不用于类目)
      c8  = Produktgruppe (类目: Hardshell Jackets / Pants / ...)
      c9  = Warengruppe (备用类目)
      c10 = Geschlecht (Male / Female / Child)
      c11 = NOS (Ja/Nein, 不进模板)
      c12 = Größe (尺码)
      c14 = Liefertermin-Slots (交付周期, 不进模板)
      c15 = Lieferfenster (交付日期, 不进模板)
      c16 = Menge (数量 — **源文件全空**)
      c17 = EK (成本)
      c18 = UVP (零售价)
      c19 = Gesamt EK (合计成本, 不用)
      c20 = Gesamt UVP (合计零售价, 不用)

    映射决策 (2026-10-21 用户确认):
      - 3 个 sheet 全转
      - 数量留空 (null) — 源文件 Menge 列全空, 等客户填
      - 类目 = Produktgruppe (c8)
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row < 2 or ws.max_column < 12:
            continue

        # R1 必须是 'SKU' + 'Produktname' + 'Größe' 才认 (防止误吃其他 sheet)
        h2 = str(ws.cell(1, 2).value or "").strip()
        h4 = str(ws.cell(1, 4).value or "").strip()
        h12 = str(ws.cell(1, 12).value or "").strip()
        if h2 != "SKU" or h4 != "Produktname" or h12 != "Größe":
            continue

        # 图片: 逐 sheet 抽 embedded media
        images_by_row: Dict[int, List[Dict[str, object]]] = {}
        try:
            for img in ws._images:
                anchor = img.anchor
                try:
                    from_row_0idx = anchor._from.row
                except Exception:
                    continue
                try:
                    img_bytes = img._data()
                except Exception:
                    continue
                fmt = (img.format or "jpeg").lower()
                if fmt == "jpg":
                    fmt = "jpeg"
                excel_row = from_row_0idx + 1
                images_by_row.setdefault(excel_row, []).append({
                    "bytes": img_bytes,
                    "format": fmt,
                })
        except Exception:
            pass

        for r in range(2, ws.max_row + 1):
            sku = normalize_text(ws.cell(r, 2).value)   # SKU
            if not sku:
                continue

            description = normalize_text(ws.cell(r, 4).value)   # Produktname
            farbe = normalize_text(ws.cell(r, 5).value)         # Farbe '19781 brunsina'
            category = normalize_text(ws.cell(r, 8).value)      # Produktgruppe
            geschlecht = normalize_text(ws.cell(r, 10).value)   # Geschlecht
            size = normalize_text(ws.cell(r, 12).value)         # Größe
            cost = to_float(ws.cell(r, 17).value)              # EK
            retail = to_float(ws.cell(r, 18).value)            # UVP

            # 性别: Male → Man / Female → Woman / Child → Kid (源文件有明确值, 不推断)
            gender = {"MALE": "Man", "FEMALE": "Woman", "CHILD": "Kid"}.get(
                geschlecht.upper(), geschlecht
            )

            # 颜色: '19781 brunsina' → 色号='19781', 颜色名='brunsina'
            color = ""
            color_code = ""
            if farbe:
                parts = farbe.split(" ", 1)
                if len(parts) == 2 and parts[0].isdigit():
                    color_code = parts[0]
                    color = parts[1]
                else:
                    color = farbe

            row_image_items = images_by_row.get(r, [])

            rows.append({
                "brand": "MAMMUT",
                "sku": sku,
                "image": "",
                "description": description,
                "category": category,
                "material": "",
                "color": color,
                "color_code": color_code,
                "gender": gender,
                "season": "FW27",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size,
                "quantity": None,   # 2026-10-21 用户确认: Menge 全空 → 留空等客户填
                "vendor_id": "V1003",
                "image_items": row_image_items,
                "source_row_height": 92.0 if row_image_items else None,
                "_sheet": sheet_name,
                "_row_num": r,
            })

    return rows


def _parse_mammut_price_list(xlsx_path: Path) -> List[Dict[str, object]]:
    """MAMMUT 意大利 PRICE LIST (2026-10-21): PRICE LIST MAMMUT MAIN FW27.xlsx。

    源文件结构 (sheet='Tabelle1', 6 列):
      R1: 标题 '     MAMMUT FW2728'
      R2-R6: 空
      R7: header (Categoria / Gruppo Prodotto / Codice Art. /
                  Descrizione Articolo / Prezzo WHS EUR / Prezzo Pubblico EUR)
      R8+: data rows, **每行一个 SKU (无尺码维度)**

    映射:
      c1 = Categoria (Apparel / Climbing Gear / Packs and Bags …)
      c2 = Gruppo Prodotto (类目: Hardshell Jackets / Harnesses …)
      c3 = Codice Art. (SKU, 如 1010-30660)
      c4 = Descrizione Articolo (描述, 内含性别词如 '... Men' / '... Women')
      c5 = Prezzo WHS EUR (cost)
      c6 = Prezzo Pubblico EUR (retail)

    映射决策 (2026-10-21 用户确认): 独立转换, 不与 LINESHEET 合并。
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    ws = wb["Tabelle1"] if "Tabelle1" in wb.sheetnames else wb.active
    if ws.max_row < 8 or ws.max_column < 6:
        return rows

    # R7 必须是 'Categoria' 才认
    if str(ws.cell(7, 1).value or "").strip() != "Categoria":
        return rows

    for r in range(8, ws.max_row + 1):
        sku = normalize_text(ws.cell(r, 3).value)   # Codice Art.
        if not sku:
            continue

        description = normalize_text(ws.cell(r, 4).value)
        category = normalize_text(ws.cell(r, 2).value)   # Gruppo Prodotto
        cost = to_float(ws.cell(r, 5).value)             # WHS
        retail = to_float(ws.cell(r, 6).value)           # Pubblico

        # 性别: 描述尾部有明确 Men / Women / Kids 字样 (源文件有, 不推断)
        gender = ""
        desc_u = description.upper()
        for kw, g in ((" WOMEN", "Woman"), ("WOMEN", "Woman"), (" MEN", "Man"),
                      ("MEN", "Man"), (" KIDS", "Kid"), ("KIDS", "Kid"),
                      (" JUNIOR", "Kid"), ("UNISEX", "Unisex")):
            if desc_u.endswith(kw):
                gender = g
                break

        rows.append({
            "brand": "MAMMUT",
            "sku": sku,
            "image": "",
            "description": description,
            "category": category,
            "material": "",
            "color": "",           # 源文件无颜色列
            "color_code": "",
            "gender": gender,
            "season": "FW27",
            "origin": "",
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            # 2026-10-22 用户纠正: 源文件**无尺码列**, 不凭空造 'U' → 留空
            "size": "",
            "quantity": None,      # 源文件无数量列
            "vendor_id": "V1003",
            "image_items": [],
            "source_row_height": None,
            "_row_num": r,
        })

    return rows


# MAMMUT LOOKBOOK 真货号格式: style-color, 6 位-5 位 (1010-33030 / 1011-02680)
_MAMMUT_LOOKBOOK_SKU_RE = re.compile(r"(?<!\d)(\d{4,6}-\d{5})(?!\d)")
# 品类路径: 'Apparel / Eiger Extreme / Hardshell Pants'
_MAMMUT_PATH_RE = re.compile(r"^([A-Z][A-Za-z ]{2,30})/([^/]{2,32})/([^/]{2,32})$")
# 颜色行: '0001 - black' / '19783 - alpina'
_MAMMUT_COLOR_RE = re.compile(r"^(\d{3,5})\s*-\s*([A-Za-z][A-Za-z ]{2,24})")


def _parse_mammut_lookbook(pdf_path: Path) -> List[Dict[str, object]]:
    """MAMMUT LOOKBOOK 图册 (2026-10-21): LOOKBOOK MAMMUT MAIN FW27.pdf。

    109 页 A3 横向画册, **每页左右双栏** (跨页拼版), 每栏上下排 2-3 张 product card:

      - 产品头: 'Eiger Nordwand Pro HS Pants Women 1020-14080' (产品名 + SKU 同一 text block)
      - 品类路径: 'Apparel / Eiger Extreme / Hardshell Pants'
      - 产品图: native >= 150x150 的 image XObject, 位于产品头左侧 ~35pt、同 y
      - 颜色行: '0001 - black' + 尺码 + 交付窗口

    实现要点 (2026-10-21):
      - **必须用 PyMuPDF (fitz)**, 不能用 pdfplumber:
        pdfplumber 的 chars 层每个字符渲染两次 (叠加), 直接抽全是
        'Aparel' / 'Hoded' 这类丢字残缺文本; fitz 的 get_text('blocks')
        返回干净文本 ('Eiger Nordwand Pro SO Hooded Jacket Men')。
      - 图片按「同半栏 + y 邻近(<22pt) + 位于文本左侧」三条件匹配, 命中率 100%。

    注意: 图册**没有 cost / retail / 数量**, 这些字段留空。
    """
    rows: List[Dict[str, object]] = []
    try:
        import fitz  # PyMuPDF
    except Exception:
        return rows

    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return rows

    seen = set()
    for page in doc:
        half = page.rect.width / 2.0

        # 1) 产品图候选 (native >= 150x150, 排除 88x88 的色卡缩略图)
        big_images: List[Dict[str, object]] = []
        try:
            for info in page.get_images(full=True):
                xref, nw, nh = info[0], info[2], info[3]
                if nw < 150 or nh < 150:
                    continue
                for rect in page.get_image_rects(xref):
                    big_images.append({"xref": xref, "rect": rect})
        except Exception:
            continue

        # 2) 文本 blocks (fitz 给干净文本)
        blocks: List[Dict[str, object]] = []
        try:
            for b in page.get_text("blocks"):
                txt = " ".join(str(b[4]).split())
                if txt:
                    blocks.append({
                        "x0": b[0], "y0": b[1], "text": txt,
                    })
        except Exception:
            continue

        # 3) 品类路径 / 颜色行的位置索引
        path_idx: List[Dict[str, object]] = []
        color_idx: List[Dict[str, object]] = []
        for b in blocks:
            t = b["text"]
            pm = _MAMMUT_PATH_RE.match(t)
            if pm:
                path_idx.append({
                    "y0": b["y0"], "x0": b["x0"],
                    "top": pm.group(1), "mid": pm.group(2), "low": pm.group(3),
                })
                continue
            cm = _MAMMUT_COLOR_RE.match(t)
            if cm and not _MAMMUT_LOOKBOOK_SKU_RE.search(t):
                color_idx.append({
                    "y0": b["y0"], "x0": b["x0"],
                    "code": cm.group(1), "name": cm.group(2).strip(),
                })

        def _nearest(lst, y0, left_side, max_dy=90.0):
            best, best_score = None, None
            for it in lst:
                dy = abs(it["y0"] - y0)
                if dy > max_dy:
                    continue
                same_side = (it["x0"] >= half) == (not left_side)
                score = dy + (0 if same_side else 500)
                if best_score is None or score < best_score:
                    best, best_score = it, score
            return best

        # 4) 产品头 (产品名 + SKU)
        for b in blocks:
            t = b["text"]
            m = _MAMMUT_LOOKBOOK_SKU_RE.search(t)
            if not m:
                continue
            if len(t) - m.end() > 8:      # SKU 必须在块尾
                continue
            desc = t[: m.start()].strip()
            if len(desc) < 8 or not re.match(r"^[A-Z][A-Za-z]", desc):
                continue

            sku = m.group(1)
            if sku in seen:
                continue
            seen.add(sku)

            # 品类 (产品头下方最近的路径行)
            category = ""
            gender = ""
            path = _nearest(path_idx, b["y0"], b["x0"] < half)
            if path:
                category = path["low"].strip()
                for kw, g in (("Women", "Woman"), ("Men", "Man"),
                              ("Kids", "Kid"), ("Junior", "Kid"),
                              ("Unisex", "Unisex")):
                    if kw in path["mid"] or kw in path["top"]:
                        gender = g
                        break

            # 性别兜底: 产品名尾部
            for kw, g in (("Women", "Woman"), ("Women'", "Woman"),
                          (" Men", "Man"), (" Kids", "Kid"), (" Kids'", "Kid"),
                          (" Junior", "Kid"), ("Unisex", "Unisex")):
                if desc.endswith(kw) or (" " + kw) in desc[-16:]:
                    gender = g
                    break

            # 颜色
            color = ""
            color_code = ""
            col = _nearest(color_idx, b["y0"], b["x0"] < half, max_dy=60.0)
            if col:
                color_code = col["code"]
                color = col["name"]

            # 产品图 (同半栏 + y 邻近 + 位于文本左侧)
            image_items: List[Dict[str, object]] = []
            side = b["x0"] >= half
            cands = [
                im for im in big_images
                if abs(im["rect"].y0 - b["y0"]) < 22
                and (im["rect"].x0 >= half) == side
                and im["rect"].x1 <= b["x0"] + 8
            ]
            if cands:
                cands.sort(key=lambda im: abs(im["rect"].y0 - b["y0"]))
                try:
                    data = doc.extract_image(cands[0]["xref"])
                    image_items = [{
                        "bytes": data["image"],
                        "format": (data.get("ext") or "jpeg").lower(),
                    }]
                except Exception:
                    image_items = []

            rows.append({
                "brand": "MAMMUT",
                "sku": sku,
                "image": "",
                "description": desc,
                "category": category,
                "material": "",
                "color": color,
                "color_code": color_code,
                "gender": gender,
                "season": "FW27",
                "origin": "",
                "cost": None,           # 图册无价格
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,   # 图册无价格
                # 2026-10-22 用户纠正: 图册无尺码列, 不凭空造 'U' → 留空
                "size": "",
                "quantity": None,       # 图册无数量
                "vendor_id": "V1003",
                "image_items": image_items,
                "source_row_height": 92.0 if image_items else None,
            })

    return rows


# Calvin Klein NOS PW27 期货订货表 (2026-10-22)
_CK_SIZE_COL_START = 7      # c7-c18 = 尺码列 (c19 是 'Total Quantities', 不算尺码)
_CK_SIZE_COL_END = 18


# Calvin Klein「NO IMAGE」占位图 (2026-10-22)
# 源文件里没上图的款会放一张 1600x1600 的 Calvin Klein NO IMAGE 占位图,
# 缩放后仍是 123x123, 在输出里会盖住多行, 视觉上像"巨大图片"。
# 纯白/灰字检测不可靠 (非白像素占比会算到 1.0), 所以按**字节指纹**精确匹配。
_CK_NO_IMAGE_MD5 = {
    "0a34e61b5fda3bbf6a18820dbb54d59f",   # Calvin Klein NO IMAGE 1600x1600
}

# 图片显示上限: 行高 92pt (framework 默认) → EMU
_CK_IMG_MAX_H_EMU = int(92.0 * 12700)


def _is_placeholder_image(blob: bytes) -> bool:
    """判断是否 NO IMAGE 占位图 (按 md5 指纹)."""
    if not isinstance(blob, (bytes, bytearray)):
        return True
    try:
        import hashlib
        return hashlib.md5(bytes(blob)).hexdigest() in _CK_NO_IMAGE_MD5
    except Exception:
        return False


def _fit_image_emu(w_px, h_px, max_h_emu: int = _CK_IMG_MAX_H_EMU):
    """按行高等比缩放 → (cx, cy) EMU; 不放大只缩小."""
    try:
        w_px, h_px = int(w_px), int(h_px)
    except (TypeError, ValueError):
        return None
    if w_px <= 0 or h_px <= 0:
        return None
    cx, cy = w_px * 9525, h_px * 9525
    if cy > max_h_emu:
        s = max_h_emu / cy
        cx, cy = int(cx * s), max_h_emu
    return cx, cy


def _parse_eu_price(v) -> Optional[float]:
    """CK 的 RRP 是欧式写法: '34,9' → 34.9, '1.234,56' → 1234.56."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 4)
    s = str(v).strip().replace("€", "").replace("EUR", "").strip()
    if not s:
        return None
    if "," in s and "." in s:
        # 点号是千分位, 逗号是小数
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return round(float(s), 4)
    except ValueError:
        return None


def _parse_calvin_klein(xlsx_path: Path) -> List[Dict[str, object]]:
    """Calvin Klein NOS 期货订货表 (2026-10-22): CK UNDERW NOS PW27- ORDER FORM.xlsx。

    源文件结构 (2 sheets):
      - **'Options'** (152 产品, 152 图): **本函数只用这个 sheet**
          R1  : header (c1=Image / c2=Image Link / c3=Season / c4=Division / c5=Gender /
                 c6=Enter Quantities / c19=Total Quantities / c23=Wholesale Price /
                 c24=Recommended Retail Price / c27=Label / c29=Sub Category /
                 c30=Style Name / c31=Style Number / c32=Option Code /
                 c34=Color Code / c36=Color Description / c41=Fabric Content /
                 c44=Country of Origin / c45=Order Reference)
          R2  : 'Totals' 汇总行 (跳过)
          R3, R5, R7, ...: 产品行 (奇数行) — c7-c18 是尺码标签
          R4, R6, R8, ...: qty 行 (偶数行) — c7-c18 是对应尺码的数量
          图片: c1 列, anchor 就在产品行 (152/152 命中)

      - 'EANs' (1770 行, 尺码/EAN 明细) — 本函数**不使用**

    映射决策 (2026-10-22 用户确认):
      - 货号 = Option Code (c32), = Style+Color, 不含尺码 (如 000NM2748E100)
      - 类目 = Sub Category (c29)
      - 只用 Options sheet

    字段细节:
      - RRP (c24) 是**欧式小数逗号字符串** ('34,9'), 要清洗成 34.9
      - qty 全为 0 (期货单, 等客户填)
      - CLOSED 尺码: qty 单元格写 'CLOSED' 而非数字, 该尺码**跳过不输出**
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    if "Options" not in wb.sheetnames:
        return rows
    ws = wb["Options"]

    # detect: R1 必须是 CK 的特征列
    h32 = str(ws.cell(1, 32).value or "").strip()
    h23 = str(ws.cell(1, 23).value or "").strip()
    if h32 != "Option Code" or h23 != "Wholesale Price":
        return rows

    # 图片: c1 列 anchor 在产品行 (0-idx)
    # 2026-10-22: 过滤 NO IMAGE 占位图 + 预置 display_emu 防止超大图盖行
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    n_placeholder = 0
    try:
        for img in ws._images:
            try:
                row_0idx = img.anchor._from.row
            except Exception:
                continue
            try:
                img_bytes = img._data()
            except Exception:
                continue
            if _is_placeholder_image(img_bytes):
                n_placeholder += 1
                continue
            fmt = (img.format or "jpeg").lower()
            if fmt == "jpg":
                fmt = "jpeg"
            item = {"bytes": img_bytes, "format": fmt}
            emu = _fit_image_emu(img.width, img.height)
            if emu:
                item["display_emu"] = emu
            images_by_row.setdefault(row_0idx, []).append(item)
    except Exception:
        pass
    if n_placeholder:
        print(f"[CK] 跳过 {n_placeholder} 张 NO IMAGE 占位图")

    for r in range(3, ws.max_row + 1):
        option_code = normalize_text(ws.cell(r, 32).value)   # c32 Option Code
        if not option_code:
            continue

        brand = normalize_text(ws.cell(r, 27).value) or "Calvin Klein"   # c27 Label
        gender_raw = normalize_text(ws.cell(r, 5).value)                # c5 Gender
        category = normalize_text(ws.cell(r, 29).value)                 # c29 Sub Category
        style_name = normalize_text(ws.cell(r, 30).value)               # c30 Style Name
        division = normalize_text(ws.cell(r, 4).value)                  # c4 Division
        color_code = normalize_text(ws.cell(r, 34).value)               # c34 Color Code
        color = normalize_text(ws.cell(r, 36).value)                    # c36 Color Description
        material = normalize_text(ws.cell(r, 41).value)                # c41 Fabric Content
        origin = normalize_text(ws.cell(r, 44).value)                  # c44 Country of Origin
        season = normalize_text(ws.cell(r, 3).value)                   # c3 Season ('NOS 9999')

        cost = _parse_eu_price(ws.cell(r, 23).value)                   # c23 Wholesale Price
        retail = _parse_eu_price(ws.cell(r, 24).value)                 # c24 RRP (欧式逗号)

        # 性别: Men/Women → Man/Woman (源文件已是英文规范)
        gender = {"MEN": "Man", "WOMEN": "Woman", "UNISEX": "Unisex"}.get(
            gender_raw.upper(), gender_raw
        )

        # 描述: Style Name + Division (源文件没有更完整的描述字段)
        description = " ".join(x for x in (style_name, division) if x)

        # 季节: c3 是 'NOS 9999', 从 Order Reference (c45 'CKU NOS SS27') 抽真实季节
        order_ref = normalize_text(ws.cell(r, 45).value)
        season_out = ""
        sm = re.search(r"\b(SS|FW|AW|PW)\s?(\d{2,4})\b", order_ref, re.IGNORECASE)
        if sm:
            season_out = f"{sm.group(1).upper()}{sm.group(2)}"
        else:
            season_out = season

        row_images = images_by_row.get(r - 1, [])

        # 尺码: 产品行 (r) c7-c18 是标签, qty 行 (r+1) c7-c18 是数量
        qty_row = r + 1
        for c in range(_CK_SIZE_COL_START, _CK_SIZE_COL_END + 1):
            size_label = normalize_text(ws.cell(r, c).value)
            if not size_label:
                continue
            # qty 单元格可能是 'CLOSED' (该尺码已关闭) → 跳过
            qty_cell = ws.cell(qty_row, c).value
            if isinstance(qty_cell, str) and qty_cell.strip().upper() == "CLOSED":
                continue
            qty = to_float(qty_cell)

            rows.append({
                "brand": brand,
                "sku": option_code,
                "image": "",
                "description": description,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code,
                "gender": gender,
                "season": season_out,
                "origin": origin,
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": int(qty) if qty is not None else None,
                "vendor_id": "V1003",
                "image_items": row_images,
                "source_row_height": 92.0 if row_images else None,
                "_row_num": r,
            })

    return rows


# CK 尺码区间常量 (2026-10-22)
_CK_SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "2XL", "3XL", "4XL"]


# 尺码标签识别 (2026-10-22)
# 覆盖: XS/S/M/L/XL/XXL/2XL…5XL / 纯数字 36,38,40 / 带小数 8.5,10.5 /
#       数字+地区后缀 "10 UK" / 意大利码 3,4,5 / "U"(单码)
_CK_SIZE_TOKEN_RE = re.compile(
    r"^(?:"
    r"\d{1,2}(?:\.\d)?(?:X{0,4}L?)"        # 36 / 8.5 / 2XL / 3XL
    r"|X{1,5}(?:S|L)"                      # XS / XL / XXS / XXL / XXXXL
    r"|\d{1,2}\s?(?:UK|US|EU)"             # 10 UK
    r"|U(?:NICA)?|OS|ONE SIZE"             # 单码
    r"|[A-Z0-9]{1,3}(?:[/\-][A-Z0-9]{1,3}){1,3}"  # 2026-10-23: 复合码 XS/S / M/L / 6/8 / 35.5-36
    r")$",
    re.IGNORECASE,
)

# 多字母尺码 (XS/XXL/2XL) 特征强, 单字母要靠邻近判断
_CK_MULTI_SIZE_RE = re.compile(r"^(?:X{1,5}S|X{1,5}L|\d{1,2}X{1,4}L)$", re.IGNORECASE)
_CK_SINGLE_SIZE = {"XS", "S", "M", "L", "XL", "XXL", "U", "OS", "ONE SIZE", "T", "V"}

# 明确的终止标记 (遇到就停, 不算尺码)
_SIZE_STOP_WORDS = {
    "TOTAL", "TOT", "QTY", "QUANTITY", "TOTAL QTY", "TOTAL QUANTITY",
    "SUBTOTAL", "SUM", "备注", "合计", "小计",
}

# 明确不是尺码的常见表头 (防止 S/M/L 这种单字母误判)
_NOT_SIZE_WORDS = {
    "S",  # 只在孤立出现且不在连续段里时才算, 这里不做全局排除
}


def _is_size_token(v) -> bool:
    """判断表头单元格是否**强匹配**尺码 (多字母/数字型, 不会误判)."""
    s = normalize_text(v).upper()
    if not s or len(s) > 8:
        return False
    if s in _SIZE_STOP_WORDS:
        return False
    if not _CK_SIZE_TOKEN_RE.match(s):
        return False
    # 数字型 (36 / 8.5 / 10 UK / 2XL) 和多字母 (XS / XXL) 直接算尺码
    return True


def _is_weak_size_token(v) -> bool:
    """单字母尺码 (S / M / L) — 只能通过「连续段」确认, 不能单独成立."""
    s = normalize_text(v).upper()
    return s in {"S", "M", "L", "T", "V"}


def _detect_size_ladder(header_values) -> tuple:
    """从表头扫出「尺码梯度段」→ (起始列下标0-idx, 尺码梯度 list)。

    为什么要扫而不是写死列号 (2026-10-22):
      不同款式的尺码列位置/数量都不一样, 写死 c13-c18 只对 CK FASHION 成立:
        - 衣服 XS/S/M/L/XL/2XL        → c13-c18 (6 列)
        - 鞋   36/37/38/39/40/41/42/43 → c8-c15 (8 列)
      扫表头对两种都能用, 以后新增款式不用改代码。

    扫法:
      把表头每个单元格分成三类 ——
        strong (数字型/多字母型: 36 / 8.5 / 10 UK / XS / XL / 2XL / XXS / U) 特征明确
        weak   (单字母: S / M / L / T / V) 单独看不能确定是尺码
        other  (SEASON / STYLE# / RRP / TOTAL …) 一定不是
      然后取**最长的「strong|weak 连续区间」且其中至少含 1 个 strong**。
      这样 XS→S→M→L→XL 这种强弱交替的梯度能完整读出,
      而表头里孤零零一个叫 S 的字段不会被误认 (区间里没有 strong)。
    """
    n = len(header_values)
    strong = [_is_size_token(v) for v in header_values]
    weak = [_is_weak_size_token(v) for v in header_values]
    candidate = [strong[i] or weak[i] for i in range(n)]

    best_start, best_len = -1, 0
    cur_start, cur_len, cur_has_strong = -1, 0, False
    for i, ok in enumerate(candidate):
        if ok:
            if cur_len == 0:
                cur_start, cur_has_strong = i, False
            cur_len += 1
            cur_has_strong = cur_has_strong or strong[i]
        else:
            if cur_len > best_len and cur_has_strong:
                best_start, best_len = cur_start, cur_len
            cur_start, cur_len, cur_has_strong = -1, 0, False
    if cur_len > best_len and cur_has_strong:
        best_start, best_len = cur_start, cur_len

    if best_len == 0:
        return -1, []

    ladder = [normalize_text(v) for v in header_values[best_start:best_start + best_len]]
    return best_start, ladder


def _expand_size_range(size_range: str, ladder: Optional[List[str]] = None) -> List[str]:
    """'XS TO 2XL' → ['XS','S','M','L','XL','2XL']; 'S TO 2XL' → ['S','M','L','XL','2XL'].

    ladder: 该文件实际的尺码梯度 (**从源表头 c13-c18 读**, 2026-10-22)。
            CK FASHION 表头是 XS/S/M/L/XL/2XL — **没有 XXL**,
            所以不能用通用梯度表, 否则 'S TO 2XL' 会多展开一个 XXL。
            不传 ladder 时回退到通用表。

    单值 (如 'U') 原样返回。无法解析时返回 [size_range]。
    """
    order = ladder or _CK_SIZE_ORDER
    s = normalize_text(size_range).upper().replace("–", "-").strip()
    if not s:
        return []
    m = re.match(r"^([A-Z0-9]{1,4})\s*(?:TO|-)\s*([A-Z0-9]{1,4})$", s)
    if not m:
        return [normalize_text(size_range)]
    lo, hi = m.group(1), m.group(2)
    if lo not in order or hi not in order:
        # 端点不在本文件梯度里 (如 '10 UK') → 整段当一个尺码
        return [normalize_text(size_range)]
    i, j = order.index(lo), order.index(hi)
    if j < i:
        i, j = j, i
    return order[i:j + 1]


def _parse_calvin_klein_style(xlsx_path: Path) -> List[Dict[str, object]]:
    """Calvin Klein FASHION style 级订货表 (2026-10-22): SUPF27 CKU FASHION - ORDER FORM.xlsx。

    源文件结构 (sheet 名带性别, 如 'CKU M' / 'CKU W'):
      R1  : 空
      R2  : 空 (merged M2:N2)
      R3  : header — c1=SEASON / c2=PROGRAM(CTB & Acc only) / c3=STYLE # /
            c4=CAD / c5=STYLE DESCRIPTION / c6=PACK / c7=SIZE RANGE /
            c8=FINAL COLOR / c9=REQUESTED IN STORE DELIVERY DATE /
            c10=RRP / c11=WHSL PRICE / c12=ADDITIONAL COMMENTARY /
            c13-c18=XS/S/M/L/XL/2XL (**订货格, 全空**) / c19=TOTAL
      R4+ : 数据行, **一行 = 一个 style + 一个色way** (20 行, 20 图)

    与 _parse_calvin_klein (NOS PW27) 的区别:
      - NOS 版: option 级 (Style+Color 一个 SKU, c7-c18 是尺码标签 + 下一行 qty)
      - 本版: **style 级**, 尺码是文字区间 (SIZE RANGE = 'XS TO 2XL'), 尺码列全空

    映射决策 (2026-10-22 用户确认):
      - 货号 = STYLE# + FINAL COLOR (同一 STYLE# 有多个色way, 必须带颜色才唯一)
      - SIZE RANGE **展开成多行** (每尺码一行, 数量留空)
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    # sheet 名带性别: 'CKU M' / 'CKU W' / 'CKU F'
    sheet_name = ""
    for sn in wb.sheetnames:
        if sn.upper().startswith("CKU"):
            sheet_name = sn
            break
    if not sheet_name:
        return rows
    ws = wb[sheet_name]

    # detect: R3 c3 含 'STYLE #' 且 c11 = 'WHSL PRICE'
    h3 = str(ws.cell(3, 3).value or "").upper()
    h11 = str(ws.cell(3, 11).value or "").upper()
    if "STYLE #" not in h3 or h11 != "WHSL PRICE":
        return rows

    # 性别: 从 sheet 名取 (CKU M → Man, CKU W → Woman)
    tail = re.sub(r"^CKU\s*", "", sheet_name.upper()).strip()
    gender = {"M": "Man", "W": "Woman", "F": "Woman", "MEN": "Man",
              "WOMEN": "Woman", "KIDS": "Kid"}.get(tail, "")

    # 尺码梯度: **自动扫表头** (2026-10-22), 不写死 c13-c18。
    # 不同款式尺码列位置/数量不同 (衣服 c13-c18 六码 / 鞋 c8-c15 八码), 扫表头通用。
    header_vals = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    size_col_start, ladder = _detect_size_ladder(header_vals)
    if not ladder:
        ladder = list(_CK_SIZE_ORDER)
        size_col_start = -1

    # 图片: anchor 在数据行 (c1 列)
    # 2026-10-22: 同样过滤 NO IMAGE 占位图 + 预置 display_emu (防超大图盖行)
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    n_placeholder = 0
    try:
        for img in ws._images:
            try:
                row_0idx = img.anchor._from.row
            except Exception:
                continue
            try:
                img_bytes = img._data()
            except Exception:
                continue
            if _is_placeholder_image(img_bytes):
                n_placeholder += 1
                continue
            fmt = (img.format or "jpeg").lower()
            if fmt == "jpg":
                fmt = "jpeg"
            item = {"bytes": img_bytes, "format": fmt}
            emu = _fit_image_emu(img.width, img.height)
            if emu:
                item["display_emu"] = emu
            images_by_row.setdefault(row_0idx, []).append(item)
    except Exception:
        pass
    if n_placeholder:
        print(f"[CK] 跳过 {n_placeholder} 张 NO IMAGE 占位图")

    for r in range(4, ws.max_row + 1):
        style_no = normalize_text(ws.cell(r, 3).value)     # c3 STYLE #
        if not style_no:
            continue

        program = normalize_text(ws.cell(r, 2).value)      # c2 PROGRAM (面料/系列)
        style_desc = normalize_text(ws.cell(r, 5).value)   # c5 STYLE DESCRIPTION
        pack = normalize_text(ws.cell(r, 6).value)         # c6 PACK
        size_range = normalize_text(ws.cell(r, 7).value)   # c7 SIZE RANGE
        final_color = normalize_text(ws.cell(r, 8).value)  # c8 FINAL COLOR
        season_raw = normalize_text(ws.cell(r, 1).value)   # c1 SEASON ('SUPF27')

        rrp = _parse_eu_price(ws.cell(r, 10).value)         # c10 RRP
        cost = _parse_eu_price(ws.cell(r, 11).value)        # c11 WHSL PRICE

        # 货号 = STYLE# + FINAL COLOR (同 STYLE# 多色way, 必须唯一)
        sku = f"{style_no} - {final_color}" if final_color else style_no

        # 描述: STYLE DESCRIPTION + PACK + PROGRAM
        description = " ".join(x for x in (style_desc, pack, program) if x)

        # 季节: 'SUPF27' → 'PF27' (SUP = Supplier/Order 前缀, 去掉 SUP; F27 = PF27)
        season = ""
        sm = re.match(r"^SUP([A-Z]{1,2})(\d{2})$", season_raw.upper().strip())
        if sm:
            season = f"P{sm.group(1)}{sm.group(2)}"
        else:
            season = season_raw

        sizes = _expand_size_range(size_range, ladder)
        if not sizes:
            sizes = [""]

        row_images = images_by_row.get(r - 1, [])

        for size in sizes:
            rows.append({
                "brand": "Calvin Klein",
                "sku": sku,
                "image": "",
                "description": description,
                "category": "",           # 源文件无类目字段
                "material": program,      # c2 PROGRAM 是面料 (100% Cotton / ICON COTTON STRETCH)
                "color": final_color,     # c8 FINAL COLOR
                "color_code": "",
                "gender": gender,         # 从 sheet 名 'CKU M' 取
                "season": season,
                "origin": "",             # 源文件无产地字段
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": rrp,
                "size": size,
                "quantity": None,         # 订货格全空, 等客户填
                "vendor_id": "V1003",
                "image_items": row_images,
                "source_row_height": 92.0 if row_images else None,
                "_row_num": r,
            })

    return rows


# SIRIO DB 期货单 (Michael Kors 等) — 2026-10-22
# 两个变体:
#   1) whls-70-T27-40.xlsx    sheet 'A1 A2   SHOES 5 12' — **有尺码列 c20-c34** (鞋码 5.0~12.0)
#   2) WOMAN ACC (1).xlsx     sheet 'NS   NO SIZE'      — **无尺码列** (NS = No Size)
# 尺码列位置/数量两者不同 → 用 _detect_size_ladder 自动扫表头识别 (2026-10-22)
_SIRIO_SIRIO_HEADERS = ("picture", "id", "sku", "style name", "style number",
                        "color code", "col desc")


def _parse_sirio_whls(xlsx_path: Path) -> List[Dict[str, object]]:
    """SIRIO DB 期货订货单 (2026-10-22/23): Michael Kors 等品牌。

    源文件结构:
      - **数据 sheet** (N 个!, sheet 名带尺码段):
          2026-10-23 发现**一份导出里可能有多个数据 sheet**, 每个尺码梯度不同:
            '1V   XS S  M L WOM'      → 2 列 (复合码 XS/S / M/L)
            'M4   XXS XXXL'            → 8 列 (XXS~XXXL)
            'W0 WB   WOMEN S 00   18'  → 11 列 (数字码 00~18)
            'A1 A2   SHOES 5 12'       → 15 列 (鞋码 5.0~12.0)
            'NS   NO SIZE'             → 无尺码列
          **本函数处理全部匹配的数据 sheet**, 不是只取第一个。
          R1 header: c1=Picture / c2=ID / c4=SKU / c6=Product Subtype /
                     c8=Style Name / c9=Style Number / c10=Color Code /
                     c11=Col Desc / c12=Delivery / c13=Material(Group) /
                     c14=Main Material Composition /
                     **尺码列(位置/数量每个 sheet 不同)** /
                     末几列 = WHLS Eur / RRP Eur / TOT QTY / TOT WHLS / TOT RRP
      - 'Info' sheet: 每行一个 JSON (本函数不用)
      - 'Totals' sheet: 汇总 (跳过)

    字段说明:
      - 数量全为 0/空 (期货单, 等客户填)
      - 无尺码列的 sheet → 尺码**留空** (2026-10-22 用户纠正: 不伪造 'U')
      - 性别: 源文件无 gender 列, 从 sheet 名 / 文件名判断, 无则留空
      - 品牌: 源文件无 brand 列 → 留空, 由用户手填 (_no_filename_brand)
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    # 2026-10-23: **所有匹配的数据 sheet 都要处理** (一份导出可能含多个)
    data_sheets = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.max_row < 2 or ws.max_column < 12:
            continue
        hdr = [normalize_text(ws.cell(1, c).value).lower()
               for c in range(1, ws.max_column + 1)]
        if hdr[0].startswith("picture") and "sku" in hdr and "style number" in hdr:
            data_sheets.append(ws)
    if not data_sheets:
        return rows

    def _cols_of(ws) -> Dict[str, int]:
        """按表头名建索引 (不写死列号 — 每个 sheet 列数不同)。"""
        m: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            nm = normalize_text(ws.cell(1, c).value).lower()
            if nm and nm not in m:
                m[nm] = c
        return m

    def _pick(m, *names) -> int:
        for n in names:
            if n in m:
                return m[n]
        return 0

    total_placeholder = 0

    for data_ws in data_sheets:
        hdr_map = _cols_of(data_ws)
        c_sku = _pick(hdr_map, "sku")
        if not c_sku:
            continue
        c_style_name = _pick(hdr_map, "style name")
        c_color_code = _pick(hdr_map, "color code")
        c_col_desc = _pick(hdr_map, "col desc")
        c_material = _pick(hdr_map, "main material composition", "material")
        c_subtype = _pick(hdr_map, "product subtype")
        c_whls = _pick(hdr_map, "whls eur", "wholesale price", "whls")
        c_rrp = _pick(hdr_map, "rrp eur", "recommended retail price", "rrp")

        # 尺码列: **每个 sheet 独立扫表头** (位置/数量随款式变, 不能写死)
        header_vals = [data_ws.cell(1, c).value
                       for c in range(1, data_ws.max_column + 1)]
        size_col_start, ladder = _detect_size_ladder(header_vals)

        # 图片: c1 (Picture) 列, anchor 在数据行
        images_by_row: Dict[int, List[Dict[str, object]]] = {}
        n_placeholder = 0
        try:
            for img in data_ws._images:
                try:
                    row_0idx = img.anchor._from.row
                except Exception:
                    continue
                try:
                    img_bytes = img._data()
                except Exception:
                    continue
                if _is_placeholder_image(img_bytes):
                    n_placeholder += 1
                    continue
                fmt = (img.format or "jpeg").lower()
                if fmt == "jpg":
                    fmt = "jpeg"
                item = {"bytes": img_bytes, "format": fmt}
                emu = _fit_image_emu(img.width, img.height)
                if emu:
                    item["display_emu"] = emu
                images_by_row.setdefault(row_0idx, []).append(item)
        except Exception:
            pass
        total_placeholder += n_placeholder

        # 性别: 从 sheet 名 / 文件名找 WOMAN/MEN (源文件无 gender 列)
        blob = f"{data_ws.title} {xlsx_path.stem}".upper()
        gender = ""
        if "WOMAN" in blob or "WOMEN" in blob or " DONNA" in blob:
            gender = "Woman"
        elif re.search(r"\bMEN\b|\bMENS\b| UOMO", blob):
            gender = "Man"

        for r in range(2, data_ws.max_row + 1):
            sku = normalize_text(data_ws.cell(r, c_sku).value)
            if not sku:
                continue

            description = normalize_text(data_ws.cell(r, c_style_name).value) if c_style_name else ""
            color = normalize_text(data_ws.cell(r, c_col_desc).value) if c_col_desc else ""
            color_code = normalize_text(data_ws.cell(r, c_color_code).value) if c_color_code else ""
            material = normalize_text(data_ws.cell(r, c_material).value) if c_material else ""
            category = normalize_text(data_ws.cell(r, c_subtype).value) if c_subtype else ""
            cost = to_float(data_ws.cell(r, c_whls).value) if c_whls else None
            retail = to_float(data_ws.cell(r, c_rrp).value) if c_rrp else None

            row_images = images_by_row.get(r - 1, [])

            def _emit(size_label: str, qty_cell):
                q = to_float(qty_cell)
                rows.append({
                    # 2026-10-22 用户确认: 源文件无品牌列且文件名是内部代号,
                    # 禁用文件名兜底 → 品牌留空, 用户在 App 手填
                    "brand": "",
                    "sku": sku,
                    "image": "",
                    "description": description,
                    "category": category,
                    "material": material,
                    "color": color,
                    "color_code": color_code,
                    "gender": gender,
                    "season": "",          # 源文件无季节列
                    "origin": "",          # 源文件无产地列
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size_label,
                    "quantity": int(q) if q is not None else None,
                    "_no_filename_brand": True,
                    "vendor_id": "V1003",
                    "image_items": row_images,
                    "source_row_height": 92.0 if row_images else None,
                    "_row_num": r,
                })

            if ladder:
                for idx, size_label in enumerate(ladder):
                    _emit(size_label, data_ws.cell(r, size_col_start + idx + 1).value)
            else:
                # 无尺码列 (NS / NO SIZE) → 尺码**留空**
                # 2026-10-22 用户纠正: 不伪造 'U', 源文件没有这个值
                _emit("", None)

    if total_placeholder:
        print(f"[SIRIO] 跳过 {total_placeholder} 张占位图")
    if len(data_sheets) > 1:
        print(f"[SIRIO] 处理了 {len(data_sheets)} 个数据 sheet")

    return rows


# Burberry FA27 (2026-10-23) — 列名与 SIRIO 完全不同, 独立解析
def _parse_burberry_bh(xlsx_path: Path) -> List[Dict[str, object]]:
    """Burberry BMB B&H 期货单 (2026-10-23): BMB B&H FA27.xlsx。

    源文件结构 (2 sheets):
      - **'B&H'** (326 SKU, 326 图, 41 列): 数据 sheet
          R1 header:
            c1=sp_id / c2=POS / c3=MPG(大类) / c4=SPG / c5=Picture /
            c6=Style No. / c7=Color('001-Black' = 色号-颜色名) /
            c8=Form Name('Brief 3P Power (259618)') / c9=WHS / c10=RRP /
            c11=Total Units / c12=Total WHS / c13=Core range / c14=Keylook /
            c15=Fit / c16=Material Comp. / c17=Quality / c18=Responsible /
            **c19=Sizes (尺码清单文本, 如 'XS, S, M, L, XL, XXL')** /
            c20=Total RRP / c21=SKU('50475273-487') / c22=Name /
            **c23-c41=尺码列 (XS/S/M/L/XL/XXL/XXXL/4XL/5XL/6XL +
                   39-42/39-40/40-46/41-42/43-46/43-44/45-46/47-50/47-48)**
      - **'index#B&H'**: 同结构的字段字典 sheet (c23+ 带 `size#` 前缀), **不使用**

    关键特点 (2026-10-23 用户确认):
      - **c23-c41 尺码数量列全空** (326 行 6194 个单元格全为空字符串)
      - 尺码清单只在 c19 的 **Sizes 文本** 里 → 按文本逗号拆行
      - **数量留空** (期货单, 等客户填)
      - 尺码文本形如 '39-42, 40-46, 43-46' (**腰码区间**), 原样保留不拆

    与 SIRIO DB 的区别:
      列名不同 (Style No./Color/Form Name vs SKU/Col Desc/Style Name),
      Picture 在 c5 而非 c1, 尺码来自 Sizes 文本而非表头尺码列。
      所以走独立函数, 不复用 _parse_sirio_whls。
    """
    rows: List[Dict[str, object]] = []

    try:
        wb = _load_wb(xlsx_path)
    except Exception:
        return rows

    ws = None
    for sn in wb.sheetnames:
        cand = wb[sn]
        if cand.max_row < 2 or cand.max_column < 19:
            continue
        hdr = [normalize_text(cand.cell(1, c).value).lower()
               for c in range(1, cand.max_column + 1)]
        # 数据 sheet 特征: C6='Style No.' + C19='Sizes' + C21='SKU'
        if "style no." in hdr and "sizes" in hdr and "sku" in hdr:
            ws = cand
            break
    if ws is None:
        return rows

    def _col(*names) -> int:
        for c in range(1, ws.max_column + 1):
            nm = normalize_text(ws.cell(1, c).value).lower()
            if nm in names:
                return c
        return 0

    c_style = _col("style no.")
    c_color = _col("color")
    c_form = _col("form name")
    c_name = _col("name")
    c_whs = _col("whs")
    c_rrp = _col("rrp")
    c_mpg = _col("mpg")
    c_material = _col("material comp.", "material comp")
    c_sizes = _col("sizes")
    c_sku = _col("sku")
    c_picture = _col("picture")

    if not c_style:
        return rows

    # 图片: Picture 列 (c5), anchor 在数据行
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    n_placeholder = 0
    try:
        for img in ws._images:
            try:
                row_0idx = img.anchor._from.row
            except Exception:
                continue
            try:
                img_bytes = img._data()
            except Exception:
                continue
            if _is_placeholder_image(img_bytes):
                n_placeholder += 1
                continue
            fmt = (img.format or "jpeg").lower()
            if fmt == "jpg":
                fmt = "jpeg"
            item = {"bytes": img_bytes, "format": fmt}
            emu = _fit_image_emu(img.width, img.height)
            if emu:
                item["display_emu"] = emu
            images_by_row.setdefault(row_0idx, []).append(item)
    except Exception:
        pass
    if n_placeholder:
        print(f"[BB] 跳过 {n_placeholder} 张占位图")

    for r in range(2, ws.max_row + 1):
        style_no = normalize_text(ws.cell(r, c_style).value)
        if not style_no:
            continue

        sku = normalize_text(ws.cell(r, c_sku).value) if c_sku else ""
        if not sku:
            sku = style_no

        # Color: '001-Black' → 色号 '001' + 颜色 'Black'
        color_raw = normalize_text(ws.cell(r, c_color).value) if c_color else ""
        color = ""
        color_code = ""
        if color_raw:
            parts = color_raw.split("-", 1)
            if len(parts) == 2 and parts[0].strip().isdigit():
                color_code = parts[0].strip()
                color = parts[1].strip()
            else:
                color = color_raw

        form_name = normalize_text(ws.cell(r, c_form).value) if c_form else ""
        name = normalize_text(ws.cell(r, c_name).value) if c_name else ""
        category = normalize_text(ws.cell(r, c_mpg).value) if c_mpg else ""
        material = normalize_text(ws.cell(r, c_material).value) if c_material else ""
        cost = to_float(ws.cell(r, c_whs).value) if c_whs else None
        retail = to_float(ws.cell(r, c_rrp).value) if c_rrp else None

        # 描述: Name 优先, 退到 Form Name
        description = name or form_name

        row_images = images_by_row.get(r - 1, [])

        # 尺码: c19 Sizes 文本逗号分隔 → 拆行
        #   'XS, S, M, L, XL, XXL' → 6 个码
        #   '39-42, 40-46, 43-46'    → 3 个腰码区间 (**原样保留, 不拆**)
        sizes_text = normalize_text(ws.cell(r, c_sizes).value) if c_sizes else ""
        size_list = [s.strip() for s in sizes_text.split(",") if s.strip()] if sizes_text else []

        if not size_list:
            rows.append({
                "brand": "BURBERRY",
                "sku": sku,
                "image": "",
                "description": description,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code,
                "gender": "",            # 源文件无性别列
                "season": "FA27",        # 文件名 BMB B&H FA27
                "origin": "",            # 源文件无产地列
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": "",              # 2026-10-22: 无尺码则留空, 不伪造
                "quantity": None,
                "vendor_id": "V1003",
                "image_items": row_images,
                "source_row_height": 92.0 if row_images else None,
                "_row_num": r,
            })
            continue

        for size_label in size_list:
            rows.append({
                "brand": "BURBERRY",
                "sku": sku,
                "image": "",
                "description": description,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code,
                "gender": "",
                "season": "FA27",
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                # 2026-10-23 用户确认: c23-c41 尺码数量列**全空** → 数量留空
                "quantity": None,
                "vendor_id": "V1003",
                "image_items": row_images,
                "source_row_height": 92.0 if row_images else None,
                "_row_num": r,
            })

    return rows


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析入口 (2026-10-26: 包一层, 保证结束一定清掉 workbook 缓存)."""
    try:
        return _parse_rows_impl(primary_file, markup_file, vendor_id)
    finally:
        _clear_wb_cache()


def _parse_rows_impl(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """Parse McQueen SS27 期货 file → output List[Dict].

    2026-10-21 扩展支持 Armani 系 PLM Report 版式 (LINESHEET EA FW27 - HOLIDAY),
    检测到 sheet 'PLM - Report' 调 _parse_plm_report。
    2026-10-21 合并 LONGCHAMP SS27 期货 (原 V1003-LC 独立适配器),
    检测 R1 C2='COMPOSIZIONE' 调 _parse_longchamp,输出 vendor_id 统一 'V1003'。
    2026-10-21 扩展支持 MAMMUT 德语 B2B 货单 (LINESHEET MAMMUT MAIN FW27),
    检测 R1 C2='SKU' + C4='Produktname' + C12='Größe' 调 _parse_mammut_linesheet。
    2026-10-21 扩展支持 MAMMUT 意大利 PRICE LIST (R7 C1='Categoria')
    与 MAMMUT LOOKBOOK 图册 PDF (.pdf)。
    2026-10-22 扩展支持 Calvin Klein NOS 期货订货表 (sheet 'Options',
    R1 C32='Option Code' + C23='Wholesale Price')。
    2026-10-22 扩展支持 Calvin Klein FASHION style 级订货表
    (sheet 名 'CKU M'/'CKU W', R3 C3 含 'STYLE #' + C11='WHSL PRICE')。
    2026-10-22 扩展支持 SIRIO DB 期货订货单 (Michael Kors 等):
    R1 C1='Picture' + C4='SKU' + C9='Style Number' + C11='Col Desc',
    有 'Info'/'Totals' 辅助 sheet。尺码列**自动扫表头**识别,
    兼容有尺码列 (鞋码 c20-c34) 和无尺码列 ('NS NO SIZE') 两种变体,
    且支持**一份文件多个数据 sheet** (2026-10-23: 每个 sheet 尺码梯度不同)。
    2026-10-23 扩展支持 Burberry BMB B&H (R1 C6='Style No.' + C19='Sizes' +
    C21='SKU')。该版尺码来自 **Sizes 文本字段**(c19), c23-c41 数量列全空,
    所以按 Sizes 逗号拆行、数量留空。
    """
    suffix = primary_file.suffix.lower()

    # 2026-10-21 detect MAMMUT LOOKBOOK 图册 PDF
    if suffix == ".pdf":
        return _parse_mammut_lookbook(primary_file)

    # 2026-10-21 detect Armani PLM Report 版式
    if suffix == ".xlsx":
        try:
            wb = openpyxl.load_workbook(primary_file, read_only=True)
            is_plm_report = "PLM - Report" in wb.sheetnames
            wb.close()
            if is_plm_report:
                return _parse_plm_report(primary_file, vendor_id)
        except Exception:
            pass

        # 2026-10-23 detect Burberry BMB B&H 期货单
        # 特征: 某 sheet R1 有 C6='Style No.' + C19='Sizes' + C21='SKU'
        try:
            wb = _load_wb(primary_file)
            for sn in wb.sheetnames:
                ws = wb[sn]
                if ws.max_row < 2 or ws.max_column < 19:
                    continue
                hdr = [normalize_text(ws.cell(1, c).value).lower()
                       for c in range(1, ws.max_column + 1)]
                if "style no." in hdr and "sizes" in hdr and "sku" in hdr:
                    return _parse_burberry_bh(primary_file)
            # 缓存的 wb 后面还要给 parse 函数用, 这里不能关
        except Exception:
            pass

        # 2026-10-22 detect SIRIO DB 期货订货单 (Michael Kors 等)
        # 特征: 某 sheet 的 R1 C1='Picture' + C4='SKU' + C9='Style Number',
        #       且有 'Info' / 'Totals' 辅助 sheet
        try:
            wb = _load_wb(primary_file)
            has_aux = ("Info" in wb.sheetnames) or ("Totals" in wb.sheetnames)
            if has_aux:
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    if ws.max_row < 2 or ws.max_column < 12:
                        continue
                    hdr = [normalize_text(ws.cell(1, c).value).lower()
                           for c in range(1, ws.max_column + 1)]
                    if (hdr[0].startswith("picture") and "sku" in hdr
                            and "style number" in hdr and "col desc" in hdr):
                        return _parse_sirio_whls(primary_file)
            # 缓存的 wb 后面还要给 parse 函数用, 这里不能关
        except Exception:
            pass

        # 2026-10-22 detect Calvin Klein NOS 期货订货表 (sheet 'Options')
        try:
            wb = _load_wb(primary_file)
            if "Options" in wb.sheetnames:
                ws = wb["Options"]
                if (ws.max_row >= 3 and ws.max_column >= 32
                        and str(ws.cell(1, 32).value or "").strip() == "Option Code"
                        and str(ws.cell(1, 23).value or "").strip() == "Wholesale Price"):
                    return _parse_calvin_klein(primary_file)

            # 2026-10-22 detect Calvin Klein FASHION style 级订货表
            # (sheet 名 'CKU M'/'CKU W', R3 c3 含 'STYLE #' + c11='WHSL PRICE')
            for sn in wb.sheetnames:
                if not sn.upper().startswith("CKU"):
                    continue
                ws = wb[sn]
                if (ws.max_row >= 4 and ws.max_column >= 11
                        and "STYLE #" in str(ws.cell(3, 3).value or "").upper()
                        and str(ws.cell(3, 11).value or "").strip().upper() == "WHSL PRICE"):
                    return _parse_calvin_klein_style(primary_file)
            # 缓存的 wb 后面还要给 parse 函数用, 这里不能关
        except Exception:
            pass

        # 2026-10-21 detect MAMMUT 意大利 PRICE LIST (R7 C1='Categoria')
        try:
            wb = _load_wb(primary_file)
            is_mammut_price = False
            for sn in wb.sheetnames:
                ws = wb[sn]
                if ws.max_row >= 8 and ws.max_column >= 6 and \
                        str(ws.cell(7, 1).value or "").strip() == "Categoria":
                    is_mammut_price = True
                    break
            if is_mammut_price:
                return _parse_mammut_price_list(primary_file)
            # 缓存的 wb 后面还要给 parse 函数用, 这里不能关
        except Exception:
            pass

        # 2026-10-21 detect MAMMUT 德语 B2B 货单 (多 sheet, R1 C2='SKU' + C4='Produktname')
        try:
            wb = _load_wb(primary_file)
            is_mammut = False
            for sn in wb.sheetnames:
                ws = wb[sn]
                if (ws.max_row >= 2 and ws.max_column >= 12
                        and str(ws.cell(1, 2).value or "").strip() == "SKU"
                        and str(ws.cell(1, 4).value or "").strip() == "Produktname"
                        and str(ws.cell(1, 12).value or "").strip() == "Größe"):
                    is_mammut = True
                    break
            if is_mammut:
                return _parse_mammut_linesheet(primary_file)
            # 缓存的 wb 后面还要给 parse 函数用, 这里不能关
        except Exception:
            pass

        # 2026-10-21 detect LONGCHAMP SS27 期货 (C2='COMPOSIZIONE')
        try:
            wb = _load_wb(primary_file)
            if _is_longchamp_format(wb.active):
                return _parse_longchamp(primary_file)
            # 缓存的 wb 后面还要给 parse 函数用, 这里不能关
        except Exception:
            pass

    if suffix == ".xls":
        # .xls 文件本身无图,从同文件夹的分类 PDF 抽图匹配 SKU
        sku_image_map = _build_sku_image_map(primary_file.parent)
        return _parse_mcqueen_xls(primary_file, sku_image_map)
    elif suffix == ".xlsx":
        return _parse_mcqueen_xlsx(primary_file)
    else:
        # .pdf or other — return empty
        return []


def get_vendor_meta() -> Dict[str, object]:
    return {
        "default_image_col_width": 40.0,
    }