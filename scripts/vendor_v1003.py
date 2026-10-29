"""V1003 McQueen SS27 vendor adapter (2026-09-29).

适用文件:
  - McQueen SS27 .xls / .xlsx (MAIN MEN/WOMEN ALL CAT 等)
  - McQueen SS27 .pdf (line sheet,需先用 scripts/preprocess_mcqueen_pdf.py 转 .xlsx)

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
from typing import Dict, List, Optional, Tuple

import openpyxl  # for .xlsx
import pandas as pd  # for .xls (xlrd backend)

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


VENDOR_ID = "V1003"

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
    "WOMAN": "Woman",
    "DONNA": "Woman",
    "MEN": "Man",
    "MAN": "Man",
    "UOMO": "Man",
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
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
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


def _norm_header(v: object) -> str:
    """表头名归一化: 小写 + 去空格/连字符/下划线."""
    if v is None:
        return ""
    return re.sub(r"[\s\-_]+", "", str(v).strip().lower())


# Beestore CATALOGO 表头别名 → 输出字段
# (意大利语表头 / 英语表头 两套, 同一字段在两种导出里名字不同)
_BEESTORE_COL_ALIASES: Dict[str, Tuple[str, ...]] = {
    "brand":       ("linea", "brand"),
    "cost":        ("costo", "cost"),
    "retail":      ("retail",),
    "gender":      ("sesso", "gender"),
    "department":  ("reparto", "department"),
    "description": ("descrizione", "description"),
    "sku":         ("modello", "model"),
    "season":      ("campionario", "samples"),
    "color":       ("variante",),
    "origin":      ("marca", "madein"),
    "material":    ("materiale", "material"),
    "category":    ("categoria", "category"),
    "total":       ("totale",),
}


def _find_beestore_header(ws) -> Tuple[Optional[int], Dict[str, int]]:
    """找 Beestore CATALOGO 表头行, 并建立 字段→列号 映射.

    2026-10-26 改造原因: 原来是**写死列号** (c9=costo / c11=linea / c14=modello ...),
    但同一套 Beestore 导出有多个变体, 列数 / 列序 / 表头语言 / 起始行都不一样:

      Mix scarves.xlsx      意大利语表头, Scalarino=c23, data R12 起
      bv-fg.xlsx            意大利语表头, Scalarino=c23, data R6  起
      BV belts.xlsx         英语表头,   Scalarino=c23, data R5  起
      GUCCI new ready goods 英语表头,   Scalarino=c22, data R12 起
      NEW SUPER OFFER W31   无 c1 标记,  整表左移 5 列, Scalarino=c17

    写死列号会把这些文件全部读错 (品牌读到成本列 / SKU 读到描述列),
    写死起始行 (R17) 还会漏掉前 5 个 SKU。改成按表头名定位列。
    """
    alias2field: Dict[str, str] = {}
    for field, names in _BEESTORE_COL_ALIASES.items():
        for n in names:
            alias2field[n] = field

    best_row: Optional[int] = None
    best_map: Dict[str, int] = {}
    best_score = 0
    for r in range(1, min(ws.max_row, 60) + 1):
        colmap: Dict[str, int] = {}
        scalar_cols: List[int] = []
        for c in range(1, ws.max_column + 1):
            key = _norm_header(ws.cell(r, c).value)
            if not key:
                continue
            if key == "descscalarino":
                colmap.setdefault("size_group", c)
            elif key == "scalarino":
                scalar_cols.append(c)
                continue
            field = alias2field.get(key)
            if field and field not in colmap:
                colmap[field] = c
        if scalar_cols:
            # 有 DescScalarino 列就取它; 没有则第一个 Scalarino 是全名 (老版式两列同名)
            colmap.setdefault("size_group", scalar_cols[0])
            colmap["size_letter"] = scalar_cols[-1]
        score = len(colmap)
        if score > best_score:
            best_row, best_map, best_score = r, colmap, score
    if best_score < 4 or "size_letter" not in best_map:
        return None, {}
    return best_row, best_map


def _parse_beestore_catalogo(
    primary_file: Path, vendor_id: str
) -> List[Dict[str, object]]:
    """Beestore 系统 CATALOGO 双 sheet 版式 (2026-10-11 接入, 2026-10-26 改表头驱动).

    适用文件: Parametri + CATALOGO 双 sheet (BURBERRY / BOSS / GUCCI / BV / A.P.C. 等).

    结构 (2026-10-26 起按**表头名**定位列, 不再写死列号 —— 同一套导出有多个变体,
    列数/列序/表头语言/起始行都不同, 见 _find_beestore_header 的说明):
      - size matrix: 表头行之上, 每行一个尺码组
          [size_group 列] = 组名 (如 'XXXS-XXXL' / 'UNICA (Borse/Acc.)')
          [size_letter 列] = 组代号 (如 'H' / 'T')
          [size_letter+1 起] = 该组的尺码标签序列
      - data 行: 表头行之下, 每行一个 SKU
          [size_letter 列] = 该 SKU 所属尺码组代号
          [size_letter+1 起] = 各尺码库存数量 (列与尺码标签一一对齐)

    字段取值 (表头名 → 输出, 意大利语/英语两套别名):
      Linea/Brand → brand | Costo/Cost → cost | Retail → retail_price
      Sesso/Gender → gender | Reparto/Department → department
      Descrizione/Description → description | Modello/Model → sku
      Campionario/Samples → season | Variante → color
      Marca/Made-in → origin | Materiale/Material → material
      Categoria/Category → category (没有该列则回落尺码组名)
    """
    rows: List[Dict[str, object]] = []
    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True)
    except Exception:
        return rows
    if "CATALOGO" not in wb.sheetnames:
        return rows
    ws = wb["CATALOGO"]
    if ws.max_row < 5 or ws.max_column < 20:
        return rows

    # 0) 表头定位 (2026-10-26: 改成按表头名找列, 支持多套导出变体)
    hdr_row, COL = _find_beestore_header(ws)
    if not hdr_row:
        return rows
    letter_col = COL["size_letter"]          # 尺码组代号列 (Scalarino)
    group_col = COL.get("size_group", letter_col - 1)   # 尺码组名列 (DescScalarino)
    qty_start = letter_col + 1               # 尺码标签 / 库存数量 起始列

    # 1) size matrix: 表头行之上, 每行一个尺码组
    #    [group_col] = 组名, [letter_col] = 组代号, [qty_start+] = 该组尺码标签序列
    #    有 c1 marker 的版本 (c1=0) 额外用 marker 过滤, 没有 marker 的老版式靠
    #    「字母列有短代码 + 后面有标签列」判断
    size_maps: Dict[str, Dict[int, str]] = {}   # size_letter → {col: size_label}
    SIZE_CODE_MAX_LEN = 4
    has_marker = any(
        ws.cell(r, 1).value in (0, 1, 2)
        for r in range(1, min(ws.max_row, 200) + 1)
    )
    for r in range(1, hdr_row):
        if has_marker and ws.cell(r, 1).value != 0:
            continue
        code_raw = ws.cell(r, letter_col).value
        if code_raw is None:
            continue
        code = str(code_raw).strip().upper()
        if not code or len(code) > SIZE_CODE_MAX_LEN:
            continue
        if not re.match(r"^[A-Z0-9]+$", code):
            continue
        col_map: Dict[int, str] = {}
        for c in range(qty_start, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                if col_map:
                    break
                continue
            sv = str(v).strip()
            if not sv:
                if col_map:
                    break
                continue
            col_map[c] = sv
        if col_map:
            size_maps[code] = col_map

    if not size_maps:
        return rows

    # 2) 图 (从 CATALOGO sheet 抽)
    # HOKA 主图 anchor col 优先级: col=3 (主图) > col=0 (部分 R21) > col=4 (副图)
    # 收集所有 col 的图,按 col 优先级排序,framework KEEP_IMAGES_PER_ROW=1 取第一张
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    COL_PRIORITY = {3: 0, 0: 1, 4: 2}  # col 越小优先级越高
    for img in ws._images:
        a = img.anchor
        if not hasattr(a, "_from") or a._from is None:
            continue
        try:
            img_bytes = img._data() if hasattr(img, "_data") else None
        except Exception:
            img_bytes = None
        if img_bytes is None:
            continue
        images_by_row.setdefault(a._from.row, []).append({
            "col": a._from.col,
            "bytes": img_bytes,
            "format": img.format or "jpeg",
        })
    # 每行按 col 优先级排序
    for r in images_by_row:
        images_by_row[r].sort(key=lambda x: COL_PRIORITY.get(x["col"], 99))
        # 去掉 col 字段,只留 framework 要的 bytes/format
        images_by_row[r] = [
            {"bytes": x["bytes"], "format": x["format"]} for x in images_by_row[r]
        ]

    # 3) Data rows: 表头行之下
    #    有 c1 marker 的版本用 c1==2 过滤; 老版式 (无 marker) 用 SKU 列非空判断
    sku_to_default_image: Dict[Tuple[str, str], List] = {}

    def _cell(r: int, field: str):
        c = COL.get(field)
        if not c:
            return None
        return ws.cell(r, c).value

    for r in range(hdr_row + 1, ws.max_row + 1):
        if has_marker:
            if ws.cell(r, 1).value != 2:
                continue
        else:
            if not normalize_text(_cell(r, "sku")):
                continue

        brand = normalize_text(_cell(r, "brand"))
        description = normalize_text(_cell(r, "description"))
        full_sku = normalize_text(_cell(r, "sku"))
        color = normalize_text(_cell(r, "color"))
        gender = _translate_gender(normalize_text(_cell(r, "gender")))
        origin = normalize_text(_cell(r, "origin"))
        material = normalize_text(_cell(r, "material"))
        season = normalize_text(_cell(r, "season"))
        scalarino_full = normalize_text(ws.cell(r, group_col).value)
        scalarino_letter = str(ws.cell(r, letter_col).value or "").strip().upper()
        cost = to_float(_cell(r, "cost"))
        retail = to_float(_cell(r, "retail"))

        # 类目: 优先 Categoria/Category 列, 没有则回落到尺码组名
        category = normalize_text(_cell(r, "category")) or scalarino_full

        # 货号拼接
        if color:
            sku_final = f"{full_sku}-{color}" if full_sku else color
        else:
            sku_final = full_sku

        # 图
        row_images = images_by_row.get(r - 1, [])
        sku_key = (full_sku, color)
        if row_images:
            sku_to_default_image[sku_key] = row_images
        final_images = row_images or sku_to_default_image.get(sku_key, [])

        # Size
        size_col_map = size_maps.get(scalarino_letter, {})

        # 出 size 展开
        emitted = False
        for c, size_label in size_col_map.items():
            v = ws.cell(r, c).value
            if v is None or not isinstance(v, (int, float)) or v <= 0:
                continue
            qty = int(round(v))
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
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
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": final_images,
            })
            emitted = True

        # Fallback: 尺码组代号在 size matrix 里没匹配上
        # 2026-10-26 修正: 原来这里输出 size='TOTAL', quantity=1 ——
        # 'TOTAL' 根本不是尺码, 是把整行库存当成一个尺码 (V1011 已修过同样的 bug)。
        # 现在尺码留空 (源文件没给), 数量取源文件的 Totale 列, 不凭空造尺码。
        if not emitted:
            total_qty = to_float(_cell(r, "total")) or 0
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": description,
                "category": category,
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
                "size": "",
                "quantity": int(round(total_qty)),
                "vendor_id": vendor_id,
                "image_items": final_images,
            })

    # 4) 合并重复 (货号+颜色+尺码) 三元组 (2026-10-26)
    # 源文件里同一个 Modello+Variante 可能出现多行 (不同 IDArticolo / 不同批次),
    # 展开后会产出完全相同的三元组。下游系统按三元组去重只保留首行,
    # 后面那行的数量会凭空消失 (V1003 mcqueen stock 版式踩过同样的坑)。
    # 这里把数量累加到首行, 并保持源行顺序。
    merged_qty: Dict[Tuple[str, str, str], int] = {}
    first_idx: Dict[Tuple[str, str, str], int] = {}
    for i, r in enumerate(rows):
        key = (str(r.get("sku") or ""), str(r.get("color") or ""), str(r.get("size") or ""))
        if key not in first_idx:
            first_idx[key] = i
            merged_qty[key] = 0
        merged_qty[key] += int(r.get("quantity") or 0)

    out_rows: List[Dict[str, object]] = []
    for i, r in enumerate(rows):
        key = (str(r.get("sku") or ""), str(r.get("color") or ""), str(r.get("size") or ""))
        if first_idx[key] != i:
            continue
        merged = dict(r)
        merged["quantity"] = merged_qty[key]
        out_rows.append(merged)

    return out_rows


def _parse_mcqueen_stock(
    primary_file: Path, vendor_id: str
) -> List[Dict[str, object]]:
    """ALEXANDER MCQUEEN STOCK list 版式 (2026-10-13 接入).

    文件结构 (.xlsx, 1 sheet):
      R1: 空 (c14=4489 无关)
      R2: header (c1=TRADING STATUS FLAG / c2=PRODUCT STATUS 2 / c3=Season /
                c4=Season Event / c5=Last Full Price Season / c6=Department /
                c7=Function / c8=SFC / c9=IMAGES / c10=Description / c11=Colour /
                c12=Sku / c13=Size / c14=In Warehouse (Qty) /
                c15=EURO RRP / c16=WHS)
      R3+: data (每行一个 size, 16 列)

    字段映射 (1-idx):
      c3=Season / c6=Department / c8=SFC (full SKU) / c10=Description /
      c11=Colour / c12=Sku / c13=Size / c14=Qty / c15=EURO RRP (retail) /
      c16=WHS (cost)

    C+ 规则 (用户 2026-10-08 确认):
      - 鞋类 (Department 含 'SHOE') → -20 (打 8 折)
      - 其他 → -30 (打 7 折)
      - 批发价 = cost × (1 + C+/100) = cost × 0.80 / 0.70

    2026-10-13 修复: 同一 (货号+颜色+尺码) 多行 qty 合并到首行
    (源文件 size=15 可能拆成 5 行不同 bin,合并后下游系统按三元组去重不丢数据)
    """
    rows: List[Dict[str, object]] = []
    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True)
    except Exception:
        return rows
    if not wb.worksheets:
        return rows
    ws = wb.worksheets[0]
    if ws.max_row < 3 or ws.max_column < 16:
        return rows

    # 图 (0-idx row, col 8 = 1-idx c9 IMAGES)
    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    for img in ws._images:
        a = img.anchor
        if not hasattr(a, "_from") or a._from is None:
            continue
        try:
            img_bytes = img._data() if hasattr(img, "_data") else None
        except Exception:
            img_bytes = None
        if img_bytes is None:
            continue
        images_by_row.setdefault(a._from.row, []).append({
            "bytes": img_bytes,
            "format": img.format or "jpeg",
        })

    sku_to_default_image: Dict[Tuple[str, str], List] = {}

    # 第一遍: 收集原始数据,按 (sku_final, color, size) 分组累加 qty
    # groups[key] = {'qty_sum': int, 'first_r': int, 'first_data': dict}
    groups: Dict[Tuple[str, str, str], Dict] = {}

    for r in range(3, ws.max_row + 1):
        sku_v = ws.cell(r, 12).value
        if not sku_v or not str(sku_v).strip():
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

        dept = normalize_text(ws.cell(r, 6).value)
        description = normalize_text(ws.cell(r, 10).value)
        color = normalize_text(ws.cell(r, 11).value)
        sfc = normalize_text(ws.cell(r, 8).value)
        season = normalize_text(ws.cell(r, 3).value)

        # 货号拼接: sfc + color
        if color:
            sku_final = f"{sfc}-{color}" if sfc else color
        else:
            sku_final = sfc or sku

        key = (sku_final, color, size)
        if key in groups:
            groups[key]["qty_sum"] += qty
        else:
            groups[key] = {
                "qty_sum": qty,
                "first_r": r,
                "dept": dept,
                "description": description,
                "color": color,
                "sfc": sfc,
                "season": season,
            }

    # 第二遍: 按 first_r 顺序输出 (保持源行顺序)
    for key in sorted(groups.keys(), key=lambda k: groups[k]["first_r"]):
        g = groups[key]
        sku_final, color, size = key
        r = g["first_r"]
        dept = g["dept"]
        sfc = g["sfc"]

        # 价格 (首行读)
        retail_price = to_float(ws.cell(r, 15).value)  # c15 EURO RRP
        cost = to_float(ws.cell(r, 16).value)  # c16 WHS
        cplus_rate = -0.20 if "SHOE" in (dept or "").upper() else -0.30
        supply_price = round(cost * (1 + cplus_rate), 2) if cost else None
        cplus_text = "-20" if "SHOE" in (dept or "").upper() else "-30"

        # 图 (首行 r-1)
        row_images = images_by_row.get(r - 1, [])
        sku_key = (sfc, color)
        if row_images:
            sku_to_default_image[sku_key] = row_images
        final_images = row_images or sku_to_default_image.get(sku_key, [])

        gender = "Man" if (dept or "").upper().startswith("MEN") else (
            "Woman" if "WOMEN" in (dept or "").upper() else ""
        )
        # 类目: 源文件没有专门的类目字段,留空 (2026-10-13 用户要求"不要自己做逻辑去猜")
        category = ""

        rows.append({
            "brand": "ALEXANDER MCQUEEN",
            "sku": sku_final,
            "image": "",
            "description": g["description"],
            "category": category,
            "material": "",
            "color": color,
            "color_code": "",
            "gender": gender,
            "season": g["season"],
            "origin": "",
            "cost": cost,
            "cplus_text": cplus_text,
            "supply_price": supply_price,
            "retail_price": retail_price,
            "size": size,
            "quantity": g["qty_sum"],
            "vendor_id": vendor_id,
            "image_items": final_images,
        })

    return rows


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """Parse McQueen SS27 file → output List[Dict].

    2026-10-11 扩展支持 Beestore 系统 CATALOGO 双 sheet 版式 (HOKA 等),
    检测后调 _parse_beestore_catalogo,原 McQueen line sheet 逻辑保留。
    2026-10-13 扩展支持 ALEXANDER MCQUEEN STOCK list 版式 (每行一个 size),
    调 _parse_mcqueen_stock (鞋 C+ -20 / 其他 C+ -30, WHS = cost, EURO RRP = retail)。
    """
    suffix = primary_file.suffix.lower()

    # 2026-10-11 detect Beestore CATALOGO 双 sheet 版式
    if suffix == ".xlsx":
        try:
            wb = openpyxl.load_workbook(primary_file, read_only=True)
            is_beestore = "Parametri" in wb.sheetnames and "CATALOGO" in wb.sheetnames
            # 2026-10-13 detect ALEXANDER MCQUEEN STOCK list 版式
            # R2 c8='SFC' + c6='Department' + c16='WHS' + c15='EURO RRP' + c13='Size'
            is_mcqueen_stock = False
            if not is_beestore and wb.worksheets:
                first = wb.worksheets[0]
                # 检查关键 header
                h8 = first.cell(2, 8).value if first.max_row >= 2 else None
                h16 = first.cell(2, 16).value if first.max_row >= 2 else None
                h13 = first.cell(2, 13).value if first.max_row >= 2 else None
                h15 = first.cell(2, 15).value if first.max_row >= 2 else None
                is_mcqueen_stock = (
                    h8 == "SFC" and h16 == "WHS" and h13 == "Size" and h15 == "EURO RRP"
                )
            wb.close()
            if is_beestore:
                return _parse_beestore_catalogo(primary_file, vendor_id)
            if is_mcqueen_stock:
                return _parse_mcqueen_stock(primary_file, vendor_id)
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