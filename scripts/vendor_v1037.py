"""V1037 多品牌 vendor adapter (2026-09-23).

覆盖 7 个意大利供应商品牌(CANADA GOOSE / GUCCI / LONGCHAMP / STONE ISLAND /
THOM BROWNE / TORY BURCH / VIVIENNE WESTWOOD),共享 NUGNES-style linesheet 结构。

源文件结构(每个文件):
  Sheet "Foglio1", 36-267 行 × 40-41 列
  R2 起 = 尺码标签行(XXXS/XXS/XS/S/M/L/XL/XXL/XXXL/XXXXL/3XL/4XL/5XL 等,
                    也可能 UNI/UK1/CIN/SCD/AME/JD/C14/DON/UO/GEN 等)
  R28 = 数据列头 (MODELLO VARIANTE COLORE / MODELLO / VARIANTE / COLORE / GRUPPO
                     / SETTORE / DESCRIZIONE / BRAND / TESSUTO / COLORE / ...
                     / PR. ACQUISTO / PR. VENDITA / COD. FOTO)
  R29 = 空行
  R30+ = 数据行

两种列布局变体:
  A. 标准版(无 STAGIONE,大多数文件):
       R28: c2=SKU|c3=MODELLO|c4=VARIANTE|c5=COLORE|c6=GRUPPO|c7=SETTORE
                c8=DESCRIZIONE|c9=BRAND|c10=TESSUTO|c11=COLORE
       R2:  c12=size_system(AME/UNI/...)|c13-c25=尺码标签
       数据:c38=PR.ACQUISTO|c39=PR.VENDITA|c40=COD.FOTO
  B. 含 STAGIONE 版(GUCCI):
       R28: c2=SKU|c3=MODELLO|c4=VARIANTE|c5=COLORE|c6=GRUPPO|c7=SETTORE
                c8=STAGIONE|c9=DESCRIZIONE|c10=BRAND|c11=TESSUTO|c12=COLORE
       R2:  c13=size_system|c14-c26=尺码标签
       数据:c39=PR.ACQUISTO|c40=PR.VENDITA|c41=COD.FOTO

适配逻辑:
  1) auto-detect 列偏移(看 STAGIONE 在不在表头)
  2) 扫所有 R2-R15 找尺码 section 行,按 col 12/13 的 system marker 建索引
     (例 GUCCI 有 UK1 / UNI / CIN / SCD / AME / DON / JD / UO / C14 / JU 共 10 段)
  3) 每个 data row 按自己 col 12/13 的 marker 选对应尺码段,展开 (size, qty>0) 多行
  4) C+ 从 品牌对应C+.xlsx 查;查不到留空(用户偏好:宁缺毋假)
  5) 图片:每个源数据行选 1 张(对应规则),每个尺码行重复同一张图(2026-09-23 用户偏好)

参考:
  - 之前 CANADA GOOSE-only 版本(2026-07-17 写)
  - 用户偏好"不凭推断填充源文件没有的字段"(2026-09-23 user memory)
  - 用户偏好"SKU 多尺码展开时,每个尺码行都带同一张图"(2026-09-23)
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    compute_supply_price,
    normalize_text,
    to_float,
)


VENDOR_ID = "V1037"


# ---------------------------------------------------------------------------
# C+ lookup(读 6.29/品牌对应C+.xlsx,查不到不 fallback)
# ---------------------------------------------------------------------------

_CPLUS_TABLE_CACHE: Optional[Dict[str, int]] = None


def _load_cplus_table() -> Dict[str, int]:
    """加载品牌 → C+ 整数 映射;读取失败返回 {} (调用方留空)。"""
    global _CPLUS_TABLE_CACHE
    if _CPLUS_TABLE_CACHE is not None:
        return _CPLUS_TABLE_CACHE
    cplus_file = Path.home() / "Desktop/表格处理/6.29/品牌对应C+.xlsx"
    table: Dict[str, int] = {}
    if cplus_file.exists():
        wb = openpyxl.load_workbook(cplus_file, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            brand_val, cplus_val = row[0], row[1]
            if brand_val and cplus_val:
                m = re.search(r'C\+(\d+)%', str(cplus_val), re.IGNORECASE)
                if m:
                    table[str(brand_val).strip().upper()] = int(m.group(1))
    _CPLUS_TABLE_CACHE = table
    return table


def _get_cplus_for_brand(brand: str) -> Optional[int]:
    """按品牌查 C+ 整数百分比;查不到返回 None(不 fallback 到 18% 等默认值)。"""
    return _load_cplus_table().get(brand.strip().upper())


# ---------------------------------------------------------------------------
# Layout detection
# ---------------------------------------------------------------------------


def _find_header_row(ws) -> int:
    """找表头行:col 2 含 'MODELLO VARIANTE COLORE' (0-indexed 返回)。"""
    for r in range(min(40, ws.max_row)):
        v = ws.cell(r + 1, 2).value
        if v and 'MODELLO VARIANTE COLORE' in str(v).upper():
            return r
    raise ValueError("Cannot find header row (no 'MODELLO VARIANTE COLORE' found)")


def _has_stagione_col(ws, header_row_0idx: int) -> bool:
    """检测 R28 是否有 'STAGIONE' 列 → 决定整体列偏移是 +0 还是 +1。"""
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row_0idx + 1, c).value
        if v and 'STAGIONE' in str(v).upper():
            return True
    return False


def _find_size_sections(ws, has_stagione: bool) -> Dict[str, Dict[int, str]]:
    """扫 R1-R15 找所有尺码 section 行,返回 {system_marker: {col_0idx: size_label}}。

    多 section 文件(GUCCI 10 段,STONE ISLAND 3 段 等):每行 R2/R3/R4... 是某个尺码段的标签头:
      - 该行 col 12 (或 col 13 if has_stagione) 是 system marker (例 "UK1", "UNI", "CIN")
      - 该行 system marker 之后的非空 cell 就是尺码 label

    Returns: {system_marker: {col_0idx: size_label}}
    """
    size_token_pattern = re.compile(
        r'^(XXXS|XXS|XS|S$|M$|XL|XXL|XXXL|XXXXL|3XL|4XL|5XL|6XL|UNI|UNICA|TU|'
        r'\d+(\.\d+)?\+?|\d+\s*½|OS|ONE.?SIZE|UK\d+|CIN|JD|C\d+|SCD|AME|BIG)$',
        re.IGNORECASE)
    system_col_0idx = 12 if has_stagione else 11  # 0-indexed

    sections: Dict[str, Dict[int, str]] = {}
    for r in range(min(20, ws.max_row)):
        system_marker = str(ws.cell(r + 1, system_col_0idx + 1).value or '').strip()
        labels: Dict[int, str] = {}
        for c in range(system_col_0idx + 1, min(system_col_0idx + 25, ws.max_column)):
            v = ws.cell(r + 1, c + 1).value
            if v is not None and str(v).strip():
                labels[c] = str(v).strip()
        if len(labels) < 2:
            continue
        # 至少 2 个 label 长得像尺码(避免抓到表头行)
        size_like = sum(1 for v in labels.values() if size_token_pattern.match(v))
        if size_like < 2:
            continue
        key = system_marker if system_marker else 'DEFAULT'
        if key not in sections:
            sections[key] = labels
    return sections


# ---------------------------------------------------------------------------
# Image loader
# ---------------------------------------------------------------------------


def _load_images(xlsx_path: Path) -> Dict[int, List[Dict]]:
    """加载嵌入图片,返回 {row_idx_0idx: [image_items]}。

    V1037 文件图片 anchor 的 `from_row` (Excel 1-idx) = 数据行 - 1 (图像视觉覆盖下一行数据)。
    所以 from_row_excel 直接作为 0-idx 数据行键。

    每行只保留第一张图(源数据每行 3-4 张,取 1 张产品主图)。
    """
    import io
    import posixpath

    from PIL import Image as PILImage

    row_images: Dict[int, List[Dict]] = {}
    target_cy_emu = 92 * 12700  # 92pt

    try:
        with zipfile.ZipFile(xlsx_path) as z:
            drawing_paths = [n for n in z.namelist()
                             if re.match(r'xl/drawings/drawing\d+\.xml$', n)]
            if not drawing_paths:
                return row_images

            drawing_path = drawing_paths[0]
            drawing_dir = posixpath.dirname(drawing_path)
            drawing_basename = posixpath.basename(drawing_path)
            drawing_rels_path = f'{drawing_dir}/_rels/{drawing_basename}.rels'
            drawing_xml = z.read(drawing_path).decode('utf-8')

            rid_to_img: Dict[str, str] = {}
            if drawing_rels_path in z.namelist():
                rels_content = z.read(drawing_rels_path).decode('utf-8')
                for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_content):
                    rid_to_img[m.group(1)] = m.group(2)

            anchor_blocks = re.findall(
                r'<xdr:(?:two|one)CellAnchor\b[^>]*>(.*?)</xdr:(?:two|one)CellAnchor>',
                drawing_xml, re.DOTALL)

            for anchor_xml in anchor_blocks:
                from_m = re.search(
                    r'<xdr:from>\s*<xdr:col>\d+</xdr:col>\s*<xdr:colOff>\d+</xdr:colOff>\s*<xdr:row>(\d+)</xdr:row>',
                    anchor_xml)
                if not from_m:
                    continue
                from_row_excel = int(from_m.group(1))
                from_row_0idx = from_row_excel  # V1037:anchor row = data row N (image 视觉覆盖 N+1)
                if from_row_0idx in row_images:
                    continue  # 每行只保留第一张图

                # srcRect cropping
                l_pct = t_pct = r_pct = b_pct = 0.0
                src_rect_m = re.search(r'<a:srcRect[^/]*/>', anchor_xml)
                if src_rect_m:
                    src = src_rect_m.group()
                    l_m = re.search(r'l="(\d+)"', src)
                    t_m = re.search(r't="(\d+)"', src)
                    r_m = re.search(r'r="(\d+)"', src)
                    b_m = re.search(r'b="(\d+)"', src)
                    if l_m and t_m and r_m and b_m:
                        l_pct = int(l_m.group(1)) / 100000.0
                        t_pct = int(t_m.group(1)) / 100000.0
                        r_pct = int(r_m.group(1)) / 100000.0
                        b_pct = int(b_m.group(1)) / 100000.0

                embed_m = re.search(r'<a:blip[^>]*r:embed="([^"]+)"', anchor_xml)
                if not embed_m:
                    continue
                rid = embed_m.group(1)
                img_rel_path = rid_to_img.get(rid, '')
                if not img_rel_path:
                    continue
                if img_rel_path.startswith('../media/'):
                    img_path = f'xl/media/{img_rel_path[9:]}'
                elif img_rel_path.startswith('media/'):
                    img_path = f'xl/media/{img_rel_path[6:]}'
                else:
                    img_path = f'xl/{img_rel_path}'

                if img_path not in z.namelist():
                    continue
                img_blob = z.read(img_path)

                img_ext = img_rel_path.rsplit('.', 1)[-1].lower()
                if img_ext == 'jpg':
                    img_ext = 'jpeg'
                if img_ext not in {'jpeg', 'png', 'gif', 'bmp', 'tiff'}:
                    img_ext = 'jpeg'

                # display_emu (跟 framework 的 embed_images_in_row 一致的算法)
                ext_size_m = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"', anchor_xml)
                orig_cx_emu = int(ext_size_m.group(1)) if ext_size_m else None
                orig_cy_emu = int(ext_size_m.group(2)) if ext_size_m else None
                display_emu: Optional[tuple] = None

                if l_pct > 0 or t_pct > 0 or r_pct > 0 or b_pct > 0:
                    pil_img = PILImage.open(io.BytesIO(img_blob))
                    w_px, h_px = pil_img.size
                    x1 = int(w_px * l_pct)
                    y1 = int(h_px * t_pct)
                    x2 = int(w_px * (1.0 - r_pct))
                    y2 = int(h_px * (1.0 - b_pct))
                    cropped = pil_img.crop((x1, y1, x2, y2))
                    out_buf = io.BytesIO()
                    fmt = 'JPEG' if img_ext == 'jpeg' else img_ext.upper()
                    cropped.save(out_buf, format=fmt)
                    img_blob = out_buf.getvalue()
                    crop_cx_emu = (x2 - x1) * 9525
                    crop_cy_emu = (y2 - y1) * 9525
                    if crop_cy_emu > 0:
                        scale = target_cy_emu / crop_cy_emu
                        display_emu = (int(crop_cx_emu * scale), target_cy_emu)
                else:
                    if orig_cx_emu and orig_cy_emu and orig_cy_emu > 0:
                        scale = target_cy_emu / orig_cy_emu
                        display_emu = (int(orig_cx_emu * scale), target_cy_emu)

                item = {
                    "bytes": img_blob,
                    "format": img_ext,
                    "display_emu": display_emu,
                }

                row_images.setdefault(from_row_0idx, []).append(item)

    except Exception:
        pass

    return row_images


# ---------------------------------------------------------------------------
# Row builder
# ---------------------------------------------------------------------------


def _build_row(sku: str, brand: str, gender: str, description: str,
               category: str, color_desc: str, color_code: str, season: str,
               cost, retail, qty: int, size_label: str, row_num: int) -> Dict:
    """构建输出 row dict。

    C+ 字段:
      - 品牌在 品牌对应C+.xlsx 表里 → 用表里的整数(如 GUCCI=23)
      - 不在表里 → 留空字符串(""),按 framework 的 C+ 整数约定(无 C+ 时留空)
    supply_price = cost × (1 + C+/100);C+ 为空时 supply_price 留空(2026-07-15 用户确认)
    """
    cplus_int = _get_cplus_for_brand(brand)
    cplus_text = cplus_int if cplus_int is not None else ""
    cplus_pct = (cplus_int / 100.0) if cplus_int is not None else None
    supply_price = compute_supply_price(cost, cplus_pct)

    return {
        "brand": brand,
        "sku": sku,
        "image": "",
        "description": normalize_text(description),
        "category": normalize_text(category),
        "material": "",
        "color": normalize_text(color_desc),
        "color_code": str(color_code).strip(),
        "season": season,
        "origin": "",
        "cost": cost,
        "cplus_text": cplus_text,
        "supply_price": supply_price,
        "retail_price": retail,
        "size": str(size_label).strip(),
        "quantity": qty,
        "vendor_id": VENDOR_ID,
        "image_items": [],
        "source_row_height": None,
        "_row_num": row_num,
    }


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """主入口:解析 V1037 文件 → 输出 List[Dict](framework 期望的格式)。"""
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # Detect layout
    header_row = _find_header_row(ws)
    has_stagione = _has_stagione_col(ws, header_row)
    size_sections = _find_size_sections(ws, has_stagione)
    if not size_sections:
        raise ValueError("No size sections found")

    # Column offsets (0-indexed)
    if has_stagione:
        COL_SKU = 1
        COL_DESC = 8
        COL_BRAND = 9
        COL_COLORE = 11
        COL_SETTORE = 6
        COL_GRUPPO = 5
        COL_STAGIONE = 7
        COL_COLOR_CODE = 4
        COL_SYSTEM = 12
        COL_COST = 38
        COL_RETAIL = 39
    else:
        COL_SKU = 1
        COL_DESC = 7
        COL_BRAND = 8
        COL_COLORE = 10
        COL_SETTORE = 6
        COL_GRUPPO = 5
        COL_STAGIONE = None
        COL_COLOR_CODE = 4
        COL_SYSTEM = 11
        COL_COST = 37
        COL_RETAIL = 38

    # Load images
    row_images = _load_images(primary_file) if primary_file.suffix.lower() == '.xlsx' else {}

    rows: List[Dict[str, object]] = []
    data_start = header_row + 2  # header + 1 empty row

    for r_0 in range(data_start, ws.max_row):
        sku = str(ws.cell(r_0 + 1, COL_SKU + 1).value or '').strip()
        if not sku:
            continue

        description = str(ws.cell(r_0 + 1, COL_DESC + 1).value or '').strip()
        brand = str(ws.cell(r_0 + 1, COL_BRAND + 1).value or '').strip()
        color = str(ws.cell(r_0 + 1, COL_COLORE + 1).value or '').strip()
        gender = str(ws.cell(r_0 + 1, COL_SETTORE + 1).value or '').strip()
        category = str(ws.cell(r_0 + 1, COL_GRUPPO + 1).value or '').strip()
        color_code = str(ws.cell(r_0 + 1, COL_COLOR_CODE + 1).value or '').strip()
        season = ""
        if COL_STAGIONE is not None:
            season = str(ws.cell(r_0 + 1, COL_STAGIONE + 1).value or '').strip()

        cost = to_float(ws.cell(r_0 + 1, COL_COST + 1).value)
        retail = to_float(ws.cell(r_0 + 1, COL_RETAIL + 1).value)

        # 该数据行的图片(只 1 张,已 _load_images 限制)
        my_image = row_images.get(r_0, [])

        # 找该数据行使用的尺码 section
        system_marker = str(ws.cell(r_0 + 1, COL_SYSTEM + 1).value or '').strip()
        size_labels = size_sections.get(system_marker)
        if size_labels is None:
            # 找不到对应 section,跳过(可能是 UNI 但没找到,源文件写错了)
            continue

        for col_idx, size_label in size_labels.items():
            qty_val = ws.cell(r_0 + 1, col_idx + 1).value
            try:
                qty = int(float(qty_val or 0))
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue
            row = _build_row(sku, brand, gender, description, category,
                             color, color_code, season, cost, retail,
                             qty, size_label, r_0)
            # 每个尺码行都重复同一张图(用户偏好 — 跟 V0000 GUCCI 腰带一致)
            if my_image:
                row["image_items"] = my_image
                row["source_row_height"] = 92.0
            rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# Vendor meta (app framework 通过 get_vendor_meta() 读 vendor 默认值)
# ---------------------------------------------------------------------------


def get_vendor_meta() -> Dict[str, object]:
    """V1037 vendor 元数据 — 给 dispatcher 用。

    V1037 是多品牌工厂代码(CANADA GOOSE / GUCCI / LONGCHAMP / STONE ISLAND /
    THOM BROWNE / TORY BURCH / VIVIENNE WESTWOOD),源 c9/c10 有 BRAND 字段从源读。
    源没 brand 时让 main 走 4 级文件名兜底,不设 default_brand。
    default_image_col_width=40:KEEP_IMAGES_PER_ROW=1,92pt 高约 17 字符,加缓冲 40 字符。
    """
    return {
        "default_image_col_width": 40.0,
    }