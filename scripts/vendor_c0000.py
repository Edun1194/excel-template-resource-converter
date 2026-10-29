"""C0000 vendor adapter — GIVENCHY China 库存清单(2026-09-23 写)。

源文件结构:
  Sheet "CHINA", 1715 行 × 19 列
  R3 = 表头
  R4+ = 数据行

  列偏移(0-idx):
      0 = SKU
      2 = Material (源为空)
      3 = Category
      4 = Family
      5 = Line
      6 = Story
      7 = Sizes (源为空)
      8 = color (例 "001 - BLACK")
      9 = model description
     10 = Made in (源为空)
     11 = HS Code
     12 = EAN
     13 = Plant
     14 = Stock type
     15 = Stock qties (单一总库存,无尺码分项)
     16 = Picture (图嵌在这一列)
     17 = Retail Full Price CNY TTC (零售价,无 cost)
     18 = Tot RRP

  图片:
    drawing1.xml 用 oneCellAnchor(no xdr: namespace prefix),
    from col=16, row=N (0-idx), 数据行也是 N (0-idx),所以 row 直接对应。
    1691 张图,24 行无图。

  已知 quirk:
    - 整张表 Material / Made in / Sizes 字段全空(用户偏好:不推断,源没有就留空)
    - 源没 brand 字段,但整张表都是 GIVENCHY,所以在 get_vendor_meta 里设 default_brand
    - C+ 表没有 Givenchy 条目 → cplus_text 留空
    - 源没 cost 列 → cost 留空(只有 retail)
    - 源没 gender 列 → gender 留空
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    normalize_text,
    to_float,
)


VENDOR_ID = "C0000"


# ---------------------------------------------------------------------------
# 字段读取规则(2026-09-23 用户明确):
#   - 描述 ← model description (col 9 0-idx)
#   - 尺码 ← Sizes (col 7 0-idx) [源通常为空,留空,不从 SKU 反推]
#   - 产地 ← Made in (col 10 0-idx) [源通常为空,留空,不从 Stock type 反推]
# 其他字段按用户偏好"不推断":空就空。
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Image loader
# ---------------------------------------------------------------------------


def _load_images(xlsx_path: Path) -> Dict[int, List[Dict]]:
    """加载嵌入图片,返回 {row_0idx: [image_items]}。

    C0000 文件的 drawing XML 用 oneCellAnchor(no namespace prefix,直接
    `<from><col>16</col><row>N</row></from>`),row N 0-idx 对应数据行 N 0-idx。

    每行只保留 1 张图(framework 的 KEEP_IMAGES_PER_ROW=1 会再 truncate)。
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
                # 注意:这个文件的 rels 是 Target 在前 Id 在后 (e.g. `Target="..." Id="rIdN"`),
                # 不是常见的 Id 在前。两种顺序都得试。
                for m in re.finditer(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels_content):
                    rid_to_img[m.group(2)] = m.group(1)
                for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_content):
                    if m.group(1) not in rid_to_img:
                        rid_to_img[m.group(1)] = m.group(2)

            # C0000 文件的 drawing 没有 xdr: namespace 前缀
            anchor_blocks = re.findall(
                r'<oneCellAnchor\b[^>]*>(.*?)</oneCellAnchor>',
                drawing_xml, re.DOTALL)

            for anchor_xml in anchor_blocks:
                from_m = re.search(
                    r'<from>\s*<col>\d+</col>\s*<colOff>\d+</colOff>\s*<row>(\d+)</row>',
                    anchor_xml)
                if not from_m:
                    continue
                from_row_0idx = int(from_m.group(1))  # 已经是 0-idx

                # srcRect(本文件未见,但保留逻辑)
                l_pct = t_pct = r_pct = b_pct = 0.0
                src_rect_m = re.search(r'<a?:?srcRect[^/]*/>', anchor_xml)
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

                embed_m = re.search(r'r:embed="([^"]+)"', anchor_xml)
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
                elif img_rel_path.startswith('/xl/'):
                    img_path = img_rel_path.lstrip('/')
                elif img_rel_path.startswith('xl/'):
                    img_path = img_rel_path
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

                # display_emu
                ext_size_m = re.search(r'<ext\s+cx="(\d+)"\s+cy="(\d+)"', anchor_xml)
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

                # 每行只保留 1 张图(framework KEEP_IMAGES_PER_ROW=1 兜底)
                if from_row_0idx in row_images:
                    continue

                item = {
                    "bytes": img_blob,
                    "format": img_ext,
                    "display_emu": display_emu,
                }
                row_images.setdefault(from_row_0idx, []).append(item)

    except Exception as e:
        print(f"  [vendor_c0000] image loading warning: {e}")

    return row_images


# ---------------------------------------------------------------------------
# Row builder
# ---------------------------------------------------------------------------


def _build_row(sku: str, description: str, category: str, color_desc: str,
               origin: str, sizes: str, retail, qty: int, row_num: int,
               my_image: List[Dict]) -> Dict:
    """构建输出 row dict(2026-09-23:字段读取严格按源)。

    description: 源 col 9 model description(直接读,不拼接 Story)
    origin:      源 col 10 Made in(直接读,通常为空,留空)
    sizes:       源 col 7 Sizes(直接读,通常为空,留空)

    已知空字段(用户偏好"宁缺毋假"):
      - material: 源字段为空
      - gender: 源无字段
      - color_code: 源只有 "001 - BLACK" 这种合并字段,不强行拆出色号
      - season: 源无字段
      - cost: 源只有零售价,无 cost
      - cplus_text: Givenchy 不在 C+ 表里
      - supply_price: 没 cost 也算不出
    """
    return {
        "brand": "",                      # 由 default_brand 兜底(见 get_vendor_meta)
        "sku": sku,
        "image": "",
        "description": normalize_text(description),
        "category": normalize_text(category),
        "material": "",                  # 源 Material 全空,留空
        "gender": "",                    # 源无字段,留空
        "color": normalize_text(color_desc),
        "color_code": "",                # 不强行从 "001 - BLACK" 拆出色号
        "season": "",                    # 源无字段,留空
        "origin": normalize_text(origin) if origin else "",
        "cost": None,                    # 源无 cost,留空
        "cplus_text": "",                # Givenchy 不在 C+ 表,留空
        "supply_price": None,            # 没 cost,supply_price 也算不出
        "retail_price": retail,
        "size": sizes if sizes else "",  # 直接读 Sizes,空就空(2026-09-23 用户偏好)
        "quantity": qty,
        "vendor_id": VENDOR_ID,
        "image_items": my_image,
        "source_row_height": 92.0 if my_image else None,
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
    """主入口:解析 GIVENCHY China 库存清单 → 输出 List[Dict]。"""
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    if "CHINA" not in wb.sheetnames:
        raise ValueError("Sheet 'CHINA' not found")
    ws = wb["CHINA"]

    # Load images
    row_images: Dict[int, List[Dict]] = {}
    if primary_file.suffix.lower() == '.xlsx':
        row_images = _load_images(primary_file)

    # 列偏移(0-idx)
    COL_SKU = 0
    COL_CATEGORY = 3
    COL_SIZES_FORMULA = 7     # 源 col 8 — 公式 =RIGHT(A<r>, LEN(A<r>)-13)
    COL_COLOR = 8
    COL_DESC = 9              # 源 col 10 (model description)
    COL_MADE_IN_FORMULA = 10  # 源 col 11 — 公式 =RIGHT(O<r>, 3)
    COL_EAN = 12
    COL_PLANT = 13
    COL_STOCK_TYPE = 14       # 源 col 15
    COL_QTY = 15
    COL_RETAIL = 17
    COL_TOT_RRP = 18

    def _eval_sizes_formula(sku: str) -> str:
        """重算 Excel 公式 =RIGHT(sku, LEN(sku)-13)。

        例:
          BH00AFH1R500441 (LEN=15) → RIGHT(., 2) = "41"
          BW709X3YNL001S  (LEN=14) → RIGHT(., 1) = "S"
          BB606CB0KA001   (LEN=13) → RIGHT(., 0) = ""
        LEN < 13 时 Excel 报 #VALUE!,Python 留空。
        """
        L = len(sku)
        if L < 13:
            return ""
        n = L - 13
        return sku[-n:] if n > 0 else ""

    def _eval_made_in_formula(stock_type: str) -> str:
        """重算 Excel 公式 =RIGHT(stock_type, 3)。

        例: "1ITA" → "ITA", "PCNA" → "CNA"
        """
        return stock_type[-3:] if len(stock_type) >= 3 else stock_type

    rows: List[Dict[str, object]] = []

    # 数据从 R4 (Excel 1-idx) = 0-idx 3 开始
    data_start_0idx = 3
    for r_0 in range(data_start_0idx, ws.max_row):
        sku = str(ws.cell(r_0 + 1, COL_SKU + 1).value or '').strip()
        if not sku:
            continue

        description = str(ws.cell(r_0 + 1, COL_DESC + 1).value or '').strip()
        category = str(ws.cell(r_0 + 1, COL_CATEGORY + 1).value or '').strip()
        color = str(ws.cell(r_0 + 1, COL_COLOR + 1).value or '').strip()
        stock_type = str(ws.cell(r_0 + 1, COL_STOCK_TYPE + 1).value or '').strip()
        qty = to_float(ws.cell(r_0 + 1, COL_QTY + 1).value)
        retail = to_float(ws.cell(r_0 + 1, COL_RETAIL + 1).value)

        try:
            qty_int = int(qty) if qty is not None else 0
        except (TypeError, ValueError):
            qty_int = 0
        if qty_int <= 0:
            continue

        # Sizes 和 Made in 是公式列,Python 重算(Excel 没缓存值,openpyxl 读不到)
        sizes = _eval_sizes_formula(sku)
        made_in = _eval_made_in_formula(stock_type)

        my_image = row_images.get(r_0, [])

        row = _build_row(
            sku=sku,
            description=description,
            category=category,
            color_desc=color,
            origin=made_in,
            sizes=sizes,
            retail=retail,
            qty=qty_int,
            row_num=r_0,
            my_image=my_image,
        )
        rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# Vendor meta (app framework 通过 get_vendor_meta() 读 vendor 默认值)
# ---------------------------------------------------------------------------


def get_vendor_meta() -> Dict[str, object]:
    """C0000 vendor 元数据 — 给 dispatcher 用。

    C0000 = GIVENCHY 中国仓:源文件没 brand 字段,整张表都是 GIVENCHY,
    所以用 default_brand 兜底。
    default_image_col_width=40:KEEP_IMAGES_PER_ROW=1,92pt 高约 17 字符,加缓冲 40 字符。
    """
    return {
        "default_brand": "Givenchy",
        "default_image_col_width": 40.0,
    }