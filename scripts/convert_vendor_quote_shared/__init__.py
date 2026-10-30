"""
Shared utilities for vendor quote conversion.

Provides constants, text/number normalization helpers, image loading,
template validation, and output writing used by all vendor adapters.
"""

from __future__ import annotations

import hashlib
import math
import re
import uuid
from collections import defaultdict
from pathlib import Path
from typing import DefaultDict, Dict, List, Optional, Sequence, Tuple

import openpyxl
import pandas as pd
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parents[2]
ASSETS_DIR = ROOT / "assets"
TEMPLATES_DIR = ASSETS_DIR / "templates"

CN_TEMPLATE = TEMPLATES_DIR / "template1-cn.xlsx"
EN_TEMPLATE = TEMPLATES_DIR / "template1-en.xlsx"

TMP_DIR = Path("/tmp") / "template-resource-converter"

# ---------------------------------------------------------------------------
# Required template columns
# ---------------------------------------------------------------------------

REQ_COLS_CN = ["品牌", "货号", "图片", "cost", "零售价", "尺码", "数量", "云仓"]
REQ_COLS_EN = [
    "Brand",
    "International code",
    "Image",
    "cost",
    "Retail price",
    "Size",
    "Quantity",
    "Vendor ID",
    "Gender",
    "Color Code",
    "Wholesale price",
]

# 每个数据行最多保留多少张图(2026-07-17 用户要求:统一只保留第一张)
# 之前 V1015 等多图产品每个数据行带 3-7 张,会撑爆图片列
# 设为 1 = 只用主图;设更大值可以保留多图(需配合更宽的图片列)
KEEP_IMAGES_PER_ROW = 1

# 货号列宽(2026-07-17 用户要求:模板默认 13 太短,扩到 26)
# 模板列宽:模板自带 13,翻倍后 26 字符(约 180px,够放 25 位货号 + 颜色后缀)
SKU_COL_WIDTH = 26.0

# ---------------------------------------------------------------------------
# Size constants
# ---------------------------------------------------------------------------

SIZE_CODE_COL = 7
SIZE_START_COL = 8

COMMON_SIZE_HEADERS = {
    "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
    "4XL", "5XL", "OS", "U", "D", "O", "US", "UK", "T", "Z2",
    "X", "Y",
    "1M", "1-3M", "3M", "3-6M", "6M", "6-9M", "9M", "9-12M",
    "12M", "12-18", "18M", "18-24", "24M", "24-36", "30M", "36M",
}

# ---------------------------------------------------------------------------
# Text / number helpers
# ---------------------------------------------------------------------------


def normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def is_missing(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def first_non_missing(*values: object, ignore_values: Optional[set] = None) -> str:
    ignore_values = ignore_values or set()
    for v in values:
        if is_missing(v):
            continue
        s = normalize_text(v)
        if s in ignore_values:
            continue
        return s
    return ""


def normalize_brand(value: object) -> str:
    return re.sub(r"\s+", " ", normalize_text(value).upper())


def to_float(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    # 去货币符号和单位 (EUR / € / USD / RMB 等)
    s = re.sub(r"^[A-Za-z$€£¥]+\s*", "", s)
    s = s.replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_cplus(rule: object) -> Optional[float]:
    s = normalize_text(rule).upper()
    if not s:
        return None
    frac = to_float(s)
    if frac is not None and 0 <= frac <= 1:
        return frac


# ---------------------------------------------------------------------------
# C+ lookup — 通用方法:从 ~/Desktop/表格处理/6.29/品牌对应C+.xlsx 读 brand→C+ 映射
# ---------------------------------------------------------------------------

_CPLUS_LOOKUP: Optional[Dict[str, str]] = None


def _load_cplus_lookup(cplus_file: Optional[Path] = None) -> Dict[str, str]:
    """Load brand -> C+ text mapping from 品牌对应C+.xlsx.

    表格结构:col1=品牌名(可大小写), col2=C+ 文本(如 'C+18%' / 'c+18%' / 'C+30%')
    返回:dict 键是品牌名 upper-case,值是原文本
    """
    if cplus_file is None:
        cplus_file = Path.home() / "Desktop/表格处理/6.29/品牌对应C+.xlsx"
    cplus_file = Path(cplus_file)
    if not cplus_file.exists():
        return {}
    import openpyxl
    wb = openpyxl.load_workbook(cplus_file, data_only=True)
    ws = wb.active
    lookup: Dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        brand_val, cplus_val = row[0], row[1]
        if brand_val and cplus_val:
            lookup[str(brand_val).strip().upper()] = str(cplus_val).strip()
    return lookup


def get_cplus_pct(brand: str, cplus_file: Optional[Path] = None, fallback: float = 0.18) -> float:
    """Return C+ percentage as float for the given brand (e.g. 0.18).

    从 品牌对应C+.xlsx 读 C+ 文本,parse 出数字百分比。
    fallback: 找不到时返回的默认值(各品牌应传自己的,如 FENDI=0.18, GUCCI=0.23, MIU MIU=0.26)
    """
    global _CPLUS_LOOKUP
    if _CPLUS_LOOKUP is None:
        _CPLUS_LOOKUP = _load_cplus_lookup(cplus_file)
    key = brand.strip().upper()
    text = _CPLUS_LOOKUP.get(key, "")
    m = re.search(r'C\+(\d+)%', text, re.IGNORECASE)
    return int(m.group(1)) / 100.0 if m else fallback


def get_cplus_int(brand: str, cplus_file: Optional[Path] = None, fallback: float = 0.18) -> int:
    """Return C+ as integer (e.g. 18) — 模板 C+- 列用纯整数,正数/负数/无 C+ 留空。"""
    pct = get_cplus_pct(brand, cplus_file, fallback)
    return int(round(pct * 100))


def compute_supply_price(cost: Optional[float], cplus_pct: Optional[float]) -> Optional[float]:
    """supply_price = cost × (1 + C+/100)。无 cost 或无 C+ 留空(None)。"""
    if cost is None or cplus_pct is None:
        return None
    return round(cost * (1 + cplus_pct), 2)
    m = re.search(r"C\+\s*(\d+(?:[.,]\d+)?)", s)
    if m:
        return float(m.group(1).replace(",", ".")) / 100.0
    return None


def calc_supply_price(cost: Optional[float], cplus: Optional[float]) -> Optional[float]:
    if cost is None:
        return None
    if cplus is None:
        return cost
    return round(cost * (1 + cplus), 4)


def format_cplus_text(cplus: Optional[float]) -> str:
    """Format C+ value for display in C+- column.
    
    Rules:
    - Positive: just the number, e.g., C+20 -> 20
    - Negative: keep the minus sign, e.g., -10 -> -10
    - None/zero: empty string
    """
    if cplus is None or cplus == 0:
        return ""
    if cplus > 0:
        # Convert to percentage and round
        return str(int(round(cplus * 100)))
    else:
        # Negative: keep the minus sign
        return str(int(round(cplus * 100)))


def is_positive_qty(value: object) -> bool:
    v = to_float(value)
    return v is not None and v > 0


def guess_product_group(description: str) -> str:
    """奢侈品牌标准分类 (参考 Net-a-Porter / Farfetch / Mytheresa / SSENSE 等)

    返回: "BAGS" | "SLG" | "SHOES" | "RTW" | "BELTS" | "JEWELRY" | "ACCESSORIES" | ""
    兜底为空 — 描述不像任何已知品类时, 留空比 "RTW" 准确。
    "RTW" 只在描述明确是成衣时才返回。

    顺序: 最具体 → 最通用 (SLG 在 BAGS 前, JEWELRY 在 ACCESSORIES 前)
    """
    d = (description or "").strip().upper()
    if not d:
        return ""

    # 关键词表 (英文 + 意大利语 + 复数前缀匹配, e.g. "PORTACART" 同时匹配 "portacarte" / "portacartes")
    CATEGORIES = [
        # SLG (Small Leather Goods) - 钱包/卡包/钥匙扣/小皮具 单独成类, 不归 BAGS
        # "slg" 是 small leather goods 行业缩写 ("borsa slg" = 小皮具包, 归 SLG 不归 BAGS)
        ("SLG", [
            "WALLET", "PORTAFOGL", "CARD HOLDER", "CARDHOLDER",
            "PORTACART", "CARD CASE", "CARDCAS", "COIN PURSE",
            "KEY CASE", "KEYCAS", "KEYRING", "KEYFOB",
            "BILLFOLD",
            " SLG", " SLGS",  # 带空格的 slg 标识, 避免误伤 "SLOGAN" 等
        ]),

        # BAGS - 手袋类(不含 SLG, SLG 已单独)
        ("BAGS", [
            "BAG", "BORSA", "BORSETTA", "TOTE", "CLUTCH", "BACKPACK",
            "ZAINO", "BAULETTO", "BAULETT", "POCHETTE",
            "HOBOS", "SATCHEL", "BUCKET", "BARREL",
        ]),

        # SHOES - 所有鞋
        ("SHOES", [
            "SHOE", "SNEAKER", "BOOT", "SANDAL", "LOAFER", "SLIPPER",
            "SCARPA", "STIVAL", "MOCASSIN", "DECOLLET", "MULES",
            "PUMP", "FLAT", "HEEL", "ESPADRIL", "ESPADRILLE", "ESPADRILLAS", "BALLERINA",
        ]),

        # BELTS - 腰带单独
        ("BELTS", [
            "BELT", "CINTUR",
        ]),

        # JEWELRY - 贵重首饰(戒指/项链/耳环/手镯/胸针/吊坠)
        # 注意: BRACELET / BRACCIAL 放 JEWELRY 而非 ACCESSORIES
        # (奢侈品牌 Bracciale 通常是金属 / 编织工艺, 属贵重)
        ("JEWELRY", [
            "RING", "ANELL", "NECKLACE", "COLLAN", "EARRING", "ORECCHIN",
            "BRACELET", "BRACCIAL", "PENDANT", "CIONDOL", "BROOCH", "SPILL",
            "CHARM", "ANKLET",
        ]),

        # ACCESSORIES - 配饰(帽子/围巾/手套/眼镜/钥匙扣非贵重)
        ("ACCESSORIES", [
            "HAT", "CAP", "BERRETT", "BEANIE", "BONNET",
            "SCARF", "SCIARP", "FOULARD", "STOLE", "PONCHO",
            "GLOVE", "GUANT", "MITTEN",
            "SUNGLASS", "OCCHIAL",
            "KEYCHAIN", "PORTACHIAV", "BOOKMARK",
            "LIGHTER", "ACCENDIN", "TIE", "CRAVATT",
        ]),

        # RTW (Ready-to-Wear) - 成衣, **只**在描述明显是服装时返回
        ("RTW", [
            "JACKET", "GIUBBOTTO", "GIUBBIN", "COAT", "CAPPOTT", "TRENCH", "PARKA",
            "BOMBER", "BLAZER", "CARDIGAN", "CARDIG", "GILET", "VEST",
            "SHIRT", "CAMICIA", "BLOUSE", "BLUSA", "T-SHIRT", "TSHIRT", "T SHIRT",
            "PANTS", "PANTALON", "JEANS", "JEA", "TROUSER", "BERMUDA",
            "SKIRT", "GONNA", "DRESS", "VESTIT", "ABIT",
            "TOP", "SWEATER", "MAGLIA", "FELPA", "TUTA", "POLO", "TANK",
            "ROMPER", "JUMPSUIT", "TUNIC", "TUNICA", "CARDIGAN",
        ]),
    ]

    for cat, keywords in CATEGORIES:
        if any(k in d for k in keywords):
            return cat
    # 兜底: 空字符串 (不要 RTW)
    return ""


# ---------------------------------------------------------------------------
# Column / template helpers
# ---------------------------------------------------------------------------


def find_col(df: pd.DataFrame, candidates: Sequence[str]) -> Optional[str]:
    lookup: Dict[str, str] = {}
    for col in df.columns:
        key = normalize_text(col).upper()
        if key:
            lookup[key] = col
    for candidate in candidates:
        hit = lookup.get(candidate.upper())
        if hit is not None:
            return hit
    return None


def pick_template(language: str) -> Path:
    return EN_TEMPLATE if language.lower() == "en" else CN_TEMPLATE


# ---------------------------------------------------------------------------
# Image loading from Excel
# ---------------------------------------------------------------------------


# --------------------------------------------------------------------------------------
# Image loading from Excel - by SKU (descr-based, not anchor-based)
# --------------------------------------------------------------------------------------

def _load_drawing_images_via_xml(primary_file: Path) -> Dict[str, Dict[str, List[Dict[str, object]]]]:
    """
    Load embedded images by parsing drawing XML directly.
    
    Returns: Dict[sheet_name -> Dict[sku_code -> List[image_items]]]
    
    This is the CORRECT method for files where anchor.row does NOT match the
    actual data row (e.g., V1051 SALOMON where row=4 anchor contains the image
    for row=6's SKU because multiple anchors share the same xml_row).
    
    Instead of using anchor._from.row, we extract the SKU code from the image's
    descr attribute (e.g., "102_L47299100.png" -> "L47299100") and look up
    the actual Excel row for that SKU.
    """
    import zipfile, re, io, posixpath
    from collections import defaultdict
    
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    
    result: Dict[str, Dict[str, List[Dict[str, object]]]] = {}
    
    with zipfile.ZipFile(primary_file) as z:
        # Build sheet_index -> sheet_name mapping (decode XML entities like &amp;)
        import html
        wb_xml = z.read('xl/workbook.xml').decode('utf-8')
        sheet_rels = {}
        for m in re.finditer(r'<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml):
            sheet_rels[m.group(2)] = html.unescape(m.group(1))
        
        # Load workbook rels to get worksheet paths
        wb_rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        rid_to_path = {}
        for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', wb_rels):
            rid_to_path[m.group(1)] = m.group(2)
        
        # For each sheet, find its drawing files
        for ws in wb.worksheets:
            row_images: Dict[str, List[Dict[str, object]]] = {}
            
            # Find the worksheet's rId
            ws_rid = None
            for rid, name in sheet_rels.items():
                if name == ws.title:
                    ws_rid = rid
                    break
            
            if ws_rid not in rid_to_path:
                result[ws.title] = row_images
                continue

            ws_rel_path = f'xl/{rid_to_path[ws_rid]}'
            ws_dir = posixpath.dirname(ws_rel_path)  # e.g. xl/worksheets (dir containing the sheet file)
            ws_rels_dir = ws_dir + '/_rels/'
            ws_rel_fn = posixpath.basename(ws_rel_path) + '.rels'
            ws_rels_full = ws_rels_dir + ws_rel_fn

            # Find drawing files for this worksheet
            drawing_files = []
            if ws_rels_full in z.namelist():
                rels_content = z.read(ws_rels_full).decode('utf-8')
                drawing_files = re.findall(r'Target="([^"]*drawing\d+\.xml)"', rels_content)
                # Resolve against the sheet file's directory (not the _rels subdirectory)
                drawing_files = [
                    posixpath.normpath(posixpath.join(ws_dir, f))
                    for f in drawing_files
                ]
                drawing_files = [f for f in drawing_files if f in z.namelist()]
            
            # Build SKU -> excel_row mapping for this sheet
            sku_to_row: Dict[str, int] = {}
            # FTW: col 8 = article code (index 7), data from row 4
            # App & Acc: col 5 = article code (index 4), data from row 3
            # etc. We auto-detect by finding the first non-empty code column
            if ws.max_row > 1:
                # Try to find the article code column (scan first 15 cols for values that look like SKU codes)
                for row_idx in range(1, min(ws.max_row + 1, 200)):
                    for col_idx in range(1, min(ws.max_column + 1, 30)):
                        cell_val = ws.cell(row=row_idx, column=col_idx).value
                        if cell_val and isinstance(cell_val, str):
                            val = str(cell_val).strip()
                            # SKU codes are like L47179800, LD0263800, etc.
                            if re.match(r'^[LH][A-Z0-9]{7,}$', val):
                                sku_to_row[val] = row_idx
                                break
                    # Only scan first 100 rows for SKU headers
                    if row_idx > 100:
                        break
            
# Parse each drawing file
            for drawing_path in drawing_files:
                drawing_xml = z.read(drawing_path).decode('utf-8')
                
                # Drawing rels are in xl/drawings/_rels/drawingN.xml.rels
                drawing_dir = posixpath.dirname(drawing_path)
                drawing_basename = posixpath.basename(drawing_path)
                drawing_rels_path = f'{drawing_dir}/_rels/{drawing_basename}.rels'
                
                # Load drawing rels to get rId -> image file mapping
                rid_to_img = {}
                if drawing_rels_path in z.namelist():
                    rels_content = z.read(drawing_rels_path).decode('utf-8')
                    for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_content):
                        rid_to_img[m.group(1)] = m.group(2)
                
                # Extract images from twoCellAnchor elements
                anchors = re.findall(r'<xdr:twoCellAnchor>(.*?)</xdr:twoCellAnchor>', drawing_xml, re.DOTALL)
                
                # First pass: collect images matched via SKU in descr
                sku_matched_rows = {}  # row -> list of image dicts
                sku_matched_hashes = set()
                
                for anchor_xml in anchors:
                    # Get image descr (contains SKU code)
                    descr_m = re.search(r'<xdr:cNvPr[^>]*descr=\"([^\"]+)\"', anchor_xml)
                    if not descr_m:
                        continue
                    descr = descr_m.group(1)
                    
                    # Extract SKU from descr (format: "NNN_SKUCODE.ext")
                    sku_m = re.search(r'_([LH][A-Z0-9]+)\.(png|jpg|jpeg)', descr)
                    if not sku_m:
                        continue
                    sku = sku_m.group(1)
                    
                    # Get image data via rId
                    embed_m = re.search(r'<a:blip[^>]*r:embed=\"([^\"]+)\"', anchor_xml)
                    if not embed_m:
                        continue
                    rid = embed_m.group(1)
                    img_rel_path = rid_to_img.get(rid, "")
                    if not img_rel_path:
                        continue
                    
                    # Resolve image path (drawing rels are relative to drawing dir)
                    if img_rel_path.startswith('../media/'):
                        img_path = f'xl/media/{img_rel_path[9:]}'
                    elif img_rel_path.startswith('media/'):
                        img_path = f'xl/media/{img_rel_path[6:]}'
                    else:
                        # relative to xl/
                        img_path = f'xl/{img_rel_path}'
                    
                    if img_path not in z.namelist():
                        continue
                    
                    img_blob = z.read(img_path)
                    
                    # Determine format
                    img_ext = img_rel_path.rsplit('.', 1)[-1].lower()
                    if img_ext == 'jpg':
                        img_ext = 'jpeg'
                    if img_ext not in {'jpeg', 'png', 'gif', 'bmp', 'tiff'}:
                        img_ext = 'png'
                    
                    # Get the excel row for this SKU (via anchor._from.row for reference only,
                    # actual row is from sku_to_row lookup)
                    from_m = re.search(r'<xdr:from>\s*<xdr:col>(\d+)</xdr:col>\s*<xdr:colOff>(\d+)</xdr:colOff>\s*<xdr:row>(\d+)</xdr:row>', anchor_xml)
                    anchor_row = int(from_m.group(3)) + 1 if from_m else None
                    
                    # Find correct excel row from sku_to_row
                    excel_row = sku_to_row.get(sku, anchor_row)
                    
                    if excel_row not in sku_matched_rows:
                        sku_matched_rows[excel_row] = []
                    sku_matched_rows[excel_row].append({
                        "bytes": img_blob,
                        "format": img_ext,
                        "_descr": descr,
                        "_sku": sku,
                    })
                    sku_matched_hashes.add(hashlib.md5(img_blob).hexdigest())
                
                # Add SKU-matched images
                for row, items in sku_matched_rows.items():
                    if row not in row_images:
                        row_images[row] = []
                    row_images[row].extend(items)
                
                # Second pass: fallback row-based matching for anchors without SKU in descr.
                # Used for sheets like BAGS & PACKS where images have no descr attribute
                # but are positioned at the correct row (not decorative elements).
                #
                # Also apply per-row limits to avoid assigning too many images to the same row.
                # In BAGS & PACKS, ~197 anchors are at row 3 but should be distributed across
                # multiple rows; limiting to 2 per row ensures correct distribution.
                row_fallback_counts: Dict[int, int] = {}
                max_fallback_per_row = 2
                
                for anchor_xml in anchors:
                    # Skip if descr has SKU (already matched above)
                    descr_m = re.search(r'<xdr:cNvPr[^>]*descr=\"([^\"]+)\"', anchor_xml)
                    if descr_m:
                        descr = descr_m.group(1)
                        if re.search(r'_([LH][A-Z0-9]+)\.(png|jpg|jpeg)', descr):
                            continue  # Already matched via SKU
                    
                    # Get anchor row
                    from_m = re.search(r'<xdr:from>\s*<xdr:col>(\d+)</xdr:col>\s*<xdr:colOff>(\d+)</xdr:colOff>\s*<xdr:row>(\d+)</xdr:row>', anchor_xml)
                    if not from_m:
                        continue
                    anchor_row = int(from_m.group(3)) + 1  # Excel is 1-indexed
                    
                    # Skip decorative elements: rotated lines (rot attribute present, cx=0 or very small)
                    # These are visible artifacts in some sheets (e.g., BAGS & PACKS drawing4 row 2)
                    has_rotation = re.search(r'<a:xfrm\s+rot=\"[^\"]+\"', anchor_xml)
                    cx_m = re.search(r'<a:ext\s+cx=\"(\d+)\"', anchor_xml)
                    cx = int(cx_m.group(1)) if cx_m else None
                    if has_rotation or (cx is not None and cx < 10000):  # < 1mm in EMU = decorative line
                        continue
                    
                    # Per-row limit to avoid flooding a single row with many images
                    count = row_fallback_counts.get(anchor_row, 0)
                    if count >= max_fallback_per_row:
                        continue
                    row_fallback_counts[anchor_row] = count + 1
                    
                    # Get image data via rId
                    embed_m = re.search(r'<a:blip[^>]*r:embed=\"([^\"]+)\"', anchor_xml)
                    if not embed_m:
                        continue
                    rid = embed_m.group(1)
                    img_rel_path = rid_to_img.get(rid, "")
                    if not img_rel_path:
                        continue
                    
                    # Resolve image path
                    if img_rel_path.startswith('../media/'):
                        img_path = f'xl/media/{img_rel_path[9:]}'
                    elif img_rel_path.startswith('media/'):
                        img_path = f'xl/media/{img_rel_path[6:]}'
                    else:
                        img_path = f'xl/{img_rel_path}'
                    
                    if img_path not in z.namelist():
                        continue
                    
                    img_blob = z.read(img_path)
                    
                    # Skip if this exact image bytes was already matched via SKU (dedup)
                    if hashlib.md5(img_blob).hexdigest() in sku_matched_hashes:
                        continue
                    
                    # Determine format
                    img_ext = img_rel_path.rsplit('.', 1)[-1].lower()
                    if img_ext == 'jpg':
                        img_ext = 'jpeg'
                    if img_ext not in {'jpeg', 'png', 'gif', 'bmp', 'tiff'}:
                        img_ext = 'png'
                    
                    if anchor_row not in row_images:
                        row_images[anchor_row] = []
                    row_images[anchor_row].append({
                        "bytes": img_blob,
                        "format": img_ext,
                        "_descr": "",
                        "_sku": "",
                    })
            
            result[ws.title] = row_images
    
    return result


def load_sheet_images_by_row(primary_file: Path) -> Dict[int, List[Dict[str, object]]]:
    """简单的图加载器:按 anchor._from.row 索引,1 图 1 行 直接拿。

    适用:每行 1 张图、图锚点行号 = 数据行号的文件(V1031 FENDI / V1010 Longchamp 等)。
    与 ``load_sheet_embedded_images`` 的区别:
      - load_sheet_embedded_images 是 V1051 SALOMON 专用,从 descr 里抽 SKU 来匹配,
        适合"图锚点和数据行号错位"的版式
      - 本函数直接用 anchor._from.row,适合"图锚点 = 数据行"的简单版式

    Returns: {row_0indexed -> [{bytes, format}]}
    对 .xls (OLE2 老格式) 返回空 dict。
    """
    path_str = str(primary_file)
    if not path_str.lower().endswith(".xlsx"):
        return {}
    try:
        wb = openpyxl.load_workbook(path_str)
    except Exception:
        return {}
    if not wb.worksheets:
        return {}
    ws = wb.worksheets[0]
    out: Dict[int, List[Dict[str, object]]] = {}
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


def load_xls_ole2_jpeg_images(
    primary_file: Path,
    sheet_index: int = 0,
    image_col: Optional[int] = None,
) -> Dict[int, List[Dict[str, object]]]:
    """OLE2 老 .xls 格式图加载器:V1042 POLO 等把 JPEG 二进制塞进 Workbook 流。

    适用文件特征:
      - 后缀 .xls (OLE2 Compound File,非 .xlsx zip)
      - 只有 Workbook 一个数据流,内含连续 JPEG bytes
      - 图片在 MSODrawing F010 Client Anchor 中,col 固定 (如 col=3)
      - JPEG 顺序 = MSODrawing F010/F00B 顺序 (F00B 含 file name 但 F010 才有 row)

    算法:
      1. olefile 打开,读 Workbook 流全部 bytes
      2. 找 SOI/EOI 标记切分 JPEG blobs
      3. xlrd 打开 + monkey-patch handle_msodrawingetc 捕获 F010 (18 字节 Client Anchor)
         unpack '<Hiiii' = unk_short(2) + col_lo(4) + row_lo(4) + col_hi(4) + row_hi(4)
         实际: vals[0]=move_type, vals[1]=col_lo, vals[2]=row_lo
      4. 顺序配对: JPEG[i] ↔ F010[i] ↔ row=F010.row_lo
      5. 返回 {row_0indexed -> [{bytes, format}]}

    Args:
        primary_file: .xls 路径
        sheet_index: sheet 下标 (默认 0)
        image_col: 过滤 anchor col,只取 col=image_col 的图 (None = 不过滤)

    Returns:
        {xlrd_0indexed_row -> [{bytes, format}]}
        没找到 JPEG / 不是 OLE2 / 解析失败 → 返回空 dict
    """
    path_str = str(primary_file)
    if not path_str.lower().endswith(".xls"):
        return {}

    # 1) OLE2 读 Workbook 流
    try:
        import olefile  # type: ignore
        ole = olefile.OleFileIO(path_str)
        if not ole.exists("Workbook"):
            ole.close()
            return {}
        wb_bytes = ole.openstream("Workbook").read()
        ole.close()
    except Exception:
        return {}

    # 2) 切 JPEG blobs
    soi_marker = b"\xff\xd8\xff"
    # 2026-09-29 fix: 别用 EOI `\xff\xd9` 切,JPEG 数据流里偶尔有 spurious EOI bytes
    # 当数据 (不是 marker),基础版 `find(eoi, s)` 会切到第一个就停,导致
    # 切短 → PIL.UnidentifiedImageError
    # 改用 SOI 区间 [SOI[i], SOI[i+1]) 切,让 PIL 自动 trim 到真 EOI
    jpegs: List[bytes] = []
    soi_positions: List[int] = []
    pos = 0
    while True:
        s = wb_bytes.find(soi_marker, pos)
        if s == -1:
            break
        soi_positions.append(s)
        pos = s + 1
    for i, s in enumerate(soi_positions):
        e = soi_positions[i + 1] if i + 1 < len(soi_positions) else len(wb_bytes)
        jpegs.append(wb_bytes[s:e])
    if not jpegs:
        return {}

    # 3) 解 F010 Client Anchors — 优先 raw 字节扫描(2026-09-12 fix: xlrd MSODrawing
    #    解析在文件 >1MB 时只解前 ~28 个 record,后面产品的 anchor 字节全在但被 xlrd 跳过;
    #    raw 找 F010 signature 一次性解全部 124 个 anchor)
    #
    # F010 ClientAnchor 实际 layout (BIFF8,实测 POLO 现货单):
    #   6 字节 header = sig(2) "10 F0" + cb(4) "12 00 00 00" (=18 payload)
    #   18 字节 data  = move(2) + col_lo(4) + row_lo(4) + col_hi(4) + row_hi(4)
    anchors: List[Tuple[int, int]] = []  # [(col_lo, row_lo)]

    # 3a) raw scan: F010 signature + cb=18 → 解 (col_lo, row_lo) from data[0:10]
    import struct
    sig = b"\x10\xf0\x12\x00\x00\x00"  # F010 + 18 bytes payload
    cursor = 0
    while True:
        idx = wb_bytes.find(sig, cursor)
        if idx == -1:
            break
        cursor = idx + 1
        if idx + 6 + 10 > len(wb_bytes):
            continue
        try:
            move, col_lo, row_lo = struct.unpack(
                "<Hii", wb_bytes[idx + 6:idx + 6 + 10]
            )
        except Exception:
            continue
        # sanity: col 0-15, row 0-3000 是 BIFF sheet 合理范围
        if 0 <= col_lo <= 15 and 0 <= row_lo <= 3000:
            anchors.append((col_lo, row_lo))

    # 3b) fallback: 如果 raw scan 找不到 anchor(早期文件),用 xlrd monkey-patch
    if not anchors:
        try:
            import xlrd  # type: ignore
            import xlrd.sheet  # type: ignore
            import xlrd.biffh  # type: ignore
        except Exception:
            pass
        else:
            xlrd_anchors: List[Tuple[int, int]] = []
            original = xlrd.sheet.Sheet.handle_msodrawingetc

            def _patched(self, recid, data_len, data):  # noqa: ANN001
                cur = 0
                while cur + 8 <= data_len:
                    try:
                        tmp, fbt, cb = xlrd.biffh.unpack(
                            "<HHI", data[cur:cur + 8]
                        )
                    except Exception:
                        break
                    ver = tmp & 0xF
                    ndb = cb if ver != 0xF else 0
                    if fbt == 0xF010 and ndb == 18:
                        try:
                            vals = xlrd.biffh.unpack(
                                "<Hiiii", data[cur + 8:cur + 8 + ndb]
                            )
                            # vals[0]=move_type(2), vals[1]=col_lo(4), vals[2]=row_lo(4)
                            xlrd_anchors.append((vals[1], vals[2]))
                        except Exception:
                            pass
                    cur += ndb + 8
                return original(self, recid, data_len, data)

            try:
                xlrd.sheet.Sheet.handle_msodrawingetc = _patched
                xlrd.open_workbook(path_str, formatting_info=True)
            except Exception:
                pass
            finally:
                try:
                    xlrd.sheet.Sheet.handle_msodrawingetc = original
                except Exception:
                    pass
            anchors = xlrd_anchors

    if image_col is not None:
        anchors = [(c, r) for c, r in anchors if c == image_col]

    # 4) 配对: JPEG[i] ↔ F010[i] (用 jpeg 索引,避免 anchor 越界)
    # 2026-09-12 加 swatch 裁剪 — POLO 等供应商源图是"产品 + 下方色卡"组合,
    # 自动检测底部 swatch 区域裁掉,只留产品主体,补白回 1:1
    out_map: Dict[int, List[Dict[str, object]]] = {}
    for jpeg_idx, jpeg_bytes in enumerate(jpegs):
        if jpeg_idx >= len(anchors):
            break
        col_lo, row_lo = anchors[jpeg_idx]
        cleaned_bytes = _strip_bottom_swatch(jpeg_bytes)
        # 2026-09-29 fix: 源 .xls 偶尔有 SOF0 声称 8224x4360 但数据只有 7KB 的无效 JPEG
        # (PIL.Image.open 会 UnidentifiedImageError,XLImage 也跟着挂)
        # 在写入 out_map 前先 PIL.verify() 一遍,无效的直接跳过
        try:
            from PIL import Image as _PILImage
            import io as _pioi
            _pi = _PILImage.open(_pioi.BytesIO(cleaned_bytes))
            _pi.verify()
        except Exception:
            continue
        out_map.setdefault(row_lo, []).append({
            "bytes": cleaned_bytes,
            "format": "jpeg",
        })
    return out_map


def _strip_bottom_swatch(jpeg_bytes: bytes) -> bytes:
    """POLO 源图常见"产品 + 下方色卡 (swatch)"组合图,裁掉 swatch 留产品主体,补白回 1:1。

    检测算法:
      1) 量化 RGB 到 32 阶(降低 JPEG 噪声)
      2) 每行算 unique color count
      3) 从下往上找连续 ≥8 行 unique ≤ 8 的"纯色条带"(swatch 特征)
      4) 如果该条带主色不是白色 (RGB min > 240 且 std < 30),就裁掉
      5) 裁后图按原宽补白成 1:1(扩到底部白色),避免 Excel cell 显示变形

    失败 / 已是 1:1 / 裁后比例 < 0.4 → 返回原 bytes
    """
    try:
        from PIL import Image
        import io as _io
        import numpy as _np
        pi = Image.open(_io.BytesIO(jpeg_bytes)).convert("RGB")
    except Exception:
        return jpeg_bytes

    w, h = pi.size
    if w < 80 or h < 80:
        return jpeg_bytes

    arr = _np.array(pi).astype(_np.int32)
    quant = (arr // 32).astype(_np.int32).reshape(h, -1, 3)
    row_uniques = _np.array([len(_np.unique(quant[r], axis=0)) for r in range(h)])

    consec = 0
    for r in range(h - 1, h // 3, -1):
        if row_uniques[r] <= 8:
            consec += 1
        else:
            break

    if consec < 8:
        return jpeg_bytes
    swatch_start = h - consec

    # 白色背景过滤:swatch 区域主色接近白色 → 是纯产品照的空白底,不裁
    swatch_region = arr[swatch_start:h]
    avg_color = swatch_region.mean(axis=(0, 1))
    is_white_bg = avg_color.min() > 240 and swatch_region.std() < 30
    if is_white_bg:
        return jpeg_bytes

    # 裁 + 补白 1:1
    cropped = pi.crop((0, 0, w, swatch_start))
    if swatch_start < h * 0.4:
        # 裁过头了(< 40% 留产品),放弃裁剪
        return jpeg_bytes
    padded = Image.new("RGB", (w, h), (255, 255, 255))
    padded.paste(cropped, (0, 0))
    buf = _io.BytesIO()
    padded.save(buf, format="JPEG", quality=92)
    return buf.getvalue()


def load_richvalue_images(primary_file: Path) -> Dict[int, List[Dict[str, object]]]:
    """Excel 365 richValue 机制:从 cell.vm → rdrichvalue → richValueRel → image file。

    适用:从 Microsoft 365 / SharePoint / OneDrive 同步下来的 xlsx。
    这种 xlsx 有两套图机制:
      - 老 drawing1.xml (大 jpg 1920x1800) - openpyxl 读得到,但 Excel 不显示
      - 新 richValue linked value (小 png 84x112) - Excel 实际显示,openpyxl 读不到
    本函数走 richValue 路径拿 *Excel 实际显示的图*。

    流程:
    1. sheet1.xml: E 列 cell 的 vm=N  (1-idx 索引到 rdrichvalue)
    2. rdrichvalue.xml: rvs[vm-1]       (vm 是 1-idx,rvs 是 0-idx)
       = (_rvRel:LocalImageIdentifier=M, CalcOrigin=...)
    3. richValueRel.xml: rels_list[M]   (0-idx 索引,直接是 rId 编号-1)
       = "rIdN"
    4. richValueRel.xml.rels: rIdN → "../media/imageK.png"
    5. 读 xl/media/imageK.png

    Returns: {row_1idx -> [{bytes, format}]} (1-idx row)
    对非 richValue 文件(无 xl/richData/ 目录)返回空 dict。
    """
    out: Dict[int, List[Dict[str, object]]] = {}
    if not str(primary_file).lower().endswith(".xlsx"):
        return out
    try:
        z = __import__("zipfile").ZipFile(primary_file, "r")
    except Exception:
        return out
    # 1) 检测是否 richValue 文件
    if "xl/richData/richValueRel.xml" not in z.namelist():
        return out
    try:
        sheet1 = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        rd = z.read("xl/richData/rdrichvalue.xml").decode("utf-8")
        rels = z.read("xl/richData/richValueRel.xml").decode("utf-8")
        rels_targets = z.read("xl/richData/_rels/richValueRel.xml.rels").decode("utf-8")
    except KeyError:
        return out

    import re as _re
    # E 列 cell → vm
    e_vm: Dict[int, int] = {}
    for m in _re.finditer(r'<c r="E(\d+)"[^>]*?vm="(\d+)"', sheet1):
        e_vm[int(m.group(1))] = int(m.group(2))

    rvs = _re.findall(r'<rv s="\d+"><v>(\d+)</v><v>(\d+)</v></rv>', rd)
    rels_list = _re.findall(r'<rel r:id="(rId\d+)"\s*/>', rels)
    rid_to_target: Dict[str, str] = {}
    for m in _re.finditer(
        r'<Relationship Id="(rId\d+)"[^>]*Target="\.\./(media/[^"]+)"', rels_targets
    ):
        rid_to_target[m.group(1)] = m.group(2)

    for row_1idx, vm in e_vm.items():
        # ★ vm 是 1-idx, rvs 是 0-idx, 差 1
        rv_idx = vm - 1
        if not (0 <= rv_idx < len(rvs)):
            continue
        lid = int(rvs[rv_idx][0])
        if not (0 <= lid < len(rels_list)):
            continue
        rid = rels_list[lid]
        target = rid_to_target.get(rid)
        if not target:
            continue
        try:
            data = z.read(f"xl/{target}")
        except KeyError:
            continue
        fmt = "png" if target.lower().endswith(".png") else "jpeg"
        out[row_1idx] = [{"bytes": data, "format": fmt}]
    return out


def has_richvalue_images(primary_file: Path) -> bool:
    """检测 xlsx 是否用 richValue 机制(快速判断,用于 vendor 适配器分支)。"""
    if not str(primary_file).lower().endswith(".xlsx"):
        return False
    try:
        z = __import__("zipfile").ZipFile(primary_file, "r")
        return "xl/richData/richValueRel.xml" in z.namelist()
    except Exception:
        return False


def load_richvalue_images_by_col(
    primary_file: Path, col_letter: str
) -> Dict[int, List[Dict[str, object]]]:
    """Excel 365 richValue 图片 —— **按指定列**取 cell.vm, 且兼容两种 rv 结构。

    与 ``load_richvalue_images`` 的区别(那个函数写死 E 列 + 只认"恰好 2 个 <v>"的 rv):
      1. **列不写死** —— 本函数按 ``col_letter`` 取(POLO RL ATS 的图片在 B 列)
      2. **按结构定义解析字段** —— 读 ``xl/richData/rdrichvaluestructure.xml``,
         认识两种结构:
           ``_webimage``   字段: WebImageIdentifier / CalcOrigin / ComputedImage
                                / ImageSizing / Text
           ``_localImage`` 字段: _rvRel:LocalImageIdentifier / CalcOrigin [/ Text]

    链路:
      cell vm=N → ``rdrichvalue.xml`` 第 N 个 ``<rv>``(vm 是 1-idx, rv 是 0-idx)
        - _webimage   → WebImageIdentifier → ``rdRichValueWebImage.xml`` 的
                        ``webImageSrd`` 序号 → ``<blip r:id="rIdX"/>``
                        → ``rdRichValueWebImage.xml.rels`` → ``xl/media/imageN.jpeg``
        - _localImage → _rvRel:LocalImageIdentifier → ``richValueRel.xml`` 的
                        ``<rel>`` 序号 → rIdX → ``richValueRel.xml.rels``
                        → ``xl/media/imageN.png``

    Returns: {row_1idx -> [{bytes, format}]}
    """
    out: Dict[int, List[Dict[str, object]]] = {}
    if not str(primary_file).lower().endswith(".xlsx"):
        return out
    try:
        z = __import__("zipfile").ZipFile(primary_file, "r")
    except Exception:
        return out
    names = z.namelist()
    if "xl/richData/rdrichvalue.xml" not in names:
        return out

    import re as _re

    try:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8", "ignore")
        rd_xml = z.read("xl/richData/rdrichvalue.xml").decode("utf-8", "ignore")
    except KeyError:
        return out

    # 1) 指定列的 cell → vm (1-idx row)
    vm_by_row: Dict[int, int] = {}
    pat = r'<c r="%s(\d+)"[^>]*?vm="(\d+)"' % _re.escape(col_letter.upper())
    for m in _re.finditer(pat, sheet_xml):
        vm_by_row[int(m.group(1))] = int(m.group(2))
    if not vm_by_row:
        return out

    # 2) 结构定义: s 序号 → 字段名列表
    struct_keys: List[List[str]] = []
    if "xl/richData/rdrichvaluestructure.xml" in names:
        try:
            st_xml = z.read("xl/richData/rdrichvaluestructure.xml").decode("utf-8", "ignore")
        except KeyError:
            st_xml = ""
        for s_blk in _re.findall(r"<s\b[^>]*>(.*?)</s>", st_xml, _re.S):
            struct_keys.append(_re.findall(r'<k n="([^"]+)"', s_blk))

    # 3) rv 列表: (s 序号, 字段值列表)
    rvs: List[Tuple[int, List[str]]] = []
    for s_idx, body in _re.findall(r'<rv\s+s="(\d+)"\s*>(.*?)</rv>', rd_xml, _re.S):
        rvs.append((int(s_idx), _re.findall(r"<v>([^<]*)</v>", body)))

    # 4) 两条 rel 链
    web_blip: List[Optional[str]] = []
    if "xl/richData/rdRichValueWebImage.xml" in names:
        try:
            wi_xml = z.read("xl/richData/rdRichValueWebImage.xml").decode("utf-8", "ignore")
        except KeyError:
            wi_xml = ""
        for blk in _re.findall(r"<webImageSrd>(.*?)</webImageSrd>", wi_xml, _re.S):
            m = _re.search(r'<blip r:id="(rId\d+)"', blk)
            web_blip.append(m.group(1) if m else None)

    local_rel: List[str] = []
    if "xl/richData/richValueRel.xml" in names:
        try:
            rr_xml = z.read("xl/richData/richValueRel.xml").decode("utf-8", "ignore")
        except KeyError:
            rr_xml = ""
        local_rel = _re.findall(r'<rel r:id="(rId\d+)"\s*/>', rr_xml)

    def _rid_targets(rels_name: str) -> Dict[str, str]:
        if rels_name not in names:
            return {}
        try:
            xml = z.read(rels_name).decode("utf-8", "ignore")
        except KeyError:
            return {}
        return {
            m.group(1): m.group(2)
            for m in _re.finditer(
                r'<Relationship Id="(rId\d+)"[^>]*Target="\.\./(media/[^"]+)"', xml
            )
        }

    web_targets = _rid_targets("xl/richData/_rels/rdRichValueWebImage.xml.rels")
    local_targets = _rid_targets("xl/richData/_rels/richValueRel.xml.rels")

    for row_1idx, vm in vm_by_row.items():
        rv_idx = vm - 1                      # vm 是 1-idx, rv 列表是 0-idx
        if not (0 <= rv_idx < len(rvs)):
            continue
        s_idx, values = rvs[rv_idx]
        keys = struct_keys[s_idx] if 0 <= s_idx < len(struct_keys) else []

        target: Optional[str] = None
        if "WebImageIdentifier" in keys:
            k = keys.index("WebImageIdentifier")
            if k < len(values):
                wi_id = int(values[k] or -1)
                if 0 <= wi_id < len(web_blip) and web_blip[wi_id]:
                    target = web_targets.get(web_blip[wi_id])
        elif "_rvRel:LocalImageIdentifier" in keys:
            k = keys.index("_rvRel:LocalImageIdentifier")
            if k < len(values):
                li = int(values[k] or -1)
                if 0 <= li < len(local_rel):
                    target = local_targets.get(local_rel[li])
        if not target:
            continue

        try:
            data = z.read(f"xl/{target}")
        except KeyError:
            continue
        fmt = "png" if target.lower().endswith(".png") else "jpeg"
        out.setdefault(row_1idx, []).append({"bytes": data, "format": fmt})
    return out


def load_images_by_row_per_sheet(primary_file: Path) -> Dict[str, Dict[int, List[Dict[str, object]]]]:
    """按 sheet 加载所有图,每 sheet 返回 {row_0indexed -> [images]}。

    适用:多 sheet 文件,每 sheet 每行 1 张图(V1044 RATTI / V1046 LOSCHI / V1048 MALTEMPI)。
    与 ``load_sheet_embedded_images`` 的区别:
      - load_sheet_embedded_images 是 V1051 SALOMON 专用,从 descr 里抽 SKU 来匹配
      - 本函数直接用 anchor._from.row,适合"图锚点 = 数据行"的简单版式
      - 而且支持多 sheet(每个 sheet 独立索引)

    Returns: {sheet_name -> {row_0indexed -> [{bytes, format}]}}
    """
    path_str = str(primary_file)
    if not path_str.lower().endswith(".xlsx"):
        return {}
    try:
        wb = openpyxl.load_workbook(path_str)
    except Exception:
        return {}
    out: Dict[str, Dict[int, List[Dict[str, object]]]] = {}
    for ws in wb.worksheets:
        sheet_map: Dict[int, List[Dict[str, object]]] = {}
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
            sheet_map.setdefault(row, []).append({
                "bytes": img_bytes,
                "format": img.format or "jpeg",
            })
        out[ws.title] = sheet_map
    return out


def load_images_by_merged_range(
    primary_file: Path,
    image_col: int = 1,
) -> Dict[str, Dict[int, List[Dict[str, object]]]]:
    """按 sheet 加载所有图,基于合并单元格范围把图"分配"到覆盖的每行。

    适用:V1043 PRITELLI 这种"图片所在单元格是几行合并"的版式:
      - 图片锚点 cell 处于一个 merged cell range(例如 A11:A13)
      - 这个 range 里的所有行(11, 12, 13)都应该用同一张图
      - 数据行落在 merge 范围外时,该行无图(不被 spillover 错配)

    与 ``load_images_by_row_per_sheet`` 的区别:
      - load_images_by_row_per_sheet 直接按 anchor._from.row 分桶,1 行 1 张图
      - 本函数按"图所在的合并区范围"分桶,N 行共享 1 张图

    实现:扫每张图的 anchor._from,找到它所在的 merged cell range,
    把图复制到 range 内每个 row 上。

    Returns: {sheet_name -> {row_0indexed -> [{bytes, format}]}}
    """
    path_str = str(primary_file)
    if not path_str.lower().endswith(".xlsx"):
        return {}
    try:
        wb = openpyxl.load_workbook(path_str)
    except Exception:
        return {}

    out: Dict[str, Dict[int, List[Dict[str, object]]]] = {}
    for ws in wb.worksheets:
        # 1) 把 merged ranges 按 "anchor cell" 索引,方便用 anchor._from 反查
        #    cell (col, row) → (min_row, max_row) 当 col == image_col
        #    坐标系说明:
        #    - ws.merged_cells.ranges 用 1-indexed Excel 坐标
        #    - img.anchor._from 用 0-indexed openpyxl 坐标
        #    所以 look up 时 anchor row 要 +1 才能匹配
        col_merges: Dict[Tuple[int, int], Tuple[int, int]] = {}
        for m in ws.merged_cells.ranges:
            if m.min_col <= image_col <= m.max_col:
                # 存 1-indexed → 1-indexed
                col_merges[(m.min_col, m.min_row)] = (m.min_row, m.max_row)

        sheet_map: Dict[int, List[Dict[str, object]]] = {}

        for img in ws._images:
            anchor = img.anchor
            if not hasattr(anchor, "_from") or anchor._from is None:
                continue
            col = anchor._from.col      # 0-indexed openpyxl
            row = anchor._from.row      # 0-indexed openpyxl
            try:
                img_bytes = img._data() if hasattr(img, "_data") else None
            except Exception:
                img_bytes = None
            if img_bytes is None:
                continue
            item = {
                "bytes": img_bytes,
                "format": img.format or "jpeg",
            }
            # 转 1-indexed,查 merge 范围
            excel_col = col + 1
            excel_row = row + 1
            row_span = col_merges.get((excel_col, excel_row))
            if row_span is None:
                # 兜底:扫所有 col 跨度上的 merge(横跨 image_col)
                for (c, _), (mn, mx) in col_merges.items():
                    if c != excel_col:
                        continue
                    if mn <= excel_row <= mx:
                        row_span = (mn, mx)
                        break
            if row_span is None:
                # 单格图(没合并)— 只覆盖 anchor 那一行
                sheet_map.setdefault(row, []).append(item)
            else:
                # 把图复制到 merge 范围的所有行(0-indexed)
                for r_excel in range(row_span[0], row_span[1] + 1):
                    r0 = r_excel - 1
                    sheet_map.setdefault(r0, []).append(item)
        out[ws.title] = sheet_map
    return out


def load_sheet_embedded_images(primary_file: Path) -> Dict[str, Dict[int, List[Dict[str, object]]]]:
    """
    Load embedded images from Excel, using descr-based SKU lookup.
    
    Returns: Dict[sheet_name -> Dict[excel_row -> List[image_items]]]
    
    Image items only contain 'bytes' and 'format' keys (descr/sku metadata
    from _descr/_sku are stripped for compatibility with write_output).
    
    When the drawing XML has per-row SKU info (descr contains SKU code like
    "100_L47179800.png"), we use that to find the correct Excel row.
    Falls back to anchor._from.row for files without SKU in descr.
    """
    raw = _load_drawing_images_via_xml(primary_file)
    
    # Convert: sheet -> {excel_row -> [image_items]} (strip internal metadata)
    out: Dict[str, Dict[int, List[Dict[str, object]]]] = {}
    for sheet_name, row_map in raw.items():
        cleaned: Dict[int, List[Dict[str, object]]] = {}
        for row, items in row_map.items():
            cleaned_items = []
            seen_hashes = set()
            for item in items:
                # Deduplicate by content hash
                h = hashlib.md5(item['bytes']).hexdigest()
                if h in seen_hashes:
                    continue
                seen_hashes.add(h)
                cleaned_items.append({
                    "bytes": item["bytes"],
                    "format": item.get("format", "png"),
                })
            cleaned[row] = cleaned_items
        out[sheet_name] = cleaned
    
    return out


def load_sheet_row_heights(primary_file: Path) -> Dict[str, Dict[int, float]]:
    wb = openpyxl.load_workbook(primary_file)
    out: Dict[str, Dict[int, float]] = {}
    for ws in wb.worksheets:
        default_height = float(ws.sheet_format.defaultRowHeight or 15)
        row_map: Dict[int, float] = {}
        for row_idx, row_dim in ws.row_dimensions.items():
            if row_dim is None:
                continue
            h = row_dim.height
            if h is None:
                continue
            row_map[int(row_idx)] = float(h)
        if default_height > 0:
            row_map.setdefault(0, default_height)
        out[ws.title] = row_map
    return out


# ---------------------------------------------------------------------------
# Scalar size-map builder (shared by V1039, V1048)
# ---------------------------------------------------------------------------


def build_scalar_size_map(
    df: pd.DataFrame,
    header_row: int,
    scalar_col: int,
    qty_start_col: int,
    qty_end_col: int,
) -> Dict[str, Dict[int, str]]:
    out: Dict[str, Dict[int, str]] = {}
    for r in range(0, header_row):
        scalar = normalize_text(df.iat[r, scalar_col] if scalar_col < df.shape[1] else None)
        if not scalar:
            continue
        col_map: Dict[int, str] = {}
        for c in range(qty_start_col, min(qty_end_col + 1, df.shape[1])):
            label = normalize_text(df.iat[r, c])
            if label:
                col_map[c] = label
        if col_map:
            out[scalar.upper()] = col_map
    return out


# ---------------------------------------------------------------------------
# Template output writing
# ---------------------------------------------------------------------------


class _SharedMediaImage(XLImage):
    """相同内容的图片共用同一个 media 文件 (2026-10-26).

    openpyxl 默认给**每个** Image 对象分配独立的媒体文件 (xl/media/imageN.ext)。
    但同一个 SKU 展开成多个尺码行时, 同一张图会被引用几十次 ——
    实测 Boss&Hugo 50572 行只用了 5304 张不同的图 (重复率 89.5%),
    写进 xlsx 就变成 50572 份几乎一样的媒体文件, 体积和写盘时间都被放大。

    Excel 里同一张图被多处引用是**合法**的: 多个 drawing anchor 各自带 rId,
    但 Target 可以指向同一个 media 文件。所以这里让相同内容的图共用 media,
    再由 _patch_openpyxl_dedup_media() 保证 zip 里不出现重复条目。
    """

    _media_idx = 1

    @property
    def path(self) -> str:
        return self._path.format(self._media_idx, self.format)


def _patch_openpyxl_dedup_media() -> None:
    """让 openpyxl 保存时同一 media 文件只写一次 (配合 _SharedMediaImage).

    同一个 media path 被多个 anchor 引用是合法的, 但 zip 归档里不能出现
    重复条目 (Excel 会报文件损坏), 所以写 media 时按 path 去重。

    只影响 _SharedMediaImage (path 由内容哈希决定); 普通 XLImage 的 path
    仍按 _id 分配, 互不冲突。
    """
    from openpyxl.writer.excel import ExcelWriter
    if getattr(ExcelWriter, "_mavis_dedup_patched", False):
        return

    def _write_images(self) -> None:
        seen = set()
        for img in self._images:
            path = img.path[1:]
            if path in seen:
                continue
            seen.add(path)
            self._archive.writestr(path, img._data())

    ExcelWriter._write_images = _write_images
    ExcelWriter._mavis_dedup_patched = True


def embed_images_in_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row_idx: int,
    col_idx: int,
    image_items: Sequence[Dict[str, object]],
    cache: Dict[str, Path],
    tmp_dir: Path,
    max_height_emu: Optional[int] = None,
    media_map: Optional[Dict[str, int]] = None,
) -> None:
    """Embed images in a worksheet row.

    - Images are written at their ORIGINAL pixel dimensions (no scaling) unless
      max_height_emu is set, in which case they're scaled down to fit that
      height (keeping aspect ratio). Image BYTES are never re-encoded —
      only the displayed <ext cx/cy> changes, so clarity is preserved.
    - Row height is set separately via write_output() using source_row_height
      or image_row_height.
    - media_map: 内容哈希 → media 序号。相同内容的图片共用同一个媒体文件
      (见 _SharedMediaImage), 传 None 则退回 openpyxl 默认的每图一份。
    """
    if not image_items:
        return
    running_col_off_emu = 0
    gap_px = 8
    for image_item in image_items:
        blob = image_item.get("bytes")
        img_format = normalize_text(image_item.get("format")).lower() or "png"
        if not isinstance(blob, (bytes, bytearray)):
            continue
        # 2026-08-18: 标准化为 openpyxl 支持的格式 (png/jpeg),避免 jp2/jpeg2000 等
        if img_format not in ("png", "jpeg", "jpg"):
            try:
                from PIL import Image
                import io as _io
                src = Image.open(_io.BytesIO(blob))
                # 转换 RGBA → RGB(避免保存 JPEG 失败)
                if src.mode in ("RGBA", "LA", "P"):
                    src = src.convert("RGB")
                buf = _io.BytesIO()
                src.save(buf, format="PNG")
                blob = buf.getvalue()
                img_format = "png"
            except Exception:
                # PIL 加载失败 — fallback 强制 png 让 openpyxl 自己决定
                img_format = "png"
        digest = hashlib.md5()
        digest.update(blob)
        digest.update(img_format.encode("utf-8"))
        key = digest.hexdigest()
        file_path = cache.get(key)
        if file_path is None:
            ext = "jpg" if img_format in {"jpeg", "jpg"} else "png"
            file_path = tmp_dir / f"{key}-{uuid.uuid4().hex[:8]}.{ext}"
            file_path.write_bytes(bytes(blob))
            cache[key] = file_path

        img = _SharedMediaImage(str(file_path))
        if media_map is not None:
            # 相同内容 → 同一个 media 序号 (媒体文件只写一份)
            if key not in media_map:
                media_map[key] = len(media_map) + 1
            img._media_idx = media_map[key]
        # Use display_emu from image_item if available (preserves source display size),
        # otherwise fall back to original image dimensions.
        display_emu = image_item.get("display_emu")
        if display_emu:
            cx, cy = display_emu
        else:
            cx = int(img.width * 9525)
            cy = int(img.height * 9525)

        # 按行高上限缩放（只改显示尺寸，不动原图字节 → 清晰度保留）
        if max_height_emu and cy > max_height_emu:
            scale = max_height_emu / cy
            cx = int(cx * scale)
            cy = max_height_emu

        anchor = OneCellAnchor(
            _from=AnchorMarker(
                col=col_idx - 1,
                row=row_idx - 1,
                colOff=running_col_off_emu,
                rowOff=0,
            ),
            ext=XDRPositiveSize2D(cx=cx, cy=cy),
        )
        img.anchor = anchor
        ws.add_image(img)
        # ⚠️ 用**显示宽度** cx 累加偏移,不能用 img.width(原图宽度)
        # 否则原图 500x500 缩到 92pt 后,colOff 仍按 500px 算,第二张图会被推到右边很远
        running_col_off_emu += cx + int(gap_px * 9525)


def map_row_to_template(row: Dict[str, object], headers: Sequence[str]) -> List[object]:
    hset = set(headers)
    is_cn = "品牌" in hset
    if is_cn:
        mapped = {
            "品牌": row["brand"],
            "货号": row["sku"],
            "图片": row["image"],
            "描述": row["description"],
            "类目": row["category"],
            "材质": row["material"],
            "颜色": row["color"],
            "色号": row.get("color_code", ""),  # 2026-07-30:暂时只有 V1050 + 中文模板填(V1050 v1 parser 输出 color_code 字段)
            "季节": row["season"],
            "产地": row["origin"],
            "cost": row["cost"],
            "C+-": row["cplus_text"],
            "批发价": row["supply_price"],
            "零售价": row["retail_price"],
            "尺码": row["size"],
            "数量": row["quantity"],
            "云仓": row["vendor_id"],
            "性别": row.get("gender", ""),  # 2026-07-31:暂时只有 V1070 / V1015 + 中文模板填(其他 vendor 没 gender 也不报错)
        }
    else:
        mapped = {
            "Brand": row["brand"],
            "International code": row["sku"],
            "Image": row["image"],
            "Described": row["description"],
            "Category": row["category"],
            "Material": row["material"],
            "Gender": row.get("gender", ""),  # 2026-07-31:暂时只有 V1070 / V1015 + 中文模板填(其他 vendor 没 gender 也不报错)
            "Color": row["color"],
            "Color Code": row.get("color_code", ""),  # 2026-07-30:暂时只有 V1050 + 中文模板填(V1050 v1 parser 输出 color_code 字段)
            "Season": row["season"],
            "Made in": row["origin"],
            "cost": row["cost"],
            "C+-": row["cplus_text"],
            "Wholesale price": row["supply_price"],  # 2026-09-09:原 Supply price 改名 Wholesale price,跟 CN "批发价" 对齐
            "Retail price": row["retail_price"],
            "Size": row["size"],
            "Quantity": row["quantity"],
            "Vendor ID": row["vendor_id"],
        }
    return [mapped.get(h, "") for h in headers]


def validate_required_headers(headers: Sequence[str]) -> None:
    hset = set(headers)
    if "品牌" in hset:
        missing = [c for c in REQ_COLS_CN if c not in hset]
    else:
        missing = [c for c in REQ_COLS_EN if c not in hset]
    if missing:
        raise ValueError(f"Template missing required fields: {', '.join(missing)}")


def write_output(
    template_file: Path,
    output_file: Path,
    rows: List[Dict[str, object]],
    image_row_height: Optional[float] = None,
    image_col_width: Optional[float] = None,
    keep_images_per_row: Optional[int] = None,
) -> None:
    """Write converted rows to a template xlsx.

    image_row_height: 有图片的数据行的目标行高（pt）。
    - 设为 92 → 92pt（POLO 临时任务，按 image-handling.md skill 规则）
    - 设为 None → 沿用 source_row_height 或模板默认
    - 图片会按 max_height_emu = 行高 × 12700 缩放，保留原图字节（清晰度不变）

    image_col_width: 图片列的目标列宽(Excel 字符单位)。
    - 多图产品(如 V1049 每个产品 3-4 张图)需要更宽的列才能完整显示
    - 设为 None → 沿用模板的默认列宽
    """
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    headers = [normalize_text(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    validate_required_headers(headers)
    image_col_name = "图片" if "图片" in headers else "Image"
    image_col_idx = headers.index(image_col_name) + 1 if image_col_name in headers else None

    tmp_dir = TMP_DIR / uuid.uuid4().hex[:12]
    tmp_dir.mkdir(parents=True, exist_ok=True)
    image_cache: Dict[str, Path] = {}
    # 2026-10-26: 相同内容的图片共用 media 文件 (同 SKU 多尺码行会重复引用同一张图)
    media_map: Dict[str, int] = {}
    _patch_openpyxl_dedup_media()

    # 行高上限（pt → EMU）。openpyxl 的 row_dimensions.height 单位是 pt。
    max_height_pt: Optional[float] = None
    max_height_emu: Optional[int] = None
    if image_row_height is not None and image_row_height > 0:
        max_height_pt = float(image_row_height)
        max_height_emu = int(image_row_height * 12700)

    # 图片列宽（Excel 字符单位）
    if image_col_width is not None and image_col_width > 0 and image_col_idx is not None:
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(image_col_idx)
        ws.column_dimensions[col_letter].width = float(image_col_width)

    # 货号列宽(2026-07-17 用户要求:模板默认 13 太短,统一扩到 26)
    # 货号列:中文模板是 col 2(货号),英文模板是 col 2(International code)
    sku_col_idx = headers.index("货号") + 1 if "货号" in headers else (
        headers.index("International code") + 1 if "International code" in headers else None
    )
    if sku_col_idx is not None and SKU_COL_WIDTH > 0:
        from openpyxl.utils import get_column_letter
        sku_col_letter = get_column_letter(sku_col_idx)
        ws.column_dimensions[sku_col_letter].width = float(SKU_COL_WIDTH)

    # 2026-10-19:新模板 max_row 可能到 140-157 (空样式占位),
    # 不能用 ws.max_row + 1,改成找最后非空行 + 1
    next_row = 2
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            next_row = r + 1
            break
    # 2026-07-21:per-vendor keep_images_per_row(默认 KEEP_IMAGES_PER_ROW=1)
    # V1003 Sirio DB 等需要每行多图(vendor_meta 设置 default_keep_images_per_row=2)
    keep_n = keep_images_per_row if keep_images_per_row is not None else KEEP_IMAGES_PER_ROW
    for r in rows:
        mapped = map_row_to_template(r, headers)
        # 2026-07-17:全局只保留第一张图,避免多图撑爆图片列
        # 2026-07-21:V1003 等 vendor 通过 keep_images_per_row 覆盖
        image_items = r.get("image_items", [])[:keep_n]
        has_image = bool(image_items)

        if has_image and max_height_pt is not None:
            # 有图片 + 指定了行高 → 用指定行高
            ws.row_dimensions[next_row].height = max_height_pt
        else:
            # 否则沿用源行高
            src_row_h = to_float(r.get("source_row_height"))
            if src_row_h is not None and src_row_h > 0:
                ws.row_dimensions[next_row].height = src_row_h

        for idx, value in enumerate(mapped, start=1):
            ws.cell(next_row, idx, value=value)
        if image_col_idx is not None:
            embed_images_in_row(
                ws, next_row, image_col_idx,
                image_items, image_cache, tmp_dir,
                max_height_emu=max_height_emu if has_image else None,
                media_map=media_map,
            )
        next_row += 1
    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)
