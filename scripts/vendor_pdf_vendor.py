"""PDF vendor adapter — 通用 PDF 货单配置器(2026-08-15 新建)

支持的版式(从 aliyun OSS top2-b2b-prod 下来的货单):
- 表格化版式,列: 序号 | 品牌 | 款号图片 | 款号 | 男女装 | 批发价 | 秒杀价 | 打包价 | 尺码 | 可配
- 一个 SKU 多个尺码(每个尺码一行 + 合计行)
- 每页 9 个产品图(x=121-191, y 方向每 84pt 一张)
- 图居中,SKU 文字在图的两侧

依赖: pdfplumber (抽文字) + PyMuPDF/fitz (抽图)
"""
from __future__ import annotations

import io
import mimetypes
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import pdfplumber

# 2026-08-18: 注册 openpyxl 缺的 MIME type,修 PyMuPDF 返回 jp2/jpeg2000 时
# openpyxl 写文件报 KeyError: '.jpeg2000'
# 同时写进 4 个地方确保生效(顶层 types_map / types_map[True] / strict / non-strict)
JP2_EXTS = [".jp2", ".jpeg2000", ".jpx", ".j2k", ".jpf"]
for ext in JP2_EXTS:
    mime = "image/jp2"
    mimetypes.add_type(mime, ext, strict=True)
    mimetypes.add_type(mime, ext)
    try:
        mimetypes.types_map[ext] = mime
    except Exception:
        pass
    try:
        mimetypes.types_map[True][ext] = mime
    except (KeyError, TypeError):
        try:
            mimetypes._db
        except Exception:
            pass
        mimetypes.add_type(mime, ext, strict=True)
        try:
            mimetypes.types_map[True] = mimetypes._db.types_map[1]
        except Exception:
            pass
    try:
        if mimetypes._db is not None and len(mimetypes._db.types_map) > 1:
            mimetypes._db.types_map[1][ext] = mime
    except Exception:
        pass

# 列 x 范围(根据 PDF1 采样)
# - 序号: ~40
# - 品牌: ~74
# - 款号: ~204-237 (2 行)
# - 男女装: ~261-289
# - 批发价: ~319
# - 尺码: ~486
# - 可配: ~546
COL_INDEX_X = 40
COL_BRAND_X = 74
COL_SKU_X_MIN = 200
COL_GENDER_X_MIN = 250
COL_COST_X_MIN = 300
COL_SIZE_X_MIN = 470
COL_QTY_X_MIN = 530

# y 方向分桶容差(2pt 内算同一行)
Y_TOLERANCE = 2.0
# 图 y 与 SKU y 配对容差(±30pt)
IMAGE_SKU_TOLERANCE = 50.0

# 尺码字符白名单(常见尺码 label)
SIZE_LABELS = {
    "UNI", "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
    "3XL", "4XL", "5XL", "6XL",
    "ITU", "ITD", "UKU", "UKD", "CMU", "CMD",  # 鞋码 / 厘米
    "PZ", "TU", "OS", "ONE",
}

VENDOR_ID = "PDF_VENDOR"  # 临时命名


def _norm_text(s: str) -> str:
    if not s:
        return ""
    return s.strip()


def _is_number(s: str) -> bool:
    if not s:
        return False
    try:
        float(s.replace(",", ""))
        return True
    except (ValueError, AttributeError):
        return False


def _group_words_by_y(words: List[Dict], tolerance: float = Y_TOLERANCE) -> List[List[Dict]]:
    """按 y (top) 坐标把 word 分桶,容差 tolerance 内算同一行。"""
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    groups: List[List[Dict]] = []
    current_group: List[Dict] = []
    current_y = None
    for w in sorted_words:
        if current_y is None or abs(w["top"] - current_y) <= tolerance:
            current_group.append(w)
            current_y = current_y if current_y is not None else w["top"]
        else:
            groups.append(current_group)
            current_group = [w]
            current_y = w["top"]
    if current_group:
        groups.append(current_group)
    return groups


def _parse_y_group(group: List[Dict]) -> Dict[str, object]:
    """解析一个 y 桶,按 x 位置分类到不同字段。返回 dict 含:
      - index: 序号(数字)
      - brand: 品牌
      - sku: 款号(多行可能,合并)
      - gender: 男女装
      - cost: 批发价
      - size: 尺码
      - qty: 可配(数字)
      - is_total: 是否合计行
      - y: 该组 top 坐标
    """
    y = sum(w["top"] for w in group) / len(group)
    parts = {"index": "", "brand": "", "sku_parts": [], "gender": "", "cost": "", "size": "", "qty": "", "is_total": False}
    for w in sorted(group, key=lambda x: x["x0"]):
        text = _norm_text(w["text"])
        if not text:
            continue
        x0 = w["x0"]
        # 合计行
        if text == "合计" and 480 <= x0 <= 510:
            parts["is_total"] = True
            continue
        # 序号
        if x0 < COL_INDEX_X + 20 and _is_number(text):
            parts["index"] = text
            continue
        # 品牌
        if COL_BRAND_X - 10 <= x0 <= COL_BRAND_X + 50:
            # 排除 header 词"品牌"
            if text in ("品牌", "品", "牌"):
                continue
            parts["brand"] = text
            continue
        # 款号
        if COL_SKU_X_MIN <= x0 <= COL_SKU_X_MIN + 50:
            # 排除 header 词"款号"
            if text in ("款号",):
                continue
            parts["sku_parts"].append(text)
            continue
        # 男女装
        if COL_GENDER_X_MIN <= x0 <= COL_GENDER_X_MIN + 50:
            if text in ("男女装",):
                continue
            parts["gender"] = text
            continue
        # 批发价
        if COL_COST_X_MIN <= x0 <= COL_COST_X_MIN + 60:
            if text in ("批发价", "秒杀价", "打包价"):
                continue
            if _is_number(text):
                parts["cost"] = text
                continue
        # 尺码
        if COL_SIZE_X_MIN <= x0 <= COL_SIZE_X_MIN + 50:
            if text in ("尺码",):
                continue
            parts["size"] = text
            continue
        # 可配
        if COL_QTY_X_MIN <= x0 <= COL_QTY_X_MIN + 30:
            if text in ("可配",):
                continue
            if _is_number(text):
                parts["qty"] = text
                continue
    parts["y"] = y
    # 合并款号(可能跨 2 行)
    parts["sku"] = "".join(parts["sku_parts"])
    return parts


def _extract_page_data(page, page_idx: int) -> List[Dict[str, object]]:
    """从一页 PDF 抽所有 SKU 块。返回 list of dict,每个 dict 含:
      - brand, sku, gender, cost (main row)
      - size, qty (per size sub-row)
      - y (main row y 坐标,用于配图)

    实际版式(2026-08-15 实测 PDF1/PDF2):
      一个 SKU 块由"size rows + 合计行 + main row"组成,其中 size/合计 在 main 的**上面**(更小 y)。
      例(SKU 1):
        y=34.8: UNI 4
        y=46.7: 合计 4
        y=66:   3B00037-        (sku part 1)
        y=72:   1/MONCLER/MEN/1670.0 (序号/品牌/男女装/批发价)
        y=78:   04863-999       (sku part 2)

    SKU main row 跨 3 个 y 桶,所以不能用"单桶有 sku+brand+gender"判定。改成:
      - 序号 (1-3 位数字在 col 40) = main row 起始
      - 同 y 桶的 brand+gender+cost 一起算
      - y 桶下面 y+0~7pt 的 sku 段(part 1+part 2) 拼起来
      - size rows = main.y 上面、prev_main.y 下面的所有 size+合计行(去合计)
    """
    words = page.extract_words()
    groups = _group_words_by_y(words)

    # 第一遍:用更宽松的判定收集每桶的字段(sku part 1/2 单独存)
    parsed: List[Dict[str, object]] = []
    for g in groups:
        y = sum(w["top"] for w in g) / len(g)
        parts = {
            "y": y,
            "index": "",
            "brand": "",
            "sku": "",  # single-line sku(如果款号只有 1 行)
            "gender": "",
            "cost": "",
            "size": "",
            "qty": "",
            "is_total": False,
        }
        for w in sorted(g, key=lambda x: x["x0"]):
            text = _norm_text(w["text"])
            if not text:
                continue
            x0 = w["x0"]
            if text == "合计" and 480 <= x0 <= 510:
                parts["is_total"] = True
                continue
            if x0 < COL_INDEX_X + 20 and _is_number(text) and not parts["index"]:
                parts["index"] = text
                continue
            if COL_BRAND_X - 10 <= x0 <= COL_BRAND_X + 120:  # 加宽到 130,容下 "LORO PIANA" 双词
                if text in ("品牌", "品", "牌"):
                    continue
                # 同一 y 桶可能多词,用空格连接
                parts["brand"] = (parts["brand"] + " " + text).strip() if parts["brand"] else text
                continue
            if COL_SKU_X_MIN <= x0 <= COL_SKU_X_MIN + 50:
                if text in ("款号",):
                    continue
                parts["sku"] = parts["sku"] + text  # 累积多行款号
                continue
            if COL_GENDER_X_MIN <= x0 <= COL_GENDER_X_MIN + 50:
                if text in ("男女装",):
                    continue
                if not parts["gender"]:
                    parts["gender"] = text
                continue
            if COL_COST_X_MIN <= x0 <= COL_COST_X_MIN + 60:
                if text in ("批发价", "秒杀价", "打包价"):
                    continue
                if _is_number(text) and not parts["cost"]:
                    parts["cost"] = text
                    continue
            if COL_SIZE_X_MIN <= x0 <= COL_SIZE_X_MIN + 50:
                if text in ("尺码",):
                    continue
                if not parts["size"]:
                    parts["size"] = text
                continue
            if COL_QTY_X_MIN <= x0 <= COL_QTY_X_MIN + 30:
                if text in ("可配",):
                    continue
                if _is_number(text) and not parts["qty"]:
                    parts["qty"] = text
                    continue
        parsed.append(parts)

    # 第二遍:找出 main rows 和 size rows
    # main row 判定:有 序号 + 品牌 + 男女装(款号可能跨多行,在前后桶里)
    # 款号(part 1+part 2)通过往前/往后扫 1-2 桶拼接
    # 品牌(如 CHRISTIAN DIOR / LOUIS VUITTON / MAX MARA)也可能跨 2 行,扫前后桶
    main_rows: List[Dict[str, object]] = []
    size_rows: List[Dict[str, object]] = []
    used_indices: set = set()  # 记录已被 main row 消耗的桶 idx
    for i, p in enumerate(parsed):
        if i in used_indices:
            continue
        if p["index"] and p["brand"] and p["gender"]:
            # 找款号:往前 1 桶 + 往后 1 桶(part 1 在前,part 2 在后)
            sku = p["sku"]
            if not sku:
                # 往前 1 桶
                if i - 1 >= 0 and parsed[i - 1]["sku"] and parsed[i - 1]["index"] == "":
                    sku = parsed[i - 1]["sku"]
                    used_indices.add(i - 1)
                # 往后 1 桶
                if i + 1 < len(parsed) and parsed[i + 1]["sku"] and parsed[i + 1]["index"] == "":
                    sku = sku + parsed[i + 1]["sku"]
                    used_indices.add(i + 1)
            if not sku:
                continue  # 没找到款号,跳过
            # 品牌:可能多词(CHRISTIAN DIOR / LOUIS VUITTON / MAX MARA 等跨 2 行)
            # 往前/后 1 桶,合并 brand(不消耗桶,只参考)
            brand = p["brand"]
            for adj in (i - 1, i + 1):
                if 0 <= adj < len(parsed) and adj not in used_indices:
                    if parsed[adj]["brand"] and not parsed[adj]["sku"] and not parsed[adj]["index"]:
                        # 邻桶只有 brand,可能是品牌第 2 词
                        if parsed[adj]["y"] < p["y"]:
                            brand = parsed[adj]["brand"] + " " + brand
                        else:
                            brand = brand + " " + parsed[adj]["brand"]
            main_rows.append({
                "brand": brand,
                "sku": sku,
                "gender": p["gender"],
                "cost": p["cost"],
                "y": p["y"],
            })
            used_indices.add(i)
        elif p["is_total"] or (p["size"] and p["qty"]):
            size_rows.append(p)

    # 第三遍:每个 main row,从 size_rows 里挑"y 在 (prev_main.y, main.y)" 的归属
    blocks: List[Dict[str, object]] = []
    for i, main in enumerate(main_rows):
        prev_main_y = main_rows[i - 1]["y"] if i > 0 else 0
        related = [
            sr for sr in size_rows
            if sr["y"] < main["y"] and sr["y"] > prev_main_y
        ]
        blocks.append({
            "brand": main["brand"],
            "sku": main["sku"],
            "gender": main["gender"],
            "cost": main["cost"],
            "y": main["y"],
            "sizes": [
                {"size": sr["size"], "qty": int(sr["qty"]), "y": sr["y"]}
                for sr in related
                if sr["size"] and sr["qty"] and not sr["is_total"]
            ],
        })
    return blocks


def _extract_images_by_y(doc: fitz.Document, page_idx: int) -> List[Dict[str, object]]:
    """从一页 PDF 抽所有图,返回 [{y, bytes, format}]。y 是图垂直中心。"""
    page = doc[page_idx]
    out = []
    for img in page.get_images(full=True):
        xref = img[0]
        bbox = page.get_image_bbox(img)
        y_center = (bbox[1] + bbox[3]) / 2
        try:
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]
            ext = base_image.get("ext", "jpeg")
        except Exception:
            continue
        out.append({
            "y": y_center,
            "bytes": img_bytes,
            "format": ext if ext in ("jpeg", "png", "gif", "bmp") else "jpeg",
        })
    return out


def _match_image_to_block(block: Dict, page_images: List[Dict]) -> Optional[Dict]:
    """找离 block y 最近的图。"""
    if not page_images:
        return None
    block_y = block["y"]
    best = None
    best_dist = float("inf")
    for img in page_images:
        dist = abs(img["y"] - block_y)
        if dist < best_dist:
            best_dist = dist
            best = img
    if best_dist > IMAGE_SKU_TOLERANCE:
        return None
    return best


def _normalize_gender(g: str) -> str:
    """MEN → Man, WOMEN → Woman, UNISEX → Unisex, etc."""
    if not g:
        return ""
    g_low = g.lower().strip()
    if g_low in ("men", "man", "uomo"):
        return "Man"
    if g_low in ("women", "woman", "donna"):
        return "Woman"
    if g_low in ("unisex", "unica"):
        return "Unisex"
    if g_low in ("boy", "kid", "kids", "junior"):
        return "Kid"
    if g_low in ("girl",):
        return "Girl"
    return g


def _normalize_brand(b: str) -> str:
    """MONCLER → Moncler, CELINE → Celine (首字母大写)"""
    if not b:
        return ""
    return b.capitalize() if b.isupper() else b.title() if b.isupper() else b


def _parse_moncler_pdf(
    primary_file: Path,
    vendor_id: str,
) -> List[Dict[str, object]]:
    """MONCLER 版式 PDF 解析(原 parse_rows,2026-08-18 拆出来)。"""
    rows: List[Dict[str, object]] = []
    primary_file = Path(primary_file)
    if not primary_file.exists() or primary_file.suffix.lower() != ".pdf":
        return rows

    try:
        doc = fitz.open(str(primary_file))
    except Exception:
        return rows

    try:
        with pdfplumber.open(str(primary_file)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                blocks = _extract_page_data(page, page_idx)
                if not blocks:
                    continue
                page_images = _extract_images_by_y(doc, page_idx)
                for block in blocks:
                    img = _match_image_to_block(block, page_images)
                    brand = _normalize_brand(block.get("brand", ""))
                    sku = block.get("sku", "")
                    gender = _normalize_gender(block.get("gender", ""))
                    try:
                        cost = float(block.get("cost", "0").replace(",", "")) if block.get("cost") else None
                    except (ValueError, AttributeError):
                        cost = None
                    for sz in block.get("sizes", []):
                        try:
                            qty = int(sz["qty"])
                        except (ValueError, TypeError):
                            continue
                        if qty <= 0:
                            continue
                        rows.append({
                            "brand": brand,
                            "sku": sku,
                            "image": "",
                            "description": "",
                            "category": "",
                            "material": "",
                            "color": "",
                            "color_code": "",
                            "season": "",
                            "origin": "",
                            "cost": cost,
                            "cplus_text": "",
                            "supply_price": None,
                            "retail_price": None,
                            "size": sz["size"],
                            "quantity": qty,
                            "vendor_id": vendor_id,
                            "gender": gender,
                            "image_items": [img] if img else [],
                        })
    finally:
        doc.close()
    return rows


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """PDF vendor 解析入口 — 自动检测 PANICALE / MONCLER 版式(2026-08-18 加)。"""
    if not primary_file.exists():
        return []
    # PANICALE 版式检测 — 扫前 3 页找 ARTICOLO/ARTICLE 文本
    if _detect_panicale_layout(primary_file):
        return _parse_panicale_pdf(primary_file, vendor_id)
    # 否则走原 MONCLER 版式
    return _parse_moncler_pdf(primary_file, vendor_id)
    rows: List[Dict[str, object]] = []
    primary_file = Path(primary_file)
    if not primary_file.exists() or primary_file.suffix.lower() != ".pdf":
        return rows

    try:
        doc = fitz.open(str(primary_file))
    except Exception:
        return rows

    try:
        with pdfplumber.open(str(primary_file)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                blocks = _extract_page_data(page, page_idx)
                if not blocks:
                    continue
                page_images = _extract_images_by_y(doc, page_idx)
                for block in blocks:
                    img = _match_image_to_block(block, page_images)
                    brand = _normalize_brand(block.get("brand", ""))
                    sku = block.get("sku", "")
                    gender = _normalize_gender(block.get("gender", ""))
                    try:
                        cost = float(block.get("cost", "0").replace(",", "")) if block.get("cost") else None
                    except (ValueError, AttributeError):
                        cost = None
                    for sz in block.get("sizes", []):
                        try:
                            qty = int(sz["qty"])
                        except (ValueError, TypeError):
                            continue
                        if qty <= 0:
                            continue
                        rows.append({
                            "brand": brand,
                            "sku": sku,
                            "image": "",
                            "description": "",
                            "category": "",
                            "material": "",
                            "color": "",
                            "color_code": "",
                            "season": "",
                            "origin": "",
                            "cost": cost,
                            "cplus_text": "",
                            "supply_price": None,
                            "retail_price": None,
                            "size": sz["size"],
                            "quantity": qty,
                            "vendor_id": vendor_id,
                            "gender": gender,
                            "image_items": [img] if img else [],
                        })
    finally:
        doc.close()
    return rows


# ===========================================================================
# PANICALE 现货单 PDF 解析(2026-08-18 加)
# 来源:SS26_STOCKUOMO_test.pdf — 意大利 PANICALE 工厂的 stock list
# 版式:每页 5-6 个 ARTICOLO 块,每块结构:
#   R(N): "ARTICOLO/ARTICLE: <SKU> - <DESC> - <MATERIAL>  Prezzo/Price: <PRICE>"
#   R(N+1): "Colore/Color 44 46 48 50 52 54 56 58 60 62" (size 标签)
#   R(N+2)..R(N+X): 每个颜色 3 行模式 —
#     - color code 行(e.g. "38U24 0290")
#     - qty 行(在 size 列 x 位置上的数字)
#     - color name 行(e.g. "CIGAR")
# 配图:每个 ARTICOLO 块有 1-2 张产品图,在块左侧 (x~50-200, y=block_y 附近)
# 关键差异 vs MONCLER 版式:
#   - 没有固定 9 张/页排版,是 N 块/页
#   - 没有 "序号 + 品牌 + 男女装" 标准列
#   - size 标签在 Colore/Color 文本里直接列出
# ===========================================================================

# PANICALE 版式:size 列 x 中心位置 — 2026-08-18 改为**动态从 Colore/Color 行取**
# 之前硬编码 386/414/442/470/498... 只对 28pt 一档的 PDF 有效,这次碰到 30pt 一档的就全错了
# Fallback 默认值(28pt 一档 — 早期 PANICALE PDF)
PANICALE_SIZE_X_CENTERS_FALLBACK = [386, 414, 442, 470, 498, 526, 554, 582, 610, 638]
PANICALE_SIZE_LABELS = ["44", "46", "48", "50", "52", "54", "56", "58", "60", "62"]
PANICALE_X_TOLERANCE = 16.0  # ±16pt 内归属某 size 列(覆盖 30pt 一档的边界)
# 颜色代码 pattern
PANICALE_COLOR_CODE_RE = re.compile(r"^\d{2}[A-Z]+\s*\d+")
# ARTICOLO 头 pattern
PANICALE_ARTICOLO_RE = re.compile(
    r"ARTICOLO/ARTICLE:\s*(\S+)\s+-\s+(.+?)\s+Prezzo/Price:\s*([\d,]+)"
)


def _detect_panicale_layout(pdf_path: Path) -> bool:
    """检测 PDF 是否是 PANICALE 版式 — 看有没有 "ARTICOLO/ARTICLE:" 文本。"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 只扫前 3 页
            for page in pdf.pages[:3]:
                text = page.extract_text() or ""
                if "ARTICOLO/ARTICLE:" in text:
                    return True
        return False
    except Exception:
        return False


def _parse_panicale_page(
    page, fitz_page, doc, season: str, vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 PANICALE 版式单页 — 返回该页的所有 rows。

    page: pdfplumber.Page (用于 extract_words)
    fitz_page: fitz.Page (用于 get_images — PyMuPDF API)
    doc: fitz.Document
    """
    page_h = page.height
    page_w = page.width

    # 1) 拿所有 word
    words = page.extract_words() or []

    # 2) 拿所有 image (xref → bytes) — 用 fitz API
    page_images: Dict[int, Dict] = {}  # xref → {bytes, format, rects}
    for img in fitz_page.get_images(full=True):
        xref = img[0]
        try:
            base = doc.extract_image(xref)
            data = base.get("image")
            ext = base.get("ext", "png")
            if data:
                page_images[xref] = {
                    "bytes": data,
                    "format": ext.lower() if ext.lower() in ("png", "jpeg", "jpg") else "png",
                    "rects": fitz_page.get_image_rects(xref),  # list[Rect]
                }
        except Exception:
            pass

    # 3) 按 y 桶(2pt 容差)分 word
    y_groups: Dict[float, List[Dict]] = {}
    for w in words:
        y = round(w["top"])
        y_groups.setdefault(y, []).append(w)

    sorted_ys = sorted(y_groups.keys())

    # 4) 找所有 ARTICOLO/ARTICLE 行(y 桶)
    art_rows: List[Tuple[float, str, str, str, str]] = []  # (y, sku, desc, material, price_str)
    for y in sorted_ys:
        bucket = y_groups[y]
        for w in bucket:
            if w["text"] == "ARTICOLO/ARTICLE:":
                # 拼一个完整行 — 同一 y 桶里所有文本
                full = " ".join(w2["text"] for w2 in sorted(bucket, key=lambda x: x["x0"]))
                m = PANICALE_ARTICOLO_RE.search(full)
                if m:
                    sku = m.group(1).strip()
                    rest = m.group(2)
                    price_str = m.group(3)
                    # 拆 desc + material:最后一个带 % 的是 material
                    parts = [p.strip() for p in rest.split(" - ")]
                    if parts and "%" in parts[-1]:
                        material = parts[-1]
                        desc = " - ".join(parts[:-1]).strip()
                    else:
                        material = ""
                        desc = rest.strip()
                    art_rows.append((y, sku, desc, material, price_str))
                break

    # 5) 找所有 Colore/Color 行(每个块一个) — 同时记录 size 列的 x 中心位置
    #    (2026-08-18 改:size 位置从每块 Colore/Color 行动态取,不是硬编码)
    cc_rows: Dict[float, Tuple[List[str], List[float]]] = {}  # y → (sizes, x_centers)
    for y in sorted_ys:
        bucket = y_groups[y]
        for w in bucket:
            if w["text"] == "Colore/Color":
                # 找同一 y 桶里的所有数字 — 同时记 x 中心
                sizes = []
                x_centers = []
                for w2 in sorted(bucket, key=lambda x: x["x0"]):
                    if w2["text"] != "Colore/Color" and w2["text"].isdigit():
                        sizes.append(w2["text"])
                        x_centers.append((w2["x0"] + w2["x1"]) / 2)
                if sizes:
                    cc_rows[y] = (sizes, x_centers)
                break

    # 6) 解析每个 ARTICOLO 块
    rows: List[Dict[str, object]] = []
    for i, (art_y, sku, desc, material, price_str) in enumerate(art_rows):
        # 块结束 y = 下一个 ARTICOLO y 或 页底
        block_end_y = art_rows[i + 1][0] if i + 1 < len(art_rows) else page_h
        # 找 Colore/Color header(在 art_y+5 ~ art_y+25 之间,稍微低于 ARTICOLO)
        cc_y = None
        for cy in sorted(cc_rows.keys()):
            if art_y < cy < block_end_y:
                cc_y = cy
                break
        if not cc_y:
            continue
        sizes, size_x_centers = cc_rows[cc_y]
        # 如果 Colore/Color 行的 size x 没拿到(防御),用 fallback
        if not size_x_centers or len(size_x_centers) != len(sizes):
            size_x_centers = PANICALE_SIZE_X_CENTERS_FALLBACK[:len(sizes)]

        # 7) 在 cc_y 之后扫所有 y 桶,识别颜色 3 行模式
        #    模式: code_row(y) → qty_row(y+~6) → name_row(y+~10)
        #    实际上 y 可能不严格,但可以用 pattern 识别
        #    用 y 桶 y_set 排序
        block_ys = [y for y in sorted_ys if cc_y < y < block_end_y]
        # 收集 code/qty/name 序列
        i_y = 0
        while i_y < len(block_ys):
            y = block_ys[i_y]
            bucket = y_groups[y]
            # 检查是否是 color code 桶:
            # - 模式 A: 一个 word 形如 "38U240290"(全连一起)
            # - 模式 B: 两个 word "38U24" + "0290" 同 y 桶(x 紧邻, 175 和 201)
            # - 模式 C: 单 "38U24" 后面紧跟数字 word 在下一行(罕见)
            code_text = None
            # 模式 A: 整个 color code 拼成一个 word (e.g. "38U240290")
            for w in bucket:
                t = w["text"]
                # 模式 A: 2 数字 + 2+ 字母 + 4 数字 (e.g. "38U240290")
                m = re.match(r"^(\d{2})([A-Z]{2,})(\d{4,})$", t)
                if m:
                    code_text = m.group(1) + m.group(2) + " " + m.group(3)
                    break
            # 模式 B: 找 "##U##" 紧邻的纯数字 word
            #   color code pattern: \d{2}[A-Z]+(\d+)? — starts with 2 digits + 1+ letters
            if not code_text:
                code_word = None
                digit_word = None
                for w in sorted(bucket, key=lambda x: x["x0"]):
                    t = w["text"]
                    # code word: starts with 2 digits, has at least 1 letter, 4-7 chars
                    if re.match(r"^\d{2}[A-Z][A-Z0-9]{0,5}$", t) and not t.isdigit() and any(c.isalpha() for c in t):
                        # 限定 x 范围(产品在左侧)
                        if w["x0"] < 250:
                            code_word = (w["text"], w["x1"])
                    elif code_word and t.isdigit() and 3 <= len(t) <= 5:
                        # 数字 word 在 code word 紧邻右侧(差距 < 30pt)
                        if w["x0"] - code_word[1] < 30:
                            digit_word = w["text"]
                if code_word and digit_word:
                    code_text = code_word[0] + " " + digit_word
            # 模式 C: 单 "##U##" 后面下一行有数字 word
            if not code_text:
                for w in bucket:
                    t = w["text"]
                    if re.match(r"^\d{2}[A-Z]+$", t):
                        if i_y + 1 < len(block_ys):
                            next_bucket = y_groups[block_ys[i_y + 1]]
                            for w2 in next_bucket:
                                if w2["text"].isdigit() and 1 <= int(w2["text"]) <= 9999:
                                    code_text = t + " " + w2["text"]
                                    break
                        if code_text:
                            break
            if not code_text:
                i_y += 1
                continue
            # 找到了 code_text — 找 qty(同一 y 桶或紧邻 y 桶,排除 code_text 中的数字)
            qty_map: Dict[str, int] = {}  # size → qty
            # 模式 1:qty 在 code_text 同一 y 桶(size 列 x 位置上的数字,排除 x<220 的 code digit)
            for w in bucket:
                t = w["text"]
                if t.isdigit() and 1 <= int(t) <= 9999:
                    cx = (w["x0"] + w["x1"]) / 2
                    # 排除 code 的数字 word(x < 250)
                    if cx < 250:
                        continue
                    # 找最接近的 size 列(动态从 Colore/Color 行取)
                    best_size = None
                    best_dist = float("inf")
                    for sidx, sx in enumerate(size_x_centers):
                        if sidx >= len(sizes):
                            break
                        d = abs(cx - sx)
                        if d < best_dist:
                            best_dist = d
                            best_size = sizes[sidx]
                    if best_size and best_dist <= PANICALE_X_TOLERANCE:
                        qty_map[best_size] = int(t)
            # 模式 2:qty 在下一个 y 桶(当 y 桶被 col14 模式 1 错过时)
            if not qty_map and i_y + 1 < len(block_ys):
                next_y = block_ys[i_y + 1]
                next_bucket = y_groups[next_y]
                for w in next_bucket:
                    t = w["text"]
                    if t.isdigit() and 1 <= int(t) <= 9999:
                        cx = (w["x0"] + w["x1"]) / 2
                        if cx < 250:
                            continue
                        best_size = None
                        best_dist = float("inf")
                        for sidx, sx in enumerate(size_x_centers):
                            if sidx >= len(sizes):
                                break
                            d = abs(cx - sx)
                            if d < best_dist:
                                best_dist = d
                                best_size = sizes[sidx]
                        if best_size and best_dist <= PANICALE_X_TOLERANCE:
                            qty_map[best_size] = int(t)
            # 找 color name — 跳过 qty 桶(code+1 或 code+2)
            color_name = ""
            name_y = None
            for off in (2, 3):  # 跳 1 个或 2 个 y 桶
                cand = block_ys[i_y + off] if i_y + off < len(block_ys) else None
                if cand is None:
                    continue
                cand_bucket = y_groups[cand]
                # 名字桶特征:有非数字 word
                has_text = any(
                    not w["text"].isdigit() and not re.match(r"^\d{2}[A-Z]+", w["text"])
                    and w["text"] != "Colore/Color"
                    for w in cand_bucket
                )
                if has_text:
                    name_y = cand
                    color_name = " ".join(
                        w["text"] for w in sorted(cand_bucket, key=lambda x: x["x0"])
                        if not w["text"].isdigit()
                        and not re.match(r"^\d{2}[A-Z]+", w["text"])
                        and w["text"] != "Colore/Color"
                    )
                    break
            # 构造 row
            block_images = _extract_panicale_block_images(
                page_images, art_y, block_end_y,
            )
            for size, qty in qty_map.items():
                full_sku = f"{sku}-{code_text}"
                rows.append({
                    "brand": "PANICALE",
                    "sku": full_sku,
                    "image": "",
                    "description": desc,
                    "category": "",
                    "material": material,
                    "color": color_name,
                    "color_code": code_text,
                    "gender": "",
                    "season": season,
                    "origin": "",
                    "cost": float(price_str.replace(",", ".")),
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": None,
                    "size": size,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "image_items": block_images,
                    "source_row_height": None,
                })
            # 跳到 name 桶之后
            if name_y:
                try:
                    i_y = block_ys.index(name_y) + 1
                    continue
                except ValueError:
                    pass
            i_y += 1
    return rows


def _extract_panicale_block_images(
    page_images: Dict[int, Dict], art_y: float, block_end_y: float,
) -> List[Dict]:
    """提取 PANICALE 块对应的图 — 2026-08-18 加大小过滤。

    PDF 里有两种图:
    - 大产品图 (model) — 73x110pt 在 x=20-93
    - 小色卡 (color swatch) — 18x18pt 在 x=115-133 (每个色一行)
    之前 cx<250 把色卡也当产品图,现在加宽高过滤:宽度>=50pt 才是产品图
    """
    block_imgs: List[Dict] = []
    for xref, info in page_images.items():
        for rect in info["rects"]:
            cy = (rect.y0 + rect.y1) / 2
            if not (art_y - 30 <= cy <= block_end_y + 5):
                continue
            cx = (rect.x0 + rect.x1) / 2
            if cx >= 250:
                continue
            w = rect.x1 - rect.x0
            h = rect.y1 - rect.y0
            if w < 50 or h < 50:
                continue
            block_imgs.append({
                "bytes": info["bytes"],
                "format": info["format"],
            })
            break
    return block_imgs

def _parse_panicale_pdf(
    primary_file: Path, vendor_id: str,
) -> List[Dict[str, object]]:
    """PANICALE 现货单 PDF 解析主入口。"""
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    # 季节从文件名提取
    season = ""
    name_upper = primary_file.stem.upper()
    m = re.search(r"(SS\d{2}|FW\d{2})", name_upper)
    if not m:
        m = re.search(r"\b(SS|FW)\b", name_upper)
    if m:
        season = m.group(1)
    if not season:
        season = "SS26"

    doc = fitz.open(primary_file)
    try:
        with pdfplumber.open(primary_file) as pdf:
            for pno, page in enumerate(pdf.pages):
                if pno < len(doc):
                    fitz_page = doc[pno]
                else:
                    fitz_page = None
                page_rows = _parse_panicale_page(page, fitz_page, doc, season, vendor_id)
                rows.extend(page_rows)
    finally:
        doc.close()
    return rows


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """PDF vendor 入口 — 自动检测 PANICALE / MONCLER 版式。"""
    if not primary_file.exists():
        return []
    # PANICALE 版式检测
    if _detect_panicale_layout(primary_file):
        return _parse_panicale_pdf(primary_file, vendor_id)
    # 否则走原 MONCLER 版式(保留)
    return _parse_moncler_pdf(primary_file, vendor_id)
