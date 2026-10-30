"""V1007 vendor adapter — GANNI 现货单 (PDF)

来源:GANNI Wholesale 货单 .pdf (Style 7 系列)
- 25.4 / 26.1 / 26.2 / 26.3 / WHOLESALE_CARRYOVER 等子品牌
- 仓库 ID 固定为 V1007
- 品牌默认 GANNI(CLI --brand 可覆盖)
- 季节从文件名提取(SS26 / FW26 / SS25 / ...)

结构(每页 3-4 个产品块,每块 4-6 行):
  R(N): SKU 在左 + Wholesale EUR <cost> 在右
  R(N+1): 描述在左 + MSRP EUR <retail> 在右
  R(N+2): size header (XXS/XS S/M L/XL 2XL/3XL TOTAL UNITS 或 32 34 36 38 40 42 44)
  R(N+3..N+X): color row(s) — 颜色名 + 色码 + 各 size qty + TOTAL UNITS
  可选 R(N+X+1): extended size header (3XL 4XL) + qty row (0 0)
  可选 R(N+X+2): total row (跨 size 汇总)

每块有 1 张产品图(大,105x144pt @ x=32-138)+ 多个色卡小图(7x7pt,不重要,丢弃)

SKU 识别:col 1 (x=148-200) 形如 ^[A-Z]+\d+$
size header 识别:R(N+2) 是纯数字或 [XXS XS S M L XL 2XL] 等尺码标签
qty 值:与 size header 同 x 位置对齐(R(N+3) 行),y 略偏(±2pt 同一组)
color name:左半区(x<200) 非数字文本(可能 truncated '...')
color code:左半区第一个 3 位数
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import pdfplumber

VENDOR_ID = "V1007"
DEFAULT_BRAND = "GANNI"

# 2026-08-26: 注册 openpyxl 缺的 MIME type,修 PyMuPDF 返回 jp2/jpeg2000 时
# openpyxl 写文件报 KeyError: '.jpeg2000'
for ext in (".jp2", ".jpeg2000", ".jpx", ".j2k", ".jpf"):
    import mimetypes
    try:
        mimetypes.types_map[True][ext] = "image/jp2"
    except (KeyError, TypeError):
        try:
            mimetypes._db
        except Exception:
            pass
        mimetypes.add_type("image/jp2", ext, strict=True)
        try:
            mimetypes.types_map[True] = mimetypes._db.types_map[1]
        except Exception:
            pass


# SKU 识别 — 字母 + 数字组合(如 W0565, A1010099, S3022)
SKU_RE = re.compile(r"^[A-Z]+\d+$")
# 价格识别 "Wholesale EUR 70.00" / "MSRP EUR 195.00"
WHOLESALE_RE = re.compile(r"Wholesale\s+EUR\s+([\d,]+(?:\.\d+)?)", re.IGNORECASE)
MSRP_RE = re.compile(r"MSRP\s+EUR\s+([\d,]+(?:\.\d+)?)", re.IGNORECASE)


def _is_size_header_word(text: str) -> bool:
    """判断 R(N+2) size header 行的 token。

    接受:
    - 纯数字 ("32", "44", "52", "2½")
    - 单/双字母尺码 (XXS, XS, S, M, L, XL, XXL, 2XL, 3XL, 4XL, ...)
    - 组合尺码 (XXS/XS, S/M, L/XL, 2XL/3XL)
    """
    if not text:
        return False
    t = text.strip()
    if not t:
        return False
    if t in ("TOTAL", "UNITS"):
        return True
    # 纯数字
    if re.fullmatch(r"\d+", t):
        return True
    # Unicode 分数 (2½)
    if re.fullmatch(r"\d+[½¼¾]", t):
        return True
    # 标准尺码标签
    if re.fullmatch(r"(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|\dXL)", t):
        return True
    # 组合尺码
    if re.fullmatch(r"[A-Z\d]+/[A-Z\d]+", t):
        return True
    return False


def _to_int_qty(s: str) -> int:
    """把 qty 字符串转 int(支持 '2', '2.0', 失败返 0)。"""
    if not s:
        return 0
    try:
        return int(round(float(str(s).strip())))
    except (ValueError, TypeError):
        return 0


def _parse_eur_price(s: str) -> Optional[float]:
    """意大利/欧洲价格格式: '70.00' / '1.400,00' / '1.500,00' / '70,00'。

    - 单 . 且 . 在中间 → 千位分隔符(欧洲写法),如 1.400 → 1400
    - 单 , 且 , 在末尾 2 位 → 小数,如 70,00 → 70.00
    - 单 . 在末尾 2 位 → 小数,如 70.00 → 70.00
    - 1.400,00 → 1400.00
    """
    if not s:
        return None
    t = str(s).strip()
    if not t:
        return None
    # 1.400,00 格式: 有点有逗 → 点千位 + 逗小数
    if "." in t and "," in t:
        return float(t.replace(".", "").replace(",", "."))
    # 只有逗号: 是千位还是小数?看位置
    if "," in t:
        if len(t) - t.rfind(",") == 3:  # 逗号后正好 2 位 → 小数
            return float(t.replace(",", "."))
        else:  # 千位分隔
            return float(t.replace(",", ""))
    # 只有点: 跟欧洲反(US/UK 写法), 是小数
    if "." in t:
        return float(t)
    # 纯数字
    try:
        return float(t)
    except (ValueError, TypeError):
        return None


def _words_near_y(y_groups: Dict[float, List[Dict]], target_y: float, tolerance: float = 2.5) -> List[Dict]:
    """返回 y 桶中 top 与 target_y 差值 <= tolerance 的所有 words(按 x 排序)。"""
    out: List[Dict] = []
    for y, words in y_groups.items():
        if abs(y - target_y) <= tolerance:
            out.extend(words)
    out.sort(key=lambda w: w["x0"])
    return out


def _extract_size_x_centers(size_words: List[Dict]) -> List[Tuple[str, float]]:
    """从 size header 行的 words 提取 (size_label, x_center) 列表,排除 'TOTAL UNITS'。"""
    out: List[Tuple[str, float]] = []
    for w in size_words:
        text = w["text"].strip()
        if text in ("TOTAL", "UNITS"):
            continue
        if not _is_size_header_word(text):
            continue
        cx = (w["x0"] + w["x1"]) / 2
        out.append((text, cx))
    return out


def _find_qty_x_centers(qty_words: List[Dict]) -> List[Tuple[int, float]]:
    """从 qty 行的 words 提取 (qty, x_center) 列表。"""
    out: List[Tuple[int, float]] = []
    for w in qty_words:
        t = w["text"].strip()
        if not t:
            continue
        try:
            v = int(round(float(t)))
            if 0 <= v < 100000:
                cx = (w["x0"] + w["x1"]) / 2
                out.append((v, cx))
        except (ValueError, TypeError):
            continue
    return out


def _map_qty_to_sizes(
    size_x: List[Tuple[str, float]],
    qty_x: List[Tuple[int, float]],
) -> List[Tuple[str, int]]:
    """把 qty 按 x 位置映射到 size label。返回 [(size, qty), ...]。"""
    if not size_x:
        return []
    # 找 TOTAL UNITS x(最后两个 qty 中的一个)
    # 实际上 size label 的 x 位置和 qty 的 x 位置应该一致或非常接近
    out: List[Tuple[str, int]] = []
    for size, sx in size_x:
        # 找最接近的 qty
        best_qty = 0
        best_dist = float("inf")
        for q, qx in qty_x:
            d = abs(sx - qx)
            if d < best_dist:
                best_dist = d
                best_qty = q
        if best_dist <= 8:  # ±8pt 容差
            out.append((size, best_qty))
    return out


def _extract_block_images(
    fitz_page, art_y: float, block_end_y: float,
) -> List[Dict]:
    """提取该块的图(产品图,大,过滤小色卡)。"""
    block_imgs: List[Dict] = []
    for img in fitz_page.get_images(full=True):
        xref = img[0]
        try:
            rects = fitz_page.get_image_rects(xref)
        except Exception:
            rects = []
        for rect in rects:
            cy = (rect.y0 + rect.y1) / 2
            if not (art_y - 5 <= cy <= block_end_y + 5):
                continue
            # 限定 x 范围(产品在左侧)
            cx = (rect.x0 + rect.x1) / 2
            if cx >= 250:
                continue
            # 尺寸过滤(产品图至少 50x50pt,排除 7x7 色卡)
            w = rect.x1 - rect.x0
            h = rect.y1 - rect.y0
            if w < 50 or h < 50:
                continue
            try:
                base = fitz_page.parent.extract_image(xref)
                data = base.get("image")
                ext = base.get("ext", "png")
                if data:
                    block_imgs.append({
                        "bytes": data,
                        "format": ext.lower() if ext.lower() in ("png", "jpeg", "jpg") else "png",
                    })
            except Exception:
                pass
            break  # 每图只取一次
    return block_imgs


def _parse_block(
    page_words: List[Dict],  # 这页所有 words,按 y 分桶前
    y_groups: Dict[float, List[Dict]],  # y → words
    sorted_ys: List[float],  # 所有有 word 的 y 排序
    art_y: float,  # 块起始 y (SKU 行)
    block_end_y: float,  # 块结束 y (下一个块开始或页底)
    fitz_page,
) -> List[Dict[str, object]]:
    """解析一个产品块 → 输出 1 个或多个 row(每个 color 1 个 row)。"""
    rows: List[Dict[str, object]] = []

    # 1) SKU 在 art_y 行 x=148-200 区间
    sku = None
    for w in y_groups.get(art_y, []):
        if SKU_RE.match(w["text"].strip()):
            sku = w["text"].strip()
            break
    if not sku:
        return rows

    # 2) 描述在 art_y+10 左右(单行多 word,x=148-350)
    desc = ""
    desc_y = art_y + 10
    desc_words = _words_near_y(y_groups, desc_y)
    if desc_words:
        desc = " ".join(
            w["text"].strip() for w in desc_words
            if w["x0"] < 400 and not WHOLESALE_RE.search(w["text"])
            and not MSRP_RE.search(w["text"])
        )

    # 3) Wholesale / MSRP 价格(在右侧 y=art_y+3 ~ art_y+13)
    # 2026-08-26: pdfplumber 把 "Wholesale EUR 158.00" 拆成 3 个 word,要 join 起来匹配
    # Wholesale 和 MSRP 在不同 y,要分两次找 — 不能 break 早了
    # 欧洲价格格式 "1.400,00" (点千位 + 逗小数),用 _parse_eur_price 解析
    wholesale = None
    msrp = None
    for target_y in (art_y + 3, art_y + 10, art_y + 13, art_y + 15):
        right_words = [w for w in _words_near_y(y_groups, target_y) if w["x0"] > 400]
        right_words.sort(key=lambda w: w["x0"])
        right_text = " ".join(w["text"] for w in right_words)
        m = WHOLESALE_RE.search(right_text)
        if m and wholesale is None:
            wholesale = _parse_eur_price(m.group(1))
        m = MSRP_RE.search(right_text)
        if m and msrp is None:
            msrp = _parse_eur_price(m.group(1))
        if wholesale is not None and msrp is not None:
            break

    # 4) 找 size header 行(在 desc 之后,有 >=2 个 size token)
    # 通常在 desc_y+20 左右
    size_header_y = None
    for cand_y in sorted_ys:
        if cand_y <= desc_y + 15:
            continue
        if cand_y >= block_end_y:
            break
        words_at_y = y_groups.get(cand_y, [])
        # 至少 2 个 size token
        size_tokens = [w for w in words_at_y if _is_size_header_word(w["text"])]
        if len(size_tokens) >= 2 and any(t in ("TOTAL", "UNITS") for t in [w["text"] for w in words_at_y]):
            size_header_y = cand_y
            break
        # 备选:有 2+ size token 且至少 1 个 x<350
        if len(size_tokens) >= 2:
            # 检查是否真有多个 size(非仅 TOTAL)
            non_total = [t for t in size_tokens if t["text"] not in ("TOTAL", "UNITS")]
            if len(non_total) >= 2:
                size_header_y = cand_y
                break
    if not size_header_y:
        return rows

    size_x = _extract_size_x_centers(y_groups[size_header_y])

    # 5) 找 color rows(size header 下方 1-2 行,每行有 qty numbers)
    # 一个块可能有多 color (多个 SKU 共享 size header)
    # 也可能有 extended size header + extra qty row
    color_rows: List[Tuple[float, List[Dict]]] = []  # [(y, qty_words), ...]
    current_y = size_header_y + 5  # 起始找 color row
    while current_y < block_end_y - 5:
        # 找下一个有效 y
        next_y = None
        for cand_y in sorted_ys:
            if cand_y <= current_y:
                continue
            if cand_y >= block_end_y:
                break
            words_at_y = y_groups.get(cand_y, [])
            qty_words = _find_qty_x_centers(words_at_y)
            # 至少 1 个 qty,且包含 >= 2 个数字(否则是 size header 之类的)
            if len(qty_words) >= 2:
                next_y = cand_y
                color_rows.append((cand_y, words_at_y))
                current_y = cand_y + 5
                break
            elif len(qty_words) == 1 and any(
                w["text"] in ("TOTAL", "UNITS") for w in words_at_y
            ):
                # total row,跳过
                current_y = cand_y + 5
                break
            elif len(qty_words) == 0:
                # 检查是否是 size header (2XL/3XL 等) — 跳
                size_tokens = [w for w in words_at_y if _is_size_header_word(w["text"])]
                if len(size_tokens) >= 2:
                    current_y = cand_y + 5
                    continue
                # 空行或纯文本 — 终止
                break
            else:
                # 1 个数字(可能是 total 但不在 TOTAL UNITS 列)— 也算 color row
                next_y = cand_y
                color_rows.append((cand_y, words_at_y))
                current_y = cand_y + 5
                break
        if next_y is None:
            break

    if not color_rows:
        return rows

    # 6) 对每个 color row 解析 color + qty
    for cy, qty_words_at_y in color_rows:
        qty_x = _find_qty_x_centers(qty_words_at_y)
        # 取色名/色码(左半区,x<200)
        # 2026-08-26:GANNI PDF 实际结构是 y(qty) 在上, y+1~2(色名) 在下
        # dedup 用 (x, text) 键 — 避免 offset 重叠时重复收集
        color_name_parts: List[str] = []
        color_code = ""
        seen: set = set()
        # 2026-08-26 修:色码可能落在 x=200~225 区间(如 "Phantom 252" 中 252 在 x=202)
        # 但 qty 值至少在 x=240+,所以放宽到 x<230 既能收色码也能挡 qty
        for offset in (0, 1, 2):
            cand_y = cy + offset
            for w in _words_near_y(y_groups, cand_y):
                if w["x0"] >= 230:
                    continue
                t = w["text"].strip()
                if not t:
                    continue
                key = (round(w["x0"], 1), t)
                if key in seen:
                    continue
                seen.add(key)
                # 数字(可能是色码,3 位数)
                if re.fullmatch(r"\d+", t) and len(t) == 3 and not color_code:
                    color_code = t
                elif t.isdigit():
                    # 数字但不是 3 位 — 跳过(可能是 qty 跑到左半区)
                    continue
                else:
                    color_name_parts.append(t)
        color_name = " ".join(color_name_parts).strip()

        # 映射 qty 到 size
        size_qty_map = _map_qty_to_sizes(size_x, qty_x)

        if not size_qty_map:
            continue

        # 7) 找 product image(块 y 范围内,左侧)
        block_images = _extract_block_images(
            fitz_page, art_y - 5, block_end_y + 5,
        )

        # 8) 每个 size 出 1 个 row
        for size, qty in size_qty_map:
            if qty <= 0:
                continue
            full_sku = f"{sku}-{color_code}" if color_code else sku
            rows.append({
                "brand": DEFAULT_BRAND,
                "sku": full_sku,
                "image": "",
                "description": desc,
                "category": "",
                "material": "",
                "color": color_name,
                "color_code": color_code,
                "gender": "",
                "season": "",  # dispatcher 走文件名兜底
                "origin": "",
                "cost": wholesale,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": msrp,
                "size": size,
                "quantity": qty,
                "vendor_id": VENDOR_ID,
                "image_items": block_images,
                "source_row_height": None,
            })
    return rows


def parse_rows(
    primary_file: Path,
    _: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1007 (GANNI) 货单 PDF。

    强制把 vendor_id 改成 V1007,保证输出 仓库ID 列稳定是 V1007。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    if primary_file.suffix.lower() != ".pdf":
        return rows

    # 季节从文件名提取
    season = ""
    name_upper = primary_file.stem.upper()
    m = re.search(r"(SS\d{2}|FW\d{2})", name_upper)
    if m:
        season = m.group(1)

    doc = fitz.open(primary_file)
    try:
        with pdfplumber.open(primary_file) as pdf:
            for pno, page in enumerate(pdf.pages):
                words = page.extract_words() or []
                # 按 y 分桶(2pt 容差)
                y_groups: Dict[float, List[Dict]] = {}
                for w in words:
                    y = round(w["top"])
                    y_groups.setdefault(y, []).append(w)
                sorted_ys = sorted(y_groups.keys())

                # 找所有 SKU 行(art_y 列表)
                art_ys: List[float] = []
                for y in sorted_ys:
                    for w in y_groups[y]:
                        if SKU_RE.match(w["text"].strip()) and w["x0"] < 200:
                            # 避免重复加同一 y
                            if y not in art_ys:
                                art_ys.append(y)
                            break

                # 解析每个块
                for i, art_y in enumerate(art_ys):
                    block_end_y = art_ys[i + 1] - 1 if i + 1 < len(art_ys) else page.height
                    fitz_page = doc[pno]
                    block_rows = _parse_block(
                        words, y_groups, sorted_ys,
                        art_y, block_end_y, fitz_page,
                    )
                    for r in block_rows:
                        if season:
                            r["season"] = season
                    rows.extend(block_rows)
    finally:
        doc.close()
    return rows


def get_vendor_meta() -> Dict[str, object]:
    return {
        "default_brand": DEFAULT_BRAND,
    }
