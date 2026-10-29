"""V1010 / V1038 共用的解析逻辑。

适用文件版式(.xls / .xlsx):
- 货号在 col 14 (N)
- 货号下 1 行是描述(含 2 个 XX,XX 格式的价格 = cost + retail)
- 货号下 3 行是尺码行(col 13+)
- 货号下 5 行是数量行(col 17+)
- .xlsx 时 图片在 col 5 (F) 附近,行号 ≈ 货号行 - 1

观察到的真实样本:
- V1010 Longchamp (.xlsx): 36 个货号,描述如 "made in china, leather Yellow 37,50 90,00"
- V1038 APC       (.xlsx): 24 个货号,描述如 "T-SHIRTS White 44,00 110,00 TURKEY 100% CO"

本模块只负责把文件解析成 List[Dict];品牌默认/仓库 ID 走各 vendor adapter 自己
声明(get_vendor_meta + VENDOR_ID)。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

# 不当作货号的"尺码标签"白名单(出现这些说明这行是尺码行,不是货号行)
SIZE_LABELS = {
    "UNI", "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
    "TU", "TU.", "ONE", "ONE SIZE", "UNICA",
}

# 价格正则:意大利/欧洲格式 `XX,XX` 或 `XXX,XX` 或 `XXXX,XX`(允许空格)
PRICE_RE = re.compile(r"\d{1,5}\s*,\s*\d{2}")

# 材质百分比正则:`100% Cotton` / `5% Elastane` 等
PCT_RE = re.compile(r"\d+\s*%")


def _norm(val) -> str:
    if val is None:
        return ""
    if not isinstance(val, str):
        val = str(val)
    return val.strip()


def _looks_like_desc(s: str) -> bool:
    """判断字符串是不是描述行内容(用于反向定位货号行)。

    不依赖货号格式,只检查"描述特征":
      - 产地短语:  made in / madei (OCR 错字容错)
      - 材质百分比: 数字+%(如 100% Cotton)
      - 价格模式:  数字.数字 数字.数字(欧式 / 美式价格,37,50 90,00 或 100.50 250.00)
    """
    if not s:
        return False
    s_low = s.lower()
    # 强信号 1: 产地短语
    if 'made in' in s_low or 'madei' in s_low:
        return True
    # 强信号 2: 材质百分比
    if PCT_RE.search(s):
        return True
    # 强信号 3: 价格(欧式 XX,XX / 美式 XX.XX)
    if re.search(r'\d+[.,]\d{2}\s+\d+[.,]\d{2}', s):
        return True
    return False


def find_sku_blocks(df: pd.DataFrame) -> List[Dict]:
    """扫整张表,找到所有"货号块"。

    块结构(col 14):
      R0: 货号
      R1: 描述(含 2 个价格)
      R2: (空)
      R3: 尺码标签
      R4: (空)
      R5: 数量(部分列)

    识别策略: 不匹配货号格式, 而是识别"描述行"(用 _looks_like_desc),
    描述行**上一行非空** = 货号行。这样任何品牌的货号格式都通用。
    """
    blocks: List[Dict] = []
    if df.shape[1] < 14:
        return blocks
    n = len(df)
    for r in range(n):
        v = df.iloc[r, 13]
        if pd.isna(v):
            continue
        desc = _norm(v)
        if not _looks_like_desc(desc):
            continue
        # 找上一行非空、且**不是描述**的行作为货号
        for sr in range(r - 1, -1, -1):
            sku_raw = df.iloc[sr, 13]
            if pd.isna(sku_raw):
                continue
            sku = _norm(sku_raw)
            if not sku:
                continue
            if _looks_like_desc(sku):
                # 上一行也是描述(不连续), 跳出找下一个块
                break
            blocks.append({
                "sku_row": sr,
                "desc_row": r,
                "sku": sku,
                "desc_text": desc,
            })
            break
    return blocks


def find_size_qty(df: pd.DataFrame, desc_row: int) -> Tuple[Dict[int, str], Dict[int, int]]:
    """在描述行下方找尺码行和数量行。

    经验(对 V1010 / V1038 都成立):
      - 尺码行:col 13+ 是非空字符串(尺码标签)
      - 数量行:尺码行下方 1-2 行,col 13+ 出现数字
    """
    size_map: Dict[int, str] = {}
    qty_map: Dict[int, int] = {}
    if df.shape[1] < 14:
        return size_map, qty_map

    def _is_blank(v) -> bool:
        """pandas NaN / None / 空字符串 统一视为空。"""
        if v is None:
            return True
        if isinstance(v, float) and pd.isna(v):
            return True
        return False

    # 找尺码行(描述行下方 1-5 行内)
    # 两种版式:
    #   模式 A (包/配饰, 单尺码): col 13 是字符串尺码标签 (UNI / XXXS / ...)
    #                            col 14+ 是 qty (单尺码)
    #   模式 B (鞋, 多尺码):   col 13 是数字尺码值 (65 / 70 / 75 ...)
    #                            col 14+ 也是尺码值; qty 在尺码行下方 1-3 行同 col
    # 2026-08-26: 加 fraction 支持 — TORY BURCH 等品牌用 "2½" "6½" "10½" 这种
    # Unicode ½ 字符表示半号鞋码,原来 mode B 只认纯数字会全跳过
    def _is_digit_or_fraction(s: str) -> bool:
        """数字 + Unicode 分数字符也算 mode B size: '2', '2.5', '2½', '10' 都算。"""
        if not s:
            return False
        # 替换常见 Unicode 分数字符为小数点
        normalized = s.replace("½", ".5").replace("¼", ".25").replace("¾", ".75")
        try:
            float(normalized)
            return True
        except (ValueError, TypeError):
            return False

    size_row_idx: Optional[int] = None
    for off in range(1, 6):
        r = desc_row + off
        if r >= len(df):
            break
        v13 = df.iloc[r, 13] if df.shape[1] > 13 else None
        s13 = _norm(v13)
        if _is_blank(v13) or not s13:
            continue  # 真空行 — 跳过

        is_label = s13 in SIZE_LABELS or (len(s13) <= 6 and not s13.isdigit())
        # 模式 B: col 13 是数字 (鞋码)— 含 ½ ¼ ¾ 等分数字符也算
        is_digit_size = _is_digit_or_fraction(s13) and 1 <= len(s13) <= 5

        if is_label or is_digit_size:
            size_row_idx = r
            for c in range(13, min(30, df.shape[1])):
                v = df.iloc[r, c]
                if _is_blank(v):
                    continue
                s = _norm(v)
                if not s or s == "nan":
                    continue
                if is_label:
                    # 模式 A: 任何非空都收 (尺码标签或 qty 数字)
                    size_map[c] = s
                else:
                    # 模式 B: 收数字或分数(2026-08-26:含 ½ ¼ ¾)
                    if _is_digit_or_fraction(s):
                        # 2026-08-26: pandas 把整数读成 6.0,转成 "6" 跟 "6½" 风格一致
                        if isinstance(v, float) and v.is_integer():
                            s = str(int(v))
                        else:
                            s = str(v).strip()
                        size_map[c] = s
            break

    if size_row_idx is None:
        return size_map, qty_map

    # 找数量行(尺码行下方 1-3 行,col 13+ 出现数字)
    for qoff in range(1, 4):
        qr = size_row_idx + qoff
        if qr >= len(df):
            break
        for c, label in size_map.items():
            v = df.iloc[qr, c] if c < df.shape[1] else None
            if _is_blank(v):
                continue
            try:
                qty = int(float(str(v).strip()))
                if qty > 0:
                    qty_map[c] = qty
            except (ValueError, TypeError):
                continue
        if qty_map:
            break

    return size_map, qty_map


def parse_description(desc: str) -> Tuple[str, Optional[float], Optional[float], str, str, str]:
    """从描述里拆出 (description, cost, retail, origin, material, color)。

    真实样本(描述末尾总是跟 2 个 XX,XX 价格):
      "made in china, leather Yellow 37,50 90,00"
        → ("made in china, leather Yellow", 37.5, 90.0, "CHINA", "leather", "YELLOW")
      "Made in France 100% Cotton Trimmed Cowhide Leather Brown 58,00 140,00"
        → ("Made in France 100% Cotton Trimmed Cowhide Leather Brown", 58.0, 140.0, "FRANCE",
           "100% Cotton Trimmed Cowhide Leather", "BROWN")
      "made in China, 100% Cotton, Trimmed with Cowhide Leather Green 104,00 250,00"
        → ("made in China, 100% Cotton, Trimmed with Cowhide Leather Green", 104.0, 250.0, "CHINA",
           "100% Cotton, Trimmed with Cowhide Leather", "GREEN")
      "JEANS Washed Indigo 108,00 270,00 EGYPT 100% CO"  (APC 版式,无 made in 开头)
        → ("JEANS Washed Indigo EGYPT 100% CO", 108.0, 270.0, "", "JEANS Washed Indigo", "EGYPT")

    拆分规则:
      1. 价格:最后 2 个 XX,XX = cost、retail
      2. 产地:开头 `made in <国家>`(大小写不敏感),没匹配就空
      3. 颜色:剩余文本的最后一个"单词"(无空格、无标点),如不在已知颜色列表里
              也照常抽取(避免漏掉)
      4. 材质:扣除产地前缀和颜色后剩下的内容
    """
    if not desc:
        return "", None, None, "", "", ""

    # 1) 抠出价格
    prices = list(PRICE_RE.finditer(desc))
    if len(prices) >= 2:
        cost_str = prices[-2].group().replace(" ", "").replace(",", ".")
        retail_str = prices[-1].group().replace(" ", "").replace(",", ".")
        try:
            cost = float(cost_str)
            retail = float(retail_str)
        except ValueError:
            return desc, None, None, "", "", ""
        text = (desc[: prices[-2].start()].rstrip() + " " + desc[prices[-1].end():])
        text = re.sub(r"\s+", " ", text).strip()
    else:
        cost = retail = None
        text = desc.strip()

    # 2) 产地:made in <国家>(可能后接逗号)
    origin = ""
    m = re.match(r"^made\s+in\s+([A-Za-z]+)[,.\s]*", text, flags=re.IGNORECASE)
    if m:
        origin = m.group(1).upper()
        text = text[m.end():].lstrip(" ,")

    # 3) 颜色:最后一个"单词"(只含字母)
    #    用 findall 拿到所有"单词",最后一个字母-only 的就是颜色候选
    color = ""
    tokens = text.split()
    if tokens:
        # 从后往前找第一个"纯字母单词"(忽略标点/数字)
        for tok in reversed(tokens):
            cleaned = re.sub(r"[^A-Za-z]", "", tok)
            if cleaned and len(cleaned) >= 2:
                color = cleaned.upper()
                # 颜色 token 也从 text 里抠掉
                text = (text[: text.rfind(tok)] + text[text.rfind(tok) + len(tok):]).strip()
                text = re.sub(r"\s+", " ", text).strip(" ,")
                break

    # 4) 材质:剩下的就是材质
    material = text.strip(" ,")

    # description 保留原始清理后的全文(让用户在「描述」列看到原话)
    if len(prices) >= 2:
        full = (desc[: prices[-2].start()].rstrip() + " " + desc[prices[-1].end():])
        full = re.sub(r"\s+", " ", full).strip()
    else:
        full = desc.strip()
    return full, cost, retail, origin, material, color


# V1038 专用:版式 `<类型> <颜色> <cost> <retail> <产地> <材料>`
# 类型在头部,颜色在中间,产地在价格之后,材料在末尾
V1038_KNOWN_TYPES = [
    # 多字类型优先(从长到短匹配)
    "TOP HANDLE BAGS",
    "DOWN JACKETS",
    "SHOULDER BAGS",
    "DENIM JACKETS",
    "T-SHIRTS",
    "KNITWEAR",
    "JEANS",
    "JACKETS",
    "SHIRTS",
    "TOTES",
    "JACKET",
    "SHIRT",
    "BAGS",
    "BAG",
]

# V1038 产地白名单(已知国家/产地,大小写归一化)
# 注:这些是 APC 表里出现过的,源数据里"Cina"=意大利语 China,会归一化到 CHINA
V1038_ORIGIN_MAP = {
    "EGYPT": "EGYPT",
    "TURKEY": "TURKEY",
    "PORTUGAL": "PORTUGAL",
    "ROMANIA": "ROMANIA",
    "ITALIA": "ITALY",
    "ITALY": "ITALY",
    "CINA": "CHINA",
    "CHINA": "CHINA",
    "MAROCCO": "MOROCCO",
    "MOROCCO": "MOROCCO",
    "VIETNAM": "VIETNAM",
    "SPAGNA": "SPAIN",
    "SPAIN": "SPAIN",
    "FRANCE": "FRANCE",
    "INDIA": "INDIA",
    "TUNISIA": "TUNISIA",
}

# 材料正则:以百分号/数字开头(允许逗号分隔多个材料)
MATERIAL_RE = re.compile(r"^[\d\s,%]+%?\s*[A-Za-z]+(?:\s*,\s*[\d\s,%]+%?\s*[A-Za-z]+)*$")


def parse_description_v1038(desc: str) -> Tuple[str, Optional[float], Optional[float], str, str, str, str]:
    """V1038 (APC) 描述解析器。

    真实样本(描述 = `<类型> <颜色> <cost> <retail> <产地> <材料>`):
      "JEANS Washed Indigo 108,00 270,00 EGYPT 100% CO"
        → ("JEANS Washed Indigo", 108.0, 270.0, "EGYPT", "JEANS", "Washed Indigo", "100% CO")
      "T-SHIRTS DARK NAVY/ECRU 44,00 110,00 TURKEY 100% CO"
        → ("T-SHIRTS DARK NAVY/ECRU", 44.0, 110.0, "TURKEY", "T-SHIRTS", "DARK NAVY/ECRU", "100% CO")
      "TOP HANDLE BAGS Indigo 132,00 330,00 Cina 100% CO"
        → ("TOP HANDLE BAGS Indigo", 132.0, 330.0, "CHINA", "TOP HANDLE BAGS", "Indigo", "100% CO")
      "SHOULDER BAGS HAZELNUT 212,00 530,00 Marocco 100% LE"
        → ("SHOULDER BAGS HAZELNUT", 212.0, 530.0, "MOROCCO", "SHOULDER BAGS", "HAZELNUT", "100% LE")
      "DOWN JACKETS Dark navy 236,00 590,00 ROMANIA 55% PL, 45% WV"
        → ("DOWN JACKETS Dark navy", 236.0, 590.0, "ROMANIA", "DOWN JACKETS", "Dark navy", "55% PL, 45% WV")
      "TOP HANDLE BAGS STEEL BLUE 204,00 510,00 Marocco 100% LE"
        → ("TOP HANDLE BAGS STEEL BLUE", 204.0, 510.0, "MOROCCO", "TOP HANDLE BAGS", "STEEL BLUE", "100% LE")
      "SHIRTS White 108,00 270,00 SPAGNA 100% CO"
        → ("SHIRTS White", 108.0, 270.0, "SPAIN", "SHIRTS", "White", "100% CO")

    返回: (description, cost, retail, origin, category, color, material)
      - description: 类型+颜色原文
      - cost/retail: 数字
      - origin: 标准化后的大写国家(没匹配到白名单就保留原大写)
      - category: 已知类型(从 V1038_KNOWN_TYPES 匹配,最长优先;没匹配用第一个词)
      - color: 类型后到价格前的内容
      - material: 末尾"百分号开头"的内容
    """
    if not desc:
        return "", None, None, "", "", "", ""

    # 1) 抠出价格
    prices = list(PRICE_RE.finditer(desc))
    if len(prices) >= 2:
        cost_str = prices[-2].group().replace(" ", "").replace(",", ".")
        retail_str = prices[-1].group().replace(" ", "").replace(",", ".")
        try:
            cost = float(cost_str)
            retail = float(retail_str)
        except ValueError:
            return desc, None, None, "", "", "", ""
        before = desc[: prices[-2].start()].rstrip()
        after = desc[prices[-1].end():].strip()
    else:
        cost = retail = None
        before = desc.strip()
        after = ""

    # 2) 拆 after:第一段是产地,剩下是材料(材料可能多段)
    after_tokens = after.split()
    origin = ""
    material = ""
    if after_tokens:
        # 第一个 token 是产地
        raw_origin = after_tokens[0]
        origin_key = raw_origin.upper()
        origin = V1038_ORIGIN_MAP.get(origin_key, origin_key)
        # 剩下的拼成材料
        material = " ".join(after_tokens[1:]).strip()

    # 3) 拆 before:已知类型匹配(最长优先)
    before_upper = before.upper()
    category = ""
    for kw in V1038_KNOWN_TYPES:
        if before_upper.startswith(kw + " ") or before_upper == kw:
            category = kw
            # 把类型从前 before 抠掉,剩下的就是颜色
            color = before[len(kw):].strip()
            break
    else:
        # 没匹配到 — 兜底:第一个词当类型
        parts = before.split(maxsplit=1)
        category = parts[0] if parts else ""
        color = parts[1] if len(parts) > 1 else ""

    # 4) description = 类型 + 颜色
    description = before.strip()

    return description, cost, retail, origin, category, color, material


def _load_images_by_row(primary_file) -> Dict[int, List[Dict]]:
    """从 .xlsx 加载所有图片,按行号索引。

    返回: {row_0indexed -> [{bytes, format}]}
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


def _find_image_for_block(
    images_by_row: Dict[int, List[Dict]],
    sku_row_0idx: int,
) -> Optional[Dict]:
    """找最接近 SKU 行的图片(±3 行内,优先同列偏移 = -1)。

    V1010 / V1038 的图片在 SKU 行上方 1 行(col 5)。±3 行是为了容忍版式偏差。
    """
    if not images_by_row:
        return None
    # 1) 优先 sku_row - 1(最常见)
    for offset in (-1, 0, -2, 1, -3, 2):
        target = sku_row_0idx + offset
        if target in images_by_row and images_by_row[target]:
            return images_by_row[target][0]
    return None


def parse_rows(
    primary_file,
    vendor_id: str,
    *,
    parser: str = "v1010",
) -> List[Dict[str, object]]:
    """V1010 / V1038 共用的 parse_rows。

    vendor_id 由 vendor adapter 强制传入(避免 dispatcher 的 --warehouse-id 影响)。
    支持 .xls / .xlsx,.xlsx 时自动抽图(.xls 抽不到)。

    parser:
      - "v1010" (默认): V1010 风格描述解析,产出 (description, cost, retail,
                       origin, material, color),无 category
      - "v1038":         V1038 风格描述解析,产出 (description, cost, retail,
                       origin, category, color, material)
    """
    rows: List[Dict[str, object]] = []
    df = pd.read_excel(primary_file, sheet_name=0, header=None)
    if df.empty:
        return rows

    # 抽图(.xlsx 才有,.xls 跳过)
    images_by_row = _load_images_by_row(primary_file)

    blocks = find_sku_blocks(df)
    for blk in blocks:
        sku = blk["sku"]
        desc_text = blk["desc_text"]
        if parser == "v1038":
            description, cost, retail, origin, category, color, material = (
                parse_description_v1038(desc_text)
            )
        else:
            description, cost, retail, origin, material, color = parse_description(desc_text)
            # v1010 (Longchamp / Gucci-Hogan) 源表描述里只有材质+产地+颜色, 无品类词
            # 优先级: 数据源中直接表明 > 类目规则分类 > 空  → 此处描述无品类信息, 类目留空
            category = ""
        size_map, qty_map = find_size_qty(df, blk["desc_row"])

        if not size_map or not qty_map:
            # 没尺码/数量 - 跳过(不算有效产品行)
            continue

        # 找本产品的图片(基于 SKU 行的位置)
        image_item = _find_image_for_block(images_by_row, blk["sku_row"])

        for col, size in size_map.items():
            qty = qty_map.get(col, 0)
            if qty > 0:
                rows.append({
                    "brand": "",
                    "sku": sku,
                    "image": "",
                    "description": description,
                    "category": category,
                    "material": material,
                    "color": color,
                    "season": "",
                    "origin": origin,
                    "cost": cost,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "image_items": [image_item] if image_item else [],
                })
    return rows
