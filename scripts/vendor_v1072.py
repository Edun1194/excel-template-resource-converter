"""V1072 vendor adapter — BURBERRY 现货单 (2026-08-13 新建)

来源:BURBERRY 工厂 V1072 的 .xlsx 现货单(BURBERRY-2.xlsx)。

跟 V1038 / V1062 高度相似(都是 "块状"版式:SKU/描述/尺码/数量四行一组),但
两个关键差异:

1. **SKU 在 col J (0-idx 9) 而不是 col N (0-idx 13)**
   V1038/V1062 走 col14_parsers;V1072 自己写 parser(共用图片加载逻辑)。

2. **描述格式是 V1072 风格,不是 V1038 风格**
   V1038 描述 = ``<类型> <颜色> <cost> <retail> <产地> <材料>``
   V1072 描述 = ``<类型> <性别> <颜色> <cost> <retail> <产地> <材料>``
   - 性别 token (Uomo/Donna) 出现在描述里 → 类目/性别能干净分开
   - 多词类型: "Borsa a Spalla" / "Borsa a Mano" / "borsa slg"
   - 产地可能是 "Regno Unito" / "RegNorvegia Unito" (源数据有 typo)
   - 材质可能是多段: "Synthetic->Polyurethane55 % Synthetic->Polyester25 % Natural (Veg)->Cotton20 %"

3. **尺码/数量两种版式**
   - 单尺码:col I=OS, col J=UNI, col K=PZ + qty → qty 在 col K(比 size 标签偏移 +1)
   - 多尺码:col I=INT, col J=XXXS/XXS/XS/S/M/L/XL/XXL/XXXL/4XL/5XL,
     qty 在 size 同一列

真实样本(BURBERRY-2.xlsx, 100 行 = 100 货号块):
- R5 img / R6 SKU / R7 desc / R9 size / R11 qty
- 描述: ``Cravatta Uomo Archive beige        88,60       195,00 Italia Natural (other)->Mulberry Silk100 %``
- 性别 Uomo → Man
- 类目 Cravatta → Cravatta
- 颜色 "Archive beige"
- 产地 Italia
- 材质 "Natural (other)->Mulberry Silk100 %"

仓库 ID 固定为 ``V1072``,品牌默认 ``BURBERRY``(可被 --brand 覆盖)。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

# V1072 SKU/描述 在 col J (0-idx 9)
SKU_COL_0IDX = 9
# 尺码/数量扫描范围
# - 0-idx 8 = col I (OS/INT/ITu/ITd/UKd/CMu prefix marker)
# - 0-idx 9 = col J (OS 时存 UNI size;多尺码时可能空)
# - 0-idx 10-30 = col K-AE (鞋码最大范围,ITu/ITd 一直延伸到 47½)
# - 0-idx 34 = col AH (源文件 max_col,留 buffer)
SIZE_COL_START = 8
SIZE_COL_END = 35  # cover up to col AI, 留 buffer 超过源文件 max_col 34 (AH)
# 已知性别 token(在描述里出现的位置)
GENDER_TOKENS = {"Uomo", "Donna", "Unisex", "Bambina", "Bambino", "Kid", "Kids", "Man", "Woman", "Men", "Women"}
# 材质段开头(看到这些 token 之前的都是 origin)
MATERIAL_STARTS = (
    "Natural", "Wool", "Synthetic", "Furs", "Cashmere", "Cotton",
    "Leather", "Polyester", "Polyamide", "Polyurethane", "Silk",
    "Linen", "Cashwool", "Foam", "Rubber", "Elastane", "Polyacrylic",
    "Viscose", "Mohair", "Alpaca", "Polypropylene", "Metal",
)
# 尺码白名单(出现这些说明这行是尺码行,col I = OS/INT 标志)
SIZE_LABELS = {
    "OS", "UNI", "PZ", "INT", "TU", "TU.", "ONE", "ONE SIZE", "UNICA",
    "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
    "3XL", "4XL", "5XL", "6XL",
}
# 价格正则(欧式 XX,XX)
PRICE_RE = re.compile(r"\d{1,5}\s*,\s*\d{2}")
# 百分号正则(材质)
PCT_RE = re.compile(r"\d+\s*%")

# 产地 typo 修正(源数据里 "RegNorvegia Unito" 应解读为 "Regno Unito" = UK)
ORIGIN_FIX = {
    "RegNorvegia Unito": "Regno Unito",
    "RegNorvegia": "Regno Unito",
    "Regno Unito": "Regno Unito",
    "Repubblica Ceca": "Repubblica Ceca",
    "Repubblica Popolare Cinese": "China",
}

# SKU 拼装:`<style>---<color_code>` → `<style>-<color_code>` (符合 `<数字>-<颜色>` 格式)
SKU_RE = re.compile(r"^([A-Za-z0-9]+?)---+([A-Za-z0-9]+)$")


def _norm(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float):
        if pd.isna(val):
            return ""
        # 整数 size (e.g. 43.0) 显示成 "43" 不要 ".0"
        if val.is_integer():
            return str(int(val))
        return str(val)
    if not isinstance(val, str):
        val = str(val)
    return val.strip()


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return False


def _to_int(v) -> Optional[int]:
    if _is_blank(v):
        return None
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return None


def _format_sku(sku_raw: str) -> Tuple[str, str]:
    """把 V1072 SKU ``8011693---A7026`` → ``(8011693-A7026, A7026)`` 拆出 style + color_code。"""
    if not sku_raw:
        return "", ""
    m = SKU_RE.match(sku_raw)
    if m:
        style, color_code = m.group(1), m.group(2)
        return f"{style}-{color_code}", color_code
    # fallback:保留原样
    return sku_raw, ""


def _parse_description_v1072(desc: str) -> Tuple[str, Optional[float], Optional[float], str, str, str, str]:
    """V1072 描述解析。

    真实样本:
      "Cravatta Uomo Archive beige        88,60       195,00 Italia Natural (other)->Mulberry Silk100 %"
        → ("Cravatta Uomo Archive beige", 88.6, 195.0, "ITALIA", "Cravatta", "Archive beige", "Uomo",
           "Natural (other)->Mulberry Silk100 %")
      "Felpa Uomo Black       375,00       825,00 Portogallo Natural (Veg)->Cotton100 %"
        → ("Felpa Uomo Black", 375.0, 825.0, "PORTUGAL", "Felpa", "Black", "Uomo",
           "Natural (Veg)->Cotton100 %")
      "Borsa a Spalla Donna Black       434,10       955,00 Romania Synthetic->Polyamide100 %"
        → ("Borsa a Spalla Donna Black", 434.1, 955.0, "ROMANIA", "Borsa a Spalla", "Black", "Donna",
           "Synthetic->Polyamide100 %")
      "Sciarpa Donna Archive beige       372,70       820,00  Wool or fine animal hair->Cashmere100 %"
        → origin 空

    返回:(description, cost, retail, origin, category, color, gender, material)
    """
    if not desc:
        return "", None, None, "", "", "", "", ""

    # 1) 抠出价格
    prices = list(PRICE_RE.finditer(desc))
    if len(prices) >= 2:
        cost_str = prices[-2].group().replace(" ", "").replace(",", ".")
        retail_str = prices[-1].group().replace(" ", "").replace(",", ".")
        try:
            cost = float(cost_str)
            retail = float(retail_str)
        except ValueError:
            return desc, None, None, "", "", "", "", ""
        before = desc[: prices[-2].start()].rstrip()
        after = desc[prices[-1].end():].strip()
    else:
        cost = retail = None
        before = desc.strip()
        after = ""

    before = re.sub(r"\s+", " ", before).strip()
    after = re.sub(r"\s+", " ", after).strip()

    # 2) after:拆 origin + material(material 以 MATERIAL_STARTS 开头)
    after_tokens = after.split() if after else []
    mat_idx: Optional[int] = None
    for i, tok in enumerate(after_tokens):
        # 材质 token 标志:以 MATERIAL_STARTS 开头 / 含 '->' / 含 '%'
        if any(tok.startswith(ms) for ms in MATERIAL_STARTS) or "->" in tok or PCT_RE.search(tok):
            mat_idx = i
            break
    if mat_idx is None and after_tokens:
        # 没找到材质段 — 整段当 material,origin 空
        origin = ""
        material = " ".join(after_tokens)
    elif mat_idx is None:
        origin = ""
        material = ""
    else:
        origin = " ".join(after_tokens[:mat_idx]).strip()
        material = " ".join(after_tokens[mat_idx:]).strip()

    # 产地 typo 修正
    if origin in ORIGIN_FIX:
        origin = ORIGIN_FIX[origin]

    # 3) before:找 gender token(在 GENDER_TOKENS 集合里)
    before_tokens = before.split() if before else []
    gender = ""
    gender_idx: Optional[int] = None
    for i, tok in enumerate(before_tokens):
        if tok in GENDER_TOKENS:
            gender = tok
            gender_idx = i
            break
    if gender_idx is None:
        # 没有 gender token — 整段当 category,color 空
        category = " ".join(before_tokens)
        color = ""
    else:
        category = " ".join(before_tokens[:gender_idx]).strip()
        color = " ".join(before_tokens[gender_idx + 1:]).strip()

    # 4) description = before 原文(类型+性别+颜色)
    description = before

    return description, cost, retail, origin, category, color, gender, material


def _load_images_by_row(primary_file) -> Dict[int, List[Dict]]:
    """从 .xlsx 加载所有图片,按行号索引(0-idx)。"""
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
    """V1072 的图片在 SKU 行上方 1 行(col D)。±3 行容忍。"""
    if not images_by_row:
        return None
    for offset in (-1, 0, -2, 1, -3, 2):
        target = sku_row_0idx + offset
        if target in images_by_row and images_by_row[target]:
            return images_by_row[target][0]
    return None


def _gender_to_english(g: str) -> str:
    """把意大利语性别 token 翻译成英文。"""
    g_low = g.lower()
    if g_low in ("uomo", "men", "man"):
        return "Man"
    if g_low in ("donna", "women", "woman"):
        return "Woman"
    if g_low in ("bambina", "bambino", "kid", "kids"):
        return "Kid"
    if g_low in ("unisex",):
        return "Unisex"
    return g


def _origin_normalize(o: str) -> str:
    """V1072 产地白名单归一化(意大利语 → 英文)。"""
    if not o:
        return ""
    mapping = {
        "ITALIA": "ITALY",
        "Italia": "ITALY",
        "Portogallo": "PORTUGAL",
        "Tailandia": "THAILAND",
        "Regno Unito": "UK",
        "Moldavia": "MOLDOVA",
        "Romania": "ROMANIA",
        "Vietnam": "VIETNAM",
        "Scozia": "SCOTLAND",
        "Repubblica Ceca": "CZECH REPUBLIC",
        "Repubblica Popolare Cinese": "CHINA",
        "China": "CHINA",
        "Spagna": "SPAIN",
        "Francia": "FRANCE",
        "Germania": "GERMANY",
        "Turchia": "TURKEY",
        "Cina": "CHINA",
        "India": "INDIA",
        "Marocco": "MOROCCO",
        "Tunisia": "TUNISIA",
        "Egitto": "EGYPT",
        "Norvegia": "NORWAY",
    }
    return mapping.get(o, o)


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1072 (BURBERRY) 现货单。

    强制把 vendor_id 改成 V1072,保证输出「云仓」列稳定。
    性别/类目/颜色 全部从描述里拆。
    """
    rows: List[Dict[str, object]] = []
    df = pd.read_excel(primary_file, sheet_name=0, header=None)
    if df.empty:
        return rows
    if df.shape[1] < SKU_COL_0IDX + 1:
        return rows

    images_by_row = _load_images_by_row(primary_file)

    n = len(df)
    for r in range(n):
        # SKU 行:col J (0-idx 9) 有非空短字符串
        v = df.iloc[r, SKU_COL_0IDX]
        if _is_blank(v):
            continue
        sku_raw = _norm(v)
        if not sku_raw or " " in sku_raw or "/" in sku_raw:
            continue
        if not SKU_RE.match(sku_raw):
            # 跳过尺码行(OS/UNI/INT/XXXS...)
            if sku_raw in SIZE_LABELS:
                continue
            # 不是已知 SKU 格式,跳过
            continue
        # 描述行(下一行)
        if r + 1 >= n:
            continue
        desc_v = df.iloc[r + 1, SKU_COL_0IDX]
        desc_text = _norm(desc_v)
        if not desc_text:
            continue
        # 尺码行(r+3) + 数量行(r+5)
        size_row = r + 3
        qty_row = r + 5
        if size_row >= n or qty_row >= n:
            continue

        # 找 size label col I (0-idx 8)
        col_i_val = _norm(df.iloc[size_row, 8]) if df.shape[1] > 8 else ""
        mode = ""
        if col_i_val == "OS":
            mode = "OS"
        elif col_i_val in ("INT", "ITu", "ITd", "UKd", "CMu", "UKu", "IT"):
            # INT (clothing intl) / ITu (Italian shoes uomo) / ITd (Italian shoes donna) /
            # UKd (UK clothing donna? kids?) / CMu (centimeters) / UKu (UK shoes uomo) /
            # IT (Italian generic)
            mode = "MULTI"
        else:
            # 不是有效尺码行,跳过
            continue

        # 解析尺码/数量
        size_entries: List[Tuple[str, int]] = []
        if mode == "OS":
            # 单尺码:size 在 col J (0-idx 9),qty 在 col K (0-idx 10)
            size_label = _norm(df.iloc[size_row, 9]) if df.shape[1] > 9 else ""
            if not size_label or size_label in ("PZ",):
                # J 也是 PZ(没 UNI 标签),跳过
                continue
            qty = _to_int(df.iloc[qty_row, 10]) if df.shape[1] > 10 else 0
            if qty and qty > 0:
                size_entries.append((size_label, qty))
        else:
            # 多尺码:尺码在 SIZE_COL_START+1 (=col J,0-idx 9) 到 SIZE_COL_END 之间,
            # 可能是 XXXS/XXS/XS/S/M/L/XL/XXL/XXXL/4XL/5XL (服装)
            # 或 37/38/39/39½/40/40½.../47½ (鞋码,延伸到 col AE 0-idx 30)
            # 或 75/80/85.../125 (厘米,腰带)
            # 或 2/4/6/8/10/12/14/16/18/20 (UK 童码)
            # qty 在同列 (size_row + 2 = qty_row, col 相同)
            for c in range(9, min(SIZE_COL_END, df.shape[1])):
                label = _norm(df.iloc[size_row, c])
                if not label:
                    continue
                # 排除 "PZ" 这种单位标签
                if label in ("PZ",):
                    continue
                # 任何非空 label 都接受
                qty = _to_int(df.iloc[qty_row, c]) if df.shape[1] > c else 0
                if qty and qty > 0:
                    size_entries.append((label, qty))

        if not size_entries:
            continue

        # 解析描述
        description, cost, retail, origin_raw, category, color, gender_token, material = (
            _parse_description_v1072(desc_text)
        )
        origin = _origin_normalize(origin_raw)
        gender = _gender_to_english(gender_token)

        # SKU 拼装:`<style>-<color_code>`
        sku_formatted, color_code = _format_sku(sku_raw)

        # 找图
        image_item = _find_image_for_block(images_by_row, r)

        for size, qty in size_entries:
            rows.append({
                "brand": "",
                "sku": sku_formatted,
                "image": "",
                "description": description,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code,
                "season": "",
                "origin": origin,
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size,
                "quantity": qty,
                "vendor_id": vendor_id,
                "gender": gender,
                "image_items": [image_item] if image_item else [],
            })
    return rows


VENDOR_ID = "V1072"
DEFAULT_BRAND = "BURBERRY"


def get_vendor_meta() -> Dict[str, object]:
    """返回该 vendor 的默认元数据。

    - default_brand: BURBERRY
    """
    return {
        "default_brand": DEFAULT_BRAND,
    }
