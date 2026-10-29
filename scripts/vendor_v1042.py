"""V1042 ALINA vendor adapter."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from convert_vendor_quote_shared import (
    is_missing,
    is_positive_qty,
    load_sheet_images_by_row,  # 抽图(.xls 走空,.xlsx 走 openpyxl)
    load_xls_ole2_jpeg_images,  # 抽图(.xls OLE2 JPEG 流 + F010 Client Anchor)
    normalize_text,
    to_float,
)

# 2026-09-12 fix: 跟 framework 主流一致 — qty=0 行跳过,不输出占位 row
# (用户偏好: 库存为 0 的行不输出,见 template-resource-converter/patterns/image-directory-matching.md §5)


def _parse_desc_with_prices(desc_line: str) -> Tuple[str, Optional[float], Optional[float], str, str, str]:
    # 百分号正则(识别材质段 token,如 "Cotton100 %")
    PCT_RE = re.compile(r"\d+\s*%")


    """V1042 描述解析器。

    真实版式(完整):
      "<类型> <性别> <颜色> <cost> <retail> <产地> <材质>"
      例:"Camicia Donna Bianco       257,00       720,00 ITALIA Natural (Veg)->Cotton72 %"
      例:"T-shirt Uomo Cobalto+grigio+chiaro+ofw 196,00 550,00 ITALIA Cotton100 %"
      例:"Borsa a Spalla Donna Antilope 257,00 720,00 ITALIA Cotton"

    拆分规则(实测 248 个独特描述,100% 都有性别词 + 100% 都有 ITALIA 产地):
      1. 抓末尾所有 XX,XX(可能 1-2 个: cost / retail)
      2. 抠掉价格
      3. 找最后 1 个性别词 → 性别词之前是 type+gender(描述)
      4. 性别词之后:第一个全大写单词(产地)之前是颜色
      5. 全大写单词 = origin(产地)
      6. 剩下的 = material(材质)

    返回: (desc, cost, retail, color, origin, material)
    """
    GENDER_WORDS = ("Uomo", "Donna", "Unisex", "Unica", "Bambino", "Bambina")
    # 时尚品牌常见产地(意大利/欧洲/亚洲) — 用集合查找,既覆盖全大写"ITALIA"又覆盖首大写"Romania"
    # 实测 POLO 现货单的产地都是首字母大写 (Indonesia / Vietnam / Cambogia / Stati Uniti 等),
    # 所以这里同时收录两种 case
    COUNTRY_WORDS = {
        # 全大写
        "ITALIA", "ITALY", "ROMANIA", "FRANCE", "SPAGNA", "SPAIN", "GERMANY",
        "PORTUGAL", "INDIA", "VIETNAM", "TURCHIA", "TURKEY", "TUNISIA",
        "MAROCCO", "MOROCCO", "CINA", "CHINA", "HONG", "MADAGASCAR",
        "MAURITIUS", "MAURIZIO", "INDONESIA", "CAMBODIA", "CAMBOGIA",
        "STATI", "UNITI", "TUNISIA", "UCRAINA", "MOLDAVIA", "BANGLADESH",
        "EGYPT", "EGITTO", "PAKISTAN", "SRI", "LANKA",
        # 首字母大写 (POLO 现货单风格)
        "Romania", "Italia", "Tunisia", "Vietnam", "Cambogia", "Indonesia",
        "Turchia", "Moldavia", "Ucraina", "Cina", "Francia", "Spagna",
        "Germania", "Portogallo", "India", "Stati", "Uniti",
        "Bangladesh", "Egitto", "Pakistan", "Sri", "Lanka",
    }

    text = normalize_text(desc_line)
    if not text:
        return "", None, None, "", "", ""

    # 1) 找所有 XX,XX 价格(从后往前扫,取最后 1-2 个当 cost/retail)
    all_prices = list(re.finditer(r"\d{1,5}[.,]\d{2}\b", text))
    if not all_prices:
        return text, None, None, "", "", ""
    # 取最后 1-2 个(cost, retail)
    last_two = all_prices[-2:] if len(all_prices) >= 2 else all_prices
    cost = to_float(last_two[0].group()) if len(last_two) >= 1 else None
    retail = to_float(last_two[1].group()) if len(last_two) >= 2 else None
    # 抠掉所有 XX,XX(把价格段移除,保留前后文字包括 ITALIA/Material)
    pre = re.sub(r"\s*\d{1,5}[.,]\d{2}\b\s*", " ", text).strip()
    pre = re.sub(r"\s+", " ", pre)  # 合并多空格

    # 2) 找最后 1 个性别词
    tokens = pre.split()
    gender_idx = -1
    for i in range(len(tokens) - 1, -1, -1):
        if tokens[i] in GENDER_WORDS:
            gender_idx = i
            break
    if gender_idx < 0:
        # 兜底:没找到性别词就当 desc 全部
        return pre, cost, retail, "", "", ""

    desc = " ".join(tokens[: gender_idx + 1])  # 包含 gender
    after_gender = tokens[gender_idx + 1:]  # 性别词之后的所有 token

    # 3) 在 after_gender 里找第一个产地(全大写 / 已知国家名 / 多词国家 "Stati Uniti" / "Repubblica Popolare Cinese")
    # 已知多词产地 (用于 2-3 词组合识别)
    MULTI_WORD_ORIGINS = [
        ("Stati", "Uniti"),
        ("Regno", "Unito"),
        ("Repubblica", "Ceca"),
        ("Repubblica", "Cinese"),
        ("Repubblica", "Dominicana"),
        ("Emirati", "Arabi"),
        ("Paesi", "Bassi"),
        ("Repubblica", "Popolare", "Cinese"),  # 3 词
        # typo: 源数据偶尔把 "Cinese" 和下一个 "India" 拼成 "CIndiaese"
        # 这种 typo 在 3 词组合里加上,优先匹配
        ("Repubblica", "Popolare", "CIndiaese"),
    ]
    upper_idx = -1
    multi_word_count = 0
    for i, tok in enumerate(after_gender):
        # 去掉尾部标点
        clean_tok = tok.rstrip(",.;:")
        if not clean_tok:
            continue
        # 跳过明显的材质 token (含 "->" 或 "%")
        if "->" in clean_tok or PCT_RE.search(clean_tok):
            # 材质段的 token 不可能是 origin
            continue
        # 先尝试匹配多词产地 (2-3 词)
        matched = False
        for combo in MULTI_WORD_ORIGINS:
            if len(combo) < 2 or len(combo) > 3:
                continue
            if i + len(combo) > len(after_gender):
                continue
            tokens_to_check = [after_gender[i + k].rstrip(",.;:") for k in range(len(combo))]
            if all(t == c for t, c in zip(tokens_to_check, combo)):
                upper_idx = i
                multi_word_count = len(combo)
                matched = True
                break
        if matched:
            break
        # 2 词产地:"Stati Uniti" — 当前 token 是 "Stati" 且下一个 token 是 "Uniti"
        if i + 1 < len(after_gender) and clean_tok in COUNTRY_WORDS:
            nxt = after_gender[i + 1].rstrip(",.;:")
            if nxt in COUNTRY_WORDS and nxt != clean_tok:
                # 2 词产地都识别到了
                upper_idx = i
                multi_word_count = 2
                break
        if clean_tok.isupper() and len(clean_tok) > 1:
            upper_idx = i
            break
        if clean_tok in COUNTRY_WORDS:
            upper_idx = i
            break

    if upper_idx < 0:
        # 没找到产地 token - 整个 after_gender 都是颜色
        color = " ".join(after_gender).strip()
        origin = ""
        material = ""
    else:
        color = " ".join(after_gender[:upper_idx]).strip()
        # 拼装多词产地
        if multi_word_count > 1:
            origin = " ".join(
                after_gender[upper_idx + k].rstrip(",.;:")
                for k in range(multi_word_count)
            )
            nxt_idx = upper_idx + multi_word_count
        else:
            origin = after_gender[upper_idx]
            nxt_idx = upper_idx + 1
        # typo 修正: "CIndiaese" 是源数据的拼写错误,实际意图 "Cinese" (China)
        if origin == "Repubblica Popolare CIndiaese":
            origin = "Repubblica Popolare Cinese"
        material = " ".join(after_gender[nxt_idx:]).strip()

    return desc.strip(), cost, retail, color, origin, material


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    # 抽图:.xls (OLE2) 走 load_xls_ole2_jpeg_images(F010 Client Anchor + JPEG SOI/EOI),
    # .xlsx 走 load_sheet_images_by_row(openpyxl)
    if str(primary_file).lower().endswith(".xls"):
        images_by_row = load_xls_ole2_jpeg_images(primary_file, sheet_index=0, image_col=3)
    else:
        images_by_row = load_sheet_images_by_row(primary_file)
    df = pd.read_excel(primary_file, header=None)
    brand = ""
    for i in range(min(30, len(df))):
        v = normalize_text(df.iloc[i, 1] if df.shape[1] > 1 else None)
        if v and len(v) > 1 and " " not in v[:1]:
            brand = v
            break

    def row_values(idx: int) -> List[object]:
        if idx < 0 or idx >= len(df):
            return []
        vals: List[object] = []
        for v in df.iloc[idx].tolist():
            if is_missing(v):
                continue
            vals.append(v)
        return vals

    # 工厂 V1042 (ALINA / POLO RALPH LAUREN) 的版式:6-row/product
    # - R+0: SKU (col 9 / 0-idx)
    # - R+1: 描述 (col 9,内嵌两个欧洲价格)
    # - R+3: 尺码 (整行所有有值的列 — 第一个是 size code, 之后是 size label)
    # - R+5: qty (整行所有有值的列,按列与 size_label 对齐)
    # 用**结构判断**找产品块(不用内容判断函数),跟 V1004/V1021 一致:
    # 1) R+1 col 9 必须含两个欧洲价格(描述行特征)
    # 2) R+3 / R+5 必须有 size + qty(用 row_values 检查)

    # 2026-09-12 POLO RALPH LAUREN fix:size row 第一个是 size code (如 INT/STN/ITD/FRD/UNI/PZ/CMU/UKU),
    # size_label 起点 = size_code_col + 1;qty row 第一个数字 col = qty 起点;
    # pair: size_label[i] ↔ qty_padded[i] (前 qty_起点 - size_label_起点 个 0 padding),
    # 保留全部 size (含 qty=0 写 0 行,让 framework 知道有货 size).
    SIZE_CODE_TOKENS = {"ITD", "FRD", "UNI", "PZ", "CMU", "UKU", "INT", "STN",
                       "FIT", "G", "FIN", "TU", "OS", "M", "L", "S", "XL", "XXL", "XXXL", "4XL", "5XL"}

    def _find_size_code_col(row_idx: int) -> Optional[int]:
        """找 size row 第一个 size code col (如 'INT'/'STN'/'UNI')."""
        if row_idx < 0 or row_idx >= len(df):
            return None
        for c in range(df.shape[1]):
            v = df.iat[row_idx, c]
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            sv = str(v).strip()
            if sv.upper() in SIZE_CODE_TOKENS:
                return c
        return None

    def _find_qty_start_col(row_idx: int) -> Optional[int]:
        """找 qty row 第一个数字 (>0) col."""
        if row_idx < 0 or row_idx >= len(df):
            return None
        for c in range(df.shape[1]):
            v = df.iat[row_idx, c]
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            f = to_float(v)
            if f is not None and f > 0:
                return c
        return None

    i = 0
    while i < len(df):
        # 2026-09-29 fix: 大多数 .xls SKU 在 col 9,但 MM6 等个别品牌在 col 8
        # 动态扫 col 7-10 找 SKU 列(找第一个非空 cell)
        sku_col = None
        for c in (9, 8, 10, 7):
            if c < df.shape[1]:
                v = normalize_text(df.iat[i, c])
                if v:
                    sku_col = c
                    break
        if sku_col is None:
            i += 1
            continue
        sku_cell = normalize_text(df.iat[i, sku_col])

        desc_line = normalize_text(df.iat[i + 1, sku_col] if i + 1 < len(df) else "")
        if not re.search(r"\d+[.,]\d+\s+\d+[.,]\d+", desc_line):
            i += 1
            continue

        # 2026-09-12 fix: 用 col 对齐, 不再 row_values 全收集 (会配错)
        size_code_col = _find_size_code_col(i + 3)
        qty_start_col = _find_qty_start_col(i + 5)
        if size_code_col is None or qty_start_col is None:
            # 兜底: 老 row_values 逻辑
            size_vals = row_values(i + 3)
            qty_vals = row_values(i + 5)
            if not size_vals or not qty_vals:
                i += 1
                continue
            sizes = [
                normalize_text(x)
                for x in size_vals
                if normalize_text(x).upper() not in SIZE_CODE_TOKENS
            ]
            qty_numbers = [to_float(x) or 0 for x in qty_vals if to_float(x) is not None]
            n = min(len(sizes), len(qty_numbers))
            pairs = [(sizes[j], int(round(qty_numbers[j]))) for j in range(n)]
        else:
            # size_label 起点 = size_code_col + 1
            size_label_start = size_code_col + 1
            # qty 数字按 col 收 (qty_start_col 起,直到行尾)
            qty_nums: List[float] = []
            for c in range(qty_start_col, df.shape[1]):
                v = df.iat[i + 5, c]
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    qty_nums.append(0.0)
                else:
                    f = to_float(v)
                    qty_nums.append(f or 0.0)
            # size_label 按 col 收 (size_label_start 起,直到行尾)
            size_labels: List[str] = []
            for c in range(size_label_start, df.shape[1]):
                v = df.iat[i + 3, c]
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    size_labels.append("")  # 末尾空 col 也算 (实际不常见)
                else:
                    size_labels.append(normalize_text(v))
            # 配对: offset = qty_起点 - size_label 起点, qty 前补 0
            offset = qty_start_col - size_label_start
            qty_padded = [0.0] * offset + qty_nums
            n = min(len(size_labels), len(qty_padded))
            pairs = [(size_labels[j], int(round(qty_padded[j]))) for j in range(n)]

        sku = sku_cell
        # 2026-08-14:从 SKU 拆 color_code (POLO 风格 "<style>---<color_code>")
        # 例: "211891376---001" → sku="211891376-001", color_code="001"
        # 没有 --- 分隔符就保留原样
        color_code = ""
        m = re.match(r"^(.+?)---+(.+)$", sku)
        if m:
            sku = f"{m.group(1)}-{m.group(2)}"
            color_code = m.group(2)
        description, cost, retail, color, origin, material = _parse_desc_with_prices(desc_line)

        # 拆 description 拿 type + gender
        # 格式: "<type_words> <gender> ..." — gender 是描述里唯一的 Uomo/Donna/Unisex 等 token
        # type 在 gender 之前,可能多词 (Borsa a Spalla / Borsa a Mano)
        desc_tokens = description.split() if description else []
        GENDER_TOKENS = ("Uomo", "Donna", "Unisex", "Unica", "Bambino", "Bambina", "Man", "Woman", "Men", "Women", "Kid", "Kids")
        GENDER_TRANSLATE = {
            "Uomo": "Man", "Men": "Man", "Man": "Man",
            "Donna": "Woman", "Women": "Woman", "Woman": "Woman",
            "Unisex": "Unisex", "Unica": "Unisex",
            "Bambino": "Kid", "Bambina": "Kid", "Kid": "Kid", "Kids": "Kid",
        }
        gender_idx = -1
        for idx, tok in enumerate(desc_tokens):
            if tok in GENDER_TOKENS:
                gender_idx = idx
                break
        if gender_idx >= 0:
            category = " ".join(desc_tokens[:gender_idx]).strip()
            gender = GENDER_TRANSLATE.get(desc_tokens[gender_idx], desc_tokens[gender_idx])
        else:
            # 没找到 gender token — 整段当 category
            category = description
            gender = ""

        for size_label, qty in pairs:
            if not size_label:
                continue
            if not is_positive_qty(qty):  # qty=0 跳过(framework 主流规则)
                continue
            rows.append(
                {
                    "brand": brand,
                    "sku": sku,
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
                    "supply_price": cost,
                    "retail_price": retail,
                    "size": size_label,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "gender": gender,
                    "image_items": images_by_row.get(i - 1, []),
                }
            )
        i += 6
    return rows
