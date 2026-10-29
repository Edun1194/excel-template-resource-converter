"""V1021 BRUNA ROSSO vendor adapter."""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
import pandas as pd

from convert_vendor_quote_shared import (
    load_sheet_embedded_images,
    load_sheet_row_heights,
    normalize_text,
    to_float,
)


def _normalize(s: str) -> str:
    """规范化字符串：去除变音符号，转大写"""
    if not s:
        return ""
    # NFD 分解字符，然后去除组合字符（去除变音符号）
    normalized = unicodedata.normalize('NFD', s)
    ascii_str = normalized.encode('ascii', 'ignore').decode()
    return ascii_str.upper().strip()


def _parse_cost_condition(condition_text: str) -> Optional[int]:
    """从 C+ 文本中提取第一个数字，如 'Cost +15%' -> 15"""
    if not condition_text:
        return None
    match = re.search(r'\+(\d+)%', str(condition_text))
    if match:
        return int(match.group(1))
    return None


def _load_cplus_map(conditions_file: Path) -> Dict[str, int]:
    """解析 Sales Conditions 文件，返回 brand -> C+ 的映射（key 已规范化）"""
    cplus_map: Dict[str, int] = {}
    try:
        df = pd.read_excel(conditions_file, sheet_name=0, header=None)
        current_group = ""
        for _, row in df.iterrows():
            brand = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            condition = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""

            if "GROUP" in brand.upper():
                current_group = brand
                continue
            if brand == "BRAND" or not brand or brand == "nan":
                continue

            cplus = _parse_cost_condition(condition)
            if cplus is not None:
                # 用规范化后的品牌名作为 key
                cplus_map[_normalize(brand)] = cplus
    except Exception as e:
        print(f"Warning: Failed to load conditions file: {e}")
    return cplus_map


def _parse_description(desc_text: str) -> tuple:
    """从描述文本中提取 description, cost, retail, origin, material
    格式如: 'Portachiavi Uomo Sellier+TBMR+P       226,00       520,00 ITALIA Furs & Skins->Calf Leather100 %'
    意大利价格格式: \d+,\d{2}  (如 226,00  520,00)
    返回 (description, cost, retail, origin, material)
    """
    if not desc_text:
        return "", None, None, "", ""

    # 意大利价格格式: 数字用逗号为小数点，如 226,00
    price_matches = re.findall(r'\d+,\d{2}', desc_text)

    if len(price_matches) >= 2:
        cost_str = price_matches[0]
        retail_str = price_matches[1]
        cost = float(cost_str.replace(',', '.'))
        retail = float(retail_str.replace(',', '.'))

        # description 是第一个价格数字之前的文本
        desc_parts = desc_text.split(cost_str, 1)
        description = desc_parts[0].strip() if desc_parts else ""

        # 提取产地 (ITALIA 或 Italy)
        origin_match = re.search(r'(ITALIA|Italy)', desc_text)
        origin = origin_match.group(1) if origin_match else ""

        # 提取材质 - 格式如: Furs & Skins->Calf Leather100 %
        # 实际材质是 "->" 后面的 "Calf Leather100 %"
        after_second_price = desc_text.split(retail_str)[-1] if len(desc_text.split(retail_str)) > 1 else ""
        
        # 先找 -> 后面的部分（包含材质名和百分比）
        mat_match = re.search(r'->\s*([A-Za-z &]+)\s*(\d+)\s*%', after_second_price)
        if mat_match:
            mat_name = mat_match.group(1).strip()
            mat_pct = mat_match.group(2)
            material = f'{mat_name} {mat_pct}%'
        else:
            # 备用：原来的正则匹配
            material_parts = re.findall(r'([A-Za-z &\-()>]+?)\s*(\d+)\s*%', after_second_price)
            materials = []
            for mat_name, pct in material_parts:
                mat_clean = mat_name.strip(' ->-')
                if mat_clean and mat_clean not in ['ITALIA', 'Italy']:
                    materials.append(f'{mat_clean} {pct}%')
            material = ', '.join(materials)

        return description, cost, retail, origin, material

    return desc_text, None, None, "", ""


def _extract_meta_from_header(df: pd.DataFrame) -> Dict[str, str]:
    """从表头几行提取元数据：season / brand
    文件第 1 行附近会有 STAGIONE / BRAND 标签，右侧就是值
    """
    meta: Dict[str, str] = {}
    for r in range(min(10, len(df))):
        for c in range(df.shape[1]):
            v = df.iloc[r, c]
            if not isinstance(v, str):
                continue
            label = v.strip().upper()
            # 找 "STAGIONE" 标签，看右侧 1-10 列找值
            if label == "STAGIONE":
                for off in range(1, 12):
                    val = df.iloc[r, c + off] if c + off < df.shape[1] else None
                    if isinstance(val, str) and val.strip() and val.strip().upper() != "STAGIONE":
                        meta["season"] = val.strip()
                        break
            elif label == "BRAND":
                for off in range(1, 12):
                    val = df.iloc[r, c + off] if c + off < df.shape[1] else None
                    if isinstance(val, str) and val.strip() and val.strip().upper() != "BRAND":
                        meta["brand"] = val.strip()
                        break
    return meta


def _find_category_for_row(df: pd.DataFrame, sku_row: int) -> str:
    """从 SKU 行往上找最近的 sub-category（Abbigliamento/Accessori/Scarpe/Maglieria/...）
    sub-category 一般在 col 3 或 col 4
    """
    candidates = []
    for r in range(max(0, sku_row - 30), sku_row):
        for c in (3, 4):
            if c >= df.shape[1]:
                continue
            v = df.iloc[r, c]
            if isinstance(v, str):
                s = v.strip()
                # 类目词：意大利语常见类目
                if s and s not in ("Donna", "Uomo", "UOMO", "DONNA", "Paia", "Totali", "Stocklist", "nan"):
                    if len(s) < 30 and any(c.isalpha() for c in s):
                        candidates.append((sku_row - r, s))
    if not candidates:
        return ""
    # 选最近的（差值最小）
    candidates.sort()
    return candidates[0][1]


def _split_color_from_description(description: str) -> tuple:
    """从 description 里把颜色拆出来

    只在"明确是 产品名+颜色"模式才拆（描述恰好 2 个 token 且没有数字）：
    - ✅ "T-shirt White"  → ("T-shirt", "White")
    - ❌ "Polo Ralph Lauren Abiti"（类目被误判为颜色）→ 不拆
    - ❌ "Felpa 100% Cotone"（数字）→ 不拆

    拆错了比漏拆危险大，宁可少拆。
    """
    if not description:
        return description, ""
    parts = description.split()
    if len(parts) != 2:
        return description, ""
    # 第二个 token 必须像颜色：首字母大写 + 纯字母（无数字） + 较短
    last = parts[-1]
    if (last and last[0].isupper()
            and any(c.isalpha() for c in last)
            and not any(c.isdigit() for c in last)
            and len(last) <= 15):
        return parts[0], last
    return description, ""


def _split_sku_color(sku: str) -> tuple:
    """把 POLO RALPH LAUREN 风格的 SKU 拆成 (数字前缀, 颜色)

    真实模式（105 个样本观察）：
    - 211870245002NEWPORT NAVY          → ("211870245002", "NEWPORT NAVY")
    - 211A12745001WHITE                  → ("211A12745001", "WHITE")
    - 211891430001BRIGHT NAVY/WHITE      → ("211891430001", "BRIGHT NAVY/WHITE")
    - 2119107435031936 SUPPLY OLIVE WHITE → ("2119107435031936", "SUPPLY OLIVE WHITE")
    - 710680784424BATH PINK/C7172        → ("710680784424", "BATH PINK/C7172")
    - 710717115024ANDOVER HEATHER/C7927  → ("710717115024", "ANDOVER HEATHER/C7927")
    - 211965843500WHITE-RED              → ("211965843500", "WHITE-RED")

    拆分规则（按优先级）：
    1. 末尾是 /C#### 这种 Ralph Lauren 色卡码 → 整段（BATH PINK/C7172）算颜色，
       最后一个 "/" 之前的数字段算 article。
    2. 否则找最后一个数字，后面是颜色（适用 95% 情况）。
    """
    if not sku or len(sku) < 8:
        return sku, ""

    # 规则 1：末尾是 /C####（色卡码）
    c_suffix_match = re.search(r'/?C\d{4}$', sku)
    if c_suffix_match:
        c_suffix_start = c_suffix_match.start()  # "/C####" 的起点
        # 往前找最后一个数字（article 末尾）
        article_end = -1
        for i in range(c_suffix_start):
            if sku[i].isdigit():
                article_end = i
        if article_end >= 0:
            prefix = sku[: article_end + 1]
            color = sku[article_end + 1 :]
            digit_count = sum(1 for c in prefix if c.isdigit())
            if digit_count >= 6 and any(c.isalpha() for c in color):
                return prefix, color.strip()

    # 规则 2：找最后一个数字
    last_digit_pos = -1
    for i in range(len(sku) - 1, -1, -1):
        if sku[i].isdigit():
            last_digit_pos = i
            break
    if last_digit_pos < 0:
        return sku, ""
    prefix = sku[: last_digit_pos + 1]
    color = sku[last_digit_pos + 1 :].strip()
    digit_count = sum(1 for c in prefix if c.isdigit())
    if digit_count < 6:
        return sku, ""
    if not color or not any(c.isalpha() for c in color):
        return sku, ""
    return prefix, color


def _extract_color_from_sku(sku: str) -> str:
    """从 SKU 拆颜色（取 _split_sku_color 第二部分）"""
    _, color = _split_sku_color(sku)
    return color


def _strip_color_from_sku(sku: str) -> str:
    """从 SKU 去掉颜色（取 _split_sku_color 第一部分）"""
    prefix, _ = _split_sku_color(sku)
    return prefix


# 欧洲十进制价格正则(用于结构判断描述行)
_EURO_PRICE_RE = re.compile(r"\d+[.,]\d{2}")


def _has_european_price(text: str) -> bool:
    """判断字符串里是否含欧洲十进制价格(如 '50,00' / '37.50')。"""
    if not text:
        return False
    return bool(_EURO_PRICE_RE.search(text))


def _find_product_blocks(df: pd.DataFrame) -> List[Dict]:
    """扫描整个 sheet,找到所有产品块(用**结构判断**,不用 _is_sku 内容判断)。

    结构判断的两条铁律:
    1) R+1..R+3 某行 col c 是描述行(同列),且含欧洲价格 → 描述行存在
    2) 描述行之后(R+2..R+12)能找到尺码+qty(用 _find_size_and_qty)→ 尺码行存在

    这两条加起来,任何"长得像内容"的行(描述、尺码、段头)都不可能通过,
    所以**根本不需要 _is_sku 内容判断** — R 那个单元格只要是非空字符串就行,
    SKU 字符串本身可以是任何长度、任何字符集(24 字符的 POLO、17 字符的
    STONE ISLAND、9 字符的 PATAGONIA 都能吃)。

    BRUNA ROSSO 不同子供应商/季度的版式差很多:
    - 老格式:SKU 下一行直接是 "产品名 价格 价格 origin"(描述含价格)
    - 新格式:SKU 下面 1-2 行是空行,再下一行是描述(价格独立在 col 35 等)
    - 这里两种都吃(offset 1-3 行)
    """
    blocks: List[Dict] = []
    max_col = len(df.columns)

    for r in range(len(df)):
        # 扫所有列(不限定 col,V1021 历史上支持多列 SKU)
        for c in range(max_col):
            val = df.iloc[r, c]
            if not (pd.notna(val) and isinstance(val, str)):
                continue
            sku = val.strip()
            if not sku:
                continue

            # 找描述:R+1..R+3 某行同列是非空字符串,且含欧洲价格
            desc_val = None
            desc_row = None
            for off in (1, 2, 3):
                if r + off >= len(df):
                    break
                nv = df.iloc[r + off, c]
                if isinstance(nv, str) and nv.strip():
                    if _has_european_price(nv):
                        desc_val = nv.strip()
                        desc_row = r + off
                        break
            if desc_val is None:
                continue

            # 找尺码+qty(用 _find_size_and_qty 灵活扫 cols 8-24)
            size_map, qty_map = _find_size_and_qty(df, desc_row)
            if not size_map or not qty_map:
                continue

            blocks.append({
                'sku_row': r,
                'desc_row': desc_row,
                'sku': sku,
                'desc_text': desc_val,
                'desc_col': c,
                'size_map': size_map,
                'qty_map': qty_map,
            })
            break  # 同一行找到一个 SKU 就够
    return blocks


def _find_size_and_qty(df: pd.DataFrame, desc_row: int) -> tuple:
    """在描述行下方查找尺码行和数量行
    返回 (size_map: {col: size_label}, qty_map: {col: qty})

    BRUNA ROSSO 的尺码/数量在描述下方（offset +2 到 +10）
    ⚠️ 不往上方搜：上面是上一个 product block 的尺码/数量，
       用错会导致用错尺码组（比如 Polo 那种）。
    """
    size_map: Dict[int, str] = {}
    qty_map: Dict[int, int] = {}
    max_col = len(df.columns)

    # 只在描述行下方搜（避免错用上一个 block 的尺码）
    for offset in range(2, 12):
        check_row = desc_row + offset
        if check_row >= len(df):
            break

        row = df.iloc[check_row]
        # 检查这一行是否有有效的尺码标签（非空字符串或数字）
        size_cols = []
        for c in range(8, min(25, max_col)):
            val = row[c]
            if pd.notna(val):
                val_str = str(val)
                # 尺码可以是字符串（如 'OS', 'UNI', 'PZ'）或数字（如 3, 3.5, 39）
                # 排除太长的字符串和太大的数字（库存会是更大的数）
                if isinstance(val, str) and len(val_str) < 20:
                    size_cols.append((c, val_str))
                elif isinstance(val, (int, float)) and not isinstance(val, bool) and abs(val) < 1000:
                    size_cols.append((c, val_str))

        if size_cols:
            # 这是尺码行
            for c, label in size_cols:
                label_str = str(label).strip()
                # 格式化：如果值是浮点数如 7.0 显示为 7，3.5 显示为 3.5
                try:
                    if isinstance(label, float) and label == int(label):
                        label_str = str(int(label))
                except:
                    pass
                if label_str and label_str not in ['nan', 'None']:
                    size_map[c] = label_str

            # 下一行应该是数量行（也可能隔一行）
            for qty_offset in [1, 2]:
                qty_row = check_row + qty_offset
                if qty_row < len(df):
                    qty_row_data = df.iloc[qty_row]
                    for c in size_map.keys():
                        if c < max_col and pd.notna(qty_row_data[c]):
                            try:
                                qty = int(float(qty_row_data[c]))
                                if qty > 0:
                                    qty_map[c] = qty
                            except:
                                pass
                    if qty_map:
                        break

            break

    return size_map, qty_map


def _find_brand_name(df: pd.DataFrame) -> str:
    """在 DataFrame 前几行中查找品牌名"""
    for r in range(min(10, len(df))):
        val = df.iloc[r, 1]  # 通常在 col 1
        if pd.notna(val) and isinstance(val, str):
            val_str = val.strip()
            if val_str and len(val_str) < 50 and not val_str.startswith('nan'):
                # 这可能是品牌名
                # 排除明显的非品牌名
                if any(c.isalpha() for c in val_str) and 'GROUP' not in val_str.upper():
                    return val_str
    return ""


def _load_images_by_row(primary_file: Path) -> Dict[int, List[Dict]]:
    """从 .xlsx 加载所有图片，按行号索引
    返回: {row_0indexed -> [{bytes, format}]}
    对 .xls (OLE2 老格式) 返回空 dict（暂不处理）
    """
    path_str = str(primary_file)
    if not path_str.endswith(".xlsx"):
        print(f"[V1021] ⚠️ 源文件是 .xls，图片无法自动抽取。请用 WPS/Excel 转成 .xlsx 再跑。")
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


def _find_image_for_block(images_by_row: Dict[int, List[Dict]], sku_row_0idx: int) -> Optional[Dict]:
    """找最接近 SKU 行的图片(sku_row 附近 ±2 行)。

    BRUNA ROSSO 的真实产品图都在 SKU 行正上方 1 行。±2 足以容忍小幅偏移,
    又能避免把"隔壁产品的图"借给"没有图的产品"(供应商用占位图的场景)。
    """
    if not images_by_row:
        return None
    candidates = []
    for img_row, items in images_by_row.items():
        dist = abs(img_row - sku_row_0idx)
        if dist <= 2:  # ±2 行内(贴近真实产品图,避免借图)
            candidates.append((dist, items[0]))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


def parse_rows(primary_file: Path, conditions_file: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []

    # 加载图片（仅 .xlsx 支持，.xls 暂不抽）
    images_by_row = _load_images_by_row(primary_file)

    # 加载 C+ 规则
    cplus_map = _load_cplus_map(conditions_file) if conditions_file else {}

    # 读取 stock 文件
    df = pd.read_excel(primary_file, sheet_name=0, header=None)
    if df.empty:
        return rows

    # 获取品牌名称（动态查找）
    brand_name_raw = _find_brand_name(df)
    brand_name = _normalize(brand_name_raw) if brand_name_raw else ""

    # 提取表头元数据（季节/品牌）
    header_meta = _extract_meta_from_header(df)
    season_value = header_meta.get("season", "")

    # 获取该品牌的 C+
    # 先精确匹配，再尝试前缀匹配（如 "MM6" 匹配 "MM6 MAISON MARGIELA"）
    cplus = cplus_map.get(brand_name)
    if cplus is None:
        for cond_brand, cond_cplus in cplus_map.items():
            if brand_name.startswith(cond_brand) or cond_brand.startswith(brand_name):
                cplus = cond_cplus
                break

    # 查找所有产品块(用结构判断,块里已经预计算 size_map/qty_map)
    blocks = _find_product_blocks(df)

    for block in blocks:
        sku = block['sku']
        desc_text = block['desc_text']
        description, cost, retail, origin, material = _parse_description(desc_text)

        # 拆出颜色：先看 SKU 后缀（POLO 模式），再看 description（2-token 模式）
        sku_color = _extract_color_from_sku(sku)
        if sku_color:
            color = sku_color
            # 货号保留完整 SKU，中间用 - 拼接（数字部分 + "-" + 颜色后缀）
            # 例: 211870245002NEWPORT NAVY → 211870245002-NEWPORT NAVY
            number_part = _strip_color_from_sku(sku)
            sku = f"{number_part}-{color}"
        else:
            description, color = _split_color_from_description(description)

        # 找该产品所属的 sub-category
        category = _find_category_for_row(df, block['sku_row'])

        # 尺码 + qty(在 _find_product_blocks 阶段已经验证过非空,这里直接用)
        size_map = block.get('size_map') or _find_size_and_qty(df, block['desc_row'])[0]
        qty_map = block.get('qty_map') or _find_size_and_qty(df, block['desc_row'])[1]

        # 找本产品的图片（基于 SKU 行的位置）
        image_item = _find_image_for_block(images_by_row, block['sku_row'])

        for col, size in size_map.items():
            qty = qty_map.get(col, 0)
            if qty > 0:
                rows.append({
                    "brand": brand_name,
                    "sku": sku,
                    "image": "",  # 图片是嵌入对象，write_output 会处理 image_items
                    "description": description,
                    "category": category,
                    "material": material,
                    "color": color,
                    "season": season_value,
                    "origin": origin,
                    "cost": cost,
                    "cplus_text": cplus if cplus else "",
                    "supply_price": None,
                    "retail_price": retail,
                    "size": size,
                    "quantity": qty,
                    "vendor_id": vendor_id,
                    "image_items": [image_item] if image_item else [],
                })

    return rows
