"""V1067 PANICALE vendor adapter (内置,2026-07-28 从 external 改内置以便 zip 打包).

适配文件:PANICALE 意大利线库存清单(.xlsx,无图,图片外部目录匹配)。

支持两种版式(2026-08-07 加 packing list mode):

【Mode 1: 横排尺码 + 双行(legacy)】
- R1 header (中文): c1=品牌, c3=商品国际编码(SKU), c4=COLOUR CODE(色码),
                    c6=商品描述, c7=分类ID(可选), c8=颜色(色名),
                    c9=SIZES(尺码起点), c19=材质, c20=cost出厂(欧元),
                    c21=加价率(±%, C+ 字段), c22=供货价, c23=原价(retail),
                    c24=库存合计
- "尺码行":c3 非空, c4 空 → 提取 c9-c18 数字/单码 = size_labels
- "颜色行":c3 非空, c4 非空 → 按 size_labels 读 c9-c18 = 各尺码 qty
- 特殊产品(D39SOCKS / D39TIE / D39TOPCRYSTAL):size 行只有 c9 一个值 (TU/XXS/XXL),
  颜色行 c9 = 单尺码 qty

【Mode 2: Detailed packing list (per-row)】(2026-08-07 加)
- R1 header (英文): c1=D.d.T. Num. / c2=Data D.d.T. / c3=Image (列)
                   c4=Description / c5=Style number / c6=Colour Code / c7=Colour
                   c8=Size / c9=Qty / c10=Unit price / c11=Discount
                   c12=Final price / c13=Box no. / c14=Barcode
                   c15=Customs code / c16=Origin
- R2+ 数据行:每行 1 个 size+qty(per-row 模式)
- 70+ 张内嵌图(ws._images 命中)
- SKU 拼法:`f"{c5}-{c6.replace(' ', '-')}"` = "U380010PA-38U24-0580"

【Mode 3: Stock list with ARTICOLO/ARTICLE blocks】(2026-08-18 加,SS26 新版式)
- 没有 R1 header,每页有 5 行页眉 (PANICALE SRL / 地址 / STAMPA...) 然后是块
- 每个块:
  - R(N): A(N):AI(N) merged, c1 = "ARTICOLO/ARTICLE:  <SKU> - <desc> - <material>  Prezzo/Price: <price>"
  - R(N+1): C(N+1):AI(N+1) merged (c3 模式) 或 c1 only (c1 模式),文本以 "Colore/Color" 开头
    size 标签从文本解析 (空格分隔)
  - R(N+2)..R(N+X): 每个颜色一行
    - c3 模式: 颜色在 c3 或 c4,格式 "<color_code>\\n<color_name>",qty 在 c11/c14/c17/c20/c23/c26/c29/c31/c33/c35
    - c1 模式: 颜色代码 + 名称拆 2 行,或合并 "code\\nname" 在 c1;qty 在随机列(只取首行 qty)
- 193 张内嵌图 (per-block anchor)
- SKU 拼法: `<SKU>-<color_code>` (c3 模式) 或 `<SKU>-<color_code>` (c1 模式,只取首行)
- 88 个 ARTICOLO 块 / 250+ c3 行 / 17+ c1 行 (例: SS26_STOCKUOMO_test.xlsx)
- 价格取 Prezzo/Price (cost 出厂)
- 季节: 文件名 "SS26" → season="SS26"

识别:看 sheet 有没有 "ARTICOLO/ARTICLE" 文本 — 有 → Mode 3,否则按 R1 c5 是不是 "Style number" 判 Mode 2 vs Mode 1。

【Mode 4: FW26 STOCK LIST 尺码行/颜色行 + T 列内嵌图】(2026-11-02 加)
- 源: "PANICALE FW26 STOCK LIST {MEN|WOMEN} 30.10.xlsx"(供应商直接给带图版本)
  单 sheet 'in', R1 中文表头:
    c1  商品国际编码 + 色号 | c2 商品描述 | c3 颜色 | c4-c13 尺码矩阵(10 列)
    | c14 材质 | c15 数量 | c16 cost出厂（欧元）| c17 加价率（+/-%）
    | c18 供货价(公式) | c19 原价(欧元)(公式) | c20 图片(内嵌图锚在这列)
- "尺码行":c1 单 token(如 "U390010PA")→ c4-c13 是该 SKU 的尺码标签
- "颜色行":c1 三 token(如 "U390010PA 39U36 0200")→ c4-c13 是各尺码数量
- 图片: 供应商已把图嵌在 c20(锚在颜色行), 不用再去图库匹配
- ⚠️ 加价率两个文件写法不同: MEN 是 50(整数百分数), WOMEN 是 0.5(小数) → 都要认
- ⚠️ c18/c19 是公式 '=P2*1.5' / '=P2*2.5', 缓存值是 '#VALUE!' → 自己按系数算
- ⚠️ "TU"(一码通)SKU 的数量可能落在没标签的列上 → 整行尺码格求和
- 货号拼法: `<国际编码>-<色号 空格改->`, 色号 = c1 的后两段("39U36 0200"), 跟客户图库文件名一致

图片外部目录(Mode 1 才有):`<V1067>/PANICALE高清图-货号+色号/FW26 {WOMEN|MEN}/`
文件名:`<SKU> <色码>.{jpg|png}`
SKU 变体映射:xlsx 的 `/M` `/NN` `/MC` 等 → 文件名 `-M` `-NN` `-MC`(`/` → `-`)
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)

# Mode 1 列定义(legacy 横排尺码)
COL_SKU = 2          # c3
COL_COLOUR_CODE = 3  # c4
COL_DESC = 5         # c6
COL_COLOR_NAME = 7   # c8
COL_SIZE_START = 8   # c9(尺码 col 起点,0-idx)
COL_SIZE_END = 18    # c18(尺码 col 终点,0-idx exclusive)
COL_MATERIAL = 18    # c19
COL_COST = 19        # c20
COL_CPLUS_PCT = 20   # c21
COL_SUPPLY = 21      # c22
COL_RETAIL = 22      # c23

# Mode 2 列定义(packing list per-row)
PL_COL_DESC = 3       # c4 Description (e.g. "TROUSERS")
PL_COL_STYLE = 4      # c5 Style number (e.g. "U380010PA")
PL_COL_COLOUR_CODE = 5  # c6 Colour Code (e.g. "38U24 0580")
PL_COL_COLOUR = 6     # c7 Colour (e.g. "DEEP" / "GRAVEL")
PL_COL_SIZE = 7       # c8 Size (e.g. "50" / "48")
PL_COL_QTY = 8        # c9 Qty (e.g. "1")
PL_COL_UNIT_PRICE = 9  # c10 Unit price (e.g. "138")
PL_COL_DISCOUNT = 10  # c11 Discount (e.g. "-0.5")
PL_COL_FINAL_PRICE = 11  # c12 Final price (e.g. "69")
PL_COL_BARCODE = 13   # c14 Barcode (e.g. "2025000593190")
PL_COL_ORIGIN = 15    # c16 Origin (e.g. "ITALY")

SIZE_LABEL = "ONE SIZE"  # 单尺码产品(SOCKS/TIE)fallback


# ---------------------------------------------------------------------------
# Mode 4 列定义 (FW26 STOCK LIST 尺码行/颜色行)
# ---------------------------------------------------------------------------
ST_COL_KEY = 1        # c1  商品国际编码 + 色号
ST_COL_DESC = 2       # c2  商品描述
ST_COL_COLOR = 3      # c3  颜色
ST_COL_SIZE_START = 4 # c4  尺码矩阵起点(10 列)
ST_COL_SIZE_END = 13  # c13 尺码矩阵终点(exclusive)
ST_COL_MATERIAL = 14  # c14 材质
ST_COL_QTY = 15       # c15 数量(合计)
ST_COL_COST = 16      # c16 cost出厂（欧元）
ST_COL_CPLUS = 17     # c17 加价率（+/-%）
ST_COL_SUPPLY = 18    # c18 供货价(公式)
ST_COL_RETAIL = 19    # c19 原价(欧元)(公式)
ST_COL_IMAGE = 20     # c20 图片(内嵌图锚在这一列)

ONE_SIZE_LABEL = "TU"  # 一码通(源里尺码标签写 'TU')


def _to_int_qty(value) -> int:
    f = to_float(value)
    if f is None:
        return 0
    return int(round(f))


def _pick_gender_dir(primary_file: Path) -> Optional[str]:
    """从 xlsx 文件名判断目录:WOMEN / MEN。

    GIANCE DONNA → WOMEN, GIACENZE UOMO → MEN。
    """
    name = primary_file.stem.upper()
    if "DONNA" in name or "DONNE" in name:
        return "FW26 WOMEN"
    if "UOMO" in name or " UOMO" in f" {name} " or name.endswith(" MEN"):
        return "FW26 MEN"
    if " MEN" in f" {name} ":
        return "FW26 MEN"
    return None


def _load_image_bytes(
    image_base: Path, sku_raw: str, colour_code: str,
) -> List[Dict[str, object]]:
    """从外部目录 rglob 找图片,文件名 <sku_safe> <colour_safe>。<ext>。

    2026-07-28:支持任意深度 — 之前只在 image_base 顶层找,现在递归整个目录
    (zip 解压后结构可能是 images/xxx.jpg 或直接 xxx.jpg 或 FW26 MEN/xxx.jpg)。

    SKU 变体映射:`/M` `/NN` `/MC` 等 → `-M` `-NN` `-MC`(匹配 PANICALE 命名规则)。
    """
    if not image_base.is_dir():
        return []
    sku_safe = sku_raw.replace("/", "-").strip()
    colour_safe = colour_code.replace("/", "-").strip()
    base = f"{sku_safe} {colour_safe}"
    for ext in ("jpg", "jpeg", "png", "JPG", "PNG", "JPEG"):
        # 先顶层精确匹配(快路径)
        top = image_base / f"{base}.{ext}"
        if top.is_file():
            return [{
                "bytes": top.read_bytes(),
                "format": "jpeg" if ext.lower() in ("jpg", "jpeg") else ext.lower(),
            }]
        # 任意深度匹配(rglob)
        for f in image_base.rglob(f"{base}.{ext}"):
            if f.is_file():
                return [{
                    "bytes": f.read_bytes(),
                    "format": "jpeg" if ext.lower() in ("jpg", "jpeg") else ext.lower(),
                }]
    return []


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 openpyxl 0-idx row 分桶图(Mode 2 per-row 模式,2026-08-07 升级)。

    PANICALE SS26 packing list 的图是"产品块共享"模式 — 一张图 anchor._from.row
    是块的第一行(只有第一行被 _data() 解析出来),但 anchor._to.row 是块的
    最后一行(openpyxl 默认不展开覆盖范围)。

    实测:
      img 1: from=R2 to=R7  → R2-R7 6 行共享同一张图
      img 3: from=R14 to=R24 → R14-R24 11 行共享同一张图

    本函数用 from → to 范围填满所有中间行,这样 _parse_packing_list 拿到
    images_by_row[r] 时任何一行(包括 _from 之后的所有尺码行)都能拿到图。
    """
    out: Dict[int, List[Dict]] = {}
    for img in ws._images:
        anchor = img.anchor
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        try:
            data = img._data() if hasattr(img, "_data") else None
        except Exception:
            data = None
        if data is None:
            continue
        from_row = anchor._from.row
        # to row(2-cell anchor 才有;one-cell anchor 没有 to)
        to_row = from_row
        if hasattr(anchor, "to") and anchor.to is not None:
            try:
                to_row = anchor.to.row
            except Exception:
                to_row = from_row
        # to_row < from_row(异常情况)就只用 from
        if to_row < from_row:
            to_row = from_row

        entry = {
            "bytes": data,
            "format": img.format or "jpeg",
        }
        # 范围填满 from..to(含)
        for r in range(from_row, to_row + 1):
            out.setdefault(r, []).append(entry)
    return out


def get_vendor_meta() -> Dict[str, object]:
    """V1067 PANICALE vendor 元数据。

    default_brand=PANICALE:固定,CLI --brand 可覆盖。
    default_image_col_width=40:跟其他 vendor 统一。
    """
    return {
        "default_brand": "PANICALE",
        "default_image_col_width": 40.0,
    }


def _detect_mode(ws) -> str:
    """检测版式:Mode 2 (packing list) 还是 Mode 1 (legacy 双行)。

    R1 c5 (0-idx 4) == "Style number" → Mode 2。
    """
    c5_header = normalize_text(ws.cell(row=1, column=PL_COL_STYLE + 1).value).upper()
    if c5_header == "STYLE NUMBER":
        return "packing_list"
    return "legacy"


def _parse_packing_list(ws, vendor_id: str) -> List[Dict[str, object]]:
    """Mode 2:Detailed packing list (per-row 模式) — 2026-08-07 加。

    R1 header(英文),R2+ 每行 1 个 size+qty。

    2026-08-07 增强:合并 (sku+size+color) 完全相同的重复行,累加库存。
    源数据里同一货号同尺码同色号会有多行(可能是分录/分箱,模板没法体现),
    累加后只留一行,避免视觉重复。
    """
    rows: List[Dict[str, object]] = []
    # key → rows index 映射(用于累加重复行)
    key_to_idx: Dict[Tuple[str, str, str], int] = {}
    images_by_row = _load_images_by_row(ws)

    for r in range(2, ws.max_row + 1):
        # 1) SKU (Style number)
        style = normalize_text(ws.cell(row=r, column=PL_COL_STYLE + 1).value)
        if not style:
            continue

        # 2) Colour Code (色号)
        colour_code_raw = normalize_text(ws.cell(row=r, column=PL_COL_COLOUR_CODE + 1).value)
        # SKU 拼法: <Style>-<ColourCode 空格改 - >
        # e.g. "U380010PA" + "38U24 0580" → "U380010PA-38U24-0580"
        colour_code_norm = colour_code_raw.replace(" ", "-") if colour_code_raw else ""
        sku = f"{style}-{colour_code_norm}" if colour_code_norm else style

        # 3) Description
        desc = normalize_text(ws.cell(row=r, column=PL_COL_DESC + 1).value)

        # 4) Colour (色名)
        colour = normalize_text(ws.cell(row=r, column=PL_COL_COLOUR + 1).value)

        # 5) Size
        size = normalize_text(ws.cell(row=r, column=PL_COL_SIZE + 1).value)
        if not size:
            continue

        # 6) Qty
        qty_v = ws.cell(row=r, column=PL_COL_QTY + 1).value
        if not is_positive_qty(qty_v):
            continue
        qty = _to_int_qty(qty_v)
        if qty <= 0:
            continue

        # 7) cost (Unit price) / retail (Final price)
        cost_v = ws.cell(row=r, column=PL_COL_UNIT_PRICE + 1).value
        cost = to_float(cost_v)
        retail_v = ws.cell(row=r, column=PL_COL_FINAL_PRICE + 1).value
        retail = to_float(retail_v)

        # 8) Origin
        origin = normalize_text(ws.cell(row=r, column=PL_COL_ORIGIN + 1).value)

        # 9) 图 (per-row,0-idx row = r - 1)
        zero_idx = r - 1
        row_images = images_by_row.get(zero_idx, [])

        # 10) 合并重复行(2026-08-07):(sku, size, colour) 三元组完全一致 → 累加 qty
        dedup_key = (sku, size, colour)
        if dedup_key in key_to_idx:
            existing_idx = key_to_idx[dedup_key]
            rows[existing_idx]["quantity"] = rows[existing_idx]["quantity"] + qty
            # 重复行不增加 rows 计数,只累加 qty
            continue

        new_row = {
            "brand": "PANICALE",
            "sku": sku,
            "image": "",
            "description": desc,
            "category": "",  # 源没类目字段
            "material": "",
            "color": colour,
            "color_code": colour_code_raw,  # 中文模板"色号"列
            "gender": "",
            "season": "",  # 留空,main 走文件名兜底 → "SS26"
            "origin": origin,
            "cost": cost,
            "cplus_text": "",
            "supply_price": None,
            "retail_price": retail,
            "size": size,
            "quantity": qty,
            "vendor_id": vendor_id,
            "image_items": row_images,
            "source_row_height": None,
        }
        rows.append(new_row)
        key_to_idx[dedup_key] = len(rows) - 1

    return rows


# Mode 3 (2026-08-18 加):ARTICOLO/ARTICLE 块模式 — SS26 新版式
ARTICOLO_PREFIX = "ARTICOLO/ARTICLE"
COLOR_HEADER_PREFIX = "Colore/Color"
# 颜色代码 pattern: 2 数字 + 大写字母 + 数字组 (e.g. "38U24 0200", "38UPOC 0555")
COLOR_CODE_RE = re.compile(r"^\d{2}[A-Z]+\s+\d+")
# 页眉识别(避免误解析):PANICALE SRL / Via del / Tel. / P.Iva / STAMPA
PAGE_HEADER_TOKENS = ("PANICALE SRL", "Via del", "Tel. +39", "P.Iva", "STAMPA", "panicale@pec.it")
# c3 模式固定 qty 列(0-idx):K(11)/N(14)/Q(17)/T(20)/W(23)/Z(26)/AC(29)/AE(31)/AG(33)/AI(35)
C3_QTY_COLS = [11, 14, 17, 20, 23, 26, 29, 31, 33, 35]


def _detect_new_layout(ws) -> bool:
    """Mode 3 检测:扫描 sheet 找 "ARTICOLO/ARTICLE" 文本 — 是 → Mode 3。"""
    for r in range(1, min(ws.max_row + 1, 200)):  # 前 200 行够了
        v = ws.cell(r, 1).value
        if v and ARTICOLO_PREFIX in str(v):
            return True
    return False


def _find_articolo_blocks(ws) -> List[Tuple[int, int, str, str, str, Optional[float]]]:
    """扫所有 ARTICOLO/ARTICLE 行,返回 (art_r, end_r, sku, desc, material, price) 列表。

    end_r 是下一个 ARTICOLO 之前的行(或 sheet 末尾)。
    """
    art_rows: List[int] = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and ARTICOLO_PREFIX in str(v):
            art_rows.append(r)

    blocks: List[Tuple[int, int, str, str, str, Optional[float]]] = []
    for i, art_r in enumerate(art_rows):
        end_r = art_rows[i + 1] - 1 if i + 1 < len(art_rows) else ws.max_row
        art_text = str(ws.cell(art_r, 1).value)
        # 优先匹配 "ARTICOLO/ARTICLE: <SKU> - <desc/mat> Prezzo/Price: <price>"
        m = re.match(
            r"ARTICOLO/ARTICLE:\s*(\S+)\s+-\s+(.+?)\s+Prezzo/Price:\s*([\d,]+)",
            art_text,
        )
        if m:
            sku = m.group(1).strip()
            rest = m.group(2)
            price = float(m.group(3).replace(",", "."))
            # 拆分 desc 和 material — 末尾含 % 的是 material
            parts = [p.strip() for p in rest.split(" - ")]
            if parts and "%" in parts[-1]:
                material = parts[-1]
                desc = " - ".join(parts[:-1])
            else:
                material = ""
                desc = rest.strip()
        else:
            # fallback: 只有 SKU,没 desc/price (例 "U381351G" 短格式)
            m2 = re.match(r"ARTICOLO/ARTICLE:\s*(\S+)", art_text)
            if not m2:
                continue
            sku = m2.group(1).strip()
            desc = ""
            material = ""
            price = None
        blocks.append((art_r, end_r, sku, desc, material, price))
    return blocks


def _is_page_header_row(ws, r: int) -> bool:
    """判断 row r 是不是页眉行(5 行 PANICALE SRL 头)。"""
    for c in range(1, 6):
        v = ws.cell(r, c).value
        if v and any(tok in str(v) for tok in PAGE_HEADER_TOKENS):
            return True
    return False


def _parse_color_text(text: str) -> Tuple[Optional[str], str]:
    """从 "color_code\\ncolor_name" 或纯 "color_code" 拆 (code, name)。

    返回 (code, name)。code 可能是 None(纯名称行)。
    """
    if not text:
        return None, ""
    s = str(text).strip()
    if not s:
        return None, ""
    if "\n" in s:
        parts = s.split("\n", 1)
        return parts[0].strip(), parts[1].strip()
    # 单行:判断是不是 color_code pattern
    if COLOR_CODE_RE.match(s):
        return s, ""
    return None, s


def _collect_c3_block(
    ws, art_r: int, end_r: int, sku: str, desc: str, material: str, price: Optional[float],
    season: str, vendor_id: str, block_images: List[Dict],
) -> List[Dict[str, object]]:
    """c3 模式解析:颜色在 c3 或 c4,格式 '<code>\\n<name>',qty 在固定列。"""
    rows: List[Dict[str, object]] = []
    # 1) 找 Colore/Color header(前 5 行内,在 c1-c4 找)
    cc_r = None
    cc_col = None
    for r in range(art_r + 1, min(art_r + 5, end_r + 1)):
        for c in (1, 2, 3, 4):
            v = ws.cell(r, c).value
            if v and COLOR_HEADER_PREFIX in str(v):
                cc_r = r
                cc_col = c
                break
        if cc_r:
            break
    if not cc_r:
        return rows
    # 2) 解析 size 标签
    cc_text = str(ws.cell(cc_r, cc_col).value)
    m = re.search(r"Colore/Color\s+(.*?)$", cc_text)
    sizes = m.group(1).split() if m else []

    # 3) 扫颜色行
    for r in range(cc_r + 1, min(cc_r + 15, end_r + 1)):
        # 页眉行/下一个 ARTICOLO 跳过
        c1_v = ws.cell(r, 1).value
        if c1_v and ARTICOLO_PREFIX in str(c1_v):
            break
        if _is_page_header_row(ws, r):
            break
        c3_v = ws.cell(r, 3).value
        c4_v = ws.cell(r, 4).value
        # 取第一个非空的颜色 cell
        color_text = c3_v if c3_v else c4_v
        if not color_text:
            continue
        color_str = str(color_text).strip()
        if COLOR_HEADER_PREFIX in color_str:
            continue
        color_code, color_name = _parse_color_text(color_str)
        if not color_code:
            continue
        # 4) 按 size 列读 qty
        for si, size in enumerate(sizes):
            if si >= len(C3_QTY_COLS):
                break
            col = C3_QTY_COLS[si]
            v = ws.cell(r, col).value
            # 接受 int / float / 数字字符串
            qty = _to_int_qty(v)
            if qty <= 0:
                continue
            full_sku = f"{sku}-{color_code}" if color_code else sku
            rows.append({
                "brand": "PANICALE",
                "sku": full_sku,
                "image": "",
                "description": desc,
                "category": "",
                "material": material,
                "color": color_name,
                "color_code": color_code,
                "gender": "",
                "season": season,
                "origin": "",
                "cost": price,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,
                "size": size,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": block_images,
                "source_row_height": None,
            })
    return rows


def _collect_c1_block(
    ws, art_r: int, end_r: int, sku: str, desc: str, material: str, price: Optional[float],
    season: str, vendor_id: str, block_images: List[Dict],
) -> List[Dict[str, object]]:
    """c1 模式解析:Colore/Color 在 c1,颜色代码 + 名称拆 2 行或合并 1 行。

    因为源数据 c1 模式 qty 列位置完全无规律,降级为:
    - 每个 color 1 行,size 用 sizes[0] (或 "ONE SIZE"),qty=1 (源数据本身缺 qty)
    - POCHETTE 之类的合并单行格式按 \\n 拆分
    """
    rows: List[Dict[str, object]] = []
    # 1) 找 Colore/Color header
    cc_r = None
    cc_col = None
    for r in range(art_r + 1, min(art_r + 5, end_r + 1)):
        for c in (1, 2, 3, 4):
            v = ws.cell(r, c).value
            if v and COLOR_HEADER_PREFIX in str(v):
                cc_r = r
                cc_col = c
                break
        if cc_r:
            break
    if not cc_r:
        return rows
    # 2) 解析 size 标签
    cc_text = str(ws.cell(cc_r, cc_col).value)
    m = re.search(r"Colore/Color\s+(.*?)$", cc_text)
    sizes = m.group(1).split() if m else []
    default_size = sizes[0] if sizes else SIZE_LABEL

    # 3) 扫颜色行 — 用"行序列配对"算法:每行有可能是 color_code 或 color_name
    pending_code: Optional[str] = None
    pending_name: str = ""
    for r in range(cc_r + 1, min(cc_r + 15, end_r + 1)):
        c1_v = ws.cell(r, 1).value
        if c1_v and ARTICOLO_PREFIX in str(c1_v):
            break
        if _is_page_header_row(ws, r):
            break
        c3_v = ws.cell(r, 3).value
        c4_v = ws.cell(r, 4).value
        # 取首非空 cell 内容
        primary = c1_v if c1_v else (c3_v if c3_v else c4_v)
        if not primary:
            continue
        primary_str = str(primary).strip()
        if not primary_str or COLOR_HEADER_PREFIX in primary_str:
            continue
        # 拆分 "code\nname" 格式
        if "\n" in primary_str:
            code, name = _parse_color_text(primary_str)
            if code:
                full_sku = f"{sku}-{code}" if code else sku
                rows.append({
                    "brand": "PANICALE",
                    "sku": full_sku,
                    "image": "",
                    "description": desc,
                    "category": "",
                    "material": material,
                    "color": name,
                    "color_code": code,
                    "gender": "",
                    "season": season,
                    "origin": "",
                    "cost": price,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": None,
                    "size": default_size,
                    "quantity": 1,
                    "vendor_id": vendor_id,
                    "image_items": block_images,
                    "source_row_height": None,
                })
            continue
        # 单行:判断是 code 还是 name
        if COLOR_CODE_RE.match(primary_str):
            pending_code = primary_str
            pending_name = ""
        else:
            # 是 name,配对上一个 pending_code
            pending_name = primary_str
            if pending_code:
                full_sku = f"{sku}-{pending_code}"
                rows.append({
                    "brand": "PANICALE",
                    "sku": full_sku,
                    "image": "",
                    "description": desc,
                    "category": "",
                    "material": material,
                    "color": pending_name,
                    "color_code": pending_code,
                    "gender": "",
                    "season": season,
                    "origin": "",
                    "cost": price,
                    "cplus_text": "",
                    "supply_price": None,
                    "retail_price": None,
                    "size": default_size,
                    "quantity": 1,
                    "vendor_id": vendor_id,
                    "image_items": block_images,
                    "source_row_height": None,
                })
                pending_code = None
                pending_name = ""
    # 4) 处理没配对 name 的尾随 code (defensive)
    return rows


def _parse_new_layout(
    primary_file: Path, ws, vendor_id: str,
) -> List[Dict[str, object]]:
    """Mode 3 解析:ARTICOLO/ARTICLE 块模式 — 2026-08-18 加。

    流程:
    1) 找所有 ARTICOLO 块
    2) 把内嵌图按 anchor 行归到对应块
    3) 对每块判断 c3 vs c1 模式,分别解析
    """
    rows: List[Dict[str, object]] = []
    blocks = _find_articolo_blocks(ws)
    if not blocks:
        return rows

    # 1) 图片 → 块 映射(按 anchor 行归属)
    images_by_block: Dict[int, List[Dict]] = {}
    art_r_list = [b[0] for b in blocks]
    for img in ws._images:
        anchor = img.anchor
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        try:
            data = img._data() if hasattr(img, "_data") else None
        except Exception:
            data = None
        if data is None:
            continue
        # 找 anchor 行对应的块(找 <= anchor 的最大 art_r)
        anchor_row = anchor._from.row + 1  # 转 1-idx
        block_idx = None
        for i, ar in enumerate(art_r_list):
            if ar <= anchor_row:
                block_idx = i
            else:
                break
        if block_idx is None:
            continue
        entry = {
            "bytes": data,
            "format": img.format or "jpeg",
        }
        images_by_block.setdefault(block_idx, []).append(entry)

    # 2) 季节从文件名提取
    season = ""
    name_upper = primary_file.stem.upper()
    # 优先 4 字符 (SS26/FW26),再回退 2 字符
    m = re.search(r"(SS\d{2}|FW\d{2})", name_upper)
    if not m:
        m = re.search(r"\b(SS|FW)\b", name_upper)
    if m:
        season = m.group(1)
    if not season:
        season = "SS26"  # fallback

    # 3) 逐块解析
    for i, (art_r, end_r, sku, desc, material, price) in enumerate(blocks):
        block_images = images_by_block.get(i, [])
        # 找 Colore/Color 找 col — 决定走 c3 还是 c1
        cc_col = None
        for r in range(art_r + 1, min(art_r + 5, end_r + 1)):
            for c in (1, 2, 3, 4):
                v = ws.cell(r, c).value
                if v and COLOR_HEADER_PREFIX in str(v):
                    cc_col = c
                    break
            if cc_col:
                break
        if cc_col is None:
            continue
        if cc_col in (1, 2):
            rows.extend(_collect_c1_block(
                ws, art_r, end_r, sku, desc, material, price,
                season, vendor_id, block_images,
            ))
        else:
            rows.extend(_collect_c3_block(
                ws, art_r, end_r, sku, desc, material, price,
                season, vendor_id, block_images,
            ))
    return rows


# ---------------------------------------------------------------------------
# Mode 4: FW26 STOCK LIST (尺码行 / 颜色行 + T 列内嵌图)
# ---------------------------------------------------------------------------

def _detect_stocklist_layout(ws) -> bool:
    """是不是 Mode 4: R1 表头 c1 含「国际编码」且 c4 是「尺码」。

    ⚠️ 必须排在 Mode 1(legacy) 前面 —— 两者都是 R1 中文表头, Mode 1 会误判。

    ⚠️ 两个文件的 c1 表头不一样(只有这一列不同, 其余列完全一致):
      MEN   → "商品国际编码 + 色号"
      WOMEN → "国际编码"
    所以按「含 国际编码」来判, 不要写死整串。
    """
    key_header = normalize_text(ws.cell(row=1, column=ST_COL_KEY).value)
    size_header = normalize_text(ws.cell(row=1, column=ST_COL_SIZE_START).value)
    return "国际编码" in key_header and size_header == "尺码"


def _stocklist_file_meta(path: Path) -> Tuple[str, str]:
    """从文件名取 (性别, 季节)。

    "PANICALE FW26 STOCK LIST MEN 30.10_with_images.xlsx" → ("Man", "FW26")
    "PANICALE FW26 STOCK LIST WOMEN 30.10_with_images.xlsx" → ("Woman", "FW26")
    """
    stem = path.stem.upper()
    gender = ""
    if re.search(r"\bWOMEN\b", stem) or re.search(r"\bDONNA\b", stem):
        gender = "Woman"
    elif re.search(r"\bMEN\b", stem) or re.search(r"\bUOMO\b", stem):
        gender = "Man"
    m = re.search(r"\b(FW|SS)\s*(\d{2})\b", stem)
    season = f"{m.group(1)}{m.group(2)}" if m else "FW26"
    return gender, season


def _normalize_cplus_pct(value: Optional[float]) -> Optional[float]:
    """加价率归一成「百分数整数」。

    ⚠️ 两个文件写法不一样: MEN 是 50(百分数), WOMEN 是 0.5(小数) → 都要认。
    """
    if value is None:
        return None
    v = float(value)
    if 0 < abs(v) <= 1:
        v *= 100
    return v


def _stocklist_price_multipliers(path: Path) -> Tuple[Optional[float], Optional[float]]:
    """读 c18(供货价) / c19(原价) 的公式系数 → (供货价系数, 原价系数)。

    源里这两列是公式 '=P2*1.5' / '=P2*2.5', 但缓存值是 '#VALUE!'(公式错误),
    data_only=True 读不到值 → 再用 read_only + data_only=False 读一遍公式文本。
    """
    supply_mult: Optional[float] = None
    retail_mult: Optional[float] = None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        ws = wb[wb.sheetnames[0]]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            for col, is_supply in ((ST_COL_SUPPLY, True), (ST_COL_RETAIL, False)):
                if (is_supply and supply_mult is not None) or (not is_supply and retail_mult is not None):
                    continue
                v = row[col - 1] if len(row) >= col else None
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                m = re.search(r"\*\s*([0-9]*\.?[0-9]+)", v)
                if m:
                    if is_supply:
                        supply_mult = float(m.group(1))
                    else:
                        retail_mult = float(m.group(1))
            if supply_mult is not None and retail_mult is not None:
                break
        wb.close()
    except Exception:
        pass
    return supply_mult, retail_mult


def _parse_stocklist(
    primary_file: Path, ws, vendor_id: str,
) -> List[Dict[str, object]]:
    """Mode 4: PANICALE FW26 STOCK LIST(尺码行/颜色行 + 内嵌图)。

    结构:
      尺码行: c1 单 token → c4-c13 是该 SKU 的尺码标签(每个 SKU 一张自己的尺码表)
      颜色行: c1 三 token → c4-c13 是各尺码数量

    货号: `<国际编码>-<色号 空格改->`
      色号 = c1 的后两段("39U36 0200") —— 跟客户图库文件名 `<国际编码> <色号>` 一致。
    """
    rows: List[Dict[str, object]] = []
    images_by_row = _load_images_by_row(ws)
    gender, season = _stocklist_file_meta(primary_file)
    supply_mult, retail_mult = _stocklist_price_multipliers(primary_file)

    size_labels: List[Tuple[int, str]] = []   # (列号, 尺码标签) —— 记住列号, 不用序号推
    size_row_style: str = ""           # 当前尺码表属于哪个 style
    last_cost: Optional[float] = None
    last_cplus: Optional[float] = None

    for r in range(2, ws.max_row + 1):
        c1 = normalize_text(ws.cell(r, ST_COL_KEY).value)
        if not c1:
            continue
        parts = c1.split()

        cost = to_float(ws.cell(r, ST_COL_COST).value)
        cplus_pct = _normalize_cplus_pct(to_float(ws.cell(r, ST_COL_CPLUS).value))

        # ---- 尺码行: 只带尺码标签 + 本 SKU 的价格信息 ----
        if len(parts) == 1:
            size_row_style = parts[0]
            size_labels = []
            for c in range(ST_COL_SIZE_START, ST_COL_SIZE_END + 1):
                sv = normalize_text(ws.cell(r, c).value)
                if sv:
                    size_labels.append((c, sv))
            if cost is not None:
                last_cost = cost
            if cplus_pct is not None:
                last_cplus = cplus_pct
            continue

        # ---- 颜色行: c1 = "<国际编码> <色号1> <色号2>" ----
        if len(parts) != 3:
            continue
        style = parts[0]
        colour_code = " ".join(parts[1:])
        sku = f"{style}-{colour_code.replace(' ', '-')}"

        if cost is None:
            cost = last_cost
        if cplus_pct is None:
            cplus_pct = last_cplus

        desc = normalize_text(ws.cell(r, ST_COL_DESC).value)
        colour = normalize_text(ws.cell(r, ST_COL_COLOR).value)
        material = normalize_text(ws.cell(r, ST_COL_MATERIAL).value)

        # 供货价 = cost × (1 + 加价率) ; 原价 = cost × 源公式系数(2.5)
        supply_price: Optional[float] = None
        if cost is not None and cplus_pct is not None:
            supply_price = round(cost * (1 + cplus_pct / 100.0), 2)
        retail_price: Optional[float] = None
        if cost is not None and retail_mult:
            retail_price = round(cost * retail_mult, 2)

        img_items = images_by_row.get(r - 1, [])

        base_row: Dict[str, object] = {
            "brand": "PANICALE",
            "sku": sku,
            "image": "",
            "description": desc,
            "category": "",              # 源没类目字段(跟 Mode 1/2/3 一致)
            "material": material,
            "color": colour,
            "color_code": colour_code,
            "gender": gender,
            "season": season,
            "origin": "",                # 源没产地列
            "cost": cost,
            "cplus_text": str(int(cplus_pct)) if cplus_pct is not None else "",
            "supply_price": supply_price,
            "retail_price": retail_price,
            "vendor_id": vendor_id,
            "image_items": img_items,
            "source_row_height": None,
        }

        # 尺码表对不上(颜色行不属于上一个尺码行) → 不能按位置取量
        if size_row_style != style or not size_labels:
            continue

        # 一码通(TU): 只有一个尺码, 但数量可能落在没标签的列上 → 整行尺码格求和
        if len(size_labels) == 1 and size_labels[0][1] == ONE_SIZE_LABEL:
            total_qty = 0.0
            for c in range(ST_COL_SIZE_START, ST_COL_SIZE_END + 1):
                qv = to_float(ws.cell(r, c).value)
                if qv and qv > 0:
                    total_qty += qv
            if total_qty > 0:
                rows.append({**base_row, "size": ONE_SIZE_LABEL,
                             "quantity": int(round(total_qty))})
            continue

        # 标准: 按列号精确取各尺码格的量
        for c, label in size_labels:
            qty_v = ws.cell(r, c).value
            if not is_positive_qty(qty_v):
                continue
            rows.append({**base_row, "size": label,
                         "quantity": _to_int_qty(qty_v)})

    return rows


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    # 0) PDF 委托 — 2026-08-18:用户传的可能是 PDF 而非 xlsx
    if primary_file.suffix.lower() == ".pdf":
        try:
            from vendor_pdf_vendor import _parse_panicale_pdf
            return _parse_panicale_pdf(primary_file, vendor_id)
        except ImportError:
            pass  # fallthrough to xlsx path

    # 1) 选图片目录(Mode 1 才有,Mode 2 用内嵌图)
    env_base = os.environ.get("PANICALE_IMAGES_DIR")
    image_base: Optional[Path] = None
    if env_base:
        candidate = Path(env_base)
        if candidate.is_dir():
            image_base = candidate
    else:
        for cand_name in ("PANICALE高清图-货号+色号", ""):
            candidate = primary_file.parent / cand_name if cand_name else primary_file.parent
            if candidate.is_dir():
                image_base = candidate
                break
    _pick_gender_dir(primary_file)  # 保留(可能未来日志要用),目前不参与图源路径

    # 2) 读 sheet
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    for ws in wb.worksheets:
        # 2.0) Mode 3 优先检测(SS26 新版式) — ARTICOLO/ARTICLE 块模式
        if _detect_new_layout(ws):
            rows.extend(_parse_new_layout(primary_file, ws, vendor_id))
            break
        # 2.0b) Mode 4: FW26 STOCK LIST(尺码行/颜色行 + T 列内嵌图)
        #      ⚠️ 必须排在 Mode 1 前面 —— 两者都是 R1 中文表头, Mode 1 会误判
        if _detect_stocklist_layout(ws):
            rows.extend(_parse_stocklist(primary_file, ws, vendor_id))
            break
        # 2.1) 判别 mode
        mode = _detect_mode(ws)
        if mode == "packing_list":
            rows.extend(_parse_packing_list(ws, vendor_id))
            break
        # 2.2) Mode 1: legacy 横排尺码 + 双行
        # 跨行状态:当前尺码行遗留的 size_labels
        size_labels: List[str] = []
        # 跨行状态:尺码行的 cost/cplus/retail(颜色行也带这些,但用作兜底)
        last_cost: Optional[float] = None
        last_cplus: Optional[float] = None
        last_retail: Optional[float] = None
        last_supply: Optional[float] = None

        for r in range(2, ws.max_row + 1):
            c3 = normalize_text(ws.cell(r, COL_SKU + 1).value)
            c4 = normalize_text(ws.cell(r, COL_COLOUR_CODE + 1).value)
            c6 = normalize_text(ws.cell(r, COL_DESC + 1).value)
            c8 = normalize_text(ws.cell(r, COL_COLOR_NAME + 1).value)
            c19 = normalize_text(ws.cell(r, COL_MATERIAL + 1).value)
            c20 = to_float(ws.cell(r, COL_COST + 1).value)
            c21 = to_float(ws.cell(r, COL_CPLUS_PCT + 1).value)
            c22 = to_float(ws.cell(r, COL_SUPPLY + 1).value)
            c23 = to_float(ws.cell(r, COL_RETAIL + 1).value)

            if not c3:
                continue

            # 3) 尺码行 vs 颜色行
            if not c4:
                # 尺码行:提取 c9-c18 尺码 label
                size_labels = []
                for c in range(COL_SIZE_START, COL_SIZE_END):
                    v = ws.cell(r, c + 1).value
                    sv = normalize_text(v)
                    if sv:
                        size_labels.append(sv)
                # 尺码行也记 cost/cplus/retail(后续颜色行若无,fallback)
                if c20 is not None:
                    last_cost = c20
                if c21 is not None:
                    last_cplus = c21
                if c23 is not None:
                    last_retail = c23
                if c22 is not None:
                    last_supply = c22
                continue

            # 颜色行 / 数据行
            cost = c20 if c20 is not None else last_cost
            cplus = c21 if c21 is not None else last_cplus
            retail = c23 if c23 is not None else last_retail
            supply = c22 if c22 is not None else last_supply

            # 4) 找图片
            image_items = _load_image_bytes(image_base, c3, c4)

            # 5) 按尺码展开
            if not size_labels:
                # 数据行(SOCKS/TIE 类):用 c9 数字当单尺码 qty
                for c in range(COL_SIZE_START, COL_SIZE_END):
                    qty_v = ws.cell(r, c + 1).value
                    if is_positive_qty(qty_v):
                        rows.append({
                            "brand": "PANICALE",
                            "sku": c3,
                            "image": "",
                            "description": c6,
                            "category": "",
                            "material": c19,
                            "color": c8,
                            "season": "FW26",
                            "origin": "",
                            "cost": cost,
                            "cplus_text": f"+{int(cplus * 100)}%" if cplus else "",
                            "supply_price": supply,
                            "retail_price": retail,
                            "size": SIZE_LABEL,
                            "quantity": _to_int_qty(qty_v),
                            "vendor_id": vendor_id,
                            "image_items": image_items,
                            "source_row_height": None,
                        })
            else:
                # 标准模式:尺码行 c9-c18 数字 = 各尺码 qty
                for i, label in enumerate(size_labels):
                    c = COL_SIZE_START + i
                    if c >= COL_SIZE_END:
                        break
                    qty_v = ws.cell(r, c + 1).value
                    if not is_positive_qty(qty_v):
                        continue
                    rows.append({
                        "brand": "PANICALE",
                        "sku": c3,
                        "image": "",
                        "description": c6,
                        "category": "",
                        "material": c19,
                        "color": c8,
                        "season": "FW26",
                        "origin": "",
                        "cost": cost,
                        "cplus_text": f"+{int(cplus * 100)}%" if cplus else "",
                        "supply_price": supply,
                        "retail_price": retail,
                        "size": label,
                        "quantity": _to_int_qty(qty_v),
                        "vendor_id": vendor_id,
                        "image_items": image_items,
                        "source_row_height": None,
                    })
        break  # 只处理第一个 sheet
    return rows
