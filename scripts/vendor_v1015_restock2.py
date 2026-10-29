"""V1015-补货单2 MIU MIU ATS vendor adapter (2026-07-31 新建)

适配文件:MIU MIU 工厂的 ATS (Available To Sell = 可售库存) Excel。
3 个 sheet (Pelletteria / Calzature / Abbigliamento) 都要转,共用同一版式:

版式 — R6 是 header,R7+ 是产品行。

[Pelletteria / Abbigliamento 基础版 header, 6 个核心列 + 尺码]
  c1 (0-idx 0)  Periodo            ← 季节, e.g. "20261" (SS) / "20262" (FW)
  c2 (0-idx 1)  Immagine           ← 图占位列
  c3 (0-idx 2)  Cod. modello var.  ← SKU 基础码, e.g. "5BA281 V    OLI" (Pelletteria) / "GWB148 S    OOO" (Abbigliamento)
  c4 (0-idx 3)  Abbinamento        ← 描述, e.g. "2IEO - SOFTY SHINE" / "10BF - DENIM"
  c5 (0-idx 4)  Colore             ← "F0009 - BIANCO" (色号 - 色名)
  c6 (0-idx 5)  Totale             ← 总库存数字
  c7+           尺码列 (TU/65/70/.../S, 或 1/2/24/25/.../XL)

[Calzature 扩展版 header, 多 5 列(2 个品牌/类目相关 + 3 个标识列) + 尺码]
  c1 (0-idx 0)  SKU                ← 完整 18 位 SKU, e.g. "5D058EFD010Z0GF0379"
  c2 (0-idx 1)  Marchio-Coll       ← 品牌, e.g. "Miu Miu - Donna" (含性别)
  c3 (0-idx 2)  Divisione          ← 部门, e.g. "CALZATURE"
  c4 (0-idx 3)  Classe comm.       ← 商业类, e.g. "CALZATURE"
  c5 (0-idx 4)  Cat. merceologica  ← 类目, e.g. "SABOT" / "MOCASSINI" / "SNEAKERS"
  c6 (0-idx 5)  Periodo
  c7 (0-idx 6)  Immagine
  c8 (0-idx 7)  Cod. modello var.
  c9 (0-idx 8)  Abbinamento
  c10(0-idx 9)  Colore
  c11(0-idx 10) Totale
  c12+          尺码列 (34/34,5/35/.../43,有 35/36 等双码)

判别基础版/扩展版:c0 (0-idx 0) header 文本:
  - 'Periodo' → 基础版 (Pelletteria / Abbigliamento)
  - 'SKU'     → 扩展版 (Calzature)

字段映射 (基本):
  SKU        ← 扩展版: c0 完整 SKU / 基础版: 拼自 c2 (Cod) + "-" + 色号(从 Colore 拆)
  颜色       ← Colore 拆出色名 (e.g. "BIANCO" from "F0009 - BIANCO")
  色号       ← Colore 拆出色号 (e.g. "F0009" from "F0009 - BIANCO")
  类目       ← 扩展版: c4 Cat. merceologica / 基础版: guess_product_group(Abbinamento)
  品牌       ← 扩展版: c1 Marchio-Coll / 基础版: "" (留空,main 走文件名兜底)
  性别       ← 扩展版: 从 Marchio-Coll 提取 "Donna"/"Uomo"/"Man"/"Woman" / 基础版: ""
  描述       ← Abbinamento (c3 / c8)
  季节       ← Periodo (c0 / c5) 留原值
  cost       ← "" (源无 cost 列)
  零售价     ← "" (源无 retail 列)
  尺码       ← 尺码列 label (从 header 提取,strip 空格)
  库存       ← 尺码列 qty

图片:用 openpyxl ws._images,按 anchor._from.row (0-idx) 分桶,
     每行 1 张图(实测 Pelletteria 280 张、Calzature 297 张、Abbigliamento 421 张,
     都接近产品行数,确认每产品 1 图)。

3 个 sheet 全部转:parse_rows 遍历 wb.worksheets,每个 sheet 独立 parse 后合并。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    guess_product_group,
    is_positive_qty,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 V1015(2026-07-31 用户要求),即使 vendor code 是 V1015-补货单 / V1015-补货单2
# 工厂是 V1015,补货单只是子类型
VENDOR_ID = "V1015"


# R6 是 header(从 R1 数起 1-idx row=6)
HEADER_ROW = 6  # 1-idx
# 颜色字段格式: "F0009 - BIANCO" → split 出 color_code='F0009' / color='BIANCO'
_COLORE_RE = re.compile(r"^([A-Z0-9]+)\s*-\s*(.+)$")
# 性别关键字(从 Marchio-Coll 提取,意大利/英语)
_GENDER_KEYWORDS = ["Donna", "Uomo", "Woman", "Man", "Bambina", "Bambino", "Unisex", "Junior"]


def get_vendor_meta() -> Dict[str, object]:
    """V1015-补货单2 vendor 元数据 — 给 dispatcher 用。

    MIUMIU ATS 每行 1 张图(嵌入 drawing),KEEP_IMAGES_PER_ROW=1 (默认)。
    default_image_col_width=40:与 V1015 现货单一致。
    """
    return {
        "default_image_col_width": 40.0,
    }


def _parse_colore(colore_raw: str) -> Tuple[str, str]:
    """拆 "F0009 - BIANCO" → (color_code='F0009', color='BIANCO')

    拆不开就 (color_code='', color=原值)
    """
    if not colore_raw:
        return "", ""
    sv = str(colore_raw).strip()
    m = _COLORE_RE.match(sv)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", sv


def _pick_gender(marchio_coll: str) -> str:
    """从 "Miu Miu - Donna" 提取 "Donna" / "Uomo" 等。

    没匹配返回 ""。
    """
    if not marchio_coll:
        return ""
    sv = str(marchio_coll).strip()
    for kw in _GENDER_KEYWORDS:
        if kw in sv:
            return kw
    return ""


def _load_images_by_row(ws) -> Dict[int, List[Dict]]:
    """按 openpyxl 0-idx row 分桶图(每行 1 张)。"""
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
        row = anchor._from.row
        out.setdefault(row, []).append({
            "bytes": data,
            "format": img.format or "jpeg",
        })
    return out


def _detect_layout(ws) -> Tuple[str, Dict[str, int]]:
    """判别 sheet 是基础版还是扩展版,返回 (layout, col_map)

    layout: 'base' (Pelletteria/Abbigliamento) / 'extended' (Calzature)
    col_map: 字段名 → 0-idx 列位置
    """
    c0_header = normalize_text(ws.cell(row=HEADER_ROW, column=1).value).upper()

    if c0_header == "SKU":
        # 扩展版 (Calzature)
        return "extended", {
            "sku": 0,            # c1 完整 SKU
            "brand": 1,          # c2 Marchio-Coll (含品牌 + 性别)
            "category": 4,       # c5 Cat. merceologica
            "periodo": 5,        # c6 Periodo
            "cod": 7,            # c8 Cod. modello var.
            "abbinamento": 8,    # c9 Abbinamento (描述)
            "colore": 9,         # c10 Colore
            "totale": 10,        # c11 Totale
            "size_start": 11,    # c12 起尺码列
        }
    elif c0_header == "PERIODO":
        # 基础版 (Pelletteria / Abbigliamento)
        return "base", {
            "periodo": 0,        # c1 Periodo
            "cod": 2,            # c3 Cod. modello var.
            "abbinamento": 3,    # c4 Abbinamento (描述)
            "colore": 4,         # c5 Colore
            "totale": 5,         # c6 Totale
            "size_start": 6,     # c7 起尺码列
        }
    else:
        # 不识别 — 跳过
        return "unknown", {}


def _build_size_labels(ws, size_start_col: int) -> Dict[int, str]:
    """从 header 行 size_start_col+ 起提取尺码 label,跳过 None/空 cell。

    label 归一化:
      - '   34'  → '34'  (strip 空格)
      - ' 34,5'  → '34.5'  (欧洲逗号 → 句点)
      - '35/36'  → '35/36'  (双码保留,例如 Calzature 鞋码 35/36 等同 35)
    """
    out: Dict[int, str] = {}
    for c in range(size_start_col, ws.max_column):
        v = ws.cell(row=HEADER_ROW, column=c + 1).value
        if v is None:
            if out:
                break  # 已有 token 后面空 → 本行结束
            continue
        sv = str(v).strip()
        if not sv:
            if out:
                break
            continue
        # 数字字符串归一化(欧洲逗号 → 句点,去掉空格)
        if re.fullmatch(r"\d+(?:[.,]\d+)?", sv):
            sv = sv.replace(",", ".")
        out[c] = sv
    return out


def _parse_sheet(ws, sheet_name: str, vendor_id: str) -> List[Dict[str, object]]:
    """解析单个 sheet,返回 row dict 列表。"""
    rows: List[Dict[str, object]] = []
    layout, col_map = _detect_layout(ws)
    if layout == "unknown" or not col_map:
        return rows

    size_start_col = col_map["size_start"]
    size_labels = _build_size_labels(ws, size_start_col)
    if not size_labels:
        return rows

    images_by_row = _load_images_by_row(ws)
    # sheet 名兜底类别(基础版 Pelletteria/Abbigliamento 的 Abbinamento 没有品类词)
    sheet_category_fallback = sheet_name.strip() if sheet_name else ""

    # 遍历产品行 (R7 起,1-idx)
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        # 0-idx = r - 1
        zero_idx = r - 1

        # 1) SKU
        if layout == "extended":
            sku = normalize_text(ws.cell(row=r, column=col_map["sku"] + 1).value)
        else:
            # 基础版:拼 Cod + "-" + 色号
            cod = normalize_text(ws.cell(row=r, column=col_map["cod"] + 1).value)
            colore_raw = ws.cell(row=r, column=col_map["colore"] + 1).value
            color_code, _ = _parse_colore(colore_raw)
            if not cod or not color_code:
                continue
            sku = f"{cod}-{color_code}"
        if not sku:
            continue

        # 2) 描述 (Abbinamento)
        desc = normalize_text(ws.cell(row=r, column=col_map["abbinamento"] + 1).value)

        # 3) 颜色 + 色号
        colore_raw = ws.cell(row=r, column=col_map["colore"] + 1).value
        color_code, color = _parse_colore(colore_raw)

        # 4) 季节 (Periodo 留原值)
        season = normalize_text(ws.cell(row=r, column=col_map["periodo"] + 1).value)

        # 5) 品牌 + 性别
        if layout == "extended":
            brand = normalize_text(ws.cell(row=r, column=col_map["brand"] + 1).value)
            gender = _pick_gender(brand)
        else:
            brand = ""  # 基础版留空,main 走文件名兜底
            gender = ""

        # 6) 类目
        if layout == "extended":
            category = normalize_text(ws.cell(row=r, column=col_map["category"] + 1).value)
        else:
            category = guess_product_group(desc) if desc else ""
        # 类目兜底:基础版的 Abbinamento 可能不含品类词(皮具描述都是 "2IEO - SOFTY SHINE" 编码),
        # 兜底用 sheet 名(让用户至少看到 "Pelletteria" / "Abbigliamento" 标识)
        if not category:
            category = sheet_category_fallback

        # 7) 图片(0-idx row = zero_idx)
        row_images = images_by_row.get(zero_idx, [])

        # 8) 按尺码展开
        size_count = 0
        for c, size_label in size_labels.items():
            qty_v = ws.cell(row=r, column=c + 1).value
            if not is_positive_qty(qty_v):
                continue
            qty = int(round(float(to_float(qty_v) or 0)))
            if qty <= 0:
                continue
            size_count += 1
            rows.append({
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": desc,
                "category": category,
                "material": "",
                "color": color,
                "color_code": color_code,  # 中文模板"色号"列
                "gender": gender,         # 2026-07-31:扩展版从 Marchio-Coll 提取
                "season": season,
                "origin": "ITALY",
                "cost": None,             # 源无 cost 列
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
                "source_row_height": None,
            })

        # 如果一行所有尺码都没 qty,跳过(空产品行)
        # 但保留 img_items / brand / sku 后续可能用作它用,这里直接 skip
        # size_count == 0 仍可继续(用户可能没填尺码但有 total)

    return rows


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 MIUMIU ATS .xlsx,3 sheet 全部转,合并返回 row dict 列表。

    仓库 ID 统一用 VENDOR_ID = "V1015"(2026-07-31),忽略传入 vendor_id。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    # 用模块级 VENDOR_ID 覆盖 vendor_id,保证「云仓」列稳定是 "V1015"
    vendor_id = VENDOR_ID

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    for ws in wb.worksheets:
        sheet_rows = _parse_sheet(ws, ws.title, vendor_id)
        rows.extend(sheet_rows)
    return rows
