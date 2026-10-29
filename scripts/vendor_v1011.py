"""V1011 MAX MARA 适配器 (意大利供应商 CATALOGO sheet 文件)。

适配文件版式 (.xlsx):
- Sheet 1 "Parametri": 5 行 1 列,参数/配置,无数据
- Sheet 2 "CATALOGO": 实际数据 sheet
  - R1-R4: size 矩阵 (4 个 size code 块)
      R1 c18='A2', c19-c29 = [3XS, 2XS, XS, S, M, L, XL, 2XL, 3XL, 4XL, 5XL]  (服饰)
      R2 c18='CA', c19-c32 = [53..66]                                     (帽码)
      R3 c18='TU', c19='TU'                                              (单尺码)
      R4 c18='V',  c19-c36 = [32..66]                                     (鞋码/腰码)
  - R5: c1=1 (size matrix 标识)
  - R6: header — c9=Costo, c10=Retail, c11=Linea, c12=Descrizione,
                    c13=Modello, c14=Variante, c15=Sesso, c16=MadeIn,
                    c17=Materiale, c18=Scalarino, c49=Totale, c50=TotalePrz
  - R7+: data (c1=2)
  - 字段直接读:
      c3=SKU 短码 / c8=SKU 短码 (重复) / c9=cost / c10=retail /
      c11=brand / c12=description / c13=full SKU / c14=color / c15=gender /
      c16=origin / c17=material / c18=size code / c19+=per-size qty /
      c49=total qty / c50=total price
  - 类目: c12 如果是 ALL CAPS (e.g. 'ABITI', 'CAPPOTTI') 当 section header;
         否则 c2 末尾 ('Donna' 之后) 拆出 section (e.g. 'ABITI', 'TOP'),
         prev_section carry-down 兜底
  - 图片: openpyxl 0-idx anchor row = r - 1 (实测: R7 anchor row=6, R9=8 ...)
          每行 2 张图 (col 0 + col 4)
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    normalize_text,
    to_float,
)


def get_vendor_meta() -> Dict[str, object]:
    return {"default_image_col_width": 40.0}


# 2026-10-13 修改: 原 V1011 写死 4 个 size groups 不够 (BURBERRY 6+, HOKA 13+),
# 改成 _build_size_label_map 动态扫 R1-R20 c18 是 size letter 的行 (c1=0)
SIZE_COL_START = 19  # 1-idx col 19 = 第一个 size label 列
MAX_MATRIX_ROW = 20   # size matrix 只在 R1-R20 (避开 data 行 R12+)


def _build_size_label_map(ws) -> Dict[str, Dict[int, str]]:
    """动态扫 R1-R20 构建 {size_code: {col_1idx: size_label}}.

    BURBERRY 之类文件有 6+ 个 size groups (A2 / AUK / FF / L / S / TU / V),
    写死 SIZE_MATRIX_ROWS 不够用。跟 V1003 一样的逻辑:
      - 只 build c1=0 的行 (避免 R12+ data 行 c18='AUK' 等被误读成 size matrix)
      - c18 是 1-6 字符的 size letter (e.g. 'A2' / 'AUK' / 'FF' / 'V' / 'TU')
      - c19+ 是 size labels (字符串)
      - 遇 c49 'Totale' 是空字符串自动 break
    """
    import re
    out: Dict[str, Dict[int, str]] = {}
    for r in range(1, MAX_MATRIX_ROW + 1):
        if ws.cell(r, 1).value != 0:
            continue
        code_raw = ws.cell(r, 18).value
        if code_raw is None:
            continue
        code = str(code_raw).strip()
        if not code or len(code) > 6:
            continue
        if not re.match(r"^[A-Z0-9]+$", code):
            continue
        col_map: Dict[int, str] = {}
        for c in range(SIZE_COL_START, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                if col_map:
                    break
                continue
            sv = str(v).strip()
            if not sv:
                if col_map:
                    break
                continue
            col_map[c] = sv
        if col_map:
            out[code] = col_map
    return out


def _load_catalogo_images(primary_file: Path) -> Dict[int, List[Dict[str, object]]]:
    """从 CATALOGO sheet 抽图,按 openpyxl 0-idx row 分桶。

    标准 load_sheet_images_by_row 只读 worksheets[0] (这里是 'Parametri' 空 sheet),
    CATALOGO sheet 有图但不是第一个,需要单独 load。

    每行有 2 张图 (col 0 主图 + col 4 副图),只取 col 0 主图避免 KEEP_IMAGES_PER_ROW=1
    限制下 col 0+col 4 被 framework 拆到不同 size 行。
    """
    out: Dict[int, List[Dict[str, object]]] = {}
    try:
        wb = openpyxl.load_workbook(str(primary_file))
    except Exception:
        return out
    if "CATALOGO" not in wb.sheetnames:
        return out
    ws = wb["CATALOGO"]
    for img in ws._images:
        anchor = img.anchor
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        # 只取 col 0 主图,跳过 col 4 副图
        if anchor._from.col != 0:
            continue
        try:
            img_bytes = img._data() if hasattr(img, "_data") else None
        except Exception:
            img_bytes = None
        if img_bytes is None:
            continue
        out.setdefault(anchor._from.row, []).append({
            "bytes": img_bytes,
            "format": img.format or "jpeg",
        })
    return out


def _parse_category(c2: str, c12: str, prev_section: str) -> tuple[str, str]:
    """决定 category + 更新 prev_section。

    策略:
      1. c12 ALL CAPS  → c12 是 section header,category = c12, prev_section = c12
      2. c2 末尾 ('Donna' 之后) 拆 section
      3. 兜底用 prev_section

    Returns: (category, new_prev_section)
    """
    if c12 and c12.isupper():
        return c12, c12
    if c2 and "Donna" in c2:
        idx = c2.find("Donna")
        section = c2[idx + 5:].strip()
        if section:
            return section, section
    return prev_section, prev_section


def parse_rows(
    primary_file: Path,
    _: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True)
    except Exception:
        return rows
    if "CATALOGO" not in wb.sheetnames:
        return rows
    ws = wb["CATALOGO"]
    if ws.max_row < 7:
        return rows

    # 1) Size 矩阵
    size_maps = _build_size_label_map(ws)

    # 2) 图片
    images_by_row = _load_catalogo_images(primary_file)

    # 3) Data rows: 2026-10-13 修复
    # 原从 R7 开始, 但 BURBERRY R7 是 header (R7 c1=0), R8 R9 是 header (c1=1),
    # R10 开始才是 data. MAX MARA 的 R7-R9 是 size matrix (R1-R4 模式), data 从 R12.
    # 改成从 R10 开始, 靠 c1=2 过滤 header
    prev_section = ""
    for r in range(10, ws.max_row + 1):
        if ws.cell(r, 1).value != 2:  # 只取 data 行
            continue

        brand = normalize_text(ws.cell(r, 11).value)
        c2 = str(ws.cell(r, 2).value or "")
        c12 = normalize_text(ws.cell(r, 12).value)
        full_sku = normalize_text(ws.cell(r, 13).value)
        color = normalize_text(ws.cell(r, 14).value)
        gender = normalize_text(ws.cell(r, 15).value)
        origin = normalize_text(ws.cell(r, 16).value)
        material = normalize_text(ws.cell(r, 17).value)
        size_code = str(ws.cell(r, 18).value or "").strip().upper()
        cost = to_float(ws.cell(r, 9).value)
        retail = to_float(ws.cell(r, 10).value)
        total_qty_v = to_float(ws.cell(r, 49).value) or 0
        total_qty = int(round(total_qty_v))

        # 类目
        category, prev_section = _parse_category(c2, c12, prev_section)

        # 货号拼接:full_sku + color
        if color:
            sku_final = f"{full_sku}-{color}"
            color_code_field = color.split()[0]
        else:
            sku_final = full_sku
            color_code_field = ""

        # 图片:openpyxl 0-idx = r - 1 (实测 CATALOGO R7 anchor=6)
        row_images = images_by_row.get(r - 1, [])

        # Size 矩阵
        col_map = size_maps.get(size_code, {})

        # 出 size 展开
        emitted = False
        for c, size_label in col_map.items():
            v = ws.cell(r, c).value
            if v is None or not isinstance(v, (int, float)) or v <= 0:
                continue
            qty = int(round(v))
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": c12,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code_field,
                "gender": gender,
                "season": "",
                "origin": origin,
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })
            emitted = True

        # 兜底: 2026-10-13 修复
        # 原 fallback 是 scalarino == 'Unica' 时出 UNI 1 行, 其它情况出 TOTAL 行
        # 但 BURBERRY 等多 size group 文件里 size_code (AUK/FF/L/S) 不在老 SIZE_MATRIX_ROWS
        # 导致 col_map={}, emitted=False, 错误地出 size='TOTAL' qty=total_qty
        # 现在改成: 只在 scalarino 真的就是 'Unica' 时出 UNI 行, 其它情况 skip
        if not emitted and scalarino == "Unica":
            rows.append({
                "brand": brand,
                "sku": sku_final,
                "image": "",
                "description": c12,
                "category": category,
                "material": material,
                "color": color,
                "color_code": color_code_field,
                "gender": gender,
                "season": "",
                "origin": origin,
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": "UNI",
                "quantity": 1,
                "vendor_id": vendor_id,
                "image_items": row_images,
            })

    return rows
