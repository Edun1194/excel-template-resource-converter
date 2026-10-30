"""category-futures: SALOMON FTW 版式.

sheet 'FTW', 结构比较怪:
  R1 = 表头行, 但 c14 起是**尺码换算表** (两行):
        R1: 36 / 36.6666666666 / 37.3333333333 / 38 …   (欧码, 只是换算参考)
        R2: 3.5 / 4 / 4.5 / 5 …                      (英码, **这才是 SIZE**)
  R2 = 尺码标签行
  R3 = 空
  R4+ = 数据行 (c1='x' 标记, c2=Category BU P, c5=Article code, c9=Color,
                 c12=WHLS €, c13=Model Euro Re, c33=Unit total, c34=Value total)

⚠️ 2026-10-30 修复: **尺码区(c14-c32) 之前被整个丢掉了**(老代码只取 c33 Unit total, size 写死 "")。
   源文件里每个 SKU 一行, c14-c32 是"每个尺码"的**订货格子**(标尺矩阵)。

⚠️ 2026-10-30 二次调整(用户决定): **这批全是期货订单 → 忽略数量, 尺码全部展开, 数量留空**。
   - 源里 152 个 SKU 中 17 个填了数量(尺码合计精确等于 c33 Unit total: 75=75/130=130/40=40),
     另外 135 个的尺码区整行全空 —— 但期货单的语义是**全尺码供应, 数量等客户填**,
     所以统一处理: **每个 SKU × 全部 19 个尺码各出一行, 数量一律留空**。
   - 结果: 152 行 → **2888 行**(152 SKU × 19 尺码), 数量合计为空。

尺码标签取 R2(英码 3.5/4/4.5…)而不是 R1(欧码 36/36.67…):
   R1 是**欧码换算参考行**, R2 才是真正的 SIZE —— 与 V1004 SALOMON Foglio 版式一致
   (那边 Foglio1 的 R4 是欧码换算表、R7 才是 3.5/4/4.5 的尺码标签行)。

图片: anchor 在 c10 'Article image' 列; 同一 SKU 拆出的多行共用这张图。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple

from convert_vendor_quote_shared import normalize_text, to_float

from .common import (
    cellv,
    extract_images,
    find_header_row,
    header_col_map,
    is_size_like,
    norm_header,
    open_safe_workbook,
)


ALIASES: Dict[str, tuple] = {
    "sku":       ("Article code",),
    "name":      ("Article name",),
    # ⚠️ 2026-10-30 修复: 别名原来写成**截断版**, 跟源表头对不上, 导致
    #    类目/零售价/数量/Value 全是空的。源表头实际是:
    #      'Category BU Perimeter' / 'Model Euro Retail Price €'
    #      'Unit total (10 pcs min)' / 'Value total (€)'
    #    (旧别名保留, 兼容别的导出变体)
    "category":  ("Category BU Perimeter", "Category BU P", "Category BU"),
    "color":     ("Color",),
    "cost":      ("WHLS €", "WHLS"),
    "retail":    ("Model Euro Retail Price €", "Model Euro Retail Price",
                  "Model Euro Re", "Model Euro Retail"),
    "units":     ("Unit total (10 pcs min)", "Unit total"),
    "value":     ("Value total (€)", "Value total"),
    "image":     ("Article image",),
}

SIGNATURE = ("Article code", "Color", "WHLS")

# 尺码区右边界: 表头里这些词开头的列是汇总列, 不是尺码
_SIZE_STOP_PREFIXES = ("unittotal", "valuetotal")


def looks_like_salomon(ws) -> bool:
    hdr = find_header_row(ws, SIGNATURE, max_scan=6)
    if hdr is None:
        return False
    return "sku" in header_col_map(ws, hdr, ALIASES)


def _size_columns(ws, hdr: int, COL: Dict[str, int]) -> List[Tuple[int, str]]:
    """尺码区列 → [(列号, 尺码标签)]。

    尺码区 = 最后一个固定字段列之后 ~ Unit total 列之前。
    右边界**不依赖 ALIASES 匹配上**(表头变体多, 匹配不稳), 而是直接扫表头行,
    找第一个以 'unittotal' / 'valuetotal' 开头的列。
    标签优先取**尺码标签行**(R2 = 表头行下一行, 英码), 取不到再退回表头行(R1 欧码)。
    """
    fixed = [c for f, c in COL.items() if f not in ("units", "value")]
    if not fixed:
        return []
    start = max(fixed) + 1
    end = ws.max_column
    for c in range(start, min(ws.max_column, 120) + 1):
        low = norm_header(ws.cell(hdr, c).value)
        if any(low.startswith(p) for p in _SIZE_STOP_PREFIXES):
            end = c - 1
            break
    if end < start:
        return []
    out: List[Tuple[int, str]] = []
    for c in range(start, end + 1):
        label = ""
        for r in (hdr + 1, hdr):          # 先尺码标签行(R2), 再表头行(R1)
            v = ws.cell(r, c).value
            if v is not None and str(v).strip():
                label = str(v).strip()
                break
        if label and is_size_like(label):
            out.append((c, label))
    return out


def parse(primary_file: Path, vendor_id: str) -> List[Dict[str, object]]:
    wb = open_safe_workbook(primary_file)
    try:
        ws = None
        for sn in wb.sheetnames:
            cand = wb[sn]
            if find_header_row(cand, SIGNATURE, max_scan=6) is not None:
                ws = cand
                break
        if ws is None:
            return []

        hdr = find_header_row(ws, SIGNATURE, max_scan=6)
        if hdr is None:
            return []
        COL = header_col_map(ws, hdr, ALIASES)
        if "sku" not in COL:
            return []

        # 数据行起点: 表头之后第一行非空的行 (R2 是尺码标签行, R3 空)
        data_start = hdr + 1
        for r in range(hdr + 1, min(hdr + 6, ws.max_row) + 1):
            if normalize_text(cellv(ws, r, COL["sku"])):
                data_start = r
                break

        size_cols = _size_columns(ws, hdr, COL)
        images = extract_images(ws)
        rows: List[Dict[str, object]] = []
        for r in range(data_start, ws.max_row + 1):
            sku = normalize_text(cellv(ws, r, COL.get("sku")))
            if not sku:
                continue
            base = {
                "brand": "",
                "sku": sku,
                "image": "",
                "description": normalize_text(cellv(ws, r, COL.get("name"))),
                "category": normalize_text(cellv(ws, r, COL.get("category"))),
                "material": "",
                "color": normalize_text(cellv(ws, r, COL.get("color"))),
                "color_code": "",
                "gender": "",
                "season": "",
                "origin": "",
                "cost": to_float(cellv(ws, r, COL.get("cost"))),
                "cplus_text": "",
                "supply_price": None,
                "retail_price": to_float(cellv(ws, r, COL.get("retail"))),
                "vendor_id": vendor_id,
                "image_items": images.get(r, []),
            }
            # 期货订单: 忽略源里的数量, **每个 SKU × 全部尺码各出一行, 数量留空等客户填**
            if size_cols:
                for _c, label in size_cols:
                    row = dict(base)
                    row["size"] = label
                    row["quantity"] = None
                    rows.append(row)
            else:
                # 兜底: 源里没识别出尺码列 → 一行, 尺码留空
                row = dict(base)
                row["size"] = ""
                row["quantity"] = None
                rows.append(row)
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
