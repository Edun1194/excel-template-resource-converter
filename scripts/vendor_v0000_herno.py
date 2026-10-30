"""V0000 — HERNO 现货单适配器 (2026-10-30 新增)。

源文件: HERNO FW 29-10.xlsx
  HERNO (意大利高端成衣/配饰品牌), 单 sheet 'Worksheet', R2 表头, R3+ 数据。
  实测: 300 行 × 33 列, 298 个数据行, 298 张内嵌图(每个数据行 1 张)。

版式 (一律**按表头名定位**, 绝不写死列号):

  R2 表头:
    c1   IMAGE          空列(真正的图在 drawing 里, 不读这一列)
    c2   SKU            货号, 例 'GI000023X12732_9018'
    c3   BRAND          品牌 (全是 'Herno')
    c4   PRODUCT        产品名 → 描述 (例 'Asymmetric jacket with collar')
    c5   GROUP          类目 ('Clothing' 291 / 'Accessories' 7)
    c6   SEASON         季节 (全是 'FW')
    c7   GENDER         性别 (Female 162 / Male 128 / W 6 / M 1 / UNI 1)
    c8   COLOR          颜色 (意大利语色名, 31 种)
    c9   COMPOSITION    材质 (例 '100% PA')
    c10  DESC_EN        英文长描述(营销文案, 太长, 不用)
    c11  MADE_IN_CODE   产地 (IT 184 / RO 64 / BG 14 / MD 14 / AM 12 / TR 4 / BA 4 / CN 2)
    c12  RETAIL PRICE   零售价
    c13  TOTAL RETAIL   = 零售价 × 总件数 (冗余列, 不用)
    c14  TOT.QTY        总件数 (= 尺码格合计)
    c15  TU             总件数, **只有 8 行填了** —— 见下面「8 行特例」
    c16-c33  尺码数量矩阵  表头是尺码 (XS/S/M/L/XL + 36…60 数字码)

⚠️ 源文件**没有 cost / 供货价列** → cost 和批发价一律留空(不凭空造, 不套默认值)。
   C+ 由用户每单单独给; 就算给了 C+, 没有 cost 也算不出批发价。

⚠️ 源文件**没有色号列** → 色号留空。
   SKU 末尾 4 位数字(9018 这种)**不是色号**: 102 个末尾码里有 35 个同时对应多个颜色名
   (如 '1000' → {bianco, beige, panna}), 它是款式内部编号。

「8 行特例」(供应商填了总数但没填尺码明细):
   有 8 个 SKU 的 TOT.QTY > 0 但尺码矩阵整行全空, 且这 8 行的 TU 值与 TOT.QTY
   **完全相等**(1/3/3/5/6/6/10/74, 合计 108 件)。
   跟框架惯例(qty<=0 的行跳过)不一样, 这里**保留**: 输出一行「尺码留空 + 数量=TOT.QTY」。
   理由: 数量是源里明写的, 尺码源里没写 —— 留空比整行丢掉 108 件真实库存更忠实。
   尺码列空着的行在模板里是合法的(其他 vendor 也会有尺码空的行)。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    is_positive_qty,
    normalize_text,
    to_float,
)


VENDOR_ID = "V0000"          # 云仓列(仓库 ID) — 临时占位码
BRAND = "HERNO"              # 源 BRAND 列全是 'Herno', 这里做大小写归一


# 性别归一: 源文件混用全称和缩写 ('Female' 和 'W' 都有)
GENDER_MAP = {
    "FEMALE": "Woman", "WOMAN": "Woman", "WOMEN": "Woman", "W": "Woman",
    "MALE": "Man", "MAN": "Man", "MEN": "Man", "M": "Man",
    "UNI": "Unisex", "UNISEX": "Unisex",
}


# 产地: 源是 ISO 3166-1 alpha-2 国家码 → 英文国名
# (仓库惯例用大写, 如 vendor_v1038 / vendor_v1062 里写的 "ITALY")
ORIGIN_MAP = {
    "IT": "ITALY",
    "RO": "ROMANIA",
    "BG": "BULGARIA",
    "MD": "MOLDOVA",
    "AM": "ARMENIA",
    "TR": "TURKEY",
    "BA": "BOSNIA AND HERZEGOVINA",
    "CN": "CHINA",
}


# 表头名(归一化后) → 字段。非尺码列, 用来定位尺码区起点。
HEADER_FIELDS: Dict[str, str] = {
    "SKU": "sku",
    "BRAND": "brand",
    "PRODUCT": "description",
    "GROUP": "category",
    "SEASON": "season",
    "GENDER": "gender",
    "COLOR": "color",
    "COMPOSITION": "material",
    "MADEINCODE": "origin",
    "RETAILPRICE": "retail",
    "TOTALRETAIL": "total_retail",
    "TOTQTY": "total_qty",
    "TU": "tu",
    "IMAGE": "image_col",
    "DESCEN": "desc_en",
}


def _norm(value: object) -> str:
    """表头归一化: 小写 + 去掉非字母数字(下划线/点/空格都去掉)。"""
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


# 元数据表头的归一化形式 —— 用来把尺码区跟固定字段区分开
# (HEADER_FIELDS 的 key 已经是归一化后的写法)
_META_HEADERS = {_norm(k) for k in HEADER_FIELDS}


def _find_header_row(ws) -> Optional[int]:
    """找表头行 —— 必须同时含 'SKU' 和 'BRAND' 两个表头。

    HERNO 的表头固定在 R2, 但上面可能压着一行残留值(R1 的 c14 有个 2560,
    正好是全表 TOT.QTY 合计 —— 供应商放错了位置), 所以要扫几行而不是写死 R2。
    """
    limit = min(ws.max_row, 8)
    for r in range(1, limit + 1):
        keys = {
            _norm(ws.cell(r, c).value)
            for c in range(1, min(ws.max_column, 60) + 1)
            if ws.cell(r, c).value not in (None, "")
        }
        if "sku" in keys and "brand" in keys:
            return r
    return None


def _header_map(ws, header_row: int) -> Dict[str, int]:
    """表头行 → {归一化表头名: 列号}。同名取最左边那个。"""
    out: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        k = _norm(ws.cell(header_row, c).value)
        if k and k not in out:
            out[k] = c
    return out


def _size_columns(ws, header_row: int) -> List[Tuple[int, str]]:
    """尺码列 = 表头不在 HEADER_FIELDS 里的非空列(按列号顺序)。

    HERNO 的尺码区固定在最后一段(c16-c33), 表头直接就是尺码标签
    ('XS' 'S' 'M' 'L' 'XL' '36' '38' … '60'), 单元格是**该尺码的数量**。
    """
    out: List[Tuple[int, str]] = []
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(header_row, c).value
        if raw in (None, ""):
            continue
        if _norm(raw) in _META_HEADERS:
            continue
        out.append((c, str(raw).strip()))
    return out


def _collect_images(ws) -> Dict[int, List[Dict[str, object]]]:
    """内嵌图 → {1-idx 行号: [{bytes, format}]}。

    实测 298 张图, 每张都锚在对应数据行上(anchor 行 == 数据行, 100% 命中)。
    """
    out: Dict[int, List[Dict[str, object]]] = {}
    try:
        images = list(ws._images)
    except Exception:
        return out
    for img in images:
        anchor = getattr(img, "anchor", None)
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        try:
            blob = img._data() if hasattr(img, "_data") else None
        except Exception:
            blob = None
        if not blob:
            continue
        fmt = (img.format or "jpeg").lower()
        if fmt == "jpg":
            fmt = "jpeg"
        out.setdefault(anchor._from.row + 1, []).append({
            "bytes": blob,
            "format": fmt,
        })
    return out


def looks_like_herno(ws) -> bool:
    """是不是 HERNO 版式。"""
    try:
        return _find_header_row(ws) is not None
    except Exception:
        return False


def parse_rows(
    primary_file: Path,
    _markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 HERNO 现货单 → 模板 row dict 列表。

    一个源行 = 一个 SKU, 按尺码矩阵展开成多行(每个有量的尺码一行)。
    """
    wb = openpyxl.load_workbook(primary_file, data_only=True)
    try:
        ws = wb.active
        header_row = _find_header_row(ws)
        if header_row is None:
            return []
        H = _header_map(ws, header_row)
        size_cols = _size_columns(ws, header_row)
        if not size_cols:
            return []

        col = {field: H.get(_norm(h)) for h, field in HEADER_FIELDS.items()}
        if not col.get("sku"):
            return []
        images_by_row = _collect_images(ws)

        rows: List[Dict[str, object]] = []
        no_size_breakdown = 0

        for r in range(header_row + 1, ws.max_row + 1):
            sku = normalize_text(ws.cell(r, col["sku"]).value) if col.get("sku") else ""
            if not sku:
                continue

            src_brand = normalize_text(ws.cell(r, col["brand"]).value) if col.get("brand") else ""
            # _norm 返回小写, 映射表 key 是大写 → 统一转大写再查
            gender_raw = _norm(ws.cell(r, col["gender"]).value).upper() if col.get("gender") else ""
            origin_raw = _norm(ws.cell(r, col["origin"]).value).upper() if col.get("origin") else ""

            base_row: Dict[str, object] = {
                "brand": (src_brand or BRAND).upper(),
                "sku": sku,
                "image": "",
                "description": normalize_text(ws.cell(r, col["description"]).value)
                if col.get("description") else "",
                "category": normalize_text(ws.cell(r, col["category"]).value)
                if col.get("category") else "",
                "material": normalize_text(ws.cell(r, col["material"]).value)
                if col.get("material") else "",
                "gender": GENDER_MAP.get(gender_raw, ""),
                "color": normalize_text(ws.cell(r, col["color"]).value)
                if col.get("color") else "",
                "color_code": "",          # 源文件没有色号列 → 留空
                "season": normalize_text(ws.cell(r, col["season"]).value)
                if col.get("season") else "",
                "origin": ORIGIN_MAP.get(origin_raw, origin_raw),
                "cost": None,              # 源文件没有 cost / 供货价列 → 留空
                "cplus_text": "",
                "supply_price": None,
                "retail_price": to_float(ws.cell(r, col["retail"]).value)
                if col.get("retail") else None,
                "vendor_id": vendor_id or VENDOR_ID,
                "image_items": images_by_row.get(r, []),
            }

            # 尺码矩阵 → 每个有量的尺码一行
            emitted = False
            for c_idx, size_label in size_cols:
                qty_v = ws.cell(r, c_idx).value
                if not is_positive_qty(qty_v):
                    continue
                rows.append({
                    **base_row,
                    "size": size_label,
                    "quantity": int(round(float(to_float(qty_v) or 0))),
                })
                emitted = True

            if not emitted:
                # 8 行特例: 有总件数但尺码明细整行空 → 出一行「尺码留空 + 数量=TOT.QTY」
                total_qty = to_float(ws.cell(r, col["total_qty"]).value) if col.get("total_qty") else None
                if total_qty is not None and total_qty > 0:
                    rows.append({
                        **base_row,
                        "size": "",
                        "quantity": int(round(total_qty)),
                    })
                    no_size_breakdown += 1

        if no_size_breakdown:
            print(f"  [V0000-HERNO] {no_size_breakdown} 行有总件数但没填尺码明细 → "
                  f"按「尺码留空 + 数量=TOT.QTY」输出")
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass


def get_vendor_meta() -> Dict[str, object]:
    """vendor 级设置。"""
    return {
        "default_image_col_width": 40.0,
    }
