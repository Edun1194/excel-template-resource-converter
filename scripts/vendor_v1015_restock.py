"""V1015R V1015 补货单 vendor adapter (2026-07-29 新建, 2026-11-09 改成按表头定位)

适配文件:V1015 工厂的**补货单 / STOCKLIST**(ATS 系意大利代理发的表)。

⚠️ ATS 的表有两种列序(列数不同), 老代码写死列位 → 换个版式就全读错:
  - 老版式(PRADA 11 列):c1=Prod Creation Season, c4=Material Code, c6=Style Des,
    c8=Prod Fabric Des, c9=Colour Des, c11=Prod Size, c12=Available Qty, c13=WHL
  - 新版式(2026-11-09 VW STOCKLIST 15 列):c1=Group, c2=Product Department,
    c3=Prod Class, c4=Generic Prod Code, c5=Material Code, c6=Product Image,
    c7=Style Code, c8=Style Des, c9=Product Collection, c10=Prod Fabric Des,
    c11=Colour Des, c12=Colour Code, c13=Prod Size, c14=Available Qty, c15=WHL

→ 现在**一律按表头名定位列**(仓库铁律), 两个版式都能吃。

图片:ATS 的图是 **Excel「嵌入单元格」图片(richValue)** —— 文件里没有 drawing part,
  openpyxl 的 ws._images 读不到, 必须走 richValue 链:
    cell.vm=N → rdrichvalue.xml → richValueRel.xml → xl/media/imageN.png
  ⚠️ 图片列**按表头名定位**(Product Image), 写死列号会读错
  (老函数 load_richvalue_images 写死 E 列, 新版式的图在 F 列 → 读出 0 张)。
  参考:references/patterns/image-handling.md §15 (richValue 机制)

字段映射(按表头名, 1-idx):
  货号 (SKU)   ← Material Code        (完整 SKU, 如 '1803008O-C00AF-003')
  描述         ← Style Des
  颜色         ← Colour Des
  色号         ← Colour Code
  材质         ← Prod Fabric Des
  尺码         ← Prod Size
  库存         ← Available Qty
  cost         ← WHL (warehouse price, 作 cost)
  性别         ← Product Department   ('WOM-Woman' / 'MAN-Man' / 'UNI-Unisex')
  类目         ← Prod Class           ('GAR-Ready to Wear' 等, 原样)
  品牌         ← Product Collection   (VW 文件里是 'Vivienne Westwood')
  季节         ← Prod Selling Season  (新版式没有这列 → 留空)
  产地         ← 源没有产地列; 老 PRADA 版式沿用原来的 'ITALY'
  图片         ← richValue(按 Product Image 列)
  仓库ID       ← 固定 "V1015"(工厂代码, 补货单只是子类型)
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
from openpyxl.utils import get_column_letter

from convert_vendor_quote_shared import (
    has_richvalue_images,
    load_richvalue_images,
    load_richvalue_images_by_col,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 V1015(2026-07-31 用户要求),即使 vendor code 是 V1015-补货单 / V1015-补货单2
# 工厂是 V1015,补货单只是子类型
VENDOR_ID = "V1015"


def get_vendor_meta() -> Dict[str, object]:
    """V1015R 补货单 vendor 元数据。"""
    return {
        "default_image_col_width": 40.0,
    }


# 字段 → 表头别名(归一化后精确匹配)。老/新两个版式都覆盖。
ALIASES: Dict[str, tuple] = {
    "sku":        ("Material Code",),
    "desc":       ("Style Des", "Style Description"),
    "material":   ("Prod Fabric Des", "Prod Fabric Description"),
    "color":      ("Colour Des", "Colour Des", "Color Des", "Colour Description"),
    "color_code": ("Colour Code", "Color Code"),
    "size":       ("Prod Size",),
    "qty":        ("Available Qty",),
    "cost":       ("WHL",),
    "season":     ("Prod Selling Season", "Prod Creation Season"),
    "gender":     ("Product Department",),
    "category":   ("Prod Class",),
    "brand":      ("Product Collection",),
    "image":      ("Product Image",),
}


def _norm_header(value: object) -> str:
    """表头归一化: 小写 + 去掉非字母数字(空格/下划线/换行/€ 之类全去掉)。"""
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _find_header_row(ws) -> Optional[int]:
    """找表头行 —— 含 'Material Code' 或 'Prod Size' 这类可识别表头的行。"""
    limit = min(ws.max_row, 6)
    for r in range(1, limit + 1):
        keys = {
            _norm_header(ws.cell(r, c).value)
            for c in range(1, min(ws.max_column, 30) + 1)
        }
        if "materialcode" in keys or "prodsize" in keys:
            return r
    return None


def _header_cols(ws, header_row: int) -> Dict[str, int]:
    """{归一化表头名: 列号}; 同名取最左边。"""
    out: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        k = _norm_header(ws.cell(header_row, c).value)
        if k and k not in out:
            out[k] = c
    return out


def _pick(cols: Dict[str, int], names: tuple) -> Optional[int]:
    """按别名取列号(第一个命中的)。"""
    for n in names:
        c = cols.get(_norm_header(n))
        if c:
            return c
    return None


def _parse_department(value: object) -> str:
    """'WOM-Woman' / 'MAN-Man' / 'UNI-Unisex' → Woman / Man / Unisex。"""
    s = normalize_text(value)
    if not s:
        return ""
    tail = s.split("-")[-1].strip().lower()
    for kw in ("unisex", "woman", "man", "kids", "junior"):
        if tail == kw:
            return kw.capitalize() if kw != "unisex" else "Unisex"
    u = s.upper()
    if "UNISEX" in u:
        return "Unisex"
    if "WOMAN" in u or "DONNA" in u:
        return "Woman"
    if "MAN" in u or "UOMO" in u:
        return "Man"
    return ""


def parse_rows(
    primary_file: Path, _: Optional[Path], vendor_id: str,
) -> List[Dict[str, object]]:
    """读补货单 .xlsx,每行 → 1 row dict。

    跳过逻辑:无 SKU 或无尺码或库存 <=0 的行。
    仓库 ID 统一用 VENDOR_ID = "V1015"(2026-07-31),忽略传入 vendor_id。
    """
    rows: List[Dict[str, object]] = []
    if not primary_file.exists():
        return rows
    # 用模块级 VENDOR_ID 覆盖 vendor_id,保证仓库ID 列稳定是 "V1015"
    vendor_id = VENDOR_ID

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    try:
        ws = wb.worksheets[0]
        header_row = _find_header_row(ws)
        if header_row is None:
            return rows

        cols = _header_cols(ws, header_row)
        C: Dict[str, Optional[int]] = {
            field: _pick(cols, names) for field, names in ALIASES.items()
        }

        # 图片:richValue 链(ATS 的图是嵌入单元格图, 没有 drawing part)
        images_by_row: Dict[int, List[Dict]] = {}
        if has_richvalue_images(primary_file):
            img_col = C.get("image")
            if img_col:
                # ⚠️ 图片列按表头名定位(Product Image)—— 写死列号会读错
                images_by_row = load_richvalue_images_by_col(
                    primary_file, get_column_letter(img_col)
                )
            else:
                # 老版式没有 Product Image 列 → 退回老函数(写死 E 列)
                images_by_row = load_richvalue_images(primary_file)

        # 产地: 源文件都没有产地列。
        #   老 PRADA 版式(11 列, **没有 Product Department 列**)→ 沿用原来的固定值 'ITALY'
        #   ATS 新版式(有 Product Department, 能认出是哪个品牌)→ 源里没产地信息 → 留空
        default_origin = "" if C.get("gender") else "ITALY"

        def cellv(r: int, field: str) -> object:
            c = C.get(field)
            return ws.cell(r, c).value if c else None

        for r in range(header_row + 1, ws.max_row + 1):
            sku = normalize_text(cellv(r, "sku"))
            size = normalize_text(cellv(r, "size"))
            if not sku or not size:
                continue
            qty = to_float(cellv(r, "qty"))
            if qty is None or qty <= 0:
                continue

            brand = normalize_text(cellv(r, "brand"))
            rows.append({
                "brand": brand,          # 源有品牌列就用源里的; 没有 → 留空走文件名兜底
                "sku": sku,
                "image": "",
                "description": normalize_text(cellv(r, "desc")),
                "category": normalize_text(cellv(r, "category")),
                "material": normalize_text(cellv(r, "material")),
                "color": normalize_text(cellv(r, "color")),
                "color_code": normalize_text(cellv(r, "color_code")),
                "gender": _parse_department(cellv(r, "gender")),
                "season": normalize_text(cellv(r, "season")),
                "origin": default_origin,
                "cost": to_float(cellv(r, "cost")),
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,
                "size": size,
                "quantity": int(qty),
                "vendor_id": vendor_id,
                "image_items": images_by_row.get(r, []),
                "source_row_height": None,
            })
        return rows
    finally:
        try:
            wb.close()
        except Exception:
            pass
