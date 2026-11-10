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
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter

from convert_vendor_quote_shared import (
    guess_product_group,
    has_richvalue_images,
    is_positive_qty,
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


# ---------------------------------------------------------------------------
# ORDER(订货单)块式版式 (2026-11-10 新增)
# ---------------------------------------------------------------------------
# 源: "ORDER ZEGNA SS27_Jeans and Rtw 091126.xlsx"
# 老 V1015 / PRADA 式块状版式, 每个产品 10 行:
#   R(N)   产品 header 行: c1=总件数, c3=类别 token(尺码类别), c8+ = 各尺码 qty
#   R(N+2) c24 = SKU 编码
#   R(N+3) c24 = 描述, c35 = 颜色
#   R(N+5) c26='Acquisto'(成本标签) / c32='Vendita'(零售价标签)
#   R(N+6) c26=cost / c32=零售价
# 尺码标签行(文件顶部): c2 = 多段 token('AIU\nJU\nSKU\nXU'), c4-c27 = 尺码(也是多段)
#   → 按产品类别的 token 选第几段(欧码 / 英码 / IT码 / 字母码)
# 标题行: c1 = 'Uomo - ZEGNA'(性别 - 品牌)
# 图片: 每个产品一张, 锚在 (from_row = 产品 header 行 - 1, to_row = 下一产品附近) 跨整块
#
# 尺码段工具复用 vendor_v1015(同一套 ATS 逻辑, 避免两处漂移)

_ORDER_GENDER = {
    "donna": "Woman", "uomo": "Man", "woman": "Woman", "man": "Man",
    "unisex": "Unisex", "kids": "Kid", "junior": "Junior",
}


def _looks_like_order_block(ws) -> bool:
    """ORDER/PRADA 式块状版式 —— 找 'Acquisto'(成本标签)/ 'Vendita'(零售价标签)。"""
    limit = min(ws.max_row, 60)
    for r in range(1, limit + 1):
        for c in range(1, min(ws.max_column, 40) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() in ("acquisto", "vendita"):
                return True
    return False


def _load_block_images(ws) -> List[Tuple[int, int, Dict[str, object]]]:
    """块共享图 → [(from_row_1idx, to_row_1idx, {bytes, format})]。

    ⚠️ ORDER 的图**跨整个产品块**(from = 产品 header 行 - 1, to = 下一产品附近),
    只记 from 行会**错位一行** → 图对不上行(老 V1015 主适配器就栽在这)。
    """
    out: List[Tuple[int, int, Dict[str, object]]] = []
    for img in ws._images:
        a = getattr(img, "anchor", None)
        if not (hasattr(a, "_from") and a._from is not None):
            continue
        try:
            blob = img._data() if hasattr(img, "_data") else None
        except Exception:
            blob = None
        if not blob:
            continue
        frm = a._from.row + 1
        to = frm
        if hasattr(a, "to") and a.to is not None:
            try:
                to = a.to.row + 1
            except Exception:
                to = frm
        if to < frm:
            to = frm
        out.append((frm, to, {"bytes": blob, "format": img.format or "jpeg"}))
    return out


def _pick_block_image(
    images: List[Tuple[int, int, Dict[str, object]]], row_1idx: int,
) -> List[Dict[str, object]]:
    """取覆盖这一行的图: 最后一个 from_row <= 该行的图, 且它的范围要盖住这一行。"""
    best: Optional[Tuple[int, int, Dict[str, object]]] = None
    for frm, to, item in images:
        if frm <= row_1idx:
            best = (frm, to, item)
    if best and best[1] >= row_1idx:
        return [best[2]]
    return []


def _order_title(df, hdr_row: int) -> Tuple[str, str]:
    """向上找最近的性别标题行('Uomo - ZEGNA') → (品牌, 性别)。"""
    for r in range(hdr_row - 1, -1, -1):
        v = df.iat[r, 0] if df.shape[1] > 0 else None
        if not isinstance(v, str):
            continue
        sv = v.strip()
        if not sv:
            continue
        head = sv.split("-")[0].strip().lower()
        gender = _ORDER_GENDER.get(head, "")
        if not gender:
            continue
        brand = sv.split("-", 1)[1].strip() if "-" in sv else ""
        return brand, gender
    return "", ""


def _order_season(primary_file: Path) -> str:
    """从文件名取季节: 'ORDER ZEGNA SS27_Jeans and Rtw 091126' → 'SS27'。

    ⚠️ 季节后面常跟下划线('SS27_Jeans') —— 用 \\b 会在"数字+下划线"之间失配
       (两者都是单词字符, 没有边界), 所以用 (?![0-9]) 收尾。
    """
    stem = primary_file.stem.upper()
    m = re.search(r"(?:^|[^0-9A-Z])(FW|SS)\s*(\d{2})(?![0-9])", stem)
    return f"{m.group(1)}{m.group(2)}" if m else ""


def _parse_order_block(
    primary_file: Path, ws, vendor_id: str,
) -> List[Dict[str, object]]:
    """ORDER(订货单)块式版式 —— 老 V1015 / PRADA 那种每产品 10 行的版式。"""
    import pandas as pd
    # 尺码段解析复用主适配器(同一套 ATS 逻辑, 别写两份)
    from vendor_v1015 import (
        COLOR_COL,
        COST_VAL_COL,
        DESC_COL,
        RETAIL_VAL_COL,
        SKU_COL,
        _build_size_labels,
        _find_size_label_row,
        _looks_like_size_token,
        _pick_seg_idx,
    )

    rows: List[Dict[str, object]] = []
    df = pd.read_excel(primary_file, sheet_name=ws.title, header=None)

    size_row = _find_size_label_row(df)
    if size_row is None:
        return rows
    size_row_c2 = df.iat[size_row, 1] if df.shape[1] > 1 else ""
    if not isinstance(size_row_c2, str):
        size_row_c2 = ""

    images = _load_block_images(ws)
    season = _order_season(primary_file)

    # 找产品 header 行: c1 = 总件数(正整数), c3 = 类别 token, 且至少一个尺码列有量
    header_rows: List[int] = []
    all_labels = _build_size_labels(df, size_row)
    for idx, row in df.iterrows():
        total_qty = row.iloc[0] if len(row) > 0 else None
        cat_raw = row.iloc[2] if len(row) > 2 else None
        cat = (
            cat_raw.split("\n")[0].strip()
            if isinstance(cat_raw, str) and cat_raw.strip() else ""
        )
        if not (isinstance(total_qty, (int, float)) and total_qty > 0
                and cat and _looks_like_size_token(cat)):
            continue
        if not any(
            isinstance(row.iloc[c], (int, float)) and row.iloc[c] > 0
            for c in all_labels.keys() if c < len(row)
        ):
            continue
        header_rows.append(int(idx))

    for hdr_row in header_rows:
        hdr = df.iloc[hdr_row]
        cat_raw = hdr.iloc[2] if len(hdr) > 2 else None
        cat = (
            cat_raw.split("\n")[0].strip()
            if isinstance(cat_raw, str) and cat_raw.strip() else ""
        )
        # 按产品类别选尺码段(欧码 / 英码 / IT码 / 字母码)
        seg_idx = _pick_seg_idx(cat, size_row_c2)
        size_labels = _build_size_labels(df, size_row, seg_idx=seg_idx)
        if not size_labels:
            continue

        sku_row = hdr_row + 2
        if sku_row >= len(df):
            continue
        sku = normalize_text(df.iat[sku_row, SKU_COL]) if SKU_COL < df.shape[1] else ""
        if not sku:
            continue

        desc_row = hdr_row + 3
        if desc_row >= len(df):
            continue
        description = (normalize_text(df.iat[desc_row, DESC_COL])
                       if DESC_COL < df.shape[1] else "")
        color = (normalize_text(df.iat[desc_row, COLOR_COL])
                 if COLOR_COL < df.shape[1] else "")

        price_row = hdr_row + 6
        cost = (to_float(df.iat[price_row, COST_VAL_COL])
                if price_row < len(df) and COST_VAL_COL < df.shape[1] else None)
        retail = (to_float(df.iat[price_row, RETAIL_VAL_COL])
                  if price_row < len(df) and RETAIL_VAL_COL < df.shape[1] else None)

        brand, gender = _order_title(df, hdr_row)
        img_items = _pick_block_image(images, hdr_row + 1)

        for col_idx, size_label in size_labels.items():
            if col_idx >= len(hdr):
                continue
            qty_raw = hdr.iloc[col_idx]
            if not is_positive_qty(qty_raw):
                continue
            qty = int(round(float(to_float(qty_raw) or 0)))
            rows.append({
                "brand": brand,
                "sku": sku,
                "image": "",
                "description": description,
                "category": guess_product_group(description) if description else "",
                "material": "",
                "color": color,
                "color_code": "",
                "gender": gender,
                "season": season,
                "origin": "ITALY",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": retail,
                "size": size_label,
                "quantity": qty,
                "vendor_id": vendor_id,
                "image_items": img_items,
                "source_row_height": None,
            })
    return rows


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

        # 2026-11-10 新增: ORDER(订货单)块式版式(老 V1015 / PRADA 那种每产品 10 行)
        # ⚠️ 要在 ATS STOCKLIST 的表头判定之前 —— 两边完全不同的版式
        if _looks_like_order_block(ws):
            return _parse_order_block(primary_file, ws, vendor_id)

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
