"""V1070-期货 vendor adapter — BRUNELLO CUCINELLI (BC) **NuORDER 期货**版式 (2026-09-02 接入)。

**独立 vendor** — 不跟 V1070 现货 (FW25 / Woman new charge) 混:
- V1070 现货 → `scripts/vendor_v1070.py`(锚点 c2=SKU / c3=BRAND / c4=PRODUCT)
- V1070-期货 → `scripts/vendor_v1070_nuorder.py`(本文件,锚点 c1=Season / c2=Color / c3=Color Code / c4=Style Number)

适用:BC 通过 NuORDER 系统导出的 SS27/SS28 等期货单
- 路径示例:`配置器测试文件/V1070/BC27SS期货/SS27_WOMAN-_BAG.xlsx`

版式(NuORDER 期货):
- R1 (1-idx) = header(22 列固定 + size 组,每组 3 列 Size N/Qty N/Size price N)
  - C1=Season (e.g. "ss27"), C2=Color (lowercase), C3=Color Code,
  - C4=Style Number (完整 SKU), C5=Image, C6=Name (产品类型 e.g. "HANDBAG"),
  - C7=Wholesale EUR (cost), C8=M.S.R.P. EUR (retail, 期货多为 0),
  - C9=Division, C10=Department, C11=Category, C12=Subcategory,
  - C13=Product Notes, C14=Ship Start, C15=Ship End, C16=Prebook, C17=Resort,
  - C18=Total Price EUR, C19=Total Units,
  - C20-C22=Size 1 / Qty 1 / Size price 1, C23-C25=Size 2 / Qty 2 / Size price 2, ...
  - 最多 11 组 size(33 列),不同文件 size 数不同(BAG=1, SHOES=13, KNITWEAR=8-9, ...)
- R2+ (1-idx) = data row
- 第二 sheet 'Summary' 跳过

**期货特性**:
- retail (M.S.R.P) 通常 = 0(期货零售价未填,保留 0)
- qty 通常 = 0 或空(期货没生产出来,保留原样,**不**像现货跳 0 库存行 — 期货 catalog 就是要所有 SKU 在册)
- image 通常源文件嵌图(WPS 生成),用 zipfile+ET 手动解析
"""

from __future__ import annotations

import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import (
    normalize_text,
    to_float,
)

# 1-idx 列偏移(版式固定)
COL_SEASON = 1
COL_COLOR = 2           # color name (lowercase, e.g. "c1829")
COL_COLOR_CODE = 3
COL_STYLE = 4           # 完整 SKU (e.g. "MBXAD2571")
COL_NAME = 6            # 产品类型 (e.g. "HANDBAG")
COL_WHOLESALE = 7       # cost
COL_MSRP = 8            # retail (期货通常 0)
COL_DEPARTMENT = 10
COL_CATEGORY = 11       # 主类目 (e.g. "Bags")
COL_SUBCATEGORY = 12
COL_NOTES = 13
COL_SHIP_START = 14
COL_SHIP_END = 15
COL_TOTAL_UNITS = 19
# Size 数组:从 C20 开始,每 3 列一组 (Size N / Qty N / Size price N)
SIZE_GROUP_START = 20
SIZE_GROUP_STEP = 3


def _find_header_row(ws) -> Optional[int]:
    """找 header 行 (0-idx):扫前 5 行找 c1='Season' c2='Color' c3='Color Code' c4='Style Number'"""
    for r in range(min(5, ws.max_row)):
        c1 = normalize_text(ws.cell(row=r + 1, column=1).value).upper()
        c2 = normalize_text(ws.cell(row=r + 1, column=2).value).upper()
        c3 = normalize_text(ws.cell(row=r + 1, column=3).value).upper()
        c4 = normalize_text(ws.cell(row=r + 1, column=4).value).upper()
        if c1 == "SEASON" and c2 == "COLOR" and c3 == "COLOR CODE" and c4 == "STYLE NUMBER":
            return r
    return None


def _load_images_by_row_zip(primary_file: Path) -> Dict[int, List[Dict]]:
    """BC NuORDER 期货 — 图加载器(2026-09-02 新增)。

    NuORDER 导出的 xlsx 用 WPS 工具生成,标准 drawing XML 但 openpyxl 解析有 quirk
    (`ws._images` 返回 0,即使 zip 里有 51 个 xl/media/image*.jpeg)。
    用 zipfile + ET 手动解析 drawing1.xml + rels,按 0-idx row 分桶。

    Returns: {row_0idx -> [{bytes, format}]}
    """
    images_by_row: Dict[int, List[Dict]] = {}
    try:
        with zipfile.ZipFile(primary_file) as z:
            # 1. rels: rId -> media path (e.g. "../media/image1.jpeg")
            try:
                rels_xml = z.read("xl/drawings/_rels/drawing1.xml.rels").decode()
            except KeyError:
                return images_by_row
            rels_root = ET.fromstring(rels_xml)
            id_to_target = {r.get("Id"): r.get("Target") for r in rels_root if r.get("Id")}

            # 2. drawing1.xml: 找每个 anchor 的 (row, col) + r:embed (rel id)
            try:
                drawing_xml = z.read("xl/drawings/drawing1.xml").decode()
            except KeyError:
                return images_by_row
            drawing_root = ET.fromstring(drawing_xml)

            NS = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
            NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

            for anchor in (
                list(drawing_root.findall(f"{NS}oneCellAnchor"))
                + list(drawing_root.findall(f"{NS}twoCellAnchor"))
            ):
                from_el = anchor.find(f"{NS}from")
                if from_el is None:
                    continue
                try:
                    row = int(from_el.find(f"{NS}row").text)
                except (AttributeError, ValueError):
                    continue
                blip = anchor.find(f".//{NS_A}blip")
                embed_id = blip.get(f"{NS_R}embed") if blip is not None else None
                if not embed_id or embed_id not in id_to_target:
                    continue
                # media path: "../media/image1.jpeg" → "xl/media/image1.jpeg"
                media_target = id_to_target[embed_id].lstrip("./").lstrip("../")
                # 拼出 zip 内完整路径
                if not media_target.startswith("xl/"):
                    full_path = "xl/" + media_target
                else:
                    full_path = media_target
                try:
                    img_bytes = z.read(full_path)
                except KeyError:
                    continue
                ext = os.path.splitext(full_path)[1].lower()
                fmt = "jpeg" if ext in (".jpg", ".jpeg") else "png"
                images_by_row.setdefault(row, []).append({
                    "bytes": img_bytes,
                    "format": fmt,
                })
    except Exception:
        return images_by_row
    return images_by_row


def _build_size_groups(ws, header_row: int) -> List[Tuple[int, int, str]]:
    """扫 R1 (header_row+1) 的 size 组,返回 [(size_col, qty_col, header_label), ...]

    R1 列 20 起: 'Size 1' / 'Qty 1' / 'Size price 1' / 'Size 2' / 'Qty 2' / ...
    只收 'Size N' 开头的 col,对应下一个 col 是 Qty。
    """
    groups: List[Tuple[int, int, str]] = []
    for c in range(SIZE_GROUP_START, ws.max_column + 1, SIZE_GROUP_STEP):
        header_val = normalize_text(ws.cell(row=header_row + 1, column=c).value)
        if not header_val or not header_val.lower().startswith("size "):
            break
        # Size label 在 data row (R2+),header 里只写 "Size 1"/"Size 2"...
        # size 实际值在 R2+ 同列(0-idx row+1 那一行)
        size_col = c
        qty_col = c + 1 if c + 1 <= ws.max_column else None
        groups.append((size_col, qty_col, header_val))
    return groups


def get_vendor_meta() -> Dict[str, object]:
    """V1070 NuORDER 期货 — 给 dispatcher 用。

    - 仓库 ID 由 dispatcher 通过 vendor_id 传进来,本 adapter 强制覆盖
    - source 文件通常无图,image_col_width 不设
    - season 由源 SEASON 列(ss27/ss28)读,app 端填写可在 main 覆盖
    """
    return {}


def parse_rows(primary_file: Path, _: Optional[Path], vendor_id: str) -> List[Dict[str, object]]:
    """解析 BC NuORDER 期货。

    注意:
    - 期货 qty=0 不跳过(保留所有 SKU 在 catalog)
    - retail=0 保留原样(期货零售价待定)
    - 季节从源 SEASON 列读
    """
    rows: List[Dict[str, object]] = []

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    # 取 'NuORDER Order Data' sheet(第二个 'Summary' sheet 跳过)
    if "NuORDER Order Data" in wb.sheetnames:
        ws = wb["NuORDER Order Data"]
    elif wb.worksheets:
        ws = wb.worksheets[0]
    else:
        return rows

    header_row = _find_header_row(ws)
    if header_row is None:
        return rows

    size_groups = _build_size_groups(ws, header_row)
    if not size_groups:
        return rows

    # 强制 vendor_id 覆盖(本 adapter 专属 V1070-期货)
    actual_vendor_id = vendor_id or "V1070-期货"

    # 2026-09-02: BC NuORDER 期货图加载(用 zipfile + ET,绕开 openpyxl 解析失败)
    # image 锚点在 C5 (col 4, 0-idx),对应 0-idx row = r (跟数据行对齐)
    images_by_row = _load_images_by_row_zip(primary_file)

    for r in range(header_row + 1, ws.max_row):
        # 1-idx row = r + 1
        excel_row = r + 1

        # 6.1 SKU = Style Number (C4),必填非空
        style = normalize_text(ws.cell(row=excel_row, column=COL_STYLE).value)
        if not style:
            continue
        # 6.2 color (C2) + color code (C3) — 拼到 sku 后面(像 V1070 现货规范,用 '-' 分隔)
        color_name = normalize_text(ws.cell(row=excel_row, column=COL_COLOR).value)
        color_code = normalize_text(ws.cell(row=excel_row, column=COL_COLOR_CODE).value)
        # 货号格式: <style>-<color_code>  (color name 跟 sku 重复,不重复拼)
        sku = f"{style}-{color_code}" if color_code else style

        # 6.3 描述 = Name (产品类型) e.g. "HANDBAG"
        name = normalize_text(ws.cell(row=excel_row, column=COL_NAME).value)

        # 6.4 wholesale = cost
        cost = to_float(ws.cell(row=excel_row, column=COL_WHOLESALE).value)

        # 6.5 msrp = retail(期货可能为 0,保留原样)
        msrp = to_float(ws.cell(row=excel_row, column=COL_MSRP).value)

        # 6.6 类目 = Category(主类目) + Subcategory — 用 "/" 拼
        category_v = normalize_text(ws.cell(row=excel_row, column=COL_CATEGORY).value)
        subcategory = normalize_text(ws.cell(row=excel_row, column=COL_SUBCATEGORY).value)
        category = f"{category_v} / {subcategory}" if subcategory else category_v

        # 6.7 部门 → gender(WOMAN/MAN/UNISEX)
        # 源没有 gender 列,只从 Division 推: "Ladies Main" / "Men Main" → WOMAN/MAN
        division = normalize_text(ws.cell(row=excel_row, column=9).value)
        gender = ""
        if "LADIES" in division.upper() or "WOMAN" in division.upper() or "WOMEN" in division.upper():
            gender = "WOMAN"
        elif "MEN" in division.upper() or "UOMO" in division.upper() or "MAN" in division.upper():
            gender = "MAN"

        # 6.8 季节从源读(SS27/SS28/FW25/...)
        season = normalize_text(ws.cell(row=excel_row, column=COL_SEASON).value)

        # 6.9 产地空(源没有 origin 字段)
        # 6.10 材质空(源 Product Notes 可能含材质但太杂,不解析)
        # 6.11 notes 可作描述后缀(可选)
        notes = normalize_text(ws.cell(row=excel_row, column=COL_NOTES).value)
        # 描述 = name + notes(可选)
        description = name
        if notes and notes not in name:
            description = f"{name} - {notes}" if name else notes

        # 6.12 按 size 组展开(期货 qty=0 也输出)
        # 取本行图片(0-idx row = r,跟数据行对齐)
        row_images = images_by_row.get(r, [])

        for sc, qc, _hdr in size_groups:
            if qc is None:
                continue
            # size label 在 data row 同一 col
            size_label_raw = ws.cell(row=excel_row, column=sc).value
            size_label = normalize_text(size_label_raw) if size_label_raw is not None else ""
            if not size_label:
                # 没 size 标签 → 这行没意义,跳过
                continue
            qty_v = ws.cell(row=excel_row, column=qc).value
            # 期货 qty=0 是正常的(还没生产),保留输出,但统一写成 0
            if qty_v is None or qty_v == "":
                qty = 0
            else:
                try:
                    qty = int(round(float(qty_v)))
                except (TypeError, ValueError):
                    qty = 0

            rows.append({
                "brand": "BRUNELLO CUCINELLI",  # BC NuORDER 都是这一个品牌
                "sku": sku,
                "image": "",
                "description": description,
                "category": category,
                "material": "",  # 源 Product Notes 杂,留空
                "color": color_name.upper() if color_name else "",  # 颜色名大写
                "color_code": color_code,
                "season": season,  # ss27/SS28 → 模板「季节」列
                "origin": "",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,  # 没 C+ → 留空
                "retail_price": msrp,  # 期货多为 0
                "size": size_label,
                "quantity": qty,  # 期货 0 也输出
                "vendor_id": actual_vendor_id,
                "image_items": row_images,  # 2026-09-02:从 drawing XML 读 (WPS 工具生成)
                "source_row_height": None,
                "gender": gender,  # 从 Division 推
            })
    return rows
