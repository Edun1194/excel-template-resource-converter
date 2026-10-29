"""V1015R PRADA ITALY 补货单 vendor adapter (2026-07-29 新建)

适配文件:Prada 意大利线的 *补货单* (Stocklist Restock)。

⚠️ 当前实现基于 ATS VW (Vivienne Westwood) 的 SharePoint 同步模式 (richValue):
- 源文件从 Microsoft 365 / SharePoint 同步,带 `xl/richData/` 目录
- E 列 cell 有 `vm` 属性,通过 richValue 链拿到 Excel 实际显示的图
- 走 *per-row* 模式:每行用自己的图(不按 SKU 前缀共享)

字段映射(参考 ATS VW V1015 现货单,1-idx 列):
  货号 (SKU)   ← c4  Material Code (完整 SKU,c3 + 色码后缀)
  描述         ← c6  Style Des
  颜色         ← c9  Colour Des
  季节         ← c1  Prod Creation Season
  材质         ← c8  Prod Fabric Des
  尺码         ← c11 Prod Size
  库存         ← c12 Available Qty
  cost         ← c13 WHL (warehouse price,作 cost)
  图片         ← richValue (Excel 实际显示,小 png 84x112)
  品牌         ←  ""  留空(让 main 走文件名兜底;V1015 是多品牌工厂代码)
  云仓          ←  ""  留空

参考:references/patterns/image-handling.md §15 (richValue 机制)
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

from convert_vendor_quote_shared import (
    has_richvalue_images,
    load_richvalue_images,
    normalize_text,
    to_float,
)


# 仓库 ID 强制为 V1015(2026-07-31 用户要求),即使 vendor code 是 V1015-补货单 / V1015-补货单2
# 工厂是 V1015,补货单只是子类型
VENDOR_ID = "V1015"


def get_vendor_meta() -> Dict[str, object]:
    """V1015R PRADA ITALY 补货单 vendor 元数据。"""
    return {
        "default_image_col_width": 40.0,
    }


# 0-indexed 关键列位置(基于 ATS VW V1015 现货单版式,1-idx - 1)
COL_SKU = 3         # c4 Material Code (完整 SKU)
COL_DESC = 5        # c6 Style Des
COL_MATERIAL = 7    # c8 Prod Fabric Des
COL_COLOR = 8       # c9 Colour Des
COL_SIZE = 10       # c11 Prod Size
COL_QTY = 11        # c12 Available Qty
COL_COST = 12       # c13 WHL
COL_SEASON = 0      # c1 Prod Creation Season


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
    # 用模块级 VENDOR_ID 覆盖 vendor_id,保证「云仓」列稳定是 "V1015"
    vendor_id = VENDOR_ID

    # richValue 图(per-row,1-idx)
    if has_richvalue_images(primary_file):
        rich_imgs = load_richvalue_images(primary_file)
    else:
        rich_imgs = {}

    wb = openpyxl.load_workbook(primary_file, data_only=True)
    for ws in wb.worksheets:
        # 当前假定单 sheet
        for r in range(2, ws.max_row + 1):  # 跳过 R1 header
            sku = normalize_text(ws.cell(r, COL_SKU + 1).value)
            desc = normalize_text(ws.cell(r, COL_DESC + 1).value)
            mat = normalize_text(ws.cell(r, COL_MATERIAL + 1).value)
            color = normalize_text(ws.cell(r, COL_COLOR + 1).value)
            size = normalize_text(ws.cell(r, COL_SIZE + 1).value)
            qty_v = ws.cell(r, COL_QTY + 1).value
            season = normalize_text(ws.cell(r, COL_SEASON + 1).value)
            cost_v = ws.cell(r, COL_COST + 1).value
            if not sku or not size:
                continue
            qty = to_float(qty_v)
            if qty is None or qty <= 0:
                continue
            cost = to_float(cost_v)
            # per-row 图(richValue 1-idx row → 直接 r)
            img_items = rich_imgs.get(r, [])
            rows.append({
                "brand": "",  # 留空,让 main 走文件名兜底
                "sku": sku,
                "image": "",
                "description": desc,
                "category": "",
                "material": mat,
                "color": color,
                "season": season,
                "origin": "ITALY",
                "cost": cost,
                "cplus_text": "",
                "supply_price": None,
                "retail_price": None,
                "size": size,
                "quantity": int(qty),
                "vendor_id": vendor_id,
                "image_items": img_items,
                "source_row_height": None,
            })
        break  # 只处理第一个 sheet
    return rows
