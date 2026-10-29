"""V1063 vendor adapter — WAREHOUSE C 多品牌现货单(2026-08-26 新建)

来源:V1063 WAREHOUSE C .xlsx 现货单,1 sheet,块状结构。
适配品牌:GANNI / POLO RALPH LAUREN / THOM BROWNE / TORY BURCH / VIVIENNE WESTWOOD

跟 V1062 (GUCCI WAREHOUSE B/C) **同版式**(同一意大利库存系统),
2026-08-26 用户上传 5 个品牌要 1 个 vendor code 走全。

差异(V1063 vs V1062):
- V1062:单品牌 GUCCI,品牌固定
- V1063:多品牌,品牌从源文件 R3 c4 读(不设 default_brand)

仓库 ID 固定 V1063
品牌走源 R3 c4 字段(GANNI / POLO RALPH LAUREN / THOM BROWNE / TORY BURCH / VIVIENNE WESTWOOD)
性别走源 R5 c5 字段(WOMAN / MAN / UNISEX 等)
颜色:本版式 c14 只放 SKU,不存色名/色码 → 留空
描述:本版式 c14 是 "<CATEGORY>  <COST>  <QTY>" 格式,只取 category word

复用 V1062 的 col14 v1038 parser(段头/品牌/性别在固定列,SKU 块在 c14)。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared.col14_parsers import parse_rows as _col14_parse


VENDOR_ID = "V1063"


def _extract_header_fields(primary_file: Path) -> Tuple[str, str]:
    """从 WAREHOUSE C 文件顶部读 brand (R3 c4) 和 gender (R5 c5)。

    Args:
        primary_file: V1063 .xlsx 路径

    Returns:
        (brand, gender) — 任一缺时返 ""
    """
    brand = ""
    gender = ""
    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True)
        ws = wb.active
        if ws.max_row >= 3:
            v = ws.cell(3, 4).value
            if v and isinstance(v, str):
                brand = v.strip()
        if ws.max_row >= 5:
            v = ws.cell(5, 5).value
            if v and isinstance(v, str):
                gender = v.strip()
    except Exception:
        pass
    return brand, gender


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1063 (WAREHOUSE C 多品牌) 现货单 — 委托 V1062 parser。

    1. 委托 V1062 parse_rows (走 col14 v1038 parser) — 仓库 ID 强制 V1063
    2. 从源文件 R3 c4 读 brand 覆盖所有 row
    3. 从源文件 R5 c5 读 gender 覆盖所有 row
    """
    rows = _col14_parse(primary_file, VENDOR_ID, parser="v1038")
    if not rows:
        return rows

    brand, gender = _extract_header_fields(primary_file)
    for row in rows:
        if brand:
            row["brand"] = brand
        if gender:
            row["gender"] = gender
    return rows


def get_vendor_meta() -> Dict[str, object]:
    """V1063 元数据。

    - 不设 default_brand:每个文件自带 brand (R3 c4)
    - 品牌走源字段 (GANNI / POLO RALPH LAUREN / THOM BROWNE / TORY BURCH / VIVIENNE WESTWOOD)
    """
    return {}
