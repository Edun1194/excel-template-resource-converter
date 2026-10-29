"""V1038 vendor adapter — col 14 版式通用配置器(多品牌)。

来源:A.P.C. / TORY BURCH / 等共享 col 14 块状版式的工厂现货单。

特殊点(与 V1021 / V1004 / V1012 / V1010 的差异):
- 仓库 ID 固定为 ``V1038``
- 货号识别:col 14 (N)
- 描述解析:``<类型> <颜色> <价格> <价格> <产地> <材料>``(产地在中后)
- 复用 ``convert_vendor_quote_shared.col14_parsers`` 的解析逻辑
- 品牌/性别:扫源文件顶部 R1-R20, col 4 / col 5 找 brand / gender 字符串
  (覆盖原 A.P.C. 默认 — 多品牌同版式: TORY BURCH WAREHOUSE A 2026-08-28 接入)
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

from convert_vendor_quote_shared import flat_stocklist
from convert_vendor_quote_shared.col14_parsers import parse_rows as _col14_parse

VENDOR_ID = "V1038"
DEFAULT_BRAND = "A.P.C."  # 兜底 — 源文件没扫到 brand 时用这个

# 性别关键词白名单(扫源文件 col 5 顶部,匹配上就当 gender)
GENDER_KEYS = {
    "WOMAN", "MAN", "UNISEX",
    "UOMO", "DONNA",  # 意大利语
    "KIDS", "BAMBINO", "BAMBINA",
}


def _scan_top_meta(primary_file: Path) -> Tuple[str, str]:
    """扫源文件顶部 R1-R20, 在 col 4 (D) / col 5 (E) 找 brand / gender 字符串。

    多品牌同版式: TORY BURCH WAREHOUSE A 顶部 R8 D='TORY BURCH', R10 E='WOMAN'
                  A.P.C. 顶部某行 D='A.P.C.', 某行 E='WOMAN' / 'MAN'

    Returns: (brand, gender) — 找不到就空字符串
    """
    brand = ""
    gender = ""
    try:
        wb = openpyxl.load_workbook(primary_file, data_only=True, read_only=True)
    except Exception:
        return brand, gender
    if not wb.worksheets:
        return brand, gender
    ws = wb.worksheets[0]
    for r in range(1, min(21, ws.max_row + 1)):
        v4 = ws.cell(r, 4).value
        v5 = ws.cell(r, 5).value
        if not brand and isinstance(v4, str):
            cand = v4.strip()
            # brand 特征: 全大写 + 多字 + 长度 >= 2 + 跟数字不一样
            if cand and cand == cand.upper() and len(cand) >= 2 and not cand.isdigit():
                brand = cand
        if not gender and isinstance(v5, str):
            cand = v5.strip().upper()
            if cand in GENDER_KEYS:
                gender = cand
    return brand, gender



def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1038 现货单(多品牌 col 14 版式)。

    强制把 vendor_id 改成 V1038,保证输出「云仓」列稳定是 V1038。
    品牌/性别:扫源文件顶部;扫到空就 fallback 到 DEFAULT_BRAND.
    """
    # 2026-10-28: GUCCI 库存清单平铺表版式(GUCCI FW26 - WAREHOUSE A.xlsx)
    # 版式与 col 14 完全不同, 先分流再走老路径。
    # 解析实现在 convert_vendor_quote_shared/flat_stocklist.py(V1061 也用它)。
    flat_rows = flat_stocklist.parse_rows(primary_file, VENDOR_ID)
    if flat_rows is not None:
        return flat_rows

    rows = _col14_parse(primary_file, VENDOR_ID, parser="v1038")
    src_brand, src_gender = _scan_top_meta(primary_file)
    if src_brand or src_gender:
        for r in rows:
            if src_brand and not r.get("brand"):
                r["brand"] = src_brand
            if src_gender and not r.get("gender"):
                r["gender"] = src_gender
    return rows


def get_vendor_meta() -> Dict[str, object]:
    """返回该 vendor 的默认元数据。

    - default_brand: 兜底品牌(源文件没扫到 brand 时用 A.P.C.)
    - 行高/图片缩放走 dispatcher 全局默认(92pt)
    """
    return {
        "default_brand": DEFAULT_BRAND,
    }
