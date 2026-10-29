"""V1004-期货 期货通道 vendor adapter (2026-10-27).

与 ``V1004``(现货通道)的**唯一区别**是「数量」列的处理,其余字段/版式/图片逻辑完全一致。

期货通道统一约定(与 ``V1003-期货`` 一致,见 vendor_v1003_futures.py):
    源文件没有真实库存 → **数量列留空**,等客户填。

覆盖范围:
    1. **期货单** — SALOMON 'Foglio1' / 'Foglio3' 版式 (FW27 / AW27 订货单)
       - Foglio3 = 明细主表 (每行 颜色 × 尺码, 含 Color Code / Size)
       - Foglio1 = 款号汇总表 (cost / retail / gender / 类目 / 内嵌图片), 按货号关联补字段
       - 源文件 c7 = 'Quantity Requested' (订货量, 全部为 0 → 尚未填) → **本通道留空**
    2. **现货单** — V1004 原有的 4 个品牌
       (ARC'TERYX / PATAGONIA / POLO RALPH LAUREN / STONE ISLAND)
       → 直接委托 ``vendor_v1004.parse_rows``,数量照写源文件里的真实库存数

「云仓」列(vendor_id)统一写 ``V1004`` —— 与 ``V1003-期货`` 写 ``V1003`` 的做法一致
(期货通道沿用所属工厂的仓库号, 不用带"-期货"后缀)。

注册位置(两个地方都要改, 本文件不含这两处):
    - ``scripts/convert_vendor_quote.py`` → ``BUILTIN_VENDORS``
    - ``app.py`` → ``VENDOR_CODES``
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import vendor_v1004


VENDOR_ID = "V1004"


def _is_futures_sheet(primary_file: Path) -> bool:
    """是不是期货版式(SALOMON 'Foglio1')。"""
    if primary_file.suffix.lower() != ".xlsx":
        return False
    try:
        return bool(vendor_v1004._is_salomon_foglio(primary_file))
    except Exception:
        return False


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
    brand_override: str = "",
) -> List[Dict[str, object]]:
    """V1004-期货 解析入口.

    - 期货单(SALOMON Foglio 版式): 复用 V1004 的 SALOMON 解析, 但**数量留空**
    - 现货单(其它 V1004 品牌):   委托 V1004 现货解析, 数量照旧
    """
    # 0) 期货单: SALOMON 'Foglio1' 版式
    if _is_futures_sheet(primary_file):
        rows = vendor_v1004._parse_salomon_foglio(primary_file, VENDOR_ID)
        for r in rows:
            # 期货模版约定: 数量留空, 等客户填(源 'Quantity Requested' 全 0)
            r["quantity"] = None
        print(f"futures_mode=SALOMON qty=blank rows={len(rows)}")
        return rows

    # 1) 现货单: 委托 V1004
    rows = vendor_v1004.parse_rows(
        primary_file, conditions_file, VENDOR_ID, brand_override
    )
    print(f"spot_mode=delegated_to_v1004 rows={len(rows)}")
    return rows


def get_vendor_meta() -> Dict[str, object]:
    """沿用 V1004 的默认元数据(不设 default_brand, 走文件名兜底品牌)。"""
    return vendor_v1004.get_vendor_meta()
