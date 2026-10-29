"""V1061 vendor adapter — GUCCI WAREHOUSE B 平铺库存清单 (2026-10-28 新建)。

来源: 工厂 V1061 发出的 ``GUCCI FW26 - WAREHOUSE B.xlsx``。

版式与 V1038 (GUCCI FW26 - WAREHOUSE A) **完全相同**(意大利语平铺库存清单),
所以解析逻辑复用 ``convert_vendor_quote_shared/flat_stocklist.py``
—— 跟 col14 版式复用 ``col14_parsers.py`` 是一个套路。

差异(V1061 vs V1038):
- **仓库 ID**: V1061(输出「云仓」列写 V1061)
- 源文件尺码列更多(148 列 vs 127 列), 多出 EU 鞋码段('14' '14+' '15++' 等)
- 半码写法: WAREHOUSE B 在 2-11 鞋码段用 ASCII 斜杠写半码('7/' '8/'),
  而同文件另一段用标准 '½'('36½' '37½') → flat_stocklist 统一归一成 '7½'

品牌/季节从源文件 R1 标题取('GUCCI — Stocklist FW26' → GUCCI / FW26),
不设 default_brand(源文件有品牌, 走源值分支)。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from convert_vendor_quote_shared import flat_stocklist


VENDOR_ID = "V1061"


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1061 (GUCCI WAREHOUSE B) 平铺库存清单。

    强制把 vendor_id 改成 V1061,保证输出「云仓」列稳定是 V1061。
    文件版式不匹配时直接报错(不静默输出空表)—— 避免选错 vendor 却看不出来。
    """
    rows = flat_stocklist.parse_rows(primary_file, VENDOR_ID)
    if rows is None:
        raise ValueError(
            "V1061 只支持 GUCCI 平铺库存清单版式"
            "(表头需同时含 'Modello/Variante' 和 'TOTALE')。当前文件不匹配 —— "
            "请确认是否选错了 vendor(GUCCI WAREHOUSE A 的平铺表用 V1038)。"
        )
    return rows


def get_vendor_meta() -> Dict[str, object]:
    """返回该 vendor 的默认元数据。

    不设 default_brand —— 品牌从源文件标题取(GUCCI)。
    行高/图片缩放走 dispatcher 全局默认(92pt)。
    """
    return {}
