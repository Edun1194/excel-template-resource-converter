"""V1010 vendor adapter — col14 parser 通用(Burberry / Gucci / Longchamp / Vivienne Westwood 等)。

来源:.xls/.xlsx 现货单。

特殊点(与 V1021 / V1004 / V1012 / V1038 的差异):
- 仓库 ID 固定为 ``V1010``
- 品牌不由 vendor meta 兜底:V1010 是多品牌工厂代码(Burberry/Gucci/Longchamp/Vivienne Westwood 等),
  不写 default_brand,让 main 走 4 级(文件名去扩展名)兜底
- 货号识别:col 14 (N),10-15 位大写字母+数字(V1021 那套 POLO 规则不适用)
- 描述解析:``made in <产地> <材质> <颜色>``(产地在头)
- 复用 ``convert_vendor_quote_shared.col14_parsers`` 的解析逻辑
  (与 V1038 同源 — 同一文件版式,不同描述结构)
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from convert_vendor_quote_shared.col14_parsers import parse_rows as _col14_parse

VENDOR_ID = "V1010"


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1010 现货单(Burberry / Gucci / Longchamp / Vivienne Westwood 等)。

    强制把 vendor_id 改成 V1010,保证输出 仓库ID 列稳定是 V1010。
    """
    return _col14_parse(primary_file, VENDOR_ID, parser="v1010")


def get_vendor_meta() -> Dict[str, object]:
    """返回该 vendor 的默认元数据。

    V1010 是多品牌工厂代码(Burberry / Gucci / Longchamp / Vivienne Westwood 等),
    不设 default_brand — main 5 级兜底中第 3 级跳过,自动走第 4 级(文件名去扩展名),
    e.g. "Longchamp.xlsx" → "Longchamp","Burberry.xlsx" → "Burberry"。
    行高/图片缩放走 dispatcher 全局默认(92pt)— Longchamp 是 .xls 无图,不影响
    """
    return {}
