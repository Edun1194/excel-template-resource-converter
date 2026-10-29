"""V1062 vendor adapter — CANADA GOOSE WAREHOUSE C 平铺库存清单 + GUCCI 老 col14 现货单。

⚠️ 2026-10-28 版式扩容: 本 vendor 现在管**两种版式**, 按文件内容自动分流:

1. **平铺库存清单**(意大利语表头, 2026-10-28 接入)
   - 来源: ``CANADA GOOSE FW26 - WAREHOUSE C.xlsx`` (单 sheet 'Stock')
   - 结构跟 GUCCI FW26 - WAREHOUSE A/B 完全一样(同一套意大利库存系统):
     R1 标题 → 品牌+季节; R3 表头(Foto | Modello/Variante | Gruppo | Descrizione
     | Colore | MADE IN | Acquisto € | Vendita € | TOTALE | <尺码列>);
     R4+ 数据行与段落标题行交替
   - 解析复用 ``convert_vendor_quote_shared/flat_stocklist.py``
     (跟 V1038 GUCCI WAREHOUSE A / V1061 GUCCI WAREHOUSE B 同一个模块)

2. **col14 块状版式**(老单, 2026-08-06 接入)
   - 来源: GUCCI 工厂 V1062 WAREHOUSE B / C 的 .xlsx 现货单
   - SKU 块在 c14, 描述在 SKU 下一行(也含价格), 尺码在 c13+
   - 解析复用 ``convert_vendor_quote_shared.col14_parsers`` 的 v1038 parser

差异(V1062 vs V1038 / V1061):
- 仓库 ID: V1062
- 老 col14 路径默认品牌 GUCCI; 平铺库存清单路径的品牌从源文件标题取
  (CANADA GOOSE → 'CANADA GOOSE'), 不走 default_brand
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from convert_vendor_quote_shared import flat_stocklist
from convert_vendor_quote_shared.col14_parsers import parse_rows as _col14_parse


VENDOR_ID = "V1062"
DEFAULT_BRAND = "GUCCI"  # 只给老 col14 路径兜底


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1062 现货单 — 按文件内容自动分流两种版式。

    强制把 vendor_id 改成 V1062, 保证输出「云仓」列稳定是 V1062。
    """
    # 1) 平铺库存清单版式 (CANADA GOOSE FW26 - WAREHOUSE C 等)
    flat_rows = flat_stocklist.parse_rows(primary_file, VENDOR_ID)
    if flat_rows is not None:
        return flat_rows

    # 2) col14 块状版式 (老 GUCCI WAREHOUSE B / C)
    return _col14_parse(primary_file, VENDOR_ID, parser="v1038")


def get_vendor_meta() -> Dict[str, object]:
    """返回该 vendor 的默认元数据。

    - default_brand: GUCCI(只兜底老 col14 路径; 平铺库存清单路径的品牌来自源文件标题)
    - 行高/图片缩放走 dispatcher 全局默认(92pt)
    """
    return {
        "default_brand": DEFAULT_BRAND,
    }
