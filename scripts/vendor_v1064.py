"""V1064 vendor adapter — Sirio DB 意大利多品牌 FW26 现货单(2026-08-26 新建)

适配文件:从意大利 Sirio DB 导出的 .xlsx(2 sheets: Parametri / CATALOGO)
多品牌:AMI / AMIRI / BOSS / BURBERRY / CUCINELLI / DOLCE & GABBANA / DSQUARED /
       EMPORIO ARMANI / FERRAGAMO / GIVENCHY / ISABEL MARANT / J.W. ANDERSON /
       JIMMY CHOO / KENZO / MARC JACOBS / MAX MARA / OFF WHITE / PARIS TEXAS /
       TODS / TORY BURCH / VERSACE / Y-3 / ZEGNA

跟 V1003 (GUCCI Sirio DB) **完全同版式**(Sirio 意大利系统导出),2026-08-26
用户上传 26 个 .xlsx,要 1 个 vendor code 走全。

跟 V1003 的差异:
- V1003:多品牌混合 (GUCCI / BV),不设 default_brand,走文件名兜底
- V1064:也是多品牌,但每个文件自带 Linea col 写明品牌,主走源字段

仓库 ID 固定 V1064(用户要求)
```

结构(2026-07-31 改版支持的 NEW SUPER OFFER 格式,V1003 验证过):

  - c1 全空,无 0/1/2 标识
  - HEADER label (意大利语):Costo / Retail / Linea(=品牌) / Sesso(=性别) /
                              Reparto(=类目) / Descrizione / Modello / Variante /
                              Marca / Materiale / Categoria / Scalarino
  - 货号:Modello-Variante(品牌 9 位 ID 模式 + 色码)
  - 尺码表行:HEADER 上方,Scalarino col 是单大写字母
  - DATA 行:HEADER 下方,cost col 是正数
```

直接 import V1003 的 parse_rows,把 vendor_id 改成 V1064,仓库 ID 列保持稳定。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from vendor_v1003 import parse_rows as _v1003_parse


VENDOR_ID = "V1064"


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1064 (Sirio DB 多品牌) 货单 — 委托 V1003,只改 vendor_id。"""
    return _v1003_parse(primary_file, conditions_file, VENDOR_ID)


def get_vendor_meta() -> Dict[str, object]:
    """V1064 元数据。

    - 不设 default_brand:每个文件自带 Linea col 写明品牌 (AMI / AMIRI / BOSS / ...)
    """
    return {}
