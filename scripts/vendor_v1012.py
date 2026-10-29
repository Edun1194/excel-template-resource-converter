"""V1012 vendor adapter — BRUNA ROSSO 通用配置器(适配 Gucci / Hogan)。

来源:从 V1021 拆分出来的子配置器,处理 BRUNA ROSSO 体系下非 POLO 的品牌
(Gucci、Hogan 等)。版式与 V1021 完全一致,只是去掉 V0000 的 92pt + 品牌默认。

行为:
- 仓库 ID 固定为 ``V1012``
- 品牌/行高走用户输入(不强制默认值)
- 复用 V1021 的 SKU/颜色/尺码/图片识别逻辑
- .xls 文件暂不抽图(.xls 是 OLE2 老格式,openpyxl 不支持),
  V1021 / V0000 同此行为
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import vendor_v1021

# 固定仓库 ID
VENDOR_ID = "V1012"


def parse_rows(
    primary_file: Path,
    conditions_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析 V1012 现货单(Gucci / Hogan 等 BRUNA ROSSO 品牌)。

    完全复用 V1021 的 parse_rows,强制 vendor_id = V1012。
    """
    return vendor_v1021.parse_rows(primary_file, conditions_file, VENDOR_ID)


def get_vendor_meta() -> Dict[str, object]:
    """V1012 没有强制默认值,所有覆盖都通过 CLI 显式传。"""
    return {}
