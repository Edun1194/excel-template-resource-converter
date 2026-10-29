#!/usr/bin/env python3
"""
Adapter skeleton for vendor VXXXX.

复制成 `scripts/vendor_vXXXX.py`,然后:
  1. 在 `scripts/convert_vendor_quote.py` 的 BUILTIN_VENDORS 注册
     "VXXXX": "vendor_vXXXX"
  2. 按 SKILL.md 的「Adaptation Workflow(新供应商 5 步)」填实现

dispatcher 要求的接口(改了就接不上):
  - parse_rows(primary_file, markup_file, vendor_id) -> List[Dict]
  - get_vendor_meta() -> Dict          (可选)

每行 row 必须有的字段(缺了对应模板列就空):
  brand, sku, image, description, category, material, gender, color,
  color_code, season, origin, cost, cplus_text, supply_price,
  retail_price, size, quantity, vendor_id, image_items

⚠️ C+ 不要在这里写死 —— C+ 是「每单单独给」的,由调用方 --cplus / 规则注入。
   没有 C+ 就写 "cplus_text": "" 和 "supply_price": None(留空,不要填 0)。
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

VENDOR_ID = "VXXXX"
BRAND = "VENDOR NAME"


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析供应商文件 → 行列表(一行一个尺码)。"""
    rows: List[Dict[str, object]] = []

    # TODO 1: 读源文件(按表头名定位列,绝不写死列号)
    # TODO 2: 取图片(见 references/patterns/image-handling.md §8-§10 的 4 个加载器)
    # TODO 3: 尺码矩阵展开成多行
    # TODO 4: 拼 row

    return rows


def get_vendor_meta() -> Dict[str, object]:
    """vendor 级设置(图片列宽、每行保留几张图等)。"""
    return {
        "default_image_col_width": 40.0,
        "default_keep_images_per_row": 1,
    }
