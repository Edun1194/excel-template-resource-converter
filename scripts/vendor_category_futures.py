"""category-futures 适配器 (2026-10-27 新增).

期货类文件的聚合通道: 一个适配器覆盖多个品牌/多种版式, 按文件特征自动路由。

解析实现 (每个都是独立模块, 内部一律**按表头名定位列**, 绝不写死列号):
  linesheet.py   LEMAIRE / Alexander Wang / GANNI
                 (表头行: Style Image | Style Name | Style Number | Color | …
                  尺码在 'Available Sizes' 文本列, 按逗号展开, 数量留空)
  nuorder.py     ACNE / BC / Men ss27 / POLO RL
                 (sheet 'NuORDER Order Data', 每行一个尺码, 列序各家不同)
  boss.py        BOSS  (sheet 'RTW', 尺码是列, 数量在尺码列)
  salomon.py     SALOMON (sheet 'FTW', 双行表头: R1 尺码标签 / R2 尺码数值)
  longchamp.py   LONGCHAMP (sheet 'QTE BDC' + 'BDC')
  vw.py          VIVIENNE WESTWOOD (sheet 'Foglio1', 每行一个尺码)

输出约定:
  解析逻辑与各品牌对应的一致, 唯一区别是**「云仓」列固定写「品牌名称（期货）」**
  —— 模板的「云仓」列映射的是 row["vendor_id"]
  (见 convert_vendor_quote_shared/__init__.py 的 map_row_to_template)。
"""

from __future__ import annotations

import importlib
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from category_futures import boss, common, linesheet, longchamp, nuorder, salomon, vw


# 「云仓」列固定写入的内容
CATEGORY_FUTURES_WAREHOUSE = "品牌名称（期货）"


# (模块, 检测函数) —— 顺序即优先级
_PARSERS: List[Tuple[object, Callable[[object], bool]]] = [
    (linesheet, linesheet.looks_like_linesheet),
    (nuorder,   nuorder.looks_like_nuorder),
    (boss,      boss.looks_like_boss),
    (salomon,   salomon.looks_like_salomon),
    (longchamp, longchamp.looks_like_longchamp),
    (vw,        vw.looks_like_vw),
]


def _detect(primary_file: Path) -> Optional[Tuple[object, Callable]]:
    """打开文件, 找出是哪个版式."""
    wb = common.open_safe_workbook(primary_file)
    try:
        for mod, looks in _PARSERS:
            for sn in wb.sheetnames:
                try:
                    ws = wb[sn]
                except Exception:
                    continue
                try:
                    if looks(ws):
                        return mod, ws
                except Exception:
                    continue
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return None


def parse_rows(
    primary_file: Path,
    markup_file: Optional[Path],
    vendor_id: str,
) -> List[Dict[str, object]]:
    """解析期货文件 (自动路由版式), 并把「云仓」列改成固定文本."""
    rows: List[Dict[str, object]] = []

    detected = _detect(primary_file)
    if detected is not None:
        mod, _ws = detected
        try:
            rows = mod.parse(primary_file, vendor_id) or []
        except Exception as exc:
            print(f"[category-futures] {mod.__name__} 解析失败: {exc}")
            rows = []

    if not rows:
        # 兜底: 逐个模块再试一遍 (detect 用的 sheet 未必是数据所在 sheet)
        for mod, _looks in _PARSERS:
            try:
                rows = mod.parse(primary_file, vendor_id) or []
            except Exception:
                rows = []
            if rows:
                break

    for r in rows:
        r["vendor_id"] = CATEGORY_FUTURES_WAREHOUSE
        q = r.get("quantity")
        if isinstance(q, float):
            r["quantity"] = int(round(q))     # 数量列要整数
    return rows


def get_vendor_meta() -> Dict[str, object]:
    """图片列宽等 vendor 级设置."""
    return {
        "default_image_col_width": 40.0,
    }
