#!/usr/bin/env python3
"""
One-off converter: Tory Burch "Stock List 4199 08.25_Ita.xlsx"
一次性转换脚本 (2026-08-26 临时写,不上 BUILTIN_VENDORS)。

源文件结构(5 个 sheet 通用):
  R1: 空
  R2: "STOCKLIST as of 08.25"
  R3: 空
  R4: header
    C1=Images, C2=Product Category, C3=Product Class, C4=Product Sub Class,
    C5=Material (其实是 "<Style Code>-<NRF Color Code>" — 直接当 sku 用),
    C6=Style Code, C7=Style Description, C8=Color Description,
    C9=Material Content, C10=NRF Color Code, C11=Seasonality,
    C12=Europe Wholesale EUR (cost), C13=Europe MSRP EUR (retail),
    C14=UK MSRP GBP, C15..C(N-1)=size 列(数字或字母), CN=Total
  R5..RN: data rows

输出:
  - 1 个 xlsx 文件,内含 5 sheet,1:1 对应源 sheet
  - 18 列(品牌/货号/图片/描述/类目/材质/颜色/色号/季节/产地/cost/C+-/供货价/零售价/尺码/库存/仓库ID/性别)
  - 品牌: TORY BURCH (硬编码)
  - 货号: C5 (`<Style Code>-<NRF Color Code>` e.g. "137312-001")
  - C+: 8 (用户指定 +8%,硬编码)
  - 供货价: cost * 1.08
  - 性别: WOMAN (Tory Burch 主线)
  - 仓库 ID: 留空(没有工厂代码)
  - 季节: C11(Seasonality 字段原样输出, e.g. "CO" / "N25")
  - 产地: 空(源文件没产地列)

复用 framework utility:
  - load_images_by_row_per_sheet (读每行图片)
  - load_sheet_row_heights (读源行高)
  - embed_images_in_row (嵌入图片)
  - map_row_to_template (字段映射)
  - SKU_COL_WIDTH=26.0, KEEP_IMAGES_PER_ROW=1
"""
from __future__ import annotations

import re
import sys
import uuid
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter

# 路径:让 framework utility 可导入
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from convert_vendor_quote_shared import (  # noqa: E402
    KEEP_IMAGES_PER_ROW,
    SKU_COL_WIDTH,
    TMP_DIR,
    embed_images_in_row,
    load_images_by_row_per_sheet,
    load_sheet_row_heights,
    map_row_to_template,
)

# ---------------------------------------------------------------------------
# 硬编码常量
# ---------------------------------------------------------------------------
BRAND = "TORY BURCH"
GENDER = "WOMAN"
CPLUS = 8  # 用户指定 +8%
CPLUS_TEXT = f"+{CPLUS}"  # 显示成 "+8" 让 UI 醒目
WAREHOUSE_ID = ""  # 没有工厂代码
DEFAULT_IMAGE_ROW_HEIGHT = 92.0  # 全局默认(有图的数据行)

# 18 列模板头(中文,跟 framework 一致,2026-09-09 跟 template1-cn.xlsx 新表头同步)
CN_HEADERS = [
    "品牌", "图片", "货号", "描述", "类目", "材质", "性别",
    "颜色", "色号", "季节", "产地", "零售价", "cost", "C+-",
    "批发价", "尺码", "数量", "云仓",
]
# 列名 -> 模板中的 1-indexed 列(用于设置列宽,2026-09-09 跟新表头同步)
# 新顺序: 品牌(1) 图片(2) 货号(3) 描述(4) ... 颜色(8) 色号(9) 季节(10) 产地(11) 零售价(12) cost(13) C+-(14) 批发价(15) 尺码(16) 数量(17) 云仓(18)
COL_SKU = 3    # 货号
COL_IMAGE = 2  # 图片


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _clean_source_name(stem: str) -> str:
    """去掉 user 上传到 ~/.minimax/v2/assets/ 时被加上的 hash 前缀。

    例: '23-17-57-910-asset_20260826-231757-910_54b2e90e9769_cc3748e5-Stock List 4199 08.25_Ita'
      → 'Stock List 4199 08.25_Ita'
    """
    m = re.match(
        r"^\d{2}-\d{2}-\d{2}-\d{3}-asset_\d{8}-\d{6}-\d{3}_[0-9a-f]+_[0-9a-f]+-(.+)$",
        stem,
    )
    return m.group(1) if m else stem


def _to_float(v) -> Optional[float]:
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        if isinstance(v, str):
            s = v.strip().replace(",", ".")
            try:
                return float(s)
            except ValueError:
                return None
        return None


def _is_positive_qty(v) -> bool:
    """qty > 0 才输出。0/None/空字符串/'0' 都不算。"""
    if v is None:
        return False
    if isinstance(v, str):
        s = v.strip()
        if not s or s == "0":
            return False
    n = _to_float(v)
    return n is not None and n > 0


def _size_label(raw) -> str:
    """size 列 R4 的 header 值,可能是 int(5)/float(5.5)/str(OS/XXS)。

    5.0 归一成 '5', 5.5 保持 '5.5', 'OS'/'XXS' 等字符串原样。
    """
    if raw is None:
        return ""
    if isinstance(raw, str):
        s = raw.strip()
        return s if s else ""
    if isinstance(raw, int):
        return str(raw)
    if isinstance(raw, float):
        if raw.is_integer():
            return str(int(raw))
        return str(raw)
    s = str(raw).strip()
    return s


# ---------------------------------------------------------------------------
# parser
# ---------------------------------------------------------------------------

def _parse_sheet(
    ws,
    sheet_name: str,
    sheet_images: Dict[int, List[Dict[str, object]]],
    sheet_row_heights: Dict[int, float],
) -> List[Dict[str, object]]:
    """扫一个 sheet,R5+ 数据行 → 18 列 row dict 列表。

    每个 (Style Code, Color Code) 多个 size → 多 row(每 size 1 row, qty>0)。
    如果所有 size 都 qty=0,整个 SKU 跳过(用户要求: 库存为 0 的不放进来)。
    """
    out: List[Dict[str, object]] = []
    # 找 size 列范围:从 c15 开始向右扫,遇 "Total" 列停。
    size_col_start = 15
    size_col_end = size_col_start - 1
    for c in range(size_col_start, ws.max_column + 1):
        header = _size_label(ws.cell(4, c).value)
        if header.upper() == "TOTAL":
            break
        size_col_end = c
    if size_col_end < size_col_start:
        return out  # 没找到任何 size 列

    for r in range(5, ws.max_row + 1):
        style_code = ws.cell(r, 6).value
        nrf_code = ws.cell(r, 10).value
        if not style_code and not nrf_code:
            continue

        # 货号:C5 是 "<Style Code>-<NRF Color Code>" e.g. "137312-001"
        sku_raw = ws.cell(r, 5).value
        if sku_raw:
            sku = str(sku_raw).strip()
        else:
            # C5 缺失时,自己拼
            sc = str(style_code).strip() if style_code else ""
            nc = str(nrf_code).strip() if nrf_code else ""
            sku = f"{sc}-{nc}" if sc and nc else (sc or nc)

        description = ws.cell(r, 7).value or ""
        color = ws.cell(r, 8).value or ""
        material = ws.cell(r, 9).value or ""
        # 类目:3 段拼起来更清楚 — "Handbags / Satchels / Satchels"
        cat_parts = [
            str(ws.cell(r, 2).value or "").strip(),
            str(ws.cell(r, 3).value or "").strip(),
            str(ws.cell(r, 4).value or "").strip(),
        ]
        category = " / ".join(p for p in cat_parts if p)
        season = ws.cell(r, 11).value or ""  # Seasonality(CO/N25)
        cost = _to_float(ws.cell(r, 12).value)
        retail = _to_float(ws.cell(r, 13).value)
        # C14=UK MSRP GBP,不输出(template 零售价是 EUR,留一个就够)

        # 供货价 = cost * (1 + C+/100)
        supply_price: Optional[float] = None
        if cost is not None:
            supply_price = round(cost * (1 + CPLUS / 100.0), 4)

        # 图片:openpyxl row r (1-indexed) → 0-indexed = r-1
        image_items = sheet_images.get(r - 1, [])

        # source row height
        src_row_h = sheet_row_heights.get(r)

        # 扫所有 size 列,收集 qty>0 的
        size_rows: List[Tuple[str, int]] = []
        for c in range(size_col_start, size_col_end + 1):
            qty_raw = ws.cell(r, c).value
            if _is_positive_qty(qty_raw):
                size_label = _size_label(ws.cell(4, c).value)
                qty = int(round(_to_float(qty_raw) or 0))
                size_rows.append((size_label, qty))

        if not size_rows:
            # 全部 size 都是 0:整个 SKU 跳过(用户要求: 库存为 0 的不放进来)
            continue

        for size_label, qty in size_rows:
            out.append(
                {
                    "brand": BRAND,
                    "sku": sku,
                    "image": "",
                    "description": str(description).strip(),
                    "category": category,
                    "material": str(material).strip(),
                    "color": str(color).strip(),
                    "color_code": str(nrf_code).strip() if nrf_code else "",
                    "season": str(season).strip(),
                    "origin": "",
                    "cost": cost,
                    "cplus_text": CPLUS_TEXT,
                    "supply_price": supply_price,
                    "retail_price": retail,
                    "size": size_label,
                    "quantity": qty,
                    "vendor_id": WAREHOUSE_ID,
                    "gender": GENDER,
                    "image_items": image_items,
                    "source_row_height": src_row_h,
                }
            )
    return out


# ---------------------------------------------------------------------------
# writer
# ---------------------------------------------------------------------------

def _write_one_sheet(
    template_ws,
    rows: List[Dict[str, object]],
    image_cache: Dict[str, Path],
    tmp_dir: Path,
) -> None:
    """把 rows 写到 template_ws(已设好 header/列宽),从 ws.max_row+1 开始追加。"""
    max_height_pt = DEFAULT_IMAGE_ROW_HEIGHT
    max_height_emu = int(max_height_pt * 12700)

    # 货号列宽(SKU_COL_WIDTH)
    sku_col_letter = get_column_letter(COL_SKU)
    template_ws.column_dimensions[sku_col_letter].width = float(SKU_COL_WIDTH)

    next_row = template_ws.max_row + 1
    keep_n = KEEP_IMAGES_PER_ROW
    for r in rows:
        mapped = map_row_to_template(r, CN_HEADERS)
        image_items = r.get("image_items", [])[:keep_n]
        has_image = bool(image_items)

        if has_image:
            template_ws.row_dimensions[next_row].height = max_height_pt
        else:
            src_row_h = r.get("source_row_height")
            if src_row_h and float(src_row_h) > 0:
                template_ws.row_dimensions[next_row].height = float(src_row_h)

        for idx, value in enumerate(mapped, start=1):
            template_ws.cell(next_row, idx, value=value)

        if has_image:
            embed_images_in_row(
                template_ws, next_row, COL_IMAGE,
                image_items, image_cache, tmp_dir,
                max_height_emu=max_height_emu,
            )
        next_row += 1


# ---------------------------------------------------------------------------
# main entry
# ---------------------------------------------------------------------------

def convert(
    source_file: Path,
    template_file: Path,
    output_file: Path,
) -> Dict[str, int]:
    """主入口:source → 1 个 output xlsx(多 sheet,1:1 对应源 sheet)。

    Returns: {sheet_name: row_count}
    """
    source_file = Path(source_file).expanduser().resolve()
    template_file = Path(template_file).expanduser().resolve()
    output_file = Path(output_file).expanduser().resolve()

    if not source_file.exists():
        raise FileNotFoundError(f"Source: {source_file}")
    if not template_file.exists():
        raise FileNotFoundError(f"Template: {template_file}")

    print(f"source: {source_file}")
    print(f"template: {template_file}")
    print(f"output: {output_file}")
    print(f"brand={BRAND} gender={GENDER} cplus=+{CPLUS}% warehouse_id='{WAREHOUSE_ID}'")

    # 1) 读源
    src_wb = openpyxl.load_workbook(source_file, data_only=True)
    sheet_images_all = load_images_by_row_per_sheet(source_file)
    sheet_heights_all = load_sheet_row_heights(source_file)

    # 2) 准备 output workbook
    out_wb = openpyxl.load_workbook(template_file)

    # 删掉模板自带的所有 sheet(模板可能有 1~N 个,例如 template1-cn.xlsx 有
    # "货单模版" 和 "Sheet1" 2 个)。只保留源 sheet 1:1 对应的新 sheet。
    for tmpl_sheet in list(out_wb.sheetnames):
        if tmpl_sheet not in src_wb.sheetnames:
            del out_wb[tmpl_sheet]

    tmp_dir = TMP_DIR / uuid.uuid4().hex[:12]
    tmp_dir.mkdir(parents=True, exist_ok=True)
    image_cache: Dict[str, Path] = {}

    counts: Dict[str, int] = {}

    for sheet_name in src_wb.sheetnames:
        src_ws = src_wb[sheet_name]
        sheet_images = sheet_images_all.get(sheet_name, {})
        sheet_heights = sheet_heights_all.get(sheet_name, {})

        rows = _parse_sheet(src_ws, sheet_name, sheet_images, sheet_heights)
        counts[sheet_name] = len(rows)

        # 创建 output sheet + 写 18 列 header
        out_ws = out_wb.create_sheet(title=sheet_name)
        for c, h in enumerate(CN_HEADERS, start=1):
            out_ws.cell(1, c, value=h)

        _write_one_sheet(out_ws, rows, image_cache, tmp_dir)
        print(f"  sheet={sheet_name}: {len(rows)} rows, "
              f"sizes_used={len({r['size'] for r in rows})}, "
              f"images_total={sum(1 for r in rows if r.get('image_items'))}")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(output_file)
    print(f"\n✓ written: {output_file}")
    print(f"  total rows: {sum(counts.values())}")
    return counts


def main() -> None:
    import argparse

    ap = argparse.ArgumentParser(description="Tory Burch Stock List one-off converter")
    ap.add_argument("--source", required=True, help="Source xlsx path")
    ap.add_argument(
        "--template",
        default=str(SCRIPTS_DIR.parent / "assets" / "templates" / "template1-cn.xlsx"),
        help="Template xlsx path (default: template1-cn.xlsx)",
    )
    ap.add_argument(
        "--output",
        default=None,
        help="Output xlsx path. Default: <配置器测试文件>/TB4199/converted/<clean_stem>.xlsx",
    )
    args = ap.parse_args()

    source = Path(args.source)
    if args.output:
        output = Path(args.output).expanduser().resolve()
    else:
        # 默认: 配置器测试文件/TB4199/converted/<clean_stem>.xlsx
        clean_stem = _clean_source_name(source.stem)
        default_root = (
            Path("/Users/superlist/Desktop/表格处理/配置器测试文件") / "TB4199" / "converted"
        )
        output = default_root / f"{clean_stem}.xlsx"

    convert(source, Path(args.template), output)


if __name__ == "__main__":
    main()
