#!/usr/bin/env python3
"""
给 PANICALE FW26 男/女 xlsx 的某个 sheet 匹配本地高清图,保留原图清晰度。

逻辑 (2026-08-28 改版: 保清晰度):
  - 读指定 sheet 的 (商品国际编码, 色号) 列
  - 在对应目录找 `<国际编码> <色号>.<ext>` (3 段文件名, sku=头 1 段, color=后 2 段)
  - 找到: 用 XLImage + OneCellAnchor 嵌入, ext=(cx, cy) 设成 92pt 边长内的 EMU
          **原图 bytes 不变** → 清晰度 100% 保留
  - 找不到: 图片列留空(行还是输出)
  - 严格匹配, 不放宽变体
  - 男走 FW26 MEN, 女走 FW26 WOMEN

输出: 另存到 配置器测试文件/V1067/converted/<clean_name>_with_images_hd.xlsx
"""
from __future__ import annotations

import mimetypes
import os
import re
import sys
import uuid
from pathlib import Path
from typing import Dict, List, Optional, Tuple  # noqa: F401

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AnchorMarker,
    OneCellAnchor,
)
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter

# 注册 openpyxl 缺的 mime
for ext in [".jp2", ".jpeg2000", ".jpx", ".j2k", ".jpf"]:
    mimetypes.add_type("image/jp2", ext, strict=True)
    mimetypes.add_type("image/jp2", ext)
    try:
        mimetypes.types_map[ext] = "image/jp2"
    except Exception:
        pass
    try:
        mimetypes.types_map[True][ext] = "image/jp2"
    except (KeyError, TypeError):
        pass

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
V1067_BASE = Path("/Users/superlist/Desktop/表格处理/配置器测试文件/V1067")
IMG_DIR_MEN = V1067_BASE / "PANICALE高清图-货号+色号" / "FW26 MEN"
IMG_DIR_WOMEN = V1067_BASE / "PANICALE高清图-货号+色号" / "FW26 WOMEN"
OUT_DIR = V1067_BASE / "converted"

# 显示设置: 列宽 92 字符 + 行高 92pt + 图片显示 92pt 边长
IMG_COL_WIDTH = 92.0
IMG_ROW_HEIGHT = 92.0
IMG_MAX_PT = 92.0  # 图片最长边 92pt
# 1 pt = 12700 EMU; 1 px @ 96 DPI = 9525 EMU
PT_TO_EMU = 12700
PX_TO_EMU = 9525

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp"}


def _clean_name(stem: str) -> str:
    m = re.match(
        r"^\d{2}-\d{2}-\d{2}-\d{3}-asset_\d{8}-\d{6}-\d{3}_[0-9a-f]+_[0-9a-f]+-(.+)$",
        stem,
    )
    return m.group(1) if m else stem


def build_image_index(img_dir: Path) -> Dict[Tuple[str, str], List[Path]]:
    """扫目录,建 (sku, color) → [file_paths] 索引。"""
    idx: Dict[Tuple[str, str], List[Path]] = {}
    if not img_dir.is_dir():
        print(f"WARNING: image dir not found: {img_dir}")
        return idx
    for f in os.listdir(img_dir):
        ext = os.path.splitext(f)[1].lower()
        if ext not in SUPPORTED_EXTS:
            continue
        base = os.path.splitext(f)[0]
        parts = base.split(" ")
        if len(parts) == 2:
            sku, color = parts
        elif len(parts) == 3:
            sku, color = parts[0], " ".join(parts[1:])
        else:
            continue
        idx.setdefault((sku, color), []).append(img_dir / f)
    return idx


def _load_image_keep_bytes(path: Path) -> Optional[Tuple[XLImage, int, int, int]]:
    """读原图,**不改 bytes**,返回 (XLImage, 原宽 px, 原高 px, 原文件 bytes)。

    openpyxl 内部会用 PIL 读 size; 不重编码。
    """
    try:
        original_bytes = path.stat().st_size
        img = XLImage(str(path))
        # openpyxl 已经从 PIL 读到了 w/h (px)
        w_px, h_px = img.width, img.height
        return img, w_px, h_px, original_bytes
    except Exception as e:
        print(f"  WARN: failed to load {path.name}: {e}")
        return None


def _compute_display_emu(w_px: int, h_px: int, max_pt: float) -> Tuple[int, int]:
    """算显示用的 (cx, cy) EMU: 保持 aspect, 最长边 ≤ max_pt。

    关键: cx,cy 改成 EMU 数字,但原图 bytes 不变 → Excel 渲染时高质量缩放。
    """
    max_emu = int(max_pt * PT_TO_EMU)
    cx_orig = w_px * PX_TO_EMU
    cy_orig = h_px * PX_TO_EMU
    if max(cx_orig, cy_orig) <= max_emu:
        return cx_orig, cy_orig
    # 按比例缩
    if cx_orig > cy_orig:
        cx = max_emu
        cy = int(cy_orig * max_emu / cx_orig)
    else:
        cy = max_emu
        cx = int(cx_orig * max_emu / cy_orig)
    return cx, cy


def process_file(
    src_xlsx: Path,
    sheet_name: str,
    sku_col: int,
    color_col: int,
    img_col: int,
    img_idx: Dict[Tuple[str, str], List[Path]],
    out_xlsx: Path,
    label: str,
) -> Dict[str, int]:
    """处理 1 个 xlsx 的指定 sheet, 在 img_col 列嵌入图。

    sku_col / color_col / img_col 都是 1-indexed。
    """
    print(f"\n--- {label} ({sheet_name}) ---")
    print(f"src: {src_xlsx.name}")
    print(f"out: {out_xlsx}")

    wb = openpyxl.load_workbook(src_xlsx)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"{sheet_name} not in {src_xlsx.name}, sheets: {wb.sheetnames}")
    ws = wb[sheet_name]

    # header
    ws.cell(1, img_col, value="图片")
    col_letter = get_column_letter(img_col)
    ws.column_dimensions[col_letter].width = IMG_COL_WIDTH

    hit = miss = embed_failed = 0
    total_bytes = 0
    sample_miss = []

    for r in range(2, ws.max_row + 1):
        sku = ws.cell(r, sku_col).value
        color = ws.cell(r, color_col).value
        if not sku or not color:
            continue
        key = (sku, str(color))
        files = img_idx.get(key, [])
        if not files:
            miss += 1
            if len(sample_miss) < 5:
                sample_miss.append((r, sku, color))
            continue

        path = files[0]
        result = _load_image_keep_bytes(path)
        if result is None:
            embed_failed += 1
            continue
        img, w_px, h_px, orig_bytes = result
        cx_emu, cy_emu = _compute_display_emu(w_px, h_px, IMG_MAX_PT)

        # 用 OneCellAnchor 精确控制 cx/cy
        img.anchor = OneCellAnchor(
            _from=AnchorMarker(
                col=img_col - 1, row=r - 1,
                colOff=0, rowOff=0,
            ),
            ext=XDRPositiveSize2D(cx=cx_emu, cy=cy_emu),
        )
        ws.add_image(img)
        ws.row_dimensions[r].height = IMG_ROW_HEIGHT
        hit += 1
        total_bytes += orig_bytes

    total = hit + miss
    print(f"  total data rows: {total}")
    print(f"  image matched:   {hit} ({hit*100/total:.1f}%)" if total else "")
    print(f"  image miss:      {miss} ({miss*100/total:.1f}%)" if total else "")
    if embed_failed:
        print(f"  embed failed:    {embed_failed}")
    if hit:
        avg_pt_w = IMG_MAX_PT  # 显示尺寸固定
        print(f"  display size:  ≤{avg_pt_w}pt (longest edge), 原图 bytes 不变")
        print(f"  嵌入图总 bytes: {total_bytes/1024/1024:.0f} MB")
    if sample_miss:
        print(f"  miss samples: {sample_miss}")

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)
    print(f"  ✓ saved: {out_xlsx}")
    return {"hit": hit, "miss": miss, "total": total, "embed_failed": embed_failed}


def main() -> None:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--men-src", default=None, help="男士 xlsx")
    ap.add_argument("--women-src", default=None, help="女士 xlsx")
    ap.add_argument("--men-sheet", default="Sheet1", help="男士目标 sheet")
    ap.add_argument("--women-sheet", default="EstrazioneGiacenza2026082609490", help="女士目标 sheet")
    args = ap.parse_args()

    men_src = (
        Path(args.men_src).expanduser().resolve() if args.men_src
        else Path("/Users/superlist/.minimax/v2/assets/2026/08/28/17-29-29-691-asset_20260828-172929-691_e3b432fb2c39_33adb773-男士FW26.xlsx")
    )
    women_src = (
        Path(args.women_src).expanduser().resolve() if args.women_src
        else Path("/Users/superlist/.minimax/v2/assets/2026/08/28/18-42-04-038-asset_20260828-184204-038_eea453a5b0c1_931e32e4-女士FW26.xlsx")
    )

    print("Building image indexes...")
    men_idx = build_image_index(IMG_DIR_MEN)
    women_idx = build_image_index(IMG_DIR_WOMEN)
    print(f"  MEN index:  {len(men_idx)} (sku, color) keys")
    print(f"  WOMEN idx:  {len(women_idx)} (sku, color) keys")

    # 根据 sheet 名字决定列配置:
    #   Sheet1 (颜色级聚合): A=国际编码, B=色号, C=库存, D=图片(新加)
    #   EstrazioneGiacenza (size 级展开): A=品牌ID, B=国际编码, C=色号, ..., E=图片(源文件空, 我们填)
    def _cols_for(sheet_name: str) -> Tuple[int, int, int, str]:
        """返回 (sku_col, color_col, img_col, output_suffix)。"""
        if sheet_name.lower().startswith("estrazionegiacenza"):
            return 2, 3, 5, "Estrazione_hd"
        # default: Sheet1
        return 1, 2, 4, "Sheet1_hd"

    men_sku, men_color, men_img, men_suffix = _cols_for(args.men_sheet)
    women_sku, women_color, women_img, women_suffix = _cols_for(args.women_sheet)

    men_out = OUT_DIR / f"{_clean_name(men_src.stem)}_{men_suffix}.xlsx"
    women_out = OUT_DIR / f"{_clean_name(women_src.stem)}_{women_suffix}.xlsx"

    if men_src.exists():
        m_stats = process_file(
            men_src, args.men_sheet,
            sku_col=men_sku, color_col=men_color, img_col=men_img,
            img_idx=men_idx, out_xlsx=men_out, label="MEN",
        )
    else:
        print(f"SKIP men: {men_src} not found")
        m_stats = {}

    if women_src.exists():
        w_stats = process_file(
            women_src, args.women_sheet,
            sku_col=women_sku, color_col=women_color, img_col=women_img,
            img_idx=women_idx, out_xlsx=women_out, label="WOMEN",
        )
    else:
        print(f"SKIP women: {women_src} not found")
        w_stats = {}

    print("\n=== SUMMARY ===")
    print(f"  MEN:   {m_stats}")
    print(f"  WOMEN: {w_stats}")


if __name__ == "__main__":
    main()
