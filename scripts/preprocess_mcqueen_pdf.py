"""McQueen SS27 PDF → xlsx 预处理器 (2026-09-29).

源 PDF 结构 (line sheet,1 个 product card = 7 行):
  Line 1: SKU (e.g., "A00L0H-QPABQ-4424")
  Line 2: DESCRIPTION (e.g., "HRNSS OFFICER SHIRT")
  Line 3: MATERIAL DESCRIPTION (e.g., "RESPONSIBLE WASHED SILK POP")
  Line 4: COLOR DESCRIPTION (e.g., "SOHO BLUE")
  Line 5: WHL PRICE (e.g., "EUR 876,00" 或 "EUR 0,00")
  Line 6: RTL PRICE (e.g., "EUR 1.990,00" 或 "EUR 0,00")
  Line 7: SIZE / SCALARINO CODE (e.g., "N4" 或 "T")

布局:3 个 product card 横排,多组之间用空行隔开

输出 xlsx schema (与 McQueen .xls 一致,直接喂给 V1003):
  R14: header (GENDER | CATEGORY | DESCRIPTION | SKU | MATERIAL | COLOR | SIZES | WHL PRICE | RTL PRICE | SIZE)
  R15+: data rows
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, Tuple

import openpyxl


EUR_PRICE_RE = re.compile(r"EUR\s*([\d.,]+)", re.IGNORECASE)


def _parse_card_block(lines: List[str]) -> List[Tuple[str, str, str, str, str, str, str]]:
    """Split a 3-column block of N lines into N product cards.

    Each line has up to 3 columns separated by 4+ spaces.
    Card fields are anchored by EUR price markers:
        SKU | DESC | MATL (1-2 lines if wraps) | COLOR | WHL(EUR) | RTL(EUR) | SIZE

    For each col, scan the per-col lines and locate WHL/RTL by EUR regex,
    then back-fill SKU/DESC/MATL/COLOR/SIZE from surrounding lines.
    """
    if not lines:
        return []

    # Split each line into up to 3 columns
    split_lines: List[List[str]] = []
    for line in lines:
        line_stripped = line.rstrip()
        # EUR 价格行 (如 "EUR   335,00") 不要在 EUR/数字 之间再切
        # 临时把 "EUR   number" 整个换成 placeholder,切完再换回来
        eur_marker = "__EUR_PLACEHOLDER__"
        eur_count = len(re.findall(r"EUR\s+[\d.,]+", line_stripped))
        line_for_split = re.sub(r"EUR\s+([\d.,]+)", lambda m: f"{eur_marker}{m.group(1)}", line_stripped)
        cols = re.split(r"\s{4,}", line_for_split)
        if len(cols) < 3:
            cols = re.split(r"\s{3,}", line_for_split)
        # Restore EUR placeholder
        cols = [c.replace(eur_marker, "EUR ") for c in cols]
        while len(cols) < 3:
            cols.append("")
        split_lines.append(cols[:3])

    n_cards = 3
    cards = []
    for c_idx in range(n_cards):
        # Per-column non-empty lines
        col_lines: List[str] = []
        for row in split_lines:
            if c_idx < len(row) and row[c_idx].strip():
                col_lines.append(row[c_idx])

        if not col_lines:
            continue

        # Find EUR positions
        eur_idx: List[int] = []
        for i, ln in enumerate(col_lines):
            if EUR_PRICE_RE.search(ln):
                eur_idx.append(i)

        if len(eur_idx) < 2:
            # Less than 2 EUR lines — invalid card, skip
            continue

        # WHL = first EUR, RTL = second EUR
        whl_idx = eur_idx[0]
        rtl_idx = eur_idx[1]

        # Field positions
        sku = col_lines[0]
        # DESC = col_lines[1] (single line after SKU, may be empty)
        desc = col_lines[1] if len(col_lines) > 1 else ""

        # MATL = lines between DESC and the line before WHL
        # (whl_idx - 2 is COLOR, so MATL is from 2 to whl_idx - 3 inclusive)
        matl_end = whl_idx - 2  # last MATL line index
        matl_start = 2          # first MATL line index
        if matl_end >= matl_start:
            matl = " ".join(col_lines[matl_start:matl_end + 1])
        else:
            matl = ""

        # COLOR = line just before WHL
        color = col_lines[whl_idx - 1] if whl_idx - 1 >= 0 else ""

        # WHL / RTL are the EUR lines (with EUR prefix stripped)
        whl = col_lines[whl_idx]
        rtl = col_lines[rtl_idx]

        # SIZE = line after RTL (if exists and doesn't look like a wrap)
        size_idx = rtl_idx + 1
        if size_idx < len(col_lines):
            size = col_lines[size_idx]
        else:
            size = ""

        # Sanity check: SKU should be alphanumeric with hyphens
        if not re.match(r"^[A-Z0-9]", sku):
            continue

        # Sanity: WHL/RTL must contain EUR
        if not EUR_PRICE_RE.search(whl) or not EUR_PRICE_RE.search(rtl):
            continue

        # Sanity: SIZE should be short (2-5 chars, uppercase or contains +/./digits)
        # If SIZE contains "EUR", it's actually a misplaced price line, skip card
        if "EUR" in size:
            continue

        cards.append((sku, desc, matl, color, whl, rtl, size))

    return cards


def _parse_pdf_to_cards(pdf_path: Path) -> List[Tuple[str, str, str, str, str, str, str]]:
    """Extract product cards from McQueen PDF.

    Uses pdftotext (poppler) for reliable column-aligned extraction;
    pdfplumber doesn't preserve layout for this PDF type.
    """
    import subprocess
    import tempfile

    cards = []
    # Run pdftotext -layout to get column-aligned text
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
        tmp_path = f.name
    try:
        subprocess.run(
            ["pdftotext", "-layout", str(pdf_path), tmp_path],
            check=True, capture_output=True
        )
        with open(tmp_path, encoding="utf-8") as f:
            text = f.read()
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    # Split into blocks by blank lines (≥2 consecutive \n)
    blocks = re.split(r"\n\s*\n", text)
    for block in blocks:
        lines = [l.rstrip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue
        # 7 lines per card expected (or 21 lines for 3 cards)
        cards.extend(_parse_card_block(lines))
    return cards


def _extract_card_images(pdf_path: Path, n_cards: int) -> List[Optional[bytes]]:
    """Extract product images from PDF in card order (col 0→row1→row2, etc.).

    Returns a list of image bytes, length = n_cards, with None for cards
    without a matching image (e.g., last page may have fewer images).
    """
    import pdfplumber
    card_images: List[Optional[bytes]] = [None] * n_cards
    if n_cards == 0:
        return card_images

    # Collect all product images per page (skip logo: top < 200)
    page_images: List[List[Tuple[float, float, bytes]]] = []
    with pdfplumber.open(pdf_path) as p:
        for page in p.pages:
            imgs: List[Tuple[float, float, bytes]] = []
            for img in page.images:
                if img['top'] < 170:
                    continue  # skip logo / header
                try:
                    data = img['stream'].get_data()
                except Exception:
                    continue
                imgs.append((img['top'], img['x0'], data))
            imgs.sort(key=lambda t: (t[0], t[1]))   # by row (top), then col (x0)
            page_images.append(imgs)

    # Flatten into card-index order (cards 0..5 per page, top-left first)
    card_idx = 0
    for imgs in page_images:
        for _top, _x0, data in imgs:
            if card_idx >= n_cards:
                break
            card_images[card_idx] = data
            card_idx += 1
        if card_idx >= n_cards:
            break

    return card_images


def _detect_gender(pdf_path: Path) -> str:
    """Detect gender from filename (used only as a fallback if explicit gender exists).

    Returns "" by default — PDF files don't have explicit gender field, so we leave
    C5 empty (V1003 will then leave gender empty in output per user pref).
    """
    return ""


def _detect_category(pdf_path: Path) -> str:
    """Detect category from filename."""
    name = pdf_path.name.upper()
    if "EYEWEAR" in name:
        return "EYEWEAR"
    if "SHOES" in name:
        return "SHOES"
    if "RTW" in name:
        return "RTW"
    if "ACC" in name and "WOMEN" in name:
        return "ACC"
    if "LG" in name:
        return "LG"
    if "OTHER" in name:
        return "OTHER"
    return "OTHER"


def pdf_to_xlsx(pdf_path: Path, out_path: Path) -> int:
    """Convert McQueen PDF → xlsx matching the .xls schema. Return row count."""
    cards = _parse_pdf_to_cards(pdf_path)
    if not cards:
        return 0

    # Extract images per card index (0-based, in card-order across pages)
    card_images = _extract_card_images(pdf_path, len(cards))

    gender = _detect_gender(pdf_path)
    category = _detect_category(pdf_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row at R14,column layout matches McQueen .xls exactly:
    #   C1=PIC, C2=PIC NUMBER, C3=LOCATION, C4=DESTINATION,
    #   C5=GENDER, C6=CATEGORY, C7=NOTE2, C8=ECCEZIONI, C9=CATEGORY2,
    #   C10=DESCRIPTION, C11=SKU, C12=MATERIAL,
    #   C13=MATERIAL DESCRIPTION, C14=COLOR DESCRIPTION,
    #   C15=SIZES, C16=WHL PRICE, C17=RTL PRICE, C18=SIZE
    HEADER_LABELS = [
        "PIC", "PIC NUMBER", "LOCATION", "DESTINATION",
        "GENDER", "CATEGORY", "NOTE", "ECCEZIONI", "CATEGORY",
        "DESCRIPTION", "SKU", "MATERIAL",
        "MATERIAL DESCRIPTION", "COLOR DESCRIPTION",
        "SIZES", "WHL PRICE", "RTL PRICE", "SIZE",
    ]
    for c, v in enumerate(HEADER_LABELS, start=1):
        ws.cell(14, c, v)

    # Data rows starting R15
    for i, card in enumerate(cards):
        sku, desc, matl, color, whl_str, rtl_str, scalarino = card
        # Clean strings of illegal chars (form-feed, control chars)
        def _clean(s: str) -> str:
            return "".join(c for c in s if c.isprintable() or c in "\t\n").strip()

        sku = _clean(sku)
        desc = _clean(desc)
        matl = _clean(matl)
        color = _clean(color)
        scalarino = _clean(scalarino)

        # Parse EUR prices
        whl_m = EUR_PRICE_RE.search(whl_str)
        rtl_m = EUR_PRICE_RE.search(rtl_str)
        whl = float(whl_m.group(1).replace(".", "").replace(",", ".")) if whl_m else 0
        rtl = float(rtl_m.group(1).replace(".", "").replace(",", ".")) if rtl_m else 0

        # Skip empty cards (padding artifacts)
        if not sku:
            continue

        r_idx = 15 + i
        ws.cell(r_idx, 5, gender)        # C5 GENDER
        ws.cell(r_idx, 6, category)      # C6 CATEGORY
        ws.cell(r_idx, 10, desc)         # C10 DESCRIPTION
        ws.cell(r_idx, 11, sku)          # C11 SKU
        ws.cell(r_idx, 13, matl)         # C13 MATERIAL DESCRIPTION
        ws.cell(r_idx, 14, color)        # C14 COLOR DESCRIPTION
        ws.cell(r_idx, 15, "")           # C15 SIZES (空,catalog 没具体尺码)
        ws.cell(r_idx, 16, whl)          # C16 WHL PRICE
        ws.cell(r_idx, 17, rtl)          # C17 RTL PRICE
        ws.cell(r_idx, 18, scalarino)    # C18 SIZE code

    # Embed images in column 2 (image column) for each card row
    try:
        from openpyxl.drawing.image import Image as XLImage
        from PIL import Image as PILImage
        import io as _io
        for card_idx, img_bytes in enumerate(card_images):
            if img_bytes is None:
                continue
            try:
                # Verify JPEG is valid
                pil_img = PILImage.open(_io.BytesIO(img_bytes))
                pil_img.verify()
            except Exception:
                continue
            r_idx = 15 + card_idx
            # Embed at C2, anchored to top-left of cell, sized to fit
            pil_img = PILImage.open(_io.BytesIO(img_bytes))
            img_obj = XLImage(_io.BytesIO(img_bytes))
            # Scale image to fit ~40 col width × 92pt row height
            target_h = 92 * 9525 * 1.0   # 92pt in EMU
            if img_obj.height > 0:
                scale = target_h / img_obj.height
                img_obj.width = int(img_obj.width * scale)
                img_obj.height = int(target_h)
            img_obj.anchor = f"B{r_idx}"
            ws.add_image(img_obj)
    except Exception as e:
        print(f"Warning: image embedding failed: {e}")

    wb.save(out_path)
    return len(cards)


def main():
    if len(sys.argv) < 3:
        print("Usage: preprocess_mcqueen_pdf.py <pdf_path> <xlsx_out_path>")
        sys.exit(1)
    pdf_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2])
    n = pdf_to_xlsx(pdf_path, out_path)
    print(f"Converted {pdf_path.name} → {out_path.name}: {n} rows")


if __name__ == "__main__":
    main()