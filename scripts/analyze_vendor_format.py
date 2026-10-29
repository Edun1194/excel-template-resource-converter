#!/usr/bin/env python3
"""
Analyze a vendor workbook and generate an adapter skeleton.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import List

import pandas as pd


def normalize(name: object) -> str:
    return re.sub(r"\s+", " ", str(name or "").strip())


def detect_vendor_hint(file_path: Path, sheet_names: List[str]) -> str:
    p = str(file_path).lower()
    s = {x.lower() for x in sheet_names}
    if "stocklist" in p or {"hoja1", "hoja3"}.issubset(s):
        return "V1049"
    if "ss2026" in p or {"m", "w"}.issubset(s):
        return "V1044"
    return "UNKNOWN"


def build_adapter_stub(vendor_id: str, sample_file: Path, out_file: Path) -> None:
    code = f'''#!/usr/bin/env python3
"""
Adapter skeleton for vendor {vendor_id}.
"""

from pathlib import Path
from typing import Dict, List

import pandas as pd


def parse_{vendor_id.lower()}(primary_file: Path, markup_file: Path, warehouse_id: str) -> List[Dict[str, object]]:
    xls = pd.ExcelFile(primary_file)
    rows: List[Dict[str, object]] = []
    # TODO 1: Parse stock/size matrix from vendor workbook
    # TODO 2: Parse cost+ rules from markup file
    # TODO 3: Append normalized rows with fields:
    # brand, sku, image, description, category, material, color, season,
    # origin, cost, cplus_text, supply_price, retail_price, size, quantity, vendor_id
    return rows
'''
    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_text(code, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze vendor format and generate adapter skeleton.")
    parser.add_argument("--primary-file", required=True, help="Path to vendor stock/quote file.")
    parser.add_argument("--vendor-id", help="Optional vendor id, defaults to detected hint.")
    parser.add_argument(
        "--output",
        help="Adapter skeleton output path. Default: outputs/generated_adapter_<vendor>.py",
    )
    args = parser.parse_args()

    primary = Path(args.primary_file).expanduser().resolve()
    xls = pd.ExcelFile(primary)
    sheet_names = xls.sheet_names

    vendor_hint = detect_vendor_hint(primary, sheet_names)
    vendor = args.vendor_id or vendor_hint
    skill_root = Path(__file__).resolve().parents[1]
    output = (
        Path(args.output).expanduser().resolve()
        if args.output
        else skill_root / "outputs" / f"generated_adapter_{vendor.lower()}.py"
    )

    print(f"file={primary}")
    print(f"detected_vendor={vendor_hint}")
    print(f"sheets={sheet_names}")
    for name in sheet_names:
        df = pd.read_excel(primary, sheet_name=name, nrows=4, header=None)
        preview = [[normalize(v) for v in row] for row in df.values.tolist()]
        print(f"sheet={name} preview_rows={preview}")

    build_adapter_stub(vendor, primary, output)
    print(f"adapter={output}")


if __name__ == "__main__":
    main()
