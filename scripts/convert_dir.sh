#!/bin/bash
# 把指定目录下所有 .xlsx/.xls 跑一遍 V### 转换
# 用法: ./convert_dir.sh <vendor_code> <input_dir> [output_subdir=converted]
#
# 例: ./convert_dir.sh V1031 ~/Desktop/表格处理/配置器测试文件/V1031/

set -e

VENDOR="${1:-V1031}"
INPUT_DIR="${2:?需要输入目录}"
OUTPUT_SUBDIR="${3:-converted}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

mkdir -p "$INPUT_DIR/$OUTPUT_SUBDIR"

for f in "$INPUT_DIR"/*.xlsx "$INPUT_DIR"/*.xls; do
    [ -f "$f" ] || continue
    base="$(basename "$f")"
    stem="${base%.*}"
    out="$INPUT_DIR/$OUTPUT_SUBDIR/${stem}_converted.xlsx"

    echo ">>> [$VENDOR] $base"
    python3 "$SCRIPT_DIR/convert_vendor_quote.py" \
        --primary-file "$f" \
        --vendor "$VENDOR" \
        --output "$out"
    echo ""
done
