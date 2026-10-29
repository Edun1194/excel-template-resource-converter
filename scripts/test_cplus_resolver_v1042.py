"""
V1042 (BOTTEGA VENETA) C+ Resolver 验证脚本

跑 V1042 源表, 用新加的 CplusResolver 算每行 C+, 输出:
  1) 总体分布 (C+25 / C+18 / C+15 各几行)
  2) 每条规则的命中数
  3) 抽样: 每类前 3 行 + Andiamo 全部

使用方法:
  python test_cplus_resolver_v1042.py
  (源文件用环境变量 V1042_SRC 指定,规则文件用 V1042_CPLUS_RULES 指定)
"""

from __future__ import annotations

import os
import sys
from collections import Counter
from pathlib import Path

# 让脚本能找到 vendor_v1042 和 shared 模块
SCRIPTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS_DIR))

from convert_vendor_quote_shared.cplus_resolver import CplusResolver
import vendor_v1042  # noqa: E402

# 源表 / 规则文件都不写死路径:默认走仓库相对 testdata/,
# 也可用环境变量指向本机实际位置。
SRC = Path(
    os.environ.get("V1042_SRC")
    or (SCRIPTS_DIR.parent / "testdata" / "V1042" / "BOTTEGA VENETA.xlsx")
)
RULES = Path(
    os.environ.get("V1042_CPLUS_RULES")
    or (SCRIPTS_DIR.parent / "assets" / "rules" / "cplus" / "BOTTEGA VENETA.yaml")
)


def main():
    print("=" * 70)
    print("V1042 C+ Resolver 验证")
    print("=" * 70)
    print(f"源文件: {SRC}")
    print(f"规则文件: {RULES}")
    print()

    # 1) 用 vendor_v1042 解析源数据
    rows = vendor_v1042.parse_rows(SRC, None, "V1042")
    print(f"源数据解析出 {len(rows)} 行")
    print()

    # 2) 构造 resolver (不带 default, 让 .* 兜底生效)
    resolver = CplusResolver.from_files(
        rules_file=RULES,
        default_cplus=None,  # 兜底由 yml 最后一条 .* 负责
    )
    print(f"resolver: {resolver.summary()}")
    print()

    # 3) 跑全表
    hits_by_value: Counter = Counter()
    hits_by_label: Counter = Counter()
    sample_per_label: dict = {}
    full_results = []  # (sku, desc, cplus, label)
    for row in rows:
        brand = row.get("brand") or ""
        sku = row.get("sku") or ""
        desc = row.get("description") or ""
        cplus_val, label = resolver.resolve(brand=brand, sku=sku, description=desc)
        hits_by_value[cplus_val] += 1
        hits_by_label[label] += 1
        if label not in sample_per_label:
            sample_per_label[label] = []
        if len(sample_per_label[label]) < 3:
            sample_per_label[label].append((sku, desc, cplus_val))
        full_results.append((sku, desc, cplus_val, label))

    # 4) 输出
    print("=" * 70)
    print("C+ 数值分布")
    print("=" * 70)
    for val in sorted(hits_by_value.keys(), key=lambda x: (x is None, x if x is not None else 0)):
        v = val if val is not None else "(none)"
        print(f"  C+{v}  →  {hits_by_value[val]} rows")

    print()
    print("=" * 70)
    print("规则命中分布 (按规则 label 排序, 命中数从高到低)")
    print("=" * 70)
    for label, cnt in sorted(hits_by_label.items(), key=lambda x: (-x[1], x[0])):
        print(f"  {label:50s}  →  {cnt:3d} rows")

    print()
    print("=" * 70)
    print("每条规则抽样 (前 3 行)")
    print("=" * 70)
    for label in sorted(sample_per_label.keys()):
        print(f"\n[{label}]")
        for sku, desc, cplus in sample_per_label[label]:
            short_desc = desc[:55] + "..." if len(desc) > 55 else desc
            print(f"  {sku:30s} | C+{cplus:+d} | {short_desc}")

    # 5) Andiamo 全部 (单独检查)
    print()
    print("=" * 70)
    print("Andiamo 全部 (验证 C+25)")
    print("=" * 70)
    andiamo_rows = [r for r in full_results if "andiamo" in (r[1] or "").lower()]
    print(f"共 {len(andiamo_rows)} 行")
    for sku, desc, cplus, label in andiamo_rows:
        ok = "✓" if cplus == 25 else "✗"
        short_desc = desc[:50] + "..." if len(desc) > 50 else desc
        print(f"  {ok} {sku:30s} | C+{cplus:+d} | {label} | {short_desc}")


if __name__ == "__main__":
    main()
