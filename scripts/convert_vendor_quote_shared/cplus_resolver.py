"""
C+ 解析器 - 3 层优先级规则匹配 (Phase 1 改版,2026-07-31 再改版:多条件并行)

优先级 (高 → 低):
  L1: 行级覆盖 (overrides CSV:  sku, cplus)
  L2: 品类规则 (rules:  [{brand, gender, category, cplus}, ...])
      每条规则 3 个 keyword 字段并列(AND 关系):
        - brand    品牌(空 = 任意品牌)
        - gender   性别(空 = 任意性别)
        - category 类目(空 = 任意类目)
      命中 = 3 个非空 keyword 各自在 row 对应字段中包含 (case-insensitive)
      列表顺序即优先级, 第一条全字段匹配的规则胜出
  L3: 品牌默认值 (调用方传入, 例如从 "品牌对应C+.xlsx" 读到的)
  兜底: 全部没命中 → cplus=None, C+- 留空

返回: (cplus_int, rule_label)
  - cplus_int: 整数 (18 / -10 / 25),  None = 无 C+
  - rule_label: 审计标签, 例如 "L1:override:740237-V2X40" / "L2:BURBERRY|DONNA|BORSE=+10" / "L3:default"

设计目标:
  - 调用方 (convert_vendor_quote.py) 集中处理, vendor adapter 不用改
  - 规则可视化, app UI 直接编辑, 每次提交时由前端传给后端
  - 不写正则, 普通运营就能维护
  - 多条件并行 (AND):一条规则 = 3 维过滤(品牌/性别/类目),实际 markup 文档(品牌对应C+.xlsx)就是这种结构
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import yaml


# 规则 3 维字段(固定不变 — 与 app UI 对应)
RULE_FIELDS = ("brand", "gender", "category")

# 字段名 → 中文标签 (UI 表头 + 审计 label 用)
FIELD_LABELS = {
    "brand": "品牌",
    "gender": "性别",
    "category": "类目",
}

# resolve 时读 row 的字段映射(规则 field → row 字段名)
# 品牌/性别/类目 与 row dict 字段名一致,不用映射
ROW_FIELD = {
    "brand": "brand",
    "gender": "gender",
    "category": "category",
}


# ---------------------------------------------------------------------------
# 配置加载
# ---------------------------------------------------------------------------

def _norm(s: str) -> str:
    return (s or "").strip().lower()


def _load_overrides_csv(csv_path: Path) -> Dict[str, int]:
    """CSV 列: sku, cplus  (大小写不敏感, 表头 sku / SKU 均可)"""
    if not csv_path or not csv_path.exists():
        return {}
    out: Dict[str, int] = {}
    with open(csv_path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sku = _norm(row.get("sku") or row.get("SKU") or row.get("货号") or "")
            cplus_raw = (row.get("cplus") or row.get("C+") or row.get("c+%") or "").strip()
            if sku and cplus_raw:
                try:
                    out[sku] = int(cplus_raw)
                except ValueError:
                    pass
    return out


def _load_rules_yaml(yml_path: Path) -> List[dict]:
    """YAML 是 list[dict]: [{brand, gender, category, cplus}, ...] (3 维并列 AND)

    文件中顺序即优先级, 第一个匹配的胜出。
    """
    if not yml_path or not yml_path.exists():
        return []
    with open(yml_path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not data or not isinstance(data, list):
        return []
    out: List[dict] = []
    for r in data:
        if not isinstance(r, dict):
            continue
        cplus = r.get("cplus")
        if cplus is None:
            continue
        try:
            cplus_int = int(cplus)
        except (TypeError, ValueError):
            continue
        # 校验 + 提取 3 维 keyword(任一空 = 不过滤该维度)
        clean = {"cplus": cplus_int}
        for f in RULE_FIELDS:
            v = str(r.get(f) or "").strip()
            clean[f] = v
        # 至少要有一个非空 keyword
        if not any(clean[f] for f in RULE_FIELDS):
            continue
        out.append(clean)
    return out


def _rule_matches(rule: dict, row: dict) -> bool:
    """判断 rule 是否匹配 row (3 维 AND)。

    任一 rule 字段非空且在 row 对应字段中不包含 → 不匹配。
    任一 rule 字段为空 → 跳过该维度(等价于 *)。
    """
    for f in RULE_FIELDS:
        kw = (rule.get(f) or "").strip()
        if not kw:
            continue  # 空 keyword = 跳过该维度
        row_field = ROW_FIELD[f]
        cell_value = str(row.get(row_field) or "")
        if kw.lower() not in cell_value.lower():
            return False
    return True


# ---------------------------------------------------------------------------
# 核心 resolver
# ---------------------------------------------------------------------------

class CplusResolver:
    """3 层优先级 C+ 解析器。构造时加载配置, resolve(row_dict) 返回 (cplus, label)。"""

    def __init__(
        self,
        rules: Optional[List[dict]] = None,
        overrides: Optional[Dict[str, int]] = None,
        default_cplus: Optional[int] = None,
    ):
        self.rules: List[dict] = rules or []
        self.overrides: Dict[str, int] = overrides or {}
        self.default_cplus: Optional[int] = default_cplus

    def resolve(self, row: dict) -> Tuple[Optional[int], str]:
        """根据 row (含 brand / gender / category 字段) 决定 C+。

        返回:
          (cplus_int_or_None, rule_label)
        """
        row = row or {}
        sku = row.get("sku") or ""

        # L1: 行级覆盖 (sku 精确匹配, 大小写不敏感)
        sku_key = _norm(sku)
        if sku_key and sku_key in self.overrides:
            return self.overrides[sku_key], f"L1:override:{sku}"

        # L2: 品类规则 (3 维 AND 匹配, 首条命中胜出)
        for r in self.rules:
            if _rule_matches(r, row):
                cplus = int(r.get("cplus") or 0)
                # 审计 label: 3 个非空 keyword 拼出来
                parts = []
                for f in RULE_FIELDS:
                    kw = r.get(f, "")
                    if kw:
                        parts.append(f"{FIELD_LABELS[f]}:{kw}")
                desc = "|".join(parts) if parts else "any"
                sign = "+" if cplus >= 0 else "-"
                return cplus, f"L2:{desc}={sign}{abs(cplus)}"

        # L3: 品牌默认值
        if self.default_cplus is not None:
            return int(self.default_cplus), "L3:default"

        return None, ""

    # ---- 工厂方法 ----

    @classmethod
    def from_rules(cls, rules: List[dict], default_cplus: Optional[int] = None) -> "CplusResolver":
        """直接接受规则列表 (前端 JSON 传来的格式)。

        规则格式: [{brand, gender, category, cplus}, ...]
        任一 keyword 字段可空(空 = 不过滤该维度)。
        至少要有一个非空 keyword。
        """
        clean: List[dict] = []
        for r in rules:
            if not isinstance(r, dict):
                continue
            try:
                cplus_int = int(r.get("cplus"))
            except (TypeError, ValueError):
                continue
            entry = {"cplus": cplus_int}
            for f in RULE_FIELDS:
                entry[f] = str(r.get(f) or "").strip()
            # 至少一个非空 keyword
            if not any(entry[f] for f in RULE_FIELDS):
                continue
            clean.append(entry)
        return cls(rules=clean, default_cplus=default_cplus)

    @classmethod
    def from_files(
        cls,
        rules_file: Optional[Path] = None,
        overrides_file: Optional[Path] = None,
        default_cplus: Optional[int] = None,
    ) -> "CplusResolver":
        rules = _load_rules_yaml(Path(rules_file)) if rules_file else []
        overrides = _load_overrides_csv(Path(overrides_file)) if overrides_file else {}
        return cls(rules=rules, overrides=overrides, default_cplus=default_cplus)

    # ---- 调试辅助 ----

    def summary(self) -> str:
        return (
            f"CplusResolver(rules={len(self.rules)}, "
            f"overrides={len(self.overrides)}, "
            f"default={self.default_cplus})"
        )
