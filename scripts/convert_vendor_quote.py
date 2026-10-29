#!/usr/bin/env python3
"""
Convert vendor quote/stock files into template1-cn.xlsx or template1-en.xlsx.

This is the main entry-point.  Each vendor's parse logic lives in its own
``vendor_v<ID>.py`` module; shared utilities are in
``convert_vendor_quote_shared.py``.
"""

from __future__ import annotations

import argparse
import importlib
import mimetypes
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from convert_vendor_quote_shared import (
    ASSETS_DIR,
    normalize_text,
    pick_template,
    write_output,
)

# 注册 openpyxl 缺的 MIME type — 2026-08-18
# 修 PyMuPDF 返回 jp2/jpeg2000 格式时,openpyxl 写文件报 KeyError: '.jpeg2000'
# openpyxl/packaging/manifest.py:177 查 mimetypes.types_map[True][ext]
# Python 3.7+ 的 mimetypes.types_map 是普通 dict(没 True 键),生产环境可能是
# Python 3.6 兼容 shim 或自定义 patch 让 types_map[True] 是个 dict
# 修复:同时写进 4 个地方 — 稳
JP2_EXTS = [".jp2", ".jpeg2000", ".jpx", ".j2k", ".jpf"]
for ext in JP2_EXTS:
    mime = "image/jp2"
    # 1) 标准 add_type(写到 _db.types_map strict & non-strict)
    mimetypes.add_type(mime, ext, strict=True)
    mimetypes.add_type(mime, ext)
    # 2) 顶层 types_map(普通 dict,直接键)
    try:
        mimetypes.types_map[ext] = mime
    except Exception:
        pass
    # 3) types_map[True] - openpyxl 实际查的
    try:
        mimetypes.types_map[True][ext] = mime
    except (KeyError, TypeError):
        # types_map[True] 不存在 — 把 strict map 挂上去
        try:
            mimetypes._db_init()  # 确保 _db 存在
        except AttributeError:
            pass
        try:
            mimetypes._db  # touch
        except Exception:
            pass
        # 强制 add_type 一次触发 _db init,然后把 strict map 挂到 types_map[True]
        mimetypes.add_type(mime, ext, strict=True)
        try:
            mimetypes.types_map[True] = mimetypes._db.types_map[1]
        except Exception:
            pass
    # 4) 防御:_db.types_map[1] 写一遍(strict map)
    try:
        if mimetypes._db is not None and len(mimetypes._db.types_map) > 1:
            mimetypes._db.types_map[1][ext] = mime
    except Exception:
        pass

# ---------------------------------------------------------------------------
# 全局默认值(可被 vendor meta / CLI 覆盖)
# ---------------------------------------------------------------------------
# 有图片的数据行默认行高(pt)。按 references/patterns/image-handling.md
# 里的"高度上限 92pt 缩放"规则,所有 vendor 统一用 92pt。
#   - CLI --image-row-height 优先
#   - 然后是 vendor 的 get_vendor_meta()["default_image_row_height"]
#   - 最后是这个全局默认
DEFAULT_IMAGE_ROW_HEIGHT: float = 92.0

# ---------------------------------------------------------------------------
# Built-in vendor adapter registry
# ---------------------------------------------------------------------------

BUILTIN_VENDORS: Dict[str, str] = {
    "V1001": "vendor_v1001",  # BARBOUR 货单 (2026-10-22,稀疏表: 每产品块 15 行,块内每尺码一行)
    "V1003": "vendor_v1003",  # McQueen SS27 期货 (.xls + .pdf preprocess,2026-09-29 接入,GUCCI Sirio DB 占位历史名)
    "V1003-期货": "vendor_v1003_futures",  # 期货系总入口: McQueen SS27 + Armani PLM Report(2026-10-21) + LONGCHAMP SS27(2026-10-21 从 V1003-LC 合并),输出 vendor_id 统一 "V1003"
    "category-futures": "vendor_category_futures",  # 2026-10-27: 期货通道, 解析复用 V1003-期货, 但「云仓」列写「品牌名称（期货）」
    "V1004": "vendor_v1004",
    "V1004-期货": "vendor_v1004_futures",  # 2026-10-27: 期货通道, SALOMON Foglio 版式数量留空等客户填; 现货 4 品牌委托 V1004
    "V1010": "vendor_v1010",
    "A1001": "vendor_a1001",  # CANADA GOOSE hk 现货 .xls (2026-10-12)
    "V1011": "vendor_v1011",  # MAX MARA 意大利供应商 CATALOGO sheet (Parametri + CATALOGO 2 sheet,2026-10-08 接入)
    "V1012": "vendor_v1012",
    "V1015": "vendor_v1015",
    "V1021": "vendor_v1021",
    "V1025": "vendor_v1025",  # POLO RALPH LAUREN 现货单 xlsx 新版式 (2026-09-16)
    "V1031": "vendor_v1031",
    "V1037": "vendor_v1037",  # SALOMON 意大利线(xlsx,WPS 从 .xls 转)
    "C0000": "vendor_c0000",  # GIVENCHY China 库存清单(2026-09-23 写)
    "V1038": "vendor_v1038",
    "V1061": "vendor_v1061",  # 2026-10-28: GUCCI FW26 - WAREHOUSE B 平铺库存清单(解析复用 shared/flat_stocklist.py)
    "V1039": "vendor_v1039",
    "V1041": "vendor_v1041",
    "V1042": "vendor_v1042",
    "V1043": "vendor_v1043",
    "V1044": "vendor_v1044",
    "V1046": "vendor_v1046",
    "V1048": "vendor_v1048",
    "V1049": "vendor_v1049",
    "V1050": "vendor_v1050",
    "V1062": "vendor_v1062",  # GUCCI WAREHOUSE B/C 现货单(共用 V1038 col14 parser)
    "V1067": "vendor_v1067",  # PANICALE 现货单 + 外部图片目录匹配
    "V1070": "vendor_v1070",  # BRUNELLO CUCINELLI (BC) 现货单(单 sheet,无 cost 列)
    "V1070-期货": "vendor_v1070_nuorder",  # 2026-09-02 接入:BC NuORDER SS27/SS28 期货(Style Number + Color Code 锚点)
    "V1072": "vendor_v1072",  # BURBERRY 现货单(SKU 在 col J,描述含类型+性别+颜色+价格+产地+材质)
    "V1072-期货": "vendor_v1072_in_arrivo",  # 2026-09-03 接入:LISTE FW26 IN ARRIVO 8 个品牌 (D8 EA / I6 L8 MAGDA BUTRYN / O8 AMIRI CASABLANCA OFFWHITE PALM ANGELS RHUDE)
    "V1015-补货单": "vendor_v1015_restock",  # V1015 工厂补货单 (richValue 路径,意大利大公司 SharePoint 风格)
    "V1015-补货单2": "vendor_v1015_restock2",  # V1015 工厂补货单 2 (MIUMIU ATS,3 sheet 多品类 — Pelletteria/Calzature/Abbigliamento)
    "V1015-补货单3": "vendor_v1015_restock3",  # V1015 工厂补货单 3 (GIORGIO ARMANI C-20,4 sheet 横排尺码)
    "V1015-补货单4": "vendor_v1015_restock4",  # V1015 工厂补货单 4 (EA7 C-20,1 sheet 纵排 17 列)
}

# ---------------------------------------------------------------------------
# Vendor detection
# ---------------------------------------------------------------------------
# NOTE: 这个函数目前没有 caller(app.py 在 form 里强制要求 --vendor)。
# 保留是因为:
#   1. CLI 手动跑脚本时方便(`python convert_vendor_quote.py file.xlsx`,不传 --vendor)
#   2. 上游有工具/脚本可能依赖此函数做"按文件名猜 vendor"
# 如果未来确认无 caller,可以删。
# ---------------------------------------------------------------------------


def detect_vendor(primary_file: Path) -> str:
    file_str = str(primary_file).lower()
    name = primary_file.name.lower()
    if "v1021" in file_str or "bruna_rosso" in file_str:
        return "V1021"
    if "stocklist" in name or "pompeu" in file_str:
        return "V1049"
    if "ss2026" in name or "ratti" in file_str:
        return "V1044"
    if "v1031" in file_str or "nugnes" in file_str:
        return "V1031"
    if "v1043" in file_str or "pritelli" in file_str:
        return "V1043"
    if "v1046" in file_str or "loschi" in file_str:
        return "V1046"
    if "v1041" in file_str or "biffi" in file_str:
        return "V1041"
    if "v1042" in file_str or "alina" in file_str:
        return "V1042"
    if "v1039" in file_str or "antonioli" in file_str:
        return "V1039"
    if "v1048" in file_str or "maltempi" in file_str:
        return "V1048"
    if "v1050" in file_str or "follifollie" in file_str:
        return "V1050"
    if "v1067" in file_str or "panicale" in file_str or "stockuomo" in file_str or "stockdonna" in file_str:
        return "V1067"
    if "v1007" in file_str or "ganni" in file_str:
        return "V1007"
    try:
        xls = pd.ExcelFile(primary_file)
        sheets = [s.lower() for s in xls.sheet_names]
        if "hoja1" in sheets and "hoja3" in sheets:
            return "V1049"
        if {"m", "w"}.issubset(set(sheets)):
            return "V1044"
        if "foglio1" in sheets:
            df = pd.read_excel(primary_file, sheet_name="Foglio1", nrows=2, header=None)
            vals = {normalize_text(v).upper() for v in df.iloc[0].tolist()}
            if {"BRAND", "ITEM", "SIZE", "QTY"}.issubset(vals):
                return "V1043"
    except Exception:
        pass
    raise ValueError(f"Cannot detect vendor from file: {primary_file}")


# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------


def resolve_paths(args: argparse.Namespace) -> Tuple[Path, Optional[Path], Path]:
    primary = Path(args.primary_file).expanduser().resolve()
    if not primary.exists():
        raise FileNotFoundError(f"Primary file does not exist: {primary}")

    vendor = args.vendor or detect_vendor(primary)
    if vendor == "V1044":
        markup = (
            Path(args.markup_file).expanduser().resolve()
            if args.markup_file
            else ASSETS_DIR / "resources" / "V1044 RATTI" / "MU_SS26 -10_02_26.xlsx"
        )
    elif vendor == "V1050":
        markup = None
    else:
        # V1049 等其他 vendor:C+ 永远走用户输入,不需要 markup 文档
        markup = Path(args.markup_file).expanduser().resolve() if args.markup_file else None

    template = (
        Path(args.template).expanduser().resolve() if args.template else pick_template(args.language)
    )
    # markup 缺失不直接报错,改成 warn:有 markup 就用,没有就 C+/供货价 留空(用户后面补)
    if markup is not None and not markup.exists():
        print(f"WARNING: Markup file does not exist: {markup} — C+/供货价 will be empty.")
        print(f"  To provide: --markup-file /path/to/MARKUP.docx")
        markup = None
    if not template.exists():
        raise FileNotFoundError(f"Template file does not exist: {template}")

    return primary, markup, template


# ---------------------------------------------------------------------------
# Adapter loading
# ---------------------------------------------------------------------------


def _load_adapter(vendor: str, primary_file: Path):
    """Import a vendor adapter module and return it.

    Built-in adapters live in the same package (``scripts/``).
    External adapters are discovered by looking for a
    ``convert_<vendor>.py`` file next to *primary_file*.
    Fallback: search this script's own directory for ``vendor_<lowercase>.py``
    (e.g. ``vendor_vc001.py`` for vendor ``VC001``) — useful for vendors
    not yet registered in BUILTIN_VENDORS.
    """
    if vendor in BUILTIN_VENDORS:
        return importlib.import_module(BUILTIN_VENDORS[vendor])

    external_script = primary_file.parent / f"convert_{vendor}.py"
    if external_script.is_file():
        spec = importlib.util.spec_from_file_location(f"convert_{vendor}", str(external_script))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod

    # Fallback: scripts/vendor_<lowercase>.py (适配未注册到 BUILTIN_VENDORS 的 vendor)
    scripts_dir = Path(__file__).parent
    local_script = scripts_dir / f"vendor_{vendor.lower()}.py"
    if local_script.is_file():
        spec = importlib.util.spec_from_file_location(f"vendor_{vendor}", str(local_script))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod

    raise ValueError(
        f"No adapter found for vendor {vendor}. "
        f"Expected built-in module, {external_script}, or {local_script}"
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert vendor quote sheets to fixed template format.",
    )
    parser.add_argument("--primary-file", required=True, help="Vendor stock/quote xlsx file.")
    parser.add_argument("--markup-file", help="Optional markup file path (.xlsx or .docx by vendor).")
    parser.add_argument(
        "--vendor",
        help="Optional vendor id. Auto-detected by default.",
    )
    parser.add_argument(
        "--language", choices=["cn", "en"], default="cn",
        help="Template language, default cn.",
    )
    parser.add_argument("--template", help="Optional template file path; overrides --language.")
    parser.add_argument("--warehouse-id", help="Optional warehouse/vendor id override.")
    parser.add_argument(
        "--output",
        help="Output file path. Default: same directory as primary file.",
    )
    parser.add_argument(
        "--cplus",
        type=int,
        default=None,
        help="Override C+ value (integer, e.g. 18 or -18). Empty=use vendor's own logic.",
    )
    parser.add_argument(
        "--cplus-rules-file",
        type=str,
        default=None,
        help="C+ 品类规则 yml 路径(CLI 用,旧路径)。前端请用 --cplus-rules-json。",
    )
    parser.add_argument(
        "--cplus-rules-json",
        type=str,
        default=None,
        help='C+ 品类规则 JSON 字符串(前端传来)。格式: [{field,keyword,sign,cplus}, ...]。',
    )
    parser.add_argument(
        "--cplus-default",
        type=int,
        default=None,
        help="Rules 模式下的兜底 C+ 值(整数, 可正可负)。没命中任何规则的行用这个;None=这些行 C+ 留空。",
    )
    parser.add_argument(
        "--cplus-overrides-file",
        type=str,
        default=None,
        help="C+ 行级覆盖 CSV 路径(Phase 1 新增,列: sku, cplus)。优先级最高。",
    )
    parser.add_argument(
        "--no-compute-supply-price",
        action="store_true",
        help="不计算供货价字段(只写 C+ 文本,留空 supply_price)。当 C+ 已有值时此 flag 才生效;无 C+ 时 supply_price 永远为空。",
    )
    parser.add_argument(
        "--brand",
        default=None,
        help="Override brand name for all rows. Empty=use vendor's detected brand.",
    )
    parser.add_argument(
        "--image-row-height",
        type=float,
        default=None,
        help="有图片的数据行的目标行高（pt）。如 92=92pt。设置后图片显示尺寸会按比例缩到该行高内，源图字节不变（清晰度保留）。默认 None=沿用源行高。",
    )
    parser.add_argument(
        "--image-col-width",
        type=float,
        default=None,
        help="图片列的列宽(Excel 字符单位)。多图产品(V1049 每个 3-4 张)需要更宽。默认 None=沿用模板列宽。",
    )
    parser.add_argument(
        "--season",
        type=str,
        default=None,
        help="覆盖季节字段(如 SS26 / FW26)。优先级:用户填写 > 数据源中表明 > 都没有则为空。",
    )
    parser.add_argument(
        "--keep-images-per-row",
        type=int,
        default=None,
        help="每行保留几张图。None=vendor_meta.default_keep_images_per_row > KEEP_IMAGES_PER_ROW=1。V1003 Sirio DB=2(主图+辅图)。",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    primary, markup, template = resolve_paths(args)
    vendor = args.vendor or detect_vendor(primary)
    warehouse_id = args.warehouse_id or vendor

    # Ensure the scripts directory is on sys.path for built-in adapter imports
    scripts_dir = str(Path(__file__).resolve().parent)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)

    adapter = _load_adapter(vendor, primary)

    # 拉取该 vendor 的默认元数据（如 V0000 默认 92pt + 品牌 POLO RALPH LAUREN）
    # CLI 显式传的参数优先于 vendor 默认值
    vendor_meta = {}
    if hasattr(adapter, "get_vendor_meta"):
        try:
            vendor_meta = adapter.get_vendor_meta() or {}
        except Exception:
            vendor_meta = {}

    # 解析行（V0000/V1012 内部会强制 vendor_id）
    rows: List[Dict[str, object]] = adapter.parse_rows(primary, markup, warehouse_id)

    # 【品牌】字段策略(5 级优先级,所有 vendor 统一):
    #   1. 用户在 CLI 显式传 --brand "xxx"           → 用这个(最高,覆盖一切)
    #   2. vendor adapter 读到了源文件里的 BRAND 字段   → 用源里的值
    #   3. vendor meta 声明的 default_brand            → POLO / LONGCHAMP / A.P.C. 等
    #   4. 文件名(去扩展名)作为兜底                     → FENDI / Gucci / APC-WAREHOUSE A 等
    #   5. 都没有                                     → 留空
    # 关键:第 2 步是"源文件如果有明确标出 BRAND 字段,先填入",而不是"永远清空"
    user_brand = args.brand if (args.brand is not None and args.brand != "") else None
    default_brand = vendor_meta.get("default_brand")
    # 兜底:文件名(去扩展名)
    filename_brand = primary.stem if primary else ""
    # 统计源文件是否有 brand(用于日志)
    src_brands = {str(r.get("brand") or "").strip() for r in rows}
    src_brands.discard("")
    if src_brands:
        print(f"source_brands={sorted(src_brands)}")

    # 2026-10-22: vendor_meta["brand_from_filename"]=False 时禁用文件名兜底。
    #   场景: SIRIO DB (Michael Kors 等) 源文件**没有品牌列**,文件名也是
    #   'whls-70-T27-40' 这种内部代号而非品牌,兜底会产出垃圾值。
    #   这种情况留空,由用户在 App 的「品牌」输入框手填。
    allow_filename_brand = vendor_meta.get("brand_from_filename", True) is not False

    for row in rows:
        src_brand = str(row.get("brand") or "").strip()
        # 2026-10-22: 逐行开关。适配器给某行标了 _no_filename_brand=True 就禁用
        #   文件名兜底 (源文件没品牌列 + 文件名是内部代号, 兜底会产垃圾值)。
        #   按行而非按 vendor, 因为同一 vendor 模块可能同时处理有/无品牌列的版式。
        row_allow_filename = allow_filename_brand and not row.get("_no_filename_brand")
        if user_brand:
            row["brand"] = user_brand
        elif src_brand:
            # 源文件有 BRAND 字段 → 保留(已由 parse_rows 写入,这里只是保证不被覆盖)
            pass
        elif default_brand:
            row["brand"] = default_brand
        elif filename_brand and row_allow_filename:
            row["brand"] = filename_brand
        else:
            row["brand"] = ""

    if user_brand:
        print(f"brand_override={user_brand}")
    else:
        # 取第一行非空的 brand 作为"实际使用的品牌"用于日志
        used = next((r["brand"] for r in rows if r.get("brand")), "")
        if used:
            src_brand_used = used in src_brands if src_brands else False
            if src_brand_used:
                print(f"brand_from_source={used}")
            elif used == default_brand:
                print(f"brand_default={used}")
            elif used == filename_brand:
                print(f"brand_from_filename={used}")
            else:
                print(f"brand_used={used}")
        else:
            print("brand_cleared=yes (no source brand, no user input, no vendor default, no filename)")

    # C+- 字段策略 (Phase 1 重构: 三层优先级)
    #   优先级(高 → 低):
    #     1) --cplus-rules-file + --cplus-overrides-file  →  CplusResolver 按规则算
    #     2) --cplus                                       →  全表同一个值(老逻辑)
    #     3) 都没传                                         →  C+ 留空, supply_price 保留 vendor 提供的 (如 A1001 批发价加币)
    #   --cplus 当 L3 兜底默认值可以跟 resolver 共存(传了 rules 又传了 cplus, resolver 用 cplus 当 default)
    for row in rows:
        # 2026-10-13:不强制清空 cplus_text — vendor 已经填的 C+ (如 ALEXANDER MCQUEEN STOCK 自己设了 -20/-30) 应该保留
        if not row.get("cplus_text"):
            row["cplus_text"] = ""
        # 2026-10-12:不强制清空 supply_price — vendor 已经填的批发价/供货价应该保留
        # (如 A1001 批发价加币、Moorer 供货价)。仅在后续 cplus 逻辑计算时覆盖。
        if "supply_price" not in row:
            row["supply_price"] = None
        # 2026-08-14:统一兜底 color_code / gender(让所有 vendor 在 cn 模板下都输出这 2 列,
        #  源里有值就用源里的,没有就空字符串。Vendor 自己的 parse_rows 也可以覆盖这两个字段。)
        if "color_code" not in row:
            row["color_code"] = ""
        if "gender" not in row:
            row["gender"] = ""

    rules_path = Path(args.cplus_rules_file).expanduser().resolve() if args.cplus_rules_file else None
    overrides_path = Path(args.cplus_overrides_file).expanduser().resolve() if args.cplus_overrides_file else None
    rules_json_raw = args.cplus_rules_json
    use_resolver = bool(rules_path or overrides_path or rules_json_raw)

    # L3 兜底: --cplus (single 模式全局值) > --cplus-default (rules 模式兜底) > None
    # rules 模式下优先用 --cplus-default, 跟单值模式区分清楚
    l3_default = args.cplus_default if rules_json_raw else args.cplus
    if rules_path and not rules_json_raw:
        l3_default = args.cplus

    if use_resolver:
        # Phase 1: 用 resolver 算每行 C+
        import json
        from convert_vendor_quote_shared.cplus_resolver import CplusResolver

        if rules_json_raw:
            # 前端 JSON 优先 (Phase 1 改版格式: {field, keyword, sign, cplus})
            try:
                rules_list = json.loads(rules_json_raw)
            except (ValueError, TypeError) as e:
                print(f"ERROR: --cplus-rules-json 解析失败: {e}")
                rules_list = []
            resolver = CplusResolver.from_rules(rules_list, default_cplus=l3_default)
            print(f"cplus_resolver=active (from json)  {resolver.summary()}")
            if l3_default is not None:
                print(f"  L3 default_cplus={l3_default} (没命中规则的行用这个)")
        elif rules_path:
            # CLI 老路径: yml 文件
            resolver = CplusResolver.from_files(
                rules_file=rules_path,
                overrides_file=overrides_path,
                default_cplus=l3_default,
            )
            print(f"cplus_resolver=active (from file)  {resolver.summary()}")
            print(f"  rules_file={rules_path}")
        else:
            # 只有 overrides
            resolver = CplusResolver(
                rules=[],
                overrides={},
                default_cplus=l3_default,
            )
            print(f"cplus_resolver=active (overrides only)  {resolver.summary()}")
        if overrides_path:
            print(f"  overrides_file={overrides_path}")
        if l3_default is not None:
            print(f"  default_cplus={l3_default}")

        compute_supply = not args.no_compute_supply_price
        # 统计每条规则的命中数(用于审计)
        rule_hits: Dict[str, int] = {}
        cplus_value_hits: Dict[int, int] = {}
        none_hits = 0
        for row in rows:
            # 构造 resolver 看的 row 字典(2026-07-31 改版:加 brand / gender,3 维并列规则用)
            resolver_row = {
                "sku": row.get("sku") or "",
                "description": row.get("description") or row.get("desc") or "",
                "category": row.get("category") or "",
                "color": row.get("color") or "",
                "material": row.get("material") or "",
                "origin": row.get("origin") or "",
                "season": row.get("season") or "",
                "brand": row.get("brand") or "",
                "gender": row.get("gender") or "",
            }
            cplus_val, label = resolver.resolve(resolver_row)
            if cplus_val is None:
                none_hits += 1
                # 2026-10-13:保留 vendor 已填的 cplus_text (如 ALEXANDER MCQUEEN STOCK 自己设了 -20/-30)
                if not row.get("cplus_text"):
                    row["cplus_text"] = ""
                # 2026-10-12:不传 --cplus 时保留 vendor 已填的 supply_price (如 CANADA GOOSE 直接给批发价加币)
                # 而不是强制设 None。只有 --no-compute-supply-price 才强制设 None
                if compute_supply:
                    pass  # 保留原 supply_price (vendor 提供的 wholesale / 批发价)
                else:
                    row["supply_price"] = None
                continue
            row["cplus_text"] = str(cplus_val)
            rule_hits[label] = rule_hits.get(label, 0) + 1
            cplus_value_hits[cplus_val] = cplus_value_hits.get(cplus_val, 0) + 1
            if not compute_supply:
                row["supply_price"] = None
                continue
            cost = row.get("cost")
            if cost is None:
                row["supply_price"] = None
            else:
                try:
                    row["supply_price"] = round(float(cost) * (1 + cplus_val / 100.0), 4)
                except (TypeError, ValueError):
                    row["supply_price"] = None

        # 审计输出
        print("cplus_rule_distribution:")
        for label, cnt in sorted(rule_hits.items(), key=lambda x: (-x[1], x[0])):
            print(f"  {label}  →  {cnt} rows")
        if cplus_value_hits:
            print("cplus_value_distribution:")
            for val in sorted(cplus_value_hits.keys()):
                print(f"  C+{val:+d}  →  {cplus_value_hits[val]} rows")
        if none_hits:
            print(f"cplus_unresolved={none_hits} rows (C+ 留空)")
        if not compute_supply:
            print("compute_supply_price=off (cplus 写入,supply_price 留空)")

    elif args.cplus is not None:
        # 老逻辑: 用户填了 C+ → 写入 + 重算供货价
        # 关键:用户填 0% 也算"填了",要写 cplus_text="0"(区别于"没填")
        cplus_int = args.cplus  # e.g. 18 or -18 or 0
        cplus_text = str(cplus_int)  # 永远写,包括 0(0% 是有效值)
        cplus_frac = cplus_int / 100.0  # e.g. 0.18 or -0.18 or 0.0
        compute_supply = not args.no_compute_supply_price  # 默认 True,加 flag 时 False
        for row in rows:
            row["cplus_text"] = cplus_text
            if not compute_supply:
                # 开关关闭 → 只写 C+ 文本,留空 supply_price
                row["supply_price"] = None
                continue
            cost = row.get("cost")
            if cost is None:
                row["supply_price"] = None
            else:
                try:
                    row["supply_price"] = round(float(cost) * (1 + cplus_frac), 4)
                except (TypeError, ValueError):
                    row["supply_price"] = None
        print(f"cplus_override={cplus_int}")
        if not compute_supply:
            print("compute_supply_price=off (cplus 写入,supply_price 留空)")

    # 有图片数据行的行高：CLI --image-row-height > vendor meta > 全局默认 92pt
    # 全局默认按 image-handling.md skill 规则:所有 vendor 输出统一用 92pt。
    image_row_height = args.image_row_height
    if image_row_height is None:
        image_row_height = vendor_meta.get("default_image_row_height")
    if image_row_height is None:
        image_row_height = DEFAULT_IMAGE_ROW_HEIGHT
    print(f"image_row_height={image_row_height}pt")

    # 图片列宽：CLI --image-col-width > vendor meta > None(沿用模板)
    image_col_width = args.image_col_width
    if image_col_width is None:
        image_col_width = vendor_meta.get("default_image_col_width")
    if image_col_width is not None:
        print(f"image_col_width={image_col_width}")

    template_label = "template1-en" if template.name == "template1-en.xlsx" else "template1-cn"
    output = (
        Path(args.output).expanduser().resolve()
        if args.output
        else primary.parent.parent / "converted" / f"{primary.stem}.xlsx"
    )

    if not rows:
        raise RuntimeError("No convertible rows found. Please verify source files.")

    # 季节字段:CLI --season 非空时覆盖所有 row(用户填写优先级最高)
    # 否则保留 vendor 自己从数据源读取的值(V1042/V1044/V1049 等空字符串=都没有则为空)
    season_override = (args.season or "").strip()
    if season_override:
        for row in rows:
            row["season"] = season_override
        print(f"season_override={season_override}")

    # 每行保留几张图:CLI > vendor meta > KEEP_IMAGES_PER_ROW 默认 1
    keep_images_per_row = args.keep_images_per_row
    if keep_images_per_row is None:
        keep_images_per_row = vendor_meta.get("default_keep_images_per_row")
    if keep_images_per_row is not None:
        print(f"keep_images_per_row={keep_images_per_row}")

    write_output(
        template, output, rows,
        image_row_height=image_row_height,
        image_col_width=image_col_width,
        keep_images_per_row=keep_images_per_row,
    )
    print(f"vendor={vendor}")
    print(f"rows={len(rows)}")
    print(f"output={output}")


if __name__ == "__main__":
    main()
