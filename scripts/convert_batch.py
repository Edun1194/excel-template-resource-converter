#!/usr/bin/env python3
"""
convert_batch.py — 批量转换: 解压 zip → 遍历内部表格 → 合并输出单个模板 xlsx.

设计 (2026-10-22 用户确认):
  - 文件数上限 50 (超出直接报错, 避免 900s 超时)
  - 单文件失败不中断整批: 跳过该文件, 继续转其余, 最后汇总失败清单
  - 一次 vendor 处理 zip 内所有表格 (当前版本)
  - 递归遍历: zip 内任意层级子文件夹 (品牌分目录很常见) 都能收集到
  - 只收 .xlsx / .xls; zip 内嵌套的 zip 不展开
  - 进度通过 stdout 打 `@@PROGRESS <done> <total>` 标记行, 供 app.py 实时读
    进度条的 current 用**短名** (文件名, 重名时退回全路径), 失败清单用**全路径**
    (要能定位到是哪个文件夹)

与 convert_vendor_quote.py 的区别:
  单文件 1 进 1 出; 本脚本多文件合并 1 个输出.

用法:
  python3 convert_batch.py \
    --zip-file "path/to/batch.zip" \
    --vendor V1003-期货 \
    --output out.xlsx \
    [--language cn] [--template template1-cn.xlsx] [--cplus 18] \
    [--brand GUCCI] [--season FW27] [--no-compute-supply-price] \
    [--cplus-rules-json '...'] [--cplus-default 15]
"""
from __future__ import annotations

import argparse
import concurrent.futures
import json
import os
import re
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# 与 convert_vendor_quote 共用 shared 工具
from convert_vendor_quote_shared import (
    ASSETS_DIR,
    pick_template,
    write_output,
)
from convert_vendor_quote import _load_adapter

# zip 内允许处理的表格类型 (2026-10-22 用户确认: **只收 .xls / .xlsx, 不处理 PDF**)
#
# 注意: 单文件上传 (/api/convert) 仍支持 .pdf —
#       目前只有 V1003-期货 (MAMMUT LOOKBOOK 图册) 能读 PDF,
#       该文件仍可单独上传转换, 只是不能放进 zip 走批量。
ALLOWED_EXT = {".xlsx", ".xls"}

MAX_FILES = 50

PROGRESS_MARK = "@@PROGRESS"

# 解压根目录下直接躺着的表格归入这一组 (单独出一份表)
ROOT_GROUP = "根目录"

# ---- 计时 (排查性能问题用) ----
_T0 = time.time()
_STAGE: Dict[str, float] = {}


def _tick(stage: str) -> None:
    """记录阶段耗时 (累积), 配合 _dump_timing 输出."""
    global _T0
    _STAGE[stage] = _STAGE.get(stage, 0.0) + (time.time() - _T0)
    _T0 = time.time()


def _dump_timing() -> None:
    total = sum(_STAGE.values())
    parts = " ".join(f"{k}={v:.1f}s" for k, v in _STAGE.items())
    print(f"[batch] 耗时: {parts} 合计 {total:.1f}s", flush=True)


class _Progress:
    """线程安全的进度计数器 (各组并行时共用)."""

    def __init__(self, total: int):
        self.total = total
        self.done = 0
        self._lock = threading.Lock()

    def tick(self, current: str) -> None:
        with self._lock:
            self.done += 1
            n = self.done
        _emit_progress(n, self.total, current)


def _resolve_brand(adapter, rows, f: Path, args) -> str:
    """品牌兜底 (与 convert_vendor_quote.main() 逻辑一致)."""
    src_brand = ""
    for r in rows:
        b = str(r.get("brand") or "").strip()
        if b:
            src_brand = b
            break
    default_brand = None
    try:
        meta = adapter.get_vendor_meta() or {}
        default_brand = meta.get("default_brand")
    except Exception:
        pass
    filename_brand = f.stem
    if args.brand:
        return args.brand
    if src_brand:
        return src_brand
    if default_brand:
        return default_brand
    return filename_brand or ""


def _process_group(
    group_key: str,
    gfiles: List[Path],
    args,
    template_path: Path,
    gpath: Path,
    extract_dir: Path,
    order: Dict[Path, int],
    short_labels: List[str],
) -> Dict[str, object]:
    """处理一个组: 解析 → 品牌/季节 → C+ → 写表. 供并行调用.

    每组完全独立 (各自 load adapter / 各自 write_output), 组失败不影响其它组。
    必须是**模块级函数**且参数可 pickle —— 进程池要用。
    """
    ok_files: List[str] = []
    failed_files: List[Dict[str, str]] = []
    group_rows: List[Dict[str, object]] = []
    group_ok = 0

    try:
        adapter = _load_adapter(args.vendor, gfiles[0])
    except Exception as exc:
        adapter = None
        print(f"[batch] 组 [{group_key}] 适配器加载失败: {exc}", file=sys.stderr)

    for f in gfiles:
        idx = order[f]
        label = _safe_stem(f, extract_dir)
        short = short_labels[idx - 1]
        try:
            if adapter is None:
                raise RuntimeError("适配器加载失败")
            rows = adapter.parse_rows(f, None, args.vendor)
            if not rows:
                raise ValueError("未解析出任何数据行")

            brand = _resolve_brand(adapter, rows, f, args)
            for r in rows:
                r["brand"] = brand
                if args.season and not r.get("season"):
                    r["season"] = args.season

            group_rows.extend(rows)
            ok_files.append(label)
            group_ok += 1
            print(f"[batch] ({idx}) ✓ {label} → {len(rows)} 行", flush=True)
        except Exception as exc:
            failed_files.append({"file": label, "error": str(exc)[:300]})
            print(f"[batch] ({idx}) ✗ {label} 失败: {exc}", file=sys.stderr, flush=True)

    if not group_rows:
        print(f"[batch] 组 [{group_key}] 没有可用数据, 不出表", file=sys.stderr)
        return {"group": group_key, "ok_files": ok_files, "failed_files": failed_files,
                "rows": 0, "files": group_ok, "written": False, "filename": ""}

    # C+ 处理 (复用主流程语义: vendor 已填则保留)
    _apply_cplus(group_rows, args)

    # ---- 写这一组的表 ----
    t_write = time.time()
    try:
        write_output(
            template_path,
            gpath,
            group_rows,
            image_row_height=args.image_row_height,
            image_col_width=args.image_col_width,
            keep_images_per_row=args.keep_images_per_row,
        )
    except Exception as exc:
        print(f"[batch] 组 [{group_key}] 写输出失败: {exc}", file=sys.stderr)
        failed_files.append({"file": gpath.name, "error": f"写输出失败: {exc}"[:300]})
        return {"group": group_key, "ok_files": ok_files, "failed_files": failed_files,
                "rows": 0, "files": group_ok, "written": False, "filename": ""}
    dt_write = time.time() - t_write
    if dt_write > 10:
        print(f"[batch]   ⚠️ 组 [{group_key}] 写表耗时 {dt_write:.1f}s "
              f"({len(group_rows)} 行) — 图片多/行数大时这里最慢",
              file=sys.stderr, flush=True)

    return {"group": group_key, "ok_files": ok_files, "failed_files": failed_files,
            "rows": len(group_rows), "files": group_ok, "written": True,
            "filename": gpath.name}


def _emit_progress(done: int, total: int, current: str) -> None:
    """打进度标记行 (flush, 保证 app.py 能实时读到)."""
    payload = json.dumps({"done": done, "total": total, "current": current},
                         ensure_ascii=False)
    print(f"{PROGRESS_MARK} {payload}", flush=True)


def _is_safe_member(name: str) -> bool:
    """防 zip-slip: 拒绝绝对路径 / 越界路径 / macOS 资源分支垃圾."""
    if name.startswith("/") or name.startswith("\\"):
        return False
    # 路径穿越
    parts = Path(name).parts
    if ".." in parts:
        return False
    # macOS __MACOSX / ._xxx 垃圾文件
    if name.startswith("__MACOSX/") or "/__MACOSX/" in name:
        return False
    if Path(name).name.startswith("._"):
        return False
    # Windows 盘符 (C:)
    if len(name) > 1 and name[1] == ":":
        return False
    return True


def _collect_files(root: Path) -> List[Path]:
    """递归收集可转换文件, 按相对路径排序保证结果稳定."""
    out: List[Path] = []
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        if p.name.startswith("."):     # .DS_Store / 临时文件
            continue
        if p.suffix.lower() not in ALLOWED_EXT:
            continue
        out.append(p)
    return out


def _safe_stem(p: Path, root: Path) -> str:
    """生成可读且不冲突的标识 (含相对路径, 避免同名文件互相覆盖)."""
    try:
        rel = p.relative_to(root)
    except ValueError:
        rel = Path(p.name)
    parts = [x for x in rel.parts if x not in ("", ".", "..")]
    # 去掉扩展名后用 ' - ' 连接
    stem_parts = []
    for i, part in enumerate(parts):
        if i == len(parts) - 1:
            stem_parts.append(Path(part).stem)
        else:
            stem_parts.append(part)
    return " - ".join(s for s in stem_parts if s)


def _short_labels(files: List[Path], root: Path) -> List[str]:
    """给每个文件生成进度条用的短名 (2026-10-23).

    zip 内常按品牌分文件夹 (BMB/ BMC/ BMO/...), 完整相对路径很长,
    在进度条里显示 "外层 - 品牌 - 分类 - 文件名" 会挤爆一行, 所以进度条
    只显示文件名 (去扩展名)。

    但不同文件夹可能有同名文件 (如 BMB/RTW.xlsx 和 BMO/RTW.xlsx),
    这种短名重名的情况退回用完整相对路径, 保证用户在进度条里能分辨。
    失败清单始终用完整相对路径 (要能定位到是哪个文件夹的)。
    """
    stems = [p.stem for p in files]
    dupes = {s for s in stems if stems.count(s) > 1}
    out: List[str] = []
    for p, stem in zip(files, stems):
        if stem in dupes:
            out.append(_safe_stem(p, root))      # 重名 → 用完整路径
        else:
            out.append(stem)                     # 唯一 → 只用文件名
    return out


def _apply_cplus(rows: List[Dict[str, object]], args) -> None:
    """C+ / 供货价处理 (复用 convert_vendor_quote.main() 的语义).

    规则: vendor 自己填了 cplus_text / supply_price 就保留, 没填才用全局参数算。
    """
    compute_supply = not args.no_compute_supply
    if args.cplus is not None:
        cplus_int = args.cplus
        cplus_frac = cplus_int / 100.0
        for r in rows:
            if not r.get("cplus_text"):
                r["cplus_text"] = str(cplus_int)
            if compute_supply:
                cost = r.get("cost")
                try:
                    r["supply_price"] = (round(float(cost) * (1 + cplus_frac), 4)
                                          if cost else None)
                except (TypeError, ValueError):
                    r["supply_price"] = None
            else:
                r["supply_price"] = None
    else:
        for r in rows:
            if not r.get("cplus_text"):
                r["cplus_text"] = ""
            if "supply_price" not in r:
                r["supply_price"] = None


def _group_by_dir(files: List[Path], root: Path) -> Dict[str, List[Path]]:
    """按目录把文件分组, 每组独立出一份表 (2026-10-26 用户需求).

    分组规则 (按用户确认):
      - 取文件相对解压根的**目录段** parts
      - len(parts)==0  → 文件直接躺在解压根 → 单独一组「根目录」
      - len(parts)==1  → 顶层文件夹下的文件 → 该文件夹自成一组
      - len(parts)>=2  → 顶层文件夹下面的**直接子文件夹**各自成一组

    举例 (Boss & Hugo Pre FW27.zip, 解压根下是 Boss & Hugo Pre FW27/):
      Boss & Hugo Pre FW27/BMB/BMB B&H FA27.xlsx  → 组 "Boss & Hugo Pre FW27/BMB"
      Boss & Hugo Pre FW27/BMC/BMC CDF FA27.xlsx  → 组 "Boss & Hugo Pre FW27/BMC"
      → 共 8 组, 出 8 份表格

    举例 (Calvin Klein zip, 文件直接躺在一层):
      Calvin Klein Underwear Pre FW27/CK NOS PW27.xlsx → 组 "Calvin Klein Underwear Pre FW27"
      → 1 组, 出 1 份表格
    """
    groups: Dict[str, List[Path]] = {}
    for p in files:
        try:
            rel = p.relative_to(root)
        except ValueError:
            rel = Path(p.name)
        parts = [x for x in rel.parts[:-1] if x not in ("", ".", "..")]
        if not parts:
            key = ROOT_GROUP
        elif len(parts) == 1:
            key = parts[0]
        else:
            key = parts[0] + "/" + parts[1]
        groups.setdefault(key, []).append(p)
    return groups


def _group_output_name(group_key: str) -> str:
    """组名 → 输出文件名 (不含扩展名)."""
    name = group_key.replace("/", " - ")
    name = re.sub(r"[\\:*?\"<>|]+", "_", name).strip()
    name = re.sub(r"\s+", " ", name)
    if not name:
        name = ROOT_GROUP
    # 文件名过长会被截断/写失败, 留出 "_converted.xlsx" 的余量
    if len(name) > 200:
        name = name[:200].rstrip()
    return f"{name}_converted"


def _fix_zipinfo_name(info: "zipfile.ZipInfo") -> str:
    """还原 zip 内中文文件名 (2026-10-22).

    zipfile 对**未标记 UTF-8 flag (0x800)** 的条目一律按 cp437 解码,
    而 Windows / 部分 Mac 打包工具写的中文名是 UTF-8 字节但没打 flag,
    于是解出来是 'σÑ╜τÜä1.xlsx' 这种乱码。
    这里用 cp437 反编码回原始字节, 再按 UTF-8 解.
    """
    if info.flag_bits & 0x800:      # 已声明 UTF-8, 直接用
        return info.filename
    try:
        return info.filename.encode("cp437").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return info.filename         # 本来就不是 UTF-8, 保持原样


def _extract_safely(z: "zipfile.ZipFile", dest: Path) -> None:
    """解压, 跳过不安全条目, 并修正中文文件名编码."""
    for info in z.infolist():
        if info.is_dir():
            continue
        if not _is_safe_member(info.filename):
            print(f"[batch] 跳过不安全条目: {info.filename}", file=sys.stderr)
            continue

        # 修正文件名编码 (中文乱码)
        fixed_name = _fix_zipinfo_name(info)
        if fixed_name != info.filename and not _is_safe_member(fixed_name):
            print(f"[batch] 修正后路径不安全, 跳过: {info.filename}", file=sys.stderr)
            continue

        new_info = zipfile.ZipInfo(fixed_name, date_time=info.date_time)
        new_info.compress_type = info.compress_type
        new_info.external_attr = info.external_attr
        new_info.flag_bits = info.flag_bits | 0x800     # 标记为 UTF-8

        target = (dest / fixed_name).resolve()
        try:
            target.relative_to(dest.resolve())           # 二次确认不越界
        except ValueError:
            print(f"[batch] 路径越界, 跳过: {fixed_name}", file=sys.stderr)
            continue

        try:
            target.parent.mkdir(parents=True, exist_ok=True)   # 嵌套目录要先建
            with z.open(info) as src, open(target, "wb") as out:
                while True:
                    chunk = src.read(1 << 20)
                    if not chunk:
                        break
                    out.write(chunk)
        except Exception as exc:
            print(f"[batch] 解压失败 {fixed_name}: {exc}", file=sys.stderr)


def main() -> None:
    p = argparse.ArgumentParser(description="Batch-convert all spreadsheets inside a zip")
    p.add_argument("--zip-file", required=True)
    p.add_argument("--vendor", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--language", choices=["cn", "en"], default="cn")
    p.add_argument("--template")
    p.add_argument("--warehouse-id", dest="warehouse_id")
    p.add_argument("--cplus", type=int)
    p.add_argument("--cplus-rules-json", dest="cplus_rules_json")
    p.add_argument("--cplus-default", dest="cplus_default", type=int)
    p.add_argument("--cplus-overrides-file", dest="cplus_overrides_file")
    p.add_argument("--no-compute-supply-price", dest="no_compute_supply",
                   action="store_true")
    p.add_argument("--brand")
    p.add_argument("--season")
    p.add_argument("--image-row-height", dest="image_row_height", type=float)
    p.add_argument("--image-col-width", dest="image_col_width", type=float)
    p.add_argument("--keep-images-per-row", dest="keep_images_per_row", type=int)
    p.add_argument("--merge-all", dest="merge_all", action="store_true",
                   help="旧行为: 整个 zip 合成一份表 (默认按目录分组, 每组一份)")
    p.add_argument("--workers", type=int, default=0,
                   help="并行进程数 (0=自动, 默认 min(8, 核数-2))")
    args = p.parse_args()

    zip_path = Path(args.zip_file).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not zip_path.exists():
        print(f"[batch] zip 不存在: {zip_path}", file=sys.stderr)
        sys.exit(2)
    if not zipfile.is_zipfile(zip_path):
        print("[batch] 不是有效的 zip 文件", file=sys.stderr)
        sys.exit(2)

    tmp_root = Path(tempfile.mkdtemp(prefix="convert_batch_"))
    extract_dir = tmp_root / "extracted"
    extract_dir.mkdir(parents=True, exist_ok=True)

    # ---- 1. 解压 ----
    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            _extract_safely(z, extract_dir)
    except zipfile.BadZipFile as exc:
        print(f"[batch] zip 损坏: {exc}", file=sys.stderr)
        sys.exit(2)
    _tick("解压")

    # ---- 2. 收集文件 ----
    files = _collect_files(extract_dir)
    if not files:
        print("[batch] zip 内没有可转换的表格文件", file=sys.stderr)
        print("@@RESULT " + json.dumps(
            {"ok": False, "total": 0, "ok_files": [], "failed_files": [],
             "total_rows": 0, "error": "zip 内没有可转换的表格文件 (仅支持 .xlsx / .xls)"},
            ensure_ascii=False), flush=True)
        sys.exit(1)

    if len(files) > MAX_FILES:
        msg = (f"zip 内文件数 {len(files)} 超过上限 {MAX_FILES}，"
               f"请拆分成多个批次上传")
        print(f"[batch] {msg}", file=sys.stderr)
        print("@@RESULT " + json.dumps(
            {"ok": False, "total": len(files), "ok_files": [], "failed_files": [],
             "total_rows": 0, "error": msg}, ensure_ascii=False), flush=True)
        sys.exit(1)

    total = len(files)

    # ---- 4. 分组 (2026-10-26): 顶层文件夹下的直接子文件夹各自成一组 ----
    groups = _group_by_dir(files, extract_dir)
    if args.merge_all:
        groups = {"全部": files}
    print(f"[batch] vendor={args.vendor}, 共 {total} 个文件, 分 {len(groups)} 组",
          flush=True)
    for gk in groups:
        print(f"[batch]   组 [{gk}] → {len(groups[gk])} 个文件", flush=True)
    _tick("收集分组")

    # ---- 5. 模板 ----
    if args.template:
        template_path = Path(args.template).expanduser().resolve()
    else:
        template_path = pick_template(args.language)

    out_dir = tmp_root / "outputs"
    out_dir.mkdir(parents=True, exist_ok=True)

    outputs: List[Dict[str, object]] = []      # 每组一份表的输出信息
    ok_files: List[str] = []
    failed_files: List[Dict[str, str]] = []
    total_rows = 0
    done = 0

    # label = 完整相对路径 (失败清单/日志用, 能定位文件夹)
    # short = 短名 (进度条用, 唯一时只显示文件名)
    short_labels = _short_labels(files, extract_dir)
    order = {f: i for i, f in enumerate(files, start=1)}

    # ---- 6. 逐组转换 (并行: 每组独立, 组失败不影响其它组) ----
    # 2026-10-26: 改成并行。串行时"写表"占 68% 总耗时, 而各组之间完全独立,
    # 并行能把这部分压到 ~1/N。写表是 IO + 释放 GIL 的活 (PIL 解码), 线程池够用。
    used_names: Dict[str, int] = {}          # 清洗后重名的组名 → 计数
    gitems: List[Tuple[str, List[Path], str, Path]] = []
    for group_key, gfiles in groups.items():
        gname = _group_output_name(group_key)
        if gname in used_names:
            used_names[gname] += 1
            gname = f"{gname}_{used_names[gname]}"
        else:
            used_names[gname] = 1
        gitems.append((group_key, gfiles, gname, out_dir / f"{gname}.xlsx"))

    prog = _Progress(total)
    cpu = os.cpu_count() or 4
    # 进程池优先 (真多核): openpyxl 写表是纯 Python, 线程池被 GIL 串行化反而更慢
    #
    # worker 数要**给 gunicorn 留出核**: service 里是 `gunicorn -w 2`, 也就是说
    # 常驻 2 个 gunicorn worker 进程, 它们各自还会再 fork 出本进程池的子进程。
    # 如果按核数-2 算 (4 核 → 2), 加上 gunicorn 的 2 个就是 4~6 个计算进程抢 4 核,
    # 一旦有并发请求就会严重超订, 反而比串行慢。
    # 所以这里按 核数//2 算 (4 核 → 2, 8 核 → 4), 上限 4。
    # 部署时可用环境变量 BATCH_WORKERS 或 --workers 精确指定。
    env_workers = os.environ.get("BATCH_WORKERS", "").strip()
    if env_workers.isdigit() and int(env_workers) > 0:
        want = int(env_workers)
    elif args.workers:
        want = args.workers
    else:
        want = min(4, max(1, cpu // 2))
    max_workers = max(1, min(want, len(gitems)))
    use_proc = max_workers > 1 and len(gitems) > 1
    print(f"[batch] {'进程池' if use_proc else '单进程'}处理 {len(gitems)} 组 "
          f"(worker={max_workers}/{cpu} 核)", flush=True)

    results: Dict[str, Dict[str, object]] = {}
    try:
        if use_proc:
            import multiprocessing as mp
            # fork: 复用已 import 的模块, 免去 spawn 重新解释器的开销
            ctx = mp.get_context("fork")
            with concurrent.futures.ProcessPoolExecutor(
                max_workers=max_workers, mp_context=ctx
            ) as ex:
                fut_map = {
                    ex.submit(_process_group, gk, gf, args, template_path, gp,
                              extract_dir, order, short_labels): gk
                    for gk, gf, _gn, gp in gitems
                }
                for fut in concurrent.futures.as_completed(fut_map):
                    gk = fut_map[fut]
                    try:
                        results[gk] = fut.result()
                    except Exception as exc:
                        results[gk] = {
                            "group": gk, "ok_files": [], "rows": 0, "files": 0,
                            "written": False, "filename": "",
                            "failed_files": [{"file": gk, "error": str(exc)[:300]}],
                        }
        else:
            for gk, gf, _gn, gp in gitems:
                try:
                    results[gk] = _process_group(
                        gk, gf, args, template_path, gp, extract_dir, order, short_labels)
                except Exception as exc:
                    results[gk] = {
                        "group": gk, "ok_files": [], "rows": 0, "files": 0,
                        "written": False, "filename": "",
                        "failed_files": [{"file": gk, "error": str(exc)[:300]}],
                    }
    except Exception as exc:
        print(f"[batch] 并行执行失败, 回退单进程: {exc}", file=sys.stderr)
        results = {}
        for gk, gf, _gn, gp in gitems:
            try:
                results[gk] = _process_group(
                    gk, gf, args, template_path, gp, extract_dir, order, short_labels)
            except Exception as exc2:
                results[gk] = {
                    "group": gk, "ok_files": [], "rows": 0, "files": 0,
                    "written": False, "filename": "",
                    "failed_files": [{"file": gk, "error": str(exc2)[:300]}],
                }
    _tick("并行转换")

    # 按组顺序汇总 (保证输出顺序稳定)
    outputs: List[Dict[str, object]] = []
    ok_files: List[str] = []
    failed_files: List[Dict[str, str]] = []
    total_rows = 0
    for gk, gf, _gn, _gp in gitems:
        r = results.get(gk) or {}
        ok_files.extend(r.get("ok_files") or [])
        failed_files.extend(r.get("failed_files") or [])
        prog.done += len(gf)
        if r.get("written"):
            total_rows += int(r.get("rows") or 0)
            outputs.append({
                "group": gk,
                "filename": r.get("filename"),
                "rows": int(r.get("rows") or 0),
                "files": int(r.get("files") or 0),
            })
            print(f"[batch] 组 [{gk}] → {r.get('filename')} "
                  f"({r.get('rows')} 行)", flush=True)
    _emit_progress(prog.done, total, "")

    if not outputs:
        msg = f"全部 {total} 个文件都转换失败，没有可输出的数据"
        print(f"[batch] {msg}", file=sys.stderr)
        print("@@RESULT " + json.dumps(
            {"ok": False, "total": total, "ok_files": ok_files,
             "failed_files": failed_files, "total_rows": 0, "outputs": [],
             "error": msg}, ensure_ascii=False), flush=True)
        sys.exit(1)

    # ---- 7. 打包成一个 zip ----
    # compresslevel=1: xlsx 内部已经是 DEFLATE 压缩过的 (媒体是 jpeg/png),
    # 再压一遍几乎没收益, 用最快的级别省 CPU (服务器上 CPU 紧张时有用)。
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED,
                             compresslevel=1) as zf:
            for info in outputs:
                zf.write(out_dir / str(info["filename"]), arcname=str(info["filename"]))
    except Exception as exc:
        print(f"[batch] 打包失败: {exc}", file=sys.stderr)
        print("@@RESULT " + json.dumps(
            {"ok": False, "total": total, "ok_files": ok_files,
             "failed_files": failed_files, "total_rows": 0, "outputs": outputs,
             "error": f"打包失败: {exc}"}, ensure_ascii=False), flush=True)
        sys.exit(1)
    _tick("打包")
    _dump_timing()

    print(f"[batch] 完成: {len(ok_files)}/{total} 文件成功, "
          f"{len(failed_files)} 失败, 共 {len(outputs)} 份表格 / {total_rows} 行", flush=True)
    print("@@RESULT " + json.dumps({
        "ok": True,
        "total": total,
        "ok_files": ok_files,
        "failed_files": failed_files,
        "total_rows": total_rows,
        "outputs": outputs,
        "group_count": len(outputs),
    }, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()