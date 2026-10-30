"""category-futures 共用工具 (2026-10-27).

这一组文件来自 6 个不同品牌、不同版式的期货/订货单, 但有几个共同点:
  - 表头列数、列序都不一样 (同一份 NuORDER, ACNE / BC / POLO RL 三家列序都不一样)
  - 所以**一律按表头名定位列**, 绝不写死列号
  - 都要按尺码展开成多行
  - 都要能取到内嵌图片 (ws._images)

本模块提供:
  - open_safe_workbook(): 处理 openpyxl 读不了的 sheetProtection (insertLinks 属性)
  - find_header_row(): 按表头名特征找表头行
  - header_col_map(): 表头名 → 列号
  - size_labels_row(): 尺码标签行 → [(列号, 尺码标签)]
  - extract_images(): 按行抽内嵌图片
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl


# ---------------- 读取 ----------------

def open_safe_workbook(path: Path, data_only: bool = True):
    """打开 workbook, 兼容 openpyxl 读不了的各种非标准 xlsx.

    部分供应商导出的 xlsx 里有 openpyxl 不认识的属性/枚举值, 直接 load_workbook 会抛:
      - <sheetProtection insertLinks="1" .../> → TypeError: unexpected keyword 'insertLinks'
      - <dataValidation type="..."> 非法枚举    → ValueError: Value must be one of {...}

    这里复制一份临时文件, 把这些标签摘掉再读; 还不行就退回 read_only 模式
    (read_only 能读出单元格值, 只是拿不到 ws._images, 图片再走 zip 解析)。
    """
    try:
        return openpyxl.load_workbook(path, data_only=data_only)
    except Exception:
        pass

    tmp = Path(tempfile.mkdtemp(prefix="catfut_"))
    try:
        fixed = tmp / "fixed.xlsx"
        with zipfile.ZipFile(path, "r") as zin, \
                zipfile.ZipFile(fixed, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                    txt = data.decode("utf-8", "ignore")
                    txt = re.sub(r"<sheetProtection\b[^>]*/>", "", txt)
                    txt = re.sub(r"<sheetProtection\b.*?</sheetProtection>", "", txt,
                                 flags=re.S)
                    txt = re.sub(r"<dataValidations\b.*?</dataValidations>", "", txt,
                                 flags=re.S)
                    txt = re.sub(r"<dataValidation\b[^>]*/>", "", txt)
                    data = txt.encode("utf-8")
                zout.writestr(item, data)
        try:
            wb = openpyxl.load_workbook(fixed, data_only=data_only)
            shutil.rmtree(tmp, ignore_errors=True)     # 普通模式读完即关, 可删
            return wb
        except Exception:
            pass
        # 最后兜底: read_only (文件要留着, 不能删)
        wb = openpyxl.load_workbook(fixed, data_only=data_only, read_only=True)
        return wb
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise



# ---------------- 表头定位 ----------------

def norm_header(v: object) -> str:
    """表头名归一化: 小写 + 去空格/连字符/下划线/括号.

    去括号是因为同一字段在不同导出里写法不同:
    'Wholesale (EUR)' / 'Wholesale Price (EUR)' / 'Sugg. Retail (EUR)'
    归一化后分别是 wholesaleeur / wholesalepriceeur / suggretaileur,
    所以各模块要把变体都列进别名里。
    """
    if v is None:
        return ""
    s = re.sub(r"[\s\-_/()\[\]]+", "", str(v).strip().lower())
    return s



def find_header_row(
    ws,
    required: Sequence[str],
    max_scan: int = 15,
    min_hits: Optional[int] = None,
) -> Optional[int]:
    """在前 max_scan 行里找表头行 —— **必须同时含全部 required 名称**才认.

    严格匹配是必须的: 'Style Number' 这类词在多个版式里都有
    (Linesheet / NuORDER 都用), 宽松匹配会让两个版式互相误判。
    """
    want = [norm_header(x) for x in required]
    limit = min(ws.max_row, max_scan)
    for r in range(1, limit + 1):
        row_vals = set()
        for c in range(1, min(ws.max_column, 80) + 1):
            k = norm_header(ws.cell(r, c).value)
            if k:
                row_vals.add(k)
        if all(any(w in v for v in row_vals) for w in want):
            return r
    return None



def header_col_map(
    ws,
    header_row: int,
    aliases: Dict[str, Sequence[str]],
    max_col: Optional[int] = None,
) -> Dict[str, int]:
    """表头行 → {字段: 列号}. aliases: 字段 → 该字段可能的表头名(多个都接受).

    一个字段的所有别名都找不到时, 该字段不出现 (调用方按"源文件没有就留空"处理)。
    """
    limit = max_col or ws.max_column
    found: Dict[str, int] = {}
    for field, names in aliases.items():
        targets = {norm_header(n) for n in names}
        for c in range(1, limit + 1):
            k = norm_header(ws.cell(header_row, c).value)
            if k and k in targets and field not in found:
                found[field] = c
                break
    return found


def cellv(ws, row: int, col: Optional[int]):
    """按列号取值, col 为 None 返回 None."""
    if not col:
        return None
    return ws.cell(row, col).value


# ---------------- 尺码标签 ----------------

def size_labels_row(
    ws,
    row: int,
    start_col: int,
    stop_col: Optional[int] = None,
    stop_labels: Sequence[str] = ("total", "units", "wholesale", "sugg", "retail",
                                 "styleminimum", "colorminimum", "discount",
                                 "finalunit", "totalwholesale", "totalretail",
                                 "price", "cost"),
) -> List[Tuple[int, str]]:
    """从某一行读尺码标签列, 返回 [(列号, 标签)].

    start_col: 尺码标签起始列
    stop_col: 尺码标签结束列(不含); None = 遇到空白/终止标签为止
    stop_labels: 遇到这些表头就停 (如 'TOTAL' / 'Wholesale (EUR)')
    """
    limit = stop_col or ws.max_column
    out: List[Tuple[int, str]] = []
    for c in range(start_col, limit + 1):
        v = ws.cell(row, c).value
        if v is None:
            if out:
                break
            continue
        s = str(v).strip()
        if not s:
            if out:
                break
            continue
        low = norm_header(s)
        if any(low.startswith(norm_header(x)) for x in stop_labels):
            break
        out.append((c, s))
    return out


def is_size_like(label: str) -> bool:
    """判断一个表头标签像不像尺码 (XXS / 32 / 36.5 / 1C / 29/30 ...)."""
    s = str(label).strip()
    if not s:
        return False
    if norm_header(s) in ("size", "sizes", "availablesizes", "bulks"):
        return False
    # 纯数字(含小数) 或 数字/字母/斜杠 组合
    if re.match(r"^\d+(\.\d+)?$", s):
        return True
    if re.match(r"^[A-Za-z]{1,4}(\s*/\s*[A-Za-z0-9]{1,4})?$", s):
        return True
    # 数字打头 + 字母: 3X / 3XL / 4XL / 5XL / 6XL …
    # (2026-10-30: 原来只认 `^\d+[A-Za-z]$`, 漏了 '3XL'/'4XL' 这类双字母尺码)
    if re.match(r"^\d+[A-Za-z]{1,2}$", s):
        return True
    return False


# ---------------- 图片 ----------------

def extract_images(ws, col_priority: Optional[Dict[int, int]] = None) -> Dict[int, List[Dict[str, object]]]:
    """抽内嵌图片 → {1-idx 行号: [{bytes, format}, ...]}.

    col_priority: {0-idx 列: 优先级}, 小的优先; 没给则按列号升序。
    """
    out: Dict[int, List[Dict[str, object]]] = {}
    try:
        images = list(ws._images)
    except Exception:
        return out
    for img in images:
        anchor = getattr(img, "anchor", None)
        if not hasattr(anchor, "_from") or anchor._from is None:
            continue
        try:
            blob = img._data() if hasattr(img, "_data") else None
        except Exception:
            blob = None
        if not blob:
            continue
        out.setdefault(anchor._from.row + 1, []).append({
            "bytes": blob,
            "format": (img.format or "jpeg"),
            "_col": anchor._from.col,
        })
    for r in out:
        if col_priority:
            out[r].sort(key=lambda x: (col_priority.get(x["_col"], 99), x["_col"]))
        else:
            out[r].sort(key=lambda x: x["_col"])
        for it in out[r]:
            it.pop("_col", None)
    return out


def load_images_by_row_zip(path: Path) -> Dict[int, List[Dict[str, object]]]:
    """直接从 xlsx 的 zip 里解析 drawing, 按 0-idx 行分桶.

    有些供应商导出的 xlsx, openpyxl 的 ws._images 会返回 0 张 (虽然 zip 里
    明明有 xl/media/image*.jpeg) —— 绕开 openpyxl 的 drawing 解析, 自己读 XML。

    Returns: {row_0idx -> [{bytes, format}]}
    """
    import xml.etree.ElementTree as ET
    import os as _os

    images_by_row: Dict[int, List[Dict[str, object]]] = {}
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            draw_names = sorted(n for n in names
                                if n.startswith("xl/drawings/drawing")
                                and n.endswith(".xml"))
            if not draw_names:
                return images_by_row
            NS = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
            NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

            for dname in draw_names:
                rels_name = (dname.rsplit("/", 1)[0] + "/_rels/"
                             + dname.rsplit("/", 1)[1] + ".rels")
                if rels_name not in names:
                    continue
                try:
                    rels_root = ET.fromstring(z.read(rels_name).decode("utf-8", "ignore"))
                    id_to_target = {r.get("Id"): r.get("Target")
                                    for r in rels_root if r.get("Id")}
                    root = ET.fromstring(z.read(dname).decode("utf-8", "ignore"))
                except Exception:
                    continue

                for anchor in (list(root.findall(f"{NS}oneCellAnchor"))
                               + list(root.findall(f"{NS}twoCellAnchor"))):
                    from_el = anchor.find(f"{NS}from")
                    if from_el is None:
                        continue
                    try:
                        row = int(from_el.find(f"{NS}row").text)
                    except (AttributeError, TypeError, ValueError):
                        continue
                    blip = anchor.find(f".//{NS_A}blip")
                    embed_id = blip.get(f"{NS_R}embed") if blip is not None else None
                    if not embed_id or embed_id not in id_to_target:
                        continue
                    target = id_to_target[embed_id].lstrip("./").lstrip("../")
                    full_path = target if target.startswith("xl/") else "xl/" + target
                    if full_path not in names:
                        continue
                    blob = z.read(full_path)
                    ext = _os.path.splitext(full_path)[1].lower()
                    fmt = "jpeg" if ext in (".jpg", ".jpeg") else "png"
                    images_by_row.setdefault(row, []).append({
                        "bytes": blob, "format": fmt,
                    })
    except Exception:
        return images_by_row
    return images_by_row

