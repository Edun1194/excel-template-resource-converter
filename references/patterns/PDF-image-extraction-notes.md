# PDF 图片提取与 SKU 坐标匹配

本笔记涵盖从品牌方 PDF（如 lookbook / line-sheet / compressed inventory PDF）提取产品图片并按 SKU 匹配的工作流。适用于 PANICALE V1052 等供应商的转换任务。

---

## 工具链

| 工具 | 用途 |
|------|------|
| `pdfplumber` | 提取文本 + 精确坐标（bottom-left y-origin） |
| `PyMuPDF` (fitz) | 提取图片二进制（按 xref） |
| `PIL` | 检测图片格式（magic bytes） |
| `pdfimages`（命令行） | 快速批量导出原始图片文件（可选） |

---

## 核心坐标陷阱

**pdfplumber 和 fitz 的 y 坐标系相同**（均为 bottom-left origin，页面左下角为原点），这是最容易搞错的地方。

```
页面高度 = page.height  (pdfplumber) /  page.rect.height  (fitz)

文字顶边（从底部算）= page_height - text["top"]
图像顶边（从底部算）= page_height - bbox[3]     (bbox[3] = top in PDF coords)
```

**图像在文字上方** = 图像顶边 > 文字顶边（在 PDF 坐标系中，值越大越靠上）

---

## SKU → 图片 匹配算法

```python
import fitz  # PyMuPDF
import pdfplumber

def match_images_to_skus(pdf_path: str, sku_to_row: dict, out_dir: str):
    """
    pdf_path: PDF 文件路径
    sku_to_row: {parent_sku: excel_row_number}  映射
    out_dir: 图片输出目录
    """
    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc[page_num]
        page_h = page.rect.height

        with pdfplumber.open(pdf_path) as pv:
            pv_page = pv.pages[page_num]
            words = pv_page.extract_words()

        # 1. 按行聚类文字（y 坐标相近的为同一行）
        rows_by_y = {}
        for w in words:
            row_key = round(w["top"] / 10) * 10  # 按 10pt 分组
            rows_by_y.setdefault(row_key, []).append(w)

        # 2. 找每个 SKU 行
        for row_key, words_in_row in rows_by_y.items():
            text_line = " ".join(w["text"] for w in sorted(words_in_row, key=lambda x: x["x0"]))
            # 匹配 SKU（如 U390010PA-39U36）
            m = re.search(r'([A-Z0-9]{4,}-\d+[A-Z0-9\s]+)', text_line)
            if not m:
                continue
            matched_sku = m.group(1).strip()

            # 3. 计算该 SKU 文字的顶边（从底部算）
            text_top = page_h - min(w["top"] for w in words_in_row)

            # 4. 找在该行上方的图片（图像顶边 > 文字顶边）
            candidates = []
            for img_info in fitz_imgs:
                xref = img_info[0]
                img_data = xref_map[xref]
                # 用 fitz 页面查找图片位置（bbox）
                for img_rect in page.get_image_info(xrefs=[xref]):
                    bbox = img_rect["bbox"]   # (x0, y0, x1, y1)
                    img_top = page_h - bbox[3]
                    if img_top > text_top and (img_top - text_top) < 200:  # 200pt 上限
                        candidates.append((img_top - text_top, xref, img_data))

            if candidates:
                # 取差值最小者
                candidates.sort(key=lambda x: x[0])
                _, xref, img_data = candidates[0]
                safe_sku = re.sub(r'[/\\:*?"<>|]', '_', matched_sku)
                ext = img_data["ext"] or "jpeg"
                out_path = os.path.join(out_dir, f"{safe_sku}.{ext}")
                with open(out_path, "wb") as f:
                    f.write(img_data["image"])
```

---

## 关键陷阱

### 1. ⚠️ `get_image_info()` 和 `get_images()` 的 xref 错位陷阱（PyMuPDF 1.26.5）

**Bug 描述**：某些 PDF（含内嵌元数据图片、ICC 配置图片、Form XObject）中，`page.get_image_info(xrefs=[...])` 返回的条目数 **大于** `page.get_images()`。多余的条目是 Form XObject 等非产品图片，且 xref 与 `get_images()` 的对应关系被打乱。

**典型症状**：
- `page.get_images()` 返回 14 个条目，`page.get_image_info(xrefs=[...])` 返回 16 个条目
- 当用 `zip(fitz_items, info_list)` 配对时，后面几个图片的 xref 完全错位
- `getitem[0]`（xref）和 `info['xref']` 不一致：`fitz_items[12][0] = 60`，但 `info_list[12]['xref'] = 58`
- 用错误的 xref 调用 `doc.extract_image()` → 提取到 swatch 图而非主图，或反过来

**错误代码**（会导致 xref 错位）：

```python
# ❌ 错误：zip 配对，当两个列表长度不一致时，后面几个图片的 xref 完全错位
for i, (item, info) in enumerate(zip(fitz_items, info_list)):
    xref = item[0]                       # ← 用这个 xref 提取图片（可能是错的）
    bbox = info.get('bbox')
    ...

# ❌ 错误：直接用 fitz_items[i][0] 去 extract_image，当 info_list 有更多条目时会错位
for i, item in enumerate(fitz_items):
    xref = item[0]                       # ← 可能是错的 xref
    img_data = doc.extract_image(xref)   # ← 提取到错误图片
```

**正确代码**：

```python
# ✅ 正确：始终用 info['xref']（而非 item[0]）作为提取图片的 xref
fitz_items = page.get_images()           # [(xref, width, height, ...), ...]
xrefs = [item[0] for item in fitz_items]
info_list = page.get_image_info(xrefs=xrefs)

for i, item in enumerate(fitz_items):
    if i < len(info_list) and info_list[i]:
        info = info_list[i]
        xref = info.get('xref', item[0])  # ✅ 优先用 info 中的 xref（权威来源）
        bbox = info.get('bbox')
    else:
        xref = item[0]
        bbox = None

    # 用 xref 提取图片（用 info['xref']，不要用 item[0]）
    try:
        img_data = doc.extract_image(xref)  # ✅ 用 info['xref']
        img_bytes = img_data['image']
    except:
        img_bytes = None

    result.append({
        'xref': xref,          # ✅ 存储 info 中的 xref
        'width': item[1],
        'height': item[2],
        'bbox': bbox,
        'bytes': img_bytes,
    })
```

**为什么 `getitem[0]` 可能错误**：当 `info_list` 数量 > `fitz_items` 时，多余的 info 条目（Orphaned Form XObject）插在中间，导致 zip 配对后几个图片的 xref 被"顶替"。

**为什么 `info['xref']` 是权威来源**：`get_image_info()` 返回的每个 dict 中 `xref` 字段才是该图片在 PDF 中的真实 xref。`getitem[0]` 只是列表顺序中的序号，当有额外条目时它被"顶替"了。

**调试方法**：比较 `getitem[0]` 和 `info['xref']`，找出所有不一致的项：

```python
for i, (item, info) in enumerate(zip(fitz_items, info_list)):
    item_xref = item[0]
    info_xref = info.get('xref') if info else None
    if item_xref != info_xref:
        print(f"MISMATCH [{i}]: item_xref={item_xref}, info_xref={info_xref}")
```

### 2. `get_image_info()` 和 `get_images()` 数量可能不一致

某些 PDF（含内嵌元数据图片、ICC 配置图片）中，`page.get_image_info()` 返回的条目数 > `page.get_images()`。多余的条目是**非产品图片**（swatch 图标、装饰图等），`get_images()` 不返回它们。**只用 `get_images()` 的 xrefs 作为数据源**。

### 3. `get_image_info()` 返回的 `display_w/display_h` 可能为 None

某些 PDF 中 `info['display_w']` 和 `info['display_h']` 为 `None`，但 `bbox` 始终有值。计算图片尺寸用 bbox：

```python
bbox = info['bbox']   # (x0, y0, x1, y1)  PDF 坐标系
img_w = bbox[2] - bbox[0]
img_h = bbox[3] - bbox[1]   # ✅ 用 bbox 计算尺寸，而不是 display_w/display_h
```

### 4. SKU 含非法文件名字符

PDF 中的 SKU 可能含 `/`（如 `U391025GLB/TU-39U01B`），但 `pdfimages` 等工具导出时会将 `/` 替换为 `_` 生成文件名。

**查找时必须统一规范化**：

```python
def sanitize_filename(sku: str) -> str:
    """将 SKU 中非法文件名字符替换为下划线"""
    return re.sub(r'[/\\:*?"<>|]', '_', sku)

# 查找图片时
safe_sku = sanitize_filename(parent_sku)
img_path = os.path.join(img_dir, f"{safe_sku}.jpeg")
```

### 5. 文字顶边计算用 pdfplumber 的 `top`，不要用 `bottom`

pdfplumber 的 word 对象有 `top`（从页面顶部到文字顶边的距离）和 `bottom`。判断图片在文字上方时用 `top`：

```python
text_top_from_bottom = page_height - text["top"]   # ✅
text_bottom_from_bottom = page_height - text["bottom"]  # ❌ 会误判
```

### 6. 图片 bbox 的 y1 是顶边（不是底边）

fitz 的 `bbox = (x0, y0, x1, y1)` 中，`y1` 是图像的**顶边**（PDF 坐标系中 y 值越大越靠上）。所以：

```python
img_top = page_height - bbox[3]  # bbox[3] = y1 = 图像顶边在 PDF 坐标中的 y
```

---

## PyMuPDF 图片提取完整流程

```python
import fitz
import pdfplumber
import re
import os
from PIL import Image as PILImage
import io

def extract_and_match_pdf_images(pdf_path: str, sku_list: list, out_dir: str):
    """
    从 PDF 提取所有产品图，按 SKU 匹配并输出到 out_dir。
    sku_list: SKU 字符串列表（用于匹配）
    """
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)

    # 一次性提取所有图片 xref，避免每页重复解析
    all_fitz_imgs = doc.get_images(full=True)
    xref_to_data = {img[0]: doc.extract_image(img[0]) for img in all_fitz_imgs}

    for page_num in range(len(doc)):
        page = doc[page_num]
        page_h = page.rect.height

        # 用 get_image_info(xrefs=[...]) 获取每页图片的 bbox
        xrefs = [item[0] for item in all_fitz_imgs]
        info_list = page.get_image_info(xrefs=xrefs)

        # pdfplumber 获取文字坐标
        with pdfplumber.open(pdf_path) as pv:
            pv_page = pv.pages[page_num]
            words = pv_page.extract_words()

        for word in words:
            text = word["text"]
            # 匹配 SKU
            m = re.search(r'([A-Z0-9]{4,}-\d+)', text)
            if not m:
                continue
            sku = m.group(1).strip()

            if sku not in sku_list:
                continue

            # 计算文字顶边（从页面底部算）
            text_top = page_h - word["top"]

            # 找在该行上方的图片
            best = None
            best_dist = float("inf")
            for i, img in enumerate(all_fitz_imgs):
                xref = img[0]
                img_data = xref_to_data[xref]
                info = info_list[i]
                if not info or 'bbox' not in info:
                    continue
                bbox = info['bbox']
                img_top = page_h - bbox[3]
                dist = img_top - text_top
                if 0 < dist < 200 and dist < best_dist:
                    best_dist = dist
                    best = (xref, img_data, img_top)

            if best:
                xref, img_data, _ = best
                safe = re.sub(r'[/\\:*?"<>|]', '_', sku)
                fmt = img_data["ext"] or "jpeg"
                out_path = os.path.join(out_dir, f"{safe}.{fmt}")
                if not os.path.exists(out_path):
                    with open(out_path, "wb") as f:
                        f.write(img_data["image"])

    doc.close()
```

---

## PIL 格式检测

提取的 bytes 需要检测格式（用于 openpyxl 嵌入）：

```python
def detect_image_format(img_bytes: bytes) -> str:
    if img_bytes.startswith(b'\xff\xd8'): return 'jpeg'
    if img_bytes.startswith(b'\x89PNG'):   return 'png'
    if img_bytes.startswith(b'GIF'):       return 'gif'
    if img_bytes.startswith(b'RIFF') and img_bytes[8:12] == b'WEBP': return 'webp'
    return 'png'  # default
```

---

## 适用适配器

- **PANICALE V1052** — ⚠️ **已变更**：图片来源从 PDF 改为供应商提供的文件系统目录（`PANICALE高清图-货号+色号/`）。当前适配器使用文件系统匹配，不再从 PDF 提取。详见 `references/PANICALE-V1052-FW26-notes.md`
