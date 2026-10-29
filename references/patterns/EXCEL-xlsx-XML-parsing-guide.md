# Excel xlsx XML Parsing Guide

When converting vendor files that openpyxl reads too slowly (e.g. 7000+ row files in read_only mode), parse the underlying XML directly via `zipfile + xml.etree.ElementTree`.

## Shared Strings — MUST Use ET, Not Regex

The naive regex approach `re.findall(r'<t[^>]*>(.*?)</t>', ss_xml)` FAILS on files whose `<si>` blocks contain self-closing `<t/>` tags or multiple `<t>` elements per `<si>`.

**Example of the failure:**
```
<!-- sharedStrings.xml -->
<si><t>Panoramica ordini</t></si>
<si><t/></si>                    ← empty string, but still counts as index 31
<si><t>Apparel</t></si>
```
Regex would give indices [0, '', 'Apparel'] — off by one from index 31 onward.

**Correct approach — ET:**
```python
import xml.etree.ElementTree as ET
ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ss_root = ET.fromstring(ss_xml)
ss_list = []
for si in ss_root.findall(f'{ns}si'):
    parts = []
    for t in si.findall(f'{ns}t'):
        if t.text:
            parts.append(t.text)
    ss_list.append(''.join(parts))
```

## Sheet Data — Iterate ET Row Elements

**Regex cell parsing fails for self-closing cells:**
- `<c r="Q11" s="14"/>` (empty qty cell) — self-closing tag, no `</c>`
- The pattern `<c r="...">(.*?)</c>` would match across cell boundaries

**Correct approach — ET row iteration:**
```python
sheet_root = ET.fromstring(sheet_xml)
for row_elem in sheet_root.iter(f'{ns}row'):
    r = int(row_elem.get('r', 0))
    if r < DATA_START_ROW:
        continue
    col_vals = {}
    for cell in row_elem:
        ref = cell.get('r', '')         # e.g. 'B11'
        col_letters = ''.join(ch for ch in ref if ch.isalpha())  # 'B'
        col_num = col_letter_to_num(col_letters)
        t_attr = cell.get('t', '')       # 's' = shared string
        v_elem = cell.find(f'{ns}v')
        v = v_elem.text if v_elem is not None else None
        if t_attr == 's' and v is not None:
            col_vals[col_num] = ss_list[int(v)]
        elif v is not None:
            col_vals[col_num] = v
        # else: formula cell with no cached value → skip
```

**Helper:**
```python
def col_letter_to_num(col: str) -> int:
    num = 0
    for ch in col:
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num
```

## Column Number Reference

| Letter | Number |
|--------|--------|
| A      | 1      |
| B      | 2      |
| ...    | ...    |
| Z      | 26     |
| AA     | 27     |

## Image EMU Source: Check Both `<xdr:ext>` and `<a:ext>`

MAMMUT and similar files store image EMU dimensions inside `<a:xfrm><a:ext cx="..." cy="..."/>` NOT `<xdr:ext>`:

```python
# Check spPr block (contains <a:ext> with DrawingML namespace)
ext_m2 = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"', block)
# Also check for <xdr:ext> as fallback
ext_m = re.search(r'<xdr:ext\s+cx="(\d+)"\s+cy="(\d+)"', block)
```

EMUs to pixels: `pixels = EMU / 9525`
Target row height: 92pt = 92 * 9525 ≈ 876,300 EMU

## Performance Comparison

| Method | 7700-row file |
|--------|--------------|
| openpyxl read_only=True | >300s (timeout) |
| zipfile + ET (this guide) | ~5s |
