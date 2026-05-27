---
name: dwg-reader
description: Read and extract information from AutoCAD DWG/DXF files. Use when the user provides a .dwg or .dxf file (CAD drawing) and asks to inspect, parse, or extract structured data — text labels, dimensions, layer info, blocks, geometric entities, drawing metadata. Handles DWG files via libredwg conversion to DXF, then parses with ezdxf.
---

# DWG/DXF Reader

## Overview
This skill reads AutoCAD CAD files (DWG/DXF) and extracts structured information.

## Tooling
- `dwg2dxf` (from libredwg, installed at /usr/local/bin) — converts DWG → DXF
- `dwgread` (from libredwg) — dumps DWG content in various formats (JSON, GeoJSON, etc.)
- `ezdxf` (Python, installed via pip) — parses DXF files into structured Python objects
- `dxfgrabber` (Python fallback) — alternative DXF parser

## Workflow

### Step 1: Convert DWG to DXF (if input is DWG)
```bash
dwg2dxf -v0 /path/to/input.dwg -o /tmp/output.dxf
```

### Step 2: Parse with ezdxf
```python
import ezdxf
doc = ezdxf.readfile("/tmp/output.dxf")
msp = doc.modelspace()

# DWG/DXF version
print(doc.dxfversion, doc.acad_release)

# Layers
for layer in doc.layers:
    print(layer.dxf.name, layer.color)

# Iterate entities
for e in msp:
    print(e.dxftype(), getattr(e.dxf, 'layer', ''))

# Extract all text (TEXT + MTEXT) — most useful for building info
for e in msp.query("TEXT MTEXT"):
    text = e.dxf.text if e.dxftype() == "TEXT" else e.text
    pos = e.dxf.insert
    print(f"[{e.dxf.layer}] @({pos.x:.1f},{pos.y:.1f}): {text}")

# Dimensions
for d in msp.query("DIMENSION"):
    print(d.dxf.text, d.dxf.actual_measurement)

# Block references (INSERT)
for ins in msp.query("INSERT"):
    print(ins.dxf.name, ins.dxf.insert)

# Drawing extents / units
print(doc.header.get("$EXTMIN"), doc.header.get("$EXTMAX"))
print("Units:", doc.header.get("$INSUNITS"))
```

### Step 3: For Taiwan / Chinese-language construction drawings
Encoding pitfalls:
- DXF text may be in Big5, CP950, UTF-8, or mojibake'd UTF-8 — try multiple decodes
- MText control codes like `\\fSimSun;` or `\\C1;` need stripping (use `ezdxf.tools.text.MTextEditor` or regex `\\\\[A-Za-z][^;]*;`)
- Common building info to look for in TEXT/MTEXT:
  - 案名/建案名稱 (project name)
  - 基地/地號 (lot number)
  - 樓層 (floor)
  - 戶別/單元 (unit type)
  - 面積/坪數 (area in 坪 = 3.305785 m²)
  - 比例尺/SCALE
  - 建蔽率/容積率 (coverage / FAR)
  - 用途 (usage)

### Step 4: Helpful extraction patterns
```python
# Group text by layer to find title-block, dimension, area-label layers
from collections import defaultdict
by_layer = defaultdict(list)
for e in msp.query("TEXT MTEXT"):
    txt = e.dxf.text if e.dxftype()=="TEXT" else e.text
    by_layer[e.dxf.layer].append(txt)

# Closed polylines = room/unit boundaries; area via shoelace
import numpy as np
for pl in msp.query("LWPOLYLINE"):
    if pl.closed:
        pts = np.array([(p[0], p[1]) for p in pl.get_points("xy")])
        area = 0.5 * abs(np.dot(pts[:,0], np.roll(pts[:,1], -1)) - np.dot(pts[:,1], np.roll(pts[:,0], -1)))
        # area is in drawing units squared
```

## Alternative: JSON dump
```bash
dwgread -O JSON -o /tmp/out.json /path/input.dwg
```
Outputs a large JSON with every entity — useful for grep'ing specific values.

## Notes
- DWG version supported: R13 through 2018 (Autodesk 2013-2017 format works)
- Always use `dwg2dxf -v0` to suppress verbose logging
- Large drawings (>10MB) can produce huge DXF — stream parse with `ezdxf.iterdxf` if memory matters
