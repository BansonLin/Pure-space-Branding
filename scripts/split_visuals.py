#!/usr/bin/env python3
"""分割中道森活主視覺 8 色版本圖 + 抓出純背景色."""
from pathlib import Path
from PIL import Image, ImageStat
import json

SRC = "/root/.claude/uploads/e7b1fd1c-d52e-46dc-9b9a-b4c9d54183c0/38fe0ccf-557748.jpg"
OUT_DIR = Path("/home/user/Pure-space-Branding/assets/visuals")
OUT_DIR.mkdir(parents=True, exist_ok=True)

img = Image.open(SRC).convert("RGB")
W, H = img.size
COLS, ROWS = 2, 4
cell_w = W // COLS
cell_h = H // ROWS
MARGIN = 4

# 採樣的純底色位置：上方中央偏右（文字上方），通常是純底色無花
def sample_bg(crop):
    w, h = crop.size
    samples = []
    # 三個位置：左上、上中、右上中間（避開角落花朵與中間文字）
    for fx, fy in [(0.40, 0.10), (0.55, 0.10), (0.50, 0.06)]:
        patch = crop.crop((int(w*fx)-15, int(h*fy)-10, int(w*fx)+15, int(h*fy)+10))
        avg = ImageStat.Stat(patch).mean
        samples.append(tuple(int(c) for c in avg[:3]))
    # 取三個樣本的中位
    r = sorted(s[0] for s in samples)[1]
    g = sorted(s[1] for s in samples)[1]
    b = sorted(s[2] for s in samples)[1]
    return r, g, b

results = []
for row in range(ROWS):
    for col in range(COLS):
        idx = row * COLS + col + 1
        x0 = col * cell_w + MARGIN
        y0 = row * cell_h + MARGIN
        x1 = (col + 1) * cell_w - MARGIN
        y1 = (row + 1) * cell_h - MARGIN
        crop = img.crop((x0, y0, x1, y1))
        r, g, b = sample_bg(crop)
        hex_color = f"#{r:02X}{g:02X}{b:02X}"
        path = OUT_DIR / f"v{idx}.jpg"
        crop.save(path, quality=92)
        results.append({"id": idx, "hex": hex_color, "rgb": [r, g, b],
                        "file": path.name, "size": crop.size})
        print(f"  v{idx}: {hex_color}  RGB({r},{g},{b})")

# Save color metadata
meta_path = OUT_DIR / "colors.json"
meta_path.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
print(f"\nWrote metadata → {meta_path}")
