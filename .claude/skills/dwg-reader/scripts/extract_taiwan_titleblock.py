#!/usr/bin/env python3
"""Extract Taiwan construction-drawing title block data from a libredwg JSON dump.

Usage: extract_taiwan_titleblock.py <dump.json> [--anchor "鼎弘健康生技"]

Reads the JSON produced by `dwgread -O JSON`, finds repeating per-unit title
blocks anchored by a fixed text (typically the 起造人/owner-company line), and
prints a markdown summary with each unit's site area, floor area, building
area, coverage ratio and FAR ratio.
"""
import argparse, json, re, sys
from collections import defaultdict


def clean(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"\\[A-Za-z][^;\\]*;", "", s)
    s = re.sub(r"\\[A-Za-z]\d*", "", s)
    s = s.replace("\\P", "\n").replace("\\~", " ")
    return "".join(c for c in s if c not in "{}").strip()


def load_texts(path: str):
    with open(path, encoding="utf-8", errors="replace") as f:
        data = json.load(f)
    texts = []
    for o in data["OBJECTS"]:
        e = o.get("entity")
        if e in ("TEXT", "MTEXT"):
            ins = o.get("ins_pt", [0, 0])
            raw = o.get("text_value", "") or o.get("text", "")
            t = clean(raw)
            if t:
                texts.append((float(ins[0]), float(ins[1]), t))
    return texts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json")
    ap.add_argument("--anchor", default="鼎弘健康生技",
                    help="substring identifying the per-unit owner row")
    ap.add_argument("--unit-pattern", default=r"起造人\s*\(([A-G]\d+)戶\)",
                    help="regex with one group capturing the unit label")
    args = ap.parse_args()

    texts = load_texts(args.json)
    anchors = [(x, y) for x, y, t in texts if args.anchor in t]
    print(f"Found {len(anchors)} title blocks", file=sys.stderr)

    def nearby(ax, ay):
        return [(x, y, t) for x, y, t in texts
                if ax - 1500 <= x <= ax + 900 and ay - 4500 <= y <= ay + 200]

    pat = re.compile(args.unit_pattern)
    out = []
    for ax, ay in anchors:
        items = nearby(ax, ay)
        rows = defaultdict(list)
        for x, y, t in items:
            rows[round(y, -1)].append((x, t))

        unit = next((pat.match(t).group(1) for x, y, t in items if pat.match(t)), None)
        if not unit:
            continue
        lot = next((t for x, y, t in items if "地號" in t and "鄉" in t and len(t) < 60), None)
        cov = next((t for x, y, t in items if "%%%" in t and "60%%%" in t), None)
        far = next((t for x, y, t in items if "%%%" in t and "180%%%" in t), None)

        floors = {}
        for label, key in [("一  層", "1F"), ("二  層", "2F"), ("三  層", "3F")]:
            for x, y, t in items:
                if t == label:
                    row = rows[round(y, -1)]
                    vals = sorted(rt for rx, rt in row if re.match(r"^[\d.]+m2$", rt))
                    if vals:
                        floors[key] = vals[0]
                    break

        out.append({
            "unit": unit, "lot": lot,
            "floors": floors,
            "coverage": (cov or "").replace("%%%", "%"),
            "far": (far or "").replace("%%%", "%"),
        })

    out.sort(key=lambda r: r["unit"])
    print("| Unit | Lot | 1F | 2F | 3F | Coverage | FAR |")
    print("|------|-----|----|----|----|----------|-----|")
    for r in out:
        f = r["floors"]
        cov = re.sub(r"\s*<.*", "", r["coverage"])
        far = re.sub(r"\s*<.*", "", r["far"])
        print(f"| {r['unit']} | {r['lot'] or '—'} | "
              f"{f.get('1F','—')} | {f.get('2F','—')} | {f.get('3F','—')} | "
              f"{cov or '—'} | {far or '—'} |")


if __name__ == "__main__":
    main()
