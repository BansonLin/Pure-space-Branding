#!/usr/bin/env python3
"""人口/家戶/所得資料下載器（網路解禁或本機跑時可用）。

執行時讀 SOURCES 列表逐一下載到 /tmp/demographics/，產生
01_市調與定位/working/壯圍人口所得_v0.1.md。

如果環境網路被封，會顯示「需要的下載清單」供使用者人工下載後 --from-local。
"""
from __future__ import annotations
import argparse, os, shutil, sys, urllib.request, urllib.error
from pathlib import Path

REPO = Path("/home/user/Pure-space-Branding")
OUT_DIR = Path("/tmp/demographics")
OUT_MD = REPO / "01_市調與定位/working/壯圍人口所得_v0.1.md"

SOURCES = [
    # (name, url, parser_hint)
    ("鄉鎮市區人口指標",
     "https://www.ris.gov.tw/rs-opendata/api/v1/datastore/ODRP019/113",
     "MOI Household Registration — 鄉鎮人口數、家戶、年齡結構"),
    ("綜所稅申報初步核定統計專冊_鄉鎮市區",
     "https://www.fia.gov.tw/WEB/fia/ias/ias3/100all.zip",
     "Ministry of Finance Fiscal Information Agency — 鄉鎮綜所稅"),
    ("宜蘭縣壯圍鄉戶政月報",
     "https://www.ris.gov.tw/rs-opendata/api/v1/datastore/ODRP054/114",
     "MOI Household Registration — 月別人口數"),
]


def try_download(url: str, out: Path) -> tuple[bool, str]:
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "PureSpace-Branding/1.0 (research; non-commercial; respect robots.txt)"
        })
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = resp.read()
        out.write_bytes(data)
        return True, f"OK ({len(data):,} bytes)"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}"
    except urllib.error.URLError as e:
        return False, f"URLError: {e.reason}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from-local", action="store_true",
                    help="Skip download, expect files under /tmp/demographics")
    args = ap.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    results = []
    for name, url, hint in SOURCES:
        target = OUT_DIR / f"{name}{Path(url).suffix or '.bin'}"
        if args.from_local:
            results.append((name, target.exists(), "local"))
            continue
        ok, msg = try_download(url, target)
        results.append((name, ok, msg))
        print(f"{'✅' if ok else '❌'} {name}: {msg}", file=sys.stderr)

    # Render whatever we have. For now just print the catalogue & blockers.
    md = ["# 壯圍人口、家戶、所得 v0.1 (download log)\n"]
    md.append("## 來源清單\n")
    md.append("| # | 資料集 | URL | 用途 | 下載狀態 |")
    md.append("|---|--------|-----|------|---------|")
    for i, ((name, url, hint), (_, ok, msg)) in enumerate(zip(SOURCES, results), 1):
        status = "✅ " + msg if ok else "❌ " + msg
        md.append(f"| {i} | {name} | `{url}` | {hint} | {status} |")
    md.append("")

    if not all(r[1] for r in results):
        md.append("## 後續步驟（網路封鎖時）\n")
        md.append("環境 sandbox 不允許 .moi.gov.tw / .fia.gov.tw / data.gov.tw。請：")
        md.append("1. 在本機下載上表 URL → ZIP/CSV")
        md.append("2. 把檔案上傳到對話，或放到 `/tmp/demographics/`，重跑 `python3 scripts/fetch_demographics.py --from-local`")
        md.append("3. 我會接手把鄉鎮級欄位過濾出來、生成表 3\n")

    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f">> wrote {OUT_MD}", file=sys.stderr)


if __name__ == "__main__":
    main()
