#!/usr/bin/env bash
# dwg_to_json.sh — convert DWG to libredwg JSON dump
# Usage: dwg_to_json.sh input.dwg output.json
set -e
in="$1"; out="$2"
if [[ -z "$in" || -z "$out" ]]; then
  echo "Usage: $0 input.dwg output.json" >&2
  exit 1
fi
if ! command -v dwgread >/dev/null; then
  echo "dwgread not found. Install libredwg first." >&2
  echo "Build: git clone https://github.com/LibreDWG/libredwg && cd libredwg && sh autogen.sh && ./configure --disable-bindings && make -j && make install" >&2
  exit 1
fi
dwgread -O JSON -o "$out" "$in"
