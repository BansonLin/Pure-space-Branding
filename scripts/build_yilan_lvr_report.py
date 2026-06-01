#!/usr/bin/env python3
"""Yilan real-estate transaction analysis from MOI LVR open data.

Inputs: list_g1NNNNNNN_NNNNNNNN.xls (Yilan county quarterly bundles, MOI 內政部
        實價登錄 open data).
Outputs:
    01_市調與定位/working/宜蘭透天行情_開放資料_v0.1.md
    01_市調與定位/working/competitors.csv

Rules enforced:
    * Government open data only.
    * 24-month window from latest 交易年月日 in the dataset.
    * 壯圍鄉 if N<30 → flagged "樣本不足，僅供參考"; do NOT mix with neighbours.
    * Every table cites: 資料截止 + 筆數 + 來源.
"""
from __future__ import annotations
import glob
import os
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd

LVR_DIR = "/tmp/lvr"
REPO = Path("/home/user/Pure-space-Branding")
OUT_MD = REPO / "01_市調與定位/working/宜蘭透天行情_開放資料_v0.1.md"
OUT_CSV = REPO / "01_市調與定位/working/competitors.csv"

SOURCE_URL = "https://plvr.land.moi.gov.tw/DownloadOpenData (內政部不動產實價登錄)"
TARGETS = ["壯圍鄉", "五結鄉", "宜蘭市"]
TRANSOM_TYPES_RE = re.compile(r"透天|別墅|獨棟")
M2_PER_PING = 3.305785


def roc_to_date(roc: str | float) -> pd.Timestamp | None:
    """Convert ROC date integers like 1141201 → 2025-12-01."""
    if pd.isna(roc):
        return None
    s = str(int(roc)).zfill(7)
    try:
        y = int(s[:3]) + 1911
        m = int(s[3:5])
        d = int(s[5:7])
        return pd.Timestamp(year=y, month=m, day=d)
    except (ValueError, OverflowError):
        return None


def load_sheet(sheet: str) -> pd.DataFrame:
    """Concatenate same-named sheet from all year files; drop the second
    Chinese header row each file ships with."""
    parts = []
    for f in sorted(glob.glob(f"{LVR_DIR}/list_g1*.xls")):
        xls = pd.ExcelFile(f)
        if sheet not in xls.sheet_names:
            continue
        df = pd.read_excel(xls, sheet_name=sheet)
        # Files have a second header row in English (e.g. "The villages and towns urban district")
        if df.iloc[0].astype(str).str.contains("village|district|land sector",
                                                case=False, regex=True).any():
            df = df.iloc[1:].reset_index(drop=True)
        df["__source_file__"] = os.path.basename(f)
        parts.append(df)
    return pd.concat(parts, ignore_index=True)


def build_unified(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    """Add 交易日 / 建坪 / 每坪單價 / 屋齡 columns."""
    df = df.copy()
    df["交易日"] = df["交易年月日"].apply(roc_to_date)
    df["總價"] = pd.to_numeric(df["總價(元)"], errors="coerce")
    df["建物面積_m2"] = pd.to_numeric(df["建物移轉總面積(平方公尺)"], errors="coerce")
    df["土地面積_m2"] = pd.to_numeric(df["土地移轉總面積(平方公尺)"], errors="coerce")
    df["建坪"] = df["建物面積_m2"] / M2_PER_PING
    df["地坪"] = df["土地面積_m2"] / M2_PER_PING
    df["單價_元每m2"] = pd.to_numeric(df["單價(元/平方公尺)"], errors="coerce")
    df["每坪單價_萬"] = df["單價_元每m2"] * M2_PER_PING / 10000

    # 屋齡 — only for second-hand
    if kind == "second":
        comp = df["建築完成年月"].apply(roc_to_date)
        df["屋齡_年"] = ((df["交易日"] - comp).dt.days / 365.25).round(1)

    df["建物型態"] = df["建物型態"].astype(str)
    df["鄉鎮市區"] = df["鄉鎮市區"].astype(str)
    return df


def desc_stats(x: pd.Series, ndigits: int = 1) -> dict:
    s = pd.to_numeric(x, errors="coerce").dropna()
    if s.empty:
        return dict(n=0, median=None, mean=None, p25=None, p75=None, min=None, max=None)
    return dict(
        n=int(s.size),
        median=round(float(s.median()), ndigits),
        mean=round(float(s.mean()), ndigits),
        p25=round(float(s.quantile(0.25)), ndigits),
        p75=round(float(s.quantile(0.75)), ndigits),
        min=round(float(s.min()), ndigits),
        max=round(float(s.max()), ndigits),
    )


def main():
    print(">> Loading second-hand transactions…", file=sys.stderr)
    sec = build_unified(load_sheet("不動產買賣"), kind="second")
    print(">> Loading pre-sale transactions…", file=sys.stderr)
    pre = build_unified(load_sheet("預售屋買賣"), kind="pre")

    # data span
    latest = max(sec["交易日"].max(), pre["交易日"].max())
    cutoff_24m = latest - pd.DateOffset(months=24)
    earliest_sec = sec["交易日"].min()
    print(f">> Data range second-hand: {earliest_sec.date()} → {latest.date()}", file=sys.stderr)
    print(f">> 24-month window:       {cutoff_24m.date()} → {latest.date()}", file=sys.stderr)

    # 24-month filter
    sec_w = sec[(sec["交易日"] >= cutoff_24m) & (sec["交易日"] <= latest)].copy()
    pre_w = pre[(pre["交易日"] >= cutoff_24m) & (pre["交易日"] <= latest)].copy()

    # transom filter
    sec_th = sec_w[sec_w["建物型態"].str.contains(TRANSOM_TYPES_RE)].copy()
    pre_th = pre_w[pre_w["建物型態"].str.contains(TRANSOM_TYPES_RE)].copy()

    # ── Table 1 — Second-hand transom transactions, by district
    print(">> Building table 1 (second-hand)…", file=sys.stderr)
    rows1 = []
    for region in TARGETS:
        r = sec_th[sec_th["鄉鎮市區"] == region]
        price_per_ping = desc_stats(r["每坪單價_萬"], 1)
        total = desc_stats(r["總價"] / 10000, 0)  # 萬
        ping = desc_stats(r["建坪"], 1)
        age = desc_stats(r["屋齡_年"], 1)
        rows1.append(dict(
            區=region,
            筆數=int(price_per_ping["n"]),
            每坪中位數=price_per_ping["median"],
            每坪P25=price_per_ping["p25"],
            每坪P75=price_per_ping["p75"],
            每坪最低=price_per_ping["min"],
            每坪最高=price_per_ping["max"],
            總價中位數=total["median"],
            總價P25=total["p25"],
            總價P75=total["p75"],
            建坪中位數=ping["median"],
            屋齡中位數=age["median"],
        ))
    t1 = pd.DataFrame(rows1)

    # ── Table 2 — Pre-sale transactions in Yilan transom
    print(">> Building table 2 (pre-sale)…", file=sys.stderr)
    # Pre-sale 建商欄位有時叫「備註」內含建商名；確切欄位是「編號」前綴。先用全部欄位顯示。
    pre_cols = pre_th.columns.tolist()
    # Identify potential pre-sale brand/case fields
    cand_brand = [c for c in pre_cols if "建商" in c or "起造" in c or "申請" in c]
    cand_case = [c for c in pre_cols if "案名" in c or "建案" in c]
    print(f"   候選建商欄位: {cand_brand}", file=sys.stderr)
    print(f"   候選案名欄位: {cand_case}", file=sys.stderr)

    rows2 = []
    for region in TARGETS:
        r = pre_th[pre_th["鄉鎮市區"] == region]
        rows2.append(dict(
            區=region,
            筆數=len(r),
            每坪中位數_萬=round(float(r["每坪單價_萬"].median()), 1) if len(r) else None,
            總價中位數_萬=round(float((r["總價"]/10000).median()), 0) if len(r) else None,
            建坪中位數=round(float(r["建坪"].median()), 1) if len(r) else None,
        ))
    t2_summary = pd.DataFrame(rows2)

    # Per-deal pre-sale list (all Yilan transom)
    presale_cols = ["鄉鎮市區", "建案名稱", "土地位置/建物門牌", "交易日", "建物型態",
                    "建坪", "每坪單價_萬", "總價", "棟及號", "編號"]
    pre_th_simple = pre_th[[c for c in presale_cols if c in pre_th.columns]].copy()
    pre_th_simple["總價_萬"] = (pre_th_simple["總價"]/10000).round(0)
    pre_th_simple["每坪單價_萬"] = pre_th_simple["每坪單價_萬"].round(1)
    pre_th_simple["建坪"] = pre_th_simple["建坪"].round(1)
    pre_th_simple["交易日"] = pre_th_simple["交易日"].dt.strftime("%Y-%m-%d")

    # ── Table 3 — population & income — leave NA with a note. We do not have that file.
    pop_note = ("> 人口/家戶/所得資料：本批次未含 data.gov.tw 鄉鎮人口與綜所稅"
                "申報資料集。資料缺口。請另行下載：\n"
                ">  - 內政部戶政司 鄉鎮人口統計\n"
                ">  - 財政部財政資訊中心 綜所稅申報初步核定統計專冊（鄉鎮市區別）")

    # ── Generate competitors.csv from pre-sale records — aggregate by 建案名稱
    print(">> Building competitors.csv…", file=sys.stderr)
    if not pre_th.empty and "建案名稱" in pre_th.columns:
        comp = pre_th.copy()
        comp["建案名稱"] = comp["建案名稱"].fillna("(未揭露)").astype(str).str.strip()
        agg = comp.groupby(["鄉鎮市區", "建案名稱"]).agg(
            戶數=("編號", "count"),
            首交易=("交易日", "min"),
            末交易=("交易日", "max"),
            建坪_中位=("建坪", "median"),
            每坪單價_萬_中位=("每坪單價_萬", "median"),
            每坪單價_萬_最低=("每坪單價_萬", "min"),
            每坪單價_萬_最高=("每坪單價_萬", "max"),
            總價_萬_中位=("總價", lambda s: round(float(s.median())/10000)),
            門牌樣本=("土地位置/建物門牌", lambda s: s.dropna().iloc[0] if s.dropna().size else ""),
            棟號樣本=("棟及號", lambda s: ";".join(s.dropna().astype(str).unique()[:5])[:120]),
            備註樣本=("備註", lambda s: ";".join(s.dropna().astype(str).unique()[:2])[:200]),
            建物型態=("建物型態", lambda s: s.mode().iloc[0] if not s.empty else ""),
            主要建材=("主要建材", lambda s: s.mode().iloc[0] if not s.empty else ""),
            解約筆數=("解約情形", lambda s: s.dropna().astype(str).ne("").sum()),
        ).reset_index()
        agg["建坪_中位"] = agg["建坪_中位"].round(1)
        agg["每坪單價_萬_中位"] = agg["每坪單價_萬_中位"].round(1)
        agg["每坪單價_萬_最低"] = agg["每坪單價_萬_最低"].round(1)
        agg["每坪單價_萬_最高"] = agg["每坪單價_萬_最高"].round(1)
        agg["focus_target"] = agg["鄉鎮市區"].isin(TARGETS)
        agg = agg.sort_values(["focus_target", "鄉鎮市區", "戶數"],
                              ascending=[False, True, False]).reset_index(drop=True)
        agg["首交易"] = pd.to_datetime(agg["首交易"]).dt.strftime("%Y-%m-%d")
        agg["末交易"] = pd.to_datetime(agg["末交易"]).dt.strftime("%Y-%m-%d")
        agg.insert(0, "資料來源", "MOI實價登錄(預售屋)")
        agg.insert(0, "type", "presale_project")
        comp_df = agg
    else:
        comp_df = pd.DataFrame()

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    comp_df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f">> wrote {OUT_CSV} ({len(comp_df)} rows)", file=sys.stderr)

    # ── Compose Markdown ────────────────────────────────────────────────
    md = []
    md.append("# 宜蘭透天行情 — 開放資料 v0.1\n")
    md.append(f"- **資料來源**：{SOURCE_URL}")
    md.append(f"- **檔案範圍**：宜蘭縣 110/01 ~ 115/05（含 6 個年度批次）")
    md.append(f"- **資料截止日**：{latest.date().isoformat()}（資料中最後一筆交易日）")
    md.append(f"- **24 個月分析區間**：{cutoff_24m.date().isoformat()} ~ {latest.date().isoformat()}")
    md.append(f"- **建物型態篩選**：建物型態欄含「透天/別墅/獨棟」")
    md.append(f"- **行政區**：壯圍鄉（主）、五結鄉、宜蘭市（參考）")
    md.append(f"- **單位**：每坪 = {M2_PER_PING} m²；單價、總價皆以萬元呈現\n")

    # Sample-size guard
    zw_n = int(t1.loc[t1["區"]=="壯圍鄉", "筆數"].iloc[0])
    if zw_n < 30:
        md.append(f"> ⚠️ **壯圍鄉 24 個月透天成交僅 {zw_n} 筆，樣本不足。**"
                  " 數字僅供參考，**不**以鄰區 (五結/宜蘭市) 樣本補齊。\n")

    # Table 1
    md.append("## 表 1 — 中古透天成交行情（24 個月）\n")
    md.append("單價單位：萬元/坪；總價單位：萬元；建坪：坪；屋齡：年。\n")
    cols = ["區","筆數","每坪中位數","每坪P25","每坪P75","每坪最低","每坪最高",
            "總價中位數","總價P25","總價P75","建坪中位數","屋齡中位數"]
    md.append(t1[cols].to_markdown(index=False, floatfmt=".1f"))
    md.append(f"\n*資料截止：{latest.date()}｜總筆數（三區）：{int(t1['筆數'].sum())}*\n")

    # Table 2
    md.append("## 表 2 — 預售透天案件（24 個月）\n")
    md.append("### 2a. 區彙整\n")
    md.append(t2_summary.to_markdown(index=False, floatfmt=".1f"))
    md.append(f"\n*資料截止：{latest.date()}｜總筆數（三區）：{int(t2_summary['筆數'].sum())}*\n")

    md.append("### 2b. 預售透天案件 — 焦點三區（依建案名稱聚合）\n")
    md.append("> 預售屋實價登錄表自 111 年起揭露「建案名稱」欄。建商名仍不揭露。\n")
    if not comp_df.empty:
        focus = comp_df[comp_df["focus_target"]].copy()
        view = focus[["鄉鎮市區","建案名稱","戶數","建坪_中位",
                     "每坪單價_萬_中位","每坪單價_萬_最低","每坪單價_萬_最高",
                     "總價_萬_中位","首交易","末交易","門牌樣本"]].copy()
        view.columns = ["區","案名","戶數","建坪",
                       "每坪中位","每坪低","每坪高",
                       "總價中位","首交易","末交易","代表門牌"]
        md.append(view.to_markdown(index=False, floatfmt=".1f"))
        md.append(f"\n*焦點三區共 {len(view)} 個預售透天建案；其他鄉鎮預售透天 {len(comp_df)-len(view)} 案於 competitors.csv（focus_target=False）*\n")
    else:
        md.append("（無資料）\n")

    # Table 3
    md.append("## 表 3 — 區域人口、家戶、所得\n")
    md.append(pop_note + "\n")

    md.append("## 4. 缺口與限制\n")
    if zw_n < 30:
        md.append(f"- **壯圍鄉中古透天樣本不足** ({zw_n} 筆 < 30)：見表 1 警示。")
    else:
        md.append(f"- **壯圍鄉中古透天樣本**：{zw_n} 筆，分布跨度大（建坪中位 {t1.loc[t1['區']=='壯圍鄉','建坪中位數'].iloc[0]} 坪、屋齡中位 {t1.loc[t1['區']=='壯圍鄉','屋齡中位數'].iloc[0]} 年）。"
                  "若要對應本案（新建透天、3 層、無車位），建議進一步剖切：")
        md.append("    - 屋齡 ≤ 5 年（接近新成屋）")
        md.append("    - 主要用途=住家用、車位數=0")
        md.append("    - 總樓層 ≤ 3 層")
    md.append("- **預售建商**：實價登錄揭露**案名**但**不揭露建商**。如需建商名稱，需另查「營建署建照核發資料集」(data.gov.tw dataset 6303) 或 591 等商業站；商業站涉反爬機制，本批次未做。")
    md.append("- **人口/所得**：本壓縮包未含 data.gov.tw 鄉鎮人口、財政部綜所稅申報資料。表 3 待另行下載：")
    md.append("    - 內政部戶政司「鄉鎮市區人口指標」")
    md.append("    - 財政部財政資訊中心「綜所稅申報初步核定統計專冊」")
    md.append("- **異常極值**：表 1 每坪最高/最低差異極大（如宜蘭市每坪 0.4 ~ 561.7 萬），含部分共有人交易、夾雜土地交易、瑕疵戶等。中位數與 P25/P75 已可吸收，但若做散佈圖建議剔除 P1/P99。\n")

    md.append("## 5. 處理腳本\n")
    md.append("- `scripts/build_yilan_lvr_report.py`（本檔生成）")
    md.append("- 來源檔放置位置：`/tmp/lvr/list_g1100101_1101231.xls` ~ `list_g1150101_1150530.xls`\n")

    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f">> wrote {OUT_MD}", file=sys.stderr)


if __name__ == "__main__":
    main()
