#!/usr/bin/env python3
"""壯圍鄉中古透天剖切分析 — 24 個月，多維度切片。

針對「鼎弘27戶住宅新建工程」做最相近樣本剖切：
  - 主要用途 = 住家用
  - 總樓層 ≤ 3（本案 3 層）
  - 屋齡 ≤ 5 年（近新成屋）
  - 排除農業用、農舍

輸出：01_市調與定位/working/壯圍中古透天剖切_v0.1.md
"""
from __future__ import annotations
import glob
import sys
from pathlib import Path

import pandas as pd

LVR_DIR = "/tmp/lvr"
REPO = Path("/home/user/Pure-space-Branding")
OUT_MD = REPO / "01_市調與定位/working/壯圍中古透天剖切_v0.1.md"
M2_PER_PING = 3.305785


def roc_to_date(roc):
    if pd.isna(roc):
        return None
    s = str(int(roc)).zfill(7)
    try:
        return pd.Timestamp(year=int(s[:3]) + 1911, month=int(s[3:5]), day=int(s[5:7]))
    except (ValueError, OverflowError):
        return None


def load_all():
    parts = []
    for f in sorted(glob.glob(f"{LVR_DIR}/list_g1*.xls")):
        df = pd.read_excel(f, sheet_name="不動產買賣")
        if df.iloc[0].astype(str).str.contains("village|district|land sector",
                                                case=False, regex=True).any():
            df = df.iloc[1:].reset_index(drop=True)
        parts.append(df)
    df = pd.concat(parts, ignore_index=True)
    df["交易日"] = df["交易年月日"].apply(roc_to_date)
    df["建成日"] = df["建築完成年月"].apply(roc_to_date)
    df["總價"] = pd.to_numeric(df["總價(元)"], errors="coerce")
    df["建物面積_m2"] = pd.to_numeric(df["建物移轉總面積(平方公尺)"], errors="coerce")
    df["土地面積_m2"] = pd.to_numeric(df["土地移轉總面積(平方公尺)"], errors="coerce")
    df["建坪"] = df["建物面積_m2"] / M2_PER_PING
    df["地坪"] = df["土地面積_m2"] / M2_PER_PING
    df["單價_元每m2"] = pd.to_numeric(df["單價(元/平方公尺)"], errors="coerce")
    df["每坪單價_萬"] = df["單價_元每m2"] * M2_PER_PING / 10000
    df["總價_萬"] = df["總價"] / 10000
    df["屋齡"] = ((df["交易日"] - df["建成日"]).dt.days / 365.25).round(1)
    df["建物型態"] = df["建物型態"].astype(str)
    df["主要用途"] = df["主要用途"].astype(str)
    return df


def stats(df: pd.DataFrame) -> dict:
    if df.empty:
        return dict(n=0)
    pp = df["每坪單價_萬"].dropna()
    tp = df["總價_萬"].dropna()
    bp = df["建坪"].dropna()
    lp = df["地坪"].dropna()
    age = df["屋齡"].dropna()
    return dict(
        n=len(df),
        每坪中位=round(float(pp.median()), 1) if len(pp) else None,
        每坪P25=round(float(pp.quantile(0.25)), 1) if len(pp) else None,
        每坪P75=round(float(pp.quantile(0.75)), 1) if len(pp) else None,
        總價中位=round(float(tp.median()), 0) if len(tp) else None,
        建坪中位=round(float(bp.median()), 1) if len(bp) else None,
        地坪中位=round(float(lp.median()), 1) if len(lp) else None,
        屋齡中位=round(float(age.median()), 1) if len(age) else None,
    )


def main():
    df = load_all()
    latest = df["交易日"].max()
    cutoff = latest - pd.DateOffset(months=24)

    base = df[(df["鄉鎮市區"] == "壯圍鄉") &
              (df["建物型態"].str.contains("透天|別墅|獨棟", regex=True)) &
              (df["交易日"] >= cutoff) & (df["交易日"] <= latest)]

    # 切片定義
    slices = [
        ("0. 全部透天 (24M)",       base),
        ("1. 住家用",                base[base["主要用途"] == "住家用"]),
        ("2. 住家用 + 總樓層≤3",     base[(base["主要用途"] == "住家用") &
                                         (base["總樓層數"].astype(str).str.match(r"^[一二三]層$"))]),
        ("3. 住家用 + 三層",         base[(base["主要用途"] == "住家用") & (base["總樓層數"] == "三層")]),
        ("4. 住家用 + 三層 + 屋齡≤5", base[(base["主要用途"] == "住家用") &
                                          (base["總樓層數"] == "三層") &
                                          (base["屋齡"] <= 5)]),
        ("5. 住家用 + 三層 + 屋齡≤2", base[(base["主要用途"] == "住家用") &
                                          (base["總樓層數"] == "三層") &
                                          (base["屋齡"] <= 2)]),
        ("6. 住家用 + 三層 + 建坪 40~55",
            base[(base["主要用途"] == "住家用") & (base["總樓層數"] == "三層") &
                 (base["建坪"].between(40, 55))]),
        ("7. ★ 最接近本案：住家用+三層+屋齡≤5+建坪40~55",
            base[(base["主要用途"] == "住家用") & (base["總樓層數"] == "三層") &
                 (base["屋齡"] <= 5) & (base["建坪"].between(40, 55))]),
    ]

    md = []
    md.append("# 壯圍鄉中古透天 — 剖切分析 v0.1\n")
    md.append(f"- **資料來源**：內政部實價登錄 (plvr.land.moi.gov.tw)")
    md.append(f"- **資料截止**：{latest.date()}")
    md.append(f"- **分析區間**：{cutoff.date()} ~ {latest.date()}（24 個月）")
    md.append(f"- **基地比較對象**：鼎弘 27 戶新建案（住家用、3 層、建坪約 38~43 坪、地坪約 72~123 m² ≈ 22~37 坪）")
    md.append(f"- **單位**：每坪 = {M2_PER_PING} m²；單價、總價皆以萬元呈現\n")

    md.append("## 各切片彙整\n")
    md.append("| 切片 | 筆數 | 每坪中位 | P25 | P75 | 總價中位 | 建坪中位 | 地坪中位 | 屋齡中位 |")
    md.append("|------|-----:|--------:|----:|----:|--------:|--------:|--------:|--------:|")

    saved_stats = []
    for name, sub in slices:
        s = stats(sub)
        saved_stats.append((name, s))
        if s["n"] == 0:
            md.append(f"| {name} | 0 | — | — | — | — | — | — | — |")
        else:
            md.append(f"| {name} | {s['n']} | {s['每坪中位']} | {s['每坪P25']} | {s['每坪P75']} | "
                      f"{s['總價中位']} | {s['建坪中位']} | {s['地坪中位']} | {s['屋齡中位']} |")

    md.append("\n*單價：萬元/坪；總價：萬元；建坪/地坪：坪；屋齡：年*\n")

    # Sample-size warning
    final = slices[-1][1]
    if len(final) < 10:
        md.append(f"> ⚠️ 切片 7（最接近本案）僅 **{len(final)} 筆**，樣本太薄，做參考值不可作結論。")
        md.append("> 建議回到切片 4（住家用+三層+屋齡≤5）作為主要參考，已是接近本案的代理。\n")

    # ── 詳細列出切片 4 & 切片 7 的交易明細
    for slice_name, sub in [("切片 4：住家用 + 三層 + 屋齡≤5",
                              base[(base["主要用途"]=="住家用") & (base["總樓層數"]=="三層") & (base["屋齡"]<=5)]),
                             ("切片 7：最接近本案",
                              base[(base["主要用途"]=="住家用") & (base["總樓層數"]=="三層") &
                                   (base["屋齡"]<=5) & (base["建坪"].between(40,55))])]:
        md.append(f"\n## {slice_name} — 交易明細\n")
        if sub.empty:
            md.append("（無資料）")
            continue
        view_cols = ["交易日", "土地位置/建物門牌", "建物型態", "屋齡",
                     "建坪", "地坪", "每坪單價_萬", "總價_萬", "建物現況格局-房",
                     "建物現況格局-廳", "建物現況格局-衛"]
        view = sub[view_cols].copy()
        view["交易日"] = view["交易日"].dt.strftime("%Y-%m-%d")
        view["建坪"] = view["建坪"].round(1)
        view["地坪"] = view["地坪"].round(1)
        view["每坪單價_萬"] = view["每坪單價_萬"].round(1)
        view["總價_萬"] = view["總價_萬"].round(0)
        view = view.sort_values("交易日", ascending=False).head(50)
        md.append(view.to_markdown(index=False, floatfmt=".1f"))
        md.append(f"\n*顯示最近 50 筆；該切片共 {len(sub)} 筆。*\n")

    # 換算：建坪 vs 地坪 vs 容積（給本案做對齊）
    md.append("\n## 與本案對齊參考\n")
    md.append("- 本案 27 戶 **建坪** 約 38~43 坪 (126~143 m²)、**地坪** 約 22~37 坪 (72~123 m²)、**3 層住家用**")
    md.append("- 切片 4 是樣本最厚的「接近新成屋」基準；切片 7 是更窄的「同建坪區段」")
    md.append("- 估價建議：")
    md.append("    - 用切片 4 的 **每坪中位** × 建坪 = 預估總價")
    md.append("    - 上下浮動以 P25/P75 為合理區間")
    md.append("    - 地坪較大或臨路條件好者，視個案疊加 5~10%\n")

    md.append("## 缺口")
    md.append("- 本切片以中古屋 (含建照已核發) 推估新成屋走勢，**仍非預售**；預售透天請看 `宜蘭透天行情_開放資料_v0.1.md` 表 2")
    md.append("- 五結、宜蘭市同切片可同腳本跑出，但鐵則：**不**混入壯圍當主樣本\n")

    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f">> wrote {OUT_MD}", file=sys.stderr)
    for name, s in saved_stats:
        print(f"  {name}: n={s['n']}", file=sys.stderr)


if __name__ == "__main__":
    main()
