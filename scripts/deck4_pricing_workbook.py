"""簡報 4：中道森活 27 戶單戶定價試算簡報 + Excel 試算表.

依 DWG 提取的 27 戶基地、建坪、建築面積，搭配三情境單價策略，
給出每一戶的建議售價區間。

輸出：
  02_簡報/04_中道森活_27戶定價試算簡報.pptx
  02_簡報/04_中道森活_27戶定價試算表.xlsx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt
import pandas as pd

REPO = Path("/home/user/Pure-space-Branding")
OUT_PPT = REPO / "02_簡報" / "04_中道森活_27戶定價試算簡報.pptx"
OUT_XLSX = REPO / "02_簡報" / "04_中道森活_27戶定價試算表.xlsx"
OUT_PPT.parent.mkdir(parents=True, exist_ok=True)
TOTAL = 10

# 27 戶資料（取自 DWG 報告）
# 欄位：戶, 地號, 基地m², 樓地板m², 建築面積m², 建蔽率%
UNITS = [
    ("A1", "789-32", 96, 126.76, 43.16, 44.96),
    ("A2", "789-31", 72, 126.10, 42.94, 59.64),
    ("A3", "789-28", 77, 126.76, 43.16, 56.05),
    ("B1", "789-23", 74, 126.91, 43.21, 58.39),
    ("B2", "789-20", 72, 126.22, 42.98, 59.69),
    ("B3", "789-19", 73, 126.91, 43.21, 59.19),
    ("B4", "789-16", 73, 126.91, 43.21, 59.19),
    ("B5", "789-15", 73, 126.91, 43.21, 59.19),
    ("B6", "789-12", 73, 126.91, 43.21, 59.19),
    ("B7", "789-11", 79, 126.91, 43.48, 55.04),
    ("C1", "789-4",  76, 126.91, 43.21, 56.86),
    ("C2", "789-3",  72, 126.22, 42.98, 59.69),
    ("C3", "789",    82, 126.91, 43.48, 53.02),
    ("D1", "789-34、35", 91, 129.45, 46.23, 50.80),
    ("D2", "789-37、36", 84, 126.76, 43.16, 51.38),
    ("D3", "789-38、39", 85, 126.76, 43.16, 50.78),
    ("D4", "789-41、40", 89, 126.76, 43.16, 48.49),
    ("E1", "789-43、42", 92, 134.98, 47.76, 51.91),
    ("E2", "789-44、45", 89, 134.98, 47.76, 53.66),
    ("E3", "789-47、46", 123, 134.98, 46.24, 37.59),
    ("F1", "789-49、50", 89, 128.92, 45.50, 51.12),
    ("F2", "789-52、51", 83, 126.49, 43.07, 51.89),
    ("F3", "789-53、54", 83, 126.49, 43.07, 51.89),
    ("F4", "789-56、55", 86, 143.50, 51.53, 59.92),
    ("G1", "789-58、57", 88, 136.93, 50.04, 56.86),
    ("G2", "789-59、60", 82, 131.02, 45.32, 55.27),
    ("G3", "789-62、61", 79, 130.62, 46.42, 58.76),
]
M2_PER_PING = 3.305785


def classify_unit(unit, site_m2, floor_m2):
    """依地坪、戶別位置分群."""
    site_ping = site_m2 / M2_PER_PING
    is_corner = unit.endswith(("1", "3")) or unit in ("F4", "D4", "F1", "F5", "G1", "G3")
    # tier rules
    if site_m2 >= 100:
        return "大地坪角間", 22.0, 22.5
    if site_m2 >= 88 and is_corner:
        return "中型角間", 21.5, 22.0
    if floor_m2 >= 135:
        return "標準大戶", 21.0, 21.5
    if site_m2 <= 73:
        return "A1~A3 入口戶" if unit.startswith("A") else "標準小戶", 20.5, 21.0
    return "標準中戶", 21.0, 21.5


def build_dataframe():
    rows = []
    for u, lot, site_m2, floor_m2, build_m2, cov in UNITS:
        tier, p_lo, p_hi = classify_unit(u, site_m2, floor_m2)
        site_ping = round(site_m2 / M2_PER_PING, 1)
        floor_ping = round(floor_m2 / M2_PER_PING, 1)
        # 三情境總價（取建坪×單價，建坪用樓地板/坪）
        price_conservative = round(floor_ping * 19.5)
        price_main = round(floor_ping * (p_lo + p_hi) / 2)
        price_premium = round(floor_ping * p_hi * 1.03)  # +3% 溢價空間
        rows.append({
            "戶": u, "地號": lot, "群組": tier,
            "基地m²": site_m2, "地坪": site_ping,
            "樓地板m²": floor_m2, "建坪": floor_ping,
            "建築面積m²": build_m2, "建蔽率%": cov,
            "建議單價_萬": (p_lo + p_hi) / 2,
            "區間下限": p_lo, "區間上限": p_hi,
            "保守總價_萬": price_conservative,
            "主力總價_萬": price_main,
            "進取總價_萬": price_premium,
        })
    return pd.DataFrame(rows)


def write_excel(df):
    # 三張表：總表、群組彙整、情境總額
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="27戶定價表", index=False)

        # 群組彙整
        grp = df.groupby("群組").agg(
            戶數=("戶", "count"),
            地坪平均=("地坪", "mean"),
            建坪平均=("建坪", "mean"),
            建議單價_萬=("建議單價_萬", "mean"),
            主力總價_萬_中位=("主力總價_萬", "median"),
            主力總價_萬_合計=("主力總價_萬", "sum"),
        ).round(1).reset_index()
        grp.to_excel(w, sheet_name="群組彙整", index=False)

        # 情境總額
        scenarios = pd.DataFrame([
            ["保守情境（單價 -1.5 萬）", df["保守總價_萬"].sum(),
             round(df["保守總價_萬"].mean(), 0), df["保守總價_萬"].min(), df["保守總價_萬"].max()],
            ["★ 主力情境", df["主力總價_萬"].sum(),
             round(df["主力總價_萬"].mean(), 0), df["主力總價_萬"].min(), df["主力總價_萬"].max()],
            ["進取情境（+3% 溢價）", df["進取總價_萬"].sum(),
             round(df["進取總價_萬"].mean(), 0), df["進取總價_萬"].min(), df["進取總價_萬"].max()],
        ], columns=["情境", "全案總銷_萬", "平均總價_萬", "最低總價_萬", "最高總價_萬"])
        scenarios.to_excel(w, sheet_name="情境總額", index=False)


def build_pptx(df):
    prs = new_presentation()

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中道森活｜栖",
        subtitle="27 戶單戶定價試算簡報",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　試算前提與三情境設定",
        "02　戶別分群邏輯（5 群）",
        "03　27 戶定價總表（建議主力情境）",
        "04　群組彙整：各群戶數、平均建坪、建議單價",
        "05　三情境全案總銷對比",
        "06　可調整參數說明（Excel 試算表）",
        "07　下一步：定價公開節奏建議",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ── 1. 試算前提
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　試算前提", "三情境單價設定（取自定價策略簡報 第 06 頁）")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(2.5),
              ["情境", "單價基準", "對標案", "用途"],
              [
                  ["保守（衝去化）", "19.5 萬/坪", "古亭郡、麒美紅葉",
                   "前 3~5 戶內部讓利、急售壓力下緣"],
                  ["★ 主力", "20.5 ~ 22.0 萬/坪", "宗硯大砌、鼎弘玉田 1",
                   "對外公開、媒體與廣告露出版"],
                  ["進取（品牌溢價）", "+3% 上限", "玉田森活 2",
                   "好區位戶（角間/大地坪/邊間）議價空間"],
              ],
              header_size=12, body_size=12, first_col_bold=True, highlight_rows=[1], body_align="left")

    add_rect(s, Inches(0.5), Inches(4.7), Inches(12.3), Inches(2.0),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.85), Inches(12), Inches(0.4),
             "計算公式", size=15, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.3), Inches(11.9), Inches(1.4), [
        "戶別總價 = 樓地板坪 × 該戶建議單價",
        "戶別單價 = 群組基準單價（依地坪、是否角間、樓地板規模分群）",
        "建坪換算：每坪 = 3.305785 m²（依政府實登口徑）",
    ], size=12)

    add_footer(s, "戶別資料源：建照核準圖 DWG（115.02.13 版）")
    add_page_number(s, 2, TOTAL)

    # ── 2. 分群邏輯
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　戶別分群邏輯", "依地坪、角間位置、樓地板規模分為 5 群")
    grp_summary = df.groupby("群組").agg(
        戶數=("戶", "count"),
        代表戶=("戶", lambda s: "、".join(list(s)[:4]) + ("…" if len(s) > 4 else "")),
        地坪範圍=("地坪", lambda s: f"{s.min():.0f}~{s.max():.0f} 坪"),
        建坪平均=("建坪", "mean"),
        建議單價=("建議單價_萬", "mean"),
    ).round(1).reset_index()
    headers = ["群組", "戶數", "代表戶", "地坪範圍", "建坪平均(坪)", "建議單價(萬/坪)"]
    rows = grp_summary.values.tolist()
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(3.5),
              headers, rows, header_size=12, body_size=12,
              first_col_bold=True, body_align="left")

    add_text(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.5),
             "分群規則（依優先順位）", size=14, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(0.5), Inches(6.1), Inches(12.3), Inches(1.0), [
        "基地 ≥ 100 m² → 大地坪角間（溢價力最強）",
        "基地 ≥ 88 m² 且為角間 → 中型角間",
        "樓地板 ≥ 135 m² → 標準大戶",
        "基地 ≤ 73 m² → 標準小戶 / A 幢入口戶（壓主力下緣）",
    ], size=11.5)
    add_footer(s, "分群依 DWG 戶別資料 + 同類型透天市場慣例")
    add_page_number(s, 3, TOTAL)

    # ── 3. 27 戶定價總表（分兩頁顯示）
    for page, (start, end) in enumerate([(0, 14), (14, 27)]):
        s = add_blank_slide(prs); set_background(s, WHITE)
        add_title_bar(s, f"03　27 戶定價總表（{page+1}/2）",
                      "依「★ 建議主力情境」單價試算")
        sub = df.iloc[start:end][["戶", "地號", "群組", "地坪", "建坪", "建議單價_萬", "主力總價_萬"]].copy()
        sub.columns = ["戶", "地號", "群組", "地坪", "建坪", "單價(萬)", "總價(萬)"]
        add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
                  sub.columns.tolist(), sub.values.tolist(),
                  header_size=12, body_size=11.5,
                  first_col_bold=True, body_align="center")
        add_text(s, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
                 f"小計（{end-start} 戶）：總銷 {sub['總價(萬)'].sum():,.0f} 萬",
                 size=12, color=DARK_BROWN, bold=True)
        add_page_number(s, 4 + page, TOTAL)

    # ── 4. 三情境全案總銷
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　三情境全案總銷對比", "27 戶 × 三情境")
    s_lo = df["保守總價_萬"].sum()
    s_mid = df["主力總價_萬"].sum()
    s_hi = df["進取總價_萬"].sum()
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(2.5),
              ["情境", "全案總銷", "27 戶平均", "最低戶", "最高戶"],
              [
                  ["保守", f"{s_lo:,.0f} 萬", f"{s_lo/27:.0f} 萬",
                   f"{df['保守總價_萬'].min():.0f} 萬", f"{df['保守總價_萬'].max():.0f} 萬"],
                  ["★ 主力", f"{s_mid:,.0f} 萬", f"{s_mid/27:.0f} 萬",
                   f"{df['主力總價_萬'].min():.0f} 萬", f"{df['主力總價_萬'].max():.0f} 萬"],
                  ["進取", f"{s_hi:,.0f} 萬", f"{s_hi/27:.0f} 萬",
                   f"{df['進取總價_萬'].min():.0f} 萬", f"{df['進取總價_萬'].max():.0f} 萬"],
              ],
              header_size=13, body_size=12.5, first_col_bold=True,
              highlight_rows=[1])

    # 大數字 — 主力情境
    add_rect(s, Inches(0.5), Inches(4.8), Inches(12.3), Inches(2.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.95), Inches(12), Inches(0.5),
             "建議主力情境總銷", size=14, color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.4), Inches(12), Inches(1.1),
             f"{s_mid:,.0f} 萬　≈　{s_mid/10000:.2f} 億",
             size=42, bold=True, color=ACCENT_RED, font=FONT_TITLE)
    add_text(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.4),
             f"27 戶平均 {s_mid/27:.0f} 萬｜100% 守千萬內｜銷售期 18~24 個月",
             size=12, color=DARK_BROWN)
    add_footer(s, "全案總銷敏感度：單價 ±0.5 萬/坪 → 全案 ±5~6%")
    add_page_number(s, 6, TOTAL)

    # ── 5. Excel 試算表說明
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　可調整參數說明",
                  "完整 Excel 試算表（隨本簡報附）")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["工作表", "內容", "業主可改處"],
              [
                  ["27戶定價表", "每戶完整參數 + 三情境總價",
                   "單價區間下限/上限｜分群名稱"],
                  ["群組彙整", "5 群戶數、平均建坪、建議單價、總額",
                   "群組分類規則"],
                  ["情境總額", "三情境全案總銷、平均、最低最高",
                   "—（自動計算）"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "Excel 為「公式型」試算 — 改單價或分群後總銷自動更新",
             size=13, color=ACCENT_RED, bold=True)
    add_footer(s, "檔名：04_中道森活_27戶定價試算表.xlsx")
    add_page_number(s, 7, TOTAL)

    # ── 6. 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "06　下一步 — 定價公開節奏建議", "三階段定價揭露")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.5),
              ["階段", "時機", "揭露單價", "操作"],
              [
                  ["1. 內部 VIP", "公開前 1~2 個月",
                   "保守情境（19.5 萬/坪）", "前 3~5 戶讓利、客戶名單預埋"],
                  ["2. 公開銷售", "Phase 1 公開",
                   "主力情境（20.5~22 萬/坪）", "依戶別分群定價、媒體露出"],
                  ["3. 後期溢價", "完銷率 ≥ 60% 後",
                   "進取情境（+3%）", "好區位戶提價、品牌溢價"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "三階段揭露 = 鼎弘玉田森活 1 已驗證的去化模式（22 年 15 戶 / 23 年 14 戶完銷）",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 8, TOTAL)

    prs.save(OUT_PPT)
    print(f"wrote {OUT_PPT}")


def main():
    df = build_dataframe()
    write_excel(df)
    print(f"wrote {OUT_XLSX}")
    build_pptx(df)


if __name__ == "__main__":
    main()
