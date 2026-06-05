"""簡報 13：中道森活 — 媒體投放預算試算簡報.

把 deck 6 客群分配（TA1 50% / TA2 30% / TA3 20%）
+ deck 12 預算 280 萬，拆解成具體通路、KPI、時程。

輸出：02_簡報/13_中道森活_媒體投放預算試算.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "13_中道森活_媒體投放預算試算.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()
    TOTAL = 12

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="媒體投放預算試算簡報　Media Budget Plan",
        footer_left="壯圍鄉五權段｜27 戶｜2026 Q4 開賣",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　預算總覽 — 280 萬｜全案 18 個月",
        "02　通路一覽 — 8 個媒體通路",
        "03　通路 1：FB / IG 廣告（TA1+TA2 主力）",
        "04　通路 2：591 案件頁（全 TA）",
        "05　通路 3：YouTube 廣告（TA1 故事版）",
        "06　通路 4：LINE 在地廣告（TA2）",
        "07　通路 5：實體 DM / 夾報 / 戶外（TA2+TA3）",
        "08　通路 6-8：公關、口碑、活動",
        "09　預算 × 通路 × Phase 完整試算表",
        "10　KPI 追蹤指標",
        "11　下一步行動",
    ]
    add_bullets(s, Inches(1.0), Inches(1.7), Inches(11), Inches(5.5),
                items, size=17, line_space=1.45)

    # ── 01 預算總覽
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　預算總覽",
                  "全案 280 萬｜總銷 2.24 億的 1.25%")

    metrics = [
        ("280 萬", "全案行銷預算", DARK_BROWN),
        ("1.25%", "佔總銷比例", ACCENT_RED),
        ("10.4 萬", "每戶 marketing cost", DARK_BROWN),
        ("18 個月", "投放期間", ACCENT_RED),
    ]
    bw = 2.95
    for i, (big, label, color) in enumerate(metrics):
        x = 0.5 + i * (bw + 0.13)
        add_rect(s, Inches(x), Inches(1.85), Inches(bw), Inches(2.0),
                 fill_color=CREAM, line_color=DARK_BROWN)
        add_text(s, Inches(x), Inches(2.0), Inches(bw), Inches(1.0),
                 big, size=40, bold=True, color=color, align="center", font=FONT_TITLE)
        add_text(s, Inches(x), Inches(3.05), Inches(bw), Inches(0.7),
                 label, size=13, color=GRAY_70, align="center", line_space=1.3)

    # TA 分配
    add_text(s, Inches(0.5), Inches(4.3), Inches(12.3), Inches(0.5),
             "預算依 TA 比例分配", size=14, bold=True, color=DARK_BROWN)
    bands = [
        ("TA 1 返鄉換屋  (50%)", "140 萬", BRAND_YELLOW_SOFT, 6.15),
        ("TA 2 在地二代  (30%)",  "84 萬",  CREAM,             3.69),
        ("TA 3 退休置產  (20%)",  "56 萬",  BRAND_YELLOW,      2.46),
    ]
    y_b = 4.85
    for label, amount, fill, w in bands:
        add_rect(s, Inches(0.5), Inches(y_b), Inches(w), Inches(0.55),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(0.65), Inches(y_b + 0.05), Inches(w - 0.3), Inches(0.45),
                 label, size=12, bold=True, color=text_color)
        add_text(s, Inches(0.65), Inches(y_b + 0.05), Inches(w - 0.3), Inches(0.45),
                 amount, size=12, bold=True, color=text_color, align="right")
        y_b += 0.65
    add_footer(s, "預算範圍：透天案合理區間為總銷的 1~2%，本案取下緣偏中")
    add_page_number(s, 2, TOTAL)

    # ── 02 通路一覽
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　通路一覽 — 8 個媒體通路",
                  "依預算佔比排序")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["#", "通路", "預算 (萬)", "佔比", "主要 TA"],
              [
                  ["1", "FB / IG 廣告",           "85", "30%",   "TA 1 + TA 2"],
                  ["2", "短影音製作（3 版本）",   "30", "11%",   "全 TA（製作費）"],
                  ["3", "591 案件頁 + 推薦",      "30", "11%",   "全 TA"],
                  ["4", "看屋會館製作",           "30", "11%",   "全 TA"],
                  ["5", "實體 DM / 夾報 / 戶外",  "25", "9%",    "TA 2 + TA 3"],
                  ["6", "YouTube 廣告",          "20", "7%",    "TA 1"],
                  ["7", "LINE 在地廣告",         "15", "5%",    "TA 2"],
                  ["8", "公關活動 / 口碑 / KOL",  "30", "11%",   "全 TA"],
                  ["—", "業主自留彈性",           "15", "5%",    "—"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="center",
              highlight_rows=[0, 1, 2, 3])

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "合計 280 萬｜前 4 大通路 (FB/IG、短影音、591、看屋會館) 佔 63%",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 3, TOTAL)

    # ── 03 通路 1 FB/IG
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　通路 1：FB / IG 廣告",
                  "85 萬｜TA1+TA2 主力｜18 個月分段投放")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["Phase", "期間", "預算 (萬)", "TA 投放比", "預估觸及 / 來客"],
              [
                  ["Phase 1 暖身", "2026 Q3 (3 月)", "15", "TA1 70% / TA2 30%",
                   "觸及 30 萬｜來客 50"],
                  ["Phase 2 主力", "2026 Q4~2027 Q1 (6 月)", "40", "TA1 60% / TA2 30% / TA3 10%",
                   "觸及 100 萬｜來客 200"],
                  ["Phase 3 精品", "2027 Q2~Q3 (6 月)", "20", "TA1 40% / TA3 60%",
                   "觸及 40 萬｜來客 80"],
                  ["Phase 4 收尾", "2027 Q4 (3 月)", "10", "全 TA",
                   "觸及 20 萬｜來客 40"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="center")

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(0.4),
             "FB/IG 投放策略", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "TA1：FB 雙北精準（30~45 歲、宜蘭出生 tag、有子女）；CTA 連短影音",
        "TA2：FB 宜蘭在地社團、LINE 群組轉發；CTA 連 LINE 加好友",
        "TA3：FB 雙北高消費族群（50+、退休、置產興趣）；CTA 連 591 案件頁",
    ], size=11.5)
    add_page_number(s, 4, TOTAL)

    # ── 04 通路 2 591
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　通路 2：591 案件頁",
                  "30 萬｜全 TA 必經點｜18 個月持續露出")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["項目", "費用 (萬)", "說明"],
              [
                  ["案件頁基本上架 (18 個月)", "10", "完整 591 案件頁 + 戶別清單"],
                  ["首頁推薦 (Phase 2)",      "8",  "宜蘭區首頁輪播、3 個月"],
                  ["關鍵字推薦",              "5",  "「壯圍透天」「宜蘭返鄉」等關鍵字"],
                  ["影音內嵌",                "3",  "60 秒故事版上傳 + 推播"],
                  ["案件 banner 設計",        "4",  "591 平台規格適配"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left",
              highlight_rows=[1])
    add_text(s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.5),
             "591 預估貢獻：每月 30 組來電｜18 個月共 500+ 組來客",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 5, TOTAL)

    # ── 05 通路 3 YouTube
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　通路 3：YouTube 廣告",
                  "20 萬｜TA1 主力｜60 秒故事版投放")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["項目", "費用 (萬)", "說明"],
              [
                  ["TrueView in-stream", "12", "可跳過廣告，按看完計費，主力 TA1"],
                  ["Bumper Ads (6秒)",   "4",  "30秒版剪成6秒，高頻曝光"],
                  ["YouTube Shorts",     "3",  "30秒直式版上傳 + 推播"],
                  ["關鍵字精準",         "1",  "「宜蘭 透天」「壯圍 新案」等"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(0.4),
             "YouTube 投放邏輯", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "TA1 在 YouTube 看：科技、家庭、宜蘭旅遊 — 投這幾個 affinity audience",
        "預估每千次曝光成本 (CPM) 約 80~150 元 — 20 萬可換 130~250 萬次曝光",
    ], size=11.5)
    add_page_number(s, 6, TOTAL)

    # ── 06 通路 4 LINE
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　通路 4：LINE 在地廣告",
                  "15 萬｜TA2 主力｜在地族群滲透")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["項目", "費用 (萬)", "說明"],
              [
                  ["LINE 官方帳號建置 + 維運", "5", "客服自動回覆、預約看屋、推播"],
                  ["LINE 廣告投放 (LAP)",     "6", "宜蘭地區精準投放、按地理位置"],
                  ["LINE 群組投放（在地 KOL）", "3", "宜蘭媽媽群組、壯圍生活圈"],
                  ["導購折扣設計",             "1", "看屋禮、加 LINE 送伴手禮"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.5),
             "LINE 官方帳號預估好友 800~1,200 人｜成交貢獻 5~10 戶",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 7, TOTAL)

    # ── 07 通路 5 實體
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　通路 5：實體 DM / 夾報 / 戶外",
                  "25 萬｜TA2 在地 + TA3 雙北")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.0),
              ["項目", "費用 (萬)", "TA", "說明"],
              [
                  ["A2 海報印刷 + 看屋會館",   "3",  "全 TA", "首批 5 張雪銅 + 亞膜"],
                  ["A4 售屋手冊 (3000 份)",   "5",  "全 TA", "20 頁全彩、含光束圖"],
                  ["宜蘭在地夾報 (3 期)",     "6",  "TA 2",  "聯合報、自由時報宜蘭地方版"],
                  ["雙北精準 DM (5000 戶)",   "8",  "TA 3",  "依社區清單、信箱投遞"],
                  ["戶外看板 (壯圍主要路口)", "3",  "TA 2",  "宜30 縣道、壯六路 3 個月"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_page_number(s, 8, TOTAL)

    # ── 08 通路 6-8 公關活動
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　通路 6-8：公關、口碑、KOL 活動",
                  "30 萬｜建立品牌信任感")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["項目", "費用 (萬)", "目的"],
              [
                  ["開賣記者會（Phase 2 啟動）",        "8",  "在地媒體報導、行業曝光"],
                  ["玉田森活屋主聯誼會 + 中道導覽",     "5",  "口碑傳播、轉介紹"],
                  ["宜蘭在地媒體合作 (3 篇)",          "6",  "宜蘭新聞、地方雜誌"],
                  ["微網紅合作 (5 位)",                "8",  "在地媽媽、宜蘭旅遊 KOC"],
                  ["看屋禮 / 抽獎活動",                "3",  "提升現場成交率"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "公關價值：每次曝光 ROI 高、品牌信任度提升 — 但難量化、需長期",
             size=12, color=GRAY_70, align="center")
    add_page_number(s, 9, TOTAL)

    # ── 09 完整預算試算表
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　預算 × 通路 × Phase 完整試算表",
                  "全案 280 萬｜18 個月分配")
    add_table(s, Inches(0.3), Inches(1.7), Inches(12.7), Inches(4.8),
              ["通路", "Phase 1", "Phase 2", "Phase 3", "Phase 4", "合計"],
              [
                  ["FB / IG 廣告",       "15", "40", "20", "10", "85"],
                  ["短影音製作",         "30", "—",  "—",  "—",  "30"],
                  ["591 案件頁",         "5",  "15", "7",  "3",  "30"],
                  ["看屋會館",           "30", "—",  "—",  "—",  "30"],
                  ["實體 DM / 戶外",     "5",  "10", "7",  "3",  "25"],
                  ["YouTube",            "—",  "12", "8",  "—",  "20"],
                  ["LINE 在地",          "3",  "8",  "3",  "1",  "15"],
                  ["公關 / KOL",         "5",  "15", "8",  "2",  "30"],
                  ["業主彈性",           "—",  "5",  "5",  "5",  "15"],
                  ["★ 小計 (萬)",        "93", "105", "58", "24", "280"],
              ],
              header_size=11, body_size=10.5, first_col_bold=True, body_align="center",
              highlight_rows=[9])

    add_text(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.4),
             "Phase 1 重看屋會館初建｜Phase 2 重廣告衝刺｜Phase 3 重精品｜Phase 4 收尾",
             size=11, color=DARK_BROWN, bold=True, align="center")
    add_page_number(s, 10, TOTAL)

    # ── 10 KPI 追蹤指標
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "10　KPI 追蹤指標",
                  "每月追蹤｜每季檢討｜可調整")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["指標", "目標", "計算方式", "預警閾值"],
              [
                  ["FB/IG 來客數",     "每月 ≥ 20 組",   "FB ad → 表單填寫",     "< 12 組"],
                  ["591 來電數",       "每月 ≥ 30 組",   "591 後台統計",         "< 18 組"],
                  ["LINE 好友數",      "每月 ≥ 50 加",   "官方帳號後台",         "< 30 加"],
                  ["看屋會館來客",     "每月 ≥ 40 組",   "現場簽到",             "< 25 組"],
                  ["看屋 → 成交轉換",  "≥ 5%",          "成交數 / 來客數",      "< 3%"],
                  ["全案完銷時間",     "≤ 24 個月",     "從 Phase 1 啟動算",    "> 30 個月"],
                  ["每戶平均成交價",   "≥ 830 萬",      "成交總額 / 戶數",      "< 800 萬"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "預警觸發時：璞域提案調整方案（增預算、改 message、降單價、延長期）",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 11, TOTAL)

    # ── 11 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "11　下一步行動", "預算定案 → 啟動投放")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "工項", "璞域提供", "業主決定"],
              [
                  ["1", "預算總額確認",       "推薦 280 萬", "_______"],
                  ["2", "通路優先順序",       "前 4 大優先", "_______"],
                  ["3", "代理商選擇",         "璞域代為操作 vs 其他代理", "_______"],
                  ["4", "Phase 1 啟動",       "Q3 2026 中", "_______"],
                  ["5", "KPI 報告頻率",       "每月一份", "_______"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "決定後 1 週內提供：媒體採購清單、執行時程表、合約草案",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 12, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
