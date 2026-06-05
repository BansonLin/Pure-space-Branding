"""簡報 12：中道森活 — 業主會議 15 分鐘執行摘要.

從 deck 1~11 濃縮為 9 頁，給業主開會用一頁一張結論。

輸出：02_簡報/12_中道森活_業主會議15分鐘執行摘要.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "12_中道森活_業主會議15分鐘執行摘要.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()
    TOTAL = 9

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 業主會議專用",
        title_main="中 道 森 活",
        subtitle="執行摘要｜15 分鐘決策版　Executive Summary",
        footer_left="壯圍鄉五權段｜27 戶｜2026 Q4 預計開賣",
        footer_right="2026.06")

    # ── 01 案件 snapshot
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　案件 Snapshot",
                  "一頁認識中道森活")

    # 4 big KPIs
    metrics = [
        ("27 戶", "獨棟透天", DARK_BROWN),
        ("38~43 坪", "建坪規模", ACCENT_RED),
        ("21 萬/坪", "建議單價", DARK_BROWN),
        ("2.24 億", "預估總銷", ACCENT_RED),
    ]
    bw = 2.95
    for i, (big, label, color) in enumerate(metrics):
        x = 0.5 + i * (bw + 0.13)
        add_rect(s, Inches(x), Inches(1.75), Inches(bw), Inches(1.7),
                 fill_color=CREAM, line_color=DARK_BROWN)
        add_text(s, Inches(x), Inches(1.85), Inches(bw), Inches(0.9),
                 big, size=36, bold=True, color=color, align="center", font=FONT_TITLE)
        add_text(s, Inches(x), Inches(2.85), Inches(bw), Inches(0.5),
                 label, size=13, color=GRAY_70, align="center")

    # 三段資訊欄
    info = [
        ("基地", "宜蘭縣壯圍鄉五權段 789、790 等 45 筆地號"),
        ("起造", "鼎弘健康生技｜代表人 陳鼎琳"),
        ("營造", "宗硯建設（同營造廠關聯案：宗硯大砌系列18 31 戶）"),
        ("規劃", "A~G 共 7 幢｜地上 3 層 + 屋頂層｜RC 結構"),
        ("目標", "100% 守千萬內｜預估 18~24 個月完銷"),
        ("信譽", "鼎弘玉田森活 1/2、學進賦 3 案已完銷"),
    ]
    y_base = 3.7
    for i, (k, v) in enumerate(info):
        row = i // 2; col = i % 2
        x = 0.5 + col * 6.3
        y = y_base + row * 0.7
        add_text(s, Inches(x), Inches(y), Inches(1.2), Inches(0.4),
                 k, size=12, color=GRAY_70)
        add_text(s, Inches(x + 1.3), Inches(y), Inches(5), Inches(0.4),
                 v, size=12, bold=True, color=DARK_BROWN)

    add_footer(s, "詳細資料：deck 01 客群定價、deck 06 客群深掘")
    add_page_number(s, 1, TOTAL)

    # ── 02 客群與定價
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　客群與定價",
                  "三組 TA × 三條價格帶｜27 戶分配")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.0),
              ["客群", "戶數", "每坪", "總價帶", "對應戶別"],
              [
                  ["TA 1 宜蘭返鄉換屋族 (45%)", "12 戶", "21.0 萬", "810~903 萬",
                   "B7、C3、D2~D4、E2、F1~F3、G1~G3"],
                  ["TA 2 壯圍/礁溪在地二代 (35%)", "10 戶", "20.5 萬", "780~840 萬",
                   "A2、A3、B1~B6、C1、C2"],
                  ["TA 3 雙北退休置產 (20%)", "5 戶", "22.0 萬", "920~990 萬",
                   "A1、D1、E1、E3、F4"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left",
              highlight_rows=[0])

    # 全案總銷
    add_rect(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(2.0),
             fill_color=BRAND_YELLOW, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.15), Inches(12), Inches(0.5),
             "★ 全案總銷（建議主力情境）", size=14, bold=True, color=WHITE)
    add_text(s, Inches(0.7), Inches(5.65), Inches(7.5), Inches(1.2),
             "22,431 萬　≈　2.24 億",
             size=40, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(s, Inches(9.0), Inches(5.7), Inches(3.6), Inches(0.5),
             "27 戶平均 831 萬", size=14, color=WHITE, bold=True)
    add_text(s, Inches(9.0), Inches(6.2), Inches(3.6), Inches(0.5),
             "100% 守千萬內", size=14, color=WHITE, bold=True)
    add_page_number(s, 2, TOTAL)

    # ── 03 競品與關係
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　競品地圖（重新分類）",
                  "宗硯大砌 = 關聯案不是對手｜真正競品是另外三案")

    # 三欄
    cols = [
        (0.5, "關聯案 (同營造)", BRAND_YELLOW_SOFT,
         ["宗硯大砌系列18-透天",
          "31 戶｜21.2 萬｜858 萬",
          "戰術：對齊定價、錯開時程",
          "資訊管道：宗硯輔銷 = 璞域客戶"]),
        (4.7, "真正競品 (千萬內)", CREAM,
         ["古亭郡（870 萬｜19.1 萬）",
          "麒美紅葉（995 萬｜19.6 萬）",
          "艾利永美（1,078 萬｜22.7 萬）",
          "赫宇聯勤2（1,018 萬）"]),
        (8.9, "鼎弘自有戰績", CREAM,
         ["玉田森活1（30 戶｜930 萬 完銷）",
          "玉田森活2（14 戶｜1,513 萬 完銷）",
          "學進賦（10 戶｜587 萬 完銷）",
          "三案皆完銷 = 鼎弘品牌"]),
    ]
    for x, title, fill, items in cols:
        add_rect(s, Inches(x), Inches(1.8), Inches(4.0), Inches(4.5),
                 fill_color=fill, line_color=DARK_BROWN)
        add_text(s, Inches(x), Inches(1.95), Inches(4.0), Inches(0.5),
                 title, size=16, bold=True, color=DARK_BROWN, align="center")
        add_bullets(s, Inches(x + 0.2), Inches(2.6), Inches(3.7), Inches(3.6),
                    items, size=11.5, line_space=1.5)

    add_rect(s, Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.6),
             fill_color=DARK_BROWN, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(6.5), Inches(12), Inches(0.6),
             "戰術：對齊宗硯定價 21 萬｜時程錯開 1~2 季｜真正要競爭的是古亭郡和麒美紅葉",
             size=12.5, bold=True, color=WHITE, anchor="middle")
    add_page_number(s, 3, TOTAL)

    # ── 04 視覺與命名
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　視覺、命名、字標",
                  "業主已選方向 — 璞域推薦下一步")

    items = [
        ("命名", "中道森活｜栖", "業主初步方向「中道森活」延續玉田森活；副標「栖」鎖定三組 TA 共通價值「想要一個落腳的家」"),
        ("主視覺", "v7 金黃·古典玫瑰", "業主先前選定。全客層覆蓋最廣，印刷與數位皆強"),
        ("字標系統", "中道森活｜栖 字標", "宋體本位、東方氣質、可橫式/直式/堆疊三變體"),
        ("色彩", "金黃 + 深棕 + 米白", "主色：#E8AC1B｜深棕：#4A3210｜米白：#FFF8E7"),
        ("延伸副系列", "v5 湖青｜v2 普藍", "高端精品場景使用（看屋會館貴賓室、年節限定 DM）"),
    ]
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["類別", "推薦", "說明"],
              [list(it) for it in items],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "字標、海報、文案、影音腳本皆已完成 v0.1 草案，業主決定方向後可立即量產",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 4, TOTAL)

    # ── 05 銷售節奏
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　銷售節奏",
                  "4 階段｜預估 18~24 個月完銷")

    phases = [
        ("Phase 1\n暖身期", "Q3 2026",   "5 戶｜TA 2 為主\n壓低單價、建立能見度", CREAM),
        ("Phase 2\n主力期", "Q4 2026~Q1 2027", "12 戶｜TA 1 為主\n主力廣告、媒體曝光", BRAND_YELLOW_SOFT),
        ("Phase 3\n精品期", "Q2~Q3 2027", "5 戶｜TA 3 為主\n品牌溢價、單客深度", BRAND_YELLOW),
        ("Phase 4\n收尾期", "Q4 2027",   "5 戶｜混合\n口碑、彈性議價", CREAM),
    ]
    bw = 2.9; bh = 3.8; y0 = 1.85
    for i, (label, when, action, fill) in enumerate(phases):
        x = 0.5 + i * (bw + 0.15)
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(x), Inches(y0 + 0.15), Inches(bw), Inches(0.85),
                 label, size=16, bold=True, color=text_color, align="center",
                 line_space=1.2, font=FONT_TITLE)
        add_text(s, Inches(x), Inches(y0 + 1.05), Inches(bw), Inches(0.5),
                 when, size=12, color=ACCENT_RED, align="center", bold=True)
        add_text(s, Inches(x + 0.2), Inches(y0 + 1.7), Inches(bw - 0.4), Inches(1.9),
                 action, size=11, color=text_color, align="center", line_space=1.4)

    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "對標：鼎弘玉田森活 1（30 戶 / 2 年完銷）— 模型可直接複製",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.4),
             "與宗硯大砌（2026 Q2 衝刺）錯開 1~2 季，不打架",
             size=11.5, color=GRAY_70, align="center")
    add_page_number(s, 5, TOTAL)

    # ── 06 故事與賣點
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　核心故事",
                  "TA1 林先生的一天 — 5 行說完位置價值")

    add_rect(s, Inches(0.5), Inches(1.85), Inches(12.3), Inches(2.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(2.0), Inches(12), Inches(2.2),
             "「在台北 8 年，我們每月房租 30,000、通勤 90 分鐘。\n"
             "搬回壯圍後：房貸 30,000、通勤一週兩次。\n"
             "省下的時間，是和爸媽、孩子的時間。」\n"
             "—— 中道森活｜千萬內｜3 層透天｜27 戶｜2026 預售",
             size=15, color=DARK_BROWN, line_space=1.7, anchor="middle")

    # 三個賣點
    sellers = [
        ("學區",   "公館 / 壯圍國小 4~5 分鐘車",     ACCENT_RED),
        ("通勤",   "國道5 8 分｜雪隧到台北 50 分",   DARK_BROWN),
        ("生活",   "壯圍沙丘 7 分、夢時代 13 分、醫院 14 分", ACCENT_RED),
    ]
    for i, (k, v, color) in enumerate(sellers):
        x = 0.5 + i * 4.2
        add_rect(s, Inches(x), Inches(4.7), Inches(4.0), Inches(1.7),
                 fill_color=CREAM, line_color=color)
        add_text(s, Inches(x), Inches(4.85), Inches(4.0), Inches(0.5),
                 k, size=18, bold=True, color=color, align="center")
        add_text(s, Inches(x + 0.2), Inches(5.4), Inches(3.6), Inches(0.95),
                 v, size=12, color=DARK_BROWN, align="center", line_space=1.45)

    add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.5),
             "完整機能光束圖（26 機能點、5 圈距離）詳 deck 08；故事擴寫成短影音腳本詳 deck 10",
             size=11, color=GRAY_70, align="center")
    add_page_number(s, 6, TOTAL)

    # ── 07 行銷預算與資源
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　行銷預算與資源需求",
                  "全案 280 萬｜總銷的 1.25%")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["項目", "預算 (萬)", "佔比", "說明"],
              [
                  ["數位廣告 (FB/IG/Google)", "120", "43%", "主力 TA1+TA2 投放"],
                  ["短影音製作 (3 版本)",     "30",  "11%", "30s/60s/90s 一次拍攝"],
                  ["591 + 房產垂直",          "30",  "11%", "12 個月案件頁 + 推薦"],
                  ["實體 DM / 戶外 / 夾報",   "40",  "14%", "在地 TA2 + 雙北 TA3"],
                  ["看屋會館製作",            "30",  "11%", "海報、模型、實品屋"],
                  ["公關活動 / 口碑",         "20",  "7%",  "開賣記者會、屋主聯誼"],
                  ["業主自留彈性",            "10",  "3%",  "突發應變"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left",
              highlight_rows=[0, 4])

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(0.4),
             "預算 ROI 預估", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "全案總銷 2.24 億｜行銷預算 280 萬 = 1.25%（透天案合理區間 1~2%）",
        "每戶 marketing cost 約 10.4 萬｜每戶毛利約 200 萬+，ROI 約 19 倍",
    ], size=12)
    add_page_number(s, 7, TOTAL)

    # ── 08 下一步五項待決策
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "08　下一步 — 五項待業主決策",
                  "今天會議後可立即推動")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["#", "決策項", "璞域建議", "業主決定"],
              [
                  ["1", "案名定案",       "中道森活｜栖", "_______"],
                  ["2", "主視覺定案",     "v7 金黃·玫瑰（同業主先前方向）", "_______"],
                  ["3", "建議主力單價",   "21.0 萬/坪", "_______"],
                  ["4", "行銷預算上限",   "280 萬", "_______"],
                  ["5", "Phase 1 啟動時點","2026 Q3", "_______"],
              ],
              header_size=12, body_size=13, first_col_bold=True, body_align="left")

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "決定後璞域 2 週內提供：宗硯輔銷協同會議、看屋會館設計圖、廣告投放計畫",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 8, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
