"""簡報 14：中道森活 — 看屋會館實體規劃簡報.

整合 deck 6（三條 TA 動線）+ deck 9（海報安裝位置），
做成看屋會館 layout 平面圖 + 區域功能說明 + 家具設備清單 + 預算。

輸出：02_簡報/14_中道森活_看屋會館實體規劃.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "14_中道森活_看屋會館實體規劃.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def draw_floorplan(slide, x, y, w, h):
    """在 slide 上畫一個簡化的接待中心平面圖."""
    # 整體外框
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h),
             fill_color=WHITE, line_color=DARK_BROWN)

    # 區域劃分（座標皆相對於 x,y，以英吋計）
    # 平面採用 4×3 grid 概念
    # 區域定義: (rx, ry, rw, rh, name, fill, label_color)
    zones = [
        # 入口（下方中央）
        (3.0, h - 1.0, 2.0, 1.0, "入口接待", BRAND_YELLOW_SOFT, DARK_BROWN),
        # 主展示中心（中央上）
        (1.0, 0.8, 5.0, 2.5, "展示中心\n主視覺 + 模型 + 機能海報", CREAM, DARK_BROWN),
        # 實品屋體驗 (右上)
        (6.2, 0.8, 2.6, 3.5, "實品屋\n（4 房格局示範）", BRAND_YELLOW, WHITE),
        # 洽談區 (左中)
        (0.3, 3.5, 3.5, 1.8, "洽談區\n（3 桌）", CREAM, DARK_BROWN),
        # 兒童區 (左下)
        (0.3, 5.5, 1.8, 1.5, "兒童區\nTA1 友善", BRAND_YELLOW_SOFT, DARK_BROWN),
        # 茶水 (中下)
        (2.3, 5.5, 1.5, 1.5, "茶水區", CREAM, DARK_BROWN),
        # 貴賓室 (右中下)
        (4.0, 3.5, 2.0, 2.0, "貴賓室\nTA3 接待", BRAND_YELLOW, WHITE),
        # 廁所 (右下)
        (6.2, 4.5, 1.2, 1.2, "WC", CREAM, DARK_BROWN),
        # 後場 (右下角)
        (7.5, 4.5, 1.3, 2.2, "後場", WHITE, GRAY_70),
    ]

    for (rx, ry, rw, rh, name, fill, label_color) in zones:
        add_rect(slide, Inches(x + rx), Inches(y + ry),
                 Inches(rw), Inches(rh), fill_color=fill, line_color=DARK_BROWN)
        add_text(slide, Inches(x + rx), Inches(y + ry + rh / 2 - 0.3),
                 Inches(rw), Inches(0.6),
                 name, size=11, bold=True, color=label_color,
                 align="center", anchor="middle", line_space=1.25)

    # 三條 TA 動線箭頭（用粗線代表）
    # TA1 路線：入口 → 展示 → 實品屋 → 兒童區 → 洽談
    # TA2 路線：入口 → 展示 → 洽談 → 出口
    # TA3 路線：入口 → 貴賓室 → 實品屋 → 洽談

    # 海報位置標記（紅星）
    poster_spots = [
        (x + 3.0, y + 0.4, "P1\n入口正面"),
        (x + 1.8, y + 3.3, "P2\n接待桌後"),
        (x + 5.3, y + 4.0, "P3\n貴賓室"),
    ]
    for sx, sy, lbl in poster_spots:
        # 星狀（用三角）
        add_rect(slide, Inches(sx - 0.12), Inches(sy - 0.12),
                 Inches(0.24), Inches(0.24), fill_color=ACCENT_RED, line_color=WHITE)
        add_text(slide, Inches(sx - 0.7), Inches(sy + 0.18),
                 Inches(1.4), Inches(0.5),
                 lbl, size=8, bold=True, color=ACCENT_RED, align="center", line_space=1.2)


def build():
    prs = new_presentation()
    TOTAL = 11

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="看屋會館實體規劃簡報　Sales Centre Layout",
        footer_left="壯圍鄉五權段｜27 戶｜2026 Q3 啟用",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　設計核心理念 — 三條 TA 動線、不互相干擾",
        "02　平面圖總覽（80~120 坪建議）",
        "03　九大區域功能說明",
        "04　TA 1 動線（返鄉換屋族）",
        "05　TA 2 動線（在地二代）",
        "06　TA 3 動線（退休置產）",
        "07　主視覺與海報安裝位置",
        "08　家具與設備清單",
        "09　預算試算（建置 + 維運）",
        "10　下一步行動",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=18, line_space=1.45)

    # ── 01 設計核心理念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　設計核心理念",
                  "三條 TA 動線並存、不互相打擾")
    add_rect(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(2.5),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(1.85), Inches(12), Inches(0.5),
             "三大設計原則", size=15, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(2.35), Inches(11.9), Inches(1.8), [
        "①  動線分流：三組 TA 各有偏好路徑，接待員依客層判斷帶往不同區域",
        "②  情境共存：兒童區（TA1）與貴賓室（TA3）必須空間隔離，避免互相影響",
        "③  視覺一致：主視覺（v7 金黃·玫瑰）滲透每個區域，但密度依場景調整",
    ], size=12)

    add_table(s, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5),
              ["TA", "動線特色", "核心區域"],
              [
                  ["TA 1 返鄉換屋", "週末攜家帶眷、希望快速看完", "展示中心 → 實品屋 → 兒童區 → 洽談"],
                  ["TA 2 在地二代", "下班後快速看、需停車", "展示中心 → 洽談 → 出口"],
                  ["TA 3 退休置產", "預約專人接待、需慢慢看", "貴賓室 → 實品屋 → 庭院區 → 洽談"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_page_number(s, 2, TOTAL)

    # ── 02 平面圖總覽
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　平面圖總覽",
                  "建議坪數 80~120 坪｜接待中心+實品屋一體")
    draw_floorplan(s, x=0.5, y=1.7, w=9.0, h=5.0)

    # 右側說明
    add_text(s, Inches(10.0), Inches(1.7), Inches(3.0), Inches(0.4),
             "圖例", size=14, bold=True, color=DARK_BROWN)
    legend = [
        ("展示／公共區", CREAM),
        ("主動線（TA1+2）", BRAND_YELLOW_SOFT),
        ("貴賓／實品", BRAND_YELLOW),
        ("海報位置 P1~P3", ACCENT_RED),
    ]
    for i, (lbl, color) in enumerate(legend):
        yy = 2.15 + i * 0.45
        add_rect(s, Inches(10.0), Inches(yy + 0.05),
                 Inches(0.25), Inches(0.25), fill_color=color, line_color=DARK_BROWN)
        add_text(s, Inches(10.35), Inches(yy), Inches(2.7), Inches(0.35),
                 lbl, size=11, color=DARK_BROWN, anchor="middle")

    add_rect(s, Inches(10.0), Inches(4.5), Inches(3.0), Inches(2.2),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(10.15), Inches(4.6), Inches(2.8), Inches(0.4),
             "尺寸建議", size=12, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(10.15), Inches(5.0), Inches(2.8), Inches(1.6), [
        "建議 100 坪",
        "層高 ≥ 3.5m",
        "正面寬 ≥ 12m",
        "停車 ≥ 8 位",
    ], size=10, line_space=1.4)
    add_footer(s, "本平面為示意 — 實際依基地條件、預算、租金可彈性調整")
    add_page_number(s, 3, TOTAL)

    # ── 03 九大區域功能
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　九大區域功能",
                  "從入口到後場 — 每區角色明確")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["#", "區域", "面積", "主要功能", "客戶第一印象"],
              [
                  ["1", "入口接待",   "5 坪",   "前台、簽到、轉接",       "親切、明亮、看到主視覺"],
                  ["2", "展示中心",   "20 坪",  "全案模型、主視覺、機能海報", "震撼、資訊密集"],
                  ["3", "實品屋",     "15 坪",  "4 房格局示範（1 間）",   "想像「住在這裡」"],
                  ["4", "洽談區",     "12 坪",  "3 桌、議價、簽約",       "正式、開放"],
                  ["5", "兒童區",     "4 坪",   "TA1 子女遊戲",           "家庭友善"],
                  ["6", "茶水區",     "3 坪",   "飲品、點心",             "在地茶、宜蘭點心"],
                  ["7", "貴賓室",     "8 坪",   "TA3 深度討論",           "典雅、私密"],
                  ["8", "WC",         "2 坪",   "—",                     "乾淨、設備新"],
                  ["9", "後場/備品",  "5 坪",   "員工、儲藏、設備",       "—（不對客戶開放）"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
             "九區合計約 74 坪（含通道約 100 坪）",
             size=11, color=GRAY_70, align="center")
    add_page_number(s, 4, TOTAL)

    # ── 04-06 三條 TA 動線
    ta_routes = [
        ("04　TA 1 返鄉換屋族動線",
         "週末攜家帶眷、目標明確、希望 30 分鐘看完",
         [
             ("Step 1", "入口", "親子歡迎、孩子畫具"),
             ("Step 2", "展示中心", "全案模型 + 4 房模型屋"),
             ("Step 3", "實品屋", "4 房格局＋兒童房示範"),
             ("Step 4", "兒童區", "孩子玩耍、大人放鬆"),
             ("Step 5", "洽談區", "學區、通勤地圖、價格"),
             ("Step 6", "出口", "親子合照 + 紀念品"),
         ]),
        ("05　TA 2 在地二代動線",
         "下班後快速看、目標性強、預算敏感",
         [
             ("Step 1", "入口", "鄰里口碑牆、在地照片"),
             ("Step 2", "展示中心", "千萬內定價對照表"),
             ("Step 3", "實品屋", "標準戶＋停車示範"),
             ("Step 4", "洽談區", "工作機能、通勤時間"),
             ("Step 5", "出口", "鼎弘玉田森活 1 完銷說明 + LINE 加好友"),
         ]),
        ("06　TA 3 退休置產動線",
         "預約專人接待、慢、看完整套",
         [
             ("Step 1", "入口", "貴賓接待員導引"),
             ("Step 2", "貴賓室", "材料樣板、庭院設計圖"),
             ("Step 3", "實品屋", "大地坪戶＋庭院實景"),
             ("Step 4", "貴賓室", "退休生活提案、家庭聚會場景"),
             ("Step 5", "茶水區", "宜蘭在地茶、深度交流"),
             ("Step 6", "出口", "保密議價、來日再訪"),
         ]),
    ]
    for idx, (title, sub, steps) in enumerate(ta_routes):
        s = add_blank_slide(prs); set_background(s, WHITE)
        add_title_bar(s, title, sub)

        # 步驟卡片
        bw = 1.95; bh = 4.4; y_b = 1.85
        # 顏色依 TA
        fills = [BRAND_YELLOW_SOFT, CREAM, BRAND_YELLOW][idx]
        for i, (step, zone, desc) in enumerate(steps):
            x = 0.4 + i * (bw + 0.07)
            add_rect(s, Inches(x), Inches(y_b), Inches(bw), Inches(bh),
                     fill_color=fills, line_color=DARK_BROWN)
            text_color = WHITE if fills == BRAND_YELLOW else DARK_BROWN
            add_text(s, Inches(x), Inches(y_b + 0.15), Inches(bw), Inches(0.4),
                     step, size=12, bold=True, color=ACCENT_RED, align="center")
            add_text(s, Inches(x), Inches(y_b + 0.6), Inches(bw), Inches(0.6),
                     zone, size=15, bold=True, color=text_color, align="center", font=FONT_TITLE)
            add_text(s, Inches(x + 0.15), Inches(y_b + 1.4), Inches(bw - 0.3), Inches(2.8),
                     desc, size=11, color=text_color, align="center", line_space=1.45)

        add_footer(s, "三條動線各自獨立、互不干擾 — 接待員依客層判斷分流")
        add_page_number(s, 5 + idx, TOTAL)

    # ── 07 海報安裝位置
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　主視覺與海報安裝位置",
                  "三張 A2 機能地圖海報｜對應三條動線")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["位置", "海報內容", "視距", "對應 TA"],
              [
                  ["P1 入口正面牆", "中道森活 主視覺 + 案件 snapshot", "3~5 m", "全 TA 第一印象"],
                  ["P2 接待桌後方", "A2 機能光束圖", "1.5~2 m", "TA1 / TA2 等待時瀏覽"],
                  ["P3 貴賓室牆面", "A2 機能光束圖 + 林先生故事", "1.5~2 m", "TA3 深度討論輔助"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left",
              highlight_rows=[1])

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(0.4),
             "裝裱與燈光", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "裝裱：深棕色木框 + 玻璃（呼應金黃主視覺）",
        "燈光：3000K 暖光｜照度 300~500 lux｜避免反光",
        "視線高度：海報中心離地 1.4~1.6 m（成人平視）",
    ], size=12)
    add_page_number(s, 8, TOTAL)

    # ── 08 家具與設備清單
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　家具與設備清單",
                  "建置一次性投資｜全清單已含預算")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["區域", "家具 / 設備", "預估費用 (萬)"],
              [
                  ["入口接待",  "前台、簽到 iPad、立式燈、迎賓椅", "5"],
                  ["展示中心",  "全案模型 (1:50)、4 房模型屋、機能海報×2", "60"],
                  ["實品屋",    "4 房格局裝潢（含廚具、衛浴、燈光）", "100"],
                  ["洽談區",    "3 套桌椅、簡報投影、文具", "15"],
                  ["兒童區",    "地墊、繪本架、小椅子、玩具", "5"],
                  ["茶水區",    "吧台、咖啡機、茶具、冰箱", "8"],
                  ["貴賓室",    "沙發、咖啡桌、樣板架、藝術品", "20"],
                  ["共用",      "空調、音響、燈光、監視器", "20"],
                  ["合計",      "—", "233"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left",
              highlight_rows=[8])
    add_text(s, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
             "另含主視覺輸出、字標立體標示、戶外招牌約 30 萬｜總計約 263 萬",
             size=11, color=GRAY_70, align="center")
    add_page_number(s, 9, TOTAL)

    # ── 09 預算試算
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　預算試算 — 建置 + 維運",
                  "全案 18 個月｜總投入 350~450 萬")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["項目", "費用 (萬)", "說明"],
              [
                  ["A. 建置 (一次)",       "—",       "—"],
                  ["  家具 / 設備",        "233",     "詳上頁清單"],
                  ["  主視覺輸出",         "30",      "海報、招牌、字標立體"],
                  ["  室內裝潢工程",       "60",      "地坪、天花板、隔間（如需）"],
                  ["  小計",               "323",     "一次性投資"],
                  ["B. 維運 (18 個月)",    "—",       "—"],
                  ["  房租或場地費",       "90",      "5 萬/月 × 18 月（依地點）"],
                  ["  水電網",             "9",       "0.5 萬/月 × 18 月"],
                  ["  接待人員",           "—",       "業主自聘，不含在此"],
                  ["  耗材 / 飲品 / 印刷補充", "9",   "0.5 萬/月 × 18 月"],
                  ["  小計",               "108",     "持續支出"],
                  ["合計",                 "431",     "—"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left",
              highlight_rows=[4, 10, 11])
    add_page_number(s, 10, TOTAL)

    # ── 10 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "10　下一步行動", "從本簡報到看屋會館啟用")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["週次", "工項", "璞域", "業主"],
              [
                  ["W1", "場地確認 + 平面定稿", "陪同勘地、修平面", "選定地點"],
                  ["W2~3", "室內設計", "提案 + 修圖", "預算上限"],
                  ["W4~6", "裝潢工程", "監工", "—"],
                  ["W6~8", "家具 / 設備進場", "代購", "—"],
                  ["W7~8", "主視覺安裝", "印刷、裝裱、安裝", "—"],
                  ["W8", "員工訓練", "三條動線教學", "聘員工"],
                  ["W9", "soft launch", "VIP 預約看屋", "—"],
                  ["W10", "正式開放", "Phase 1 暖身期啟動", "—"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "全程約 10 週｜銜接 2026 Q3 Phase 1 暖身期",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 11, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
