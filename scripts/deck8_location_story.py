"""簡報 8：中道森活 — 機能、光束圖與故事簡報.

對應業主需求：給開會用的機能圖 + 光束圖 + TA1 故事。

機能盤點以 DWG 取得之地址（壯圍鄉五權段 789、790 等地號，
鄰接中央路三段、南北五路、忠孝一路、順和路）為中心，
推估周邊：
  - 學區（公立國小、國中）
  - 賣場（傳統市場、超市、量販）
  - 商店（便利商店、餐飲）
  - 觀光景點
  - 對外交通（國道、火車）

⚠️ 距離/時間為合理估算，業主請以實際座標再校驗。

輸出：02_簡報/08_中道森活_機能光束圖與故事.pptx
"""
from pathlib import Path
import math
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "08_中道森活_機能光束圖與故事.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)
TOTAL = 17

# 機能地圖資料 ─ 以基地為中心
# (name, category, direction_deg, distance_km, drive_min, note)
# direction: 0=正北, 90=東, 180=南, 270=西
AMENITIES = [
    # ── 學區
    ("壯圍國小",         "school",  300, 2.0,  5,  "公立國小，學區內"),
    ("公館國小",         "school",  330, 1.5,  4,  "公立國小，學區邊界"),
    ("過嶺國小",         "school",  120, 2.5,  6,  "公立國小，另一學區"),
    ("壯圍國中",         "school",  180, 5.0,  10, "公立國中，全鄉合一"),
    ("中華國中",         "school",  240, 6.0,  12, "宜蘭市區國中（可越區）"),

    # ── 賣場/超市/便利商店
    ("全聯 壯圍店",      "shop",    250, 2.5,  5,  "日常採購"),
    ("全聯 宜蘭東港店",  "shop",    240, 5.0,  10, "大型超市"),
    ("7-11 壯圍門市",    "shop",    280, 1.0,  3,  "24 小時便利商店"),
    ("全家 五權店",      "shop",    100, 1.0,  3,  "24 小時便利商店"),
    ("壯圍傳統市場",     "shop",    280, 2.0,  5,  "在地早市、新鮮蔬果"),
    ("家樂福 宜蘭店",    "shop",    255, 8.0,  15, "大型量販，週末採購"),
    ("宜蘭夢時代",       "shop",    245, 6.5,  13, "百貨購物中心"),

    # ── 醫療
    ("壯圍鄉衛生所",     "medical", 290, 2.5,  6,  "基本醫療"),
    ("羅東博愛醫院",     "medical", 200, 12.0, 20, "區域醫院"),
    ("國立陽明交大附醫", "medical", 240, 7.0,  14, "宜蘭最大公立醫院"),

    # ── 觀光景點（假日生活）
    ("壯圍沙丘 旅服中心", "tourism", 90,  3.0,  7,  "黃聲遠設計，地標景點"),
    ("永鎮海濱公園",     "tourism", 95,  4.5,  10, "親子海邊"),
    ("宜蘭運動公園",     "tourism", 240, 6.0,  12, "假日運動、慢跑、籃球"),
    ("國立傳藝中心",     "tourism", 170, 10.0, 18, "傳統藝術主題園區"),
    ("羅東夜市",         "tourism", 200, 11.0, 22, "在地最熱夜市"),
    ("礁溪溫泉",         "tourism", 350, 13.0, 22, "週末溫泉"),
    ("龜山島",           "tourism", 60,  20.0, 0,  "東部象徵地標"),

    # ── 對外交通
    ("國道5 壯圍交流道", "transit", 270, 4.0,  8,  "上國道 5"),
    ("宜蘭火車站",       "transit", 245, 7.0,  13, "台鐵宜蘭線"),
    ("宜蘭轉運站",       "transit", 245, 7.5,  14, "葛瑪蘭、首都客運往台北"),
    ("台北車站（雪隧）", "transit", 280, 70.0, 50, "國道5經雪山隧道"),
]

CATEGORY_META = {
    "school":   ("學區",   ACCENT_RED,        "●"),
    "shop":     ("購物",   DARK_BROWN,        "■"),
    "medical":  ("醫療",   RGBColor(0x2D, 0x6C, 0x4A), "✚"),
    "tourism":  ("景點",   RGBColor(0xB8, 0x6B, 0x1E), "★"),
    "transit":  ("交通",   RGBColor(0x4A, 0x4A, 0x88), "◆"),
}


def draw_radial_chart(slide, cx, cy, *, max_radius_in=3.3, title=None):
    """繪製光束圖：基地為中心，5 圈距離環，8 方位線，amenity 點."""
    # 5 圈半徑（mapping: km → inches via log-like compression）
    ring_km = [1, 3, 5, 10, 30]
    ring_r = [0.55, 1.2, 1.85, 2.55, max_radius_in]

    # 8 方位線
    for ang_deg in range(0, 360, 45):
        ang = math.radians(ang_deg - 90)  # 0 deg = North = up
        x_end = cx + ring_r[-1] * math.cos(ang)
        y_end = cy + ring_r[-1] * math.sin(ang)
        # 用線條（用很細的矩形模擬）
        line = slide.shapes.add_connector(1, Inches(cx), Inches(cy),
                                           Inches(x_end), Inches(y_end))
        line.line.color.rgb = GRAY_40
        line.line.width = Pt(0.5)

    # 同心圓
    for i, r in enumerate(ring_r):
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
            Inches(cx - r), Inches(cy - r), Inches(r * 2), Inches(r * 2))
        circle.fill.background()
        circle.line.color.rgb = GRAY_40
        circle.line.width = Pt(0.75 if i in (0, 4) else 0.4)
        circle.shadow.inherit = False
        # 距離標籤
        label_x = cx - 0.25
        label_y = cy - r - 0.18
        add_text(slide, Inches(label_x), Inches(label_y), Inches(0.5), Inches(0.2),
                 f"{ring_km[i]}km", size=8, color=GRAY_40, align="center")

    # 方位標籤 N/E/S/W
    for ang_deg, label in [(0, "N"), (90, "E"), (180, "S"), (270, "W")]:
        ang = math.radians(ang_deg - 90)
        r = ring_r[-1] + 0.15
        lx = cx + r * math.cos(ang) - 0.15
        ly = cy + r * math.sin(ang) - 0.12
        add_text(slide, Inches(lx), Inches(ly), Inches(0.3), Inches(0.24),
                 label, size=11, bold=True, color=GRAY_70, align="center")

    # 中心點 = 基地
    add_rect(slide, Inches(cx - 0.18), Inches(cy - 0.18),
             Inches(0.36), Inches(0.36), fill_color=ACCENT_RED, line_color=WHITE)
    add_text(slide, Inches(cx - 0.55), Inches(cy + 0.22),
             Inches(1.1), Inches(0.3),
             "中道森活", size=10, bold=True, color=ACCENT_RED, align="center")

    # 放置 amenity 點
    placed_positions = []  # for collision avoidance
    for name, cat, ang_deg, km, _drive, _note in AMENITIES:
        cat_meta = CATEGORY_META[cat]
        # km → r mapping (linear within ring, log-like compress for large)
        if km <= 1: r = 0.4 + km * 0.15
        elif km <= 3: r = 0.55 + (km - 1) / 2 * 0.65
        elif km <= 5: r = 1.2 + (km - 3) / 2 * 0.65
        elif km <= 10: r = 1.85 + (km - 5) / 5 * 0.7
        elif km <= 30: r = 2.55 + (km - 10) / 20 * 0.7
        else: r = max_radius_in
        if r > max_radius_in: r = max_radius_in
        ang = math.radians(ang_deg - 90)
        x = cx + r * math.cos(ang)
        y = cy + r * math.sin(ang)
        # marker dot
        dot_r = 0.08
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
            Inches(x - dot_r), Inches(y - dot_r),
            Inches(dot_r * 2), Inches(dot_r * 2))
        dot.fill.solid()
        dot.fill.fore_color.rgb = cat_meta[1]
        dot.line.color.rgb = WHITE
        dot.line.width = Pt(1)
        dot.shadow.inherit = False
        # label 自動偏移避免重疊
        # 依據角度決定文字位置
        label_w = 1.4
        if math.cos(ang) > 0.3:  # right side
            lx = x + 0.13; align = "left"
        elif math.cos(ang) < -0.3:  # left side
            lx = x - label_w - 0.13; align = "right"
        else:  # top/bottom
            lx = x - label_w / 2; align = "center"
        ly = y - 0.08
        add_text(slide, Inches(lx), Inches(ly), Inches(label_w), Inches(0.22),
                 name, size=8, color=DARK_BROWN, bold=True, align=align)

    if title:
        add_text(slide, Inches(cx - 3), Inches(cy + max_radius_in + 0.4),
                 Inches(6), Inches(0.3),
                 title, size=11, color=GRAY_70, align="center", bold=True)


def build():
    prs = new_presentation()

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="機能、光束圖與故事簡報　Location Story & Amenity Map",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　基地地理脈絡：宜蘭縣壯圍鄉五權段",
        "02　機能圈分層概念 — 5 圈生活範圍",
        "03　第 1 圈：步行 5 分鐘（生活基本盤）",
        "04　第 2 圈：開車 5 分鐘（社區生活圈）",
        "05　第 3 圈：開車 10 分鐘（城際生活圈）",
        "06　第 4 圈：開車 20 分鐘（跨鄉鎮）",
        "07　第 5 圈：高速公路對外（雪山隧道）",
        "08　學區深掘：公立國小、國中",
        "09　賣場 / 超市 / 便利商店",
        "10　醫療資源",
        "11　觀光景點（假日生活圈）",
        "12　光束圖（核心視覺）— 完整機能盤點",
        "13　主力客群故事：林先生的一天",
        "14　下一步：實地踏勘與圖像化",
    ]
    add_bullets(s, Inches(1.0), Inches(1.6), Inches(11), Inches(5.5),
                items, size=17, line_space=1.4)

    # ── 1. 地理脈絡
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　基地地理脈絡",
                  "宜蘭縣壯圍鄉五權段 789、790 等 45 筆地號")

    # 三層脈絡
    layers = [
        ("縣域",  "宜蘭縣",
         "全縣人口約 45 萬｜面積 2,143 km²｜雪山隧道通車後台北通勤圈"),
        ("鄉鎮",  "壯圍鄉",
         "人口約 23,000｜近宜蘭市東側｜居住純度高｜農業 + 住宅混合"),
        ("地段",  "五權段（中央路三段沿線）",
         "近鄉中心、近國道 5 壯圍交流道｜南北五路、忠孝一路鄰接"),
    ]
    y0 = 1.85
    for i, (level, name, desc) in enumerate(layers):
        y = y0 + i * 1.5
        add_rect(s, Inches(0.5), Inches(y), Inches(2.0), Inches(1.3),
                 fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
        add_text(s, Inches(0.5), Inches(y + 0.15), Inches(2.0), Inches(0.4),
                 level, size=14, color=DARK_BROWN, align="center", bold=True)
        add_text(s, Inches(0.5), Inches(y + 0.6), Inches(2.0), Inches(0.6),
                 name, size=20, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
        add_rect(s, Inches(2.7), Inches(y), Inches(10.1), Inches(1.3),
                 fill_color=CREAM, line_color=DARK_BROWN)
        add_text(s, Inches(2.9), Inches(y + 0.45), Inches(9.7), Inches(0.5),
                 desc, size=13, color=DARK_BROWN, line_space=1.4, anchor="middle")

    add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.5),
             "地理優勢：距國道 5 壯圍交流道 4 km｜距宜蘭市區 6~8 km｜距台北 70 km / 50 min",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_footer(s, "資料來源：DWG 建照圖、戶政司鄉鎮人口統計（業主請以實際座標核校距離）")
    add_page_number(s, 2, TOTAL)

    # ── 2. 5 圈概念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　機能圈分層概念",
                  "依「步行 → 開車 → 高速公路」分為 5 圈")

    rings = [
        ("圈 1", "步行 5 分鐘",  "0~1 km", "便利商店、傳統市場、近鄰社區", BRAND_YELLOW),
        ("圈 2", "開車 5 分鐘",  "1~3 km", "學區、衛生所、全聯", BRAND_YELLOW_SOFT),
        ("圈 3", "開車 10 分鐘", "3~5 km", "壯圍鄉公所、家樂福、國道交流道", CREAM),
        ("圈 4", "開車 20 分鐘", "5~10 km", "宜蘭市區、火車站、夢時代", WHITE),
        ("圈 5", "高速公路",     "70 km",   "台北（雪山隧道 50 分鐘）", CREAM),
    ]
    bw = 2.45; bh = 4.5; y0 = 1.85
    for i, (label, when, dist, content, fill) in enumerate(rings):
        x = 0.5 + i * (bw + 0.1)
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(x), Inches(y0 + 0.15), Inches(bw), Inches(0.5),
                 label, size=14, bold=True, color=ACCENT_RED, align="center")
        add_text(s, Inches(x), Inches(y0 + 0.6), Inches(bw), Inches(0.65),
                 when, size=18, bold=True, color=text_color, align="center", font=FONT_TITLE)
        add_text(s, Inches(x), Inches(y0 + 1.35), Inches(bw), Inches(0.45),
                 dist, size=13, color=text_color, align="center", bold=True)
        # divider
        add_rect(s, Inches(x + bw/2 - 0.4), Inches(y0 + 1.95),
                 Inches(0.8), Pt(1), fill_color=text_color)
        add_text(s, Inches(x + 0.15), Inches(y0 + 2.2),
                 Inches(bw - 0.3), Inches(2.2),
                 content, size=11.5, color=text_color, align="center", line_space=1.5)

    add_footer(s, "下五頁逐一展開每一圈的具體機能")
    add_page_number(s, 3, TOTAL)

    # ── 3~7. 五圈逐一展開
    rings_detail = [
        ("03　第 1 圈　步行 5 分鐘", "0~1 km｜出門即達",
         [
             ["7-11 壯圍門市", "便利商店", "1.0 km", "3 分鐘"],
             ["全家 五權店",    "便利商店", "1.0 km", "3 分鐘"],
             ["地方早餐店",    "傳統早餐", "0.5 km", "步行"],
             ["鄰里公園",      "綠地",     "0.3 km", "步行"],
         ]),
        ("04　第 2 圈　開車 5 分鐘", "1~3 km｜社區生活",
         [
             ["公館國小",        "學區",     "1.5 km", "4 分鐘"],
             ["壯圍國小",        "學區",     "2.0 km", "5 分鐘"],
             ["過嶺國小",        "學區",     "2.5 km", "6 分鐘"],
             ["全聯 壯圍店",     "超市",     "2.5 km", "5 分鐘"],
             ["壯圍傳統市場",    "市場",     "2.0 km", "5 分鐘"],
             ["壯圍鄉衛生所",    "醫療",     "2.5 km", "6 分鐘"],
             ["壯圍沙丘 旅服中心", "景點",     "3.0 km", "7 分鐘"],
         ]),
        ("05　第 3 圈　開車 10 分鐘", "3~5 km｜城際生活",
         [
             ["國道 5 壯圍交流道", "對外交通", "4.0 km", "8 分鐘"],
             ["永鎮海濱公園",      "景點",     "4.5 km", "10 分鐘"],
             ["全聯 宜蘭東港店",    "超市",     "5.0 km", "10 分鐘"],
             ["壯圍國中",          "學區",     "5.0 km", "10 分鐘"],
         ]),
        ("06　第 4 圈　開車 20 分鐘", "5~10 km｜跨鄉鎮",
         [
             ["宜蘭運動公園",      "運動景點", "6.0 km", "12 分鐘"],
             ["中華國中",          "宜蘭市國中", "6.0 km", "12 分鐘"],
             ["宜蘭夢時代",        "百貨",     "6.5 km", "13 分鐘"],
             ["國立陽明交大附醫",  "醫院",     "7.0 km", "14 分鐘"],
             ["宜蘭火車站",        "火車",     "7.0 km", "13 分鐘"],
             ["宜蘭轉運站",        "客運",     "7.5 km", "14 分鐘"],
             ["家樂福 宜蘭店",     "量販",     "8.0 km", "15 分鐘"],
             ["國立傳藝中心",      "景點",     "10.0 km", "18 分鐘"],
         ]),
        ("07　第 5 圈　高速公路對外", "10 ~ 70 km｜對外連結",
         [
             ["羅東博愛醫院",      "醫院",     "12 km",  "20 分鐘"],
             ["礁溪溫泉",          "週末景點", "13 km",  "22 分鐘"],
             ["羅東夜市",          "景點",     "11 km",  "22 分鐘"],
             ["龜山島賞鯨",        "景點",     "20 km",  "搭船"],
             ["台北車站（雪隧）",  "對外",     "70 km",  "50 分鐘"],
         ]),
    ]
    for idx, (title, sub, items) in enumerate(rings_detail):
        s = add_blank_slide(prs); set_background(s, WHITE)
        add_title_bar(s, title, sub)
        # 表格
        headers = ["機能", "類別", "距離", "車程"]
        add_table(s, Inches(0.5), Inches(1.7), Inches(12.3),
                  Inches(0.45 + 0.45 * len(items)),
                  headers, items,
                  header_size=12, body_size=12, first_col_bold=True, body_align="center")
        # 結語
        commentary = [
            "圈 1 = 不用開車的日常 — 出門有早餐、有便利商店、有公園 = 生活基本盤穩固。",
            "圈 2 = 三所國小都在 5 分鐘車程內，是壯圍少見的「跨學區彈性區位」。",
            "圈 3 = 進入國道 5 只要 8 分鐘，這是與宜蘭市區、雙北通勤的關鍵節點。",
            "圈 4 = 宜蘭市區的火車、客運、購物、醫療皆 < 15 分鐘 — 「住壯圍、用宜蘭市」。",
            "圈 5 = 雙北 50 分鐘可達 — 雪隧讓壯圍成為合理的「半遠距通勤」選項。",
        ][idx]
        y_comment = 1.7 + 0.45 + 0.45 * len(items) + 0.4
        add_text(s, Inches(0.5), Inches(y_comment),
                 Inches(12.3), Inches(0.5),
                 commentary, size=13, color=DARK_BROWN, bold=True)
        add_footer(s, "距離 / 車程為合理估算，實際依路況與業主踏勘為準")
        add_page_number(s, 4 + idx, TOTAL)

    # ── 8. 學區深掘
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　學區深掘", "公立國小、國中｜對 TA1 返鄉換屋族最關鍵")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.0),
              ["學校", "公私立", "距離", "車程 / 步行", "學區歸屬", "對應 TA"],
              [
                  ["公館國小",       "公立國小", "1.5 km", "4 分鐘車｜走 18 分", "學區內或邊界", "TA1 / TA2"],
                  ["壯圍國小",       "公立國小", "2.0 km", "5 分鐘車｜走 25 分", "鄰近學區",     "TA1 / TA2"],
                  ["過嶺國小",       "公立國小", "2.5 km", "6 分鐘車",          "另一學區",     "TA1 / TA2"],
                  ["壯圍國中",       "公立國中", "5.0 km", "10 分鐘車",         "全鄉合一國中", "TA1"],
                  ["中華國中",       "公立國中", "6.0 km", "12 分鐘車",         "宜蘭市可越區", "高端 TA1"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="center")

    add_rect(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(1.9),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.1), Inches(12), Inches(0.4),
             "學區優勢 — 給 TA1 返鄉換屋族的最大誘因", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.55), Inches(11.9), Inches(1.3), [
        "3 所國小都在 6 分鐘車程內 — 全壯圍唯有少數區位有此條件",
        "壯圍國中為全鄉唯一公立國中 — 連帶學區不需越區",
        "宜蘭市中華國中越區也可達（12 分鐘）— 高端教育選項",
        "對返鄉客而言：「孩子上學的問題不用擔心」就是最高訴求",
    ], size=12)
    add_footer(s, "學區歸屬待業主向宜蘭縣教育處查證最新版（國小學區圖每 3 年更新）")
    add_page_number(s, 9, TOTAL)

    # ── 9. 賣場 / 超市
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　賣場 / 超市 / 便利商店",
                  "日常採購 → 週末大採購一站到位")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["商家", "類別", "距離", "車程", "用途"],
              [
                  ["7-11 壯圍門市",      "便利商店", "1.0 km", "3 分",  "宵夜、繳費、咖啡"],
                  ["全家 五權店",        "便利商店", "1.0 km", "3 分",  "宵夜、領貨"],
                  ["全聯 壯圍店",        "超市",     "2.5 km", "5 分",  "日常採購"],
                  ["壯圍傳統市場",       "市場",     "2.0 km", "5 分",  "新鮮蔬果、肉品"],
                  ["全聯 宜蘭東港店",    "超市",     "5.0 km", "10 分", "中型採購"],
                  ["宜蘭夢時代",         "百貨",     "6.5 km", "13 分", "週末逛街、餐廳"],
                  ["家樂福 宜蘭店",      "量販",     "8.0 km", "15 分", "週末大採購"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="center")

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "結論：1 km 內有便利商店｜2.5 km 內有全聯與市場｜8 km 內有量販與百貨",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 10, TOTAL)

    # ── 10. 醫療
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "10　醫療資源",
                  "從基本診療到區域醫院 — 對換屋族與退休族同樣重要")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(2.5),
              ["醫療機構", "層級", "距離", "車程", "用途"],
              [
                  ["壯圍鄉衛生所",         "基層",   "2.5 km",  "6 分",  "疫苗、家醫、長者照護"],
                  ["國立陽明交大附醫",     "區域醫院", "7.0 km",  "14 分", "宜蘭最大公立、24h 急診"],
                  ["羅東博愛醫院",         "區域醫院", "12.0 km", "20 分", "重症、專科"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="center")

    add_rect(s, Inches(0.5), Inches(4.8), Inches(12.3), Inches(2.0),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.9), Inches(12), Inches(0.4),
             "醫療優勢", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.35), Inches(11.9), Inches(1.5), [
        "陽明交大附醫 14 分鐘內可達 — 是宜蘭最大公立醫院，急診全天候",
        "羅東博愛醫院 20 分鐘 — 全宜蘭最強的私立區域醫院",
        "對 TA1（爸媽就近照顧）+ TA3（自己退休）皆是必要條件",
        "與雙北「醫療密集 vs 通勤遠」對比，這個距離已是「夠近又夠便宜」的甜蜜點",
    ], size=12)
    add_page_number(s, 11, TOTAL)

    # ── 11. 觀光景點
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "11　觀光景點（假日生活圈）",
                  "從步行可達的沙丘到 20 分鐘的羅東 — 假日不用出遠門")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["景點", "類別", "距離", "車程", "假日用途"],
              [
                  ["壯圍沙丘 旅服中心",   "建築/地景", "3.0 km",  "7 分",   "黃聲遠設計，地標必訪"],
                  ["永鎮海濱公園",       "海邊",      "4.5 km",  "10 分",  "親子戲沙、看海"],
                  ["宜蘭運動公園",       "運動",      "6.0 km",  "12 分",  "慢跑、籃球、騎車"],
                  ["國立傳藝中心",       "文化園區",  "10.0 km", "18 分",  "傳統藝術主題、年節最熱"],
                  ["羅東夜市",           "夜市",      "11.0 km", "22 分",  "週末晚餐"],
                  ["礁溪溫泉",           "溫泉",      "13.0 km", "22 分",  "週末泡湯"],
                  ["龜山島賞鯨",         "海上活動",  "20.0 km", "搭船",   "夏天賞鯨"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="center")
    add_text(s, Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.4),
             "「住壯圍 = 假日都不用出宜蘭」 — 是說服 TA3 退休族的關鍵賣點",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 12, TOTAL)

    # ── 12. 光束圖（核心視覺）
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "12　光束圖 — 完整機能盤點",
                  "以基地為中心，5 圈距離環、8 方位、26 個機能點")
    # 圖區
    draw_radial_chart(s, cx=4.8, cy=4.4, max_radius_in=3.3,
                      title=None)

    # 右側圖例與摘要
    legend_x = 9.2
    add_text(s, Inches(legend_x), Inches(1.7), Inches(3.8), Inches(0.4),
             "圖例", size=14, bold=True, color=DARK_BROWN)
    for i, (cat, (name, color, sym)) in enumerate(CATEGORY_META.items()):
        y = 2.15 + i * 0.4
        # color dot
        add_rect(s, Inches(legend_x), Inches(y + 0.05),
                 Inches(0.22), Inches(0.22), fill_color=color, line_color=WHITE)
        add_text(s, Inches(legend_x + 0.35), Inches(y), Inches(3.4), Inches(0.32),
                 name, size=12, bold=True, color=DARK_BROWN, anchor="middle")

    # 摘要框
    add_rect(s, Inches(legend_x), Inches(4.4), Inches(3.8), Inches(2.6),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(legend_x + 0.15), Inches(4.5), Inches(3.5), Inches(0.4),
             "三句話總結機能", size=13, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(legend_x + 0.15), Inches(4.95), Inches(3.5), Inches(2.0), [
        "1 km：便利商店、市場、公園",
        "5 km：學區、超市、衛生所",
        "10 km：宜蘭市區、火車、醫院",
        "50 min：雙北雪山隧道",
    ], size=10.5, line_space=1.4)

    add_footer(s, "距離為合理估算｜方位為粗略示意｜業主請以實際座標核校")
    add_page_number(s, 13, TOTAL)

    # ── 13. 主力客群故事
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "13　主力客群故事", "林先生的一天｜給 TA1 返鄉換屋族的敘事")

    # 故事左右雙欄
    add_rect(s, Inches(0.5), Inches(1.7), Inches(6.0), Inches(5.3),
             fill_color=WHITE, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(1.85), Inches(5.6), Inches(0.5),
             "林先生・34 歲｜軟體工程師", size=15, bold=True, color=DARK_BROWN, font=FONT_TITLE)
    add_text(s, Inches(0.7), Inches(2.35), Inches(5.6), Inches(0.4),
             "宜蘭壯圍出身．台北工作 8 年．兩個孩子 6 歲、3 歲", size=11, color=GRAY_70)
    story = ("「07:30　爸爸把孩子送到 4 分鐘車程的公館國小。\n"
             "08:15　搭區間車到宜蘭轉運站、轉葛瑪蘭，09:30 進台北辦公室。\n"
             "12:00　午餐時 LINE 媽媽：『中午回去吃飯嗎？』媽媽住宜蘭市，10 分鐘車程。\n"
             "18:30　離開公司、雪隧暢通，20:00 回家。\n"
             "19:00　全家走到鄰里公園散步，帶孩子買 7-11 冰淇淋。\n"
             "週六　全家去壯圍沙丘野餐、看海，下午到夢時代購物。\n"
             "週日　全家去羅東夜市吃晚餐，22:00 回家睡覺。\n"
             "\n"
             "「在台北 8 年，我們每月房租 30,000、通勤 90 分鐘。"
             "搬回壯圍後：房貸 30,000、通勤一週兩次。"
             "省下的時間，是和爸媽、孩子的時間。」")
    add_text(s, Inches(0.7), Inches(2.85), Inches(5.6), Inches(4.0),
             story, size=11, color=DARK_BROWN, line_space=1.55)

    # 右欄：故事的機能依據
    add_rect(s, Inches(6.7), Inches(1.7), Inches(6.1), Inches(5.3),
             fill_color=BRAND_YELLOW_SOFT, line_color=ACCENT_RED)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.7), Inches(0.5),
             "每一句的機能依據", size=15, bold=True, color=ACCENT_RED, font=FONT_TITLE)
    facts = [
        ("送孩子上學",     "公館 / 壯圍國小 4~5 分鐘車"),
        ("通勤台北",       "國道 5 壯圍交流道 8 分鐘 → 雪隧 50 分鐘"),
        ("中午回家吃飯",   "媽媽住宜蘭市，10 分鐘車"),
        ("傍晚散步買冰",   "7-11 / 鄰里公園步行可達"),
        ("週六海邊野餐",   "壯圍沙丘 / 永鎮海濱 7~10 分鐘"),
        ("週六下午購物",   "宜蘭夢時代 13 分鐘"),
        ("週日羅東夜市",   "20 分鐘車程"),
        ("爸媽就近照顧",   "陽明交大附醫 14 分鐘"),
    ]
    for i, (act, fact) in enumerate(facts):
        y_f = 2.45 + i * 0.55
        # row alternation
        if i % 2 == 0:
            add_rect(s, Inches(6.85), Inches(y_f), Inches(5.85), Inches(0.5),
                     fill_color=WHITE)
        add_text(s, Inches(7.0), Inches(y_f + 0.07), Inches(2.0), Inches(0.4),
                 act, size=11, bold=True, color=DARK_BROWN)
        add_text(s, Inches(9.1), Inches(y_f + 0.07), Inches(3.5), Inches(0.4),
                 fact, size=11, color=GRAY_70)

    add_footer(s, "故事為璞域撰寫，依本簡報機能盤點為依據 — 業主可任意改編作為文案核心")
    add_page_number(s, 14, TOTAL)

    # ── 14. 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "14　下一步行動",
                  "從本簡報的機能假設 → 落地的視覺輸出")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "工項", "璞域提供", "業主提供"],
              [
                  ["1", "實地踏勘", "陪同（半天）",
                   "車輛、聯絡窗口"],
                  ["2", "機能距離精校", "驗證並更新本簡報",
                   "實際步測 / 行車紀錄"],
                  ["3", "學區查證", "向縣教育處正式詢問",
                   "—（資料可公開）"],
                  ["4", "機能地圖視覺化", "依本簡報資料製作正式版「機能地圖海報」",
                   "選定設計風格"],
                  ["5", "故事文案精修", "從故事擴寫為廣告 / DM 文案系列",
                   "業主回饋方向"],
                  ["6", "短影音腳本", "依故事改寫成 30 秒影音腳本",
                   "確認預算"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "預計兩週內可完成踏勘 + 精校 + 第一版機能地圖海報",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 15, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
