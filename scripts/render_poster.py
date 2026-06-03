"""中道森活 機能地圖海報 — A2 直式 (420 × 594 mm @ 150 dpi).

輸出：assets/posters/中道森活_機能地圖海報_A2.png  (約 2480 × 3508 px)

設計：
  - 上 1/4：品牌標題 + slogan
  - 中 1/2：光束圖（radial），周圍標示 26 個機能點
  - 下 1/4：林先生的一天時間軸 + 鼎弘建設資訊
"""
from pathlib import Path
import math
from PIL import Image, ImageDraw, ImageFont

ROOT = Path("/home/user/Pure-space-Branding")
OUT_DIR = ROOT / "assets/posters"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT = OUT_DIR / "中道森活_機能地圖海報_A2.png"

# A2 portrait @ 150 dpi
W, H = 2480, 3508
DPI = 150

# Brand palette
BG       = (255, 248, 231)   # cream
YELLOW   = (232, 172, 27)
DARK     = (74,  50,  16)
ACCENT   = (139, 58,  15)
GRAY70   = (85,  85,  85)
GRAY40   = (160, 160, 160)
WHITE    = (255, 255, 255)
RING_GRAY = (200, 195, 180)

# Category colors
CAT = {
    "school":  ACCENT,
    "shop":    DARK,
    "medical": (45, 108, 74),
    "tourism": (184, 107, 30),
    "transit": (74, 74, 136),
}

FONT_BOLD = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc"
FONT_REG  = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_HEAVY = "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc"

def f(size, weight="reg"):
    p = FONT_HEAVY if weight == "heavy" else (FONT_BOLD if weight == "bold" else FONT_REG)
    return ImageFont.truetype(p, size, index=2)  # index 2 = TC (Traditional Chinese)


def text(draw, xy, s, *, font, fill=DARK, anchor="la"):
    draw.text(xy, s, font=font, fill=fill, anchor=anchor)


# Amenity list — same as deck 8, but with slight angle offsets to spread clusters
AMENITIES = [
    ("壯圍國小",        "school", 300, 2.0),
    ("公館國小",        "school", 330, 1.5),
    ("過嶺國小",        "school", 120, 2.5),
    ("壯圍國中",        "school", 180, 5.0),
    ("中華國中",        "school", 220, 6.0),
    ("全聯壯圍店",      "shop",   270, 2.5),
    ("全聯宜蘭東港店",  "shop",   235, 5.0),
    ("7-11壯圍",       "shop",   295, 1.0),
    ("全家五權店",      "shop",   100, 1.0),
    ("壯圍傳統市場",    "shop",   258, 2.0),
    ("家樂福宜蘭店",    "shop",   258, 8.0),
    ("宜蘭夢時代",      "shop",   230, 6.5),
    ("壯圍鄉衛生所",    "medical",285, 2.5),
    ("羅東博愛醫院",    "medical",205, 12.0),
    ("陽明交大附醫",    "medical",250, 7.0),
    ("壯圍沙丘",        "tourism",90,  3.0),
    ("永鎮海濱公園",    "tourism",110, 4.5),
    ("宜蘭運動公園",    "tourism",215, 6.0),
    ("國立傳藝中心",    "tourism",175, 10.0),
    ("羅東夜市",        "tourism",195, 11.0),
    ("礁溪溫泉",        "tourism",355, 13.0),
    ("龜山島",          "tourism",70,  20.0),
    ("國道5壯圍交流道", "transit",278, 4.0),
    ("宜蘭火車站",      "transit",265, 7.0),
    ("宜蘭轉運站",      "transit",272, 7.5),
    ("台北車站(雪隧)", "transit",295, 70.0),
]

def km_to_r(km, r_max):
    """km → 圖上半徑（px）— 分段壓縮，遠處越壓越緊"""
    if km <= 1: t = 0.18 * km
    elif km <= 3: t = 0.18 + 0.13 * (km - 1) / 2
    elif km <= 5: t = 0.31 + 0.12 * (km - 3) / 2
    elif km <= 10: t = 0.43 + 0.18 * (km - 5) / 5
    elif km <= 30: t = 0.61 + 0.25 * (km - 10) / 20
    else: t = 0.86 + 0.12 * min((km - 30) / 40, 1)
    return t * r_max


def draw_radial(img, draw, cx, cy, r_max):
    """畫光束圖到指定中心"""
    # 8 方位線
    for ang_deg in range(0, 360, 45):
        ang = math.radians(ang_deg - 90)
        x_end = cx + r_max * math.cos(ang)
        y_end = cy + r_max * math.sin(ang)
        draw.line([(cx, cy), (x_end, y_end)], fill=RING_GRAY, width=2)

    # 5 圈
    rings_km = [1, 3, 5, 10, 30]
    for km in rings_km:
        r = km_to_r(km, r_max)
        draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                      outline=RING_GRAY, width=2)
        # 距離標籤
        text(draw, (cx + 20, cy - r - 6), f"{km}km",
             font=f(28), fill=GRAY40, anchor="ls")

    # 方位 N/E/S/W
    fn = f(50, "bold")
    for ang_deg, lbl in [(0, "N"), (90, "E"), (180, "S"), (270, "W")]:
        ang = math.radians(ang_deg - 90)
        r = r_max + 80
        lx = cx + r * math.cos(ang)
        ly = cy + r * math.sin(ang)
        text(draw, (lx, ly), lbl, font=fn, fill=GRAY70, anchor="mm")

    # Amenity points — group + offset to avoid label overlap
    fa = f(42, "bold")
    placed_labels = []  # (lx, ly, w_est) for collision check

    def overlaps(lx, ly, w_est=200, h_est=50):
        for px, py, pw in placed_labels:
            if abs(py - ly) < h_est and (
                lx < px + pw + 20 and lx + w_est + 20 > px
            ):
                return True
        return False

    for name, cat, ang_deg, km in AMENITIES:
        r = km_to_r(km, r_max)
        ang = math.radians(ang_deg - 90)
        x = cx + r * math.cos(ang)
        y = cy + r * math.sin(ang)
        # dot
        dot_r = 22
        color = CAT[cat]
        draw.ellipse([x - dot_r, y - dot_r, x + dot_r, y + dot_r],
                      fill=color, outline=WHITE, width=4)
        # label
        ox, oy = math.cos(ang), math.sin(ang)
        w_est = len(name) * 48
        if ox > 0.3:
            anchor = "lm"; lx = x + 36; ly = y
        elif ox < -0.3:
            anchor = "rm"; lx = x - 36 - w_est; ly = y
        else:
            anchor = "mb" if oy < 0 else "mt"
            lx = x - w_est / 2; ly = y + (-40 if oy < 0 else 40)

        # vertical nudge to avoid overlap
        nudge = 0
        while overlaps(lx, ly + nudge, w_est) and abs(nudge) < 250:
            nudge += 56 if oy >= 0 else -56
        ly += nudge

        if anchor == "lm":
            text_xy = (x + 36, ly)
        elif anchor == "rm":
            text_xy = (x - 36, ly)
        else:
            text_xy = (x, ly)
        text(draw, text_xy, name, font=fa, fill=DARK, anchor=anchor)
        placed_labels.append((lx, ly, w_est))

    # 中心 — 基地
    side = 60
    draw.rectangle([cx - side/2, cy - side/2, cx + side/2, cy + side/2],
                    fill=ACCENT, outline=WHITE, width=4)
    text(draw, (cx, cy + side/2 + 50), "中道森活",
         font=f(54, "bold"), fill=ACCENT, anchor="ma")


def main():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    # 上邊條
    draw.rectangle([0, 0, W, 36], fill=YELLOW)

    # 標題區
    title_y = 180
    text(draw, (W // 2, title_y), "中　道　森　活",
         font=f(220, "bold"), fill=DARK, anchor="ma")
    text(draw, (W // 2, title_y + 290), "理想居所　最懂生活",
         font=f(80, "bold"), fill=DARK, anchor="ma")
    text(draw, (W // 2, title_y + 410), "在繁華之中收藏寧靜　在日常之間遇見家的溫度",
         font=f(48), fill=GRAY70, anchor="ma")

    # 機能圖
    chart_cx = W // 2
    chart_cy = 1700
    r_max = 950
    draw_radial(img, draw, chart_cx, chart_cy, r_max)

    # 圖例（左下角）
    legend_x = 120
    legend_y = chart_cy + r_max + 200
    text(draw, (legend_x, legend_y), "圖例", font=f(48, "bold"), fill=DARK, anchor="la")
    legend_items = [
        ("學區", CAT["school"]),
        ("購物", CAT["shop"]),
        ("醫療", CAT["medical"]),
        ("景點", CAT["tourism"]),
        ("交通", CAT["transit"]),
    ]
    for i, (lbl, color) in enumerate(legend_items):
        yy = legend_y + 90 + i * 70
        draw.ellipse([legend_x, yy, legend_x + 36, yy + 36],
                      fill=color, outline=WHITE, width=2)
        text(draw, (legend_x + 60, yy + 18), lbl,
             font=f(38, "bold"), fill=DARK, anchor="lm")

    # 重點數字 (右下角)
    metrics = [
        ("4 分鐘", "公館國小 車程"),
        ("8 分鐘", "上國道5 雪隧"),
        ("14 分鐘", "陽明交大附醫"),
        ("50 分鐘", "雙北 雪山隧道"),
    ]
    box_x = W - 800
    box_y = chart_cy + r_max + 200
    text(draw, (box_x, box_y), "關鍵車程", font=f(48, "bold"), fill=ACCENT, anchor="la")
    for i, (num, lbl) in enumerate(metrics):
        yy = box_y + 90 + i * 110
        text(draw, (box_x, yy), num,
             font=f(64, "heavy"), fill=DARK, anchor="la")
        text(draw, (box_x + 350, yy + 14), lbl,
             font=f(36), fill=GRAY70, anchor="la")

    # 故事時間軸 (底部)
    story_y = H - 720
    draw.rectangle([60, story_y, W - 60, H - 200], outline=DARK, width=3)
    text(draw, (120, story_y + 60), "返鄉換屋族・林先生的一天",
         font=f(54, "bold"), fill=ACCENT, anchor="la")
    timeline = [
        ("07:30", "送孩子到公館國小　4 分鐘車"),
        ("08:15", "宜蘭轉運站→葛瑪蘭→台北 50 分鐘"),
        ("12:00", "媽媽在宜蘭市　10 分鐘車回家吃飯"),
        ("19:00", "全家走鄰里公園散步買 7-11"),
        ("週六", "壯圍沙丘野餐＋夢時代購物 13 分鐘"),
        ("週日", "羅東夜市晚餐 22 分鐘"),
    ]
    for i, (t, content) in enumerate(timeline):
        col = i % 3
        row = i // 3
        cx = 130 + col * 770
        cy = story_y + 200 + row * 130
        text(draw, (cx, cy), t,
             font=f(56, "heavy"), fill=ACCENT, anchor="la")
        text(draw, (cx + 190, cy + 12), content,
             font=f(36), fill=DARK, anchor="la")

    # 底部識別資訊
    footer_y = H - 150
    text(draw, (W // 2, footer_y), "鼎弘建設　壯圍鄉五權段　27 戶透天　2026 預售",
         font=f(36, "bold"), fill=DARK, anchor="ma")
    text(draw, (W // 2, footer_y + 70), "璞域品牌策略｜www.pure-branding.com",
         font=f(28), fill=GRAY70, anchor="ma")

    # 下邊條
    draw.rectangle([0, H - 36, W, H], fill=YELLOW)

    img.save(OUT, "PNG", optimize=True)
    print(f"wrote {OUT}  ({W}×{H} px, A2 @ {DPI} dpi)")
    # 也另存 PDF for printing
    pdf_out = OUT.with_suffix(".pdf")
    img.save(pdf_out, "PDF", resolution=DPI)
    print(f"wrote {pdf_out}")


if __name__ == "__main__":
    main()
