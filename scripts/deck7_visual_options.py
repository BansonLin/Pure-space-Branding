"""簡報 7：中道森活 — 8 色主視覺提案說明.

對應業主提供的 38fe0ccf-557748.jpg（8 版本提案）。
分割 + 命名 + 情緒分析 + 對應 TA 客群 + 推薦排序。

輸出：02_簡報/07_中道森活_主視覺8色提案說明.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "07_中道森活_主視覺8色提案說明.pptx"
ASSETS = REPO / "assets/visuals"
OUT.parent.mkdir(parents=True, exist_ok=True)

# 8 版本完整資料：(id, 命名, hex, RGB, 花卉系, 情緒詞, 主對應TA, 印刷適合, 數位適合, 評語)
VARIANTS = [
    (1, "酒紅·古典玫瑰",       "#8D3F32", "玫瑰系", ["沉穩", "典雅", "莊重"],
     "TA 3 退休置產",
     "★★★★★", "★★★",
     "高端文化氣質。如同博物館圖鑑封面，傳遞「傳家」「典藏」訊息。"),
    (2, "普魯士藍·水仙",       "#395277", "水仙系", ["靜謐", "知性", "克制"],
     "TA 3 退休置產 / 高端 TA 1",
     "★★★★★", "★★★★",
     "深藍×黃花的歐洲博物學插畫感。最有品牌獨特性，但暗色印刷成本較高。"),
    (3, "森林綠·古典玫瑰",     "#325C34", "玫瑰系", ["沉著", "自然", "穩定"],
     "TA 3 退休置產",
     "★★★★★", "★★★",
     "深綠承接「森活」字意，與宜蘭自然意象呼應。略偏成熟、年輕客層可能無感。"),
    (4, "霧灰藍·水仙",         "#C0CBD1", "水仙系", ["輕盈", "清新", "現代"],
     "TA 1 返鄉換屋 / TA 2 在地",
     "★★★", "★★★★★",
     "最現代、最年輕。網路曝光優勢大，但實體招牌略顯弱、品牌記憶點低。"),
    (5, "湖青·古典玫瑰",       "#3A7B90", "玫瑰系", ["寧靜", "深邃", "東方"],
     "TA 3 退休置產 / 高端 TA 1",
     "★★★★", "★★★★",
     "水墨意境＋玫瑰，東方氣質最濃。與業主希望的「在繁華中收藏寧靜」高度契合。"),
    (6, "橄欖綠·水仙",         "#707952", "水仙系", ["沉澱", "自然", "成熟"],
     "TA 1 返鄉換屋 / TA 3 退休",
     "★★★★", "★★★",
     "橄欖綠較少建案使用，識別度高。情緒偏低調奢華，跨客層接受度好。"),
    (7, "金黃·古典玫瑰",       "#EABE46", "玫瑰系", ["溫暖", "明亮", "豐盛"],
     "TA 1 返鄉換屋 / TA 2 在地",
     "★★★★", "★★★★★",
     "業主先前主推版本。最明亮溫暖，全客層皆能接受，戶外與數位皆強。"),
    (8, "金黃·水仙",           "#EBBD42", "水仙系", ["明亮", "歡愉", "親切"],
     "TA 1 返鄉換屋 / TA 2 在地",
     "★★★★", "★★★★★",
     "同 v7 色基底，但花材改水仙，視覺更輕盈年輕。與 v7 可作系列搭配。"),
]


def hex_to_rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_image_box(slide, x, y, w, h, image_path, line_color=DARK_BROWN):
    """插入圖片並加細邊框."""
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h),
             fill_color=WHITE, line_color=line_color)
    slide.shapes.add_picture(str(image_path),
                              Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def build():
    prs = new_presentation()
    TOTAL = 14

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="主視覺 8 色提案說明簡報　Visual Direction Review",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　8 版本一覽 — 8 個方向、2 個花卉系統",
        "02　兩大花卉系統 — 玫瑰系 vs 水仙系",
        "03~10　逐一說明（每版本：色彩、情緒、客群、適用通路）",
        "11　8 版本 × 客群對應矩陣",
        "12　TOP 3 推薦排序",
        "13　落地使用建議 — 哪個通路用哪個",
        "14　下一步行動",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ── 1. 8 版本一覽（一頁顯示 8 縮圖）
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　8 版本一覽",
                  "業主提供之 8 色彩 / 花卉組合提案")

    # 2 排 4 欄縮圖
    thumb_w = 2.95; thumb_h = 1.25
    cols = 4
    for i, v in enumerate(VARIANTS):
        row = i // cols
        col = i % cols
        x = 0.4 + col * (thumb_w + 0.12)
        y = 1.85 + row * (thumb_h + 0.7)
        # 縮圖
        add_image_box(s, x, y, thumb_w, thumb_h, ASSETS / f"v{v[0]}.jpg")
        # 標籤
        add_text(s, Inches(x), Inches(y + thumb_h + 0.05),
                 Inches(thumb_w), Inches(0.32),
                 f"v{v[0]}　{v[1]}", size=11, bold=True, color=DARK_BROWN, align="center")
        # hex 色塊
        chip_x = x + thumb_w / 2 - 0.4
        add_rect(s, Inches(chip_x), Inches(y + thumb_h + 0.4),
                 Inches(0.3), Inches(0.18), fill_color=hex_to_rgb(v[2]), line_color=DARK_BROWN)
        add_text(s, Inches(chip_x + 0.35), Inches(y + thumb_h + 0.37),
                 Inches(1.2), Inches(0.25),
                 v[2], size=10, color=GRAY_70, font="Courier New")

    add_footer(s, "縮圖切割自業主提供之 38fe0ccf-557748.jpg")
    add_page_number(s, 2, TOTAL)

    # ── 2. 兩大花卉系統
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　兩大花卉系統",
                  "8 版本可分為兩條設計路線 — 玫瑰系（典藏感）vs 水仙系（明朗感）")

    # 左：玫瑰系（v1, v3, v5, v7）
    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.3),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.5),
             "玫瑰系｜白玫瑰", size=20, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
    add_text(s, Inches(0.5), Inches(2.4), Inches(5.9), Inches(0.4),
             "v1 v3 v5 v7", size=11, color=GRAY_70, align="center")
    # 4 縮圖小排版
    for i, vid in enumerate([1, 3, 5, 7]):
        x = 0.7 + (i % 2) * 2.7
        y = 2.85 + (i // 2) * 1.6
        add_image_box(s, x, y, 2.5, 1.4, ASSETS / f"v{vid}.jpg")
    add_bullets(s, Inches(0.7), Inches(6.0), Inches(5.5), Inches(1.0), [
        "氣質：典雅、東方、靜謐",
        "對應：TA3 退休 ＋ 高端 TA1",
        "印刷感最強，戶外效果穩定",
    ], size=11.5)

    # 右：水仙系（v2, v4, v6, v8）
    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.3),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.5),
             "水仙系｜黃水仙＋鬱金香", size=20, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
    add_text(s, Inches(6.9), Inches(2.4), Inches(5.9), Inches(0.4),
             "v2 v4 v6 v8", size=11, color=GRAY_70, align="center")
    for i, vid in enumerate([2, 4, 6, 8]):
        x = 7.1 + (i % 2) * 2.7
        y = 2.85 + (i // 2) * 1.6
        add_image_box(s, x, y, 2.5, 1.4, ASSETS / f"v{vid}.jpg")
    add_bullets(s, Inches(7.1), Inches(6.0), Inches(5.5), Inches(1.0), [
        "氣質：明朗、輕盈、北歐插畫感",
        "對應：TA1 返鄉 ＋ TA2 在地",
        "數位/社群效果最強",
    ], size=11.5)

    add_footer(s, "兩系統可單獨使用，亦可組合（如「主視覺=玫瑰系、社群=水仙系」分通路投放）")
    add_page_number(s, 3, TOTAL)

    # ── 3~10. 逐一說明（每頁一版本）
    for page_idx, v in enumerate(VARIANTS):
        vid, name, hex_, system, moods, ta, pr, dg, note = v
        s = add_blank_slide(prs); set_background(s, WHITE)
        add_title_bar(s, f"{page_idx + 3:02}　v{vid}　{name}",
                      f"{system}｜{hex_}")

        # 大圖
        add_image_box(s, 0.5, 1.7, 8.8, 3.7, ASSETS / f"v{vid}.jpg")

        # 右側資料卡
        card_x = 9.5; card_y = 1.7
        add_rect(s, Inches(card_x), Inches(card_y), Inches(3.3), Inches(3.7),
                 fill_color=CREAM, line_color=DARK_BROWN)
        # color chip 大塊
        add_rect(s, Inches(card_x + 0.2), Inches(card_y + 0.2),
                 Inches(2.9), Inches(0.8), fill_color=hex_to_rgb(hex_), line_color=DARK_BROWN)
        # 在色塊上放 hex
        text_color = WHITE if sum(hex_to_rgb(hex_)) / 3 < 130 else DARK_BROWN
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 0.2),
                 Inches(2.9), Inches(0.8),
                 hex_, size=18, bold=True, color=text_color,
                 align="center", anchor="middle", font="Courier New")
        # 情緒詞
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 1.15),
                 Inches(2.9), Inches(0.35),
                 "情緒詞", size=11, color=GRAY_70)
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 1.5),
                 Inches(2.9), Inches(0.4),
                 "　·　".join(moods), size=14, bold=True, color=DARK_BROWN)
        # 主對應 TA
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 2.0),
                 Inches(2.9), Inches(0.35),
                 "主對應 TA", size=11, color=GRAY_70)
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 2.35),
                 Inches(2.9), Inches(0.6),
                 ta, size=12, bold=True, color=ACCENT_RED, line_space=1.3)
        # 印刷/數位
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 3.0),
                 Inches(1.4), Inches(0.3),
                 "印刷適用", size=10, color=GRAY_70)
        add_text(s, Inches(card_x + 1.6), Inches(card_y + 3.0),
                 Inches(1.7), Inches(0.3),
                 pr, size=12, color=ACCENT_RED, bold=True)
        add_text(s, Inches(card_x + 0.2), Inches(card_y + 3.3),
                 Inches(1.4), Inches(0.3),
                 "數位適用", size=10, color=GRAY_70)
        add_text(s, Inches(card_x + 1.6), Inches(card_y + 3.3),
                 Inches(1.7), Inches(0.3),
                 dg, size=12, color=ACCENT_RED, bold=True)

        # 底部評語
        add_rect(s, Inches(0.5), Inches(5.65), Inches(12.3), Inches(1.3),
                 fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
        add_text(s, Inches(0.7), Inches(5.78), Inches(12), Inches(0.4),
                 "璞域評語", size=13, bold=True, color=DARK_BROWN)
        add_text(s, Inches(0.7), Inches(6.2), Inches(11.9), Inches(0.8),
                 note, size=12.5, color=DARK_BROWN, line_space=1.4)
        add_page_number(s, 4 + page_idx, TOTAL)

    # ── 11. 8 版本 × TA 對應矩陣
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "11　8 版本 × 客群對應矩陣",
                  "縱軸 = 客群｜橫軸 = 版本｜★ 數量 = 契合度")

    matrix = [
        ["", "v1\n酒紅", "v2\n普藍", "v3\n森綠", "v4\n霧灰", "v5\n湖青", "v6\n橄欖", "v7\n金黃Rose", "v8\n金黃水仙"],
        ["TA 1 返鄉換屋", "★", "★★", "★★", "★★★★", "★★★", "★★★★", "★★★★★", "★★★★★"],
        ["TA 2 在地二代", "★", "★★", "★", "★★★★★", "★★", "★★★", "★★★★★", "★★★★★"],
        ["TA 3 退休置產", "★★★★★", "★★★★★", "★★★★★", "★★", "★★★★★", "★★★★", "★★★", "★★★"],
    ]
    headers = matrix[0]
    rows = matrix[1:]
    add_table(s, Inches(0.3), Inches(1.7), Inches(12.7), Inches(3.5),
              headers, rows, header_size=10.5, body_size=11,
              first_col_bold=True, body_align="center")

    # 結論
    add_rect(s, Inches(0.5), Inches(5.45), Inches(12.3), Inches(1.5),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.55), Inches(12), Inches(0.4),
             "矩陣解讀", size=13, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "若主打全客層 → v7 / v8（金黃系）契合度最廣",
        "若主打高客單（TA3） → v1 / v2 / v3 / v5 是最強選項",
        "若想兼顧 → 用 v7 為主視覺、v2 或 v5 為「精品系列」副視覺",
    ], size=12)

    add_page_number(s, 12, TOTAL)

    # ── 12. TOP 3 推薦
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "12　TOP 3 推薦排序", "綜合三組 TA 涵蓋度 + 通路效果")

    rec = [
        ("第一推薦", "v7　金黃·古典玫瑰",
         "全客層接受度最高｜印刷與數位皆強｜業主先前選定方向",
         BRAND_YELLOW, WHITE, 7),
        ("第二推薦", "v5　湖青·古典玫瑰",
         "東方氣質最濃｜契合主文案「在繁華中收藏寧靜」｜TA3 主力",
         CREAM, DARK_BROWN, 5),
        ("第三推薦", "v2　普魯士藍·水仙",
         "獨特性最高｜博物學插畫感｜可作精品副視覺",
         CREAM, DARK_BROWN, 2),
    ]
    box_h = 1.55; y0 = 1.85
    for i, (rank, name, msg, fill, text_color, vid) in enumerate(rec):
        y = y0 + i * (box_h + 0.1)
        add_rect(s, Inches(0.5), Inches(y), Inches(12.3), Inches(box_h),
                 fill_color=fill, line_color=DARK_BROWN)
        # 縮圖
        add_image_box(s, 0.65, y + 0.12, 2.4, 1.3, ASSETS / f"v{vid}.jpg")
        # rank + name
        add_text(s, Inches(3.3), Inches(y + 0.18), Inches(2.5), Inches(0.4),
                 rank, size=14, bold=True, color=text_color)
        add_text(s, Inches(3.3), Inches(y + 0.55), Inches(7.0), Inches(0.55),
                 name, size=20, bold=True, color=text_color, font=FONT_TITLE)
        add_text(s, Inches(3.3), Inches(y + 1.05), Inches(8.9), Inches(0.45),
                 msg, size=12, color=text_color, line_space=1.35)
    add_page_number(s, 13, TOTAL)

    # ── 13. 落地使用建議
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "13　落地使用建議",
                  "若選 v7 為主視覺 — 其他版本如何被使用")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["通路 / 載體", "建議版本", "理由"],
              [
                  ["接待會館主視覺", "v7（金黃Rose）", "戶外辨識度最高，全客層覆蓋"],
                  ["大型戶外帆布", "v7", "色彩飽和、遠距視認佳"],
                  ["DM / 售屋手冊封面", "v7", "印刷成本可控"],
                  ["591 案件首圖", "v8（金黃水仙）", "輕盈年輕，網路曝光更佳"],
                  ["Facebook 廣告投放", "v8", "明亮、適合動態縮圖"],
                  ["Instagram 帳號頭像", "v4 或 v8", "輕盈、適合社群調性"],
                  ["精品 DM / 信封", "v2（普藍）", "高端對應 TA3"],
                  ["看屋會館貴賓室", "v5（湖青）", "東方氣質、退休客接待氛圍"],
                  ["年節限定 / 開賣首發", "v1（酒紅）", "莊重、儀式感、限定感"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_footer(s, "原則：v7 為主、其他版本依場景互補 — 形成系列家族")
    add_page_number(s, 14, TOTAL)

    # ── 14. 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "14　下一步行動", "視覺方向定稿 → CI 製作")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "決策項", "璞域建議", "業主回覆"],
              [
                  ["1", "主視覺定案", "建議 v7（金黃·玫瑰）", "_______"],
                  ["2", "副視覺定案", "v8（社群）+ v5 或 v2（精品）", "_______"],
                  ["3", "字標版本", "與 v7 搭配深棕字標（已於 deck 3 提案）", "_______"],
                  ["4", "字體系統", "思源宋體 + 思源黑體（deck 3 細部）", "_______"],
                  ["5", "向量檔交付", "8 版本皆可，但僅選定 1~3 個正式量產", "_______"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "視覺方向定稿後，後續延伸 logo、CI 系統、印刷品、數位資產約 6 週可完成",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 15, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
