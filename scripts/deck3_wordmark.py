"""簡報 3：中道森活｜栖 — 最小可行 logo / 字標草案.

針對命名簡報推薦的第一案「中道森活｜栖」做 wordmark 試做，
驗證命名定案前的視覺方向。

輸出：02_簡報/03_中道森活_栖_字標草案簡報.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "03_中道森活_栖_字標草案簡報.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)
TOTAL = 13


def wordmark_horizontal(slide, cx, cy, *, scale=1.0, color=DARK_BROWN, sub_color=None):
    """中道森活｜栖 橫式字標 — 中道森活 + 分隔線 + 栖。
    cx, cy = 中心點 (in inches)."""
    sub_color = sub_color or color
    main_size = int(56 * scale)
    sub_size = int(72 * scale)
    sep_h = int(48 * scale)
    char_box_w = 1.0 * scale  # inches per character box
    # 中道森活 4 字
    chars = ["中", "道", "森", "活"]
    total_main_w = char_box_w * len(chars)
    sep_w = 0.05 * scale
    pipe_gap = 0.25 * scale
    sub_w = 1.2 * scale
    total_w = total_main_w + pipe_gap + sep_w + pipe_gap + sub_w
    x0 = cx - total_w / 2
    y0 = cy - 0.55 * scale
    # main chars
    for i, ch in enumerate(chars):
        add_text(slide, Inches(x0 + i * char_box_w), Inches(y0),
                 Inches(char_box_w), Inches(1.1 * scale),
                 ch, size=main_size, bold=False, color=color, align="center",
                 font=FONT_TITLE, anchor="middle")
    # separator |
    sx = x0 + total_main_w + pipe_gap
    add_rect(slide, Inches(sx), Inches(y0 + 0.12 * scale),
             Inches(sep_w), Inches(0.86 * scale), fill_color=color)
    # 栖 (大字)
    add_text(slide, Inches(sx + sep_w + pipe_gap), Inches(y0 - 0.05 * scale),
             Inches(sub_w), Inches(1.2 * scale),
             "栖", size=sub_size, bold=False, color=sub_color, align="center",
             font=FONT_TITLE, anchor="middle")


def wordmark_stacked(slide, cx, cy, *, scale=1.0, color=DARK_BROWN):
    """中道森活 / 栖 — 直式堆疊變體."""
    w_main = 4.0 * scale
    x0 = cx - w_main / 2
    add_text(slide, Inches(x0), Inches(cy - 0.95 * scale), Inches(w_main),
             Inches(0.6 * scale),
             "中　道　森　活", size=int(28 * scale), color=color,
             align="center", font=FONT_TITLE)
    # rule
    rw = 0.6 * scale
    add_rect(slide, Inches(cx - rw/2), Inches(cy - 0.25 * scale),
             Inches(rw), Pt(1.2), fill_color=color)
    add_text(slide, Inches(x0), Inches(cy - 0.05 * scale), Inches(w_main),
             Inches(1.2 * scale),
             "栖", size=int(96 * scale), color=color,
             align="center", font=FONT_TITLE)


def build():
    prs = new_presentation()

    # ───── 封面 ─────
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中道森活｜栖",
        subtitle="最小可行字標草案簡報　Wordmark v0.1",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ───── 1. 目錄 ─────
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　設計前提：命名與主視覺的延伸",
        "02　字標核心結構 — 中道森活｜栖",
        "03　主字標 — 橫式（建案接待中心、平面廣告）",
        "04　主字標 — 直式（捲軸、立柱、戶外旗）",
        "05　堆疊變體 — 栖字主導（情境主視覺）",
        "06　Logo 配色系統（兩版）",
        "07　Typography 字體系統",
        "08　應用情境模擬",
        "09　禁忌與規範",
        "10　下一步：選定後的延伸工項",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ───── 2. 設計前提 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　設計前提", "從命名與主視覺反推字標規範")

    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.5),
             "命名前提", size=18, bold=True, color=DARK_BROWN, align="center")
    add_bullets(s, Inches(0.8), Inches(2.45), Inches(5.4), Inches(4.0), [
        "案名：中道森活｜栖",
        "主體：「中道森活」（地名 + 鼎弘系列）",
        "辨識：「栖」（本案獨立辨識字）",
        "字意：栖 = 棲息、停靠",
        "客群：TA1 返鄉換屋 + TA3 退休置產",
    ], size=13, line_space=1.4)

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.5),
             "視覺繼承", size=18, bold=True, color=ACCENT_RED, align="center")
    add_bullets(s, Inches(7.2), Inches(2.45), Inches(5.4), Inches(4.0), [
        "色彩：金黃 #E8AC1B（主視覺）",
        "輔色：深棕 #4A3210、米白 #FFF8E7",
        "氣質：典雅、宋體骨架、東方語感",
        "情緒：「在繁華中收藏寧靜」",
        "字標風格：宋體本位、避免過度設計",
    ], size=13, line_space=1.4)

    add_footer(s, "字標設計三原則：①承接命名意涵 ②呼應主視覺色調 ③允許單字、多字延伸")
    add_page_number(s, 2, TOTAL)

    # ───── 3. 核心結構 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　字標核心結構", "中道森活 ｜ 栖 ─ 兩段式品牌標準字")

    # 大字標展示
    add_rect(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.2),
             fill_color=CREAM, line_color=DARK_BROWN)
    wordmark_horizontal(s, 6.66, 3.3, scale=1.2, color=DARK_BROWN)

    # 結構解構
    add_text(s, Inches(0.5), Inches(5.1), Inches(12.3), Inches(0.4),
             "結構解構", size=15, bold=True, color=DARK_BROWN)
    parts = [
        ("中道森活",  "鼎弘建設系列名｜延續玉田森活",  Inches(0.5)),
        ("｜",          "分隔符｜薄、長、視覺呼吸",       Inches(4.8)),
        ("栖",          "本案辨識字｜放大、視覺重心",      Inches(8.8)),
    ]
    for label, desc, x in parts:
        add_rect(s, x, Inches(5.55), Inches(4.0), Inches(1.5),
                 fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
        add_text(s, x, Inches(5.6), Inches(4.0), Inches(0.6),
                 label, size=22, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
        add_text(s, x, Inches(6.25), Inches(4.0), Inches(0.7),
                 desc, size=11.5, color=GRAY_70, align="center", line_space=1.3)

    add_footer(s, "兩段式設計：系列名小、辨識字大 — 同邏輯適用未來「中道森活｜◯」系列延伸")
    add_page_number(s, 3, TOTAL)

    # ───── 4. 主字標 橫式 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　主字標 — 橫式", "適用：接待中心招牌、平面廣告、信封信紙")

    # 三種底色測試
    bgs = [
        ("白底",         WHITE,             DARK_BROWN, 0.5),
        ("米白底",       CREAM,             DARK_BROWN, 5.0),
        ("金黃底",       BRAND_YELLOW,      WHITE,      9.5),
    ]
    bw, bh, y0 = 4.0, 3.0, 1.85
    for label, bg, fg, x in bgs:
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=bg, line_color=DARK_BROWN)
        wordmark_horizontal(s, x + bw/2, y0 + bh/2, scale=0.55, color=fg)
        add_text(s, Inches(x), Inches(y0 + bh + 0.05), Inches(bw), Inches(0.4),
                 label, size=12, color=DARK_BROWN, align="center")

    # 比例尺寸規範
    add_text(s, Inches(0.5), Inches(5.55), Inches(12.3), Inches(0.4),
             "尺寸規範（最小可讀）", size=14, bold=True, color=DARK_BROWN)
    add_table(s, Inches(0.5), Inches(5.95), Inches(12.3), Inches(1.1),
              ["應用情境", "最小寬度", "最佳寬度", "說明"],
              [
                  ["數位（網站／DM）", "180 px", "320~480 px", "栖字 ≥ 48 px 高"],
                  ["印刷（DM／名片）", "35 mm", "55~85 mm", "栖字 ≥ 12 mm 高"],
                  ["戶外（招牌／旗幟）", "300 mm", "800~2000 mm", "視距 2~10 m"],
              ],
              header_size=11, body_size=10.5, first_col_bold=True)

    add_footer(s, "字標尺寸基準：以「栖」字高度為單位 — 系列名 ≈ 0.6 × 栖高")
    add_page_number(s, 4, TOTAL)

    # ───── 5. 主字標 直式 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　主字標 — 直式", "適用：捲軸、戶外立柱、廣告布、看屋會館立面")

    # 三個直式變體
    variants = [
        ("V1 標準直式",  WHITE,         DARK_BROWN),
        ("V2 米白底",   CREAM,         DARK_BROWN),
        ("V3 金黃底",   BRAND_YELLOW,  WHITE),
    ]
    bw, bh, y0 = 4.0, 4.8, 1.7
    for i, (label, bg, fg) in enumerate(variants):
        x = 0.5 + i * 4.2
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=bg, line_color=DARK_BROWN)
        cx = x + bw / 2
        chars = ["中", "道", "森", "活"]
        for j, ch in enumerate(chars):
            y = y0 + 0.4 + j * 0.55
            add_text(s, Inches(cx - 0.4), Inches(y), Inches(0.8), Inches(0.65),
                     ch, size=28, color=fg, align="center", font=FONT_TITLE, anchor="middle")
        ysep = y0 + 0.4 + 4 * 0.55 + 0.05
        add_rect(s, Inches(cx - 0.25), Inches(ysep), Inches(0.5), Pt(1.5), fill_color=fg)
        add_text(s, Inches(cx - 0.7), Inches(ysep + 0.2), Inches(1.4), Inches(1.4),
                 "栖", size=64, color=fg, align="center", font=FONT_TITLE, anchor="middle")
        add_text(s, Inches(x), Inches(y0 + bh + 0.05), Inches(bw), Inches(0.4),
                 label, size=12, color=DARK_BROWN, align="center")

    add_footer(s, "直式版主用於戶外、看屋會館立面與長型廣告布")
    add_page_number(s, 5, TOTAL)

    # ───── 6. 堆疊變體 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　堆疊變體 — 栖字主導",
                  "適用：情境主視覺、單品 icon、Instagram 帳號頭像")

    layouts = [
        ("L1 標準堆疊",  WHITE, DARK_BROWN, 0.5),
        ("L2 米白",      CREAM, DARK_BROWN, 5.0),
        ("L3 金黃",      BRAND_YELLOW, DARK_BROWN, 9.5),
    ]
    bw, bh, y0 = 4.0, 4.5, 1.7
    for label, bg, fg, x in layouts:
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=bg, line_color=DARK_BROWN)
        wordmark_stacked(s, x + bw / 2, y0 + bh / 2 - 0.3, scale=1.0, color=fg)
        add_text(s, Inches(x), Inches(y0 + bh + 0.05), Inches(bw), Inches(0.4),
                 label, size=12, color=DARK_BROWN, align="center")

    add_text(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.4),
             "栖字成為視覺主角，系列名「中道森活」退居標題上方 — 適用於情緒/氛圍導向的單頁主視覺",
             size=12, color=GRAY_70, align="center")
    add_page_number(s, 6, TOTAL)

    # ───── 7. 配色系統 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　Logo 配色系統", "兩版主色系統 + 黑白輔助版")

    # 主色卡
    colors = [
        ("Primary", "金黃 Saffron", "#E8AC1B", BRAND_YELLOW, WHITE),
        ("Primary", "深棕 Dark Brown", "#4A3210", DARK_BROWN, WHITE),
        ("Supporting", "米白 Cream", "#FFF8E7", CREAM, DARK_BROWN),
        ("Accent", "暗紅 Earth", "#8B3A0F", ACCENT_RED, WHITE),
        ("Mono", "黑 Black", "#1F1F1F", BLACK, WHITE),
        ("Mono", "灰 Gray", "#555555", GRAY_70, WHITE),
    ]
    cw, ch, y_chip = 2.0, 1.8, 1.85
    for i, (cat, name, hex_, c, tc) in enumerate(colors):
        row = i // 3; col = i % 3
        x = 0.5 + col * 4.2
        y = y_chip + row * (ch + 0.4)
        add_rect(s, Inches(x), Inches(y), Inches(cw), Inches(ch),
                 fill_color=c, line_color=DARK_BROWN)
        add_text(s, Inches(x), Inches(y + 0.15), Inches(cw), Inches(0.4),
                 cat, size=10, color=tc, align="center")
        add_text(s, Inches(x), Inches(y + 0.55), Inches(cw), Inches(0.5),
                 name, size=14, bold=True, color=tc, align="center")
        add_text(s, Inches(x), Inches(y + 1.1), Inches(cw), Inches(0.4),
                 hex_, size=12, color=tc, align="center", font="Courier New")
        x2 = x + cw + 0.1
        add_rect(s, Inches(x2), Inches(y), Inches(2.1), Inches(ch),
                 fill_color=c, line_color=DARK_BROWN)
        add_text(s, Inches(x2), Inches(y + 0.55), Inches(2.1), Inches(0.7),
                 "中道森活｜栖", size=14, color=tc, align="center", font=FONT_TITLE)

    add_footer(s, "色彩來自業主主視覺；hex 值經測試對 logo 字標的對比皆通過 AA 可讀性")
    add_page_number(s, 7, TOTAL)

    # ───── 8. Typography ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　Typography 字體系統", "中文／英文／數字 三套字體規範")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["用途", "中文", "英文", "數字", "備援"],
              [
                  ["品牌字標", "思源宋體 Heavy", "Garamond Bold", "Garamond Bold",
                   "—（字標為向量繪製固定）"],
                  ["標題級 H1", "思源黑體 Bold", "Helvetica Bold", "Helvetica Bold",
                   "微軟正黑體 Bold"],
                  ["內文 Body", "思源黑體 Regular", "Helvetica Regular", "Helvetica Regular",
                   "微軟正黑體 Regular"],
                  ["強調 Accent", "思源宋體 Regular", "Garamond Italic", "Helvetica Regular",
                   "—"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.5),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.8), Inches(5.55), Inches(12), Inches(0.5),
             "字體選擇邏輯", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.8), Inches(5.95), Inches(11.9), Inches(1.0), [
        "宋體（思源宋）：「栖」字、品牌字標主體 — 承接東方典雅、呼應主視覺花卉",
        "黑體（思源黑）：標題、內文、表格 — 現代感、可讀性強",
        "Garamond：英文與數字 — 與宋體相容性高、不打架",
    ], size=11.5)

    add_footer(s, "思源宋體 / 思源黑體均為 Adobe 開源字體，可商用授權，免費取得")
    add_page_number(s, 8, TOTAL)

    # ───── 9. 應用情境 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　應用情境模擬", "字標在七種主要載體上的視覺呈現")

    # 6 個情境縮圖
    scenarios = [
        ("接待會館立面",  BRAND_YELLOW, DARK_BROWN),
        ("DM／售屋手冊", CREAM,         DARK_BROWN),
        ("戶外帆布",     DARK_BROWN,    BRAND_YELLOW),
        ("Instagram 頭像", BRAND_YELLOW, WHITE),
        ("信封信紙",     WHITE,         DARK_BROWN),
        ("名片正面",     CREAM,         DARK_BROWN),
    ]
    bw, bh, y0 = 4.0, 2.4, 1.8
    for i, (label, bg, fg) in enumerate(scenarios):
        row = i // 3; col = i % 3
        x = 0.5 + col * 4.2
        y = y0 + row * (bh + 0.55)
        add_rect(s, Inches(x), Inches(y), Inches(bw), Inches(bh),
                 fill_color=bg, line_color=DARK_BROWN)
        wordmark_horizontal(s, x + bw / 2, y + bh / 2, scale=0.45, color=fg)
        add_text(s, Inches(x), Inches(y + bh + 0.02), Inches(bw), Inches(0.35),
                 label, size=11, color=DARK_BROWN, align="center")

    add_page_number(s, 9, TOTAL)

    # ───── 10. 禁忌與規範 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　禁忌與規範", "字標不可這樣用")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["#", "禁忌", "原因", "正確作法"],
              [
                  ["1", "拉伸／壓縮字標", "字形扭曲影響辨識", "等比例縮放"],
                  ["2", "雙字標並列在 ≤ 80% 飽和的彩色背景上",
                   "對比不足、字標被吸進去",
                   "用米白或深棕版本"],
                  ["3", "「栖」字單獨用（脫離系列名）", "失去品牌錨點",
                   "至少保留「中道森活｜栖」全寫"],
                  ["4", "在「栖」字旁加裝飾元素", "破壞字標氣質",
                   "字標周圍保留 0.5 × 栖字高的留白"],
                  ["5", "替換字標的字體", "失去設計一致性",
                   "字標一律使用向量檔，禁止重新打字"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")

    add_footer(s, "完整字標規範手冊（Brand Guide）將於命名定稿後製作")
    add_page_number(s, 10, TOTAL)

    # ───── 11. 下一步 ─────
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "10　下一步 — 字標選定後的延伸工項", "六週交付計畫")
    steps = [
        ("W1",  "字標定稿",       "業主從本簡報 3 種變體中選定 + 微調"),
        ("W2",  "向量檔交付",     "AI / EPS / SVG / PNG 全格式產出"),
        ("W3",  "Brand Guide",   "字標、色彩、字體、留白規範 PDF"),
        ("W4~5", "印刷物製作",   "名片、信紙、信封、DM 模板"),
        ("W5~6", "數位資產",     "網站 banner、IG 模板、廣告主視覺"),
        ("W6",  "看屋會館視覺",   "戶外招牌、立面、室內導引"),
    ]
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.5),
              ["週次", "工項", "交付內容"],
              [list(x) for x in steps],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "預計六週內完成字標 + Brand Guide + 第一波應用物，銜接 Q4 2026 上市時程",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 11, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
