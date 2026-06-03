"""簡報 6：中道森活 — 定價×客群深掘分析簡報.

deck 1 的進階版。聚焦於：
  1. 三組 TA 的「persona 卡」（含人物 quote、決策權重、消費邏輯）
  2. 每組 TA 對應具體哪幾戶 + 哪個價格帶
  3. 27 戶 × TA 配置矩陣
  4. 銷售動線與時序（誰先賣、誰後賣、看屋會館引導）

輸出：02_簡報/06_中道森活_定價客群深掘簡報.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "06_中道森活_定價客群深掘簡報.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)
TOTAL = 19

# ── TA × 戶別配置 ──────────────────────────────
# Total 27 戶，依鼎弘玉田森活 1 客層比例反推：
# TA3 premium 5 戶 / TA1 12 戶 / TA2 10 戶
TA3_UNITS = ["A1", "E3", "D1", "E1", "F4"]
TA1_UNITS = ["B7", "C3", "D2", "D3", "D4", "E2", "F1", "F2", "F3", "G1", "G2", "G3"]
TA2_UNITS = ["A2", "A3", "B1", "B2", "B3", "B4", "B5", "B6", "C1", "C2"]


def persona_card(slide, x, y, w, h, *, label, color_bg, color_accent,
                 name, tagline, quote, kpis):
    """單一 persona 大卡片."""
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h),
             fill_color=color_bg, line_color=color_accent)
    # label tag
    add_rect(slide, Inches(x + 0.2), Inches(y + 0.2), Inches(1.2), Inches(0.4),
             fill_color=color_accent)
    add_text(slide, Inches(x + 0.2), Inches(y + 0.2), Inches(1.2), Inches(0.4),
             label, size=12, bold=True, color=WHITE, align="center", anchor="middle")
    # name
    add_text(slide, Inches(x + 0.3), Inches(y + 0.7), Inches(w - 0.6), Inches(0.55),
             name, size=20, bold=True, color=DARK_BROWN, font=FONT_TITLE)
    # tagline
    add_text(slide, Inches(x + 0.3), Inches(y + 1.25), Inches(w - 0.6), Inches(0.35),
             tagline, size=11, color=color_accent, bold=True)
    # quote
    add_text(slide, Inches(x + 0.3), Inches(y + 1.65), Inches(w - 0.6), Inches(0.95),
             f"“{quote}”", size=11, color=GRAY_70, line_space=1.35)
    # kpis — denser layout, smaller font
    kpi_y_start = y + 2.75
    kpi_line_h = 0.28
    for i, (k, v) in enumerate(kpis):
        ky = kpi_y_start + i * kpi_line_h
        add_text(slide, Inches(x + 0.3), Inches(ky), Inches(1.6), Inches(0.25),
                 k, size=10, color=GRAY_70)
        add_text(slide, Inches(x + 1.9), Inches(ky), Inches(w - 2.2), Inches(0.25),
                 v, size=10.5, bold=True, color=DARK_BROWN)


def build():
    prs = new_presentation()

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="定價 × 客群深掘簡報　Pricing-by-Persona Deep-Dive",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　核心理念：為什麼客群決定定價",
        "02　壯圍透天買家 — 真實客層輪廓",
        "03　三組 TA 一覽 — 從 198 筆實登推算",
        "04　TA1 深掘：宜蘭返鄉換屋族",
        "05　TA2 深掘：壯圍/礁溪在地二代",
        "06　TA3 深掘：雙北退休置產",
        "07　三組 TA × 27 戶配置矩陣",
        "08　三組 TA × 價格落點圖",
        "09　三組 TA × Message × 通路",
        "10　買家旅程地圖 — 從看到買的 7 個觸點",
        "11　上市時序：誰先賣、誰後賣",
        "12　看屋會館動線設計",
        "13　風險與三組 TA 的後備計畫",
        "14　下一步行動清單",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=18, line_space=1.4)

    # ── 01 核心理念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　核心理念",
                  "為什麼客群決定定價 — 不是反過來")

    # 兩欄對比
    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=CREAM, line_color=ACCENT_RED)
    add_text(s, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.5),
             "❌ 常見錯誤", size=18, bold=True, color=ACCENT_RED, align="center")
    add_bullets(s, Inches(0.8), Inches(2.45), Inches(5.4), Inches(4.0), [
        "先定總價、再找客人",
        "看競品定價、平均拉中位",
        "業主想賣多少 → 業主開價",
        "結果：總價對的、客人卻不來",
        "或客人來了、但你不知道他為什麼買",
    ], size=12.5, line_space=1.5, bullet_color=ACCENT_RED)

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.5),
             "✅ 客群驅動定價", size=18, bold=True, color=DARK_BROWN, align="center")
    add_bullets(s, Inches(7.2), Inches(2.45), Inches(5.4), Inches(4.0), [
        "先定客人是誰、再定總價",
        "客人預算與決策權重決定價格帶",
        "每組客人 → 對應戶別 → 對應價格",
        "結果：訂價有依據、銷售話術有目標",
        "鼎弘玉田森活1 已驗證此模式（2 年完銷）",
    ], size=12.5, line_space=1.5)

    add_footer(s, "本簡報以「客群」為主軸，反推定價、戶別配置、銷售時序")
    add_page_number(s, 2, TOTAL)

    # ── 02 壯圍買家真實輪廓
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　壯圍透天買家 — 真實客層輪廓",
                  "從 198 筆 24 個月中古成交 + 鼎弘玉田 30 筆實登交叉推算")

    metrics = [
        ("47%", "屋齡 ≤5 年 (近新成屋)", "新成屋需求最強"),
        ("53%", "建坪 40~55 坪", "4 房透天主流"),
        ("78%", "成交價 800~1,100 萬", "千萬內為甜蜜帶"),
        ("65%+", "宜蘭縣內買家", "返鄉與在地比北部置產多"),
    ]
    bw, bh = 2.9, 2.0
    for i, (big, label, sub) in enumerate(metrics):
        x = 0.5 + i * (bw + 0.13)
        add_rect(s, Inches(x), Inches(1.85), Inches(bw), Inches(bh),
                 fill_color=CREAM, line_color=DARK_BROWN)
        add_text(s, Inches(x), Inches(2.0), Inches(bw), Inches(0.9),
                 big, size=40, bold=True, color=ACCENT_RED, align="center", font=FONT_TITLE)
        add_text(s, Inches(x), Inches(2.95), Inches(bw), Inches(0.5),
                 label, size=11.5, bold=True, color=DARK_BROWN, align="center")
        add_text(s, Inches(x), Inches(3.4), Inches(bw), Inches(0.4),
                 sub, size=10.5, color=GRAY_70, align="center")

    add_rect(s, Inches(0.5), Inches(4.3), Inches(12.3), Inches(2.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.4), Inches(12), Inches(0.5),
             "推論結論", size=15, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(0.7), Inches(4.85), Inches(11.9), Inches(2.0), [
        "壯圍的買家不是雙北「找便宜」客，是宜蘭縣自己的居住升級需求",
        "本案 4 房 / 38~43 坪 / 透天 3 層 = 命中 78% 主流甜蜜帶",
        "新成屋訴求要強過鄰區中古（屋齡中位 10 年） — 「新成屋」是天然賣點",
        "返鄉、在地、退休三種動機可彼此區分，**不是同一群人**",
    ], size=12)

    add_footer(s, "資料源：實價登錄 198 筆 + 鼎弘玉田 30 筆 + 璞域市場推算（百分比為合理估算）")
    add_page_number(s, 3, TOTAL)

    # ── 03 三組 TA 一覽
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　三組 TA 一覽",
                  "從 198 筆實登推算的客群結構 — 共佔市場 100%")

    tas_meta = [
        ("TA 1", "宜蘭返鄉換屋族",  "45%",  "約 12 戶",  "30~45 歲｜換屋升級",
         BRAND_YELLOW_SOFT, DARK_BROWN),
        ("TA 2", "壯圍/礁溪在地二代", "35%",  "約 10 戶",  "35~50 歲｜成家換新",
         CREAM, DARK_BROWN),
        ("TA 3", "雙北退休置產",      "20%",  "約 5 戶",   "50+ 歲｜退休回鄉",
         BRAND_YELLOW, WHITE),
    ]
    cw, ch, y0 = 4.0, 5.0, 1.7
    for i, (label, name, pct, count, sub, bg, fg) in enumerate(tas_meta):
        x = 0.5 + i * 4.2
        add_rect(s, Inches(x), Inches(y0), Inches(cw), Inches(ch),
                 fill_color=bg, line_color=DARK_BROWN)
        add_text(s, Inches(x), Inches(y0 + 0.3), Inches(cw), Inches(0.4),
                 label, size=14, color=fg, align="center")
        add_text(s, Inches(x), Inches(y0 + 0.7), Inches(cw), Inches(0.7),
                 name, size=22, bold=True, color=fg, align="center", font=FONT_TITLE)
        # big percent
        add_text(s, Inches(x), Inches(y0 + 1.7), Inches(cw), Inches(1.3),
                 pct, size=68, bold=True, color=fg, align="center", font=FONT_TITLE)
        # count
        add_text(s, Inches(x), Inches(y0 + 3.2), Inches(cw), Inches(0.4),
                 count, size=15, bold=True, color=fg, align="center")
        # sub
        add_text(s, Inches(x), Inches(y0 + 3.75), Inches(cw), Inches(0.5),
                 sub, size=12, color=fg, align="center")

    add_text(s, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.4),
             "下三頁逐一深掘每組 TA 的 persona、決策邏輯、對應戶別與價格。",
             size=12, color=GRAY_70, align="center")
    add_page_number(s, 4, TOTAL)

    # ── 04 TA1 深掘
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　TA 1 — 宜蘭返鄉換屋族",
                  "12 戶｜45% 主力客群")

    persona_card(s, 0.5, 1.7, 5.9, 5.0,
        label="TA 1", color_bg=BRAND_YELLOW_SOFT, color_accent=DARK_BROWN,
        name="林先生・34 歲", tagline="軟體工程師・宜蘭出生・台北工作 8 年",
        quote="爸媽年紀大了，我們想搬回來。但要 4 房，孩子上小學要有自己房間。預算極限就是千萬。",
        kpis=[
            ("年齡帶", "30~45 歲"),
            ("家庭", "夫妻 + 1~2 子女"),
            ("家庭年收", "100~150 萬"),
            ("北部租屋月租", "25,000~35,000"),
            ("自備款", "200~300 萬"),
            ("預算上限", "800~950 萬"),
            ("關鍵需求", "4 房、學區、爸媽近"),
            ("決策時間", "看屋 3~6 個月"),
        ])

    # 決策權重
    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(2.4),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(1.8), Inches(5.5), Inches(0.5),
             "決策權重（重要性 1~10 分）", size=14, bold=True, color=DARK_BROWN)
    weights_1 = [("總價守千萬", 10), ("4 房格局", 9), ("子女學區/通勤", 8),
                 ("爸媽就近照顧", 8), ("建材品牌", 6), ("車位", 4)]
    for i, (item, score) in enumerate(weights_1):
        y_w = 2.3 + i * 0.27
        add_text(s, Inches(7.1), Inches(y_w), Inches(2.2), Inches(0.25),
                 item, size=10.5, color=DARK_BROWN)
        bar_w = score / 10 * 2.8
        add_rect(s, Inches(9.5), Inches(y_w + 0.05), Inches(bar_w), Inches(0.18),
                 fill_color=ACCENT_RED)
        add_text(s, Inches(9.5 + bar_w + 0.05), Inches(y_w), Inches(0.5), Inches(0.25),
                 str(score), size=10, color=DARK_BROWN)

    # 對應戶別
    add_rect(s, Inches(6.9), Inches(4.25), Inches(5.9), Inches(2.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(4.35), Inches(5.5), Inches(0.5),
             "對應戶別 → 12 戶", size=14, bold=True, color=ACCENT_RED)
    add_text(s, Inches(7.1), Inches(4.85), Inches(5.5), Inches(0.5),
             "、".join(TA1_UNITS), size=12, color=DARK_BROWN, line_space=1.4)
    add_text(s, Inches(7.1), Inches(5.55), Inches(5.5), Inches(0.4),
             "建議價格帶", size=12, bold=True, color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(5.95), Inches(5.5), Inches(0.7),
             "每坪 21.0 萬｜總價 810~903 萬",
             size=18, bold=True, color=ACCENT_RED, font=FONT_TITLE)

    add_footer(s, "TA1 是主力客群 — 看屋會館動線、廣告主圖、文案 message 應以此族群為核心")
    add_page_number(s, 5, TOTAL)

    # ── 05 TA2 深掘
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　TA 2 — 壯圍/礁溪在地二代",
                  "10 戶｜35% 次主力")

    persona_card(s, 0.5, 1.7, 5.9, 5.0,
        label="TA 2", color_bg=CREAM, color_accent=DARK_BROWN,
        name="陳小姐・41 歲", tagline="羅東護理師・壯圍土生土長",
        quote="老家公寓住到滿、孩子要房間。看了五結兩個案，建商都不認識。中道森活我聽過鼎弘玉田。",
        kpis=[
            ("年齡帶", "35~50 歲"),
            ("家庭", "夫妻 + 1~2 子女"),
            ("家庭年收", "80~130 萬"),
            ("現居", "公寓或老透天"),
            ("自備款", "150~250 萬"),
            ("預算上限", "850~1,000 萬"),
            ("關鍵需求", "通勤、停車、生活機能"),
            ("決策時間", "看屋 2~4 個月"),
        ])

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(2.4),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(1.8), Inches(5.5), Inches(0.5),
             "決策權重（重要性 1~10 分）", size=14, bold=True, color=DARK_BROWN)
    weights_2 = [("通勤距離（職場）", 9), ("總價 < 1000 萬", 9), ("自家停車", 8),
                 ("品牌信賴", 7), ("學區", 6), ("建材", 5)]
    for i, (item, score) in enumerate(weights_2):
        y_w = 2.3 + i * 0.27
        add_text(s, Inches(7.1), Inches(y_w), Inches(2.2), Inches(0.25),
                 item, size=10.5, color=DARK_BROWN)
        bar_w = score / 10 * 2.8
        add_rect(s, Inches(9.5), Inches(y_w + 0.05), Inches(bar_w), Inches(0.18),
                 fill_color=ACCENT_RED)
        add_text(s, Inches(9.5 + bar_w + 0.05), Inches(y_w), Inches(0.5), Inches(0.25),
                 str(score), size=10, color=DARK_BROWN)

    add_rect(s, Inches(6.9), Inches(4.25), Inches(5.9), Inches(2.5),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(4.35), Inches(5.5), Inches(0.5),
             "對應戶別 → 10 戶", size=14, bold=True, color=ACCENT_RED)
    add_text(s, Inches(7.1), Inches(4.85), Inches(5.5), Inches(0.5),
             "、".join(TA2_UNITS), size=12, color=DARK_BROWN, line_space=1.4)
    add_text(s, Inches(7.1), Inches(5.55), Inches(5.5), Inches(0.4),
             "建議價格帶", size=12, bold=True, color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(5.95), Inches(5.5), Inches(0.7),
             "每坪 20.5 萬｜總價 780~840 萬",
             size=18, bold=True, color=ACCENT_RED, font=FONT_TITLE)

    add_footer(s, "TA2 是價格最敏感的客群 — 對應 A、B、C 幢小地坪戶（10 戶）")
    add_page_number(s, 6, TOTAL)

    # ── 06 TA3 深掘
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　TA 3 — 雙北退休置產",
                  "5 戶｜20% 高客單價輔助")

    persona_card(s, 0.5, 1.7, 5.9, 5.0,
        label="TA 3", color_bg=BRAND_YELLOW, color_accent=DARK_BROWN,
        name="王先生・58 歲", tagline="北部退休醫師・宜蘭老家",
        quote="孩子都大了，老婆愛宜蘭。想要個自己的透天，安靜、有院子，不在乎多花 100 萬。",
        kpis=[
            ("年齡帶", "50+ 歲"),
            ("家庭", "夫妻（子女已成年）"),
            ("家庭資產", "雄厚"),
            ("付款方式", "多為自備、低貸款"),
            ("自備款", "500 萬以上"),
            ("預算上限", "900~1,000 萬"),
            ("關鍵需求", "寧靜、面寬、不要太密"),
            ("決策時間", "看屋 6~12 個月"),
        ])

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(2.4),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(1.8), Inches(5.5), Inches(0.5),
             "決策權重（重要性 1~10 分）", size=14, bold=True, color=DARK_BROWN)
    weights_3 = [("面寬感/採光", 10), ("社區密度低", 9), ("品牌信譽", 8),
                 ("總價彈性", 5), ("車位", 7), ("通勤", 3)]
    for i, (item, score) in enumerate(weights_3):
        y_w = 2.3 + i * 0.27
        add_text(s, Inches(7.1), Inches(y_w), Inches(2.2), Inches(0.25),
                 item, size=10.5, color=DARK_BROWN)
        bar_w = score / 10 * 2.8
        add_rect(s, Inches(9.5), Inches(y_w + 0.05), Inches(bar_w), Inches(0.18),
                 fill_color=ACCENT_RED)
        add_text(s, Inches(9.5 + bar_w + 0.05), Inches(y_w), Inches(0.5), Inches(0.25),
                 str(score), size=10, color=DARK_BROWN)

    add_rect(s, Inches(6.9), Inches(4.25), Inches(5.9), Inches(2.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(4.35), Inches(5.5), Inches(0.5),
             "對應戶別 → 5 戶（大地坪／角間）", size=14, bold=True, color=ACCENT_RED)
    add_text(s, Inches(7.1), Inches(4.85), Inches(5.5), Inches(0.5),
             "、".join(TA3_UNITS), size=12, color=DARK_BROWN, line_space=1.4)
    add_text(s, Inches(7.1), Inches(5.55), Inches(5.5), Inches(0.4),
             "建議價格帶（溢價 +3~5%）", size=12, bold=True, color=DARK_BROWN)
    add_text(s, Inches(7.1), Inches(5.95), Inches(5.5), Inches(0.7),
             "每坪 22.0 萬｜總價 920~990 萬",
             size=18, bold=True, color=ACCENT_RED, font=FONT_TITLE)

    add_footer(s, "TA3 客單價最高 — 5 戶溢價戶為全案毛利最高來源")
    add_page_number(s, 7, TOTAL)

    # ── 07 TA × 戶別配置矩陣
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　TA × 27 戶配置矩陣",
                  "每一戶都有明確指定客群與目標價格")

    add_text(s, Inches(0.5), Inches(1.65), Inches(12.3), Inches(0.4),
             "27 戶配置（顏色 = 客群歸屬）", size=13, bold=True, color=DARK_BROWN)

    # 戶別顏色
    color_map = {
        "TA1": (BRAND_YELLOW_SOFT, DARK_BROWN),
        "TA2": (CREAM, DARK_BROWN),
        "TA3": (BRAND_YELLOW, DARK_BROWN),
    }
    unit_to_ta = {}
    for u in TA1_UNITS: unit_to_ta[u] = "TA1"
    for u in TA2_UNITS: unit_to_ta[u] = "TA2"
    for u in TA3_UNITS: unit_to_ta[u] = "TA3"

    # 排成 A~G 7 列，每列若干戶
    rows_by_block = {
        "A 幢": ["A1", "A2", "A3"],
        "B 幢": ["B1", "B2", "B3", "B4", "B5", "B6", "B7"],
        "C 幢": ["C1", "C2", "C3"],
        "D 幢": ["D1", "D2", "D3", "D4"],
        "E 幢": ["E1", "E2", "E3"],
        "F 幢": ["F1", "F2", "F3", "F4"],
        "G 幢": ["G1", "G2", "G3"],
    }
    box_w = 0.85; box_h = 0.55
    x_left = 1.4; y_top = 2.1
    for i, (block, units) in enumerate(rows_by_block.items()):
        y_b = y_top + i * 0.65
        add_text(s, Inches(0.5), Inches(y_b + 0.1), Inches(0.8), Inches(0.4),
                 block, size=13, bold=True, color=DARK_BROWN)
        for j, u in enumerate(units):
            ta = unit_to_ta[u]
            bg, fg = color_map[ta]
            x_u = x_left + j * (box_w + 0.05)
            add_rect(s, Inches(x_u), Inches(y_b), Inches(box_w), Inches(box_h),
                     fill_color=bg, line_color=DARK_BROWN)
            add_text(s, Inches(x_u), Inches(y_b), Inches(box_w), Inches(box_h),
                     u, size=12, bold=True, color=fg, align="center", anchor="middle")

    # 圖例
    legend_y = 6.6
    legends = [("TA 1 返鄉換屋｜12 戶｜21.0 萬", BRAND_YELLOW_SOFT, 0.5),
               ("TA 2 在地二代｜10 戶｜20.5 萬", CREAM, 5.0),
               ("TA 3 退休置產｜5 戶｜22.0 萬", BRAND_YELLOW, 9.5)]
    for label, color, x in legends:
        add_rect(s, Inches(x), Inches(legend_y), Inches(0.4), Inches(0.3),
                 fill_color=color, line_color=DARK_BROWN)
        add_text(s, Inches(x + 0.5), Inches(legend_y - 0.05), Inches(3.8), Inches(0.4),
                 label, size=12, color=DARK_BROWN, anchor="middle")

    add_footer(s, "配置邏輯：地坪 ≥88m² 或角間 → TA3｜小地坪 → TA2｜中型主流 → TA1")
    add_page_number(s, 8, TOTAL)

    # ── 08 TA 價格落點圖
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　TA × 價格落點圖",
                  "三組 TA 對應三條清晰的價格帶")

    # 水平價格軸（780~990 萬）
    x_left = 1.5; x_right = 12.5; y_axis = 5.6
    # 軸線
    add_rect(s, Inches(x_left), Inches(y_axis), Inches(x_right - x_left), Pt(2),
             fill_color=DARK_BROWN)
    # tick marks
    tick_values = [780, 820, 860, 900, 940, 980]
    for tv in tick_values:
        pos = x_left + (tv - 770) / (1000 - 770) * (x_right - x_left)
        add_rect(s, Inches(pos), Inches(y_axis), Pt(1), Inches(0.08), fill_color=DARK_BROWN)
        add_text(s, Inches(pos - 0.3), Inches(y_axis + 0.15), Inches(0.6), Inches(0.3),
                 f"{tv}", size=10, color=GRAY_70, align="center")
    add_text(s, Inches(x_left - 0.7), Inches(y_axis - 0.1), Inches(0.7), Inches(0.3),
             "總價", size=11, color=GRAY_70)
    add_text(s, Inches(x_right + 0.05), Inches(y_axis - 0.1), Inches(0.7), Inches(0.3),
             "（萬）", size=11, color=GRAY_70)

    # 三條條狀帶
    bands = [
        ("TA 2 在地二代",   780, 840,  20.5, "10 戶｜A2 A3 B1~B6 C1 C2", CREAM, BRAND_YELLOW_SOFT),
        ("TA 1 返鄉換屋",   810, 903,  21.0, "12 戶｜B7 C3 D2~D4 E2 F1~F3 G1~G3", BRAND_YELLOW_SOFT, BRAND_YELLOW),
        ("TA 3 退休置產",   920, 990,  22.0, "5 戶｜A1 D1 E1 E3 F4", BRAND_YELLOW, DARK_BROWN),
    ]
    band_h = 0.85
    for i, (name, lo, hi, ppk, units, fill, line) in enumerate(bands):
        y_b = 1.85 + i * 1.05
        x_lo = x_left + (lo - 770) / (1000 - 770) * (x_right - x_left)
        x_hi = x_left + (hi - 770) / (1000 - 770) * (x_right - x_left)
        add_rect(s, Inches(x_lo), Inches(y_b), Inches(x_hi - x_lo), Inches(band_h),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(x_lo + 0.15), Inches(y_b + 0.07), Inches(x_hi - x_lo - 0.3), Inches(0.4),
                 name, size=13, bold=True, color=text_color)
        add_text(s, Inches(x_lo + 0.15), Inches(y_b + 0.45), Inches(x_hi - x_lo - 0.3), Inches(0.35),
                 f"{lo}~{hi} 萬｜{ppk} 萬/坪", size=11, color=text_color)
        # 戶數標籤在右側
        add_text(s, Inches(0.5), Inches(y_b + 0.25), Inches(0.9), Inches(0.4),
                 units.split("｜")[0], size=10, color=GRAY_70, align="right", anchor="middle")

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.4),
             "三條價格帶清楚分隔｜TA1 和 TA2 略有重疊，可作為「跨族群選擇」緩衝",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_footer(s, "全案 27 戶皆守在 1,000 萬內")
    add_page_number(s, 9, TOTAL)

    # ── 09 TA × Message × 通路
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　TA × Message × 通路",
                  "每組 TA 對應的主打話術與媒體投放")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.7),
              ["TA", "主打 message", "主視覺角度", "媒體通路", "預估預算佔比"],
              [
                  ["TA 1 返鄉換屋",
                   "「回宜蘭，4 房一次到位」",
                   "親子 + 庭院 + 早晨光",
                   "Facebook 雙北精準｜591 列表｜LINE 群組",
                   "50%"],
                  ["TA 2 在地二代",
                   "「不到千萬，住進新成屋透天」",
                   "夫妻 + 新家鑰匙 + 房間",
                   "宜蘭地方廣告｜LINE｜口碑｜社群媒體",
                   "30%"],
                  ["TA 3 退休置產",
                   "「在繁華之中，找到一個栖息」",
                   "庭院 + 花卉 + 寧靜清晨",
                   "雙北高消費精準｜591 高總價｜實體 DM",
                   "20%"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")

    add_text(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5),
             "媒體預算分配建議：50% 主力（TA1）／ 30% 在地（TA2）／ 20% 高客單（TA3）",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_footer(s, "主視覺已選用花卉典雅插畫 — 同視覺可同時打 TA1（家的溫度）+ TA3（寧靜）")
    add_page_number(s, 10, TOTAL)

    # ── 10 買家旅程
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "10　買家旅程地圖",
                  "從看到、到買、到推薦 — 7 個關鍵觸點")

    journey = [
        ("觸發",    "看到廣告／聽到口碑", BRAND_YELLOW_SOFT),
        ("認識",    "上 591／網站查資料", CREAM),
        ("考慮",    "加入 LINE｜接電", BRAND_YELLOW_SOFT),
        ("看屋",    "預約看屋會館",       CREAM),
        ("比較",    "看 2~3 個競品",     BRAND_YELLOW_SOFT),
        ("決策",    "回看屋會館下訂",     BRAND_YELLOW),
        ("交屋",    "完工驗收→口碑推薦",  CREAM),
    ]
    bw = 1.65; gap = 0.1; y_b = 2.0
    for i, (stage, desc, fill) in enumerate(journey):
        x = 0.5 + i * (bw + gap)
        # circle/box
        add_rect(s, Inches(x), Inches(y_b), Inches(bw), Inches(2.2),
                 fill_color=fill, line_color=DARK_BROWN)
        # step number
        add_text(s, Inches(x), Inches(y_b + 0.15), Inches(bw), Inches(0.4),
                 f"Step {i+1}", size=11, color=ACCENT_RED, align="center", bold=True)
        # stage name
        add_text(s, Inches(x), Inches(y_b + 0.6), Inches(bw), Inches(0.55),
                 stage, size=18, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
        # desc
        add_text(s, Inches(x), Inches(y_b + 1.25), Inches(bw), Inches(0.85),
                 desc, size=10.5, color=DARK_BROWN, align="center", line_space=1.3)

    # 每階段對應的關鍵動作
    add_text(s, Inches(0.5), Inches(4.5), Inches(12.3), Inches(0.4),
             "每階段三組 TA 的差異化策略", size=14, bold=True, color=DARK_BROWN)
    add_table(s, Inches(0.5), Inches(4.95), Inches(12.3), Inches(1.95),
              ["階段", "TA 1 返鄉", "TA 2 在地", "TA 3 退休"],
              [
                  ["看到", "FB 廣告", "LINE 群組／親友", "DM／實體看板"],
                  ["看屋", "週末攜家帶眷", "下班後快速看", "預約專人接待"],
                  ["決策", "夫妻共同決定", "夫妻 + 父母", "夫妻"],
              ],
              header_size=11, body_size=11, first_col_bold=True, body_align="center")
    add_footer(s, "看屋會館應準備：兒童區（TA1）｜在地茶水（TA2）｜貴賓室（TA3）")
    add_page_number(s, 11, TOTAL)

    # ── 11 上市時序
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "11　上市時序：誰先賣、誰後賣",
                  "不同 TA 的去化節奏不同 — 時序排列影響全案毛利")

    phases = [
        ("Phase 1\n暖身期",  "Q3 2026",   "5 戶（TA2 為主）",
         "壓低單價衝量、建立社區能見度", CREAM),
        ("Phase 2\n主力期",  "Q4 2026~\nQ1 2027",
         "12 戶（TA1 為主）", "主力廣告、媒體曝光、TA1 集中收割", BRAND_YELLOW_SOFT),
        ("Phase 3\n精品期",  "Q2~Q3 2027", "5 戶（TA3 為主）",
         "保留好戶位、品牌溢價、單客深度服務", BRAND_YELLOW),
        ("Phase 4\n收尾期",  "Q4 2027",    "5 戶（混合）",
         "口碑推薦、彈性議價、最後清盤", CREAM),
    ]
    bw, bh = 2.9, 4.0; y0 = 1.85
    for i, (label, when, target, action, fill) in enumerate(phases):
        x = 0.5 + i * (bw + 0.15)
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(x), Inches(y0 + 0.15), Inches(bw), Inches(0.85),
                 label, size=16, bold=True, color=text_color,
                 align="center", line_space=1.2, font=FONT_TITLE)
        add_text(s, Inches(x), Inches(y0 + 1.05), Inches(bw), Inches(0.6),
                 when, size=12, color=ACCENT_RED, align="center", bold=True, line_space=1.2)
        add_text(s, Inches(x), Inches(y0 + 1.75), Inches(bw), Inches(0.5),
                 target, size=12, color=text_color, align="center", bold=True)
        add_text(s, Inches(x + 0.2), Inches(y0 + 2.35), Inches(bw - 0.4), Inches(1.6),
                 action, size=10.5, color=text_color, align="center", line_space=1.4)

    add_text(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.5),
             "全案 27 戶 / 預估 4 階段 18~24 個月完銷",
             size=14, bold=True, color=DARK_BROWN, align="center")
    add_text(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.4),
             "Phase 3 / Phase 4 的溢價戶可拉高全案毛利 3~5%",
             size=12, color=GRAY_70, align="center")
    add_page_number(s, 12, TOTAL)

    # ── 12 看屋會館動線
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "12　看屋會館動線設計",
                  "依三組 TA 設計三條看屋路徑")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["區域", "TA 1 路徑", "TA 2 路徑", "TA 3 路徑"],
              [
                  ["入口", "親子歡迎牌、孩子畫具",
                   "鄰里口碑牆", "貴賓接待員導引"],
                  ["展示中心", "全案模型 + 4 房模型屋",
                   "千萬內定價對照表", "材料樣板 + 庭院設計圖"],
                  ["實品屋體驗", "4 房格局＋兒童房示範",
                   "標準戶＋停車示範", "大地坪戶＋庭院實景"],
                  ["回到接待", "教育主題（學區、通勤地圖）",
                   "工作機能（通勤時間、生活圈）",
                   "退休生活提案（讀書、種花、家庭聚會）"],
                  ["結束", "親子合照 + 紀念品",
                   "本案 + 鼎弘玉田森活 1 完銷說明",
                   "保密議價、來日再訪"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_footer(s, "三條路徑可共用空間 — 接待員依客層判斷帶往不同區域")
    add_page_number(s, 13, TOTAL)

    # ── 13 風險與後備
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "13　風險與三組 TA 的後備計畫",
                  "若任一組 TA 不如預期，如何補位")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["風險情境", "影響戶數", "判斷指標", "補位策略"],
              [
                  ["TA1 返鄉客來客數低",
                   "12 戶",
                   "Q4 2026 看屋會館 TA1 占比 < 30%",
                   "加碼 FB 雙北精準｜降單價 0.5 萬｜延長銷售期"],
                  ["TA2 在地客觀望（等降價）",
                   "10 戶",
                   "Q1 2027 在地客成交率 < 預期",
                   "啟動「在地優惠」package｜LINE 限時讓利"],
                  ["TA3 退休客遲不出手",
                   "5 戶",
                   "Q2 2027 大坪數戶未售出",
                   "改為精選釋出｜或拉大坪數戶溢價 +5%｜先賣再給優惠"],
                  ["三組 TA 都疲軟（最壞情境）",
                   "27 戶",
                   "Phase 2 完銷率 < 40%",
                   "全案重新定位｜或啟動「鼎弘 + 宗硯雙案聯賣」"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "三組 TA 之間有「跨族群選擇」緩衝 — TA1 戶可下移給 TA2，TA3 戶可下移給 TA1",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_page_number(s, 14, TOTAL)

    # ── 14 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "14　下一步行動清單", "兩週內可決策的 6 項")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.5),
              ["#", "決策項", "璞域建議", "決策者"],
              [
                  ["1", "三組 TA 比例（45%/35%/20%）",
                   "依本簡報配置，玉田1 客層佐證", "鼎弘 + 璞域"],
                  ["2", "27 戶配置矩陣定稿",
                   "本簡報 § 7 矩陣，業主可微調", "鼎弘"],
                  ["3", "三條價格帶（20.5/21/22 萬）",
                   "本簡報 § 8 價格落點", "鼎弘"],
                  ["4", "媒體預算分配（50/30/20）",
                   "TA1 主力、媒體向北部精準投放", "鼎弘 + 璞域"],
                  ["5", "看屋會館三路徑設計",
                   "本簡報 § 12 動線", "璞域 + 接待中心"],
                  ["6", "Phase 1 暖身期啟動時點",
                   "建議 Q3 2026", "鼎弘"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "本簡報所有客群假設將由「宗硯輔銷資料請求清單」(05 簡報) 進一步驗證",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 15, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
