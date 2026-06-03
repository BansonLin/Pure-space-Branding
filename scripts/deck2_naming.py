"""簡報 2：中道森活 — 建案命名策略提案 (給鼎弘建設用).

依據簡報 1 客群定位 (TA1 返鄉換屋 / TA2 在地二代 / TA3 退休置產)，
從業主初步方向「中道森活」延伸三條命名策略路線，給出推薦 TOP 3。

輸出：02_簡報/02_中道森活_命名提案簡報.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "02_中道森活_命名提案簡報.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()
    TOTAL = 12

    # ───── 封面 ─────
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="建案命名策略提案",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ───── 1. 目錄 ─────
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　命名前提：業主方向與客群定位",
        "02　「中道」二字的雙重內涵",
        "03　「森活」二字的品牌資產",
        "04　三條命名策略路線總覽",
        "05　路線 A — 純延續「中道森活」",
        "06　路線 B — 副標延伸「中道森活｜◯」",
        "07　路線 C — 系列編號「中道森活 NO.x」",
        "08　推薦 TOP 3 與 message 對應",
        "09　主視覺與命名的搭配建議",
        "10　商標／域名／法規檢查清單",
        "11　下一步推進計畫",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ───── 2. 命名前提 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　命名前提", "從業主方向 + 客群定位反推命名標準")

    # 左：業主方向 / 右：客群定位回顧
    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.5),
             "業主初步方向", size=18, bold=True, color=DARK_BROWN, align="center")
    add_bullets(s, Inches(0.8), Inches(2.45), Inches(5.4), Inches(4.0), [
        "案名：「中道森活」",
        "延續「玉田森活」品牌系列",
        "命名邏輯：地名 + 森活",
        "主視覺：黃底花卉典雅插畫",
        "主文案：「在繁華之中收藏寧靜，在日常之間遇見家的溫度」",
    ], size=13, line_space=1.4)

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.5),
             "客群定位（取自簡報 1）", size=18, bold=True, color=ACCENT_RED, align="center")
    add_bullets(s, Inches(7.2), Inches(2.45), Inches(5.4), Inches(4.0), [
        "TA1 返鄉換屋族（45%）— 30~45 歲，要 4 房",
        "TA2 在地二代成家（35%）— 35~50 歲，要新成屋",
        "TA3 雙北退休置產（20%）— 50+，要寧靜面寬",
        "三組 TA 都偏「沉靜、品味、家的溫度」",
        "命名須符合 → 安靜、優雅、不浮誇",
    ], size=13, line_space=1.4)

    add_footer(s, "命名標準三條：①延續品牌、②呼應主視覺、③契合客群調性")
    add_page_number(s, 2, TOTAL)

    # ───── 3. 「中道」的內涵 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　「中道」二字", "地理與哲學的雙重意涵")

    layers = [
        ("地理層", DARK_BROWN, [
            "壯圍鄉「中道」為當地地段／聚落名",
            "鄰近順和路、南北五路，",
            "壯圍鄉內歷史聚落之一",
            "對在地客（TA2）而言：辨識度極高",
        ]),
        ("哲學層", ACCENT_RED, [
            "「中庸之道」— 不偏不倚、平衡",
            "對應客群心境：在工作與家、",
            "繁華與寧靜、北漂與返鄉之間",
            "對返鄉客（TA1）與退休客（TA3）：",
            "情感共鳴",
        ]),
        ("品牌層", DARK_BROWN, [
            "兩字結構穩定、字形對稱",
            "視覺上呼應主視覺花卉的均衡感",
            "與「玉田」（地名）系列延續性強",
            "與其他建案不易混淆",
        ]),
    ]
    col_w = Inches(4.0); col_h = Inches(4.8); col_y = Inches(1.8)
    for i, (label, color, bullets) in enumerate(layers):
        x = Inches(0.5) + i * Inches(4.2)
        add_rect(s, x, col_y, col_w, col_h, fill_color=CREAM, line_color=color)
        add_text(s, x, col_y + Inches(0.15), col_w, Inches(0.5),
                 label, size=20, bold=True, color=color, align="center")
        add_bullets(s, x + Inches(0.25), col_y + Inches(0.85), col_w - Inches(0.5), Inches(3.7),
                    bullets, size=12.5, line_space=1.45, bullet_color=color)

    add_footer(s, "中道之於本案：地理錨點 + 心境共鳴 + 品牌延續，三重資產。")
    add_page_number(s, 3, TOTAL)

    # ───── 4. 「森活」的內涵 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　「森活」二字", "鼎弘已建立的品牌資產")

    # 左：森活的意涵
    add_text(s, Inches(0.5), Inches(1.7), Inches(6.5), Inches(0.5),
             "「森活」= 森林 × 生活", size=20, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(0.5), Inches(2.3), Inches(6.3), Inches(3.8), [
        "「森」字：取「叢生、繁茂」 — 對應宜蘭的綠意",
        "「活」字：取「生活的態度」 — 對應 TA 的價值觀",
        "音韻上「森活」≈「生活」雙關，記憶點高",
        "玉田森活 1 + 2 期合計 44 戶完銷 — 已有品牌 awareness",
        "新案沿用「森活」= 直接繼承玉田客群口碑",
    ], size=13.5, line_space=1.5)

    # 右：品牌系列圖
    add_rect(s, Inches(7.5), Inches(1.7), Inches(5.3), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(7.5), Inches(1.85), Inches(5.3), Inches(0.5),
             "鼎弘「森活」系列", size=17, bold=True, color=DARK_BROWN, align="center")
    series = [
        ("玉田森活 1 期", "礁溪 30 戶 / 已完銷"),
        ("玉田森活 2 期", "礁溪 14 戶 / 已完銷"),
        ("中道森活（本案）", "壯圍 27 戶 / 推案中"),
    ]
    for i, (name, status) in enumerate(series):
        y = Inches(2.6) + i * Inches(1.25)
        h = Inches(1.0)
        fill = ACCENT_RED if i == 2 else WHITE
        text_color = WHITE if i == 2 else DARK_BROWN
        add_rect(s, Inches(7.8), y, Inches(4.7), h,
                 fill_color=fill, line_color=DARK_BROWN)
        add_text(s, Inches(7.95), y + Inches(0.12), Inches(4.5), Inches(0.45),
                 name, size=16, bold=True, color=text_color)
        add_text(s, Inches(7.95), y + Inches(0.55), Inches(4.5), Inches(0.4),
                 status, size=12, color=text_color)

    add_footer(s, "「森活」= 鼎弘建設不可放棄的品牌資產")
    add_page_number(s, 4, TOTAL)

    # ───── 5. 三條路線總覽 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　三條命名策略路線", "依「品牌延續性」與「辨識度」雙軸定位")

    routes = [
        ("路線 A\n純延續", "「中道森活」", BRAND_YELLOW_SOFT, [
            "極簡、直接、易讀",
            "完全沿用業主方向",
            "辨識度：中（會與玉田森活混淆）",
        ]),
        ("路線 B\n副標延伸", "「中道森活 | ◯」", CREAM, [
            "保留品牌系列",
            "副標字加上獨特辨識",
            "辨識度：高｜推薦度：★★★",
        ]),
        ("路線 C\n系列編號", "「中道森活 NO.x」", CREAM, [
            "工業／極簡感",
            "與玉田 1/2 同邏輯",
            "辨識度：中｜年輕感較強",
        ]),
    ]
    col_w = Inches(4.0); col_h = Inches(5.2); col_y = Inches(1.7)
    for i, (label, sample, fill, bullets) in enumerate(routes):
        x = Inches(0.5) + i * Inches(4.2)
        line_color = ACCENT_RED if i == 1 else DARK_BROWN
        add_rect(s, x, col_y, col_w, col_h, fill_color=fill, line_color=line_color)
        add_text(s, x, col_y + Inches(0.15), col_w, Inches(0.95),
                 label, size=20, bold=True, color=DARK_BROWN, align="center", line_space=1.2)
        add_text(s, x, col_y + Inches(1.25), col_w, Inches(0.7),
                 sample, size=22, bold=True, color=ACCENT_RED, align="center",
                 font=FONT_TITLE)
        add_bullets(s, x + Inches(0.25), col_y + Inches(2.2), col_w - Inches(0.5), Inches(2.8),
                    bullets, size=12.5, line_space=1.45)

    add_text(s, Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.4),
             "下三頁逐一展開每條路線的子提案。", size=13, color=DARK_BROWN, bold=True, align="center")
    add_page_number(s, 5, TOTAL)

    # ───── 6. 路線 A 細節 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　路線 A — 純延續", "「中道森活」")

    add_text(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(1.0),
             "中  道  森  活", size=66, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
    add_text(s, Inches(0.5), Inches(2.95), Inches(12.3), Inches(0.5),
             "ZHONG DAO SEN HUO", size=18, color=GRAY_70, align="center")

    # pros / cons
    add_rect(s, Inches(0.5), Inches(3.8), Inches(5.9), Inches(2.8),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.5), Inches(3.95), Inches(5.9), Inches(0.4),
             "優點", size=16, bold=True, color=DARK_BROWN, align="center")
    add_bullets(s, Inches(0.8), Inches(4.4), Inches(5.4), Inches(2.2), [
        "完全延續業主品牌方向",
        "案名極簡，辨識成本低",
        "與玉田森活高度延續性",
    ], size=13)

    add_rect(s, Inches(6.9), Inches(3.8), Inches(5.9), Inches(2.8),
             fill_color=CREAM, line_color=ACCENT_RED)
    add_text(s, Inches(6.9), Inches(3.95), Inches(5.9), Inches(0.4),
             "顧慮", size=16, bold=True, color=ACCENT_RED, align="center")
    add_bullets(s, Inches(7.2), Inches(4.4), Inches(5.4), Inches(2.2), [
        "591/Google 搜尋與玉田森活混淆",
        "若兩案銷售期重疊，客戶會問是哪一期",
        "無獨立 narrative，僅靠地名分辨",
    ], size=13, bullet_color=ACCENT_RED)

    add_text(s, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
             "適用情境：玉田 1 / 2 已交屋並淡出市場，本案不需獨立 narrative 時。",
             size=12, color=GRAY_70, align="center")
    add_page_number(s, 6, TOTAL)

    # ───── 7. 路線 B 細節 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　路線 B — 副標延伸 ★ 推薦", "「中道森活｜◯」")

    candidates = [
        ("栖", "中道森活｜栖", "棲息、停靠", "TA3 退休置產", "★★★★★"),
        ("序", "中道森活｜序", "序章、開始", "TA1 返鄉換屋", "★★★★"),
        ("誠", "中道森活｜誠", "誠意、用心", "TA2 在地二代", "★★★★"),
        ("拾光", "中道森活｜拾光", "拾起時光", "TA1+TA3", "★★★"),
        ("植光", "中道森活｜植光", "植下陽光", "TA2 年輕家庭", "★★★"),
    ]
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.2),
              ["副標字", "全名", "字意", "對應 TA", "推薦度"],
              [list(c) for c in candidates],
              header_size=13, body_size=14, first_col_bold=True,
              highlight_rows=[0])

    add_rect(s, Inches(0.5), Inches(6.1), Inches(12.3), Inches(1.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(6.18), Inches(12), Inches(0.4),
             "推薦理由：「中道森活｜栖」", size=14, bold=True, color=ACCENT_RED)
    add_text(s, Inches(0.7), Inches(6.6), Inches(12), Inches(0.5),
             "「栖」承接主視覺「在繁華之中收藏寧靜」— 一個字打中三組 TA 的共同價值觀「想要一個落腳的家」。",
             size=12, color=DARK_BROWN, line_space=1.35)
    add_page_number(s, 7, TOTAL)

    # ───── 8. 路線 C 細節 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　路線 C — 系列編號", "「中道森活 NO.x」")

    options = [
        ("中道森活 NO.1", "與玉田森活 1/2 同邏輯，本案=系列第三作。直觀。"),
        ("中道森活 III",  "羅馬數字 III，呼應「鼎弘第 3 個森活」。視覺有設計感。"),
        ("中道森活 27",   "強調戶數 27 戶。年輕、極簡。"),
        ("中道森活｜小宇宙", "副標走詩意路線，年輕客層共鳴。"),
    ]
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.0),
              ["命名", "賣點 / 風格"],
              [list(o) for o in options],
              header_size=13, body_size=14, first_col_bold=True)

    add_text(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(0.5),
             "潛在風險：系列編號顯得「公式化」，與業主希望的「典雅生活感」主視覺有違和。",
             size=13, color=ACCENT_RED, bold=True)
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.5),
             "建議：若選 C，建議 III（最有設計感、與羅馬數字典雅感呼應主視覺花卉）。",
             size=13, color=DARK_BROWN)
    add_page_number(s, 8, TOTAL)

    # ───── 9. TOP 3 推薦 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　推薦 TOP 3", "與三組 TA 之 message 對應")

    rec = [
        ("第一推薦", "中道森活｜栖",
         "TA3 退休置產 + TA1 返鄉換屋",
         "「在繁華之中，找到一個栖息」",
         BRAND_YELLOW, WHITE),
        ("第二推薦", "中道森活 III",
         "TA1 返鄉換屋 + TA2 在地二代",
         "「鼎弘森活，第三個家的故事」",
         CREAM, DARK_BROWN),
        ("第三推薦", "中道森活",
         "全 TA 安全牌",
         "「玉田之後，中道再續」",
         CREAM, DARK_BROWN),
    ]
    box_h = Inches(1.55); y0 = Inches(1.85)
    for i, (rank, name, ta, msg, fill, text_color) in enumerate(rec):
        y = y0 + i * (box_h + Inches(0.1))
        add_rect(s, Inches(0.5), y, Inches(12.3), box_h, fill_color=fill, line_color=DARK_BROWN)
        add_text(s, Inches(0.7), y + Inches(0.15), Inches(2.0), Inches(0.5),
                 rank, size=15, bold=True, color=text_color)
        add_text(s, Inches(2.6), y + Inches(0.1), Inches(4.5), Inches(0.6),
                 name, size=22, bold=True, color=text_color, font=FONT_TITLE)
        add_text(s, Inches(2.6), y + Inches(0.8), Inches(4.5), Inches(0.4),
                 ta, size=11, color=text_color)
        add_text(s, Inches(7.3), y + Inches(0.45), Inches(4.9), Inches(0.7),
                 msg, size=14, color=text_color, line_space=1.3)
    add_page_number(s, 9, TOTAL)

    # ───── 10. 主視覺對應 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　主視覺與命名搭配",
                  "業主已選方向（黃底花卉典雅插畫） — 兩個版本的對應建議")

    # 兩欄
    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW, line_color=DARK_BROWN)
    add_text(s, Inches(0.5), Inches(2.4), Inches(5.9), Inches(0.7),
             "中  道  森  活", size=32, bold=True, color=WHITE, align="center", font=FONT_TITLE)
    add_text(s, Inches(0.5), Inches(3.2), Inches(5.9), Inches(0.4),
             "（鬱金香 + 水仙｜白色標題）", size=11, color=WHITE, align="center")
    add_text(s, Inches(0.7), Inches(4.0), Inches(5.5), Inches(0.5),
             "版本 A", size=18, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(0.7), Inches(4.5), Inches(5.5), Inches(2.0), [
        "色調明亮、年輕感較強",
        "適合：TA2 在地二代、TA1 年輕返鄉",
        "建議命名：中道森活｜栖、植光",
    ], size=11.5, line_space=1.35, color=DARK_BROWN)

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(6.9), Inches(2.4), Inches(5.9), Inches(0.7),
             "中  道  森  活", size=32, bold=True, color=DARK_BROWN, align="center", font=FONT_TITLE)
    add_text(s, Inches(6.9), Inches(3.2), Inches(5.9), Inches(0.4),
             "（玫瑰｜深棕標題）", size=11, color=DARK_BROWN, align="center")
    add_text(s, Inches(7.1), Inches(4.0), Inches(5.5), Inches(0.5),
             "版本 B", size=18, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(7.1), Inches(4.5), Inches(5.5), Inches(2.0), [
        "色調沉穩、典雅感較強",
        "適合：TA3 退休置產、高品味 TA1",
        "建議命名：中道森活｜栖、序、誠",
    ], size=11.5, line_space=1.35, color=DARK_BROWN)

    add_text(s, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
             "建議：先選命名 → 再依命名調主視覺色調強弱（版本 B 配「栖」協同最佳）",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 10, TOTAL)

    # ───── 11. 法規／域名檢查清單 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "10　商標 / 域名 / 法規檢查清單",
                  "命名定案前的 5 項驗證")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["#", "檢查項", "工具", "現況", "風險"],
              [
                  ["1", "商標檢索（智財局）", "tipo.gov.tw 商標檢索系統",
                   "待查「中道森活」、「中道森活｜栖」", "可能與既有商標衝突"],
                  ["2", "591 案名查重", "market.591.com.tw",
                   "已查無同名宜蘭案", "✅ 通過"],
                  ["3", "Google 搜尋衝突", "Google",
                   "需確認無同名公司／品牌", "依命名而定"],
                  ["4", "域名可用性", "Google Domains、Gandi",
                   "建議買 zhong-dao-sen-huo.com / .tw", "易被搶註"],
                  ["5", "Facebook / IG 帳號", "Meta",
                   "建議同步註冊", "易被搶註"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True)
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "璞域可代執行第 1、4、5 項；建議命名定案後 5 個工作天內完成。",
             size=13, color=DARK_BROWN, bold=True)
    add_page_number(s, 11, TOTAL)

    # ───── 12. 下一步 ─────
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "11　下一步推進計畫", "命名定案 → CI → 主視覺 → 上市")
    steps = [
        ("Step 1\nW1",  "本簡報業主回饋",  "確認 1 案名（A / B / C 路線）", BRAND_YELLOW_SOFT),
        ("Step 2\nW1~W2", "商標／域名驗證", "璞域代執行 5 項檢查", CREAM),
        ("Step 3\nW2~W3", "案名最終定稿",  "若有衝突回到備案", CREAM),
        ("Step 4\nW3~W6", "CI / Logo 設計", "案名 + 副標 + 字體系統", BRAND_YELLOW_SOFT),
        ("Step 5\nW4~W8", "主視覺迭代",     "Banner / DM / 看屋會館", CREAM),
    ]
    box_w = Inches(2.4); y0 = Inches(2.0); box_h = Inches(3.6)
    for i, (label, action, body, fill) in enumerate(steps):
        x = Inches(0.5) + i * (box_w + Inches(0.13))
        add_rect(s, x, y0, box_w, box_h, fill_color=fill, line_color=DARK_BROWN)
        add_text(s, x, y0 + Inches(0.15), box_w, Inches(0.8),
                 label, size=14, bold=True, color=ACCENT_RED, align="center", line_space=1.2)
        add_text(s, x, y0 + Inches(1.0), box_w, Inches(0.6),
                 action, size=13, bold=True, color=DARK_BROWN, align="center", line_space=1.3)
        add_text(s, x + Inches(0.15), y0 + Inches(1.7), box_w - Inches(0.3), Inches(1.8),
                 body, size=11, color=DARK_BROWN, line_space=1.4, align="center")

    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "目標：8 週內完成命名定稿與主視覺 1.0，銜接 Q4 2026 上市時程。",
             size=14, bold=True, color=DARK_BROWN, align="center")
    add_text(s, Inches(0.5), Inches(6.8), Inches(12.3), Inches(0.4),
             "璞域品牌策略｜crazywofes@gmail.com", size=11, color=GRAY_70, align="right")
    add_page_number(s, 12, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
