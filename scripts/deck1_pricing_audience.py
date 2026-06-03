"""簡報 1：中道森活 — 客群定位與售價策略簡報 (給鼎弘建設用).

輸出：02_簡報/01_中道森活_客群與定價簡報.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "01_中道森活_客群與定價簡報.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()

    # ───── 封面 ─────
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="客群定位與售價策略簡報",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ───── 1. 簡報目錄 ─────
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　本案概況：中道森活 27 戶透天",
        "02　宜蘭壯圍透天市場全景（24 個月）",
        "03　競爭地圖：關聯案、競品、自有戰績",
        "04　鼎弘自有戰績驗證 — 三案完銷的市場意義",
        "05　核心客群：三組 TA 與其消費邏輯",
        "06　售價策略：三情境、戶別差異化",
        "07　與宗硯大砌的協同操作建議",
        "08　去化預估與行銷時程",
        "09　下一步：待業主決策事項",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ───── 2. 本案概況 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　本案概況", "中道森活 — 鼎弘建設新推案基本資料")
    headers = ["欄位", "內容"]
    rows = [
        ["案名（業主初步方向）", "中道森活"],
        ["起造人", "鼎弘健康生技有限公司　代表人 陳鼎琳"],
        ["營造", "宗硯建設"],
        ["區位", "宜蘭縣壯圍鄉五權段 789、790 等 45 筆地號"],
        ["建照核準日", "民國 115.01.30（114.09.24 府建都字第 1140160617 號）"],
        ["規模", "27 戶（A~G 共 7 幢，透天獨棟）"],
        ["建築型態", "地上 3 層 + 屋頂層，無地下層，RC 構造"],
        ["建坪／地坪", "建坪 38~43 坪｜地坪 22~37 坪"],
        ["用途", "H-2 住家用（特定農業區甲種建築用地）"],
        ["法定建蔽率／容積率", "60% / 180%（實設 38~60% / 154~178%，均符合）"],
        ["銷售方式", "自售（沿用鼎弘玉田森活模式）"],
        ["目標", "總價守在千萬內，27 戶 18~24 個月完銷"],
    ]
    add_table(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0),
              headers, rows, header_size=13, body_size=12, first_col_bold=True,
              body_align="left")
    add_footer(s, "資料來源：建照核準圖 DWG（115.02.13 版）")
    add_page_number(s, 2, 12)

    # ───── 3. 壯圍透天市場全景 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　宜蘭壯圍透天市場全景",
                  "24 個月｜內政部實價登錄 2024/05/14 – 2026/05/14")
    # 三大數字
    metrics = [
        ("198", "壯圍鄉 24 個月\n中古透天成交筆數", DARK_BROWN),
        ("14", "壯圍鄉 24 個月\n預售透天案數", ACCENT_RED),
        ("21.9", "壯圍透天成交\n單價中位（萬/坪）", DARK_BROWN),
        ("1,050", "壯圍透天成交\n總價中位（萬元）", ACCENT_RED),
    ]
    box_w = Inches(2.9)
    box_h = Inches(2.0)
    x0 = Inches(0.5)
    for i, (big, label, color) in enumerate(metrics):
        x = x0 + i * (box_w + Inches(0.13))
        add_rect(s, x, Inches(1.85), box_w, box_h, fill_color=CREAM, line_color=DARK_BROWN)
        add_text(s, x, Inches(2.0), box_w, Inches(1.0),
                 big, size=44, bold=True, color=color, align="center", font=FONT_TITLE)
        add_text(s, x, Inches(3.05), box_w, Inches(0.85),
                 label, size=12, color=GRAY_70, align="center", line_space=1.3)

    add_text(s, Inches(0.5), Inches(4.2), Inches(12.3), Inches(0.4),
             "焦點三區（壯圍、五結、宜蘭市）中古透天成交筆數", size=14, bold=True, color=DARK_BROWN)
    add_table(s, Inches(0.5), Inches(4.6), Inches(12.3), Inches(1.7),
              ["區", "筆數", "每坪中位(萬)", "P25 ~ P75", "總價中位(萬)", "建坪中位", "屋齡中位"],
              [
                  ["壯圍鄉 (主)", "198", "21.9", "16.6 ~ 25.1", "1,050", "53.5 坪", "10.0 年"],
                  ["五結鄉", "449", "20.8", "16.3 ~ 25.2", "1,066", "51.1 坪", "15.9 年"],
                  ["宜蘭市", "546", "25.2", "19.6 ~ 31.0", "1,200", "49.2 坪", "30.8 年"],
              ],
              header_size=11, body_size=11, first_col_bold=True, highlight_rows=[0])
    add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
             "結論：壯圍中古透天屋齡中位僅 10 年，是宜蘭三區最年輕的市場；行情貼近五結略低於宜蘭市。",
             size=12, color=DARK_BROWN, bold=True)
    add_footer(s, "資料來源：內政部實價登錄 plvr.land.moi.gov.tw（4 季批次）")
    add_page_number(s, 3, 12)

    # ───── 4. 競爭地圖 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　競爭地圖", "依「關係」重新分類 — 關聯案 ≠ 競品")

    # 三欄
    col_w = Inches(4.0); col_h = Inches(5.2); col_y = Inches(1.7)
    cols = [
        (Inches(0.5), "關聯案", "同營造廠（宗硯）", BRAND_YELLOW_SOFT, [
            "宗硯大砌系列18-透天",
            "31 戶｜建坪 40.7｜21.2 萬/坪",
            "總價中位 858 萬",
            "2026.02 開賣中",
            "",
            "戰術：對齊定價、分隔賣點，",
            "不打價格戰",
        ]),
        (Inches(4.7), "真正競品", "不同建設、千萬內同帶", CREAM, [
            "古亭郡（12 戶｜19.1 萬｜870 萬）",
            "  ↳ 低價帶代表，永美路",
            "麒美紅葉（7 戶｜19.6 萬｜995 萬）",
            "  ↳ 紅葉路，低單價策略",
            "艾利永美（21 戶｜22.7 萬｜1,078 萬）",
            "  ↳ 略超千萬，銷售期 8 個月",
            "赫宇聯勤2（7 戶｜21.2 萬｜1,018 萬）",
            "  ↳ 千萬內上緣參考",
        ]),
        (Inches(8.9), "自有戰績", "鼎弘 3 案均已完銷", CREAM, [
            "玉田森活1（30 戶｜20.8 萬｜930 萬）",
            "  ↳ 2 年完銷，30 筆實登驗證",
            "玉田森活2（14 戶｜22.0 萬｜1,513 萬）",
            "  ↳ 建材升級拉高單價",
            "學進賦（10 戶｜19.9 萬｜587 萬）",
            "  ↳ 公寓型，五結二結",
            "",
            "三案皆完銷 = 品牌資產",
        ]),
    ]
    for x, title, sub, fill, items in cols:
        add_rect(s, x, col_y, col_w, col_h, fill_color=fill, line_color=DARK_BROWN)
        add_text(s, x, col_y + Inches(0.15), col_w, Inches(0.5),
                 title, size=20, bold=True, color=DARK_BROWN, align="center")
        add_text(s, x, col_y + Inches(0.6), col_w, Inches(0.4),
                 sub, size=11, color=GRAY_70, align="center")
        body_lines = "\n".join(items)
        add_text(s, x + Inches(0.2), col_y + Inches(1.15), col_w - Inches(0.4), Inches(4),
                 body_lines, size=12, color=DARK_BROWN, line_space=1.45)

    add_footer(s, "資料來源：內政部實價登錄 + 業主提供推案總表")
    add_page_number(s, 4, 12)

    # ───── 5. 自有戰績 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　鼎弘自有戰績是最好的市場驗證",
                  "三案皆完銷｜玉田森活1 模式可複製")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(2.5),
              ["", "玉田森活1", "玉田森活2", "學進賦"],
              [
                  ["區域", "礁溪鄉", "礁溪鄉", "五結鄉"],
                  ["類別／型態", "預售透天 3 層", "預售透天 4 層", "預售公寓 5 層"],
                  ["戶數", "30 戶", "14 戶", "10 戶"],
                  ["建坪", "42~49 坪", "65~70 坪", "29~30 坪"],
                  ["格局", "4 房", "4 房", "—"],
                  ["實登中位總價", "930 萬", "—（591：1,468~1,558）", "—"],
                  ["實登中位單價", "20.8 萬/坪", "—（591：22 萬）", "—（591：19.9 萬）"],
                  ["去化", "2 年完銷", "完銷", "完銷"],
                  ["銷售方式", "自售", "自售", "自售"],
              ],
              header_size=12, body_size=11, first_col_bold=True, highlight_rows=[0])

    add_text(s, Inches(0.5), Inches(5.45), Inches(12.3), Inches(0.4),
             "本案模型可直接對標：玉田森活1", size=15, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.5), Inches(5.85), Inches(12.3), Inches(1.0), [
        "本案建坪 38~43 坪 < 玉田森活1 42~49 坪 → 相同單價下，總價自然更低，守千萬內有把握",
        "本案 27 戶 ≈ 玉田森活1 30 戶 → 規模相近，去化模型可直接套用",
        "自售 + 4 房 + 透天 3 層 = 鼎弘已驗證的「壯圍/礁溪首購升級」配方",
    ], size=12, line_space=1.3)

    add_footer(s, "資料來源：業主提供推案總表（591 × 政府實登 × 房地王交叉）")
    add_page_number(s, 5, 12)

    # ───── 6. 核心客群 — 三組 TA ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　核心客群", "從鼎弘自有戰績反推 — 三組目標客群")

    # 三組 TA
    tas = [
        ("TA1\n宜蘭返鄉換屋族",
         "主力 45%",
         BRAND_YELLOW_SOFT,
         ["30~45 歲，已婚已育",
          "雙北或南部工作 5~10 年",
          "家庭年收 100~150 萬",
          "想回宜蘭照顧長輩",
          "預算 800~950 萬",
          "在乎：4 房、學區、爸媽近"]),
        ("TA2\n壯圍/礁溪在地二代",
         "主力 35%",
         CREAM,
         ["35~50 歲，宜蘭土生土長",
          "已有第一房（公寓/老透天）",
          "想換新成屋透天",
          "家庭年收 80~130 萬",
          "預算 850~1,000 萬",
          "在乎：通勤、停車、生活機能"]),
        ("TA3\n雙北退休置產",
         "輔助 20%",
         CREAM,
         ["50+ 歲，子女已成年",
          "北部退休或半退休",
          "想回宜蘭養老",
          "資產雄厚不貸款居多",
          "預算 900~1,000 萬",
          "在乎：好天氣、寧靜、面寬"]),
    ]
    col_w = Inches(4.0); col_h = Inches(5.3); col_y = Inches(1.7)
    for i, (title, ratio, fill, bullets) in enumerate(tas):
        x = Inches(0.5) + i * Inches(4.2)
        add_rect(s, x, col_y, col_w, col_h, fill_color=fill, line_color=DARK_BROWN)
        add_text(s, x, col_y + Inches(0.15), col_w, Inches(0.95),
                 title, size=18, bold=True, color=DARK_BROWN, align="center", line_space=1.15)
        add_text(s, x, col_y + Inches(1.1), col_w, Inches(0.4),
                 ratio, size=14, bold=True, color=ACCENT_RED, align="center")
        add_bullets(s, x + Inches(0.25), col_y + Inches(1.65), col_w - Inches(0.5), Inches(3.5),
                    bullets, size=12, line_space=1.35)
    add_footer(s, "客群推算依據：壯圍鄉透天交易屋齡中位 10 年 + 鼎弘玉田森活 1 / 2 期客層")
    add_page_number(s, 6, 12)

    # ───── 7. 客群與賣點對應 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05a　客群 × 賣點對應", "每組 TA 應該被哪一句話打中")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["客群", "他們在意的", "本案能給的", "主打 message"],
              [
                  ["TA1 返鄉換屋",
                   "4 房、爸媽通勤近、子女學區",
                   "4 房格局、壯圍離宜蘭市 15 分鐘、近五權國小",
                   "「回宜蘭，4 房一次到位」"],
                  ["TA2 在地二代",
                   "通勤、生活機能、停車",
                   "壯圍郊區但近台 2 線、自家前院停車、千萬有找",
                   "「不到千萬，住進新成屋透天」"],
                  ["TA3 退休置產",
                   "寧靜、面寬感、不要太密",
                   "27 戶分 7 幢、外牆退縮、屋頂層可作休憩",
                   "「在繁華之中收藏寧靜」"],
              ],
              header_size=12, body_size=12, first_col_bold=True)

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.6),
             "三組 message 對應三個媒體投放路線（詳第 11 頁）", size=13, color=DARK_BROWN, bold=True)
    add_footer(s, "Message 設計：呼應業主主視覺文案「在繁華之中收藏寧靜，在日常之間遇見家的溫度」")
    add_page_number(s, 7, 12)

    # ───── 8. 售價策略 — 三情境 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　售價策略 — 三情境",
                  "三情境總價皆守在千萬內｜建議主力 21 萬/坪")

    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(2.6),
              ["情境", "每坪(萬)", "40 坪總價", "43 坪總價", "對標案", "評估"],
              [
                  ["保守（衝去化）", "19.5", "780 萬", "839 萬",
                   "古亭郡、麒美紅葉", "不建議：破壞宗硯定價、影響社群"],
                  ["★ 建議主力", "21.0", "840 萬", "903 萬",
                   "宗硯大砌、自有玉田1", "對齊關聯案、複製自有戰績"],
                  ["進取（品牌溢價）", "22.0", "880 萬", "946 萬",
                   "玉田森活2、艾利永美", "可在好區位戶嘗試"],
              ],
              header_size=12, body_size=12, first_col_bold=True, highlight_rows=[1])

    add_rect(s, Inches(0.5), Inches(4.7), Inches(12.3), Inches(2.0),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.85), Inches(12), Inches(0.45),
             "為什麼推薦 21 萬/坪？", size=16, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.3), Inches(11.9), Inches(1.4), [
        "對齊宗硯大砌 21.2 ±0.5 萬/坪，與關聯案不打架",
        "對標鼎弘玉田森活1 實登 20.8 萬/坪（已完銷）",
        "對標壯圍中古近新成屋 22.7 萬/坪（屋齡 ≤5 年、42 筆）",
        "27 戶 × 平均 41 坪 × 21 萬 = 預估全案銷售總額 約 2.32 億",
    ], size=13)
    add_footer(s, "資料截止：2026-05-14｜口徑：每坪 = 3.305785 m²")
    add_page_number(s, 8, 12)

    # ───── 9. 戶別差異化 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06a　戶別差異化定價",
                  "27 戶 × 7 幢｜地坪 72~123 m² 差異大，齊頭一價會虧")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["戶型分群", "代表戶", "地坪", "建議單價", "預估總價"],
              [
                  ["大地坪角間",     "E3",            "37 坪 (123 m²)", "22.0 萬", "≈ 990 萬"],
                  ["中型角間",       "E1、D1、E2",    "27~33 坪",       "21.5 萬", "≈ 925 萬"],
                  ["標準大戶",       "F4、G1",        "25~30 坪",       "21.5 萬", "≈ 925 萬"],
                  ["標準中戶",       "B1~B7、C1~C3", "22~23 坪",        "21.0 萬", "≈ 870 萬"],
                  ["A1~A3 入口戶",  "A1、A2、A3",   "22~29 坪",       "20.5 萬", "≈ 850 萬"],
              ],
              header_size=12, body_size=12, first_col_bold=True, highlight_rows=[3])

    add_text(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.4),
             "預估全案均價 870~900 萬｜27 戶總銷 2.30~2.45 億", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.5), Inches(6.05), Inches(12.3), Inches(1.0), [
        "屋頂層、邊間、雙面採光的戶可再加 3~5% 溢價",
        "前 5 戶（公開前內部讓利）可低 0.5~1 萬/坪 — 是鼎弘玉田1 的去化技巧",
    ], size=12)

    add_footer(s, "戶別細節依 DWG（A1~G3 共 27 戶），地坪數依「建蔽率 / 建築面積」反推")
    add_page_number(s, 9, 12)

    # ───── 10. 與宗硯協同 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　與宗硯大砌的協同操作",
                  "同營造、同價帶、同客層 — 對齊單價、分隔賣點、錯開時程")

    # 對照欄
    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(4.8),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(4.8),
             fill_color=CREAM, line_color=DARK_BROWN)

    add_text(s, Inches(0.5), Inches(1.8), Inches(5.9), Inches(0.5),
             "宗硯大砌系列18-透天", size=18, bold=True, color=DARK_BROWN, align="center")
    add_text(s, Inches(0.5), Inches(2.25), Inches(5.9), Inches(0.4),
             "31 戶｜21.2 萬/坪｜858 萬", size=13, color=GRAY_70, align="center")
    add_bullets(s, Inches(0.8), Inches(2.85), Inches(5.4), Inches(3.5), [
        "賣點：「最大量體 + 最低單價」",
        "進度：2026 Q2 衝刺中",
        "強項：價格戰、宗硯品牌",
        "弱項：宗硯首次自推、無自售戰績",
    ], size=13)

    add_text(s, Inches(6.9), Inches(1.8), Inches(5.9), Inches(0.5),
             "中道森活（本案）", size=18, bold=True, color=ACCENT_RED, align="center")
    add_text(s, Inches(6.9), Inches(2.25), Inches(5.9), Inches(0.4),
             "27 戶｜21 萬/坪｜870~900 萬", size=13, color=GRAY_70, align="center")
    add_bullets(s, Inches(7.2), Inches(2.85), Inches(5.4), Inches(3.5), [
        "賣點：「鼎弘 3 案完銷 + 自售品牌信譽」",
        "進度：2026 Q3~Q4 後段推（錯開 1~2 季）",
        "強項：鼎弘玉田森活完銷實績、自售議價彈性",
        "弱項：戶數較少、建坪偏小",
    ], size=13)

    add_text(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.4),
             "兩案合 58 戶≈壯圍 24 個月透天市場 29%；錯開時程是降低自相競爭的關鍵",
             size=12, color=DARK_BROWN, bold=True)
    add_footer(s, "本頁前提：宗硯大砌與本案為同營造廠關聯案（業主補充）")
    add_page_number(s, 10, 12)

    # ───── 11. 去化預估與時程 ─────
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　去化預估與行銷時程",
                  "對標玉田森活1（30 戶 / 2 年完銷）")

    # 時程圖
    add_text(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(0.4),
             "建議 18~24 個月完銷的時程帶", size=14, bold=True, color=DARK_BROWN)

    phases = [
        ("Pre-launch\n預備期",      "Q3 2026", "•定案、命名、CI 完成\n•接待中心籌備\n•VIP 名單預埋", BRAND_YELLOW_SOFT),
        ("Phase 1\n公開預售",       "Q4 2026", "•首批 8~10 戶（A、B 幢）\n•最強優惠／前 5 戶讓利\n•目標 3 個月去化 30%", CREAM),
        ("Phase 2\n結構成型",       "Q1~Q2 2027", "•中段 10 戶（C、D 幢）\n•主視覺 2.0、實品屋\n•目標 6 個月再去化 40%", CREAM),
        ("Phase 3\n完工交屋",       "Q3~Q4 2027", "•尾段 8~9 戶（E、F、G 幢）\n•品牌溢價、口碑帶單\n•目標 6 個月完銷", BRAND_YELLOW_SOFT),
    ]
    box_w = Inches(2.9); box_h = Inches(4.0); y0 = Inches(2.3)
    for i, (title, when, body, fill) in enumerate(phases):
        x = Inches(0.5) + i * (box_w + Inches(0.13))
        add_rect(s, x, y0, box_w, box_h, fill_color=fill, line_color=DARK_BROWN)
        add_text(s, x, y0 + Inches(0.15), box_w, Inches(0.8),
                 title, size=15, bold=True, color=DARK_BROWN, align="center", line_space=1.2)
        add_text(s, x, y0 + Inches(1.0), box_w, Inches(0.4),
                 when, size=12, color=ACCENT_RED, align="center", bold=True)
        add_text(s, x + Inches(0.2), y0 + Inches(1.55), box_w - Inches(0.4), Inches(2.3),
                 body, size=11, color=DARK_BROWN, line_space=1.4)
    add_text(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5),
             "全案 27 戶 / 預估 18~24 個月完銷｜銷售總額 2.30~2.45 億",
             size=14, bold=True, color=ACCENT_RED)
    add_footer(s, "節奏設計參考：玉田森活1 = 22 年 15 戶 + 23 年 14 戶 + 24 年完銷")
    add_page_number(s, 11, 12)

    # ───── 12. 下一步 ─────
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "09　下一步 — 待業主決策事項", "兩週內可確認的關鍵 5 項")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "決策項", "璞域建議", "業主回覆"],
              [
                  ["1", "案名定案", "「中道森活」+ 副標（詳命名提案簡報）", "_______"],
                  ["2", "主力單價", "21 萬/坪（與宗硯對齊）", "_______"],
                  ["3", "戶別定價", "依地坪/區位 20.5~22 萬分級", "_______"],
                  ["4", "推案時程", "Q4 2026 公開，2027 Q4 完銷", "_______"],
                  ["5", "向宗硯輔銷取得 大砌 18 內部數據", "成交曲線、客層、議價底線", "_______"],
              ],
              header_size=12, body_size=12, first_col_bold=True)
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "後續可同步處理：人口/家戶/所得補件、商標檢索、命名最終定案、CI 啟動",
             size=12, color=GRAY_70)
    add_text(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.4),
             "璞域品牌策略｜www.pure-branding.com", size=11, color=GRAY_70, align="right")
    add_page_number(s, 12, 12)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
