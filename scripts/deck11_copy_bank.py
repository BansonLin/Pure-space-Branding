"""簡報 11：中道森活 — DM 與廣告文案系列簡報.

包含：3 個 taglines、6 個標題、副標、body copy、CTA、
社群短文案、LINE 訊息、售屋手冊文案。

輸出：02_簡報/11_中道森活_廣告文案系列.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "11_中道森活_廣告文案系列.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()
    TOTAL = 14

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="DM 與廣告文案系列簡報　Copy Bank",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　文案系統架構 — 從 tagline 到 long copy",
        "02　Tagline 三案（品牌一句話）",
        "03　主標題系列（6 個方向）",
        "04　副標題系列",
        "05　Body Copy × TA1 返鄉換屋族",
        "06　Body Copy × TA2 在地二代",
        "07　Body Copy × TA3 退休置產",
        "08　CTA（行動呼籲）系列",
        "09　LINE 群組訊息（轉發引導）",
        "10　FB / IG 文案（含 hashtag）",
        "11　售屋手冊文案",
        "12　通路 × 文案配對表",
        "13　下一步",
    ]
    add_bullets(s, Inches(1.0), Inches(1.7), Inches(11), Inches(5.5),
                items, size=17, line_space=1.45)

    # ── 01 文案系統架構
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　文案系統架構",
                  "從一句話 → 一段話 → 完整故事")

    layers = [
        ("Tagline",   "10~15 字", "品牌一句話、招牌印象",        ACCENT_RED),
        ("主標題",     "15~25 字", "廣告主視覺打標、抓眼球",       DARK_BROWN),
        ("副標題",     "25~40 字", "解釋主標、補充訊息",          DARK_BROWN),
        ("Body Copy", "100~300 字", "情境式敘事、說服重點",        DARK_BROWN),
        ("CTA",       "5~10 字",  "明確下一步行動",              ACCENT_RED),
    ]
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["層級", "字數", "用途", "範例"],
              [
                  ["Tagline", "10~15 字", "品牌印象", "「在繁華之中，找一個栖息。」"],
                  ["主標題", "15~25 字", "廣告抓眼", "「中午能回家吃飯，是我搬回宜蘭最大的理由。」"],
                  ["副標題", "25~40 字", "解釋主標", "中道森活｜壯圍 27 戶｜透天 3 層｜千萬內｜2026 預售"],
                  ["Body Copy", "100~300 字", "情境敘事", "詳第 5-7 頁，依 TA 分版"],
                  ["CTA", "5~10 字", "下一步", "「預約看屋 →」/「了解更多 →」"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")

    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "原則：每層都來自相同的故事內核（林先生的一天），不會自相矛盾",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 2, TOTAL)

    # ── 02 Tagline
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　Tagline 三案",
                  "業主從 3 案中選定 — 將成為所有廣告物的標準字")

    taglines = [
        ("第一案", "在繁華之中，找一個栖息。",
         "承接業主主視覺文案「在繁華中收藏寧靜」",
         "全客層適用、典藏感", BRAND_YELLOW),
        ("第二案", "回宜蘭，4 房一次到位。",
         "直接訴求 TA1 主力（4 房+返鄉）",
         "TA1 主力訴求", BRAND_YELLOW_SOFT),
        ("第三案", "中道森活｜你想要的家，剛好在這裡。",
         "強調品牌、走親近路線",
         "全客層、品牌延伸", CREAM),
    ]
    box_h = 1.5; y0 = 1.85
    for i, (rank, tag, source, ta, fill) in enumerate(taglines):
        y = y0 + i * (box_h + 0.15)
        add_rect(s, Inches(0.5), Inches(y), Inches(12.3), Inches(box_h),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(0.7), Inches(y + 0.12), Inches(2), Inches(0.5),
                 rank, size=14, bold=True, color=text_color)
        add_text(s, Inches(0.7), Inches(y + 0.5), Inches(11.6), Inches(0.7),
                 f"「{tag}」", size=24, bold=True, color=text_color, font=FONT_TITLE)
        add_text(s, Inches(0.7), Inches(y + 1.15), Inches(7), Inches(0.3),
                 source, size=11, color=text_color)
        add_text(s, Inches(8.5), Inches(y + 1.15), Inches(4.3), Inches(0.3),
                 ta, size=11, bold=True, color=text_color)
    add_page_number(s, 3, TOTAL)

    # ── 03 主標題系列
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　主標題系列（6 個方向）",
                  "依不同訴求角度切入 — 可循環投放、避免疲乏")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["#", "主標題", "訴求角度", "對應 TA"],
              [
                  ["1", "「中午能回家吃飯，是我搬回宜蘭最大的理由。」",
                   "家人陪伴", "TA1"],
                  ["2", "「孩子的學校，4 分鐘車。」",
                   "學區優勢", "TA1 / TA2"],
                  ["3", "「在台北 8 年，省下來的時間，現在都還給家裡。」",
                   "時間換算", "TA1"],
                  ["4", "「住壯圍，假日不用出宜蘭。」",
                   "假日生活", "TA3 / TA1"],
                  ["5", "「不用千萬，住進新成屋透天。」",
                   "價格甜蜜點", "TA2"],
                  ["6", "「鼎弘 3 案完銷的承諾，繼續寫下去。」",
                   "品牌信賴", "全 TA"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_footer(s, "主標題建議搭配「中道森活｜栖」字標 + 業主已選定主視覺")
    add_page_number(s, 4, TOTAL)

    # ── 04 副標題系列
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　副標題系列",
                  "解釋主標、補資訊 — 通常與主標一起出現")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["#", "副標題", "搭配主標"],
              [
                  ["1", "中道森活｜宜蘭壯圍五權段｜27 戶獨棟透天｜2026 預售", "通用"],
                  ["2", "千萬內｜建坪 38~43 坪｜4 房格局｜公館國小學區",     "主標 2"],
                  ["3", "雪山隧道 50 分鐘到台北｜陽明交大附醫 14 分鐘",       "主標 1, 4"],
                  ["4", "鼎弘建設出品｜玉田森活 1/2 完銷｜自售品牌", "主標 6"],
                  ["5", "壯圍沙丘 7 分鐘｜羅東夜市 22 分鐘｜全宜蘭都在身邊", "主標 4"],
                  ["6", "首批 5 戶 VIP 優惠｜現場洽詢｜接待中心開放預約",   "全主標"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_page_number(s, 5, TOTAL)

    # ── 05 Body Copy × TA1
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　Body Copy × TA 1 返鄉換屋族",
                  "200 字版｜DM / 售屋手冊 / Facebook 廣告")

    body = ("在台北 8 年，我們每月房租 30,000、通勤 90 分鐘。\n\n"
            "搬回壯圍後：房貸 30,000、通勤一週兩次。\n\n"
            "省下的時間，是和爸媽、孩子的時間。\n\n"
            "中道森活，鼎弘建設第 3 個「森活」案。\n"
            "27 戶獨棟透天，建坪 38~43 坪，4 房格局。\n"
            "公館國小 4 分鐘車程、國道 5 壯圍交流道 8 分鐘、雪隧到台北 50 分鐘。\n\n"
            "玉田森活 1 期、2 期全數完銷。\n"
            "鼎弘的承諾，在中道繼續寫下去。")
    add_text(s, Inches(0.5), Inches(1.7), Inches(7.5), Inches(5.5),
             body, size=14, color=DARK_BROWN, line_space=1.6)

    add_rect(s, Inches(8.3), Inches(1.7), Inches(4.6), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(8.5), Inches(1.85), Inches(4.2), Inches(0.4),
             "文案結構", size=13, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(8.5), Inches(2.3), Inches(4.2), Inches(4.0), [
        "段 1：pain（台北高房租、長通勤）",
        "段 2：reframe（同房貸、短通勤）",
        "段 3：reward（時間=家人）",
        "段 4：產品介紹（戶數、規格）",
        "段 5：機能事實（學校、交通）",
        "段 6：品牌信譽（玉田完銷）",
        "段 7：呼籲（鼎弘繼續寫下去）",
    ], size=11, line_space=1.5)
    add_page_number(s, 6, TOTAL)

    # ── 06 Body Copy × TA2
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　Body Copy × TA 2 在地二代",
                  "180 字版｜在地版本、語氣更直接")

    body = ("在壯圍住了三十幾年，我們的第一個家，從爸媽的房子搬出來。\n\n"
            "看了五結兩個案、宜蘭市三個案。建商都不認識。\n\n"
            "中道森活｜鼎弘建設。\n"
            "玉田森活 1 期 30 戶、2 期 14 戶都完銷的那個鼎弘。\n\n"
            "不到千萬，住進新成屋透天。\n"
            "4 房格局、3 層、自家前院停車。\n"
            "公館國小 4 分鐘車、全聯 5 分鐘車、家樂福 15 分鐘車。\n\n"
            "在地人，住在地的家。")
    add_text(s, Inches(0.5), Inches(1.7), Inches(7.5), Inches(5.5),
             body, size=14, color=DARK_BROWN, line_space=1.65)

    add_rect(s, Inches(8.3), Inches(1.7), Inches(4.6), Inches(5.0),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(8.5), Inches(1.85), Inches(4.2), Inches(0.4),
             "文案語氣", size=13, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(8.5), Inches(2.3), Inches(4.2), Inches(4.0), [
        "用「在壯圍住了三十幾年」開場",
        "強調「建商都不認識」的不安",
        "鼎弘玉田完銷實績解套",
        "「不到千萬」是價格甜蜜點",
        "在地語感：「住在地的家」",
        "整段避免使用艱深詞彙",
    ], size=11, line_space=1.5)
    add_page_number(s, 7, TOTAL)

    # ── 07 Body Copy × TA3
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　Body Copy × TA 3 雙北退休置產",
                  "200 字版｜典雅、慢、不急著賣")

    body = ("孩子大了，老婆愛宜蘭。\n\n"
            "想要一個自己的透天，安靜、有院子、面寬好。\n\n"
            "中道森活｜栖。\n"
            "鼎弘建設在壯圍五權段的第 27 戶住宅。\n\n"
            "不到千萬。但每一戶都有自己的庭院。\n"
            "陽明交大附醫 14 分鐘車、雪山隧道 50 分鐘到台北。\n"
            "壯圍沙丘 7 分鐘、永鎮海濱 10 分鐘、礁溪溫泉 22 分鐘。\n\n"
            "在繁華之中，找一個栖息。")
    add_text(s, Inches(0.5), Inches(1.7), Inches(7.5), Inches(5.5),
             body, size=14, color=DARK_BROWN, line_space=1.65)

    add_rect(s, Inches(8.3), Inches(1.7), Inches(4.6), Inches(5.0),
             fill_color=BRAND_YELLOW, line_color=DARK_BROWN)
    add_text(s, Inches(8.5), Inches(1.85), Inches(4.2), Inches(0.4),
             "文案氣質", size=13, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(8.5), Inches(2.3), Inches(4.2), Inches(4.0), [
        "開頭一句話 — 不急著賣",
        "「面寬好」「有院子」對應 TA3 在意",
        "醫療放在前面（TA3 痛點）",
        "假日景點豐富 — 退休生活想像",
        "結尾「栖息」呼應 tagline",
        "整段節奏慢、留白多",
    ], size=11, line_space=1.5, color=DARK_BROWN)
    add_page_number(s, 8, TOTAL)

    # ── 08 CTA
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　CTA（行動呼籲）系列",
                  "從「看更多」到「下訂」")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["階段", "CTA 文字", "適用情境", "TA"],
              [
                  ["認知期", "了解更多 →",         "FB 廣告、IG 限動", "全 TA"],
                  ["認知期", "看看林先生的一天 →", "影音廣告連結",    "TA1"],
                  ["興趣期", "查看戶別與價格 →",   "591 案件頁、官網", "全 TA"],
                  ["興趣期", "下載售屋手冊 →",    "LINE 群組訊息",   "TA2"],
                  ["考慮期", "預約看屋 →",        "DM、廣告下單",    "全 TA"],
                  ["考慮期", "VIP 優惠登記 →",    "口碑、親友推薦",  "TA1 / TA2"],
                  ["決策期", "立即訂下 1 戶 →",   "現場銷售",         "全 TA"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_footer(s, "CTA 三層級對應「Phase 1 暖身 / Phase 2 主力 / Phase 3 精品」")
    add_page_number(s, 9, TOTAL)

    # ── 09 LINE 訊息
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　LINE 群組訊息",
                  "可直接轉發、口碑擴散用")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["版本", "訊息內容", "字數"],
              [
                  ["短版（朋友圈）",
                   "鼎弘建設新案「中道森活」要開賣了\n壯圍 27 戶透天｜千萬內｜4 房\n看屋連結 →",
                   "40 字"],
                  ["中版（家族群）",
                   "玉田森活鼎弘的新案，這次在壯圍\n27 戶透天｜建坪 38~43 坪\n公館國小 4 分車、雪隧 50 分到台北\n預約看屋 →",
                   "70 字"],
                  ["長版（在地人脈）",
                   "鼎弘建設第三個「森活」開賣了\n玉田 1 期 30 戶、2 期 14 戶都完銷\n這次在壯圍五權段，27 戶獨棟透天\n建坪 38~43 坪、4 房、3 層、千萬內\n公館國小 4 分車、家樂福 15 分車\n壯圍沙丘 7 分車、雪隧 50 分到台北\n看屋會館已開放 →",
                   "150 字"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_page_number(s, 10, TOTAL)

    # ── 10 FB / IG 文案
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "10　FB / IG 文案（含 hashtag）",
                  "社群投放用｜每週 2~3 篇輪播")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["平台", "文案範本", "Hashtag"],
              [
                  ["FB 廣告",
                   "「中午能回家吃飯，是我搬回宜蘭最大的理由。」\n\n中道森活｜壯圍 27 戶透天｜千萬內｜2026 預售\n了解更多 →",
                   "#中道森活 #鼎弘建設\n#壯圍透天 #宜蘭新案"],
                  ["IG 主貼",
                   "在台北 8 年，省下來的時間，現在都還給家裡。\n\n中道森活｜鼎弘建設第 3 個「森活」\n\n點擊 bio 連結看屋",
                   "#中道森活 #鼎弘建設\n#宜蘭 #壯圍 #透天\n#返鄉 #宜蘭生活"],
                  ["IG 限動",
                   "🌼 中道森活｜栖\n壯圍 27 戶｜千萬內\n2026 預售開始\n\n→ 滑動連結了解更多",
                   "—"],
                  ["FB 在地社團",
                   "壯圍五權段新案開賣\n鼎弘建設玉田森活的同主體\n27 戶透天｜4 房｜千萬內\n看屋會館已開放",
                   "—"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_page_number(s, 11, TOTAL)

    # ── 11 售屋手冊文案
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "11　售屋手冊文案",
                  "20 頁手冊的章節文案大綱")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["頁", "章節標題", "內容方向"],
              [
                  ["P1", "中道森活｜栖",                         "封面：tagline + 主視覺"],
                  ["P2~3", "鼎弘建設 — 第 3 個「森活」",         "玉田 1/2 完銷實績"],
                  ["P4~5", "壯圍 — 我們選擇這裡的理由",         "地理脈絡、5 圈生活範圍"],
                  ["P6~7", "「中道」二字的意涵",                "命名故事、東方文化"],
                  ["P8~9", "27 戶配置圖",                      "全區配置、A~G 幢"],
                  ["P10~11", "建築規格",                       "RC 結構、3 層、屋頂層、富彥代跑照"],
                  ["P12~13", "公設與庭院",                    "前院、屋頂層、共用空間"],
                  ["P14~15", "戶別格局",                      "4 房範例、平面圖"],
                  ["P16~17", "機能光束圖",                    "26 個點、5 圈"],
                  ["P18", "林先生的一天",                     "故事頁、純文字"],
                  ["P19", "建造、用心",                       "工序、建材、櫻花/TOTO"],
                  ["P20", "下一步｜接待中心資訊",              "聯絡方式、地址、QR"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_page_number(s, 12, TOTAL)

    # ── 12 通路 × 文案配對
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "12　通路 × 文案配對表",
                  "哪個文案、用在哪個通路 — 一目了然")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["通路", "主標", "副標", "Body", "CTA"],
              [
                  ["FB 廣告 (TA1)", "主標 1 (中午回家)", "副標 3 (雪隧)",
                   "—", "了解更多"],
                  ["FB 廣告 (TA2)", "主標 5 (千萬內)",  "副標 2 (學區+坪數)",
                   "—", "查看戶別"],
                  ["IG 主貼", "主標 3 (省下時間)", "副標 1 (通用)",
                   "—", "點 bio 連結"],
                  ["591 案件頁", "tagline 1 (栖息)", "副標 4 (鼎弘信譽)",
                   "TA1 版", "預約看屋"],
                  ["售屋手冊", "tagline 1", "副標 1, 3, 4",
                   "全 3 個 TA 版", "聯絡接待中心"],
                  ["看屋會館海報", "tagline 1", "副標 1",
                   "—", "—（現場接待）"],
                  ["LINE 群組轉發", "—", "副標 6 (VIP 優惠)",
                   "短/中版", "看屋連結"],
                  ["YouTube 影片", "主標 1 / 3", "副標 1",
                   "60 秒腳本", "了解更多"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_page_number(s, 13, TOTAL)

    # ── 13 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "13　下一步行動", "從文案系統到上線投放")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "決策項", "璞域建議", "業主回覆"],
              [
                  ["1", "Tagline 定案", "第一案：在繁華之中找一個栖息", "_______"],
                  ["2", "主標題確認", "推薦 6 個全部使用、輪播", "_______"],
                  ["3", "Body Copy 三版", "TA1/TA2/TA3 各一版皆製作", "_______"],
                  ["4", "售屋手冊製作", "20 頁版本、預估 2 週完成設計", "_______"],
                  ["5", "投放時程", "Q3 2026 上線、配合 Phase 1 暖身期", "_______"],
                  ["6", "代理投放 vs 自投", "建議委由璞域代為投放", "_______"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "本簡報所有文案皆可直接使用 — 業主決定 tagline 後，3 天內可完成全套校稿",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 14, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
