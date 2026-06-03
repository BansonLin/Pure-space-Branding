"""簡報 10：中道森活 — 短影音腳本簡報.

從林先生的一天故事擴寫成 30/60/90 秒三版腳本，
含分鏡、旁白、字幕、B-roll 建議、拍攝計畫。

輸出：02_簡報/10_中道森活_短影音腳本.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "10_中道森活_短影音腳本.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()
    TOTAL = 13

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="短影音腳本簡報　30s / 60s / 90s 三版本",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　三版本概念與用途",
        "02　30 秒｜純情緒版 — 用於 FB / IG 動態",
        "03　30 秒分鏡表",
        "04　60 秒｜故事版 — 用於 YouTube / 591",
        "05　60 秒分鏡表",
        "06　90 秒｜完整版 — 用於主視覺、看屋會館循環播放",
        "07　90 秒分鏡表",
        "08　共通素材清單",
        "09　拍攝計畫與時程",
        "10　預算試算（自製 vs 委製）",
        "11　下一步",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=18, line_space=1.5)

    # ── 01 三版本概念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　三版本概念與用途",
                  "不同長度對應不同觀看情境")

    versions = [
        ("30 秒\n純情緒",   "FB / IG 動態廣告",   "高頻曝光、抓眼球",
         "畫面 + 字幕為主、無旁白\n音樂帶動情緒",        BRAND_YELLOW_SOFT),
        ("60 秒\n故事",     "YouTube 廣告 / 591 內嵌",  "建立認知、引導點擊",
         "林先生 narrator 旁白\n林先生生活畫面 + 機能 B-roll", BRAND_YELLOW),
        ("90 秒\n完整",     "看屋會館循環 / 官網",       "深度說服、轉化下訂",
         "完整故事 + 產品介紹\n含建商資訊與 CTA",          CREAM),
    ]
    bw = 4.0; bh = 5.0; y0 = 1.7
    for i, (label, channel, goal, content, fill) in enumerate(versions):
        x = 0.5 + i * 4.2
        add_rect(s, Inches(x), Inches(y0), Inches(bw), Inches(bh),
                 fill_color=fill, line_color=DARK_BROWN)
        text_color = WHITE if fill == BRAND_YELLOW else DARK_BROWN
        add_text(s, Inches(x), Inches(y0 + 0.2), Inches(bw), Inches(1.1),
                 label, size=28, bold=True, color=text_color, align="center",
                 line_space=1.15, font=FONT_TITLE)
        # divider
        add_rect(s, Inches(x + bw/2 - 0.5), Inches(y0 + 1.45),
                 Inches(1.0), Pt(1), fill_color=text_color)
        add_text(s, Inches(x + 0.2), Inches(y0 + 1.75), Inches(bw - 0.4), Inches(0.45),
                 "通路", size=11, color=text_color, align="center")
        add_text(s, Inches(x + 0.2), Inches(y0 + 2.15), Inches(bw - 0.4), Inches(0.45),
                 channel, size=14, bold=True, color=text_color, align="center")
        add_text(s, Inches(x + 0.2), Inches(y0 + 2.75), Inches(bw - 0.4), Inches(0.45),
                 "目的", size=11, color=text_color, align="center")
        add_text(s, Inches(x + 0.2), Inches(y0 + 3.15), Inches(bw - 0.4), Inches(0.45),
                 goal, size=13, bold=True, color=text_color, align="center")
        add_text(s, Inches(x + 0.2), Inches(y0 + 3.75), Inches(bw - 0.4), Inches(0.45),
                 "手法", size=11, color=text_color, align="center")
        add_text(s, Inches(x + 0.2), Inches(y0 + 4.15), Inches(bw - 0.4), Inches(0.8),
                 content, size=11, color=text_color, align="center", line_space=1.4)

    add_footer(s, "三版本共用同一批素材（拍攝計畫詳第 09 頁）— 製作成本可控")
    add_page_number(s, 2, TOTAL)

    # ── 02 30 秒概念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　30 秒｜純情緒版",
                  "Hook → Insight → CTA｜FB/IG 動態廣告")
    add_text(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(0.4),
             "概念", size=15, bold=True, color=ACCENT_RED)
    add_text(s, Inches(0.5), Inches(2.1), Inches(12.3), Inches(1.5),
             "「在台北 8 年。我們每月房租 30,000、通勤 90 分鐘。」\n"
             "「搬回壯圍後，房貸 30,000、通勤一週兩次。」\n"
             "「省下的時間，是和爸媽、孩子的時間。」\n"
             "→ 中道森活｜千萬內、3 層透天、27 戶｜2026 預售",
             size=14, color=DARK_BROWN, line_space=1.65)

    add_rect(s, Inches(0.5), Inches(4.0), Inches(12.3), Inches(2.8),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.15), Inches(12), Inches(0.4),
             "情緒節奏（3 拍）", size=14, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(0.7), Inches(4.6), Inches(11.9), Inches(2.0), [
        "Hook (0~8s)：「在台北 8 年」 → 揭出 pain point（高房租、長通勤）",
        "Insight (8~22s)：「搬回壯圍後」 → 翻轉場景（房貸相同、通勤縮短）",
        "Reward (22~28s)：「省下的時間」 → 情緒高點（和家人在一起）",
        "CTA (28~30s)：案名 + 售屋資訊 + 看屋連結",
    ], size=12)
    add_page_number(s, 3, TOTAL)

    # ── 03 30 秒分鏡表
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　30 秒分鏡表", "FB / IG 9:16 直式")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["時間", "畫面", "字幕", "音樂"],
              [
                  ["00~04s", "台北捷運擁擠站台、雨天通勤", "「在台北 8 年」", "鋼琴單音、低沉"],
                  ["04~08s", "公寓房門關上、累倒沙發",     "「房租 30,000　通勤 90 分鐘」", "鋼琴持續"],
                  ["08~14s", "鏡頭切換 — 壯圍田野晨光",     "「搬回壯圍後」", "提升、加入弦樂"],
                  ["14~20s", "孩子騎腳踏車、媽媽招手",     "「房貸 30,000　通勤一週兩次」", "弦樂展開"],
                  ["20~26s", "家人圍坐餐桌、爸媽抱孫",     "「省下的時間」", "情緒最高點"],
                  ["26~30s", "中道森活 logo 浮現 + 資訊卡", "中道森活｜千萬內｜2026 預售\n了解更多 →", "音樂淡出"],
              ],
              header_size=11, body_size=10.5, first_col_bold=True, body_align="left")
    add_footer(s, "9:16 直式｜建議解析度 1080×1920｜可加製 1:1 方版供 IG 動態")
    add_page_number(s, 4, TOTAL)

    # ── 04 60 秒概念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　60 秒｜故事版",
                  "林先生的一天｜YouTube 廣告 / 591 內嵌")
    add_text(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(0.4),
             "概念", size=15, bold=True, color=ACCENT_RED)
    add_text(s, Inches(0.5), Inches(2.1), Inches(12.3), Inches(2.5),
             "用「林先生的一天」走完一個完整的循環。\n\n"
             "上半 (0~25s) — 早晨送孩子、通勤台北、與媽媽 LINE。\n"
             "中段 (25~45s) — 傍晚回家、全家散步、週末野餐。\n"
             "下半 (45~60s) — 林先生獨白：「省下的時間」+ 中道森活資訊。",
             size=13, color=DARK_BROWN, line_space=1.6)

    add_rect(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(1.8),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.1), Inches(12), Inches(0.4),
             "敘事設計", size=14, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(0.7), Inches(5.55), Inches(11.9), Inches(1.2), [
        "第一人稱旁白（林先生自述）— 親近感、信任感",
        "每段配機能 B-roll（國道、學校、醫院、景點）",
        "畫面節奏：上半快、中段中、下半慢 — 由急轉緩",
    ], size=11.5)
    add_page_number(s, 5, TOTAL)

    # ── 05 60 秒分鏡表
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　60 秒分鏡表", "YouTube 16:9 橫式")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["時間", "畫面", "旁白 / 字幕"],
              [
                  ["00~05s", "07:30 林先生送孩子到公館國小",
                   "旁白：「孩子的學校，4 分鐘車。」"],
                  ["05~15s", "08:15 林先生上葛瑪蘭、雪隧、台北辦公室",
                   "旁白：「我每週兩次去台北。雪隧 50 分鐘。」"],
                  ["15~25s", "12:00 林先生 LINE 媽媽、媽媽煮飯",
                   "旁白：「中午能回家吃飯。媽媽 10 分鐘車。」"],
                  ["25~35s", "19:00 全家從家走到鄰里公園、買 7-11 冰",
                   "旁白：「傍晚的散步，不需要規劃。」"],
                  ["35~45s", "週六全家壯圍沙丘野餐、看龜山島",
                   "旁白：「週末的海，5 分鐘就到。」"],
                  ["45~55s", "晚上一家圍坐餐桌、夜燈下",
                   "旁白：「在台北 8 年，省下來的時間，現在都還回家裡。」"],
                  ["55~60s", "中道森活 logo + 資訊",
                   "字幕：中道森活｜壯圍 27 戶｜千萬內｜2026 預售"],
              ],
              header_size=12, body_size=10.5, first_col_bold=True, body_align="left")
    add_footer(s, "16:9 橫式｜建議 1920×1080｜可加 9:16 版供 YouTube Shorts")
    add_page_number(s, 6, TOTAL)

    # ── 06 90 秒概念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　90 秒｜完整版",
                  "故事 + 產品｜看屋會館循環 / 官網 / 接待中心")
    add_text(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(0.4),
             "概念", size=15, bold=True, color=ACCENT_RED)
    add_text(s, Inches(0.5), Inches(2.1), Inches(12.3), Inches(3.5),
             "在 60 秒故事版基礎上，加入 30 秒產品介紹段：\n\n"
             "00~60s — 林先生的一天（同 60 秒版）\n\n"
             "60~75s — 中道森活產品介紹\n"
             "  ↳ 空拍：基地俯瞰、7 幢配置、屋頂層庭院\n"
             "  ↳ 字幕：27 戶 透天 3 層｜建坪 38~43 坪\n\n"
             "75~85s — 鼎弘建設信譽\n"
             "  ↳ 玉田森活 1/2 完銷照片、屋主訪談片段\n\n"
             "85~90s — CTA + 接待中心資訊",
             size=12.5, color=DARK_BROWN, line_space=1.55)

    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "適合需要說明的場景 — 看屋會館循環、官網嵌入、接待中心開場片",
             size=12, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 7, TOTAL)

    # ── 07 90 秒分鏡表
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　90 秒分鏡表 — 新增段落",
                  "60~90 秒（前 60 秒同上一頁）")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.5),
              ["時間", "畫面", "旁白 / 字幕"],
              [
                  ["60~67s", "空拍：壯圍五權段、7 幢配置鳥瞰",
                   "字幕：壯圍鄉五權段　27 戶獨棟透天"],
                  ["67~75s", "空拍環繞：A、B、C、D、E、F、G 幢",
                   "字幕：建坪 38~43 坪｜地坪 22~37 坪｜3 層"],
                  ["75~82s", "玉田森活 1 完銷實景 + 屋主家庭照",
                   "旁白：「鼎弘建設，30+14 戶完銷的承諾。」"],
                  ["82~88s", "玉田森活 2 完銷實景 + 鼎弘 logo",
                   "字幕：鼎弘建設 玉田森活 1 / 2 全數完銷"],
                  ["88~90s", "中道森活 logo + 接待中心地址 + QR",
                   "字幕：接待中心｜開放預約｜www.pure-branding.com"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")
    add_footer(s, "用於看屋會館 — 建議 4K 解析度、循環播放、可關閉旁白純看畫面")
    add_page_number(s, 8, TOTAL)

    # ── 08 共通素材清單
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　共通素材清單",
                  "三版本共用 — 拍攝一次、剪三版本")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(5.0),
              ["#", "素材類別", "具體場景", "拍攝注意"],
              [
                  ["1", "林先生家庭",
                   "夫妻、兩個孩子、爸媽（雇用 1 個演員家庭）",
                   "需簽肖像授權"],
                  ["2", "通勤畫面",
                   "雪隧入口、葛瑪蘭客運、宜蘭轉運站",
                   "選平日離峰時段、避免人潮"],
                  ["3", "學校 / 機能",
                   "公館國小門口、7-11、全聯、傳統市場",
                   "不需建築物特寫、避免商家招牌"],
                  ["4", "假日景點",
                   "壯圍沙丘、永鎮海濱、夢時代、羅東夜市",
                   "選晨光 / 黃金時刻拍攝"],
                  ["5", "基地空拍",
                   "壯圍五權段 27 戶配置鳥瞰",
                   "需業主協調、空拍機申請、無人機證照"],
                  ["6", "玉田森活素材",
                   "已有檔案、向業主索取",
                   "若無高解析度，需重新空拍"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_page_number(s, 9, TOTAL)

    # ── 09 拍攝計畫
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　拍攝計畫與時程",
                  "3 天拍攝 + 2 週後製")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.0),
              ["階段", "工項", "天數", "產出"],
              [
                  ["前置", "選角、場景勘景、分鏡 storyboard", "5 天", "完整分鏡稿"],
                  ["拍攝 D1", "林先生家庭日常（早晨、傍晚、週末）", "1 天", "家庭素材"],
                  ["拍攝 D2", "通勤、機能、景點 B-roll", "1 天", "B-roll 素材"],
                  ["拍攝 D3", "基地空拍 + 玉田森活補拍", "1 天", "空拍 + 建築素材"],
                  ["後製", "剪輯 30/60/90 三版", "10 天", "完成片"],
                  ["上線", "FB/IG/YouTube/591/官網", "1 週", "全通路上稿"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "全程約 5 週｜可銜接 Q4 2026 主力期推案前完成",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_page_number(s, 10, TOTAL)

    # ── 10 預算試算
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "10　預算試算",
                  "自製 vs 委製｜建議選「半委製」中間方案")
    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.0),
              ["方案", "預算範圍", "團隊", "適合情境"],
              [
                  ["A. 全自製", "5~10 萬",
                   "業主自拍 + iPhone + 簡易剪輯",
                   "預算極低、可接受平實質感"],
                  ["B. 半委製（建議）", "20~35 萬",
                   "璞域企劃 + 在地小型製作公司",
                   "中等預算、品質可控、彈性大"],
                  ["C. 全委製", "60~100 萬",
                   "知名製作公司、明星演員",
                   "高預算、品牌升級為主"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.5),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(0.4),
             "建議 B 方案：半委製", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "璞域協助：腳本、選角、督導拍攝、後製方向",
        "在地團隊：拍攝、剪輯、執行（成本只有北部 50~60%）",
        "預估 25 萬 → 30 秒×3 + 60 秒×1 + 90 秒×1 全套交付",
    ], size=11.5)

    add_page_number(s, 11, TOTAL)

    # ── 11 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "11　下一步行動", "從本簡報到上線播放")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "決策項", "璞域建議", "業主回覆"],
              [
                  ["1", "三版本要做幾版", "推薦三版全做（25 萬）", "_______"],
                  ["2", "預算方案", "B 方案（半委製）", "_______"],
                  ["3", "選角方向", "真實素人家庭、非演員", "_______"],
                  ["4", "拍攝時程", "8 月選角、9 月拍攝、10 月上線", "_______"],
                  ["5", "首發通路", "FB 廣告 + 591 內嵌 + YouTube", "_______"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "確認後璞域提供：完整分鏡 storyboard、選角資料、製作公司提案",
             size=12, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 12, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
