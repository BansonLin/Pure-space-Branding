"""簡報 5：給宗硯輔銷的資料請求清單.

對象：宗硯建設的輔銷企劃（璞域既有客戶）
目的：取得宗硯大砌系列18-透天的真實銷售數據，作為中道森活｜栖定價精校的依據。

輸出：02_簡報/05_給宗硯輔銷_資料請求清單簡報.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "05_給宗硯輔銷_資料請求清單簡報.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)
TOTAL = 11


def build():
    prs = new_presentation()

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="璞域品牌策略 → 宗硯輔銷",
        title_main="資料請求清單",
        subtitle="協助鼎弘建設新案「中道森活」精校定價",
        footer_left="壯圍鄉同營造廠｜雙案協同",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　請求背景：為何兩案需要資料協同",
        "02　請求對象與保密前提",
        "03　請求項 1 — 成交去化曲線",
        "04　請求項 2 — 來客客層輪廓",
        "05　請求項 3 — 看屋熱區戶別分布",
        "06　請求項 4 — 議價底線與付款條件",
        "07　請求項 5 — 媒體效果（選擇性）",
        "08　交付格式與時程",
        "09　雙方互惠：璞域可回給宗硯什麼",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ── 1. 請求背景
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　請求背景", "兩案的關係結構與協同的必要性")

    # 關係圖
    add_rect(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(2.2),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(1.85), Inches(12), Inches(0.5),
             "關係結構", size=15, bold=True, color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(2.3), Inches(12), Inches(1.5),
             "鼎弘建設 ──業主社群── 宗硯建設\n"
             "    │                      │\n"
             "中道森活｜栖（本案，27戶）  ◄═同營造廠═►  宗硯大砌系列18-透天（31戶）\n"
             "    │                      │\n"
             "璞域品牌策略  ◄══輔銷顧問══►  宗硯輔銷企劃",
             size=13, color=DARK_BROWN, line_space=1.4, font="Courier New")

    add_rect(s, Inches(0.5), Inches(4.1), Inches(12.3), Inches(2.7),
             fill_color=CREAM, line_color=ACCENT_RED)
    add_text(s, Inches(0.7), Inches(4.25), Inches(12), Inches(0.5),
             "為什麼要協同", size=15, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(4.7), Inches(11.9), Inches(2.0), [
        "兩案同營造廠、同價帶（21 萬/坪上下）、同客層 — 不協同會自相打架",
        "兩案合計 58 戶 ≈ 壯圍 24 個月透天市場吸納量 29% — 同期推案有供給壓力",
        "宗硯大砌已開賣 4 個月（2026.02 起），有真實去化數據；中道森活預計 Q4 上市 — **時間差優勢**：可用宗硯實戰數據精校定價",
        "璞域同時服務兩案：對宗硯有「定價對齊不被低價案打」的價值",
    ], size=12.5)
    add_footer(s, "本資料請求清單為雙方互惠協同的工具 — 非單向索取")
    add_page_number(s, 2, TOTAL)

    # ── 2. 對象與保密
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　請求對象與保密前提", "資料邊界與使用範圍")

    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.5),
             "對象", size=18, bold=True, color=DARK_BROWN, align="center")
    add_bullets(s, Inches(0.8), Inches(2.45), Inches(5.4), Inches(4.0), [
        "宗硯建設輔銷企劃部門",
        "璞域既有合作關係",
        "經宗硯內部許可後提供",
        "聯絡窗口：（業主確認後填入）",
    ], size=13, line_space=1.4)

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=CREAM, line_color=ACCENT_RED)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.5),
             "保密前提", size=18, bold=True, color=ACCENT_RED, align="center")
    add_bullets(s, Inches(7.2), Inches(2.45), Inches(5.4), Inches(4.0), [
        "資料僅供中道森活定價校準",
        "璞域內部最多 2 人接觸",
        "成果不對第三方披露",
        "不轉述、不二次散布",
        "全部資料於本案上市後 6 個月內銷毀",
        "簽訂保密附約後生效",
    ], size=13, line_space=1.4, bullet_color=ACCENT_RED)

    add_footer(s, "保密附約建議：補簽於既有輔銷合約之 NDA 附件")
    add_page_number(s, 3, TOTAL)

    # ── 3. 請求項 1
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　請求項 1 — 成交去化曲線",
                  "宗硯大砌系列18 自開賣至今的每月新增成交數")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(2.4),
              ["欄位", "格式", "用途"],
              [
                  ["月份", "YYYY-MM（2026-02 起）", "看去化節奏"],
                  ["當月新增成交戶", "整數", "看是否衝刺、是否疲軟"],
                  ["累積成交戶", "整數（最終 ≤ 31）", "看完銷進度"],
                  ["當月來客數", "整數（如有）", "看廣告效果"],
                  ["當月成交/來客 (%)", "百分比", "看轉換率"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_rect(s, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.2),
             fill_color=CREAM, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(4.65), Inches(12), Inches(0.5),
             "對中道森活的價值", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.15), Inches(11.9), Inches(1.5), [
        "知道宗硯實際多快被消化 → 預測壯圍同價帶 6 個月吸納量",
        "判斷中道森活 Q4 上市時，市場是否還有需求殘量",
        "若宗硯前 4 個月已賣 60%（=18戶以上）→ 強市場訊號，可拉中道單價",
        "若宗硯前 4 個月只賣 30%（=9戶以下）→ 弱市場訊號，必須對齊更低單價",
    ], size=12)
    add_footer(s, "格式建議：Excel 或 csv；單月一列；至少含開賣以來逐月資料")
    add_page_number(s, 4, TOTAL)

    # ── 4. 請求項 2
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　請求項 2 — 來客客層輪廓",
                  "宗硯大砌系列18 真實買家的人口輪廓統計")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.5),
              ["客層維度", "分組方式", "用途"],
              [
                  ["年齡帶", "30-、30~40、40~50、50+", "對齊 TA1/TA2/TA3 比例"],
                  ["居住地", "宜蘭、台北/新北、其他", "返鄉客 vs 在地客比例"],
                  ["家庭結構", "首購夫妻、有子女、退休", "格局偏好驗證"],
                  ["購屋動機", "自住、置產、子女、退休回鄉", "賣點調整方向"],
                  ["原住型態", "公寓、老透天、租屋、其他",
                   "換屋客 vs 首購客比例"],
                  ["購買頻率", "首購 / 第二房 / 投資", "預算敏感度"],
                  ["付款方式", "自備款比例、貸款成數", "金流壓力測試"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.5),
             "用途：直接校準客群定位簡報 TA1/TA2/TA3 的 45% / 35% / 20% 比例假設",
             size=12, color=ACCENT_RED, bold=True)
    add_footer(s, "格式建議：每維度一張小表，匿名統計即可，不需個資")
    add_page_number(s, 5, TOTAL)

    # ── 5. 請求項 3
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　請求項 3 — 看屋熱區戶別分布",
                  "宗硯大砌系列18 哪幾戶先賣完")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(3.5),
              ["欄位", "格式", "用途"],
              [
                  ["戶別代號", "A1、B2、… 或宗硯編碼",
                   "對應戶別位置"],
                  ["地坪 / 建坪", "坪", "驗證大地坪是否真有溢價"],
                  ["定價時單價", "萬/坪", "比對戶別差異化定價合理性"],
                  ["成交順序", "1, 2, 3, …", "找出熱門戶別共通特徵"],
                  ["最終成交價", "萬", "對比定價，看議價空間"],
                  ["備註", "客戶特徵或交易特殊狀況", "（選填）"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_rect(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(1.3),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.7), Inches(12), Inches(0.4),
             "對中道森活的價值", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(6.1), Inches(11.9), Inches(0.7), [
        "反推中道森活 27 戶哪些戶位最有競爭力（角間、雙面採光、屋頂層）",
        "依此調整 27 戶單戶定價試算表的「群組分級」",
    ], size=12)
    add_footer(s, "格式建議：每戶一列，可匿去客戶資訊但留戶別/區位")
    add_page_number(s, 6, TOTAL)

    # ── 6. 請求項 4
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　請求項 4 — 議價底線與付款條件",
                  "宗硯大砌系列18 真實接受的議價範圍")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(3.5),
              ["欄位", "範圍", "用途"],
              [
                  ["議價幅度", "牌價 vs 成交價的 %",
                   "知道實際成交相對牌價打多少折"],
                  ["前 5 戶讓利幅度", "比例 / 萬",
                   "驗證「前期讓利衝量」策略效果"],
                  ["後段戶有無提價", "%",
                   "驗證「後段溢價」策略"],
                  ["付款條件", "自備、貸款、分期",
                   "看客戶能負擔的付款方式"],
                  ["送禮 / 加贈", "家電、裝潢、車位",
                   "看是否需配套包套"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.5),
             "用途：中道森活的議價底線設定 + 看屋會館的優惠話術設計",
             size=12, color=ACCENT_RED, bold=True)
    add_text(s, Inches(0.5), Inches(6.15), Inches(12.3), Inches(0.5),
             "建議揭露程度：給「區間 + 平均」即可，不需逐戶議價細節",
             size=11.5, color=GRAY_70)
    add_footer(s, "敏感度最高的請求項 — 務必經宗硯內部許可")
    add_page_number(s, 7, TOTAL)

    # ── 7. 請求項 5 (選擇性)
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　請求項 5 — 媒體效果（選擇性）",
                  "宗硯大砌系列18 廣告投放效果回饋")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(3.5),
              ["欄位", "用途"],
              [
                  ["廣告通路與佔比", "看 591、地方廣告、戶外、社群比例"],
                  ["每通路來客單成本", "知道哪個通路 CPL 最低"],
                  ["每通路成交轉換率", "知道哪個通路客質最高"],
                  ["主視覺曝光效果", "若有 A/B test 結果可參考"],
                  ["最有效的賣點 message", "用於中道森活廣告文案"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.5),
             "標記為選擇性 — 若宗硯認為涉及代銷祕辛，可暫不提供",
             size=12, color=GRAY_70)
    add_footer(s, "互惠版本：璞域可回給宗硯中道森活的同類數據作為交換")
    add_page_number(s, 8, TOTAL)

    # ── 8. 交付格式與時程
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　交付格式與時程",
                  "璞域接收後 5 個工作天內提供整合分析")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["階段", "時程", "工項"],
              [
                  ["W0", "本簡報遞交", "宗硯內部評估、簽 NDA"],
                  ["W1", "請求項 1+2+3 交付", "去化曲線、客層、熱區（核心 3 項）"],
                  ["W2", "請求項 4 交付（選擇性）", "議價底線（最敏感）"],
                  ["W2", "請求項 5 交付（選擇性）", "媒體效果（選擇性）"],
                  ["W3", "璞域產出整合分析", "對中道森活的定價建議調整"],
                  ["W4", "璞域回饋宗硯", "中道森活相對數據（互惠）"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "全程約 4 週｜核心 3 項在 W1 完成已足以做關鍵調整",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_footer(s, "PDF / Excel / csv 皆可接收，視宗硯方便")
    add_page_number(s, 9, TOTAL)

    # ── 9. 互惠
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "09　雙方互惠 — 璞域可回給宗硯什麼",
                  "不是單向索取，是雙向協同")
    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["互惠項", "內容", "對宗硯的價值"],
              [
                  ["中道森活開賣後同類數據",
                   "去化曲線、客層、熱區（W12+ 交付）",
                   "看自己銷售相對位置"],
                  ["壯圍區整體市場分析",
                   "璞域已產出之 24M 實登分析（v0.3 報告）",
                   "市場全景驗證"],
                  ["共同避險的定價戰略",
                   "兩案不打價格戰、不重疊期推案",
                   "宗硯後段戶不被低價案打"],
                  ["共用接待中心廣告版位",
                   "視兩案是否同時段運作而定",
                   "省媒體預算 30~50%"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
             "立場：璞域同時服務兩案，目標是兩案都好賣 — 不是讓中道贏宗硯",
             size=12.5, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 10, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
