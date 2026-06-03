"""簡報 9：中道森活 — 機能地圖海報設計簡報.

輸出：02_簡報/09_中道森活_機能地圖海報設計.pptx
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pptx_helper import *
from pptx.util import Inches, Pt

REPO = Path("/home/user/Pure-space-Branding")
OUT = REPO / "02_簡報" / "09_中道森活_機能地圖海報設計.pptx"
POSTER = REPO / "assets/posters/中道森活_機能地圖海報_A2.png"
OUT.parent.mkdir(parents=True, exist_ok=True)


def build():
    prs = new_presentation()
    TOTAL = 11

    # ── 封面
    s = add_blank_slide(prs)
    add_cover(s,
        title_top="鼎弘建設 × 璞域品牌策略",
        title_main="中 道 森 活",
        subtitle="機能地圖海報設計簡報　A2 印刷規格",
        footer_left="壯圍鄉五權段｜27 戶｜透天 3 層",
        footer_right="2026.06")

    # ── 目錄
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "簡報目錄", "AGENDA")
    items = [
        "01　設計理念：為什麼要做這張海報",
        "02　海報全圖 (A2 直式)",
        "03　分區解讀 1 — 上區：品牌標題",
        "04　分區解讀 2 — 中區：機能光束圖",
        "05　分區解讀 3 — 下區：故事時間軸",
        "06　印刷規格與材質建議",
        "07　看屋會館安裝建議",
        "08　同系列延伸 — A3、A4、數位版",
        "09　下一步行動",
    ]
    add_bullets(s, Inches(1.0), Inches(1.8), Inches(11), Inches(5),
                items, size=20, line_space=1.5)

    # ── 01 設計理念
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "01　設計理念", "一張海報抵 30 分鐘簡介")

    add_rect(s, Inches(0.5), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=CREAM, line_color=ACCENT_RED)
    add_text(s, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.5),
             "海報要解決的問題", size=18, bold=True, color=ACCENT_RED, align="center")
    add_bullets(s, Inches(0.8), Inches(2.45), Inches(5.4), Inches(4.0), [
        "客戶走進看屋會館 5 秒內要知道：本案在哪裡",
        "客戶 30 秒內要看完：周邊機能",
        "客戶 1 分鐘內要被打中：為什麼這裡值得買",
        "客戶離開時要記得：能講給家人聽的故事",
    ], size=13, line_space=1.5, bullet_color=ACCENT_RED)

    add_rect(s, Inches(6.9), Inches(1.7), Inches(5.9), Inches(5.0),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.5),
             "三區設計策略", size=18, bold=True, color=DARK_BROWN, align="center")
    add_bullets(s, Inches(7.2), Inches(2.45), Inches(5.4), Inches(4.0), [
        "上區（品牌）：1 秒辨識 — 案名、slogan、視覺氛圍",
        "中區（機能）：30 秒掃描 — 光束圖讓眼睛自動找重點",
        "下區（故事）：1 分鐘共鳴 — 林先生時間軸，從理性走到感性",
        "整張海報的閱讀動線：上 → 中 → 下，與客戶決策流程一致",
    ], size=13, line_space=1.5)

    add_footer(s, "海報為實體看屋會館主視覺輔助 — 不取代簡報，但能延伸記憶")
    add_page_number(s, 2, TOTAL)

    # ── 02 海報全圖
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "02　海報全圖", "A2 直式 (420 × 594 mm @ 150 dpi)")

    # 整張海報置中縮放展示 — A2 比例 1:1.414
    # 16:9 slide 13.33" × 7.5" — 預留說明欄
    poster_h = 6.5  # inches
    poster_w = poster_h / 1.414
    poster_x = 1.0
    poster_y = 1.0
    slide.shapes if False else None
    s.shapes.add_picture(str(POSTER), Inches(poster_x), Inches(poster_y),
                          height=Inches(poster_h))

    # 右側說明
    info_x = 6.0
    add_text(s, Inches(info_x), Inches(1.4), Inches(7), Inches(0.5),
             "海報三區架構", size=18, bold=True, color=DARK_BROWN)

    sections = [
        ("上區", "品牌標題", "中道森活 + slogan + 副標"),
        ("中區", "機能光束圖", "26 個機能點 / 5 圈距離 / 8 方位"),
        ("下區", "故事時間軸", "林先生的一天 + 鼎弘建設識別"),
    ]
    for i, (zone, name, desc) in enumerate(sections):
        y = 2.1 + i * 1.5
        add_rect(s, Inches(info_x), Inches(y), Inches(1.0), Inches(1.2),
                 fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
        add_text(s, Inches(info_x), Inches(y + 0.4), Inches(1.0), Inches(0.4),
                 zone, size=16, bold=True, color=DARK_BROWN, align="center")
        add_text(s, Inches(info_x + 1.1), Inches(y + 0.1), Inches(6.2), Inches(0.5),
                 name, size=15, bold=True, color=ACCENT_RED)
        add_text(s, Inches(info_x + 1.1), Inches(y + 0.55), Inches(6.2), Inches(0.7),
                 desc, size=12, color=DARK_BROWN, line_space=1.4)

    add_footer(s, "完整高解析度 PNG/PDF：assets/posters/中道森活_機能地圖海報_A2.{png,pdf}")
    add_page_number(s, 3, TOTAL)

    # ── 03 上區
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "03　分區解讀 1 — 上區", "品牌標題｜1 秒辨識")
    # 海報上半段
    s.shapes.add_picture(str(POSTER),
                          Inches(0.5), Inches(1.7),
                          height=Inches(4.0))
    # 局部用裁圖 — 直接放整張縮小
    # 旁邊說明
    add_text(s, Inches(5.5), Inches(1.7), Inches(7.5), Inches(0.5),
             "視覺元素", size=16, bold=True, color=DARK_BROWN)
    add_table(s, Inches(5.5), Inches(2.2), Inches(7.5), Inches(3.5),
              ["元素", "規格", "目的"],
              [
                  ["主標題 中道森活", "宋體 / 220pt", "品牌辨識"],
                  ["slogan 理想居所最懂生活", "黑體 / 80pt", "情緒引導"],
                  ["副標 在繁華...遇見家", "黑體 / 48pt", "完整 narrative"],
                  ["金黃上下邊條", "高度 36px", "色彩識別"],
              ],
              header_size=12, body_size=11, first_col_bold=True, body_align="left")

    add_text(s, Inches(5.5), Inches(6.0), Inches(7.5), Inches(0.5),
             "為何用宋體？", size=13, bold=True, color=ACCENT_RED)
    add_text(s, Inches(5.5), Inches(6.45), Inches(7.5), Inches(0.5),
             "與業主主視覺花卉插畫氣質一致 — 東方典雅、避免商業感",
             size=12, color=GRAY_70)
    add_page_number(s, 4, TOTAL)

    # ── 04 中區
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "04　分區解讀 2 — 中區", "機能光束圖｜30 秒掃描")
    s.shapes.add_picture(str(POSTER),
                          Inches(0.5), Inches(1.7),
                          height=Inches(5.0))

    add_text(s, Inches(5.5), Inches(1.7), Inches(7.5), Inches(0.5),
             "資訊密度設計", size=16, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(5.5), Inches(2.2), Inches(7.5), Inches(4.5), [
        "中心點：基地（紅方塊 + 中道森活）",
        "5 圈距離環：1 / 3 / 5 / 10 / 30 km",
        "8 方位線：N / NE / E / SE / S / SW / W / NW",
        "26 個機能點：5 種顏色分類",
        "  ↳ 學區（暗紅）、購物（深棕）、醫療（綠）、景點（橙）、交通（紫）",
        "圖例固定在左下角 — 不影響中心構圖",
        "「關鍵車程」4 個數字在右下 — 視覺重心穩定",
    ], size=12, line_space=1.5)
    add_page_number(s, 5, TOTAL)

    # ── 05 下區
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "05　分區解讀 3 — 下區", "故事時間軸｜1 分鐘共鳴")
    s.shapes.add_picture(str(POSTER),
                          Inches(0.5), Inches(1.7),
                          height=Inches(5.0))

    add_text(s, Inches(5.5), Inches(1.7), Inches(7.5), Inches(0.5),
             "故事設計邏輯", size=16, bold=True, color=DARK_BROWN)
    add_bullets(s, Inches(5.5), Inches(2.2), Inches(7.5), Inches(4.5), [
        "6 個時間點 × 一句話 — 不講道理、講畫面",
        "排版 3 欄 × 2 列 — 視覺平衡、不擁擠",
        "每段都對應上方光束圖的機能 — 圖文互相驗證",
        "刻意省略：價格、坪數、戶數 — 「先打感性、再講理性」",
        "底部「鼎弘建設」+ 「璞域」識別 — 信任背書",
    ], size=12, line_space=1.5)

    add_rect(s, Inches(5.5), Inches(5.8), Inches(7.5), Inches(0.9),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(5.7), Inches(5.92), Inches(7.3), Inches(0.65),
             "客戶離開後能記得的不會是「壯圍鄉五權段」，會是「林先生中午能回家吃飯」。",
             size=12, color=DARK_BROWN, bold=True, line_space=1.4)
    add_page_number(s, 6, TOTAL)

    # ── 06 印刷規格
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "06　印刷規格與材質建議", "從紙張到亮膜的選擇邏輯")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.0),
              ["項目", "規格", "成本（單張）", "適用情境"],
              [
                  ["主檔尺寸", "A2 直式 420×594 mm", "—", "標準展示用"],
                  ["紙張 (推薦)", "250g 雪面銅版", "60~120 元", "看屋會館長期展示"],
                  ["紙張 (預算版)", "150g 道林紙", "30~60 元", "DM 隨身發放"],
                  ["護膜", "亞膜（霧面）", "+30%", "防指紋、典雅感"],
                  ["輸出方式", "數位輸出 / 大圖輸出", "店家代工", "彩雷或膠印皆可"],
                  ["檔案交付", "PDF 300dpi + PNG 150dpi", "—", "印刷廠 / 數位皆涵蓋"],
                  ["顏色管理", "CMYK + 出血 3mm", "—", "印刷品質確保"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")

    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "建議：先印 5 張 A2 雪銅 + 亞膜當看屋會館展示，再評估是否量產",
             size=13, color=ACCENT_RED, bold=True, align="center")
    add_footer(s, "印刷估價以宜蘭在地印刷廠為準（建議：宜蘭快速印刷、明傑數位）")
    add_page_number(s, 7, TOTAL)

    # ── 07 安裝建議
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "07　看屋會館安裝建議", "海報該掛在哪、怎麼看")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(3.5),
              ["位置", "建議", "視覺距離", "目的"],
              [
                  ["入口正面牆", "立刻看到", "3~5 m", "第一印象：在哪、長怎樣"],
                  ["接待桌後方", "客戶坐下時面對", "1.5~2 m", "客戶等待時自然瀏覽"],
                  ["實品屋走道", "從房間出來經過", "2~3 m", "看完格局後串連故事"],
                  ["貴賓室牆面", "TA3 客戶接待", "1.5~2 m", "深度討論輔助"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")

    add_rect(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(1.6),
             fill_color=BRAND_YELLOW_SOFT, line_color=DARK_BROWN)
    add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(0.4),
             "燈光與裝裱建議", size=14, bold=True, color=ACCENT_RED)
    add_bullets(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(1.0), [
        "裝裱：木框 + 玻璃（建議深棕 / 米白框）",
        "燈光：3000K 暖光｜照度 300~500 lux｜避免反光",
        "視線高度：海報中心離地 1.4~1.6 m（成人平視）",
    ], size=12)
    add_page_number(s, 8, TOTAL)

    # ── 08 系列延伸
    s = add_blank_slide(prs); set_background(s, WHITE)
    add_title_bar(s, "08　同系列延伸",
                  "從 A2 海報 → A3 DM → A4 文宣 → 數位版")

    add_table(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.0),
              ["版本", "尺寸", "用途", "差異"],
              [
                  ["A2 主視覺", "420 × 594 mm", "看屋會館主牆",     "完整 3 區"],
                  ["A3 售屋手冊封面", "297 × 420 mm", "現場帶走 DM",   "簡化光束圖"],
                  ["A4 文宣單張", "210 × 297 mm", "夾報 / 桌上",     "故事為主、機能簡化"],
                  ["IG 直式 (4:5)", "1080 × 1350 px", "社群投放",      "光束圖簡化版"],
                  ["FB 橫式 (16:9)", "1920 × 1080 px", "FB 廣告主視覺", "故事為主"],
                  ["591 首圖", "1920 × 1200 px", "591 案件列表",    "簡化光束圖"],
              ],
              header_size=12, body_size=11.5, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "六個延伸版本共用「機能光束圖」核心視覺 — 確保品牌一致",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_page_number(s, 9, TOTAL)

    # ── 09 下一步
    s = add_blank_slide(prs); set_background(s, CREAM)
    add_title_bar(s, "09　下一步行動", "兩週內可完成海報量產")

    add_table(s, Inches(0.5), Inches(1.8), Inches(12.3), Inches(4.0),
              ["#", "工項", "璞域提供", "業主提供"],
              [
                  ["1", "海報設計微調", "依踏勘修正、調整字體版面", "回饋方向"],
                  ["2", "印刷廠詢價",   "3 家比價",                  "預算上限"],
                  ["3", "首批 5 張 A2", "代為印刷下單",              "—"],
                  ["4", "看屋會館安裝", "陪同督導（半天）",          "現場、木框、燈光"],
                  ["5", "A3 / A4 延伸", "完成 DM 與文宣設計",        "—"],
                  ["6", "數位版（IG/FB/591）", "適配尺寸、上傳上稿",  "—"],
              ],
              header_size=12, body_size=12, first_col_bold=True, body_align="left")
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.5),
             "建議排程：本週確認設計 → 下週印刷 → 第三週安裝",
             size=13, color=DARK_BROWN, bold=True, align="center")
    add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
             f"{BRAND_NAME}｜{BRAND_URL}", size=11, color=GRAY_70, align="right")
    add_page_number(s, 10, TOTAL)

    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
