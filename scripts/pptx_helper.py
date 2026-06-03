"""中道森活 簡報共通版型 helper.

主視覺色系（取自業主主視覺提案）：
  - 黃色背景：#E8AC1B (saffron)
  - 米白：#FFF8E7
  - 深棕：#4A3210 (標題)
  - 暗紅：#8B3A0F (重點)
  - 灰：#666666 (內文)

字型：
  - 標題：思源宋體 / Noto Serif TC / Yu Mincho — pptx 宣告為「Microsoft JhengHei」+ 西文「Garamond」備援
  - 內文：「Microsoft JhengHei」(微軟正黑體)
"""
from __future__ import annotations
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# Brand palette
BRAND_YELLOW = RGBColor(0xE8, 0xAC, 0x1B)   # 主視覺金黃
BRAND_YELLOW_SOFT = RGBColor(0xF5, 0xCB, 0x4F)
CREAM = RGBColor(0xFF, 0xF8, 0xE7)
DARK_BROWN = RGBColor(0x4A, 0x32, 0x10)
ACCENT_RED = RGBColor(0x8B, 0x3A, 0x0F)
GRAY_70 = RGBColor(0x55, 0x55, 0x55)
GRAY_40 = RGBColor(0xA0, 0xA0, 0xA0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1F, 0x1F, 0x1F)

# 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT_TITLE = "Microsoft JhengHei"  # 正黑體
FONT_BODY = "Microsoft JhengHei"
FONT_SERIF = "Microsoft JhengHei"

BRAND_URL = "www.pure-branding.com"
BRAND_NAME = "璞域品牌策略"


def new_presentation() -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def add_blank_slide(prs: Presentation):
    blank = prs.slide_layouts[6]
    return prs.slides.add_slide(blank)


def set_background(slide, color: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, fill_color: RGBColor | None = None, line_color: RGBColor | None = None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill_color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.75)
    shape.shadow.inherit = False
    return shape


def add_text(slide, x, y, w, h, text: str, *,
             size: int = 18, bold: bool = False, color: RGBColor = DARK_BROWN,
             font: str = FONT_BODY, align: str = "left",
             anchor: str = "top", line_space: float = 1.2):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    if anchor == "middle":
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    elif anchor == "bottom":
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if align == "center":
            p.alignment = PP_ALIGN.CENTER
        elif align == "right":
            p.alignment = PP_ALIGN.RIGHT
        elif align == "justify":
            p.alignment = PP_ALIGN.JUSTIFY
        p.line_spacing = line_space
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        # 中文需要設定 eastasia font
        _set_east_asia_font(r, font)
    return tb


def _set_east_asia_font(run, font_name: str):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", font_name)


def add_bullets(slide, x, y, w, h, items: list[str], *,
                size: int = 16, color: RGBColor = DARK_BROWN,
                bullet_color: RGBColor | None = None, line_space: float = 1.35):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    bc = bullet_color or ACCENT_RED
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line_space
        r1 = p.add_run()
        r1.text = "• "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(size)
        r1.font.color.rgb = bc
        r1.font.bold = True
        _set_east_asia_font(r1, FONT_BODY)
        r2 = p.add_run()
        r2.text = item
        r2.font.name = FONT_BODY
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
        _set_east_asia_font(r2, FONT_BODY)
    return tb


def add_table(slide, x, y, w, h, headers: list[str], rows: list[list[str]], *,
              header_fill=DARK_BROWN, header_color=WHITE,
              row_fill=CREAM, alt_row_fill=WHITE,
              header_size=12, body_size=11, font=FONT_BODY,
              first_col_bold=False, highlight_rows: list[int] | None = None,
              body_align: str = "center"):
    nrows = len(rows) + 1
    ncols = len(headers)
    table = slide.shapes.add_table(nrows, ncols, x, y, w, h).table

    # header
    for j, h_text in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        tf = cell.text_frame
        tf.margin_left = Pt(4); tf.margin_right = Pt(4)
        tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h_text
        r.font.name = font; r.font.size = Pt(header_size); r.font.bold = True
        r.font.color.rgb = header_color
        _set_east_asia_font(r, font)

    # body
    highlight_rows = highlight_rows or []
    for i, row in enumerate(rows):
        fill = row_fill if i % 2 == 0 else alt_row_fill
        if i in highlight_rows:
            fill = BRAND_YELLOW_SOFT
        for j, cell_val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            tf = cell.text_frame
            tf.margin_left = Pt(4); tf.margin_right = Pt(4)
            tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
            p = tf.paragraphs[0]
            if j == 0:
                p.alignment = PP_ALIGN.LEFT
            elif body_align == "left":
                p.alignment = PP_ALIGN.LEFT
            else:
                p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = str(cell_val)
            r.font.name = font; r.font.size = Pt(body_size)
            r.font.bold = (j == 0 and first_col_bold) or i in highlight_rows
            r.font.color.rgb = DARK_BROWN
            _set_east_asia_font(r, font)
    return table


def add_page_number(slide, n: int, total: int, color: RGBColor = GRAY_40):
    add_text(slide, Inches(12.5), Inches(7.0), Inches(0.7), Inches(0.3),
             f"{n} / {total}", size=10, color=color, align="right")


def add_footer(slide, text: str, color: RGBColor = GRAY_40):
    add_text(slide, Inches(0.5), Inches(7.0), Inches(8.0), Inches(0.3),
             text, size=9, color=color)


def add_title_bar(slide, title: str, subtitle: str | None = None, *,
                  bar_color: RGBColor = BRAND_YELLOW):
    """Top brand bar + slide title."""
    # Top thin color bar
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.18), fill_color=bar_color)
    # Title block
    add_text(slide, Inches(0.5), Inches(0.35), Inches(12.0), Inches(0.7),
             title, size=24, bold=True, color=DARK_BROWN, font=FONT_TITLE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.95), Inches(12.0), Inches(0.4),
                 subtitle, size=13, color=GRAY_70)
    # Title underline (thin gold rule)
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(0.8), Pt(2.5),
             fill_color=ACCENT_RED)


def add_cover(slide, title_top: str, title_main: str, subtitle: str,
              footer_left: str, footer_right: str):
    set_background(slide, BRAND_YELLOW)
    # decorative corner rect (mimic flower frame)
    add_rect(slide, 0, 0, Inches(0.6), SLIDE_H, fill_color=DARK_BROWN)
    add_rect(slide, SLIDE_W - Inches(0.15), 0, Inches(0.15), SLIDE_H, fill_color=DARK_BROWN)
    # main title block
    add_text(slide, Inches(1.5), Inches(2.3), Inches(10), Inches(0.5),
             title_top, size=18, color=DARK_BROWN, align="center",
             font=FONT_BODY)
    add_text(slide, Inches(1.5), Inches(2.9), Inches(10), Inches(1.4),
             title_main, size=64, bold=True, color=DARK_BROWN, align="center",
             font=FONT_TITLE)
    add_text(slide, Inches(1.5), Inches(4.4), Inches(10), Inches(0.6),
             subtitle, size=20, color=DARK_BROWN, align="center")
    # rule
    add_rect(slide, Inches(6.17), Inches(5.2), Inches(1.0), Pt(1.5), fill_color=DARK_BROWN)
    # footer
    add_text(slide, Inches(0.9), Inches(6.7), Inches(6), Inches(0.4),
             footer_left, size=12, color=DARK_BROWN)
    add_text(slide, Inches(7), Inches(6.7), Inches(5.5), Inches(0.4),
             footer_right, size=12, color=DARK_BROWN, align="right")


def add_section_divider(slide, label_top: str, label_main: str):
    set_background(slide, BRAND_YELLOW)
    add_rect(slide, 0, Inches(3.0), SLIDE_W, Inches(1.5), fill_color=CREAM)
    add_text(slide, Inches(0.5), Inches(3.1), Inches(12.3), Inches(0.5),
             label_top, size=14, color=GRAY_70, align="center")
    add_text(slide, Inches(0.5), Inches(3.55), Inches(12.3), Inches(1.0),
             label_main, size=44, bold=True, color=DARK_BROWN, align="center",
             font=FONT_TITLE)
