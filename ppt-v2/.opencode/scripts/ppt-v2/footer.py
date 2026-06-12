from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import theme


def add_footer(slide, page_num=None, total=None):
    left = theme.MARGIN_L
    top = theme.FOOTER_Y
    width = theme.CONTENT_W
    height = theme.FOOTER_H

    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)

    p = tf.paragraphs[0]
    p.font.size = Pt(8)
    p.font.color.rgb = theme.MEDIUM_RGB
    p.font.name = theme.FONT_FAMILY
    p.alignment = PP_ALIGN.LEFT

    left_part = "C2 – Usage restreint"
    if page_num is not None and total is not None:
        right_part = f"{page_num} / {total}"
    elif page_num is not None:
        right_part = str(page_num)
    else:
        right_part = ""

    p.text = left_part

    line = tf.add_paragraph()
    line.font.size = Pt(8)
    line.font.color.rgb = theme.MEDIUM_RGB
    line.font.name = theme.FONT_FAMILY
    line.alignment = PP_ALIGN.RIGHT
    line.text = right_part

    sep = slide.shapes.add_shape(
        1, left, int(theme.FOOTER_Y - Inches(0.02)),
        width, Pt(0.5)
    )
    sep.fill.solid()
    sep.fill.fore_color.rgb = theme.BORDER_RGB
    sep.line.fill.background()
