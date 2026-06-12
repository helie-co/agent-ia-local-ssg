from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import theme
import icons


def add_card(slide, left, top, width, height, icon_name, title, body_text,
             icon_size=Inches(0.4)):
    bg = slide.shapes.add_shape(
        1, left, top, width, height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = theme.LIGHT_BG_RGB
    bg.line.color.rgb = theme.BORDER_RGB
    bg.line.width = Pt(0.5)
    bg.shadow.inherit = False

    icon_path = icons.get_icon_path(icon_name) if icon_name else None
    icon_top = top + Inches(0.15)
    icon_left = left + Inches(0.2)
    if icon_path:
        try:
            slide.shapes.add_picture(icon_path, icon_left, icon_top,
                                     icon_size, icon_size)
        except Exception:
            pass

    title_left = left + Inches(0.2)
    if icon_path:
        title_left = icon_left + icon_size + Inches(0.1)
        title_w = width - icon_size - Inches(0.4)
    else:
        title_w = width - Inches(0.4)

    title_box = slide.shapes.add_textbox(
        title_left, icon_top, title_w, Inches(0.35)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = theme.DARK_RGB
    p.font.name = theme.FONT_FAMILY

    body_top = icon_top + Inches(0.45)
    body_box = slide.shapes.add_textbox(
        left + Inches(0.2), body_top, width - Inches(0.4),
        height - Inches(0.55)
    )
    tf2 = body_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = body_text
    p2.font.size = Pt(12)
    p2.font.color.rgb = theme.MEDIUM_RGB
    p2.font.name = theme.FONT_FAMILY
    p2.space_after = Pt(0)
    p2.space_before = Pt(0)
    tf2.margin_left = Pt(0)
    tf2.margin_right = Pt(0)
    tf2.margin_top = Pt(0)
    tf2.margin_bottom = Pt(0)

    return bg
