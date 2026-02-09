"""Core engine functions for HR Slide Engine."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

from . import design as D


def create_presentation():
    """Create a new 16:9 presentation."""
    prs = Presentation()
    prs.slide_width = D.SLIDE_WIDTH
    prs.slide_height = D.SLIDE_HEIGHT
    return prs


def save_presentation(prs, filename):
    """Save presentation to file. Appends .pptx if missing."""
    if not filename.endswith(".pptx"):
        filename += ".pptx"
    prs.save(filename)
    return filename


def _add_blank_slide(prs):
    """Add a blank slide to the presentation."""
    layout = prs.slide_layouts[6]  # Blank layout
    return prs.slides.add_slide(layout)


def _set_slide_background(slide, color):
    """Set the background color of a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_textbox(slide, left, top, width, height, text,
                 font_size=D.BODY_SIZE, font_color=D.DARK_TEXT,
                 bold=False, alignment=PP_ALIGN.LEFT,
                 font_name=D.FONT_FAMILY, anchor=MSO_ANCHOR.TOP):
    """Add a textbox with formatted text to a slide."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    # Set vertical alignment
    txBox.text_frame._txBody.bodyPr.set("anchor", {
        MSO_ANCHOR.TOP: "t",
        MSO_ANCHOR.MIDDLE: "ctr",
        MSO_ANCHOR.BOTTOM: "b",
    }.get(anchor, "t"))

    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment

    return txBox


def _add_multiline_textbox(slide, left, top, width, height, lines,
                           font_size=D.BODY_SIZE, font_color=D.DARK_TEXT,
                           bold=False, alignment=PP_ALIGN.LEFT,
                           font_name=D.FONT_FAMILY, line_spacing=None,
                           bullet_color=None, bullet_char=None):
    """Add a textbox with multiple paragraphs."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        if bullet_char:
            run_bullet = p.add_run()
            run_bullet.text = f"{bullet_char} "
            run_bullet.font.size = font_size
            run_bullet.font.bold = bold
            run_bullet.font.name = font_name
            run_bullet.font.color.rgb = bullet_color or font_color

            run_text = p.add_run()
            run_text.text = line
            run_text.font.size = font_size
            run_text.font.bold = False
            run_text.font.name = font_name
            run_text.font.color.rgb = font_color
        else:
            p.text = line
            p.font.size = font_size
            p.font.color.rgb = font_color
            p.font.bold = bold
            p.font.name = font_name

        p.alignment = alignment
        if line_spacing:
            p.space_after = line_spacing

    return txBox


def _add_speaker_notes(slide, notes_text):
    """Add speaker notes to a slide."""
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


def _add_rectangle(slide, left, top, width, height, fill_color):
    """Add a filled rectangle shape."""
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()  # No border
    return shape


def _add_line(slide, left, top, width, height, color, line_width=Pt(2)):
    """Add a line shape."""
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape
