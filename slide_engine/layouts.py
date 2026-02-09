"""Layout functions for HR Slide Engine — 8 professional slide types."""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from . import design as D
from .engine import (
    _add_blank_slide,
    _set_slide_background,
    _add_textbox,
    _add_multiline_textbox,
    _add_speaker_notes,
    _add_rectangle,
    _add_line,
)


def add_title_slide(prs, title, subtitle="", notes=""):
    """Slide 1 — Title: navy background, white centered text."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.NAVY)

    # Title
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=Inches(2.2),
        width=D.CONTENT_WIDTH, height=Inches(1.5),
        text=title,
        font_size=Pt(36), font_color=D.WHITE,
        bold=True, alignment=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.BOTTOM,
    )

    # Orange accent line
    line_width = Inches(3)
    line_left = (D.SLIDE_WIDTH - line_width) // 2
    _add_line(slide, line_left, Inches(3.8), line_width, Pt(3), D.ORANGE)

    # Subtitle
    if subtitle:
        _add_textbox(
            slide,
            left=D.MARGIN_LEFT, top=Inches(4.1),
            width=D.CONTENT_WIDTH, height=Inches(1.0),
            text=subtitle,
            font_size=D.SUBTITLE_SIZE, font_color=D.LIGHT_GRAY,
            bold=False, alignment=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.TOP,
        )

    _add_speaker_notes(slide, notes)
    return slide


def add_agenda_slide(prs, items, title="Agenda", notes=""):
    """Slide 2 — Agenda: numbered list with orange numbers."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.WHITE)

    # Title
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=D.MARGIN_TOP,
        width=D.CONTENT_WIDTH, height=Inches(0.8),
        text=title,
        font_size=D.TITLE_SIZE, font_color=D.NAVY,
        bold=True, alignment=PP_ALIGN.LEFT,
    )

    # Underline
    _add_line(slide, D.MARGIN_LEFT, Inches(1.5), Inches(2), Pt(3), D.ORANGE)

    # Items
    y_start = Inches(2.0)
    item_height = Inches(0.6)
    for i, item in enumerate(items, 1):
        y = y_start + (i - 1) * item_height

        # Orange number
        _add_textbox(
            slide,
            left=D.MARGIN_LEFT, top=y,
            width=Inches(0.6), height=item_height,
            text=f"{i:02d}",
            font_size=Pt(22), font_color=D.ORANGE,
            bold=True, alignment=PP_ALIGN.LEFT,
        )

        # Item text
        _add_textbox(
            slide,
            left=D.MARGIN_LEFT + Inches(0.7), top=y,
            width=D.CONTENT_WIDTH - Inches(0.7), height=item_height,
            text=item,
            font_size=D.BODY_SIZE, font_color=D.DARK_TEXT,
            bold=False, alignment=PP_ALIGN.LEFT,
        )

    _add_speaker_notes(slide, notes)
    return slide


def add_section_slide(prs, title, subtitle="", notes=""):
    """Slide 3 — Section divider: navy bar on the left, large title."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.WHITE)

    # Navy vertical bar
    _add_rectangle(
        slide,
        left=Inches(0.4), top=Inches(1.5),
        width=D.SECTION_BAR_WIDTH, height=Inches(4.5),
        fill_color=D.NAVY,
    )

    # Title
    _add_textbox(
        slide,
        left=Inches(1.0), top=Inches(2.5),
        width=Inches(10.5), height=Inches(1.5),
        text=title,
        font_size=Pt(32), font_color=D.NAVY,
        bold=True, alignment=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.BOTTOM,
    )

    # Subtitle
    if subtitle:
        _add_textbox(
            slide,
            left=Inches(1.0), top=Inches(4.2),
            width=Inches(10.5), height=Inches(0.8),
            text=subtitle,
            font_size=D.SUBTITLE_SIZE, font_color=D.GRAY,
            bold=False, alignment=PP_ALIGN.LEFT,
        )

    _add_speaker_notes(slide, notes)
    return slide


def add_bullets_slide(prs, title, bullets, notes=""):
    """Slide 4 — Bullet points: orange bullets, gray text."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.WHITE)

    # Title
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=D.MARGIN_TOP,
        width=D.CONTENT_WIDTH, height=Inches(0.8),
        text=title,
        font_size=D.TITLE_SIZE, font_color=D.NAVY,
        bold=True, alignment=PP_ALIGN.LEFT,
    )

    # Underline
    _add_line(slide, D.MARGIN_LEFT, Inches(1.5), Inches(2), Pt(3), D.ORANGE)

    # Bullet items
    _add_multiline_textbox(
        slide,
        left=D.MARGIN_LEFT + Inches(0.3), top=Inches(2.0),
        width=D.CONTENT_WIDTH - Inches(0.3), height=Inches(4.5),
        lines=bullets,
        font_size=D.BODY_SIZE, font_color=D.GRAY,
        bullet_char=D.BULLET_CHAR, bullet_color=D.ORANGE,
        line_spacing=D.PARAGRAPH_SPACING,
    )

    _add_speaker_notes(slide, notes)
    return slide


def add_two_columns_slide(prs, title, left_title, left_items,
                          right_title, right_items, notes=""):
    """Slide 5 — Two columns: separated by a thin gray line."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.WHITE)

    # Title
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=D.MARGIN_TOP,
        width=D.CONTENT_WIDTH, height=Inches(0.8),
        text=title,
        font_size=D.TITLE_SIZE, font_color=D.NAVY,
        bold=True, alignment=PP_ALIGN.LEFT,
    )

    _add_line(slide, D.MARGIN_LEFT, Inches(1.5), Inches(2), Pt(3), D.ORANGE)

    col_width = (D.CONTENT_WIDTH - D.COLUMN_GAP) / 2
    left_x = D.MARGIN_LEFT
    right_x = D.MARGIN_LEFT + col_width + D.COLUMN_GAP

    # Vertical separator
    sep_x = D.MARGIN_LEFT + col_width + (D.COLUMN_GAP // 2)
    _add_line(slide, sep_x, Inches(2.0), Pt(1), Inches(4.5), D.LIGHT_GRAY)

    # Left column title
    _add_textbox(
        slide,
        left=left_x, top=Inches(2.0),
        width=col_width, height=Inches(0.6),
        text=left_title,
        font_size=Pt(22), font_color=D.NAVY,
        bold=True, alignment=PP_ALIGN.LEFT,
    )

    # Left column items
    _add_multiline_textbox(
        slide,
        left=left_x + Inches(0.2), top=Inches(2.7),
        width=col_width - Inches(0.2), height=Inches(3.8),
        lines=left_items,
        font_size=Pt(18), font_color=D.GRAY,
        bullet_char=D.BULLET_CHAR, bullet_color=D.ORANGE,
        line_spacing=D.LINE_SPACING,
    )

    # Right column title
    _add_textbox(
        slide,
        left=right_x, top=Inches(2.0),
        width=col_width, height=Inches(0.6),
        text=right_title,
        font_size=Pt(22), font_color=D.NAVY,
        bold=True, alignment=PP_ALIGN.LEFT,
    )

    # Right column items
    _add_multiline_textbox(
        slide,
        left=right_x + Inches(0.2), top=Inches(2.7),
        width=col_width - Inches(0.2), height=Inches(3.8),
        lines=right_items,
        font_size=Pt(18), font_color=D.GRAY,
        bullet_char=D.BULLET_CHAR, bullet_color=D.ORANGE,
        line_spacing=D.LINE_SPACING,
    )

    _add_speaker_notes(slide, notes)
    return slide


def add_key_stat_slide(prs, stat, description, notes=""):
    """Slide 6 — Key statistic: large orange number centered."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.WHITE)

    # Big stat
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=Inches(1.8),
        width=D.CONTENT_WIDTH, height=Inches(2.5),
        text=stat,
        font_size=D.STAT_SIZE, font_color=D.ORANGE,
        bold=True, alignment=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.BOTTOM,
    )

    # Description
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=Inches(4.5),
        width=D.CONTENT_WIDTH, height=Inches(1.5),
        text=description,
        font_size=D.BODY_SIZE, font_color=D.GRAY,
        bold=False, alignment=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.TOP,
    )

    _add_speaker_notes(slide, notes)
    return slide


def add_quote_slide(prs, quote, author="", notes=""):
    """Slide 7 — Quote: light gray background, decorative quotation mark."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.LIGHT_GRAY)

    # Decorative large quote mark
    _add_textbox(
        slide,
        left=Inches(1.0), top=Inches(1.0),
        width=Inches(2.0), height=Inches(2.0),
        text=D.QUOTE_CHAR,
        font_size=Pt(120), font_color=D.ORANGE,
        bold=False, alignment=PP_ALIGN.LEFT,
    )

    # Quote text
    _add_textbox(
        slide,
        left=Inches(2.0), top=Inches(2.5),
        width=Inches(9.0), height=Inches(2.5),
        text=quote,
        font_size=D.QUOTE_SIZE, font_color=D.NAVY,
        bold=False, alignment=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Author
    if author:
        _add_textbox(
            slide,
            left=Inches(2.0), top=Inches(5.3),
            width=Inches(9.0), height=Inches(0.6),
            text=f"— {author}",
            font_size=D.SUBTITLE_SIZE, font_color=D.GRAY,
            bold=False, alignment=PP_ALIGN.LEFT,
        )

    _add_speaker_notes(slide, notes)
    return slide


def add_conclusion_slide(prs, title, points, notes=""):
    """Slide 8 — Conclusion: navy banner at top, checkmarks orange."""
    slide = _add_blank_slide(prs)
    _set_slide_background(slide, D.WHITE)

    # Navy banner
    _add_rectangle(
        slide,
        left=Inches(0), top=Inches(0),
        width=D.SLIDE_WIDTH, height=Inches(1.8),
        fill_color=D.NAVY,
    )

    # Title on banner
    _add_textbox(
        slide,
        left=D.MARGIN_LEFT, top=Inches(0.4),
        width=D.CONTENT_WIDTH, height=Inches(1.0),
        text=title,
        font_size=D.TITLE_SIZE, font_color=D.WHITE,
        bold=True, alignment=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Conclusion points with checkmarks
    _add_multiline_textbox(
        slide,
        left=D.MARGIN_LEFT + Inches(0.3), top=Inches(2.3),
        width=D.CONTENT_WIDTH - Inches(0.3), height=Inches(4.5),
        lines=points,
        font_size=D.BODY_SIZE, font_color=D.DARK_TEXT,
        bullet_char=D.CHECKMARK_CHAR, bullet_color=D.ORANGE,
        line_spacing=D.PARAGRAPH_SPACING,
    )

    _add_speaker_notes(slide, notes)
    return slide
