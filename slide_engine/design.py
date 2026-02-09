"""Design system constants for HR Slide Engine."""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

# === Slide Dimensions (16:9) ===
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# === Color Palette ===
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
GRAY = RGBColor(0x6B, 0x72, 0x80)
ORANGE = RGBColor(0xE8, 0x7C, 0x3E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF3, 0xF4, 0xF6)
DARK_TEXT = RGBColor(0x1F, 0x2A, 0x37)

# === Typography ===
FONT_FAMILY = "Calibri"

TITLE_SIZE = Pt(28)
TITLE_BOLD = True

BODY_SIZE = Pt(20)
BODY_BOLD = False

SUBTITLE_SIZE = Pt(16)
SUBTITLE_BOLD = False

NOTES_SIZE = Pt(11)

STAT_SIZE = Pt(72)
QUOTE_SIZE = Pt(24)

# === Margins & Spacing ===
MARGIN_LEFT = Inches(0.8)
MARGIN_TOP = Inches(0.6)
MARGIN_RIGHT = Inches(0.8)
MARGIN_BOTTOM = Inches(0.5)

CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT  # ~11.733 inches
CONTENT_HEIGHT = SLIDE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM  # ~6.4 inches

LINE_SPACING = Pt(6)
PARAGRAPH_SPACING = Pt(12)

# === Layout-specific ===
SECTION_BAR_WIDTH = Inches(0.15)
COLUMN_GAP = Inches(0.5)
BULLET_CHAR = "\u2022"  # •
CHECKMARK_CHAR = "\u2713"  # ✓
QUOTE_CHAR = "\u201C"  # "
