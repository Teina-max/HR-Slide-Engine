"""HR Slide Engine — Professional PowerPoint generation for HR presentations."""

from .engine import create_presentation, save_presentation
from .layouts import (
    add_title_slide,
    add_agenda_slide,
    add_section_slide,
    add_bullets_slide,
    add_two_columns_slide,
    add_key_stat_slide,
    add_quote_slide,
    add_conclusion_slide,
)

__all__ = [
    "create_presentation",
    "save_presentation",
    "add_title_slide",
    "add_agenda_slide",
    "add_section_slide",
    "add_bullets_slide",
    "add_two_columns_slide",
    "add_key_stat_slide",
    "add_quote_slide",
    "add_conclusion_slide",
]
