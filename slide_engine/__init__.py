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
    # Visual layouts
    add_process_flow_slide,
    add_timeline_slide,
    add_matrix_slide,
    add_pyramid_slide,
    add_bar_chart_slide,
    add_pie_chart_slide,
    add_icon_cards_slide,
    add_org_chart_slide,
    add_funnel_slide,
    add_team_grid_slide,
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
    # Visual layouts
    "add_process_flow_slide",
    "add_timeline_slide",
    "add_matrix_slide",
    "add_pyramid_slide",
    "add_bar_chart_slide",
    "add_pie_chart_slide",
    "add_icon_cards_slide",
    "add_org_chart_slide",
    "add_funnel_slide",
    "add_team_grid_slide",
]
