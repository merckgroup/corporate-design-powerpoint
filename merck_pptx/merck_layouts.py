"""Merck Layouts: rendering engine for Merck-branded PowerPoint decks.

Produces consistent 16:9 slides across four locked styles:
  merck_executive | merck_corporate | merck_storytelling | merck_science

This module is a public re-export facade. All implementation lives in
the private _ml_* submodules. Import from this module as before — the
public API is identical to v3.
"""

# --- Constants & palette ---
from ._ml_constants import (
    FONT_HEAD, FONT_BODY, SLIDE_W, SLIDE_H,
    MERCK_PURPLE, MERCK_BLUE, MERCK_GOLD, PURPLE_DEEP, PURPLE_MUTED,
    MERCK_YELLOW, MERCK_AQUA, LIGHT_GRAY, PANEL_LIGHT, WHITE,
    INK_DARK, INK_GRAY, BAD_RED, GOOD_GREEN,
    ACT_PURPLE, LY_CYAN, OP_LIME, FC_PINK, DEV_POS_BLUE, DEV_NEG_RED,
    CHART_1, CHART_2, CHART_3, CHART_4, CHART_5, CHART_6,
    CHART_7, CHART_8, CHART_9, CHART_10, CHART_11, CHART_12,
    CHART_PALETTE, PHASE_1_COLOR, PHASE_2_COLOR, PHASE_3_COLOR,
    PALETTES, AUTO_PROMOTE_EXECUTIVE,
    GOLD_RULE_X, GOLD_RULE_Y, GOLD_RULE_W, GOLD_RULE_H,
    CLASS_BADGE_X, CLASS_BADGE_Y, CLASS_BADGE_W,
    BREADCRUMB_X, BREADCRUMB_Y, BREADCRUMB_W, BREADCRUMB_H,
    SECTION_CIRCLE_X, SECTION_CIRCLE_Y, SECTION_CIRCLE_D,
    SECTION_TAG_X, SECTION_TAG_Y, SECTION_TAG_W, SECTION_TAG_H,
    TITLE_X, TITLE_Y_NUMBERED, TITLE_Y_UNNUMBERED, TITLE_W, TITLE_H,
    SUB_X, SUB_W, SUB_H, SUB_GAP,
    CONTENT_X, CONTENT_Y, CONTENT_Y_SUBTITLE, CONTENT_W, CONTENT_H,
    SOURCE_Y, SOURCE_H, TAKEAWAY_Y, TAKEAWAY_H, PHASE_Y, PHASE_H,
    FOOTER_Y, FOOTER_H, FOOTER_TEXT_Y,
    _palette_for, _rgb_tuple, _is_dark, _DARK_STYLES,
    rgb,
)

# --- Primitives ---
from ._ml_primitives import (
    _apply_fill, _apply_border, _emu,
    rect, rounded, oval, circle, line, hairline, txt, _add_run, _freeform_poly,
    draw_harvey_ball,
)

# --- Icons ---
from ._ml_icons import (
    icon_chart_bar, icon_chart_line, icon_chart_pie,
    icon_arrow_up, icon_arrow_down, icon_arrow_right,
    icon_check, icon_x, icon_alert, icon_info, icon_target,
    icon_gear, icon_users, icon_calendar, icon_clock,
    icon_lightbulb, icon_lock, icon_globe, icon_search,
    icon_money, icon_trending_up, icon_trending_down,
    icon_shield, icon_flag, icon_doc,
    ICON_REGISTRY, draw_icon,
)

# --- Deck lifecycle & slide utilities ---
from ._ml_deck import (
    open_deck, save_deck,
    _blank_layout, _new_slide, _intro_layout,
    _divider_layout, _cover_picture_layout, _populate_placeholder,
    add_image, add_slide_jump_hyperlink, add_speaker_notes,
)

# --- Chrome, cards, bullets ---
from ._ml_chrome import (
    merck_stub, category_tag, action_title, subtitle_line,
    takeaway_band, source_line, footnotes_block, page_number,
    apply_chrome, stub_and_flag, decimal_align,
    _top_chrome, _bottom_chrome, _section_marker, _render_action_title,
    _subtitle, _source_line, _methodology_note, _takeaway_band,
    _superscript, _tracked, _track_letters,
    _estimate_card_content_h, _compute_row_card_height,
    _pad_int, _format_section_number,
    _draw_card, _gold_square_bullet, _bulleted_list, _phase_progress,
    statement_card, in_slide_section,
)

# --- Chart helpers ---
from ._ml_charts import (
    add_slope_chart, add_dot_plot, add_marimekko,
    add_waterfall, add_small_multiples, add_simple_bar,
    _render_chart, _scale,
)

# --- Style helpers ---
from ._ml_helpers import (
    _style_or_promote, _tone_color, _rag_color, _norm_key,
)

# --- Layout functions — core structural slides ---
from ._ml_layouts_core import (
    _draw_cover_keymessages_grid, _draw_cover_authors,
    build_cover, build_exec_summary, build_agenda,
    build_section_divider, build_close,
)

# --- Layout functions — chart & data slides ---
from ._ml_layouts_chart import (
    _auto_callout_for_chart,
    build_chart_slide, build_waterfall_slide,
    build_hero_stat, build_stat_strip,
    build_donut_chart, build_risk_heatmap, build_radar_chart,
)

# --- Layout functions — column & list slides ---
from ._ml_layouts_column import (
    _COLUMN_PRESETS, _NUMBERED_PRESETS, _two_or_three_column_card,
    build_two_column, build_three_column, build_four_column,
    build_vertical_numbered, build_label_rows,
)

# --- Layout functions — decision & analysis slides ---
from ._ml_layouts_decision import (
    _MATRIX_PRESETS,
    build_2x2_matrix, build_decision_rows, build_before_after,
    build_comparison_table, build_score_table,
    build_influence_diagram, build_pros_cons,
)

# --- Layout functions — process & timeline slides ---
from ._ml_layouts_process import (
    _CIRCULAR_PRESETS, _draw_check_mark,
    build_phase_process, build_gantt, build_milestone_timeline,
    build_circular_flow, build_arrow_chain, build_funnel, build_journey_map,
    build_road_to_success,
)

# --- Layout functions — organizational slides ---
from ._ml_layouts_org import (
    _SPOKE_PRESETS, _dashed_line,
    build_status_table, build_hub_spoke, build_pillar_detail,
    build_org_chart, build_topic_set,
    build_kpi_dashboard, build_icon_grid,
)

# --- Layout functions — visual & story slides ---
from ._ml_layouts_visual import (
    build_pull_quote, build_word_cloud, build_pyramid, build_venn,
    build_layered_stack, build_photo_text, build_fishbone,
    build_key_question,
)

# --- Layout functions — science slides (merck_science style) ---
from ._ml_layouts_science import (
    build_figure_panel, build_methods_box, build_sar_table, build_multi_chart,
)
