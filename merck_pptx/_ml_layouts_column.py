from __future__ import annotations
import math
import sys
from typing import Optional
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt
from ._ml_constants import (
    FONT_HEAD, FONT_BODY,
    MERCK_PURPLE, MERCK_GOLD, MERCK_YELLOW, MERCK_BLUE, MERCK_AQUA,
    WHITE, INK_DARK, INK_GRAY, LIGHT_GRAY, PANEL_LIGHT,
    PURPLE_DEEP, PURPLE_MUTED, BAD_RED, GOOD_GREEN,
    ACT_PURPLE, LY_CYAN, OP_LIME, FC_PINK, DEV_POS_BLUE, DEV_NEG_RED,
    CHART_PALETTE, PHASE_1_COLOR, PHASE_2_COLOR, PHASE_3_COLOR,
    CONTENT_X, CONTENT_Y, CONTENT_Y_SUBTITLE, CONTENT_W, CONTENT_H,
    SOURCE_Y, SOURCE_H, TAKEAWAY_Y, TAKEAWAY_H, PHASE_Y, PHASE_H,
    FOOTER_Y, FOOTER_H, FOOTER_TEXT_Y,
    TITLE_X, TITLE_Y_NUMBERED, TITLE_Y_UNNUMBERED, TITLE_W, TITLE_H,
    SUB_X, SUB_W, SUB_H, SUB_GAP,
    AUTO_PROMOTE_EXECUTIVE,
    _palette_for, _rgb_tuple, _is_dark,
)
from ._ml_primitives import (
    rect, rounded, oval, circle, line, hairline, txt, _add_run,
    _freeform_poly, _emu, _apply_fill, _apply_border,
)
from ._ml_icons import draw_icon, ICON_REGISTRY
from ._ml_deck import (
    _new_slide, _intro_layout, _divider_layout,
    _cover_picture_layout, _populate_placeholder,
)
from ._ml_chrome import (
    apply_chrome, _phase_progress, _draw_card, _bulleted_list,
    stub_and_flag, footnotes_block, _section_marker,
    _top_chrome, _bottom_chrome, _gold_square_bullet,
    _tracked, _track_letters, _format_section_number, _pad_int,
    _render_action_title, _source_line,
    statement_card, in_slide_section,
    _takeaway_band, _superscript, _callout_block,
)
from ._ml_charts import (
    add_slope_chart, add_dot_plot, add_marimekko, add_waterfall,
    add_small_multiples, add_simple_bar, _render_chart,
)
from ._ml_helpers import _style_or_promote, _tone_color, _rag_color, _norm_key

# ===========================================================================
# Framework presets for non-matrix layouts.
# Set "framework" in content to auto-populate columns/items/phases/spokes.
# ===========================================================================

_COLUMN_PRESETS = {
    "4ps": {
        "columns": [
            {"label": "PRODUCT",   "body": "What you sell"},
            {"label": "PRICE",     "body": "What you charge"},
            {"label": "PLACE",     "body": "How you distribute"},
            {"label": "PROMOTION", "body": "How you communicate"},
        ],
    },
    "value_disciplines": {
        "columns": [
            {"label": "PRODUCT LEADERSHIP",    "body": "New products, new markets, experimentation"},
            {"label": "OPERATIONAL EXCELLENCE", "body": "Efficiency, cost, volume, narrow product lines"},
            {"label": "CUSTOMER INTIMACY",      "body": "Deep customer focus, high expertise in chosen areas"},
        ],
    },
    "balanced_scorecard": {
        "columns": [
            {"label": "FINANCIAL",  "body": "How do we look to shareholders?"},
            {"label": "CUSTOMERS",  "body": "How do customers see us?"},
            {"label": "PROCESSES",  "body": "What must we excel at?"},
            {"label": "LEARNING",   "body": "How can we improve and create value?"},
        ],
    },
}

_NUMBERED_PRESETS = {
    "kotter_8": {
        "items": [
            {"number": "1", "title": "Create Urgency",         "body": "Help others see the need for change"},
            {"number": "2", "title": "Build the Coalition",    "body": "Assemble a group with power to lead change"},
            {"number": "3", "title": "Form Strategic Vision",  "body": "Shape a vision and initiatives"},
            {"number": "4", "title": "Enlist a Volunteer Army","body": "Raise a large force of people"},
            {"number": "5", "title": "Enable Action",          "body": "Remove barriers to change"},
            {"number": "6", "title": "Generate Short-Term Wins","body": "Recognise and reward progress"},
            {"number": "7", "title": "Sustain Acceleration",   "body": "Press harder and faster"},
            {"number": "8", "title": "Institute Change",       "body": "Reinforce new behaviours"},
        ],
    },
    "adkar": {
        "items": [
            {"number": "1", "title": "Awareness",   "body": "Of the need for change"},
            {"number": "2", "title": "Desire",      "body": "To participate and support the change"},
            {"number": "3", "title": "Knowledge",   "body": "On how to change"},
            {"number": "4", "title": "Ability",     "body": "To implement required skills and behaviours"},
            {"number": "5", "title": "Reinforcement","body": "To sustain the change"},
        ],
    },
}


# ===========================================================================
# Layout: TWO COLUMN
# ===========================================================================

def _two_or_three_column_card(slide, x, y, w, h, col, palette):
    """Render a column card with dark-bar header + tone dot treatment.

    Both normal and highlighted cards use PANEL_LIGHT fill. The PURPLE_DEEP
    header bar blends into the card chrome. A small dot signals tone:
    PURPLE_MUTED (neutral/normal) or theme hot colour (highlighted).

    Highlight derives from tone="positive" automatically.
    An explicit "highlighted" key overrides tone when present.
    """
    pal = _palette_for(palette)
    if "highlighted" in col:
        highlighted = bool(col["highlighted"])
    else:
        highlighted = col.get("tone") == "positive"

    # Card background — always PANEL_LIGHT.
    card = rounded(slide, x, y, w, h, fill=PANEL_LIGHT)
    _apply_border(card, LIGHT_GRAY, Pt(0.5))

    # Header bar: MERCK_PURPLE (#503291) normal, LY_CYAN (#2DBECD) highlighted
    HDR_H    = Inches(0.40)
    DOT_SZ   = Inches(0.16)
    bar_fill = LY_CYAN if highlighted else MERCK_PURPLE
    lbl_col  = PURPLE_DEEP if highlighted else WHITE   # dark text on cyan, white on purple
    rect(slide, x, y, w, HDR_H, fill=bar_fill)

    # Tone dot
    dot_col = pal["hot"] if highlighted else PURPLE_MUTED
    dot_x   = x + Inches(0.16)
    dot_y   = y + (HDR_H - DOT_SZ) / 2
    circle(slide, dot_x, dot_y, DOT_SZ, fill=dot_col)

    # Label text on bar — accept "label" (internal name) or "header" (schema name).
    label = col.get("label") or col.get("header", "")
    if label:
        hdr_box = txt(slide, dot_x + DOT_SZ + Inches(0.10), y,
                      w - DOT_SZ - Inches(0.40), HDR_H,
                      _tracked(label), sz=9, color=lbl_col, bold=True,
                      font=FONT_BODY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        try:
            from pptx.enum.text import MSO_AUTO_SIZE
            hdr_box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass

    # Body content below the bar
    pad   = Inches(0.20)
    cur_y = y + HDR_H + Inches(0.18)

    body_text = col.get("body", "")
    items     = col.get("items") or col.get("bullets") or []

    if body_text:
        hero_h = Inches(0.95)
        if len(str(body_text)) > 60:
            hero_h = Inches(1.30)
        if len(str(body_text)) > 120:
            hero_h = Inches(1.55)
        hero_sz = 16 if len(str(body_text)) <= 60 else 14
        txt(slide, x + pad, cur_y, w - pad * 2, hero_h,
            body_text, sz=hero_sz, color=MERCK_PURPLE, bold=True,
            font=FONT_BODY)
        cur_y += hero_h + Inches(0.05)

    if items:
        avail_h = (y + h) - cur_y - Inches(0.50)
        _bulleted_list(slide, x + pad, cur_y, w - pad * 2,
                       max(avail_h, Inches(0.30)),
                       items, palette, text_color=INK_DARK, sz=11,
                       bullet_color=pal["highlight"])

    # Timestamp pill at bottom
    ts = col.get("timestamp")
    if ts:
        pill_w = Inches(1.6)
        pill_h = Inches(0.32)
        pill_x = x + pad
        pill_y = y + h - pill_h - Inches(0.18)
        rect(slide, pill_x, pill_y, pill_w, pill_h,
             fill=pal["hot"] if highlighted else PURPLE_MUTED)
        txt(slide, pill_x, pill_y, pill_w, pill_h,
            str(ts), sz=10,
            color=PURPLE_DEEP if highlighted else WHITE,
            bold=True, font=FONT_BODY,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def build_two_column(prs, meta, action_title=None, left=None, right=None, takeaway=None,
                     source=None, category=None, style="merck_executive",
                     page=None, total=None, section_number=None,
                     methodology_note=None, subtitle=None, content=None):
    if content:
        if "left"     in content: left     = content["left"]
        if "right"    in content: right    = content["right"]
        if "takeaway" in content: takeaway = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number,
                 methodology_note=methodology_note)
    zone_top = Inches(2.95) if subtitle else Inches(2.55)
    zone_h = Inches(6.50) - zone_top
    gap = Inches(0.30)
    col_w = (Inches(12.0) - gap) / 2
    _two_or_three_column_card(slide, Inches(0.65), zone_top, col_w, zone_h,
                              left or {}, style)
    _two_or_three_column_card(slide, Inches(0.65) + col_w + gap, zone_top,
                              col_w, zone_h, right or {}, style)
    return slide


def build_three_column(prs, meta, action_title=None, columns=None, takeaway=None,
                       source=None, category=None, style="merck_executive",
                       page=None, total=None, section_number=None,
                       methodology_note=None, subtitle=None, content=None):
    if content:
        fw = content.get("framework")
        if fw and fw in _COLUMN_PRESETS and not columns:
            columns = _COLUMN_PRESETS[fw]["columns"]
        if "columns"  in content: columns  = content["columns"]
        if "takeaway" in content: takeaway = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number,
                 methodology_note=methodology_note)
    cols = (columns or [])[:3]
    while len(cols) < 3:
        cols.append({})
    zone_top = Inches(2.95) if subtitle else Inches(2.55)
    zone_h = Inches(6.50) - zone_top
    gap = Inches(0.25)
    col_w = (Inches(12.0) - gap * 2) / 3
    for i, col in enumerate(cols):
        cx = Inches(0.65) + i * (col_w + gap)
        _two_or_three_column_card(slide, cx, zone_top, col_w, zone_h, col, style)
    return slide


# ===========================================================================
# Layout: VERTICAL NUMBERED
# ===========================================================================

def build_vertical_numbered(prs, meta, action_title=None, items=None, takeaway=None,
                            source=None, category=None, style="merck_executive",
                            page=None, total=None, section_number=None,
                            methodology_note=None, subtitle=None, content=None):
    if content:
        fw = content.get("framework")
        if fw and fw in _NUMBERED_PRESETS and not items:
            items = _NUMBERED_PRESETS[fw]["items"]
        if "items"    in content: items    = content["items"]
        if "takeaway" in content: takeaway = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number, methodology_note=methodology_note)

    rows = (items or [])[:5]
    n = max(len(rows), 1)
    # Scale number and title font sizes down for dense layouts so badges
    # don't overflow into adjacent rows.
    num_sz   = 34 if n >= 5 else (38 if n == 4 else 44)
    title_sz = 13 if n >= 5 else (14 if n == 4 else 15)
    zone_top = Inches(2.95) if subtitle else Inches(2.55)
    zone_h = Inches(6.50) - zone_top
    gap = Inches(0.12)
    row_h = (zone_h - gap * (n - 1)) / n
    for i, row in enumerate(rows):
        ry = zone_top + i * (row_h + gap)
        num_w = Inches(0.90)
        # Number badge in MERCK_PURPLE.
        txt(slide, Inches(0.65), ry, num_w, row_h, str(i + 1),
            sz=num_sz, color=MERCK_PURPLE, bold=True, font=FONT_BODY,
            anchor=MSO_ANCHOR.TOP)
        bx = Inches(0.65) + num_w + Inches(0.25)
        bw = Inches(12.0) - num_w - Inches(0.35)
        txt(slide, bx, ry + Inches(0.04), bw, Inches(0.36),
            row.get("title", ""), sz=title_sz, color=INK_DARK, bold=True,
            font=FONT_BODY)
        txt(slide, bx, ry + Inches(0.42), bw, row_h - Inches(0.46),
            row.get("body", ""), sz=11, color=INK_GRAY, font=FONT_BODY)
        if i < n - 1:
            hairline(slide, Inches(0.65),
                     ry + row_h + gap / 2 - Emu(int(Pt(0.5))),
                     Inches(12.0), Emu(int(Pt(0.5))), LIGHT_GRAY)
    return slide



# ===========================================================================
# build_four_column: 4-up cards in a row
# ===========================================================================

def build_four_column(prs, meta, action_title=None, columns=None, takeaway=None,
                      source=None, category=None, style="merck_executive",
                      page=None, total=None, section_number=None,
                      methodology_note=None, subtitle=None, content=None):
    """4-up columns. Same card chrome as three_column.
    columns = list of dicts: {label, title, body, tone, items}.
    Set "framework" in content for auto-configured presets (e.g. "4ps", "balanced_scorecard").
    """
    if content:
        fw = content.get("framework")
        if fw and fw in _COLUMN_PRESETS and not columns:
            columns = _COLUMN_PRESETS[fw]["columns"]
        if "columns"  in content: columns  = content["columns"]
        if "takeaway" in content: takeaway = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category, subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number,
                 methodology_note=methodology_note)
    cols = (columns or [])[:4]
    while len(cols) < 4:
        cols.append({})
    zone_top = Inches(2.55) if not subtitle else Inches(2.95)
    zone_h = Inches(6.30) - zone_top
    gap = Inches(0.18)
    col_w = (Inches(12.0) - gap * 3) / 4
    for i, col in enumerate(cols):
        cx = Inches(0.65) + i * (col_w + gap)
        _two_or_three_column_card(slide, cx, zone_top, col_w, zone_h, col, style)
    return slide


# ===========================================================================
# build_label_rows: colored label card on left + body on right per row
# ===========================================================================

def build_label_rows(prs, meta, action_title=None, rows=None, takeaway=None,
                     source=None, category=None, style="merck_executive",
                     page=None, total=None, section_number=None,
                     methodology_note=None, subtitle=None,
                     label_color=None, content=None):
    """Row-per-item layout with a colored LABEL CARD on the left and BODY text on
    the right. Use for diagnosis / root cause / risk taxonomy slides where the
    name of each item is the anchor.

    rows = list of {"label": str, "body": str} (3-6 rows).
    label_color: color tuple for the label card fill. Defaults to MERCK_PURPLE.
        Set explicitly for a one-off accent (e.g. BAD_RED for negative tone).
    """
    if content:
        if "rows"        in content: rows        = content["rows"]
        if "takeaway"    in content: takeaway    = content["takeaway"]
        if "label_color" in content: label_color = content["label_color"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category, subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number,
                 methodology_note=methodology_note)

    rs = list(rows or [])[:6]
    n = max(len(rs), 1)
    zone_top = Inches(2.55) if not subtitle else Inches(2.95)
    zone_h = SOURCE_Y - zone_top - Inches(0.20)
    gap = Inches(0.14)
    row_h = (zone_h - gap * (n - 1)) / n

    label_w = Inches(3.10)
    body_x = Inches(0.65) + label_w + Inches(0.30)
    body_w = Inches(12.0) - label_w - Inches(0.30)
    fill_color = label_color if label_color is not None else MERCK_PURPLE
    # On dark / storytelling bg, body text needs to be light.
    body_color = WHITE if _is_dark(style) else INK_DARK

    for i, row in enumerate(rs):
        ry = zone_top + i * (row_h + gap)
        # Label card: colored fill, rounded corners, white bold text centered.
        card = rounded(slide, Inches(0.65), ry, label_w, row_h,
                       fill=fill_color, adj=2000)
        _apply_border(card, None)
        txt(slide, Inches(0.65) + Inches(0.12), ry,
            label_w - Inches(0.24), row_h,
            str(row.get("label", "")), sz=14, color=WHITE, bold=True,
            font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Body text on right.
        txt(slide, body_x, ry, body_w, row_h,
            str(row.get("body", "")), sz=12, color=body_color,
            font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)
        # Hairline below row (except last).
        if i < n - 1:
            rule_color = PURPLE_MUTED if _is_dark(style) else LIGHT_GRAY
            hairline(slide, Inches(0.65), ry + row_h + gap / 2 - Emu(int(Pt(0.5))),
                     Inches(12.0), Emu(int(Pt(0.5))), rule_color)
    callout = (content or {}).get("callout")
    if isinstance(callout, dict):
        _callout_block(slide, callout.get("type", "conclusion"),
                       callout.get("text", ""), pal,
                       has_takeaway=bool(takeaway))
    return slide



