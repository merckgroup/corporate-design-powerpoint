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
    _freeform_poly, _emu, _apply_fill, _apply_border, draw_harvey_ball,
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
# Layout: 2x2 MATRIX
# ===========================================================================

# Framework presets: set "framework" in content to auto-configure axes and quadrants.
# Explicit x_axis/y_axis/quadrants in content always override the preset.
_MATRIX_PRESETS = {
    "bcg": {
        "x_axis": "Relative Market Share",
        "y_axis": "Market Growth Rate",
        "quadrants": {
            "top_left":     {"label": "Stars",           "highlighted": True},
            "top_right":    {"label": "Question Marks"},
            "bottom_left":  {"label": "Cash Cows"},
            "bottom_right": {"label": "Dogs"},
        },
    },
    "swot": {
        "x_axis": "Internal / External",
        "y_axis": "Positive / Negative",
        "quadrants": {
            "top_left":     {"label": "Strengths",      "highlighted": True},
            "top_right":    {"label": "Opportunities"},
            "bottom_left":  {"label": "Weaknesses"},
            "bottom_right": {"label": "Threats"},
        },
    },
    "ansoff": {
        "x_axis": "Products (Existing → New)",
        "y_axis": "Markets (Existing → New)",
        "quadrants": {
            "top_left":     {"label": "Market Development"},
            "top_right":    {"label": "Diversification",   "highlighted": True},
            "bottom_left":  {"label": "Market Penetration"},
            "bottom_right": {"label": "Product Development"},
        },
    },
    "risk": {
        "x_axis": "Probability (Low → High)",
        "y_axis": "Impact (Low → High)",
        "quadrants": {
            "top_left":     {"label": "Monitor"},
            "top_right":    {"label": "Critical",          "highlighted": True},
            "bottom_left":  {"label": "Accept"},
            "bottom_right": {"label": "Mitigate"},
        },
    },
}


def build_2x2_matrix(prs, meta, action_title=None, x_axis=None, y_axis=None, quadrants=None,
                     takeaway=None, source=None, category=None,
                     style="merck_executive", page=None, total=None,
                     section_number=None, methodology_note=None,
                     subtitle=None, content=None):
    if content:
        framework = content.get("framework")
        if framework and framework in _MATRIX_PRESETS:
            preset = _MATRIX_PRESETS[framework]
            if not x_axis:    x_axis    = preset.get("x_axis")
            if not y_axis:    y_axis    = preset.get("y_axis")
            if not quadrants: quadrants = preset.get("quadrants")
        if "x_axis"    in content: x_axis    = content["x_axis"]
        if "y_axis"    in content: y_axis    = content["y_axis"]
        if "quadrants" in content: quadrants = content["quadrants"]
        if "takeaway"  in content: takeaway  = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number,
                 methodology_note=methodology_note)

    margin_left = Inches(0.55)
    margin_bottom = Inches(0.55)
    zone_top = Inches(2.95) if subtitle else Inches(2.55)
    zone_h = Inches(6.50) - zone_top
    plot_x = Inches(0.65) + margin_left
    plot_y = zone_top + Inches(0.05)
    plot_w = Inches(12.0) - margin_left
    plot_h = zone_h - margin_bottom - Inches(0.10)
    cell_w = plot_w / 2
    cell_h = plot_h / 2

    cells = ["top_left", "top_right", "bottom_left", "bottom_right"]
    positions = [
        (plot_x, plot_y),
        (plot_x + cell_w, plot_y),
        (plot_x, plot_y + cell_h),
        (plot_x + cell_w, plot_y + cell_h),
    ]
    for k, key in enumerate(cells):
        cx, cy = positions[k]
        q = (quadrants or {}).get(key, {})
        highlighted = bool(q.get("highlighted"))
        if highlighted:
            quad_fill = pal["accent_2"] if _is_dark(style) else pal["accent"]
            rounded(slide, cx, cy, cell_w, cell_h, fill=quad_fill)
            rounded(slide, cx, cy, cell_w, Inches(0.06), fill=pal["hot"], adj=50000)
            tcolor = WHITE
            btext = PANEL_LIGHT
            label_color = pal["hot"]
        else:
            cell = rounded(slide, cx, cy, cell_w, cell_h, fill=PANEL_LIGHT)
            _apply_border(cell, LIGHT_GRAY, Pt(0.5))
            rounded(slide, cx, cy, cell_w, Inches(0.06), fill=pal["highlight"], adj=50000)
            tcolor = MERCK_PURPLE
            btext = INK_DARK
            label_color = pal["highlight"]
        label = q.get("label", "")
        items = q.get("items") or q.get("bullets") or []
        body  = q.get("body", "")
        pad = Inches(0.18)
        txt(slide, cx + pad, cy + Inches(0.18), cell_w - pad * 2, Inches(0.26),
            _tracked(label), sz=10, color=label_color, bold=True,
            font=FONT_BODY)
        if items:
            _bulleted_list(slide, cx + pad, cy + Inches(0.50),
                           cell_w - pad * 2, cell_h - Inches(0.65),
                           items, palette=style, text_color=btext, sz=10,
                           bullet_color=label_color)
        elif body:
            txt(slide, cx + pad, cy + Inches(0.50),
                cell_w - pad * 2, cell_h - Inches(0.65),
                str(body), sz=11, color=btext, font=FONT_BODY,
                anchor=MSO_ANCHOR.TOP)

    # Axes lines (light).
    line(slide, plot_x, plot_y + plot_h, plot_x + plot_w, plot_y + plot_h,
         PURPLE_MUTED, Pt(0.75))
    line(slide, plot_x, plot_y, plot_x, plot_y + plot_h, PURPLE_MUTED, Pt(0.75))
    ah = Inches(0.12)
    _freeform_poly(slide,
                   [(plot_x + plot_w + ah, plot_y + plot_h),
                    (plot_x + plot_w, plot_y + plot_h - ah / 2),
                    (plot_x + plot_w, plot_y + plot_h + ah / 2)],
                   fill=PURPLE_MUTED)
    _freeform_poly(slide,
                   [(plot_x, plot_y - ah),
                    (plot_x - ah / 2, plot_y),
                    (plot_x + ah / 2, plot_y)],
                   fill=PURPLE_MUTED)

    x_ax = x_axis or {}
    y_ax = y_axis or {}
    txt(slide, plot_x, plot_y + plot_h + Inches(0.08), plot_w, Inches(0.26),
        _tracked(x_ax.get("label", "")), sz=10, color=pal["highlight"], bold=True,
        font=FONT_BODY, align=PP_ALIGN.CENTER)
    txt(slide, plot_x, plot_y + plot_h + Inches(0.30), Inches(1.5),
        Inches(0.22),
        x_ax.get("low", ""), sz=9, color=PURPLE_MUTED, font=FONT_BODY)
    txt(slide, plot_x + plot_w - Inches(1.5), plot_y + plot_h + Inches(0.30),
        Inches(1.5), Inches(0.22),
        x_ax.get("high", ""), sz=9, color=PURPLE_MUTED, font=FONT_BODY,
        align=PP_ALIGN.RIGHT)
    txt(slide, Inches(0.65), plot_y, margin_left, plot_h,
        _tracked(y_ax.get("label", "")), sz=10, color=pal["highlight"], bold=True,
        font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return slide



# ===========================================================================
# Layout: DECISION ROWS
# ===========================================================================

def build_decision_rows(prs, meta, action_title=None, decisions=None, takeaway=None,
                        source=None, category=None, page=None, total=None,
                        style=None, section_number=None, methodology_note=None,
                        subtitle=None, content=None):
    """Auto-promoted to executive style (unless overridden by an exec-eligible
    category). The style kwarg is honored when the category does not trigger
    auto-promotion.
    """
    if content:
        if "decisions" in content: decisions = content["decisions"]
        if "takeaway"  in content: takeaway  = content["takeaway"]
    style = _style_or_promote(category or "Decision Request", style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title,
                 category=category or "Decision Request",
                 subtitle=subtitle,
                 takeaway=takeaway, source=source, page=page, total=total,
                 palette=style, section_number=section_number, methodology_note=methodology_note)

    rows = (decisions or [])[:5]
    n = max(len(rows), 1)
    zone_top = Inches(2.95) if subtitle else Inches(2.55)
    zone_h = Inches(6.50) - zone_top
    gap = Inches(0.12)
    row_h = (zone_h - gap * (n - 1)) / n
    for i, d in enumerate(rows):
        ry = zone_top + i * (row_h + gap)
        # Card chrome (panel + hairline + theme-accent stripe on top).
        card = rect(slide, Inches(0.65), ry, Inches(12.0), row_h,
                    fill=PANEL_LIGHT)
        _apply_border(card, LIGHT_GRAY, Pt(0.5))
        rect(slide, Inches(0.65), ry, Inches(12.0), Inches(0.06),
             fill=pal["highlight"])
        # Number circle.
        tone = d.get("tone", "neutral")
        tone_color = _tone_color(tone, style)
        circ_d = Inches(0.55)
        cx = Inches(0.95)
        cy = ry + (row_h - circ_d) / 2
        circle(slide, cx, cy, circ_d, fill=tone_color)
        txt(slide, cx, cy, circ_d, circ_d, str(d.get("number") or (i + 1)),
            sz=18, color=WHITE, bold=True, font=FONT_BODY,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Title + description.
        bx = cx + circ_d + Inches(0.25)
        owner = d.get("owner", "")
        owner_w = Inches(2.4) if owner else Inches(0.0)
        bw = Inches(12.0) - (bx - Inches(0.65)) - owner_w - Inches(0.20)
        title_y = ry + Inches(0.20)
        txt(slide, bx, title_y, bw, Inches(0.34),
            d.get("title") or d.get("decision", ""), sz=14, color=MERCK_PURPLE, bold=True,
            font=FONT_BODY)
        txt(slide, bx, title_y + Inches(0.36), bw,
            row_h - Inches(0.60),
            d.get("desc") or d.get("body") or d.get("description", ""), sz=11, color=INK_GRAY, font=FONT_BODY)
        if owner:
            txt(slide, Inches(0.65) + Inches(12.0) - owner_w - Inches(0.15),
                ry, owner_w, row_h,
                f"Owner: {owner}", sz=10, color=PURPLE_MUTED, italic=True,
                font=FONT_BODY, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    callout = (content or {}).get("callout")
    if isinstance(callout, dict):
        _callout_block(slide, callout.get("type", "conclusion"),
                       callout.get("text", ""), pal,
                       has_takeaway=bool(takeaway))
    return slide



# ===========================================================================
# Layout: BEFORE / AFTER (Today vs Tomorrow with gold arrow)
# ===========================================================================

def build_before_after(prs, meta, action_title=None, before=None, after=None, takeaway=None,
                       source=None, category=None, subtitle=None,
                       style="merck_executive", page=None, total=None,
                       section_number=None,
                       before_label="TODAY", after_label="TOMORROW", methodology_note=None,
                       content=None):
    """Two cards side by side with a gold right-arrow between them.

    before, after: dicts {title, items} where title is the headline (Verdana
    18pt bold MERCK_PURPLE) and items is a small bullet list.

    The BEFORE card is muted (PANEL_LIGHT bg, hairline border, PURPLE_MUTED
    text). The AFTER card is highlighted (PANEL_LIGHT bg, MERCK_GOLD top
    accent stripe, MERCK_GOLD label, MERCK_PURPLE bold title, INK_DARK body).

    A MERCK_GOLD right-arrow shape sits in the gutter between the cards.
    """
    if content:
        if "before"       in content: before       = content["before"]
        if "after"        in content: after        = content["after"]
        if "takeaway"     in content: takeaway     = content["takeaway"]
        if "before_label" in content: before_label = content["before_label"]
        if "after_label"  in content: after_label  = content["after_label"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle, takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number, methodology_note=methodology_note)

    zone_top = Inches(2.95) if subtitle else Inches(2.50)
    # Equal-height enforcement.
    base_h = Inches(2.55)
    estimated = []
    for col in (before or {}, after or {}):
        est = Inches(0.60)  # label + title
        title_len = len(str(col.get("title", "")))
        if title_len:
            title_lines = max(1, (title_len // 30) + (1 if title_len % 30 else 0))
            est += Inches(0.40) * title_lines
        items = col.get("items") or []
        est += Inches(0.10) + Inches(0.32) * len(items)
        est += Inches(0.30)
        estimated.append(est)
    card_h = max(base_h, max(estimated))
    # Clamp.
    max_h_avail = Inches(6.40) - zone_top
    if card_h > max_h_avail:
        card_h = max_h_avail

    # Card geometry: x=0.65 w=5.7 for BEFORE; x=7.10 w=5.7 for AFTER.
    before_x = Inches(0.65)
    before_w = Inches(5.70)
    after_x = Inches(7.10)
    after_w = Inches(5.70)

    # --- BEFORE card (muted) ---
    bcard = rounded(slide, before_x, zone_top, before_w, card_h,
                 fill=PANEL_LIGHT)
    _apply_border(bcard, LIGHT_GRAY, Pt(0.5))
    # No gold top stripe on the muted card.
    pad = Inches(0.30)
    txt(slide, before_x + pad, zone_top + Inches(0.22),
        before_w - pad * 2, Inches(0.26),
        _track_letters(before_label), sz=10, color=PURPLE_MUTED, bold=True,
        font=FONT_BODY)
    btitle = (before or {}).get("title", "")
    txt(slide, before_x + pad, zone_top + Inches(0.60),
        before_w - pad * 2, Inches(0.95),
        str(btitle), sz=18, color=PURPLE_MUTED, bold=True, font=FONT_BODY)
    bitems = (before or {}).get("items") or (before or {}).get("bullets") or []
    if bitems:
        _bulleted_list(slide, before_x + pad, zone_top + Inches(1.60),
                       before_w - pad * 2, card_h - Inches(1.85),
                       bitems, palette=style,
                       text_color=PURPLE_MUTED, sz=12,
                       bullet_color=PURPLE_MUTED)

    # --- AFTER card (highlighted) ---
    acard = rounded(slide, after_x, zone_top, after_w, card_h, fill=PANEL_LIGHT)
    _apply_border(acard, LIGHT_GRAY, Pt(0.5))
    rounded(slide, after_x, zone_top, after_w, Inches(0.06), fill=pal["highlight"], adj=50000)
    txt(slide, after_x + pad, zone_top + Inches(0.22),
        after_w - pad * 2, Inches(0.26),
        _track_letters(after_label), sz=10, color=pal["highlight"], bold=True,
        font=FONT_BODY)
    atitle = (after or {}).get("title", "")
    txt(slide, after_x + pad, zone_top + Inches(0.60),
        after_w - pad * 2, Inches(0.95),
        str(atitle), sz=18, color=MERCK_PURPLE, bold=True, font=FONT_BODY)
    aitems = (after or {}).get("items") or (after or {}).get("bullets") or []
    if aitems:
        _bulleted_list(slide, after_x + pad, zone_top + Inches(1.60),
                       after_w - pad * 2, card_h - Inches(1.85),
                       aitems, palette=style,
                       text_color=INK_DARK, sz=12,
                       bullet_color=pal["highlight"])

    # --- Theme-accent right-arrow shape between the cards ---
    ar_w = Inches(0.65)
    ar_h = Inches(0.85)
    ar_x = (before_x + before_w + after_x) / 2 - ar_w / 2
    ar_y = zone_top + (card_h - ar_h) / 2
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                   ar_x, ar_y, ar_w, ar_h)
    arrow.shadow.inherit = False
    _apply_fill(arrow, pal["highlight"])
    _apply_border(arrow, None)

    return slide


# End of module



# ===========================================================================
# Layout: COMPARISON TABLE (feature matrix)
# ===========================================================================

def build_comparison_table(prs, meta, action_title=None, options=None,
                            features=None, takeaway="", source=None,
                            subtitle=None, methodology_note=None,
                            style="merck_executive", page=None, total=None,
                            section_number=None, category=None):
    """Feature matrix with check/cross/partial marks.

    options:  list of str — column headers (2-4 options to compare)
    features: list of dicts with keys:
        label       (str) — row label
        values      (list of str) — one per option: 'yes'/'no'/'partial'/text
        highlighted (bool, opt) — highlight this row
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    options  = list(options  or [])[:4]
    features = list(features or [])[:12]
    n_opts   = max(len(options), 1)

    MARKS = {
        "yes":     ("✓", GOOD_GREEN),
        "no":      ("✗", BAD_RED),
        "partial": ("∼", MERCK_YELLOW),
    }

    zone_x    = Inches(0.55)
    zone_y    = Inches(2.95) if subtitle else Inches(2.55)
    zone_w    = Inches(12.2)
    label_w   = Inches(4.20)
    opt_w     = (zone_w - label_w) / n_opts
    header_h  = Inches(0.52)
    avail_h   = SOURCE_Y - zone_y - header_h - Inches(0.15)
    row_h     = max(avail_h / max(len(features), 1), Inches(0.30))

    rect(slide, zone_x, zone_y, label_w, header_h, fill=PURPLE_DEEP)
    for j, opt in enumerate(options):
        px = zone_x + label_w + j * opt_w
        rect(slide, px, zone_y, opt_w, header_h, fill=MERCK_PURPLE)
        txt(slide, px, zone_y, opt_w, header_h,
            str(opt), sz=11, color=WHITE, bold=True,
            font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    for i, feat in enumerate(features):
        ry          = zone_y + header_h + i * row_h
        highlighted = bool(feat.get("highlighted"))
        row_bg      = (MERCK_PURPLE if highlighted else
                       PANEL_LIGHT  if i % 2 == 0 else WHITE)
        lbl_color   = WHITE if highlighted else INK_DARK

        rect(slide, zone_x, ry, label_w, row_h, fill=row_bg)
        txt(slide, zone_x + Inches(0.16), ry, label_w - Inches(0.20), row_h,
            str(feat.get("label", "")),
            sz=10, color=lbl_color, bold=highlighted,
            font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)

        vals = list(feat.get("values", []))
        for j in range(n_opts):
            px = zone_x + label_w + j * opt_w
            rect(slide, px, ry, opt_w, row_h,
                 fill=row_bg, border=LIGHT_GRAY, border_w=Pt(0.25))
            val = vals[j] if j < len(vals) else ""
            mark, m_color = MARKS.get(str(val).lower(), (str(val), INK_DARK))
            if highlighted:
                m_color = WHITE
            txt(slide, px, ry, opt_w, row_h,
                mark, sz=13, color=m_color, bold=True,
                font=FONT_BODY, align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE)

    return slide


# ===========================================================================
# Layout: SCORE / RATING TABLE
# ===========================================================================

def build_score_table(prs, meta, action_title=None, rows=None,
                      scale=5, scale_label=None,
                      takeaway="", source=None, subtitle=None,
                      methodology_note=None, style="merck_executive",
                      page=None, total=None, section_number=None,
                      category=None, content=None):
    """Audit/maturity table with visual dot ratings per row.

    rows: list of dicts with keys:
        label    (str) — item name
        score    (float 0-scale) — rating
        category (str, opt) — category tag
        note     (str, opt) — brief annotation
    scale: int — max rating (default 5)
    """
    if content:
        if "rows"     in content: rows     = content["rows"]
        if "takeaway" in content: takeaway = content["takeaway"]
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    rows = list(rows or [])[:12]
    if not rows:
        return slide

    zone_x  = Inches(0.55)
    zone_y  = Inches(1.62)
    label_w = Inches(4.80)
    cat_w   = Inches(1.60)
    note_w  = Inches(2.80)
    score_w = Inches(2.90)
    header_h = Inches(0.40)
    row_h    = min((Inches(4.72) - header_h) / len(rows), Inches(0.52))

    col_starts = [zone_x,
                  zone_x + label_w,
                  zone_x + label_w + cat_w,
                  zone_x + label_w + cat_w + note_w]
    col_widths = [label_w, cat_w, note_w, score_w]
    col_labels = ["ITEM", "CATEGORY", "NOTES",
                  scale_label or f"RATING (/{scale})"]

    for j, (cx, cw, cl) in enumerate(zip(col_starts, col_widths, col_labels)):
        rect(slide, cx, zone_y, cw, header_h,
             fill=PURPLE_DEEP if j == 3 else MERCK_PURPLE)
        txt(slide, cx + Inches(0.10), zone_y, cw - Inches(0.10), header_h,
            cl, sz=9, color=WHITE, bold=True, font=FONT_BODY,
            align=PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.MIDDLE)

    dot_sz, dot_gap = Inches(0.16), Inches(0.06)
    rating_type = (content or {}).get("rating_type", "dots")

    for i, row in enumerate(rows):
        ry = zone_y + header_h + i * row_h
        bg = PANEL_LIGHT if i % 2 == 0 else WHITE

        rect(slide, zone_x, ry, label_w, row_h, fill=bg)
        txt(slide, zone_x + Inches(0.14), ry, label_w - Inches(0.18), row_h,
            str(row.get("label", "")),
            sz=10, color=INK_DARK, font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)

        rect(slide, col_starts[1], ry, cat_w, row_h,
             fill=bg, border=LIGHT_GRAY, border_w=Pt(0.25))
        if row.get("category"):
            txt(slide, col_starts[1] + Inches(0.06), ry,
                cat_w - Inches(0.12), row_h, str(row["category"]),
                sz=8, color=PURPLE_MUTED, font=FONT_BODY,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        rect(slide, col_starts[2], ry, note_w, row_h,
             fill=bg, border=LIGHT_GRAY, border_w=Pt(0.25))
        if row.get("note"):
            txt(slide, col_starts[2] + Inches(0.10), ry,
                note_w - Inches(0.16), row_h, str(row["note"]),
                sz=9, color=INK_GRAY, italic=True,
                font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)

        rect(slide, col_starts[3], ry, score_w, row_h,
             fill=bg, border=LIGHT_GRAY, border_w=Pt(0.25))
        score = float(row.get("score", 0))
        if rating_type == "harvey":
            # Harvey Ball: score is 0.0-1.0 fill fraction
            hb_d = Inches(0.30)
            hb_x = col_starts[3] + (score_w - hb_d) / 2
            hb_y = ry + (row_h - hb_d) / 2
            fill_pct = max(0.0, min(1.0, score))
            draw_harvey_ball(slide, hb_x, hb_y, hb_d, fill_pct,
                             filled_color=pal["accent"],
                             border_color=pal["accent"])
        else:
            total_dot_w = scale * (dot_sz + dot_gap) - dot_gap
            dot_x0      = col_starts[3] + (score_w - total_dot_w) / 2
            dot_y       = ry + (row_h - dot_sz) / 2
            for d in range(scale):
                filled = (d + 1) <= score
                half   = not filled and d < score < d + 1
                fill_c = (MERCK_PURPLE if filled else
                          MERCK_GOLD   if half   else LIGHT_GRAY)
                circle(slide, dot_x0 + d * (dot_sz + dot_gap), dot_y,
                       dot_sz, fill=fill_c)

    return slide


# ===========================================================================
# Layout: INFLUENCE DIAGRAM
# ===========================================================================

def build_influence_diagram(prs, meta, action_title=None, center=None,
                             forces=None, takeaway="", source=None,
                             subtitle=None, methodology_note=None,
                             style="merck_executive", page=None,
                             total=None, section_number=None,
                             category=None):
    """Central topic box with labeled force arrows pushing in from sides.

    center: dict with label (str) and body (str)
    forces: list of dicts with:
        label (str), body (str),
        side  ('left'/'right'/'top'/'bottom'),
        tone  ('positive'/'negative'/'neutral')
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    center = center or {}
    cx, cy = Inches(6.666), Inches(3.80)
    cw, ch = Inches(3.20),  Inches(1.60)
    rounded(slide, cx - cw / 2, cy - ch / 2, cw, ch, fill=MERCK_PURPLE)
    txt(slide, cx - cw / 2 + Inches(0.12), cy - ch / 2 + Inches(0.10),
        cw - Inches(0.24), Inches(0.30),
        str(center.get("label", "")).upper(),
        sz=9, color=MERCK_YELLOW, bold=True, font=FONT_BODY)
    txt(slide, cx - cw / 2 + Inches(0.12), cy - ch / 2 + Inches(0.38),
        cw - Inches(0.24), ch - Inches(0.46),
        str(center.get("body", "")),
        sz=13, color=WHITE, bold=True, font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

    TONE_COLOR = {"positive": GOOD_GREEN, "negative": BAD_RED,
                  "neutral": MERCK_BLUE}
    SIDE_BASE  = {
        "left":   (Inches(0.55), Inches(2.30)),
        "right":  (Inches(9.80), Inches(2.30)),
        "top":    (Inches(3.80), Inches(1.62)),
        "bottom": (Inches(3.80), Inches(5.60)),
    }
    force_w, force_h, gap = Inches(2.60), Inches(1.00), Inches(0.22)
    processed = {s: 0 for s in SIDE_BASE}

    for f in list(forces or []):
        side = str(f.get("side", "left")).lower()
        if side not in SIDE_BASE:
            side = "left"
        idx  = processed[side]
        col  = TONE_COLOR.get(str(f.get("tone", "neutral")).lower(), MERCK_BLUE)
        bx0, by0 = SIDE_BASE[side]
        if side in ("left", "right"):
            bx, by = bx0, by0 + idx * (force_h + gap)
        else:
            bx, by = bx0 + idx * (force_w + gap), by0

        rounded(slide, bx, by, force_w, force_h, fill=col)
        txt(slide, bx + Inches(0.10), by + Inches(0.08),
            force_w - Inches(0.20), Inches(0.24),
            str(f.get("label", "")).upper(),
            sz=8, color=WHITE, bold=True, font=FONT_BODY)
        txt(slide, bx + Inches(0.10), by + Inches(0.30),
            force_w - Inches(0.20), force_h - Inches(0.36),
            str(f.get("body", "")),
            sz=9, color=WHITE, font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

        if side == "left":
            ax1, ay1 = bx + force_w, by + force_h / 2
            ax2, ay2 = cx - cw / 2,  cy
        elif side == "right":
            ax1, ay1 = bx,           by + force_h / 2
            ax2, ay2 = cx + cw / 2,  cy
        elif side == "top":
            ax1, ay1 = bx + force_w / 2, by + force_h
            ax2, ay2 = cx,               cy - ch / 2
        else:
            ax1, ay1 = bx + force_w / 2, by
            ax2, ay2 = cx,               cy + ch / 2
        line(slide, ax1, ay1, ax2, ay2, color=col, weight=Pt(1.5))
        processed[side] += 1

    return slide



# ===========================================================================
# Layout: PROS / CONS
# ===========================================================================

def build_pros_cons(prs, meta, action_title=None, pros=None, cons=None,
                    pros_label="ADVANTAGES", cons_label="RISKS",
                    subject=None, takeaway="", source=None, subtitle=None,
                    methodology_note=None, style="merck_executive",
                    page=None, total=None, section_number=None,
                    category=None):
    """Two-panel green pros vs red cons layout.

    pros:    list of str — advantage bullets
    cons:    list of str — risk/disadvantage bullets
    subject: str — optional dark banner above both panels
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    pros = list(pros or [])
    cons = list(cons or [])
    zone_x = Inches(0.55)
    zone_y = Inches(2.10) if not subtitle else Inches(2.55)
    zone_w = Inches(12.2)
    zone_h = Inches(6.50) - zone_y - Inches(1.10)   # respect takeaway/source at bottom
    gap    = Inches(0.14)

    if subject:
        subj_h = Inches(0.46)
        rect(slide, zone_x, zone_y, zone_w, subj_h, fill=PURPLE_DEEP)
        txt(slide, zone_x + Inches(0.20), zone_y, zone_w, subj_h,
            str(subject), sz=13, color=WHITE, bold=True,
            font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)
        panel_y = zone_y + subj_h + gap
        panel_h = zone_h - subj_h - gap
    else:
        panel_y, panel_h = zone_y, zone_h

    panel_w = (zone_w - gap) / 2
    HDR_H   = Inches(0.46)
    BULLET  = "•  "
    PANEL_BG  = {"pros": (0xE8, 0xF5, 0xE9), "cons": (0xFF, 0xEB, 0xEB)}
    HDR_COLOR = {"pros": GOOD_GREEN, "cons": BAD_RED}
    ICON      = {"pros": "✓", "cons": "✗"}

    for side in ("pros", "cons"):
        items = pros if side == "pros" else cons
        px    = zone_x if side == "pros" else zone_x + panel_w + gap
        lbl   = pros_label if side == "pros" else cons_label
        rect(slide, px, panel_y, panel_w, panel_h, fill=PANEL_BG[side])
        rect(slide, px, panel_y, panel_w, HDR_H,   fill=HDR_COLOR[side])
        txt(slide, px + Inches(0.18), panel_y, panel_w, HDR_H,
            f"{ICON[side]}  {lbl}",
            sz=12, color=WHITE, bold=True, font=FONT_BODY,
            anchor=MSO_ANCHOR.MIDDLE)
        body_y = panel_y + HDR_H + Inches(0.12)
        body_h = panel_h - HDR_H - Inches(0.16)
        item_h = min(body_h / max(len(items), 1), Inches(0.60))
        for i, item in enumerate(items):
            txt(slide, px + Inches(0.18), body_y + i * item_h,
                panel_w - Inches(0.30), item_h,
                BULLET + str(item),
                sz=10, color=INK_DARK, font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

    return slide


