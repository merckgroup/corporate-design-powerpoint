from __future__ import annotations
import math
import sys
from typing import Optional
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
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
    _takeaway_band, _superscript,
)
from ._ml_charts import (
    add_slope_chart, add_dot_plot, add_marimekko, add_waterfall,
    add_small_multiples, add_simple_bar, _render_chart,
)
from ._ml_helpers import _style_or_promote, _tone_color, _rag_color, _norm_key

_SPOKE_PRESETS = {
    "porters_5": {
        "hub": {"label": "Competitive Rivalry", "body": "Intensity among existing competitors"},
        "spokes": [
            {"label": "Supplier Power",     "body": "Bargaining power of suppliers"},
            {"label": "Buyer Power",        "body": "Bargaining power of customers"},
            {"label": "Threat of New Entry","body": "Ease of entry for new competitors"},
            {"label": "Substitution Risk",  "body": "Threat from substitute products"},
        ],
    },
}


def build_status_table(prs, meta, action_title=None, columns=None, rows=None, takeaway=None,
                       source=None, methodology_note=None,
                       category=None, subtitle=None,
                       style="merck_executive", page=None, total=None,
                       section_number=None, content=None):
    """RAG status / dependency table with header row + RAG pills.
    Caps at 7 rows; extra rows are dropped.
    """
    if content:
        if "columns"  in content: columns  = content["columns"]
        if "rows"     in content: rows     = content["rows"]
        if "takeaway" in content: takeaway = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle, takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    cols = list(columns or [])
    if not cols:
        return slide
    rs = list(rows or [])[:7]

    # Detect RAG columns by header text.
    rag_set = {"status", "rag", "health", "severity",
               "priority", "risk level", "rag status"}
    rag_cols = [i for i, c in enumerate(cols)
                if str(c).strip().lower() in rag_set]

    # Compute column widths. RAG columns are narrower.
    zone_x = CONTENT_X
    zone_w = CONTENT_W
    zone_y = Inches(2.55) if not subtitle else Inches(2.95)
    n = len(cols)
    rag_w = Inches(1.20)
    non_rag_count = n - len(rag_cols)
    remaining = zone_w - rag_w * len(rag_cols)
    first_col_share = 0.40 if non_rag_count >= 2 else 1.0
    other_col_share = (1.0 - first_col_share) / max(non_rag_count - 1, 1)
    col_widths = []
    seen_first = False
    for i, c in enumerate(cols):
        if i in rag_cols:
            col_widths.append(rag_w)
        else:
            if not seen_first:
                col_widths.append(remaining * first_col_share)
                seen_first = True
            else:
                col_widths.append(remaining * other_col_share)

    # Header row.
    header_h = Inches(0.42)
    avail_h = SOURCE_Y - zone_y - header_h - Inches(0.15)
    n_rows = max(len(rs), 1)
    row_h = min(Inches(0.78), avail_h / n_rows)
    row_h = max(row_h, Inches(0.34))  # floor for legibility
    rect(slide, zone_x, zone_y, zone_w, header_h, fill=MERCK_PURPLE)
    cx = zone_x
    for i, c in enumerate(cols):
        txt(slide, cx + Inches(0.18), zone_y, col_widths[i] - Inches(0.20),
            header_h, _track_letters(str(c).upper()),
            sz=10, color=WHITE, bold=True, font=FONT_BODY,
            anchor=MSO_ANCHOR.MIDDLE)
        cx = cx + col_widths[i]

    # Body rows.
    PANEL_LAVENDER = RGBColor(0xF7, 0xF4, 0xFA)
    for ri, row in enumerate(rs):
        ry = zone_y + header_h + row_h * ri
        rect(slide, zone_x, ry, zone_w, row_h, fill=tuple(PANEL_LAVENDER))
        # Bottom hairline.
        hairline(slide, zone_x, ry + row_h - Emu(int(Pt(0.5))),
                 zone_w, Emu(int(Pt(0.5))), LIGHT_GRAY)
        cx = zone_x
        first_text_col = True
        for i, c in enumerate(cols):
            col_key = _norm_key(c)
            val = ""
            for k, v in row.items():
                if _norm_key(k) == col_key:
                    val = v
                    break
            if val == "":
                # Substring fallback: match if either key contains the other
                for k, v in row.items():
                    nk = _norm_key(k)
                    # Only match when one key is a full prefix of the other
                    # (avoids "status" falsely matching "programme_status").
                    if nk and col_key and (
                        col_key.startswith(nk) or nk.startswith(col_key)
                    ):
                        val = v
                        break
            if val == "":
                # Fallback to first-word match or exact-lowercase match.
                _parts = str(c).strip().lower().split() if c else []
                first_word = _parts[0] if _parts else ""
                val = row.get(first_word) or row.get(str(c).strip().lower()) or ""
            cell_w = col_widths[i]
            if i in rag_cols:
                # RAG pill.
                pill_w = Inches(0.85)
                pill_h = Inches(0.35)
                pill_x = cx + (cell_w - pill_w) / 2
                pill_y = ry + (row_h - pill_h) / 2
                pill = rounded(slide, pill_x, pill_y, pill_w, pill_h,
                               fill=_rag_color(val), adj=5000)
                _apply_border(pill, None)
                txt(slide, pill_x, pill_y, pill_w, pill_h,
                    str(val).upper(), sz=10, color=WHITE, bold=True,
                    font=FONT_BODY, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
            else:
                # Text cell. First column anchors TOP so long labels don't clip.
                color = MERCK_PURPLE if i == 0 else INK_DARK
                bold = (i == 0)
                first_text_col = False
                anchor = MSO_ANCHOR.TOP if i == 0 else MSO_ANCHOR.MIDDLE
                cell_sz = 9 if (i == 0 and len(str(val)) > 30) else 10
                txt(slide, cx + Inches(0.18), ry + Inches(0.06),
                    cell_w - Inches(0.20), row_h - Inches(0.08),
                    str(val), sz=cell_sz, color=color, bold=bold,
                    font=FONT_BODY, anchor=anchor)
            cx = cx + cell_w
    return slide



def _dashed_line(slide, x1, y1, x2, y2, color, weight=Pt(1.0)):
    """Add a dashed connector between two points."""
    from lxml import etree
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      _emu(x1), _emu(y1),
                                      _emu(x2), _emu(y2))
    conn.line.color.rgb = _rgb_tuple(color)
    conn.line.width = weight
    # Set dash style via XML.
    ln = conn.line._get_or_add_ln()
    pr_dash = ln.find(qn("a:prstDash"))
    if pr_dash is None:
        pr_dash = etree.SubElement(ln, qn("a:prstDash"))
    pr_dash.set("val", "dash")
    return conn


def build_hub_spoke(prs, meta, action_title=None, hub=None, spokes=None, takeaway=None,
                    source=None, methodology_note=None,
                    category=None, subtitle=None,
                    style="merck_executive", page=None, total=None,
                    section_number=None, content=None):
    """Center hub oval + 4 corner cards with dashed connectors."""
    if content:
        fw = content.get("framework")
        if fw and fw in _SPOKE_PRESETS and not hub and not spokes:
            hub    = _SPOKE_PRESETS[fw]["hub"]
            spokes = _SPOKE_PRESETS[fw]["spokes"]
        if "hub"      in content: hub      = content["hub"]
        if "spokes"   in content: spokes   = content["spokes"]
        if "takeaway" in content: takeaway = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle, takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    sp = list(spokes or [])[:8]
    n_sp = len(sp)

    # Center hub oval — use theme accent
    hub_w = Inches(2.80)
    hub_h = Inches(2.00)
    hub_x = Inches(13.333) / 2 - hub_w / 2
    hub_y = Inches(4.55) - hub_h / 2
    hub_shape = oval(slide, hub_x, hub_y, hub_w, hub_h, fill=pal["accent"])
    _apply_border(hub_shape, pal["highlight"], Pt(2.25))

    # Hub contents.
    hub_label = (hub or {}).get("label", "")
    hub_title = (hub or {}).get("title", "")
    hub_sub = (hub or {}).get("subtitle", "")
    if hub_label:
        txt(slide, hub_x, hub_y + Inches(0.35), hub_w, Inches(0.30),
            _track_letters(str(hub_label).upper()),
            sz=10, color=pal["highlight"], bold=True, font=FONT_BODY,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if hub_title:
        txt(slide, hub_x, hub_y + Inches(0.70), hub_w, Inches(0.55),
            str(hub_title), sz=18, color=WHITE, bold=True,
            font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if hub_sub:
        txt(slide, hub_x, hub_y + Inches(1.30), hub_w, Inches(0.45),
            str(hub_sub), sz=11, color=WHITE, italic=True,
            font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Dynamic positions: distribute cards evenly
    # For 1-4: corners; for 5-8: add mid-side positions
    # Scale card size down for higher spoke counts to prevent overlap
    if n_sp <= 4:
        card_w, card_h = Inches(3.40), Inches(1.55)
    elif n_sp <= 6:
        card_w, card_h = Inches(2.80), Inches(1.30)
    else:
        card_w, card_h = Inches(2.40), Inches(1.10)
    MARGIN_X = Inches(0.30)
    MARGIN_Y = Inches(2.50)
    ZONE_BOT = Inches(6.80)
    RIGHT_X  = Inches(13.333) - MARGIN_X - card_w

    all_positions = [
        (MARGIN_X,  MARGIN_Y),                          # 0 top-left
        (RIGHT_X,   MARGIN_Y),                          # 1 top-right
        (MARGIN_X,  ZONE_BOT - card_h),                 # 2 bottom-left
        (RIGHT_X,   ZONE_BOT - card_h),                 # 3 bottom-right
        (MARGIN_X,  (MARGIN_Y + ZONE_BOT - card_h)/2),  # 4 mid-left
        (RIGHT_X,   (MARGIN_Y + ZONE_BOT - card_h)/2),  # 5 mid-right
        ((Inches(13.333) - card_w)/2, MARGIN_Y),         # 6 top-center
        ((Inches(13.333) - card_w)/2, ZONE_BOT - card_h),# 7 bottom-center
    ]
    positions = all_positions[:max(n_sp, 1)]

    hub_cx = hub_x + hub_w / 2
    hub_cy = hub_y + hub_h / 2
    title_sz = 13 if n_sp <= 4 else (11 if n_sp <= 6 else 10)
    body_sz  = 10 if n_sp <= 4 else 9
    for i, item in enumerate(sp):
        if i >= len(positions):
            break
        cx, cy = positions[i]
        card = rounded(slide, cx, cy, card_w, card_h, fill=pal["panel"])
        _apply_border(card, LIGHT_GRAY, Pt(0.5))
        # Theme-accent top stripe.
        rounded(slide, cx, cy, card_w, Inches(0.06), fill=pal["highlight"], adj=50000)
        # Title + body.
        pad = Inches(0.22)
        txt(slide, cx + pad, cy + Inches(0.20),
            card_w - pad * 2, Inches(0.45),
            str(item.get("title", "")), sz=title_sz, color=pal["accent"], bold=True,
            font=FONT_BODY)
        txt(slide, cx + pad, cy + Inches(0.72),
            card_w - pad * 2, card_h - Inches(0.80),
            str(item.get("body", "")), sz=body_sz, color=INK_GRAY, font=FONT_BODY)
        # Dashed connector: route horizontally when card is primarily to the
        # side of the hub, vertically when it is above or below (centre slots).
        card_cx = cx + card_w / 2
        card_cy = cy + card_h / 2
        dx = card_cx - hub_cx
        dy = card_cy - hub_cy
        if abs(dx) >= abs(dy):          # primarily left / right
            if dx < 0:
                edge_x, edge_y       = cx + card_w, card_cy
                hub_edge_x, hub_edge_y = hub_x, hub_cy
            else:
                edge_x, edge_y       = cx, card_cy
                hub_edge_x, hub_edge_y = hub_x + hub_w, hub_cy
        else:                           # primarily above / below (centre slots)
            if dy < 0:
                edge_x, edge_y       = card_cx, cy + card_h
                hub_edge_x, hub_edge_y = hub_cx, hub_y
            else:
                edge_x, edge_y       = card_cx, cy
                hub_edge_x, hub_edge_y = hub_cx, hub_y + hub_h
        _dashed_line(slide, edge_x, edge_y, hub_edge_x, hub_edge_y,
                     PURPLE_MUTED, weight=Pt(0.75))
    return slide



def build_pillar_detail(prs, meta, action_title=None, pillar_number=None, pillar_label=None,
                        owner=None, sections=None, takeaway=None,
                        source=None, methodology_note=None,
                        category=None, subtitle=None,
                        style="merck_executive", page=None, total=None,
                        section_number=None, content=None):
    """Dedicated page per pillar: left purple panel with huge number, right
    column with sectioned content.

    pillar_number: e.g. "02"
    pillar_label:  e.g. "PILLAR"
    owner: {"label": "OWNER", "name": "Natalia Mocan, Future Platform Architecture"}
    sections: list of {"label", "body"} dicts
    """
    if content:
        if "pillar_number" in content: pillar_number = content["pillar_number"]
        if "pillar_label"  in content: pillar_label  = content["pillar_label"]
        if "owner"         in content: owner         = content["owner"]
        if "sections"      in content: sections      = content["sections"]
        if "takeaway"      in content: takeaway      = content["takeaway"]
    style = _style_or_promote(category, style)
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 subtitle=subtitle, takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    # Left purple panel.
    panel_x = Inches(0.65)
    panel_y = Inches(2.95) if subtitle else Inches(2.55)
    panel_w = Inches(3.40)
    # Panel bottom snaps just above source / takeaway band.
    panel_h = SOURCE_Y - panel_y - Inches(0.30)
    rect(slide, panel_x, panel_y, panel_w, panel_h, fill=MERCK_PURPLE)

    # Pillar label at top.
    txt(slide, panel_x + Inches(0.30), panel_y + Inches(0.25),
        panel_w - Inches(0.40), Inches(0.30),
        _track_letters(str(pillar_label or "PILLAR").upper()),
        sz=11, color=pal["highlight"], bold=True, font=FONT_BODY)

    # Huge number (Merck Web — hero moment).
    txt(slide, panel_x, panel_y + Inches(0.65),
        panel_w, panel_h - Inches(1.80),
        str(pillar_number or ""), sz=120, color=WHITE, bold=True,
        font=FONT_HEAD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Owner block at panel bottom.
    if owner:
        owner_y = panel_y + panel_h - Inches(0.95)
        owner_label = (owner or {}).get("label", "OWNER")
        owner_name = (owner or {}).get("name", "")
        txt(slide, panel_x + Inches(0.30), owner_y,
            panel_w - Inches(0.40), Inches(0.30),
            _track_letters(str(owner_label).upper()),
            sz=10, color=pal["highlight"], bold=True, font=FONT_BODY)
        txt(slide, panel_x + Inches(0.30), owner_y + Inches(0.30),
            panel_w - Inches(0.40), Inches(0.55),
            str(owner_name), sz=13, color=WHITE, bold=True,
            font=FONT_BODY)

    # Right column: sectioned content blocks.
    sec = list(sections or [])[:5]
    right_x = panel_x + panel_w + Inches(0.40)
    right_w = Inches(13.333) - right_x - Inches(0.65)
    right_y = panel_y
    right_h = panel_h
    if sec:
        block_gap = Inches(0.18)
        block_h = (right_h - block_gap * (len(sec) - 1)) / len(sec)
        for i, s in enumerate(sec):
            by = right_y + (block_h + block_gap) * i
            stripe_w = Inches(0.05)
            rect(slide, right_x, by, stripe_w, block_h, fill=pal["highlight"])
            txt(slide, right_x + stripe_w + Inches(0.18), by,
                right_w - stripe_w - Inches(0.20), Inches(0.28),
                _track_letters(str(s.get("label", "")).upper()),
                sz=10, color=pal["highlight"], bold=True, font=FONT_BODY)
            txt(slide, right_x + stripe_w + Inches(0.18),
                by + Inches(0.32),
                right_w - stripe_w - Inches(0.20),
                block_h - Inches(0.36),
                str(s.get("body", "")), sz=11, color=INK_DARK,
                font=FONT_BODY)
    return slide




def build_org_chart(prs, meta, action_title, root, children,
                    takeaway="", source=None, subtitle=None,
                    methodology_note=None, style="merck_corporate",
                    page=None, total=None, section_number=None, category=None):
    """Two-level org chart.

    root: dict with keys: name (str), title (str, optional).
    children: list of dicts with keys: name (str), title (str, optional),
              reports (list of dicts with name/title, optional).
    Returns the slide.
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    children = children[:6]
    n = len(children)

    # Root box geometry
    root_w = Inches(2.5)
    root_h = Inches(0.7)
    root_x = Inches(13.333 / 2) - root_w / 2
    root_y = Inches(2.0)
    root_cx = root_x + root_w / 2   # horizontal centre of root

    root_box = rounded(slide, root_x, root_y, root_w, root_h,
                       fill=pal["accent"], adj=6000)
    txt(slide, root_x, root_y, root_w, root_h,
        str(root.get("name", "")),
        sz=13, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, font=FONT_BODY,
        anchor=MSO_ANCHOR.MIDDLE)

    # Optional title text beneath root box
    root_title = str(root.get("title", ""))
    if root_title:
        txt(slide, root_x, root_y + root_h + Inches(0.04),
            root_w, Inches(0.25),
            root_title, sz=9, color=pal.get("ink", INK_DARK),
            align=PP_ALIGN.CENTER, font=FONT_BODY,
            anchor=MSO_ANCHOR.TOP)

    if n == 0:
        return slide

    # Child row geometry
    child_w = Inches(1.7)
    child_h = Inches(0.65)
    child_y = Inches(3.4)
    available_w = Inches(12.0)
    col_w = available_w / n
    start_x = (Inches(13.333) - available_w) / 2

    # Vertical trunk: root bottom to midpoint row
    trunk_top_y = root_y + root_h
    trunk_bot_y = child_y                          # top of child boxes
    mid_y = trunk_top_y + (trunk_bot_y - trunk_top_y) / 2

    # Draw trunk line from root bottom down to mid_y
    line(slide, root_cx, trunk_top_y, root_cx, mid_y,
         color=PURPLE_MUTED, weight=Pt(1.0))

    for i, child in enumerate(children):
        child_cx = start_x + col_w * i + col_w / 2
        child_x = child_cx - child_w / 2

        # Horizontal branch at mid_y
        line(slide, root_cx, mid_y, child_cx, mid_y,
             color=PURPLE_MUTED, weight=Pt(1.0))
        # Vertical drop from mid_y to child top
        line(slide, child_cx, mid_y, child_cx, child_y,
             color=PURPLE_MUTED, weight=Pt(1.0))

        # Child box
        rounded(slide, child_x, child_y, child_w, child_h,
                fill=PANEL_LIGHT, adj=6000)
        txt(slide, child_x, child_y, child_w, child_h,
            str(child.get("name", "")),
            sz=11, color=pal["accent"], bold=True,
            align=PP_ALIGN.CENTER, font=FONT_BODY,
            anchor=MSO_ANCHOR.MIDDLE)

        child_title = str(child.get("title", ""))
        if child_title:
            txt(slide, child_x, child_y + child_h + Inches(0.03),
                child_w, Inches(0.22),
                child_title, sz=8, color=pal.get("ink", INK_DARK),
                align=PP_ALIGN.CENTER, font=FONT_BODY,
                anchor=MSO_ANCHOR.TOP)

        # Grandchild boxes
        reports = list(child.get("reports", []))[:2]
        if reports:
            gc_w = Inches(1.4)
            gc_h = Inches(0.5)
            gc_y = Inches(4.5)
            gc_spacing = Inches(0.12)
            total_gc_w = len(reports) * gc_w + (len(reports) - 1) * gc_spacing
            gc_start_x = child_cx - total_gc_w / 2

            # Vertical connector from child bottom to grandchild row
            line(slide, child_cx, child_y + child_h, child_cx, gc_y,
                 color=LIGHT_GRAY, weight=Pt(0.75))

            for j, rep in enumerate(reports):
                gx = gc_start_x + j * (gc_w + gc_spacing)
                gc_cx = gx + gc_w / 2

                # Horizontal branch at gc_y
                if len(reports) > 1:
                    line(slide, child_cx, gc_y, gc_cx, gc_y,
                         color=LIGHT_GRAY, weight=Pt(0.75))

                rounded(slide, gx, gc_y, gc_w, gc_h,
                        fill=LIGHT_GRAY, adj=6000)
                txt(slide, gx, gc_y, gc_w, gc_h,
                    str(rep.get("name", "")),
                    sz=9, color=INK_DARK, bold=False,
                    align=PP_ALIGN.CENTER, font=FONT_BODY,
                    anchor=MSO_ANCHOR.MIDDLE)

    return slide


# ---------------------------------------------------------------------------
# build_topic_set — 3–6 topic cards with numbered circles
# ---------------------------------------------------------------------------


def build_topic_set(prs, meta, action_title, topics, takeaway="", source=None,
                    subtitle=None, methodology_note=None, style="merck_corporate",
                    page=None, total=None, section_number=None, category=None):
    """3-6 topic cards with numbered circles, title and body.

    topics: list of dicts with keys:
        label (str) — short uppercase label shown above title
        title (str) — bold topic title
        body  (str) — 1-2 sentence description
        icon  (str, optional) — icon name for draw_icon; if None, shows sequential number
    Returns the slide.
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    topics = list(topics)[:6]
    n = max(len(topics), 1)

    # Decide row layout
    if n <= 3:
        rows = [topics]
    else:
        rows = [topics[:3], topics[3:]]

    row_count = len(rows)
    zone_x = Inches(0.65)
    zone_y = Inches(2.4)
    zone_w = Inches(12.0)
    zone_h = Inches(3.8)

    gap = Inches(0.25)
    row_h = zone_h / row_count

    circle_size = Inches(0.75)
    circle_r = circle_size / 2

    global_idx = 0

    for row_idx, row_topics in enumerate(rows):
        rn = len(row_topics)
        card_w = (zone_w - gap * (rn - 1)) / rn
        row_y = zone_y + row_idx * row_h

        for col_idx, topic in enumerate(row_topics):
            global_idx += 1
            card_x = zone_x + col_idx * (card_w + gap)

            # Circle centred horizontally on the card
            circ_x = card_x + card_w / 2 - circle_r
            circ_y = row_y
            circle(slide, circ_x, circ_y, circle_size, fill=pal["accent"])

            # Number or icon inside circle
            icon_name = topic.get("icon")
            if icon_name:
                icon_margin = Inches(0.12)
                draw_icon(slide, icon_name,
                          circ_x + icon_margin,
                          circ_y + icon_margin,
                          circle_size - icon_margin * 2,
                          color=WHITE)
            else:
                txt(slide, circ_x, circ_y, circle_size, circle_size,
                    str(global_idx),
                    sz=18, color=WHITE, bold=True,
                    align=PP_ALIGN.CENTER, font=FONT_BODY,
                    anchor=MSO_ANCHOR.MIDDLE)

            cursor_y = circ_y + circle_size + Inches(0.1)

            # Label
            label_text = str(topic.get("label", "")).upper()
            if label_text:
                label_h = Inches(0.22)
                txt(slide, card_x, cursor_y, card_w, label_h,
                    label_text,
                    sz=9, color=pal["highlight"], bold=False,
                    align=PP_ALIGN.CENTER, font=FONT_BODY,
                    anchor=MSO_ANCHOR.TOP)
                cursor_y += label_h + Inches(0.04)

            # Title
            title_h = Inches(0.36)
            txt(slide, card_x, cursor_y, card_w, title_h,
                str(topic.get("title", "")),
                sz=13, color=pal["accent"], bold=True,
                align=PP_ALIGN.CENTER, font=FONT_BODY,
                anchor=MSO_ANCHOR.TOP)
            cursor_y += title_h + Inches(0.08)

            # Body
            remaining_h = (row_y + row_h) - cursor_y - Inches(0.1)
            txt(slide, card_x, cursor_y, card_w, max(remaining_h, Inches(0.4)),
                str(topic.get("body", "")),
                sz=10, color=INK_DARK, bold=False,
                align=PP_ALIGN.CENTER, font=FONT_BODY,
                anchor=MSO_ANCHOR.TOP)

    return slide


# ---------------------------------------------------------------------------
# build_arrow_chain — 3–5 horizontal boxes connected by triangle arrows
# ---------------------------------------------------------------------------


# ===========================================================================
# Layout: KPI DASHBOARD
# ===========================================================================

def build_kpi_dashboard(prs, meta, action_title=None, kpis=None,
                        takeaway="", source=None, subtitle=None,
                        methodology_note=None, style="merck_executive",
                        page=None, total=None, section_number=None,
                        category=None):
    """4-6 KPI metric cards with RAG status and trend direction.

    kpis: list of dicts with keys:
        label     (str) — metric name
        value     (str) — big display number
        unit      (str, opt) — unit suffix shown smaller
        status    (str) — 'green', 'amber', or 'red'
        trend     (str) — 'up', 'down', or 'flat'
        context   (str, opt) — one-line note below the value
        sparkline (list of float, opt) — 5-8 points for mini bar
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    kpis = list(kpis or [])[:6]
    if not kpis:
        return slide

    dark      = _is_dark(style)
    card_fill = PURPLE_DEEP if dark else PANEL_LIGHT
    val_color = pal["hot"]    if dark else pal["accent"]
    lab_color = PANEL_LIGHT  if dark else INK_DARK
    ctx_color = PANEL_LIGHT  if dark else INK_GRAY
    RAG       = {"green": GOOD_GREEN, "amber": MERCK_YELLOW, "red": BAD_RED}
    TREND_G   = {"up": "↑", "down": "↓", "flat": "→"}

    n    = len(kpis)
    cols = min(n, 3)
    rows = (n + cols - 1) // cols
    zone_x, zone_y = Inches(0.55), Inches(1.62)
    zone_w, zone_h = Inches(12.2), Inches(4.70)
    gap    = Inches(0.18)
    card_w = (zone_w - gap * (cols - 1)) / cols
    card_h = (zone_h - gap * (rows - 1)) / rows

    for i, kpi in enumerate(kpis):
        row = i // cols
        col = i % cols
        cx  = zone_x + col * (card_w + gap)
        cy  = zone_y + row * (card_h + gap)

        rounded(slide, cx, cy, card_w, card_h, fill=card_fill)

        rag_col = RAG.get(str(kpi.get("status", "")).lower(), PURPLE_MUTED)
        dot_sz  = Inches(0.20)
        circle(slide, cx + card_w - dot_sz - Inches(0.14),
               cy + Inches(0.14), dot_sz, fill=rag_col)

        t_str   = str(kpi.get("trend", "")).lower()
        glyph   = TREND_G.get(t_str, "")
        t_col   = (GOOD_GREEN if t_str == "up" else
                   BAD_RED    if t_str == "down" else PURPLE_MUTED)
        if glyph:
            txt(slide, cx + Inches(0.12), cy + Inches(0.10),
                Inches(0.30), Inches(0.30), glyph,
                sz=14, color=t_col, bold=True, font=FONT_BODY)

        txt(slide, cx + Inches(0.16), cy + Inches(0.12),
            card_w - Inches(0.60), Inches(0.28),
            str(kpi.get("label", "")),
            sz=9, color=lab_color, font=FONT_BODY, align=PP_ALIGN.LEFT)

        val_y = cy + Inches(0.44)
        txt(slide, cx + Inches(0.16), val_y,
            card_w - Inches(0.32), Inches(0.80),
            str(kpi.get("value", "")),
            sz=34, color=val_color, bold=True, font=FONT_HEAD,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

        if kpi.get("unit"):
            txt(slide, cx + Inches(0.16), val_y + Inches(0.62),
                card_w - Inches(0.32), Inches(0.26),
                str(kpi["unit"]), sz=10, color=ctx_color, font=FONT_BODY)

        if kpi.get("context"):
            txt(slide, cx + Inches(0.16), cy + card_h - Inches(0.58),
                card_w - Inches(0.32), Inches(0.44),
                str(kpi["context"]), sz=9, color=ctx_color,
                italic=True, font=FONT_BODY)

        sp = kpi.get("sparkline")
        if sp and len(sp) >= 2:
            sp   = [float(v) for v in sp]
            mn, mx = min(sp), max(sp)
            rng  = (mx - mn) or 1.0
            sp_x = cx + Inches(0.16)
            sp_y = cy + card_h - Inches(0.30)
            sp_w = card_w - Inches(0.32)
            sp_h = Inches(0.16)
            bar_w = sp_w / len(sp) - Inches(0.02)
            for j, v in enumerate(sp):
                bh = max(sp_h * (v - mn) / rng, Emu(28000))
                rect(slide,
                     sp_x + j * (bar_w + Inches(0.02)), sp_y + sp_h - bh,
                     bar_w, bh,
                     fill=pal["highlight"] if j == len(sp) - 1 else PURPLE_MUTED)

    return slide



# ===========================================================================
# Layout: ICON GRID
# ===========================================================================

def build_icon_grid(prs, meta, action_title=None, items=None,
                    columns=3, takeaway="", source=None, subtitle=None,
                    methodology_note=None, style="merck_executive",
                    page=None, total=None, section_number=None,
                    category=None):
    """6 or 9 equal icon-card cells in a 2x3 or 3x3 grid.

    items: list of dicts with keys:
        icon        (str) — single emoji
        title       (str) — bold card title
        body        (str) — 1-2 sentence description
        highlighted (bool, opt) — MERCK_PURPLE card fill
    columns: 2 or 3 (default 3)
    """
    pal = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])
    apply_chrome(slide, meta, action_title, category=category,
                 takeaway=takeaway, source=source, subtitle=subtitle,
                 methodology_note=methodology_note,
                 page=page, total=total, section_number=section_number,
                 palette=style)

    items = list(items or [])[:9]
    if not items:
        return slide

    cols         = min(int(columns), 3)
    rows_needed  = (len(items) + cols - 1) // cols
    zone_x, zone_y = Inches(0.55), Inches(1.62)
    zone_w, zone_h = Inches(12.2), Inches(4.72)
    gap    = Inches(0.15)
    card_w = (zone_w - gap * (cols - 1)) / cols
    card_h = (zone_h - gap * (rows_needed - 1)) / rows_needed
    icon_sz = Inches(0.55)
    dark    = _is_dark(style)

    for i, item in enumerate(items):
        row = i // cols
        col = i % cols
        cx  = zone_x + col * (card_w + gap)
        cy  = zone_y + row * (card_h + gap)
        highlighted = bool(item.get("highlighted"))
        card_fill   = MERCK_PURPLE if highlighted else PANEL_LIGHT
        icon_bg     = pal["highlight"] if highlighted else MERCK_PURPLE
        title_col   = WHITE        if highlighted else MERCK_PURPLE
        body_col    = PANEL_LIGHT  if highlighted else INK_DARK

        rounded(slide, cx, cy, card_w, card_h, fill=card_fill)
        icon_cx, icon_cy = cx + Inches(0.22), cy + Inches(0.20)
        circle(slide, icon_cx, icon_cy, icon_sz, fill=icon_bg)
        icon_name = str(item.get("icon", ""))
        if icon_name and icon_name in ICON_REGISTRY:
            margin = icon_sz * 0.18
            draw_icon(slide, icon_name,
                      icon_cx + margin, icon_cy + margin,
                      icon_sz - margin * 2, color=WHITE)
        else:
            # Fallback: only render the value as text when it is a single
            # emoji or ≤2-char symbol.  Longer strings are internal icon-key
            # names that the LLM invented; rendering them overflows the small
            # circle, so substitute a neutral dot instead.
            is_short_symbol = icon_name and len(icon_name.strip()) <= 2
            txt(slide, icon_cx, icon_cy, icon_sz, icon_sz,
                (icon_name.strip() if is_short_symbol else "●"),
                sz=18, color=WHITE, font=FONT_BODY,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(slide,
            icon_cx + icon_sz + Inches(0.12), icon_cy,
            card_w - icon_sz - Inches(0.46), icon_sz,
            str(item.get("title", "")),
            sz=12, color=title_col, bold=True, font=FONT_BODY,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        body_y = cy + icon_sz + Inches(0.32)
        txt(slide, cx + Inches(0.18), body_y,
            card_w - Inches(0.36), card_h - icon_sz - Inches(0.46),
            str(item.get("body", "")),
            sz=10, color=body_col, font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

    return slide


