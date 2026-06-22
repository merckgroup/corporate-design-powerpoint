"""Science-specific slide layouts for merck_science deck style.

Four layouts designed for pharma early-research lab progress reports:
  figure_panel  — multi-panel figure grid (2- or 3-column, up to 6 panels)
  methods_box   — experimental conditions table + key result card
  sar_table     — SAR / ADMET wide data table with two-tier headers
  multi_chart   — 2 or 4 small charts on one slide (1×2 or 2×2)
"""
from __future__ import annotations

import pathlib

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from ._ml_constants import (
    FONT_BODY,
    MERCK_BLUE, MERCK_AQUA, MERCK_PURPLE,
    WHITE, INK_DARK, INK_GRAY, LIGHT_GRAY,
    CONTENT_X, CONTENT_W,
    SOURCE_Y,
    _palette_for,
)
from ._ml_primitives import rect, rounded, hairline, txt, _apply_border
from ._ml_deck import _new_slide, add_image
from ._ml_chrome import (
    apply_chrome, _chrome, _content_y,
)
from ._ml_charts import _render_chart

# Pale blue used throughout science layouts (distinct from purple-tinted PANEL_LIGHT).
_SCI_PANEL = (0xEA, 0xF2, 0xFB)
_SCI_MUTED = (0x5A, 0x82, 0xA8)

_SAFE_IMG_EXT = frozenset({
    ".png", ".jpg", ".jpeg", ".gif", ".bmp",
    ".tiff", ".tif", ".webp", ".emf", ".wmf",
})


def _place_img(slide, path_str, x, y, w, h):
    """Embed an image, silently skipping on extension check or read error."""
    if not path_str:
        return
    p = pathlib.Path(path_str).resolve()
    if p.suffix.lower() not in _SAFE_IMG_EXT:
        return
    try:
        add_image(slide, str(p), x, y, w=w, h=h)
    except Exception:
        pass


# ===========================================================================
# Layout: FIGURE PANEL
# ===========================================================================

def build_figure_panel(prs, meta, action_title=None, panels=None, columns=2,
                       takeaway=None, source=None, subtitle=None,
                       style="merck_science",
                       page=None, total=None, section_number=None,
                       category=None, methodology_note=None, content=None):
    """Multi-panel figure grid for showing experimental images or plots.

    Each panel renders an image placeholder (or embedded image if image_path
    is provided) with a bold panel label ("A", "B", ...) and a caption line.

    Content schema:
        panels:  [{label, caption, image_path (optional)}]   max 6
        columns: 2 (default) or 3
    """
    if content:
        panels  = content.get("panels",  panels  or [])
        columns = content.get("columns", columns)

    panels  = list(panels or [])[:6]
    columns = max(1, min(3, int(columns or 2)))

    pal   = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])

    apply_chrome(slide, meta, action_title,
                 category=category, subtitle=subtitle,
                 takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    if not panels:
        return slide

    # Content zone.
    zone_top = _content_y(meta, subtitle=bool(subtitle))
    zone_bot = SOURCE_Y - (Inches(0.30) if (takeaway and _chrome(meta, "takeaway_bands")) else Inches(0.0))
    zone_h   = zone_bot - zone_top

    n_panels  = len(panels)
    n_rows    = max(1, -(-n_panels // columns))   # ceiling division
    h_gap     = Inches(0.18)
    v_gap     = Inches(0.18)
    caption_h = Inches(0.32)
    label_h   = Inches(0.24)

    panel_w = (CONTENT_W - h_gap * (columns - 1)) / columns
    panel_h = (zone_h - v_gap * (n_rows - 1)) / n_rows
    img_h   = panel_h - caption_h - label_h

    for i, panel in enumerate(panels):
        row = i // columns
        col = i % columns
        px = CONTENT_X + col * (panel_w + h_gap)
        py = zone_top + row * (panel_h + v_gap)

        # Image area — pale blue placeholder or embedded image.
        img_path = panel.get("image_path") or panel.get("path")
        if img_path:
            _place_img(slide, img_path, px, py, panel_w, img_h)
        else:
            img_shp = rect(slide, px, py, panel_w, img_h,
                           fill=_SCI_PANEL, border=MERCK_AQUA,
                           border_w=Pt(0.75))
            txt(slide, px, py, panel_w, img_h,
                "[Image placeholder]", sz=9, color=_SCI_MUTED,
                italic=True, font=FONT_BODY,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Panel label "(A)" at top-left of image area.
        label_str = f"({panel.get('label', chr(65 + i))})"
        txt(slide, px + Inches(0.08), py + Inches(0.06),
            Inches(0.50), label_h,
            label_str, sz=10, color=MERCK_BLUE, bold=True, font=FONT_BODY,
            anchor=MSO_ANCHOR.TOP)

        # Caption below the image area.
        caption = str(panel.get("caption") or "")
        if caption:
            txt(slide, px, py + img_h + label_h * 0.1,
                panel_w, caption_h,
                caption, sz=9, color=INK_GRAY, font=FONT_BODY,
                anchor=MSO_ANCHOR.TOP)

    return slide


# ===========================================================================
# Layout: METHODS BOX
# ===========================================================================

def build_methods_box(prs, meta, action_title=None, conditions=None,
                      result=None, takeaway=None, source=None,
                      subtitle=None, style="merck_science",
                      page=None, total=None, section_number=None,
                      category=None, methodology_note=None, content=None):
    """Experimental conditions table (left) + key result card (right).

    Content schema:
        conditions: [{key, value}]           max 12 rows
        result:     {label, value, note}     the headline outcome
    """
    if content:
        conditions = content.get("conditions", conditions or [])
        result     = content.get("result",     result)

    conditions = list(conditions or [])[:12]

    pal   = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])

    apply_chrome(slide, meta, action_title,
                 category=category, subtitle=subtitle,
                 takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    zone_top = _content_y(meta, subtitle=bool(subtitle))
    zone_bot = SOURCE_Y - (Inches(0.30) if (takeaway and _chrome(meta, "takeaway_bands")) else Inches(0.0))
    zone_h   = zone_bot - zone_top

    # Left panel: conditions table (0.65" to 8.40").
    left_x = CONTENT_X
    left_w = Inches(7.60)
    # Right panel: result card (8.90" to 12.65").
    right_x = Inches(8.90)
    right_w = Inches(3.75)

    # ---- Left: conditions table ----
    hdr_h  = Inches(0.30)
    row_h  = Inches(0.28) if len(conditions) > 8 else Inches(0.32)
    key_w  = left_w * 0.42
    val_w  = left_w - key_w

    # Header row.
    rect(slide, left_x, zone_top, left_w, hdr_h,
         fill=MERCK_BLUE, border=None)
    txt(slide, left_x + Inches(0.10), zone_top, key_w, hdr_h,
        "PARAMETER", sz=9, color=WHITE, bold=True, font=FONT_BODY,
        anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, left_x + key_w, zone_top, val_w - Inches(0.05), hdr_h,
        "VALUE / CONDITION", sz=9, color=WHITE, bold=True, font=FONT_BODY,
        anchor=MSO_ANCHOR.MIDDLE)

    # Data rows with alternating pale-blue shading.
    for i, cond in enumerate(conditions):
        ry = zone_top + hdr_h + i * row_h
        row_fill = _SCI_PANEL if i % 2 == 1 else WHITE
        rect(slide, left_x, ry, left_w, row_h,
             fill=row_fill, border=LIGHT_GRAY, border_w=Pt(0.5))
        txt(slide, left_x + Inches(0.08), ry, key_w - Inches(0.08), row_h,
            str(cond.get("key", "")), sz=9, color=INK_DARK,
            font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)
        txt(slide, left_x + key_w + Inches(0.05), ry,
            val_w - Inches(0.10), row_h,
            str(cond.get("value", "")), sz=9, color=INK_GRAY, font=FONT_BODY,
            anchor=MSO_ANCHOR.MIDDLE)

    # ---- Right: key result card ----
    if result:
        card_pad = Inches(0.20)
        card_shp = rounded(slide, right_x, zone_top, right_w, zone_h,
                            fill=_SCI_PANEL, adj=20000)
        _apply_border(card_shp, MERCK_BLUE, Pt(1.0))

        label_text = str(result.get("label", "KEY RESULT"))
        txt(slide, right_x + card_pad, zone_top + Inches(0.20),
            right_w - card_pad * 2, Inches(0.30),
            label_text, sz=9, color=MERCK_BLUE, bold=True,
            font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

        value_text = str(result.get("value", ""))
        txt(slide, right_x + card_pad, zone_top + Inches(0.60),
            right_w - card_pad * 2, Inches(1.20),
            value_text, sz=28, color=INK_DARK, bold=True,
            font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

        note_text = str(result.get("note", ""))
        if note_text:
            txt(slide, right_x + card_pad,
                zone_top + Inches(0.60) + Inches(1.30),
                right_w - card_pad * 2, Inches(0.60),
                note_text, sz=9, color=INK_GRAY, italic=True,
                font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

    return slide


# ===========================================================================
# Layout: SAR TABLE
# ===========================================================================

def build_sar_table(prs, meta, action_title=None, header_groups=None,
                    columns=None, rows=None,
                    takeaway=None, source=None, subtitle=None,
                    style="merck_science",
                    page=None, total=None, section_number=None,
                    category=None, methodology_note=None, content=None):
    """Wide data table for SAR / ADMET data with optional two-tier headers.

    Rich format (two-tier headers):
        header_groups: [{label, columns: [str]}]
        rows:          [{label, values: [str], highlighted: bool}]

    Simple format:
        columns: [str]
        rows:    [{label, values: [str], highlighted: bool}]

    Highlighted rows receive a blue left-edge accent bar and bold label.
    """
    if content:
        header_groups = content.get("header_groups", header_groups)
        columns       = content.get("columns",       columns)
        rows          = content.get("rows",          rows or [])

    rows = list(rows or [])

    pal   = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])

    apply_chrome(slide, meta, action_title,
                 category=category, subtitle=subtitle,
                 takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    zone_top = _content_y(meta, subtitle=bool(subtitle))
    zone_bot = SOURCE_Y - (Inches(0.30) if (takeaway and _chrome(meta, "takeaway_bands")) else Inches(0.0))

    # Normalise to flat column list + optional group spans.
    has_groups  = bool(header_groups)
    if has_groups:
        flat_cols   = []
        group_spans = []
        for g in (header_groups or []):
            cols = g.get("columns", [])
            group_spans.append((str(g.get("label", "")), len(cols)))
            flat_cols.extend(str(c) for c in cols)
    else:
        flat_cols   = [str(c) for c in (columns or [])]
        group_spans = []

    n_data_cols  = len(flat_cols)
    n_total_cols = n_data_cols + 1   # +1 for the row-label column

    if n_total_cols < 2:
        return slide

    # Column widths.
    table_x     = CONTENT_X
    table_w     = CONTENT_W
    label_col_w = min(Inches(2.20), table_w * 0.25)
    data_col_w  = (table_w - label_col_w) / max(n_data_cols, 1)

    # Row heights.
    group_hdr_h = Inches(0.26) if has_groups else Inches(0.0)
    col_hdr_h   = Inches(0.28)
    n_rows_data = len(rows)
    avail_h     = zone_bot - zone_top - group_hdr_h - col_hdr_h
    data_row_h  = (
        min(Inches(0.28), avail_h / max(n_rows_data, 1))
        if n_rows_data else Inches(0.28)
    )

    cur_y = zone_top

    # ---- Group header row (optional) ----
    if has_groups:
        rect(slide, table_x, cur_y, label_col_w, group_hdr_h,
             fill=MERCK_BLUE, border=None)
        gx = table_x + label_col_w
        for g_label, g_ncols in group_spans:
            gw = data_col_w * g_ncols
            rect(slide, gx, cur_y, gw, group_hdr_h,
                 fill=MERCK_BLUE, border=MERCK_PURPLE, border_w=Pt(0.5))
            if g_label:
                txt(slide, gx + Inches(0.06), cur_y, gw - Inches(0.06), group_hdr_h,
                    g_label, sz=8, color=WHITE, bold=True,
                    font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)
            gx += gw
        cur_y += group_hdr_h

    # ---- Column header row ----
    rect(slide, table_x, cur_y, label_col_w, col_hdr_h,
         fill=_SCI_PANEL, border=MERCK_AQUA, border_w=Pt(0.5))
    txt(slide, table_x + Inches(0.06), cur_y, label_col_w, col_hdr_h,
        "COMPOUND / ID", sz=8, color=MERCK_BLUE, bold=True,
        font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)
    for ci, col_label in enumerate(flat_cols):
        cx = table_x + label_col_w + ci * data_col_w
        rect(slide, cx, cur_y, data_col_w, col_hdr_h,
             fill=_SCI_PANEL, border=MERCK_AQUA, border_w=Pt(0.5))
        txt(slide, cx + Inches(0.04), cur_y, data_col_w - Inches(0.04), col_hdr_h,
            col_label, sz=8, color=MERCK_BLUE, bold=True,
            font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cur_y += col_hdr_h

    # ---- Data rows ----
    for ri, row in enumerate(rows):
        row_fill    = _SCI_PANEL if ri % 2 == 1 else WHITE
        highlighted = bool(row.get("highlighted"))
        row_label   = str(row.get("label", ""))
        values      = [str(v) for v in row.get("values", [])]

        # Label cell.
        rect(slide, table_x, cur_y, label_col_w, data_row_h,
             fill=row_fill, border=LIGHT_GRAY, border_w=Pt(0.4))
        if highlighted:
            rect(slide, table_x, cur_y, Inches(0.06), data_row_h,
                 fill=MERCK_BLUE, border=None)
        txt(slide, table_x + Inches(0.10), cur_y,
            label_col_w - Inches(0.12), data_row_h,
            row_label, sz=9,
            color=MERCK_BLUE if highlighted else INK_DARK,
            bold=highlighted, font=FONT_BODY, anchor=MSO_ANCHOR.MIDDLE)

        # Data cells.
        for ci in range(n_data_cols):
            cx  = table_x + label_col_w + ci * data_col_w
            val = values[ci] if ci < len(values) else ""
            rect(slide, cx, cur_y, data_col_w, data_row_h,
                 fill=row_fill, border=LIGHT_GRAY, border_w=Pt(0.4))
            txt(slide, cx + Inches(0.04), cur_y,
                data_col_w - Inches(0.04), data_row_h,
                val, sz=9, color=INK_GRAY, font=FONT_BODY,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        cur_y += data_row_h

    return slide


# ===========================================================================
# Layout: MULTI CHART
# ===========================================================================

def build_multi_chart(prs, meta, action_title=None, charts=None,
                      layout="1x2",
                      takeaway=None, source=None, subtitle=None,
                      style="merck_science",
                      page=None, total=None, section_number=None,
                      category=None, methodology_note=None, content=None):
    """Two or four small charts on one slide for side-by-side comparison.

    Content schema:
        charts: [{title, chart: {type, data}}]   2 or 4 entries
        layout: "1x2" (default) or "2x2"

    Each chart entry's `chart` field uses the same schema as
    chart_slide.content.chart.
    """
    if content:
        charts = content.get("charts", charts or [])
        layout = content.get("layout", layout or "1x2")

    charts = list(charts or [])
    layout = str(layout or "1x2").strip().lower()

    pal   = _palette_for(style)
    slide = _new_slide(prs, bg_color=pal["bg"])

    apply_chrome(slide, meta, action_title,
                 category=category, subtitle=subtitle,
                 takeaway=takeaway, source=source,
                 page=page, total=total, palette=style,
                 section_number=section_number,
                 methodology_note=methodology_note)

    if not charts:
        return slide

    zone_top = _content_y(meta, subtitle=bool(subtitle))
    zone_bot = SOURCE_Y - (Inches(0.30) if (takeaway and _chrome(meta, "takeaway_bands")) else Inches(0.0))
    zone_h   = zone_bot - zone_top

    h_gap   = Inches(0.20)
    v_gap   = Inches(0.20)
    title_h = Inches(0.25)

    if layout == "2x2":
        n_cols, n_rows = 2, 2
        charts = charts[:4]
    else:
        n_cols, n_rows = 2, 1
        charts = charts[:2]

    chart_w = (CONTENT_W - h_gap * (n_cols - 1)) / n_cols
    row_h   = (zone_h - v_gap * (n_rows - 1)) / n_rows
    chart_h = row_h - title_h

    for i, entry in enumerate(charts):
        row = i // n_cols
        col = i % n_cols
        cx  = CONTENT_X + col * (chart_w + h_gap)
        cy  = zone_top + row * (row_h + v_gap)

        # Per-chart title.
        chart_title = str(entry.get("title", ""))
        if chart_title:
            txt(slide, cx, cy, chart_w, title_h,
                chart_title, sz=10, color=MERCK_BLUE, bold=True,
                font=FONT_BODY, anchor=MSO_ANCHOR.TOP)

        chart_spec = entry.get("chart") or {}
        rendered = False
        if chart_spec:
            try:
                rendered = _render_chart(slide, chart_spec,
                                         cx, cy + title_h,
                                         chart_w, chart_h, style)
            except Exception:
                rendered = False

        if not rendered:
            rect(slide, cx, cy + title_h, chart_w, chart_h,
                 fill=_SCI_PANEL, border=MERCK_AQUA, border_w=Pt(0.75))
            txt(slide, cx, cy + title_h, chart_w, chart_h,
                "[Chart placeholder]", sz=9, color=_SCI_MUTED, italic=True,
                font=FONT_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    return slide
