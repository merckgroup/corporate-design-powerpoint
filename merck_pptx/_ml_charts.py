from __future__ import annotations
import math
import sys
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt
from ._ml_constants import (
    FONT_HEAD, FONT_BODY, WHITE, INK_DARK, INK_GRAY, LIGHT_GRAY, PANEL_LIGHT,
    MERCK_PURPLE, MERCK_GOLD, MERCK_YELLOW, MERCK_BLUE, MERCK_AQUA,
    PURPLE_MUTED, BAD_RED, GOOD_GREEN,
    ACT_PURPLE, LY_CYAN, OP_LIME, FC_PINK, DEV_POS_BLUE, DEV_NEG_RED,
    CHART_PALETTE, PHASE_1_COLOR, PHASE_2_COLOR, PHASE_3_COLOR,
    _palette_for, _rgb_tuple, _is_dark,
)
from ._ml_primitives import (
    rect, rounded, oval, circle, line, hairline, txt, _add_run,
    _freeform_poly, _emu, _apply_fill, _apply_border,
)
from ._ml_chrome import apply_chrome, stub_and_flag, footnotes_block

# ===========================================================================
# Chart helpers
# ===========================================================================

def _scale(value, vmin, vmax, lo, hi):
    if vmax == vmin:
        return lo + (hi - lo) / 2
    return lo + (value - vmin) * (hi - lo) / (vmax - vmin)


def add_slope_chart(slide, x, y, w, h, before_label, after_label, items,
                    palette, highlight_indices=None):
    """Slope chart with anti-collision label stacking on both sides."""
    pal = _palette_for(palette)
    if not items:
        return
    highlight_set = set(highlight_indices or [])

    label_gutter = Inches(1.7)
    left_axis = x + label_gutter
    right_axis = x + w - label_gutter
    plot_top = y + Inches(0.65)
    plot_bot = y + h - Inches(0.20)

    vals = []
    for it in items:
        vals.append(it[1])
        vals.append(it[2])
    vmin = min(vals)
    vmax = max(vals)
    if vmin == vmax:
        vmin -= 1
        vmax += 1

    # Gold "Before" / "After" headers at top of axes.
    txt(slide, left_axis - Inches(1.1), plot_top - Inches(0.55),
        Inches(2.2), Inches(0.30),
        _tracked(before_label), sz=11, color=pal["highlight"], bold=True,
        font=FONT_BODY, align=PP_ALIGN.CENTER)
    txt(slide, right_axis - Inches(1.1), plot_top - Inches(0.55),
        Inches(2.2), Inches(0.30),
        _tracked(after_label), sz=11, color=pal["highlight"], bold=True,
        font=FONT_BODY, align=PP_ALIGN.CENTER)

    # Axis hairlines.
    hairline(slide, left_axis, plot_top, Emu(int(Pt(0.75))),
             plot_bot - plot_top, LIGHT_GRAY)
    hairline(slide, right_axis, plot_top, Emu(int(Pt(0.75))),
             plot_bot - plot_top, LIGHT_GRAY)

    # Natural positions on each side, with index.
    left_entries = []
    right_entries = []
    for idx, item in enumerate(items):
        label, bv, av = item[0], item[1], item[2]
        y1 = _scale(bv, vmin, vmax, plot_bot, plot_top)
        y2 = _scale(av, vmin, vmax, plot_bot, plot_top)
        left_entries.append({"idx": idx, "label": label, "value": bv, "y": y1})
        right_entries.append({"idx": idx, "label": label, "value": av, "y": y2})

    def _enforce_gap(entries, min_gap_emu, lo, hi):
        if not entries:
            return entries
        entries = sorted(entries, key=lambda e: e["y"])
        # First pass: push down if too close to prior.
        for i in range(1, len(entries)):
            if entries[i]["y"] - entries[i - 1]["y"] < min_gap_emu:
                entries[i]["y"] = entries[i - 1]["y"] + min_gap_emu
        # Clamp to hi.
        if entries[-1]["y"] > hi:
            entries[-1]["y"] = hi
            for i in range(len(entries) - 2, -1, -1):
                if entries[i + 1]["y"] - entries[i]["y"] < min_gap_emu:
                    entries[i]["y"] = entries[i + 1]["y"] - min_gap_emu
        # Clamp to lo.
        if entries[0]["y"] < lo:
            entries[0]["y"] = lo
            for i in range(1, len(entries)):
                if entries[i]["y"] - entries[i - 1]["y"] < min_gap_emu:
                    entries[i]["y"] = entries[i - 1]["y"] + min_gap_emu
        return entries

    min_gap = Inches(0.24)
    label_lo = plot_top - Inches(0.10)
    label_hi = plot_bot + Inches(0.10)
    left_entries = _enforce_gap(left_entries, min_gap, label_lo, label_hi)
    right_entries = _enforce_gap(right_entries, min_gap, label_lo, label_hi)

    # Map back by idx for label positions.
    left_label_y = {e["idx"]: e["y"] for e in left_entries}
    right_label_y = {e["idx"]: e["y"] for e in right_entries}

    # Draw slope lines and endpoints using NATURAL positions (not adjusted).
    for idx, item in enumerate(items):
        label, bv, av = item[0], item[1], item[2]
        y1 = _scale(bv, vmin, vmax, plot_bot, plot_top)
        y2 = _scale(av, vmin, vmax, plot_bot, plot_top)
        highlighted = idx in highlight_set
        if highlighted:
            line_color = MERCK_YELLOW
            weight = Pt(2.5)
            dot_d = Inches(0.14)
        else:
            line_color = PURPLE_MUTED
            weight = Pt(1.25)
            dot_d = Inches(0.10)
        line(slide, left_axis, y1, right_axis, y2, line_color, weight)
        circle(slide, left_axis - dot_d / 2, y1 - dot_d / 2, dot_d, fill=line_color)
        circle(slide, right_axis - dot_d / 2, y2 - dot_d / 2, dot_d, fill=line_color)

    # Render labels at adjusted positions.
    for idx, item in enumerate(items):
        label, bv, av = item[0], item[1], item[2]
        highlighted = idx in highlight_set
        text_color = MERCK_PURPLE if highlighted else PURPLE_MUTED
        bold = highlighted
        # Left label: "label  value", right-aligned to left axis.
        ly = left_label_y[idx]
        txt(slide, x, ly - Inches(0.11), left_axis - x - Inches(0.20),
            Inches(0.22),
            f"{label}   {bv}", sz=10, color=text_color, bold=bold,
            font=FONT_BODY, align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE)
        # Right label: "value  label", left-aligned to right axis.
        ry = right_label_y[idx]
        txt(slide, right_axis + Inches(0.20), ry - Inches(0.11),
            x + w - right_axis - Inches(0.20), Inches(0.22),
            f"{av}   {label}", sz=10, color=text_color, bold=bold,
            font=FONT_BODY, align=PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.MIDDLE)


def add_dot_plot(slide, x, y, w, h, categories, items, palette):
    """Horizontal dot plot."""
    pal = _palette_for(palette)
    if not items:
        return
    left = x + Inches(2.2)
    right = x + w - Inches(0.4)
    axis_y = y + h - Inches(0.40)
    top = y + Inches(0.20)
    rows = len(items)
    row_h = (axis_y - top) / max(rows, 1)

    hairline(slide, left, axis_y, right - left, Emu(int(Pt(0.75))), LIGHT_GRAY)
    try:
        cat_nums = [float(str(c).replace("%", "").strip()) for c in categories]
        vmin = min(cat_nums)
        vmax = max(cat_nums)
    except Exception:
        cat_nums = list(range(len(categories)))
        vmin = 0
        vmax = len(categories) - 1
    for i, c in enumerate(categories):
        tx = _scale(cat_nums[i], vmin, vmax, left, right)
        line(slide, tx, axis_y, tx, axis_y + Inches(0.08), LIGHT_GRAY, Pt(0.5))
        txt(slide, tx - Inches(0.4), axis_y + Inches(0.10),
            Inches(0.8), Inches(0.24),
            str(c), sz=9, color=INK_GRAY, font=FONT_BODY, align=PP_ALIGN.CENTER)

    for i, item in enumerate(items):
        label, value = item[0], item[1]
        ry = top + row_h * i + row_h / 2
        line(slide, left, ry, right, ry, LIGHT_GRAY, Pt(0.5))
        txt(slide, x, ry - Inches(0.13), left - x - Inches(0.10),
            Inches(0.26),
            str(label), sz=10, color=INK_DARK, font=FONT_BODY,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        try:
            v = float(str(value).replace("%", "").strip())
        except Exception:
            v = vmin
        dx = _scale(v, vmin, vmax, left, right)
        dot = Inches(0.18)
        circle(slide, dx - dot / 2, ry - dot / 2, dot, fill=MERCK_PURPLE)
        txt(slide, dx + Inches(0.12), ry - Inches(0.12),
            Inches(1.0), Inches(0.24),
            str(value), sz=9, color=PURPLE_DEEP, bold=True, font=FONT_BODY,
            anchor=MSO_ANCHOR.MIDDLE)


def add_marimekko(slide, x, y, w, h, columns, palette):
    """Mosaic chart with proportional column widths and stacked segments.

    Each column has a top label (uppercase tracked) and a bottom share label.
    Segments inside each column are filled with the brand palette in order;
    a label and percentage are rendered inside if the segment is tall enough,
    otherwise an external callout-style label is drawn beside the column.
    A consistent legend appears below the chart.
    """
    pal = _palette_for(palette)
    if not columns:
        return
    palette_colors = CHART_PALETTE
    total_weight = sum(max(c.get("weight", 0), 0) for c in columns) or 1.0
    label_h = Inches(0.30)
    share_h = Inches(0.26)
    legend_h = Inches(0.30)
    top = y + label_h
    bot = y + h - share_h - legend_h - Inches(0.10)
    col_h = bot - top
    gap = Inches(0.03)
    n_cols = len(columns)

    # Build a unified legend from segment label order across columns.
    legend_keys = []
    seen = set()
    for col in columns:
        for seg in col.get("segments", []):
            key = seg[0]
            if key not in seen:
                seen.add(key)
                legend_keys.append(key)
    color_by_key = {k: palette_colors[i % len(palette_colors)]
                    for i, k in enumerate(legend_keys)}

    cx = x
    for ci, col in enumerate(columns):
        weight = max(col.get("weight", 0), 0) / total_weight
        cw = w * weight - (gap if ci < n_cols - 1 else 0)
        # Column header in MERCK_PURPLE tracked bold.
        txt(slide, cx, y, cw + (gap if ci < n_cols - 1 else 0), label_h,
            _tracked(col.get("label", "")), sz=10,
            color=MERCK_PURPLE, bold=True, font=FONT_BODY,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        segs = col.get("segments", [])
        seg_total = sum(max(s[1], 0) for s in segs) or 1.0

        sy = top
        for si, seg in enumerate(segs):
            seg_label, pct = seg[0], max(seg[1], 0)
            sh = col_h * (pct / seg_total)
            color = color_by_key.get(seg_label,
                                     palette_colors[si % len(palette_colors)])
            rect(slide, cx, sy, cw, sh, fill=color)
            # Inside label only when the segment is tall AND wide enough.
            if sh >= Inches(0.40) and cw >= Inches(0.80):
                txt(slide, cx + Inches(0.06), sy, cw - Inches(0.12), sh,
                    f"{seg_label}  {int(round(pct))}%", sz=10, color=WHITE,
                    bold=True, font=FONT_BODY, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
            elif sh >= Inches(0.18) and cw >= Inches(0.60):
                txt(slide, cx + Inches(0.04), sy, cw - Inches(0.08), sh,
                    f"{int(round(pct))}%", sz=9, color=WHITE,
                    bold=True, font=FONT_BODY, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
            sy += sh

        # Share label below the column (% of total weight).
        share_pct = int(round(weight * 100))
        txt(slide, cx, bot + Inches(0.06),
            cw + (gap if ci < n_cols - 1 else 0), share_h,
            f"{share_pct}% of total", sz=9,
            color=INK_GRAY, italic=True, font=FONT_BODY, align=PP_ALIGN.CENTER)
        cx += cw + gap

    # Unified legend below the share row.
    legend_y = bot + share_h + Inches(0.12)
    if legend_keys:
        swatch = Inches(0.16)
        # Estimate text widths so chips don't overlap.
        chip_pad = Inches(0.20)
        chips = []
        for k in legend_keys:
            chip_w = swatch + Inches(0.08) + max(Inches(0.6),
                                                 Inches(0.10) * (len(str(k)) + 2))
            chips.append((k, chip_w))
        total_chip_w = sum(c[1] for c in chips) + chip_pad * (len(chips) - 1)
        lx = x + (w - total_chip_w) / 2
        for k, chip_w in chips:
            rect(slide, lx, legend_y + (legend_h - swatch) / 2,
                 swatch, swatch, fill=color_by_key.get(k, MERCK_PURPLE))
            txt(slide, lx + swatch + Inches(0.06), legend_y,
                chip_w - swatch - Inches(0.06), legend_h,
                str(k), sz=9, color=INK_DARK, font=FONT_BODY,
                anchor=MSO_ANCHOR.MIDDLE)
            lx += chip_w + chip_pad


def add_waterfall(slide, x, y, w, h, bars, palette):
    """Waterfall chart with floating bars at running totals.

    bars: list of {label, value, type} where type is start, up, down, end.
    """
    pal = _palette_for(palette)
    if not bars:
        return
    n = len(bars)
    plot_top = y + Inches(0.55)
    plot_bot = y + h - Inches(0.60)
    gap = Inches(0.10)
    bar_w = (w - gap * (n + 1)) / n

    # Compute running totals.
    running = 0.0
    levels = []  # (low, high) for each bar
    running_totals = []
    for b in bars:
        t = b.get("type", "up")
        v = float(b.get("value", 0))
        if t == "start":
            low, high = 0.0, v
            running = v
        elif t == "end":
            low, high = 0.0, v
            running = v
        elif t == "down":
            high = running
            running -= abs(v)
            low = running
        else:  # up
            low = running
            running += v
            high = running
        levels.append((low, high))
        running_totals.append(running)

    all_y = []
    for lo, hi in levels:
        all_y.append(lo)
        all_y.append(hi)
    vmin = min(all_y + [0.0])
    vmax = max(all_y + [0.0]) * 1.10
    if vmax == vmin:
        vmax = vmin + 1.0

    baseline_y = _scale(0, vmin, vmax, plot_bot, plot_top)
    hairline(slide, x, baseline_y, w, Emu(int(Pt(0.75))), LIGHT_GRAY)

    # Connector dotted lines between adjacent bars at their running totals.
    bx = x + gap
    for i, b in enumerate(bars):
        if i < n - 1:
            t = b.get("type", "up")
            next_t = bars[i + 1].get("type", "up")
            connect_y = _scale(running_totals[i], vmin, vmax, plot_bot, plot_top)
            cx1 = bx + bar_w
            cx2 = bx + bar_w + gap
            # Use a thin dashed-look hairline by drawing tiny segments.
            seg_w = Inches(0.05)
            ax = cx1
            while ax < cx2:
                ax2 = min(ax + seg_w, cx2)
                line(slide, ax, connect_y, ax2, connect_y, PURPLE_MUTED, Pt(0.5))
                ax += seg_w * 2
        bx += bar_w + gap

    # Draw bars.
    bx = x + gap
    for i, b in enumerate(bars):
        t = b.get("type", "up")
        label = b.get("label", "")
        v = float(b.get("value", 0))
        low, high = levels[i]
        y_high = _scale(high, vmin, vmax, plot_bot, plot_top)
        y_low = _scale(low, vmin, vmax, plot_bot, plot_top)
        top_px = min(y_high, y_low)
        bar_h = abs(y_high - y_low)
        bar_h = max(bar_h, Inches(0.04))

        if t in ("start", "end"):
            color = MERCK_PURPLE
            shown = f"{v:g}"
            label_above = True
            label_color = MERCK_PURPLE
        elif t == "down":
            color = BAD_RED
            shown = f"-{abs(v):g}"
            label_above = False
            label_color = BAD_RED
        else:  # up — GOOD_GREEN for positive/growth bars (CD traffic-light convention)
            color = GOOD_GREEN
            shown = f"+{v:g}"
            label_above = True
            label_color = MERCK_PURPLE

        rect(slide, bx, top_px, bar_w, bar_h, fill=color)

        # Value label (always one line; above for up/start/end, below for down).
        if label_above:
            ly = top_px - Inches(0.32)
            anchor = MSO_ANCHOR.BOTTOM
        else:
            ly = top_px + bar_h + Inches(0.04)
            anchor = MSO_ANCHOR.TOP
        txt(slide, bx - Inches(0.30), ly, bar_w + Inches(0.60), Inches(0.28),
            shown, sz=10, color=label_color, bold=True, font=FONT_BODY,
            align=PP_ALIGN.CENTER, anchor=anchor)

        # X-axis label (with optional footnote superscript).
        fn = b.get("footnote") if isinstance(b, dict) else None
        label_text = str(label)
        if fn is not None:
            label_text = f"{label_text} {fn}"
        txt(slide, bx - Inches(0.20), plot_bot + Inches(0.08),
            bar_w + Inches(0.40), Inches(0.40),
            label_text, sz=9, color=INK_DARK, font=FONT_BODY,
            align=PP_ALIGN.CENTER)
        bx += bar_w + gap


def add_small_multiples(slide, x, y, w, h, datasets, palette, grid=(2, 3)):
    """Grid of mini line charts with titles and shared visual rhythm.

    Each cell is a small panel with:
      - a card chrome background (PANEL_LIGHT)
      - a title in MERCK_PURPLE bold
      - a baseline hairline at the bottom of the plot
      - left-edge value labels (min, max) for context
      - a MERCK_PURPLE line trace with a MERCK_GOLD endpoint dot
      - the latest value as a small label at the endpoint
    """
    pal = _palette_for(palette)
    if not datasets:
        return
    rows, cols = grid
    cell_pad = Inches(0.10)
    cell_w = (w - cell_pad * (cols - 1)) / cols
    cell_h = (h - cell_pad * (rows - 1)) / rows
    pad_x = Inches(0.16)
    title_h = Inches(0.30)
    bottom_h = Inches(0.10)

    for i, ds in enumerate(datasets):
        r = i // cols
        c = i % cols
        if r >= rows:
            break
        cx = x + c * (cell_w + cell_pad)
        cy = y + r * (cell_h + cell_pad)

        # Cell chrome.
        card = rounded(slide, cx, cy, cell_w, cell_h, fill=PANEL_LIGHT)
        _apply_border(card, LIGHT_GRAY, Pt(0.5))
        rect(slide, cx, cy, cell_w, Inches(0.04), fill=pal["highlight"])

        title = ds[0]
        values = list(ds[1] or [])

        # Cell title.
        txt(slide, cx + pad_x, cy + Inches(0.10), cell_w - pad_x * 2,
            title_h, title, sz=10, color=MERCK_PURPLE, bold=True,
            font=FONT_BODY)

        # Plot area.
        px = cx + pad_x + Inches(0.32)  # leave room for left value labels
        py = cy + title_h + Inches(0.12)
        pw = cell_w - pad_x * 2 - Inches(0.34)
        ph = cell_h - title_h - bottom_h - Inches(0.22)
        # Baseline hairline.
        hairline(slide, px, py + ph, pw, Emu(int(Pt(0.5))), LIGHT_GRAY)

        if len(values) >= 2:
            vmin = min(values)
            vmax = max(values)
            if vmin == vmax:
                vmin -= 1
                vmax += 1
            # Y-axis labels (max top, min bottom) on the left.
            txt(slide, cx + Inches(0.05), py - Inches(0.06),
                pad_x + Inches(0.22), Inches(0.18),
                f"{vmax:g}", sz=7, color=INK_GRAY, font=FONT_BODY,
                align=PP_ALIGN.RIGHT)
            txt(slide, cx + Inches(0.05), py + ph - Inches(0.10),
                pad_x + Inches(0.22), Inches(0.18),
                f"{vmin:g}", sz=7, color=INK_GRAY, font=FONT_BODY,
                align=PP_ALIGN.RIGHT)
            pts = []
            for j, v in enumerate(values):
                gx = px + (pw * j / (len(values) - 1))
                gy = _scale(v, vmin, vmax, py + ph - Inches(0.04),
                            py + Inches(0.04))
                pts.append((gx, gy))
            for a, b in zip(pts, pts[1:]):
                line(slide, a[0], a[1], b[0], b[1], MERCK_PURPLE, Pt(1.5))
            # Start dot.
            sdot = Inches(0.07)
            circle(slide, pts[0][0] - sdot / 2, pts[0][1] - sdot / 2,
                   sdot, fill=PURPLE_MUTED)
            # End dot (highlight).
            dot = Inches(0.10)
            circle(slide, pts[-1][0] - dot / 2, pts[-1][1] - dot / 2,
                   dot, fill=pal["highlight"])
            # Endpoint value label, placed inside the cell.
            label_w = Inches(0.70)
            lx = pts[-1][0] - label_w + Inches(0.05)
            ly = pts[-1][1] - Inches(0.26)
            if ly < py:
                ly = py + Inches(0.02)
            txt(slide, lx, ly, label_w, Inches(0.22),
                f"{values[-1]:g}", sz=9, color=PURPLE_DEEP, bold=True,
                font=FONT_BODY, align=PP_ALIGN.RIGHT)


def add_simple_bar(slide, x, y, w, h, items=None, palette=None,
                   horizontal=False, highlight_index=None,
                   series=None, categories=None, chart_type_name="column"):
    """Simple bar / column / line chart as a NATIVE PowerPoint chart.

    Accepts two data formats:
      Legacy format  — items: list of [label, value] pairs (single series).
      Standard format — categories: list of str  +  series: list of
                        {name: str, values: list of float}.

    chart_type_name: "column" | "bar" | "line"  (ignored for legacy callers
    that set horizontal=True, which stays as BAR_CLUSTERED).
    """
    _CHART_TYPES = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar":    XL_CHART_TYPE.BAR_CLUSTERED,
        "line":   XL_CHART_TYPE.LINE,
    }

    # ------------------------------------------------------------------
    # Normalise to (cats, series_list) regardless of which format arrived.
    # ------------------------------------------------------------------
    if series is not None and categories is not None:
        # Standard {categories, series} format.
        cats        = [str(c) for c in categories]
        series_list = list(series)
    elif items:
        # Legacy [label, value] pairs → single series named "Value".
        cats        = [str(it[0]) for it in items]
        series_list = [{"name": "Value", "values": [float(it[1]) for it in items]}]
    else:
        return None

    if not cats or not series_list:
        return None

    chart_data = CategoryChartData()
    chart_data.categories = cats
    for s in series_list:
        vals = [float(v) for v in s.get("values", [])]
        chart_data.add_series(str(s.get("name", "")), vals)

    # ------------------------------------------------------------------
    # Determine PowerPoint chart type.
    # ------------------------------------------------------------------
    if horizontal:
        xl_type = XL_CHART_TYPE.BAR_CLUSTERED
    else:
        xl_type = _CHART_TYPES.get(str(chart_type_name).lower(),
                                   XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_shape = slide.shapes.add_chart(
        xl_type, _emu(x), _emu(y), _emu(w), _emu(h), chart_data
    )
    chart = chart_shape.chart

    # Clean styling — no legend when single series; show for multi-series.
    chart.has_title = False
    chart.has_legend = len(series_list) > 1
    if chart.has_legend:
        try:
            chart.legend.position    = 2  # BOTTOM
            chart.legend.include_in_layout = False
        except Exception:
            pass

    try:
        for axis in (chart.category_axis, chart.value_axis):
            try:
                axis.tick_labels.font.size = Pt(10)
                axis.tick_labels.font.name = FONT_BODY
            except Exception:
                pass
    except Exception:
        pass

    # ------------------------------------------------------------------
    # Color each series.  For a single series use the highlight logic;
    # for multiple series rotate through the Merck palette.
    # ------------------------------------------------------------------
    # Use the official Merck Corporate Design 12-color chart palette so
    # multi-series charts follow CD specifications (orange / teal / dark-cyan …).
    _MULTI_SERIES_COLORS = CHART_PALETTE
    for si, s_obj in enumerate(chart.series):
        series_color = _MULTI_SERIES_COLORS[si % len(_MULTI_SERIES_COLORS)]
        try:
            s_obj.format.line.color.rgb = _rgb_tuple(series_color)
            s_obj.format.fill.solid()
            s_obj.format.fill.fore_color.rgb = _rgb_tuple(series_color)
        except Exception:
            pass
        if si == 0 and highlight_index is not None:
            for i, pt in enumerate(s_obj.points):
                fill = pt.format.fill
                fill.solid()
                fill.fore_color.rgb = _rgb_tuple(
                    MERCK_GOLD if i == int(highlight_index) else PURPLE_MUTED
                )
        elif si == 0 and len(series_list) == 1:
            for pt in s_obj.points:
                fill = pt.format.fill
                fill.solid()
                fill.fore_color.rgb = _rgb_tuple(MERCK_PURPLE)

    # Data labels — only on bar/column charts (not line).
    if xl_type not in (XL_CHART_TYPE.LINE,):
        try:
            from pptx.enum.chart import XL_LABEL_POSITION
            plot = chart.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True
            dl.font.size = Pt(9)
            dl.font.bold = True
            try:
                dl.font.color.rgb = _rgb_tuple(INK_DARK)
            except Exception:
                pass
            try:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass
        except Exception:
            pass

    return chart_shape


# Chart dispatcher used by build_chart_slide

def _render_chart(slide, chart, x, y, w, h, palette):
    """Render chart content onto slide. Returns True if a chart was drawn."""
    if not chart:
        return False
    ctype = chart.get("type", "column")
    data = chart.get("data", {})

    if ctype == "slope":
        add_slope_chart(slide, x, y, w, h,
                        data.get("before_label", "Before"),
                        data.get("after_label", "After"),
                        data.get("items", []),
                        palette,
                        highlight_indices=data.get("highlight_indices"))
        return True
    elif ctype == "dot":
        add_dot_plot(slide, x, y, w, h,
                     data.get("categories", ["0", "25", "50", "75", "100"]),
                     data.get("items", []), palette)
        return True
    elif ctype == "marimekko":
        add_marimekko(slide, x, y, w, h, data.get("columns", []), palette)
        return True
    elif ctype == "waterfall":
        add_waterfall(slide, x, y, w, h, data.get("bars", []), palette)
        return True
    elif ctype == "small_multiples":
        add_small_multiples(slide, x, y, w, h,
                            data.get("datasets", []), palette,
                            grid=tuple(data.get("grid", (2, 3))))
        return True
    elif ctype in ("bar", "column", "line"):
        # Standard {categories, series} format — preferred for LLM-generated plans.
        # Falls back to legacy [label, value] items format when series is absent.
        series_data  = data.get("series")
        categories   = data.get("categories")
        legacy_items = data.get("items")
        if series_data and categories:
            add_simple_bar(slide, x, y, w, h,
                           categories=categories,
                           series=series_data,
                           palette=palette,
                           horizontal=(ctype == "bar"),
                           highlight_index=data.get("highlight_index"),
                           chart_type_name=ctype)
            return True
        elif legacy_items:
            add_simple_bar(slide, x, y, w, h,
                           items=legacy_items,
                           palette=palette,
                           horizontal=(ctype == "bar"),
                           highlight_index=data.get("highlight_index"),
                           chart_type_name=ctype)
            return True
        else:
            import sys
            print(f"WARNING: chart_slide received type='{ctype}' but no "
                  f"'series'+'categories' or 'items' data found — slide "
                  f"will render without a chart.", file=sys.stderr)
            return False
    else:
        # Unknown type: attempt legacy items format and warn.
        import sys
        items = data.get("items", [])
        if items:
            add_simple_bar(slide, x, y, w, h,
                           items=items, palette=palette,
                           horizontal=bool(data.get("horizontal", False)),
                           highlight_index=data.get("highlight_index"))
            return True
        else:
            print(f"WARNING: chart_slide received unknown chart type "
                  f"'{ctype}' with no renderable data.", file=sys.stderr)
            return False

