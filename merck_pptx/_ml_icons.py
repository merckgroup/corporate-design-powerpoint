from __future__ import annotations
from pptx.util import Emu, Pt
from ._ml_constants import WHITE
from ._ml_primitives import rect, rounded, oval, circle, line, hairline, _freeform_poly

# ===========================================================================
# Vector icons (unchanged structure; pen-quality silhouettes)
# ===========================================================================

def icon_chart_bar(slide, x, y, size, color):
    n = 3
    gap = size / 9
    bar_w = (size - (n + 1) * gap) / n
    base_y = y + size * 0.95
    heights = [size * 0.40, size * 0.60, size * 0.85]
    for i in range(n):
        bx = x + gap + i * (bar_w + gap)
        by = base_y - heights[i]
        rect(slide, bx, by, bar_w, heights[i], fill=color)
    hairline(slide, x + size * 0.05, base_y, size * 0.90, Emu(int(Pt(0.75))), color)


def icon_chart_line(slide, x, y, size, color):
    pts = [
        (x + size * 0.08, y + size * 0.80),
        (x + size * 0.32, y + size * 0.55),
        (x + size * 0.55, y + size * 0.65),
        (x + size * 0.92, y + size * 0.20),
    ]
    for a, b in zip(pts, pts[1:]):
        line(slide, a[0], a[1], b[0], b[1], color, Pt(1.5))
    dot = size * 0.10
    circle(slide, pts[-1][0] - dot / 2, pts[-1][1] - dot / 2, dot, fill=color)


def icon_chart_pie(slide, x, y, size, color):
    oval(slide, x, y, size, size, fill=color)
    cx = x + size / 2
    cy = y + size / 2
    pts = [(cx, cy), (cx + size / 2, cy), (cx + size / 2, cy + size / 2),
           (cx, cy + size / 2)]
    _freeform_poly(slide, pts, fill=WHITE)


def icon_arrow_up(slide, x, y, size, color):
    pts = [
        (x + size / 2, y + size * 0.08),
        (x + size * 0.90, y + size * 0.65),
        (x + size * 0.65, y + size * 0.65),
        (x + size * 0.65, y + size * 0.92),
        (x + size * 0.35, y + size * 0.92),
        (x + size * 0.35, y + size * 0.65),
        (x + size * 0.10, y + size * 0.65),
    ]
    _freeform_poly(slide, pts, fill=color)


def icon_arrow_down(slide, x, y, size, color):
    pts = [
        (x + size / 2, y + size * 0.92),
        (x + size * 0.90, y + size * 0.35),
        (x + size * 0.65, y + size * 0.35),
        (x + size * 0.65, y + size * 0.08),
        (x + size * 0.35, y + size * 0.08),
        (x + size * 0.35, y + size * 0.35),
        (x + size * 0.10, y + size * 0.35),
    ]
    _freeform_poly(slide, pts, fill=color)


def icon_arrow_right(slide, x, y, size, color):
    pts = [
        (x + size * 0.92, y + size / 2),
        (x + size * 0.55, y + size * 0.10),
        (x + size * 0.55, y + size * 0.35),
        (x + size * 0.08, y + size * 0.35),
        (x + size * 0.08, y + size * 0.65),
        (x + size * 0.55, y + size * 0.65),
        (x + size * 0.55, y + size * 0.90),
    ]
    _freeform_poly(slide, pts, fill=color)


def icon_check(slide, x, y, size, color):
    p1 = (x + size * 0.10, y + size * 0.55)
    p2 = (x + size * 0.40, y + size * 0.85)
    p3 = (x + size * 0.90, y + size * 0.18)
    line(slide, p1[0], p1[1], p2[0], p2[1], color, Pt(2.0))
    line(slide, p2[0], p2[1], p3[0], p3[1], color, Pt(2.0))


def icon_x(slide, x, y, size, color):
    line(slide, x + size * 0.12, y + size * 0.12, x + size * 0.88, y + size * 0.88, color, Pt(2.0))
    line(slide, x + size * 0.88, y + size * 0.12, x + size * 0.12, y + size * 0.88, color, Pt(2.0))


def icon_alert(slide, x, y, size, color):
    pts = [(x + size / 2, y + size * 0.08), (x + size * 0.95, y + size * 0.90),
           (x + size * 0.05, y + size * 0.90)]
    _freeform_poly(slide, pts, fill=color)
    bar_w = size * 0.10
    bar_h = size * 0.32
    bx = x + size / 2 - bar_w / 2
    by = y + size * 0.35
    rect(slide, bx, by, bar_w, bar_h, fill=WHITE)
    dot_size = size * 0.10
    circle(slide, x + size / 2 - dot_size / 2, y + size * 0.74, dot_size, fill=WHITE)


def icon_info(slide, x, y, size, color):
    circle(slide, x, y, size, fill=color)
    dot = size * 0.10
    circle(slide, x + size / 2 - dot / 2, y + size * 0.22, dot, fill=WHITE)
    bar_w = size * 0.10
    bar_h = size * 0.36
    rect(slide, x + size / 2 - bar_w / 2, y + size * 0.42, bar_w, bar_h, fill=WHITE)


def icon_target(slide, x, y, size, color):
    outer = size
    mid = size * 0.66
    inner = size * 0.30
    circle(slide, x, y, outer, fill=color)
    circle(slide, x + (outer - mid) / 2, y + (outer - mid) / 2, mid, fill=WHITE)
    circle(slide, x + (outer - inner) / 2, y + (outer - inner) / 2, inner, fill=color)


def icon_gear(slide, x, y, size, color):
    cx = x + size / 2
    cy = y + size / 2
    circle(slide, x, y, size, fill=color)
    hub = size * 0.32
    circle(slide, cx - hub / 2, cy - hub / 2, hub, fill=WHITE)
    notch_w = size * 0.14
    notch_h = size * 0.16
    positions = [
        (cx - notch_w / 2, y - notch_h * 0.3),
        (cx - notch_w / 2, y + size - notch_h * 0.7),
        (x - notch_w * 0.3, cy - notch_h / 2),
        (x + size - notch_w * 0.7, cy - notch_h / 2),
    ]
    for px, py in positions:
        rect(slide, px, py, notch_w, notch_h, fill=color)


def icon_users(slide, x, y, size, color):
    head_d = size * 0.32
    body_w = size * 0.50
    body_h = size * 0.32
    circle(slide, x + size * 0.30, y + size * 0.08, head_d, fill=color)
    rect(slide, x + size * 0.18, y + size * 0.46, body_w, body_h, fill=color)
    circle(slide, x + size * 0.55, y + size * 0.18, head_d, fill=color)
    rect(slide, x + size * 0.43, y + size * 0.56, body_w, body_h, fill=color)


def icon_calendar(slide, x, y, size, color):
    body_y = y + size * 0.18
    body_h = size * 0.80
    rect(slide, x + size * 0.05, body_y, size * 0.90, body_h, fill=color)
    rect(slide, x + size * 0.08, body_y + size * 0.18, size * 0.84, body_h - size * 0.22, fill=WHITE)
    rect(slide, x + size * 0.22, y + size * 0.04, size * 0.08, size * 0.22, fill=color)
    rect(slide, x + size * 0.70, y + size * 0.04, size * 0.08, size * 0.22, fill=color)
    hairline(slide, x + size * 0.15, body_y + body_h * 0.55, size * 0.70, Emu(int(Pt(0.75))), color)


def icon_clock(slide, x, y, size, color):
    circle(slide, x, y, size, fill=color)
    inner = size * 0.86
    off = (size - inner) / 2
    circle(slide, x + off, y + off, inner, fill=WHITE)
    cx = x + size / 2
    cy = y + size / 2
    line(slide, cx, cy, cx, cy - size * 0.28, color, Pt(1.5))
    line(slide, cx, cy, cx + size * 0.32, cy, color, Pt(1.0))
    dot = size * 0.08
    circle(slide, cx - dot / 2, cy - dot / 2, dot, fill=color)


def icon_lightbulb(slide, x, y, size, color):
    circle(slide, x + size * 0.18, y + size * 0.06, size * 0.64, fill=color)
    rect(slide, x + size * 0.34, y + size * 0.70, size * 0.32, size * 0.12, fill=color)
    rect(slide, x + size * 0.38, y + size * 0.82, size * 0.24, size * 0.10, fill=color)


def icon_lock(slide, x, y, size, color):
    rect(slide, x + size * 0.15, y + size * 0.42, size * 0.70, size * 0.50, fill=color)
    rect(slide, x + size * 0.25, y + size * 0.18, size * 0.08, size * 0.30, fill=color)
    rect(slide, x + size * 0.67, y + size * 0.18, size * 0.08, size * 0.30, fill=color)
    rect(slide, x + size * 0.25, y + size * 0.15, size * 0.50, size * 0.08, fill=color)
    kh = size * 0.10
    circle(slide, x + size / 2 - kh / 2, y + size * 0.55, kh, fill=WHITE)


def icon_globe(slide, x, y, size, color):
    circle(slide, x, y, size, fill=color)
    cx = x + size / 2
    cy = y + size / 2
    line(slide, x + size * 0.05, cy, x + size * 0.95, cy, WHITE, Pt(0.75))
    line(slide, cx, y + size * 0.05, cx, y + size * 0.95, WHITE, Pt(0.75))


def icon_search(slide, x, y, size, color):
    ring_d = size * 0.70
    rx = x
    ry = y
    circle(slide, rx, ry, ring_d, fill=color)
    hole = ring_d * 0.66
    off = (ring_d - hole) / 2
    circle(slide, rx + off, ry + off, hole, fill=WHITE)
    line(slide, rx + ring_d * 0.78, ry + ring_d * 0.78,
         x + size * 0.98, y + size * 0.98, color, Pt(2.0))


def icon_money(slide, x, y, size, color):
    circle(slide, x, y, size, fill=color)
    bar_w = size * 0.08
    rect(slide, x + size / 2 - bar_w / 2, y + size * 0.18, bar_w, size * 0.64, fill=WHITE)
    rect(slide, x + size * 0.30, y + size * 0.30, size * 0.40, size * 0.08, fill=WHITE)
    rect(slide, x + size * 0.30, y + size * 0.46, size * 0.40, size * 0.08, fill=WHITE)
    rect(slide, x + size * 0.30, y + size * 0.62, size * 0.40, size * 0.08, fill=WHITE)


def icon_trending_up(slide, x, y, size, color):
    pts = [
        (x + size * 0.05, y + size * 0.85),
        (x + size * 0.35, y + size * 0.55),
        (x + size * 0.55, y + size * 0.70),
        (x + size * 0.90, y + size * 0.20),
    ]
    for a, b in zip(pts, pts[1:]):
        line(slide, a[0], a[1], b[0], b[1], color, Pt(1.75))
    tip = pts[-1]
    head = [(tip[0], tip[1]), (tip[0] - size * 0.20, tip[1] + size * 0.02),
            (tip[0] - size * 0.02, tip[1] + size * 0.20)]
    _freeform_poly(slide, head, fill=color)


def icon_trending_down(slide, x, y, size, color):
    pts = [
        (x + size * 0.05, y + size * 0.20),
        (x + size * 0.35, y + size * 0.50),
        (x + size * 0.55, y + size * 0.40),
        (x + size * 0.90, y + size * 0.85),
    ]
    for a, b in zip(pts, pts[1:]):
        line(slide, a[0], a[1], b[0], b[1], color, Pt(1.75))
    tip = pts[-1]
    head = [(tip[0], tip[1]), (tip[0] - size * 0.20, tip[1] - size * 0.02),
            (tip[0] - size * 0.02, tip[1] - size * 0.20)]
    _freeform_poly(slide, head, fill=color)


def icon_shield(slide, x, y, size, color):
    pts = [
        (x + size * 0.50, y + size * 0.05),
        (x + size * 0.92, y + size * 0.20),
        (x + size * 0.92, y + size * 0.55),
        (x + size * 0.50, y + size * 0.95),
        (x + size * 0.08, y + size * 0.55),
        (x + size * 0.08, y + size * 0.20),
    ]
    _freeform_poly(slide, pts, fill=color)


def icon_flag(slide, x, y, size, color):
    rect(slide, x + size * 0.18, y + size * 0.08, size * 0.06, size * 0.84, fill=color)
    pts = [(x + size * 0.24, y + size * 0.10), (x + size * 0.92, y + size * 0.28),
           (x + size * 0.24, y + size * 0.46)]
    _freeform_poly(slide, pts, fill=color)


def icon_doc(slide, x, y, size, color):
    pts = [(x + size * 0.18, y + size * 0.08), (x + size * 0.66, y + size * 0.08),
           (x + size * 0.86, y + size * 0.30), (x + size * 0.86, y + size * 0.92),
           (x + size * 0.18, y + size * 0.92)]
    _freeform_poly(slide, pts, fill=color)
    fold = [(x + size * 0.66, y + size * 0.08), (x + size * 0.66, y + size * 0.30),
            (x + size * 0.86, y + size * 0.30)]
    _freeform_poly(slide, fold, fill=WHITE)
    for k in range(3):
        hairline(slide, x + size * 0.28, y + size * (0.46 + k * 0.12),
                 size * 0.50, Emu(int(Pt(0.75))), WHITE)


ICON_REGISTRY = {
    "chart_bar": icon_chart_bar,
    "chart_line": icon_chart_line,
    "chart_pie": icon_chart_pie,
    "arrow_up": icon_arrow_up,
    "arrow_down": icon_arrow_down,
    "arrow_right": icon_arrow_right,
    "check": icon_check,
    "x": icon_x,
    "alert": icon_alert,
    "info": icon_info,
    "target": icon_target,
    "gear": icon_gear,
    "users": icon_users,
    "calendar": icon_calendar,
    "clock": icon_clock,
    "lightbulb": icon_lightbulb,
    "lock": icon_lock,
    "globe": icon_globe,
    "search": icon_search,
    "money": icon_money,
    "trending_up": icon_trending_up,
    "trending_down": icon_trending_down,
    "shield": icon_shield,
    "flag": icon_flag,
    "doc": icon_doc,
}


def draw_icon(slide, name, x, y, size, color):
    fn = ICON_REGISTRY.get(name, icon_target)
    fn(slide, x, y, size, color)


