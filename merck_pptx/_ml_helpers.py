from __future__ import annotations
from typing import Optional
from ._ml_constants import (
    MERCK_PURPLE, MERCK_GOLD, MERCK_YELLOW, MERCK_BLUE, PURPLE_MUTED,
    BAD_RED, GOOD_GREEN, LY_CYAN, _MC_GREY, _MC_PINK,
    AUTO_PROMOTE_EXECUTIVE,
)

# ===========================================================================
# Named colour → RGB for per-item colour overrides (label_rows, funnel, etc.)
# Keys match the colour names emitted by _shape_color_name() in generate.py.
# ===========================================================================

_NAMED_COLOR = {
    "gray":   _MC_GREY,
    "grey":   _MC_GREY,
    "teal":   LY_CYAN,
    "blue":   MERCK_BLUE,
    "green":  GOOD_GREEN,
    "yellow": MERCK_YELLOW,
    "orange": (0xFF, 0x82, 0x00),
    "red":    BAD_RED,
    "pink":   _MC_PINK,    # #EB3C96 — Merck hot-pink (MERCK_GOLD is a historical alias)
    "purple": MERCK_PURPLE,
}


def _named_color(name: Optional[str], fallback):
    """Resolve a color name string to an RGB tuple, falling back to ``fallback``.

    Unknown names silently fall back so a typo in a plan JSON never crashes
    the build — the deck just uses the layout's default color.
    """
    if not name:
        return fallback
    return _NAMED_COLOR.get(str(name).lower().strip(), fallback)

# ===========================================================================
# Style helpers
# ===========================================================================

def _style_or_promote(category, style):
    # Never override an explicit storytelling style — visual consistency
    # across the deck matters more than per-category promotion.
    if style == "merck_storytelling":
        return style
    if category and str(category) in AUTO_PROMOTE_EXECUTIVE:
        return "merck_executive"
    return style


def _tone_color(tone, palette):
    if tone == "positive":
        return GOOD_GREEN
    if tone == "negative":
        return BAD_RED
    v = str(tone).strip().lower()
    # Decision-row semantic tones (plan schema: approve / discuss / note).
    if v in ("approve", "approved", "yes", "accept", "green", "g", "good"):
        return GOOD_GREEN
    if v in ("discuss", "review", "consider", "pending", "open",
             "amber", "yellow", "a"):
        return MERCK_YELLOW
    if v in ("note", "fyi", "info", "neutral", "observation"):
        return PURPLE_MUTED
    if v in ("red", "r"):
        from_rag = _rag_color(tone)
        return from_rag
    return MERCK_PURPLE


def _rag_color(value):
    v = str(value or "").strip().upper()
    if v in ("RED", "R", "HIGH"):
        return BAD_RED
    if v in ("AMBER", "A", "YELLOW", "MEDIUM", "MED"):
        return MERCK_GOLD
    if v in ("GREEN", "G", "LOW", "OK"):
        return GOOD_GREEN
    return PURPLE_MUTED


def _norm_key(s):
    return "".join(ch for ch in str(s).lower() if ch.isalnum())

