from __future__ import annotations
from ._ml_constants import (
    MERCK_PURPLE, MERCK_GOLD, MERCK_YELLOW, PURPLE_MUTED, BAD_RED, GOOD_GREEN,
    AUTO_PROMOTE_EXECUTIVE,
)

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

