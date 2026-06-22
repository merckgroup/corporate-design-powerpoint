from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

# ===========================================================================
# Fonts and palette
# ===========================================================================

FONT_HEAD = "Verdana"  # BinaryFiles carry KR_Merck theme (major=Noto Sans CJK KR Bold); always override with Verdana
FONT_BODY = "Verdana"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ---------------------------------------------------------------------------
# Merck Corporate Design official palette (exact hex from brand guidelines).
# Full set:  violet, blue, green, red, pink, lightblue, lightgreen, yellow,
#            palepink, paleblue, palegreen, paleyellow, grey
# ---------------------------------------------------------------------------
MERCK_PURPLE  = (0x50, 0x32, 0x91)   # #503291 — violet     PRIMARY
MERCK_BLUE    = (0x0F, 0x69, 0xAF)   # #0F69AF — blue       SECONDARY
MERCK_GOLD    = (0xEB, 0x3C, 0x96)   # #EB3C96 — pink       TERTIARY (was MERCK_GOLD)
PURPLE_DEEP   = (0x3A, 0x24, 0x68)   # #3A2468 — footer / dark backgrounds
PURPLE_MUTED  = (0x7D, 0x74, 0xA0)   # #7D74A0 — muted / separator
MERCK_YELLOW  = (0xFF, 0xC8, 0x32)   # #FFC832 — yellow
MERCK_AQUA    = (0x96, 0xD7, 0xD2)   # #96D7D2 — paleblue   (sensitive-blue)
LIGHT_GRAY    = (0xE0, 0xE0, 0xE0)   # #E0E0E0 — rules / borders
PANEL_LIGHT   = (0xF4, 0xF2, 0xF8)   # #F4F2F8 — card panel background
WHITE         = (0xFF, 0xFF, 0xFF)
INK_DARK      = (0x1A, 0x16, 0x26)   # #1A1626 — primary body text
INK_GRAY      = (0x55, 0x5D, 0x6E)   # #555D6E — secondary text / sources
BAD_RED       = (0xE6, 0x1E, 0x50)   # #E61E50 — red
GOOD_GREEN    = (0x14, 0x9B, 0x5F)   # #149B5F — green

# Full official Merck Corporate Design colour names (for theme palettes).
_MC_VIOLET     = MERCK_PURPLE          # #503291
_MC_BLUE       = MERCK_BLUE            # #0F69AF
_MC_GREEN      = GOOD_GREEN            # #149B5F
_MC_RED        = BAD_RED               # #E61E50
_MC_PINK       = MERCK_GOLD            # #EB3C96
_MC_LIGHTBLUE  = (0x2D, 0xBE, 0xCD)   # #2DBECD
_MC_LIGHTGREEN = (0xA5, 0xCD, 0x50)   # #A5CD50  (= OP_LIME)
_MC_YELLOW     = MERCK_YELLOW          # #FFC832
_MC_PALEPINK   = (0xE1, 0xC3, 0xCD)   # #E1C3CD
_MC_PALEBLUE   = MERCK_AQUA            # #96D7D2
_MC_PALEGREEN  = (0xB4, 0xDC, 0x96)   # #B4DC96
_MC_PALEYELLOW = (0xFF, 0xDC, 0xB9)   # #FFDCB9
_MC_GREY       = (0x99, 0x99, 0x99)   # #999999

# Liquid Carbon LS dashboard data-series palette (from LS dashboard legend).
# Use these for chart series, RAG indicators, and data comparisons.
ACT_PURPLE   = MERCK_PURPLE          # #503291 — Actual (same as primary)
LY_CYAN      = (0x2D, 0xBE, 0xCD)   # #2DBECD — Last Year
OP_LIME      = (0xA5, 0xCD, 0x50)   # #A5CD50 — Operating Plan
FC_PINK      = (0xEB, 0x3C, 0x96)   # #EB3C96 — Forecast
DEV_POS_BLUE = (0x0F, 0x69, 0xAF)   # #0F69AF — Deviation positive
DEV_NEG_RED  = BAD_RED               # #E61E50 — Deviation negative (same as BAD_RED)

# Liquid Carbon 12-color chart palette (resolved from tailwind-colors.json).
# chart-1 through chart-12 match the LC setup.json chart token order.
# chart-8 (sensitive-blue) and chart-10 (sensitive-yellow) are custom tokens
# approximated from the brand palette.
CHART_1  = (0xDE, 0x7A, 0x21)   # #DE7A21 — orange-600
CHART_2  = (0x12, 0x87, 0x9D)   # #12879D — teal-600
CHART_3  = (0x18, 0x42, 0x51)   # #184251 — cyan-900
CHART_4  = (0xFE, 0xC7, 0x31)   # #FEC731 — amber-400
CHART_5  = (0xF7, 0xA2, 0x16)   # #F7A216 — amber-500
CHART_6  = (0x0E, 0x69, 0xAF)   # #0E69AF — blue-700
CHART_7  = (0xE6, 0x1E, 0x50)   # #E61E50 — red-600 (same as BAD_RED)
CHART_8  = (0x96, 0xD7, 0xD2)   # #96D7D2 — sensitive-blue (exact from vibrant-m)
CHART_9  = (0x14, 0x9B, 0x5F)   # #149B5F — green-500
CHART_10 = (0xFF, 0xDC, 0xB9)   # #FFDCB9 — sensitive-yellow (exact from vibrant-m)
CHART_11 = (0x9D, 0x80, 0xE6)   # #9D80E6 — violet-400
CHART_12 = (0x8C, 0x22, 0x35)   # #8C2235 — red-800

# Ordered list for cycling through chart series automatically.
CHART_PALETTE = [
    CHART_1, CHART_2, CHART_3, CHART_4,
    CHART_5, CHART_6, CHART_7, CHART_8,
    CHART_9, CHART_10, CHART_11, CHART_12,
]

PHASE_1_COLOR = MERCK_PURPLE
PHASE_2_COLOR = MERCK_AQUA
PHASE_3_COLOR = MERCK_GOLD


PALETTES = {
    "merck_executive": {
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    MERCK_PURPLE,
        "accent_2":  PURPLE_DEEP,
        "accent_3":  LY_CYAN,
        "highlight": MERCK_GOLD,
        "hot":       MERCK_YELLOW,
        "rule":      LIGHT_GRAY,
        "panel":     PANEL_LIGHT,
        "muted":     PURPLE_MUTED,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      OP_LIME,
    },
    "merck_corporate": {
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    MERCK_PURPLE,
        "accent_2":  PURPLE_DEEP,
        "accent_3":  LY_CYAN,
        "highlight": MERCK_GOLD,
        "hot":       MERCK_YELLOW,
        "rule":      LIGHT_GRAY,
        "panel":     PANEL_LIGHT,
        "muted":     PURPLE_MUTED,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      OP_LIME,
    },
    "merck_storytelling": {
        "bg":        MERCK_PURPLE,
        "ink":       WHITE,
        "ink_2":     PANEL_LIGHT,
        "ink_3":     PURPLE_MUTED,
        "accent":    MERCK_GOLD,
        "accent_2":  MERCK_YELLOW,
        "accent_3":  LY_CYAN,
        "highlight": MERCK_GOLD,
        "hot":       MERCK_YELLOW,
        "rule":      PURPLE_MUTED,
        "panel":     PANEL_LIGHT,
        "muted":     PURPLE_MUTED,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      OP_LIME,
    },
    # ------------------------------------------------------------------
    # Official Merck Corporate Design themes (6 variants).
    # Theme names match the empower library folder names exactly.
    # Light themes use WHITE content bg; dark themes use MERCK_PURPLE.
    # ------------------------------------------------------------------
    "functional": {       # lightgreen bg cover, teal accent, organic cells
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    _MC_LIGHTBLUE,
        "accent_2":  MERCK_PURPLE,
        "accent_3":  _MC_LIGHTGREEN,
        "highlight": _MC_LIGHTBLUE,
        "hot":       _MC_LIGHTGREEN,
        "rule":      LIGHT_GRAY,
        "panel":     (0xF0, 0xF9, 0xFA),    # very pale teal
        "muted":     _MC_PALEBLUE,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      _MC_LIGHTGREEN,
    },
    "organic": {          # paleyellow bg cover, red accent, warm life-science
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    _MC_RED,
        "accent_2":  MERCK_PURPLE,
        "accent_3":  _MC_PALEYELLOW,
        "highlight": _MC_RED,
        "hot":       _MC_PALEYELLOW,
        "rule":      LIGHT_GRAY,
        "panel":     (0xFD, 0xF5, 0xED),    # very pale warm
        "muted":     _MC_PALEPINK,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       _MC_RED,
        "lime":      _MC_LIGHTGREEN,
    },
    "plastic": {          # lightgreen bg cover, pink accent — current default
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    MERCK_PURPLE,
        "accent_2":  PURPLE_DEEP,
        "accent_3":  LY_CYAN,
        "highlight": _MC_PINK,
        "hot":       _MC_LIGHTGREEN,
        "rule":      LIGHT_GRAY,
        "panel":     PANEL_LIGHT,
        "muted":     PURPLE_MUTED,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      OP_LIME,
    },
    "synthetic": {        # dark violet bg, yellow accent, industrial
        "bg":        MERCK_PURPLE,
        "ink":       WHITE,
        "ink_2":     PANEL_LIGHT,
        "ink_3":     PURPLE_MUTED,
        "accent":    _MC_YELLOW,
        "accent_2":  _MC_LIGHTBLUE,
        "accent_3":  _MC_LIGHTGREEN,
        "highlight": _MC_YELLOW,
        "hot":       _MC_YELLOW,      # yellow: section number + takeaway band
        "rule":      PURPLE_MUTED,
        "panel":     (0x3F, 0x28, 0x70),    # dark panel (lighter than bg)
        "muted":     PURPLE_MUTED,
        "good":      GOOD_GREEN,
        "warn":      _MC_YELLOW,
        "bad":       BAD_RED,
        "lime":      _MC_LIGHTGREEN,
    },
    "technical": {        # paleyellow bg, teal/lightblue, angular shapes
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    _MC_LIGHTBLUE,
        "accent_2":  MERCK_PURPLE,
        "accent_3":  _MC_PALEYELLOW,
        "highlight": _MC_LIGHTBLUE,
        "hot":       _MC_PALEYELLOW,
        "rule":      LIGHT_GRAY,
        "panel":     (0xF0, 0xF9, 0xFA),    # very pale teal (same as functional)
        "muted":     _MC_PALEBLUE,
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      _MC_LIGHTGREEN,
    },
    "electronics": {      # dark violet bg, yellow accent, photo placeholder on cover
        "bg":        MERCK_PURPLE,
        "ink":       WHITE,
        "ink_2":     PANEL_LIGHT,
        "ink_3":     PURPLE_MUTED,
        "accent":    _MC_YELLOW,
        "accent_2":  _MC_LIGHTBLUE,
        "accent_3":  _MC_LIGHTGREEN,
        "highlight": _MC_YELLOW,
        "hot":       _MC_YELLOW,      # yellow: section number + takeaway band
        "rule":      PURPLE_MUTED,
        "panel":     (0x3F, 0x28, 0x70),
        "muted":     PURPLE_MUTED,
        "good":      GOOD_GREEN,
        "warn":      _MC_YELLOW,
        "bad":       BAD_RED,
        "lime":      _MC_LIGHTGREEN,
    },
    "merck_science": {    # white bg, blue accent — data-first scientific reporting
        "bg":        WHITE,
        "ink":       INK_DARK,
        "ink_2":     INK_GRAY,
        "ink_3":     LIGHT_GRAY,
        "accent":    MERCK_BLUE,         # #0F69AF — blue as primary accent
        "accent_2":  MERCK_PURPLE,       # purple demoted to secondary
        "accent_3":  GOOD_GREEN,
        "highlight": MERCK_AQUA,         # #96D7D2 — soft teal for subtle highlights
        "hot":       MERCK_BLUE,         # blue as "hot" (no dramatic theatrics)
        "rule":      LIGHT_GRAY,
        "panel":     (0xEA, 0xF2, 0xFB), # pale blue panel (vs purple-tinted PANEL_LIGHT)
        "muted":     (0x5A, 0x82, 0xA8), # blue-gray for captions and methodology notes
        "good":      GOOD_GREEN,
        "warn":      MERCK_YELLOW,
        "bad":       BAD_RED,
        "lime":      OP_LIME,
    },
}

AUTO_PROMOTE_EXECUTIVE = {
    "Executive Summary",
    "Recommendation",
    "Decision Request",
    "Risk",
    "Tradeoff",
}


# ===========================================================================
# Layout constants
# ===========================================================================

# Top chrome (universal)
GOLD_RULE_X = Inches(0.65)
GOLD_RULE_Y = Inches(0.30)
GOLD_RULE_W = Inches(3.5)
GOLD_RULE_H = Inches(0.04)

CLASS_BADGE_X = Inches(11.5)
CLASS_BADGE_Y = Inches(0.20)
CLASS_BADGE_W = Inches(1.7)

BREADCRUMB_X = Inches(0.65)
BREADCRUMB_Y = Inches(0.80)
BREADCRUMB_W = Inches(9.0)
BREADCRUMB_H = Inches(0.30)

# Section marker zone (numbered content slides). Top breadcrumb removed,
# so positions shifted up by 0.90" to reclaim that vertical space.
SECTION_CIRCLE_X = Inches(0.65)
SECTION_CIRCLE_Y = Inches(0.40)
SECTION_CIRCLE_D = Inches(0.55)

SECTION_TAG_X = Inches(1.35)
SECTION_TAG_Y = Inches(0.52)
SECTION_TAG_W = Inches(8.0)
SECTION_TAG_H = Inches(0.30)

# Action title zone (32pt bold serif takes ~0.55in/line; allow 2 lines + slack)
TITLE_X = Inches(0.65)
TITLE_Y_NUMBERED = Inches(1.15)
TITLE_Y_UNNUMBERED = Inches(0.45)
TITLE_W = Inches(12.0)
TITLE_H = Inches(1.30)

# Subtitle zone (placed beneath the title block)
SUB_X = Inches(0.65)
SUB_W = Inches(12.0)
SUB_H = Inches(0.40)
SUB_GAP = Inches(1.30)   # distance from title_y to subtitle_y

# Content zone (between title block and source/takeaway/footer chrome)
CONTENT_X = Inches(0.65)
CONTENT_Y = Inches(2.55)
CONTENT_Y_SUBTITLE = Inches(2.95)
CONTENT_W = Inches(12.0)
CONTENT_H = Inches(4.00)

# Below-content chrome positions (above the footer band)
SOURCE_Y    = Inches(6.55)
SOURCE_H    = Inches(0.22)
TAKEAWAY_Y  = Inches(6.83)
TAKEAWAY_H  = Inches(0.22)

# Phase progress (above footer band)
PHASE_Y = Inches(6.25)
PHASE_H = Inches(0.55)

# Footer band. Reduced from 0.40" to 0.30" height (top dropped from 7.10
# to 7.20) so the band no longer overlaps the master slide's Merck "M"
# logo, which spans y=6.73 to y=7.15. Band now sits cleanly below the
# logo at 7.20 → 7.50. Text inside re-centered to 7.24.
FOOTER_Y = Inches(7.20)
FOOTER_H = Inches(0.30)
FOOTER_TEXT_Y = Inches(7.24)


# ===========================================================================
# Color helpers
# ===========================================================================

def _rgb_tuple(t) -> RGBColor:
    if isinstance(t, RGBColor):
        return t
    if isinstance(t, str):
        h = t.lstrip("#")
        if len(h) == 6:
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    return RGBColor(*t)


def _palette_for(style: str) -> dict:
    return PALETTES.get(style, PALETTES["merck_executive"])


_DARK_STYLES = frozenset({"merck_storytelling", "synthetic", "electronics"})


def _is_dark(style: str) -> bool:
    return style in _DARK_STYLES


def rgb(palette_name: str, key: str) -> RGBColor:
    """Return an RGBColor from a named palette and key. Falls back to ink."""
    pal = PALETTES.get(palette_name, PALETTES["merck_executive"])
    tup = pal.get(key) or pal.get("ink") or (0, 0, 0)
    return RGBColor(*tup)


