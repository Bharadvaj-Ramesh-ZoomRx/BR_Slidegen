"""
brand.py — Client brand definitions, slide constants, and color palette.

PRD §4.3: BRAND{} dict keyed by client identifier.
Legacy aliases (C_RYB_Q4, etc.) remain for backward compatibility.

BRAND entries were expanded from just "jnj" + "default" to 17 client entries
using real-deck analysis of 32 PET decks. See
`experiments/deck_analysis/outputs/ACTIONABLE_FINDINGS.md §3` for the source
data. The original "jnj" and "default" entries are preserved unchanged for
backward compatibility.
"""

from pptx.dml.color import RGBColor
from pptx.util import Emu

# ── Slide dimensions ─────────────────────────────────────────────────────────

SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.500
SLIDE_W_EMU = Emu(12192000)
SLIDE_H_EMU = Emu(6858000)

# ── Unit conversion ──────────────────────────────────────────────────────────

IN = 72            # 1 inch = 72 points (COM)
EMU_PER_IN = 914400

# ── Universal colors observed across all 32 client decks ────────────────────

POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)     # 494 occurrences — standard positive delta
NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)       # 913 occurrences — standard negative delta
NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)  #  61 occurrences — alternate negative

# Standard greys (appear across all decks)
GREY_DARK = RGBColor(0x40, 0x40, 0x40)
GREY_MID = RGBColor(0x59, 0x59, 0x59)
GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)
GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)       # 229 occurrences — default table alt-row

# ── BRAND{} dict (PRD §4.3) ─────────────────────────────────────────────────
#
# Each BRAND entry provides:
#   - primary_current: main series color (for current wave bars)
#   - primary_prior:   tint color (for prior wave bars)
#   - competitor:      secondary accent color (cross-brand comparison)
#   - competitor_prior:tint of secondary
#   - accent:          brand accent (used in headlines, title bars)
#   - positive:        POSITIVE_GREEN (universal)
#   - negative:        NEGATIVE_RED (universal)
#   - neutral_grey:    body text default
#   - font_display:    headline font
#   - font_body:       body text font
#   - template_path:   default slide master
#
# Legacy entries (jnj, default) retained unchanged for backward compat.

BRAND = {
    "jnj": {
        "primary_current":  RGBColor(0xF7, 0x58, 0x24),  # deep orange (legacy Rybrevant)
        "primary_prior":    RGBColor(0xFF, 0xC1, 0x99),  # pale orange
        "competitor":       RGBColor(0x70, 0x30, 0xA0),  # purple (Tagrisso)
        "competitor_prior": RGBColor(0xAD, 0x88, 0xC8),  # light purple
        "accent":           RGBColor(0xFF, 0x00, 0x00),  # J&J red
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     RGBColor(0x50, 0x50, 0x50),
        "font_display":     "Johnson Display",
        "font_body":        "Johnson Text",
        "template_path":    "templates/template.pptx",
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },
    "default": {
        "primary_current":  RGBColor(0x44, 0x72, 0xC4),  # blue
        "primary_prior":    RGBColor(0xA9, 0xC5, 0xE8),  # light blue
        "competitor":       RGBColor(0xED, 0x7D, 0x31),  # orange
        "competitor_prior": RGBColor(0xF5, 0xBE, 0x97),  # light orange
        "accent":           RGBColor(0x44, 0x72, 0xC4),  # blue
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     RGBColor(0x50, 0x50, 0x50),
        "font_display":     "Calibri",
        "font_body":        "Calibri",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # ── Pharma client brands (from deck analysis — 32 PET decks, 17 clients) ──

    # Johnson & Johnson — 4 PET decks (Rybrevant, SFEA, Tepezza, Ojjaara)
    "JJ": {
        "primary_current":  RGBColor(0x00, 0x63, 0xC3),  # JJ official blue
        "primary_prior":    RGBColor(0x7F, 0xB1, 0xE1),  # light blue tint
        "competitor":       RGBColor(0x6F, 0xC6, 0xC1),  # teal
        "competitor_prior": RGBColor(0xCB, 0xF3, 0xF1),  # pale teal
        "accent":           RGBColor(0x00, 0x1E, 0x60),  # deep navy
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Johnson Display",
        "font_body":        "Johnson Text",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # AstraZeneca — 7 PET decks (Calquence, Lokelma, Lynparza, Truqap, Tezspire, Enhertu, Datroway)
    "AZN": {
        "primary_current":  RGBColor(0x00, 0xB0, 0x50),  # AZN green
        "primary_prior":    RGBColor(0x94, 0xD4, 0x48),  # pale green
        "competitor":       RGBColor(0xFF, 0x88, 0x13),  # orange
        "competitor_prior": RGBColor(0xFF, 0xC5, 0x8A),  # pale orange
        "accent":           GREY_MID,
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Pfizer — 2 decks (Abrysvo, Bavencio via EMD)
    "PFIZER": {
        "primary_current":  RGBColor(0x00, 0x00, 0xC9),  # Pfizer navy
        "primary_prior":    RGBColor(0x6D, 0x5A, 0x7A),  # purple-grey
        "competitor":       RGBColor(0x43, 0x96, 0x4A),  # green
        "competitor_prior": RGBColor(0xA0, 0xCC, 0xA5),  # pale green
        "accent":           RGBColor(0x00, 0x63, 0xC3),  # brand blue
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Regeneron — 2 decks (Libtayo, Dupixent)
    "REGENERON": {
        "primary_current":  RGBColor(0x21, 0x94, 0x91),  # teal
        "primary_prior":    RGBColor(0x7A, 0xBF, 0xBD),  # light teal
        "competitor":       RGBColor(0xF7, 0x96, 0x46),  # orange
        "competitor_prior": RGBColor(0xFB, 0xC4, 0x95),  # pale orange
        "accent":           RGBColor(0x00, 0x74, 0x5A),  # deep teal
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Trade Gothic LT Std",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Amgen — 2 decks (Uplizna, Otezla)
    "AMGEN": {
        "primary_current":  RGBColor(0x1F, 0x49, 0x7D),  # Amgen navy
        "primary_prior":    RGBColor(0x6F, 0x8D, 0xB4),  # navy tint
        "competitor":       RGBColor(0x81, 0x3F, 0x97),  # purple
        "competitor_prior": RGBColor(0xB8, 0x8A, 0xC4),  # light purple
        "accent":           RGBColor(0x00, 0x3C, 0x71),  # deep navy
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Century Gothic",
        "font_body":        "Century Gothic",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # GSK — 2 decks (Blenrep, Jemperli-Zejula)
    "GSK": {
        "primary_current":  RGBColor(0xF3, 0x66, 0x33),  # GSK orange (official)
        "primary_prior":    RGBColor(0xF8, 0xA6, 0x85),  # pale orange
        "competitor":       RGBColor(0x66, 0x8E, 0xDD),  # blue
        "competitor_prior": RGBColor(0xA8, 0xBE, 0xE8),  # pale blue
        "accent":           RGBColor(0xF3, 0x66, 0x33),  # GSK orange
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # LEO Pharma — 2 decks (Adbry, Adbry W1'26)
    "LEO": {
        "primary_current":  RGBColor(0xC0, 0x14, 0xA3),  # LEO magenta
        "primary_prior":    RGBColor(0xEA, 0xAF, 0xE0),  # pale pink
        "competitor":       RGBColor(0xD5, 0x68, 0x5F),  # salmon
        "competitor_prior": RGBColor(0xE8, 0xA8, 0xA2),  # pale salmon
        "accent":           RGBColor(0xC0, 0x14, 0xA3),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Daiichi Sankyo — 1 deck (Datroway)
    "DSI": {
        "primary_current":  RGBColor(0x1E, 0x22, 0xAA),  # DSI navy
        "primary_prior":    RGBColor(0x85, 0x88, 0xCA),  # navy tint
        "competitor":       RGBColor(0x64, 0x34, 0x66),  # purple
        "competitor_prior": RGBColor(0xB2, 0x9C, 0xB5),  # purple tint
        "accent":           RGBColor(0x1E, 0x22, 0xAA),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Novartis — 1 deck (Rhapsido)
    "NOVARTIS": {
        "primary_current":  RGBColor(0x01, 0x8E, 0x86),  # Novartis teal
        "primary_prior":    RGBColor(0x50, 0xE2, 0xD0),  # pale teal
        "competitor":       RGBColor(0x00, 0x70, 0xFE),  # blue
        "competitor_prior": RGBColor(0x61, 0xA8, 0xFF),  # pale blue
        "accent":           RGBColor(0x01, 0x8E, 0x86),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Alexion — 1 deck (PNH)
    "ALEXION": {
        "primary_current":  RGBColor(0x0E, 0x87, 0x79),  # teal
        "primary_prior":    RGBColor(0x84, 0xCA, 0x68),  # light green
        "competitor":       RGBColor(0x70, 0x30, 0xA0),  # purple
        "competitor_prior": RGBColor(0xAD, 0x88, 0xC8),  # light purple
        "accent":           RGBColor(0x0E, 0x87, 0x79),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial Black",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Bausch + Lomb — 1 deck (DED)
    "BL": {
        "primary_current":  RGBColor(0x00, 0xA9, 0xEB),  # Bausch blue
        "primary_prior":    RGBColor(0x81, 0xC1, 0xFF),  # pale blue
        "competitor":       RGBColor(0x40, 0x02, 0x86),  # deep purple
        "competitor_prior": RGBColor(0xA8, 0x10, 0x9D),  # magenta
        "accent":           RGBColor(0x40, 0x02, 0x86),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Avenir Next LT Pro",
        "font_body":        "Century Gothic",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Apellis — 1 deck (Empaveli)
    "APELLIS": {
        "primary_current":  RGBColor(0xFC, 0x3B, 0x6E),  # Apellis pink
        "primary_prior":    RGBColor(0xFD, 0x7F, 0xA0),  # light pink
        "competitor":       RGBColor(0x30, 0xCF, 0xD0),  # teal
        "competitor_prior": RGBColor(0x81, 0xE3, 0xE3),  # pale teal
        "accent":           RGBColor(0xFC, 0x3B, 0x6E),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Calibri Light",
        "font_body":        "Calibri",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Otsuka — 1 deck (Abilify LAI)
    "OTSUKA": {
        "primary_current":  RGBColor(0xFF, 0xC0, 0x00),  # Abilify yellow
        "primary_prior":    RGBColor(0xFF, 0xE2, 0x80),  # pale yellow
        "competitor":       RGBColor(0xC0, 0x00, 0x00),  # deep red
        "competitor_prior": RGBColor(0xDB, 0x80, 0x80),  # pale red
        "accent":           RGBColor(0x00, 0x00, 0x00),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Calibri",
        "font_body":        "Arial",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Ipsen — 1 deck (Onivyde)
    "IPSEN": {
        "primary_current":  RGBColor(0x54, 0xAC, 0x65),  # Ipsen green
        "primary_prior":    RGBColor(0x8E, 0xC8, 0x99),  # light green
        "competitor":       RGBColor(0xC8, 0x48, 0x74),  # pink
        "competitor_prior": RGBColor(0xE0, 0xA1, 0xB8),  # pale pink
        "accent":           RGBColor(0x54, 0xAC, 0x65),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Rethink Sans",
        "font_body":        "Calibri",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # CCA (unknown sub-client) — 1 deck
    "CCA": {
        "primary_current":  RGBColor(0xF2, 0x8E, 0x2B),  # orange
        "primary_prior":    RGBColor(0xFF, 0x8F, 0x43),  # alt orange
        "competitor":       RGBColor(0x8F, 0xAA, 0xDC),  # blue
        "competitor_prior": RGBColor(0xBD, 0xD7, 0xEE),  # pale blue
        "accent":           RGBColor(0xF2, 0x8E, 0x2B),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Century Gothic",
        "font_body":        "Century Gothic",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },

    # Bone HCP — 1 deck (client unknown)
    "BONE_HCP": {
        "primary_current":  RGBColor(0xFF, 0x99, 0x33),  # orange
        "primary_prior":    RGBColor(0xFF, 0xC2, 0x85),  # pale orange
        "competitor":       RGBColor(0x00, 0x92, 0x01),  # green
        "competitor_prior": RGBColor(0x84, 0xCA, 0x68),  # light green
        "accent":           RGBColor(0xFF, 0x99, 0x33),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "neutral_grey":     GREY_MID,
        "font_display":     "Arial",
        "font_body":        "Century Gothic",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },
}

# ── Client lookup helpers ────────────────────────────────────────────────────


def get_brand(client_key_or_alias: str) -> dict:
    """Look up brand config by client key (case-insensitive). Raises KeyError with available keys."""
    key = client_key_or_alias.replace(" ", "_").replace("-", "_")
    if key in BRAND:
        return BRAND[key]
    # Try uppercase version for new pharma-client entries
    upper = key.upper()
    if upper in BRAND:
        return BRAND[upper]
    # Try lowercase for legacy entries
    lower = key.lower()
    if lower in BRAND:
        return BRAND[lower]
    raise KeyError(
        f"Unknown client {client_key_or_alias!r}. Available: {sorted(BRAND.keys())}"
    )


def get_color(client: str, role: str) -> RGBColor:
    """Shortcut: `get_color("JJ", "primary_current")` → `RGBColor`."""
    return get_brand(client)[role]


# ── Legacy color aliases (derived from BRAND["jnj"]) ─────────────────────────
# These exist for backward compat; new code should prefer BRAND[client_key].

_jnj = BRAND["jnj"]

C_RYB_Q4   = _jnj["primary_current"]       # deep orange  — Q4 bars, primary accent
C_RYB_Q3   = _jnj["primary_prior"]         # pale orange  — Q3 bars
C_TAG      = _jnj["competitor"]            # purple       — AZ / Tagrisso
C_RED      = _jnj["accent"]                # J&J red      — title bar, headline
C_GREEN    = _jnj["positive"]              # positive delta
C_GREY     = _jnj["neutral_grey"]          # body text

# Structural colors (client-independent)
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_FTGREY   = RGBColor(0x7F, 0x7F, 0x7F)   # footer / faint text
C_LBGREY   = RGBColor(0xF4, 0xF4, 0xF4)   # alternating table row bg
C_HDRGREY  = RGBColor(0x40, 0x40, 0x40)   # delta table header bg
C_LTGREY   = RGBColor(0xBF, 0xBF, 0xBF)   # gridlines / borders

# ── COM colour equivalents (BGR order) ───────────────────────────────────────

COM_RED    = 0x0000FF
COM_ORANGE = 0x2458F7
COM_GREEN  = 0x50B000
COM_GREY   = 0x505050

# ── Fonts (derived from BRAND["jnj"]) ────────────────────────────────────────

FONT_DISPLAY = _jnj["font_display"]
FONT_TEXT    = _jnj["font_body"]
