"""
brand.py — Client brand definitions, slide constants, and color palette.

PRD §4.3: BRAND{} dict keyed by client identifier.
Legacy aliases (C_RYB_Q4, etc.) remain for backward compatibility.
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

# ── BRAND{} dict (PRD §4.3) ─────────────────────────────────────────────────

BRAND = {
    "jnj": {
        "primary_current":  RGBColor(0xF7, 0x58, 0x24),  # deep orange
        "primary_prior":    RGBColor(0xFF, 0xC1, 0x99),  # pale orange
        "competitor":       RGBColor(0x70, 0x30, 0xA0),  # purple (Tagrisso)
        "competitor_prior": RGBColor(0xAD, 0x88, 0xC8),  # light purple
        "accent":           RGBColor(0xFF, 0x00, 0x00),  # J&J red
        "positive":         RGBColor(0x00, 0xB0, 0x50),  # green
        "negative":         RGBColor(0xFF, 0x00, 0x00),  # red
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
        "positive":         RGBColor(0x00, 0xB0, 0x50),
        "negative":         RGBColor(0xFF, 0x00, 0x00),
        "neutral_grey":     RGBColor(0x50, 0x50, 0x50),
        "font_display":     "Calibri",
        "font_body":        "Calibri",
        "template_path":    None,
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,
        "slide_margin_bottom": 0.20,
    },
}

# ── Legacy color aliases (backward compatibility) ────────────────────────────

C_RYB_Q4   = RGBColor(0xF7, 0x58, 0x24)   # deep orange  — Q4 bars, primary accent
C_RYB_Q3   = RGBColor(0xFF, 0xC1, 0x99)   # pale orange  — Q3 bars
C_TAG      = RGBColor(0x70, 0x30, 0xA0)   # purple       — AZ / Tagrisso
C_RED      = RGBColor(0xFF, 0x00, 0x00)   # J&J red      — title bar, headline
C_GREEN    = RGBColor(0x00, 0xB0, 0x50)   # positive delta
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_GREY     = RGBColor(0x50, 0x50, 0x50)   # body text
C_FTGREY   = RGBColor(0x7F, 0x7F, 0x7F)   # footer / faint text
C_LBGREY   = RGBColor(0xF4, 0xF4, 0xF4)   # alternating table row bg
C_HDRGREY  = RGBColor(0x40, 0x40, 0x40)   # delta table header bg
C_LTGREY   = RGBColor(0xBF, 0xBF, 0xBF)   # gridlines / borders

# ── COM colour equivalents (BGR order) ───────────────────────────────────────

COM_RED    = 0x0000FF
COM_ORANGE = 0x2458F7
COM_GREEN  = 0x50B000
COM_GREY   = 0x505050

# ── Fonts (legacy aliases) ───────────────────────────────────────────────────

FONT_DISPLAY = "Johnson Display"
FONT_TEXT    = "Johnson Text"
