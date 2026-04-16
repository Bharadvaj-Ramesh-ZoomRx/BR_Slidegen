"""
Brand definitions per pharmaceutical client.

Generated from real-deck analysis of 32 PET decks across 17 clients.
See experiments/deck_analysis/outputs/deep_brand_proposals.md for
the observation data that drove these constants.

Each BRAND entry provides:
  - primary:      main series color (used for 'current wave' bars)
  - prior:        secondary/tint color (used for 'prior wave' bars)
  - secondary:    additional accent for comparisons
  - positive:     delta positive (universal green across decks)
  - negative:     delta negative (universal red across decks)
  - heading_color: headline text color
  - font_heading: headline font
  - font_body:    body text font

Growth model: new clients are added to BRAND{} as their first engagement
uses the system. Existing entries are updated when brand guides change.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor


# Universal delta colors — observed across all 32 decks
POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)  # 494 occurrences
NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)    # 913 occurrences
NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)  # 61 occurrences (alt)

# Standard greys (appear across all decks)
GREY_DARK = RGBColor(0x40, 0x40, 0x40)
GREY_MID = RGBColor(0x59, 0x59, 0x59)
GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)
GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)  # Standard table alt-row fill


BRAND = {
    # --- AZN (7 decks) ---
    "AZN": {
        "primary":       RGBColor(0x00, 0xB0, 0x50),
        "secondary":     RGBColor(0xFF, 0x88, 0x13),
        "prior":         RGBColor(0x94, 0xD4, 0x48),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x59, 0x59, 0x59),
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "_observed_palette": ["#00B050", "#FF8813", "#94D448", "#E15759", "#250E62", "#33A4FF", "#59A14F", "#9467BD"],  # for reference
    },

    # --- JJ (4 decks) ---
    "JJ": {
        "primary":       RGBColor(0x00, 0x63, 0xC3),
        "secondary":     RGBColor(0x00, 0x63, 0xC3),
        "prior":         RGBColor(0x6F, 0xC6, 0xC1),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x00, 0x1E, 0x60),
        "font_heading":  "Johnson Display",
        "font_body":     "Johnson Text",
        "_observed_palette": ["#7FB1E1", "#0063C3", "#6FC6C1", "#37CFCA", "#BD05ED", "#228D8A", "#A0D0FF", "#FF6A5A"],  # for reference
    },

    # --- Pfizer (2 decks) ---
    "PFIZER": {
        "primary":       RGBColor(0x00, 0x00, 0xC9),
        "secondary":     RGBColor(0x43, 0x96, 0x4A),
        "prior":         RGBColor(0xE3, 0x56, 0xDC),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x00, 0x63, 0xC3),
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "_observed_palette": ["#0000C9", "#43964A", "#E356DC", "#F49C34", "#BA2A4C", "#C00000", "#6D5A7A", "#4060AF"],  # for reference
    },

    # --- Regeneron (2 decks) ---
    "REGENERON": {
        "primary":       RGBColor(0x21, 0x94, 0x91),
        "secondary":     RGBColor(0xF7, 0x96, 0x46),
        "prior":         RGBColor(0xE4, 0x6C, 0x0A),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x00, 0x74, 0x5A),
        "font_heading":  "Trade Gothic LT Std",
        "font_body":     "Arial",
        "_observed_palette": ["#219491", "#F79646", "#E46C0A", "#DC0077", "#8EB4E3", "#00745A", "#7030A0", "#004F6F"],  # for reference
    },

    # --- Amgen (2 decks) ---
    "AMGEN": {
        "primary":       RGBColor(0x1F, 0x49, 0x7D),
        "secondary":     RGBColor(0x81, 0x3F, 0x97),
        "prior":         RGBColor(0x00, 0x7D, 0x6D),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x00, 0x3C, 0x71),
        "font_heading":  "Century Gothic",
        "font_body":     "Century Gothic",
        "_observed_palette": ["#1F497D", "#813F97", "#007D6D", "#6F8DB4", "#003C71", "#DA205F", "#00A3DF", "#4C35DE"],  # for reference
    },

    # --- Unknown (2 decks) ---
    "UNKNOWN": {
        "primary":       RGBColor(0xE7, 0x00, 0x4C),
        "secondary":     RGBColor(0xC7, 0xA0, 0x13),
        "prior":         RGBColor(0x6B, 0xC9, 0xEF),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading":  "Century Gothic",
        "font_body":     "Arial",
        "_observed_palette": ["#E7004C", "#C7A013", "#6BC9EF", "#006838", "#FF6969", "#00607C", "#4E79A7", "#00A44A"],  # for reference
    },

    # --- GSK (2 decks) ---
    "GSK": {
        "primary":       RGBColor(0xF3, 0x66, 0x33),
        "secondary":     RGBColor(0x66, 0x8E, 0xDD),
        "prior":         RGBColor(0x66, 0x58, 0xA6),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0xF3, 0x66, 0x33),
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "_observed_palette": ["#3AB51D", "#668EDD", "#6658A6", "#5C103B", "#7030A0", "#D51900", "#17B3AF", "#1B3B7A"],  # for reference
    },

    # --- LEO (2 decks) ---
    "LEO": {
        "primary":       RGBColor(0xC0, 0x14, 0xA3),
        "secondary":     RGBColor(0xD5, 0x68, 0x5F),
        "prior":         RGBColor(0x6F, 0x43, 0x9A),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0xC0, 0x14, 0xA3),
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "_observed_palette": ["#C014A3", "#D5685F", "#6F439A", "#FED006", "#EAAFE0", "#009B77", "#D561C1", "#90117A"],  # for reference
    },

    # --- Bone-HCP (1 deck) ---
    "BONE_HCP": {
        "primary":       RGBColor(0xFF, 0x99, 0x33),
        "secondary":     RGBColor(0x00, 0x92, 0x01),
        "prior":         RGBColor(0x00, 0xB0, 0xF0),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0xFF, 0x99, 0x33),
        "font_heading":  "Arial",
        "font_body":     "Century Gothic",
        "_observed_palette": ["#FF9933", "#009201", "#00B0F0", "#58193D", "#A6A6A6", "#0063C3", "#4BACC6", "#E46C0A"],  # for reference
    },

    # --- Otsuka (1 deck) ---
    "OTSUKA": {
        "primary":       RGBColor(0xFF, 0xC0, 0x00),
        "secondary":     RGBColor(0xC0, 0x00, 0x00),
        "prior":         RGBColor(0x2D, 0x5C, 0xA2),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading":  "Calibri",
        "font_body":     "Arial",
        "_observed_palette": ["#FFC000", "#C00000", "#2D5CA2", "#7F47AA", "#7D44A9", "#3B5998", "#7E97CD", "#B3C9EA"],  # for reference
    },

    # --- Alexion (1 deck) ---
    "ALEXION": {
        "primary":       RGBColor(0x0E, 0x87, 0x79),
        "secondary":     RGBColor(0x70, 0x30, 0xA0),
        "prior":         RGBColor(0xE3, 0x51, 0x05),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x0E, 0x87, 0x79),
        "font_heading":  "Arial Black",
        "font_body":     "Arial",
        "_observed_palette": ["#0E8779", "#7030A0", "#E35105", "#799A01", "#FFA300", "#FF8181", "#CCCCCC", "#84CA68"],  # for reference
    },

    # --- BL (1 deck) ---
    "BL": {
        "primary":       RGBColor(0x00, 0xA9, 0xEB),
        "secondary":     RGBColor(0x40, 0x02, 0x86),
        "prior":         RGBColor(0xA8, 0x10, 0x9D),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x40, 0x02, 0x86),
        "font_heading":  "Avenir Next LT Pro",
        "font_body":     "Century Gothic",
        "_observed_palette": ["#00A9EB", "#400286", "#A8109D", "#00B050", "#A6A6A6", "#FF1119", "#413A5F", "#000F9F"],  # for reference
    },

    # --- Novartis (1 deck) ---
    "NOVARTIS": {
        "primary":       RGBColor(0x01, 0x8E, 0x86),
        "secondary":     RGBColor(0x00, 0x70, 0xFE),
        "prior":         RGBColor(0x00, 0x20, 0x68),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x01, 0x8E, 0x86),
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "_observed_palette": ["#018E86", "#0070FE", "#002068", "#8F2DDE", "#852065", "#B56FA3", "#61A8FF", "#50E2D0"],  # for reference
    },

    # --- CCA (1 deck) ---
    "CCA": {
        "primary":       RGBColor(0xF2, 0x8E, 0x2B),
        "secondary":     RGBColor(0xFF, 0xFF, 0xFF),
        "prior":         RGBColor(0xA6, 0xA6, 0xA6),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0xF2, 0x8E, 0x2B),
        "font_heading":  "Century Gothic",
        "font_body":     "Century Gothic",
        "_observed_palette": ["#F28E2B", "#FFFFFF", "#A6A6A6", "#8FAADC", "#BDD7EE", "#AFED5D", "#FFABD5", "#BFBFBF"],  # for reference
    },

    # --- DSI (1 deck) ---
    "DSI": {
        "primary":       RGBColor(0x1E, 0x22, 0xAA),
        "secondary":     RGBColor(0x64, 0x34, 0x66),
        "prior":         RGBColor(0xFB, 0xCC, 0xB1),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x1E, 0x22, 0xAA),
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "_observed_palette": ["#1E22AA", "#643466", "#FBCCB1", "#A6A6A6", "#B253DE", "#353236", "#CC0066", "#75E6E9"],  # for reference
    },

    # --- Ipsen (1 deck) ---
    "IPSEN": {
        "primary":       RGBColor(0x54, 0xAC, 0x65),
        "secondary":     RGBColor(0xC8, 0x48, 0x74),
        "prior":         RGBColor(0x8E, 0xC8, 0x99),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0x54, 0xAC, 0x65),
        "font_heading":  "Rethink Sans",
        "font_body":     "Calibri",
        "_observed_palette": ["#54AC65", "#C84874", "#8EC899", "#595959", "#BBDEC2", "#D3D3D3", "#D2E9D6", "#A6A6A6"],  # for reference
    },

    # --- Apellis (1 deck) ---
    "APELLIS": {
        "primary":       RGBColor(0xFC, 0x3B, 0x6E),
        "secondary":     RGBColor(0xE7, 0xE6, 0xE6),
        "prior":         RGBColor(0x30, 0xCF, 0xD0),
        "positive":      POSITIVE_GREEN,
        "negative":      NEGATIVE_RED,
        "heading_color": RGBColor(0xFC, 0x3B, 0x6E),
        "font_heading":  "Calibri Light",
        "font_body":     "Calibri",
        "_observed_palette": ["#FC3B6E", "#E7E6E6", "#30CFD0", "#81E3E3", "#5A4986", "#FD7FA0", "#FEB0C4", "#FFC000"],  # for reference
    },

}


def get_brand(client_key_or_alias: str) -> dict:
    """Look up brand config by client key. Raises KeyError with available keys."""
    key = client_key_or_alias.upper().replace(' ', '_').replace('-', '_')
    if key not in BRAND:
        raise KeyError(
            f"Unknown client {client_key_or_alias!r}. Available: {sorted(BRAND.keys())}"
        )
    return BRAND[key]


def get_color(client: str, role: str) -> RGBColor:
    """Shortcut: get_color("JJ", "primary") -> RGBColor."""
    return get_brand(client)[role]
