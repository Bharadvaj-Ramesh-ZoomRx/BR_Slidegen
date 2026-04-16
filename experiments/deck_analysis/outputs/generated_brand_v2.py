"""
Pharmaceutical brand definitions (one entry per brand, not per client).

Generated from real-deck analysis of 32 PET decks. Each deck is a single-brand
PET, so the series colors extracted per deck represent that brand's palette.

A CLIENT has many BRANDS with different colors. The client-level defaults
(fonts, heading colors, templates) live in CLIENT{} and are referenced by each
brand via its "client" field.

Lookup: `get_brand("RYBREVANT")` -> merged dict with brand colors + client
fonts/templates already resolved.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor


# Universal delta colors (observed across all 32 decks)
POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)
NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)
NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)

GREY_DARK = RGBColor(0x40, 0x40, 0x40)
GREY_MID = RGBColor(0x59, 0x59, 0x59)
GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)
GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)


# ─────────────────────────────────────────────────────────────────────
# CLIENT-level defaults (fonts, heading color, template path)
# ─────────────────────────────────────────────────────────────────────

CLIENT = {
    "JJ": {
        "font_heading":  "Johnson Display",
        "font_body":     "Johnson Text",
        "heading_color": RGBColor(0x00, 0x1E, 0x60),
        "template_path": None,
    },
    "AZN": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x59, 0x59, 0x59),
        "template_path": None,
    },
    "DSI": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x1E, 0x22, 0xAA),
        "template_path": None,
    },
    "DSI_AZN": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x1E, 0x22, 0xAA),
        "template_path": None,
    },
    "GSK": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0xF3, 0x66, 0x33),
        "template_path": None,
    },
    "PFIZER": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x00, 0x63, 0xC3),
        "template_path": None,
    },
    "EMD": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "template_path": None,
    },
    "REGENERON": {
        "font_heading":  "Trade Gothic LT Std",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x00, 0x74, 0x5A),
        "template_path": None,
    },
    "REGENERON_SANOFI": {
        "font_heading":  "Trade Gothic LT Std",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x00, 0x74, 0x5A),
        "template_path": None,
    },
    "NOVARTIS": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x01, 0x8E, 0x86),
        "template_path": None,
    },
    "AMGEN": {
        "font_heading":  "Century Gothic",
        "font_body":     "Century Gothic",
        "heading_color": RGBColor(0x00, 0x3C, 0x71),
        "template_path": None,
    },
    "LEO": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0xC0, 0x14, 0xA3),
        "template_path": None,
    },
    "ALEXION": {
        "font_heading":  "Arial Black",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x0E, 0x87, 0x79),
        "template_path": None,
    },
    "BL": {
        "font_heading":  "Avenir Next LT Pro",
        "font_body":     "Century Gothic",
        "heading_color": RGBColor(0x40, 0x02, 0x86),
        "template_path": None,
    },
    "OTSUKA": {
        "font_heading":  "Calibri",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "template_path": None,
    },
    "APELLIS": {
        "font_heading":  "Calibri Light",
        "font_body":     "Calibri",
        "heading_color": RGBColor(0xFC, 0x3B, 0x6E),
        "template_path": None,
    },
    "IPSEN": {
        "font_heading":  "Rethink Sans",
        "font_body":     "Calibri",
        "heading_color": RGBColor(0x54, 0xAC, 0x65),
        "template_path": None,
    },
    "CCA": {
        "font_heading":  "Century Gothic",
        "font_body":     "Century Gothic",
        "heading_color": RGBColor(0xF2, 0x8E, 0x2B),
        "template_path": None,
    },
    "UNKNOWN": {
        "font_heading":  "Arial",
        "font_body":     "Arial",
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "template_path": None,
    },
}


# ─────────────────────────────────────────────────────────────────────
# BRAND-level definitions (one per pharma brand, not per client)
# ─────────────────────────────────────────────────────────────────────

BRAND = {
    # ADBRY — LEO / Dermatology/AD (2 decks)
    "ADBRY": {
        "client":           "LEO",
        "therapy_area":     "Dermatology/AD",
        "primary_current":  RGBColor(0xC0, 0x14, 0xA3),  # brand primary
        "primary_prior":    RGBColor(0x6F, 0x43, 0x9A),  # tint for prior wave
        "secondary":        RGBColor(0xD5, 0x68, 0x5F),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#C014A3", "#D5685F", "#6F439A", "#FED006", "#EAAFE0", "#009B77", "#D561C1", "#90117A"],
    },

    # DATROWAY — DSI_AZN / Oncology/mBC (2 decks)
    "DATROWAY": {
        "client":           "DSI_AZN",
        "therapy_area":     "Oncology/mBC",
        "primary_current":  RGBColor(0x1E, 0x22, 0xAA),  # brand primary
        "primary_prior":    RGBColor(0x64, 0x34, 0x66),  # tint for prior wave
        "secondary":        RGBColor(0xEE, 0x76, 0x23),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#1E22AA", "#EE7623", "#643466", "#FBCCB1", "#018EAF", "#00ABC7", "#94A61F", "#F5AD7B"],
    },

    # ABILIFY_MAINTENA — OTSUKA / Psychiatry (1 deck)
    "ABILIFY_MAINTENA": {
        "client":           "OTSUKA",
        "therapy_area":     "Psychiatry",
        "primary_current":  RGBColor(0xFF, 0xC0, 0x00),  # brand primary
        "primary_prior":    RGBColor(0x7F, 0x47, 0xAA),  # tint for prior wave
        "secondary":        RGBColor(0x2D, 0x5C, 0xA2),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FFC000", "#2D5CA2", "#7F47AA", "#7D44A9", "#3B5998", "#7E97CD", "#B3C9EA", "#0000FF"],
    },

    # ABRYSVO — PFIZER / Vaccines/Maternal (1 deck)
    "ABRYSVO": {
        "client":           "PFIZER",
        "therapy_area":     "Vaccines/Maternal",
        "primary_current":  RGBColor(0x00, 0x00, 0xC9),  # brand primary
        "primary_prior":    RGBColor(0x43, 0x96, 0x4A),  # tint for prior wave
        "secondary":        RGBColor(0xE3, 0x56, 0xDC),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#0000C9", "#E356DC", "#43964A", "#F49C34", "#BA2A4C", "#0095FF", "#D3770B", "#0070BF"],
    },

    # ALL_PET_KPI — UNKNOWN / Cross-brand (1 deck)
    "ALL_PET_KPI": {
        "client":           "UNKNOWN",
        "therapy_area":     "Cross-brand",
        "primary_current":  RGBColor(0x29, 0x49, 0x83),  # brand primary
        "primary_prior":    RGBColor(0x8F, 0xAA, 0xDC),  # tint for prior wave
        "secondary":        RGBColor(0x5C, 0x84, 0xCC),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#294983", "#5C84CC", "#8FAADC", "#F0406E", "#82BC00", "#133053"],
    },

    # BAVENCIO — EMD / Oncology (1 deck)
    "BAVENCIO": {
        "client":           "EMD",
        "therapy_area":     "Oncology",
        "primary_current":  RGBColor(0x00, 0x00, 0xC9),  # brand primary
        "primary_prior":    RGBColor(0x40, 0x60, 0xAF),  # tint for prior wave
        "secondary":        RGBColor(0x6D, 0x5A, 0x7A),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#0000C9", "#6D5A7A", "#4060AF", "#081E3D", "#808080", "#6FBF4A", "#E4B77F", "#43964A"],
    },

    # BLENREP — GSK / Oncology/MM (1 deck)
    "BLENREP": {
        "client":           "GSK",
        "therapy_area":     "Oncology/MM",
        "primary_current":  RGBColor(0x3A, 0xB5, 0x1D),  # brand primary
        "primary_prior":    RGBColor(0x5C, 0x10, 0x3B),  # tint for prior wave
        "secondary":        RGBColor(0x66, 0x58, 0xA6),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#3AB51D", "#6658A6", "#5C103B", "#668EDD", "#7030A0", "#17B3AF", "#1B3B7A", "#E21860"],
    },

    # BONE_HCP_TRACKER — UNKNOWN / Bone (1 deck)
    "BONE_HCP_TRACKER": {
        "client":           "UNKNOWN",
        "therapy_area":     "Bone",
        "primary_current":  RGBColor(0xFF, 0x99, 0x33),  # brand primary
        "primary_prior":    RGBColor(0x00, 0xB0, 0xF0),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x92, 0x01),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FF9933", "#009201", "#00B0F0", "#58193D", "#0063C3", "#4BACC6", "#E46C0A", "#425563"],
    },

    # CALQUENCE — AZN / Oncology/CLL (1 deck)
    "CALQUENCE": {
        "client":           "AZN",
        "therapy_area":     "Oncology/CLL",
        "primary_current":  RGBColor(0x00, 0xA2, 0xE0),  # brand primary
        "primary_prior":    RGBColor(0xDC, 0x44, 0x05),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x20, 0x60),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#00A2E0", "#002060", "#DC4405", "#FFC72C", "#EE3EDD", "#E44405", "#001F58", "#E97142"],
    },

    # CCA_PP_TRACKER — CCA / Cross-brand (1 deck)
    "CCA_PP_TRACKER": {
        "client":           "CCA",
        "therapy_area":     "Cross-brand",
        "primary_current":  RGBColor(0xF2, 0x8E, 0x2B),  # brand primary
        "primary_prior":    RGBColor(0xBD, 0xD7, 0xEE),  # tint for prior wave
        "secondary":        RGBColor(0x8F, 0xAA, 0xDC),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#F28E2B", "#8FAADC", "#BDD7EE", "#AFED5D", "#FFABD5", "#FF6700", "#FF8F43", "#FFBC8F"],
    },

    # DARZALEX — JJ / Oncology/MM (1 deck)
    "DARZALEX": {
        "client":           "JJ",
        "therapy_area":     "Oncology/MM",
        "primary_current":  RGBColor(0x6F, 0xC6, 0xC1),  # brand primary
        "primary_prior":    RGBColor(0x2A, 0x41, 0xC1),  # tint for prior wave
        "secondary":        RGBColor(0xBD, 0x05, 0xED),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#6FC6C1", "#BD05ED", "#2A41C1", "#FF6A5A", "#FFC000", "#0302C8", "#2D5BEF", "#AC4069"],
    },

    # DUPIXENT_EoE — REGENERON_SANOFI / Immunology/EoE (1 deck)
    "DUPIXENT_EoE": {
        "client":           "REGENERON_SANOFI",
        "therapy_area":     "Immunology/EoE",
        "primary_current":  RGBColor(0x21, 0x94, 0x91),  # brand primary
        "primary_prior":    RGBColor(0xE4, 0x6C, 0x0A),  # tint for prior wave
        "secondary":        RGBColor(0xF7, 0x96, 0x46),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#219491", "#F79646", "#E46C0A", "#8EB4E3", "#00745A", "#7030A0", "#B370A4", "#7ABFBD"],
    },

    # EMPAVELI — APELLIS / Hematology (1 deck)
    "EMPAVELI": {
        "client":           "APELLIS",
        "therapy_area":     "Hematology",
        "primary_current":  RGBColor(0xFC, 0x3B, 0x6E),  # brand primary
        "primary_prior":    RGBColor(0x81, 0xE3, 0xE3),  # tint for prior wave
        "secondary":        RGBColor(0x30, 0xCF, 0xD0),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FC3B6E", "#30CFD0", "#81E3E3", "#5A4986", "#FD7FA0", "#FEB0C4", "#FFC000", "#163A6B"],
    },

    # ENHERTU — DSI_AZN / Oncology/HER2 (1 deck)
    "ENHERTU": {
        "client":           "DSI_AZN",
        "therapy_area":     "Oncology/HER2",
        "primary_current":  RGBColor(0xFF, 0x88, 0x13),  # brand primary
        "primary_prior":    RGBColor(0x33, 0xA4, 0xFF),  # tint for prior wave
        "secondary":        RGBColor(0xE1, 0x57, 0x59),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FF8813", "#E15759", "#33A4FF", "#59A14F", "#9467BD", "#1D7B87", "#91C3D5", "#F0FF4D"],
    },

    # JEMPERLI-ZEJULA — GSK / Oncology/Endometrial (1 deck)
    "JEMPERLI-ZEJULA": {
        "client":           "GSK",
        "therapy_area":     "Oncology/Endometrial",
        "primary_current":  RGBColor(0xD5, 0x19, 0x00),  # brand primary
        "primary_prior":    RGBColor(0x70, 0x63, 0x52),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x84, 0x7C),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#D51900", "#00847C", "#706352", "#A9A197", "#3F348E", "#E7004E", "#668EDD", "#FFC000"],
    },

    # LIBTAYO — REGENERON / Oncology/NMSC (1 deck)
    "LIBTAYO": {
        "client":           "REGENERON",
        "therapy_area":     "Oncology/NMSC",
        "primary_current":  RGBColor(0xDC, 0x00, 0x77),  # brand primary
        "primary_prior":    RGBColor(0xC8, 0x33, 0x33),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x4F, 0x6F),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#DC0077", "#004F6F", "#C83333", "#96BBBB", "#C19875", "#338533", "#D34D4D", "#668998"],
    },

    # LOKELMA — AZN / Nephrology (1 deck)
    "LOKELMA": {
        "client":           "AZN",
        "therapy_area":     "Nephrology",
        "primary_current":  RGBColor(0x00, 0x94, 0x7F),  # brand primary
        "primary_prior":    RGBColor(0x4A, 0x2C, 0x8C),  # tint for prior wave
        "secondary":        RGBColor(0x8F, 0x20, 0x64),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#00947F", "#8F2064", "#4A2C8C", "#00EAC9", "#B6FFF4", "#003965", "#00B0F0", "#D1C5ED"],
    },

    # LYNPARZA — AZN / Oncology (1 deck)
    "LYNPARZA": {
        "client":           "AZN",
        "therapy_area":     "Oncology",
        "primary_current":  RGBColor(0x94, 0xD4, 0x48),  # brand primary
        "primary_prior":    RGBColor(0x00, 0xB3, 0xFF),  # tint for prior wave
        "secondary":        RGBColor(0xFF, 0xBA, 0x00),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#94D448", "#FFBA00", "#00B3FF", "#34A355", "#D2CEC8", "#F74548", "#D5D5D5", "#9FD448"],
    },

    # MIEBO — BL / Ophthalmology/DED (1 deck)
    "MIEBO": {
        "client":           "BL",
        "therapy_area":     "Ophthalmology/DED",
        "primary_current":  RGBColor(0x00, 0xA9, 0xEB),  # brand primary
        "primary_prior":    RGBColor(0xA8, 0x10, 0x9D),  # tint for prior wave
        "secondary":        RGBColor(0x40, 0x02, 0x86),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#00A9EB", "#400286", "#A8109D", "#FF1119", "#413A5F", "#000F9F", "#FFC000", "#F78626"],
    },

    # OJJAARA — GSK / Hematology/MF (1 deck)
    "OJJAARA": {
        "client":           "GSK",
        "therapy_area":     "Hematology/MF",
        "primary_current":  RGBColor(0x6B, 0xB2, 0x3E),  # brand primary
        "primary_prior":    RGBColor(0xEE, 0x53, 0x40),  # tint for prior wave
        "secondary":        RGBColor(0x61, 0x14, 0xFF),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#6BB23E", "#6114FF", "#EE5340", "#3A5C6D", "#6014FF", "#DDDDDD", "#FC6666", "#AB162C"],
    },

    # ONIVYDE — IPSEN / Oncology/Pancreatic (1 deck)
    "ONIVYDE": {
        "client":           "IPSEN",
        "therapy_area":     "Oncology/Pancreatic",
        "primary_current":  RGBColor(0x54, 0xAC, 0x65),  # brand primary
        "primary_prior":    RGBColor(0x8E, 0xC8, 0x99),  # tint for prior wave
        "secondary":        RGBColor(0xC8, 0x48, 0x74),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#54AC65", "#C84874", "#8EC899", "#BBDEC2", "#D2E9D6", "#A9D5B2", "#00B039", "#387343"],
    },

    # OTEZLA — AMGEN / Dermatology/Psoriasis (1 deck)
    "OTEZLA": {
        "client":           "AMGEN",
        "therapy_area":     "Dermatology/Psoriasis",
        "primary_current":  RGBColor(0x1F, 0x49, 0x7D),  # brand primary
        "primary_prior":    RGBColor(0x00, 0xA3, 0xDF),  # tint for prior wave
        "secondary":        RGBColor(0x6F, 0x8D, 0xB4),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#1F497D", "#6F8DB4", "#00A3DF", "#4C35DE", "#A5A5A5", "#2E75B6", "#FFB81C", "#C303D4"],
    },

    # PHYSICIANS_SFE_PET — UNKNOWN / Cross-brand (1 deck)
    "PHYSICIANS_SFE_PET": {
        "client":           "UNKNOWN",
        "therapy_area":     "Cross-brand",
        "primary_current":  RGBColor(0xE7, 0x00, 0x4C),  # brand primary
        "primary_prior":    RGBColor(0x6B, 0xC9, 0xEF),  # tint for prior wave
        "secondary":        RGBColor(0xC7, 0xA0, 0x13),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#E7004C", "#C7A013", "#6BC9EF", "#006838", "#FF6969", "#00607C", "#4E79A7", "#00A44A"],
    },

    # RHAPSIDO — NOVARTIS / Immunology (1 deck)
    "RHAPSIDO": {
        "client":           "NOVARTIS",
        "therapy_area":     "Immunology",
        "primary_current":  RGBColor(0x01, 0x8E, 0x86),  # brand primary
        "primary_prior":    RGBColor(0x00, 0x20, 0x68),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x70, 0xFE),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#018E86", "#0070FE", "#002068", "#8F2DDE", "#852065", "#B56FA3", "#61A8FF", "#50E2D0"],
    },

    # RYBREVANT — JJ / Oncology/NSCLC (1 deck)
    "RYBREVANT": {
        "client":           "JJ",
        "therapy_area":     "Oncology/NSCLC",
        "primary_current":  RGBColor(0xF7, 0x58, 0x24),  # brand primary
        "primary_prior":    RGBColor(0xF9, 0x59, 0x24),  # tint for prior wave
        "secondary":        RGBColor(0xFB, 0xAB, 0x91),  # accent / comparison
        "competitor_brand": "TAGRISSO",
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#F75824", "#FBAB91", "#F95924", "#7030A0", "#BAB0AC", "#1F698F", "#77BEE3", "#7F46AA"],
    },

    # TEPEZZA — JJ / Endocrinology/TED (1 deck)
    "TEPEZZA": {
        "client":           "JJ",
        "therapy_area":     "Endocrinology/TED",
        "primary_current":  RGBColor(0x7F, 0xB1, 0xE1),  # brand primary
        "primary_prior":    RGBColor(0x37, 0xCF, 0xCA),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x63, 0xC3),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#7FB1E1", "#0063C3", "#37CFCA", "#228D8A", "#A0D0FF", "#CBF3F1", "#00A3C4", "#ADADAD"],
    },

    # TEZSPIRE — AZN / Respiratory (1 deck)
    "TEZSPIRE": {
        "client":           "AZN",
        "therapy_area":     "Respiratory",
        "primary_current":  RGBColor(0x00, 0x38, 0x65),  # brand primary
        "primary_prior":    RGBColor(0xFF, 0xD0, 0x5D),  # tint for prior wave
        "secondary":        RGBColor(0x85, 0xBF, 0xFF),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#003865", "#85BFFF", "#FFD05D", "#00B5E2", "#C688E8", "#830051", "#9B37FF", "#EF426F"],
    },

    # TRUQAP — AZN / Oncology/Breast (1 deck)
    "TRUQAP": {
        "client":           "AZN",
        "therapy_area":     "Oncology/Breast",
        "primary_current":  RGBColor(0x25, 0x0E, 0x62),  # brand primary
        "primary_prior":    RGBColor(0x00, 0x69, 0x37),  # tint for prior wave
        "secondary":        RGBColor(0x65, 0x66, 0x81),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#250E62", "#656681", "#006937", "#FFCB05", "#F5A899", "#EE7623", "#05CDCD", "#006680"],
    },

    # ULTOMIRIS — ALEXION / Hematology/PNH (1 deck)
    "ULTOMIRIS": {
        "client":           "ALEXION",
        "therapy_area":     "Hematology/PNH",
        "primary_current":  RGBColor(0x0E, 0x87, 0x79),  # brand primary
        "primary_prior":    RGBColor(0xE3, 0x51, 0x05),  # tint for prior wave
        "secondary":        RGBColor(0x70, 0x30, 0xA0),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#0E8779", "#7030A0", "#E35105", "#799A01", "#FFA300", "#FF8181", "#84CA68", "#009886"],
    },

    # UPLIZNA — AMGEN / Neurology/NMOSD (1 deck)
    "UPLIZNA": {
        "client":           "AMGEN",
        "therapy_area":     "Neurology/NMOSD",
        "primary_current":  RGBColor(0x81, 0x3F, 0x97),  # brand primary
        "primary_prior":    RGBColor(0x00, 0x3C, 0x71),  # tint for prior wave
        "secondary":        RGBColor(0x00, 0x7D, 0x6D),  # accent / comparison
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#813F97", "#007D6D", "#003C71", "#DA205F", "#00B0F0", "#E46C0A", "#D9D9D9", "#425563"],
    },

}


def get_brand(brand_name: str) -> dict:
    """Look up brand config. Merges BRAND entry with its CLIENT entry so callers
    get fonts/heading_color/template_path inline without a second lookup.
    """
    key = brand_name.upper().replace(' ', '_').replace('-', '_')
    if key not in BRAND:
        raise KeyError(f"Unknown brand {brand_name!r}. Available: {sorted(BRAND.keys())}")
    brand = BRAND[key].copy()
    client_key = brand["client"]
    if client_key in CLIENT:
        # Client-level defaults merged in (brand can override, but doesn't today)
        for k, v in CLIENT[client_key].items():
            brand.setdefault(k, v)
    return brand


def get_competitor(brand_name: str) -> dict | None:
    """If the brand has a known competitor, return that competitor's BRAND dict."""
    brand = get_brand(brand_name)
    comp = brand.get('competitor_brand')
    if comp and comp.upper() in BRAND:
        return get_brand(comp)
    return None
