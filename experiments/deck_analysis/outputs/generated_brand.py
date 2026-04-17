"""
Client-level brand defaults — fonts, heading colors, observed palettes.

Generated from deck analysis of 551 decks across 78 clients.
See experiments/deck_analysis/mass_deck_scanner.py for the analysis.

CLIENT{} provides client-level defaults (shared across all brands for that client).
Per-brand entries (product-level colors) live in BRAND{} and are hand-curated.

Each CLIENT entry provides:
  - font_heading:        headline font
  - font_body:           body text font
  - heading_color:       headline text color
  - observed_palette:    top 8 series colors observed across all decks for this client
  - deck_count:          number of decks analyzed
"""
from __future__ import annotations

from pptx.dml.color import RGBColor


# Universal delta colors — observed across all decks
POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)
NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)
NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)

# Standard greys
GREY_DARK = RGBColor(0x40, 0x40, 0x40)
GREY_MID = RGBColor(0x59, 0x59, 0x59)
GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)
GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)


CLIENT = {
    # --- Amgen (AMG) (55 decks) ---
    "AMGEN": {
        "heading_color": RGBColor(0x00, 0x3C, 0x71),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 55,
        "_observed_palette": ["0063C3", "1F497D", "D6D2D0", "76C269", "813F97", "92D050", "C00000", "00B050"],
    },

    # --- daiichi sankyo (31 decks) ---
    "DAIICHI_SANKYO": {
        "heading_color": RGBColor(0x40, 0x40, 0x40),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 31,
        "_observed_palette": ["FF8813", "D9D9D9", "7F7F7F", "FFBFFF", "34A355", "33A4FF", "EE7623", "8F66A9"],
    },

    # --- AstraZeneca (AZN) (30 decks) ---
    "ASTRAZENECA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 30,
        "_observed_palette": ["00B050", "EE7623", "BFBFBF", "1E22AA", "A6A6A6", "FF0000", "643466", "002060"],
    },

    # --- AbbVie (ABV) (28 decks) ---
    "ABBVIE": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 28,
        "_observed_palette": ["99CC00", "007B80", "2091AC", "96D801", "D73749", "FFD100", "FFC000", "532F88"],
    },

    # --- Gilead (GLD) (22 decks) ---
    "GILEAD": {
        "heading_color": RGBColor(0x54, 0x56, 0x5B),
        "font_heading": "Trebuchet MS",
        "font_body": "Trebuchet MS",
        "template_path": None,
        "deck_count": 22,
        "_observed_palette": ["D11241", "A6A6A6", "E7751F", "002060", "00B050", "B9BBBF", "FF0000", "E86A6A"],
    },

    # --- Merck (MER) (19 decks) ---
    "MERCK": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 19,
        "_observed_palette": ["00B050", "801851", "E886BC", "014C35", "BFBFBF", "228848", "004D74", "00857C"],
    },

    # --- Novartis (NVS) (19 decks) ---
    "NOVARTIS": {
        "heading_color": RGBColor(0x1F, 0x49, 0x7D),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 19,
        "_observed_palette": ["018E86", "852065", "92D050", "0C68B0", "830F66", "0000C9", "7C1C33", "002068"],
    },

    # --- Deciphera (18 decks) ---
    "DECIPHERA": {
        "heading_color": RGBColor(0x17, 0x46, 0x6B),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 18,
        "_observed_palette": ["5D2759", "F1B914", "5DBEBF", "EB8125", "C72750", "92D050", "C00000", "008A00"],
    },

    # --- Agios _ Servier (16 decks) ---
    "AGIOS___SERVIER": {
        "heading_color": RGBColor(0x24, 0x22, 0x69),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 16,
        "_observed_palette": ["242269", "D9D9D9", "BFBFBF", "A4C4EE", "1A355D", "A6A6A6", "242168", "D3DEF2"],
    },

    # --- Pfizer (13 decks) ---
    "PFIZER": {
        "heading_color": RGBColor(0x00, 0x63, 0xC3),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 13,
        "_observed_palette": ["0000C9", "602A81", "B3C9EA", "FFFFFF", "05BE66", "6B0834", "43964A", "E9EED9"],
    },

    # --- Sanofi - Genzyme (SAN) (13 decks) ---
    "SANOFI___GENZYME": {
        "heading_color": RGBColor(0x7F, 0x7F, 0x7F),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 13,
        "_observed_palette": ["00B050", "219491", "BFBFBF", "525CA3", "D9D9D9", "C00000", "249592", "2A3C97"],
    },

    # --- BioMarin (12 decks) ---
    "BIOMARIN": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial ",
        "font_body": "Arial ",
        "template_path": None,
        "deck_count": 12,
        "_observed_palette": ["FBD401", "28509C", "091A79", "ED1849", "ED037C", "A9208E", "081A79", "B60018"],
    },

    # --- BridgeBio (12 decks) ---
    "BRIDGEBIO": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 12,
        "_observed_palette": ["507892", "AF1634", "E62E4B", "2F6938", "5B2478", "DF7B17", "5AC2ED", "051D74"],
    },

    # --- Blueprint Medicines (BPM) (11 decks) ---
    "BLUEPRINT_MEDICINES": {
        "heading_color": RGBColor(0x00, 0x26, 0x3D),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 11,
        "_observed_palette": ["1E5271", "BBB5AF", "638326", "C2CA7E", "67ADD7", "C6BFB6", "7030A0", "722E3E"],
    },

    # --- Genentech (GNE) (11 decks) ---
    "GENENTECH": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Gene Sans",
        "font_body": "Gene Sans",
        "template_path": None,
        "deck_count": 11,
        "_observed_palette": ["C00000", "0E326F", "7B4E96", "9ECEEB", "00ADA6", "CCCDCD", "999B9B", "0070C0"],
    },

    # --- Alnylam (ALN) (10 decks) ---
    "ALNYLAM": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 10,
        "_observed_palette": ["E7004C", "00607C", "6BC9EF", "C00000", "006838", "002060", "FF6969", "A6A6A6"],
    },

    # --- Foundation Medicine (10 decks) ---
    "FOUNDATION_MEDICINE": {
        "heading_color": RGBColor(0x32, 0x3F, 0x4A),
        "font_heading": "Gotham Book",
        "font_body": "Gotham Book",
        "template_path": None,
        "deck_count": 10,
        "_observed_palette": ["435363", "A0D8B3", "FF9466", "267270", "38ACA8", "00B0F0", "A746F0", "64CCC9"],
    },

    # --- Takeda (TAK) (10 decks) ---
    "TAKEDA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 10,
        "_observed_palette": ["00B050", "BFBFBF", "00823B", "C00000", "731013", "E1242A", "00708A", "0EB2CD"],
    },

    # --- BMS (9 decks) ---
    "BMS": {
        "heading_color": RGBColor(0x59, 0x54, 0x54),
        "font_heading": "Trebuchet MS",
        "font_body": "Trebuchet MS",
        "template_path": None,
        "deck_count": 9,
        "_observed_palette": ["BFBFBF", "00B050", "0E2D75", "FB796E", "780081", "002140", "FFD186", "7F7F7F"],
    },

    # --- GSK (9 decks) ---
    "GSK": {
        "heading_color": RGBColor(0xF3, 0x66, 0x33),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 9,
        "_observed_palette": ["00B050", "A6A6A6", "EE5340", "C00000", "0B2745", "D51900", "D0400C", "6014FF"],
    },

    # --- Seagen (Pfizer Onc) (9 decks) ---
    "SEAGEN": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 9,
        "_observed_palette": ["8275E9", "F7C99B", "ADADAD", "A2BDE2", "EBC871", "7F7F7F", "FFE5E5", "BA2A4C"],
    },

    # --- UCB (9 decks) ---
    "UCB": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Tahoma",
        "font_body": "Tahoma",
        "template_path": None,
        "deck_count": 9,
        "_observed_palette": ["7030A0", "92D050", "646D6D", "C198E0", "E74F32", "8D1838", "445A6A", "90959C"],
    },

    # --- ArgenX (AGX) (8 decks) ---
    "ARGENX": {
        "heading_color": RGBColor(0x00, 0x2B, 0x48),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["9137AF", "002B48", "1F5FA0", "FF9300", "8E3F9A", "780032", "00223C", "B00075"],
    },

    # --- Ipsen (8 decks) ---
    "IPSEN": {
        "heading_color": RGBColor(0x00, 0x0E, 0x56),
        "font_heading": "Lato",
        "font_body": "Lato",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["00B050", "54AC65", "461C6F", "D38A4F", "C84874", "C00000", "BFBFBF", "0052B2"],
    },

    # --- ITF Therapeutics (7 decks) ---
    "ITF_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["0E8B37", "A6A6A6", "840B55", "004F9E", "2CB6B3", "0E8D38", "A7C978", "1F0064"],
    },

    # --- Jazz (7 decks) ---
    "JAZZ": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["0070C0", "8497B0", "D6EBD4", "FFC000", "FF69D6", "9A0072", "00B050", "C00000"],
    },

    # --- LEO Pharma (7 decks) ---
    "LEO_PHARMA": {
        "heading_color": RGBColor(0x20, 0x28, 0x2F),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["FF5067", "C014A3", "D5685F", "9AE649", "5387A6", "009B77", "260077", "6F439A"],
    },

    # --- Regeneron (RGN) (7 decks) ---
    "REGENERON": {
        "heading_color": RGBColor(0x00, 0x74, 0x5A),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["33B412", "D50056", "002060", "FEB8B8", "DC0077", "0074CC", "A6A6A6", "00B050"],
    },

    # --- Sobi Pharma (SOB) (7 decks) ---
    "SOBI_PHARMA": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["2D5E77", "003B5C", "F15A22", "00B050", "4CA1A6", "ABBFC9", "F79C7A", "6CB33E"],
    },

    # --- Arcutis (ARC) (6 decks) ---
    "ARCUTIS": {
        "heading_color": RGBColor(0xE0, 0xB4, 0x1C),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 6,
        "_observed_palette": ["FFC000", "229719", "7F7F7F", "D34C4C", "CB35DB", "B4C7E7", "DE5754", "8271E5"],
    },

    # --- CSL Behring (6 decks) ---
    "CSL_BEHRING": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Montserrat",
        "font_body": "Montserrat",
        "template_path": None,
        "deck_count": 6,
        "_observed_palette": ["601D72", "52636C", "FEA3A6", "890207", "F16B50", "00B0F0", "952DB1", "ED7A2B"],
    },

    # --- Genmab (6 decks) ---
    "GENMAB": {
        "heading_color": RGBColor(0xFF, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 6,
        "_observed_palette": ["E6AF00", "818285", "7294A5", "CAEDF7", "FF5050", "B3B4B6", "7030A0", "FFC000"],
    },

    # --- Ascendis (5 decks) ---
    "ASCENDIS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": [],
    },

    # --- Corcept (5 decks) ---
    "CORCEPT": {
        "heading_color": RGBColor(0x23, 0x00, 0x4C),
        "font_heading": "Lato",
        "font_body": "Lato",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": ["28378E", "2A3C97", "EE7623", "1E40BE", "5BC67C", "43248C", "C00000", "FF9393"],
    },

    # --- Ionis (5 decks) ---
    "IONIS": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": [],
    },

    # --- Neurocrine (5 decks) ---
    "NEUROCRINE": {
        "heading_color": RGBColor(0x2B, 0x5F, 0x51),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": ["439F5B", "D34D4D", "A6A6A6", "ED9C97", "B6465F", "B092B4", "ED7D31", "F2BAB6"],
    },

    # --- Alexion Pharmaceuticals (4 decks) ---
    "ALEXION_PHARMACEUTICALS": {
        "heading_color": RGBColor(0x00, 0x1E, 0x60),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["009886", "415C79", "BC3666", "76C23A", "EC9DB8", "59B7C7", "8B2DAB", "E35105"],
    },

    # --- Celltrion (4 decks) ---
    "CELLTRION": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["00B050", "FF0000", "92D050", "D63042", "AE178D", "7BD7D1", "AD641E", "8064A2"],
    },

    # --- Eisai (4 decks) ---
    "EISAI": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["880043", "7C164A", "FCBA3D", "FFA700", "7F7F7F", "D9D9D9", "8C3360", "006600"],
    },

    # --- Exelixis (4 decks) ---
    "EXELIXIS": {
        "heading_color": RGBColor(0x33, 0x33, 0x33),
        "font_heading": "Arial(Body)",
        "font_body": "Arial(Body)",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["233973", "00B050", "A11940", "595959", "00B0F0", "ED7D31", "BFBFBF", "F4B183"],
    },

    # --- IntraCellular (4 decks) ---
    "INTRACELLULAR": {
        "heading_color": RGBColor(0x23, 0x24, 0x26),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["04566E", "00ABA6", "23668A", "7F7F7F", "E7BA55", "D9D9D9", "249592", "31BDE8"],
    },

    # --- Sumitomo (SMPA) (4 decks) ---
    "SUMITOMO": {
        "heading_color": RGBColor(0x54, 0x56, 0x5B),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["F05562", "00B050", "59877D", "F7C001", "FFC000", "071D49", "F75865", "007A4C"],
    },

    # --- Alkermes (ALK) (3 decks) ---
    "ALKERMES": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["95054B", "BF85B6", "088DD0", "7030A0", "E65D02", "3C7E79", "6096C8", "173D8D"],
    },

    # --- Bayer (BAY) (3 decks) ---
    "BAYER": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Arial (Body)",
        "font_body": "Arial (Body)",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["54306A", "06AACD", "D21D00", "1F4288", "002060", "42A2FF", "F4B183", "92D050"],
    },

    # --- Novartis Gene Therapies (AveXis) (3 decks) ---
    "NOVARTIS_GENE_THERAPIES": {
        "heading_color": RGBColor(0x40, 0x40, 0x40),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["FF5500", "9D0275", "0F5496", "7F7F7F", "59A14F", "E15759", "EDC948", "BFBFBF"],
    },

    # --- Otsuka (3 decks) ---
    "OTSUKA": {
        "heading_color": RGBColor(0x49, 0x57, 0x6F),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["C00000", "2C976D", "D0CECE", "E3C032", "EC002F", "B2B2B2", "006666", "FFC000"],
    },

    # --- Theravance (THV) (3 decks) ---
    "THERAVANCE": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["005EB8", "BF2F01", "8FD572", "72D1DD", "006241", "8D4EDE", "FF9933", "009E82"],
    },

    # --- Travere (3 decks) ---
    "TRAVERE": {
        "heading_color": RGBColor(0x00, 0x5E, 0x98),
        "font_heading": "Verdana",
        "font_body": "Verdana",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["005E98", "3C9674", "C00000", "7F7F7F", "28ADFF", "B7E4FF", "A3DCFF", "0C5E7E"],
    },

    # --- Vera Therapeutics (3 decks) ---
    "VERA_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Univers",
        "font_body": "Univers",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["105686", "3080E4", "98F883", "E3923D", "FFD11C", "7F7F7F", "66A4FF", "0068FF"],
    },

    # --- Apellis (APL) (2 decks) ---
    "APELLIS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["FC3B6E", "E7E6E6", "30CFD0", "81E3E3", "5A4986", "FD7FA0", "FEB0C4", "FFC000"],
    },

    # --- Astellas (2 decks) ---
    "ASTELLAS": {
        "heading_color": RGBColor(0x00, 0x20, 0x60),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["92D050", "E37B13", "960048", "F26F15", "002060", "CA7FA3", "9C2AB5", "FFC000"],
    },

    # --- AVEO Oncology (2 decks) ---
    "AVEO_ONCOLOGY": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["395F79", "D6037D", "A6103D", "6F1A44", "333F8D", "6096C8", "99B3C1", "4D7A93"],
    },

    # --- Cytokinetics (CYTO) (2 decks) ---
    "CYTOKINETICS": {
        "heading_color": RGBColor(0x75, 0xA7, 0x33),
        "font_heading": "Aptos",
        "font_body": "Aptos",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["0A367D", "BE215E", "D9D9D9", "859BBE", "125FDC", "80ACF4", "CADCFA", "FFABD3"],
    },

    # --- EMD Serono (2 decks) ---
    "EMD_SERONO": {
        "heading_color": RGBColor(0x65, 0x18, 0x5A),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["672A76", "65185A", "4060AF", "801514", "EC8484", "BFBFBF", "54468E", "A52A2A"],
    },

    # --- Kyowa Kyrin (KKN) (2 decks) ---
    "KYOWA_KYRIN": {
        "heading_color": RGBColor(0x38, 0xC7, 0xA5),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": [],
    },

    # --- Natera (NAT) (2 decks) ---
    "NATERA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["44C04F", "0070F2", "9E0000", "FA18FA", "FF494B", "A973FD", "FFAB00", "008A00"],
    },

    # --- Novo Nordisk (2 decks) ---
    "NOVO_NORDISK": {
        "heading_color": RGBColor(0x00, 0x19, 0x65),
        "font_heading": "Apis For Office",
        "font_body": "Apis For Office",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["001965", "219491", "2D79C1", "0C337E", "2A918B", "E888A8", "00B050", "DE5281"],
    },

    # --- PTC Therapeutics (PTC) (2 decks) ---
    "PTC_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["804080", "48BC6F", "BFBFBF", "FF7F7F", "224468", "346323", "23466C", "C00000"],
    },

    # --- SpringWorks Therapeutics (2 decks) ---
    "SPRINGWORKS_THERAPEUTICS": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["B88504"],
    },

    # --- TG Therapeutics (2 decks) ---
    "TG_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Crimson Pro ExtraBold",
        "font_body": "Crimson Pro ExtraBold",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["00B050", "0E326F", "FF9393", "92D050", "85C226", "40BB9A", "83D977", "6990E7"],
    },

    # --- Vertex (2 decks) ---
    "VERTEX": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["C366A9", "3B7ADA", "EC002F", "FFC000", "7F7F7F", "92D050", "06BED6", "B6B6B6"],
    },

    # --- Arcellx (ARLX) (1 deck) ---
    "ARCELLX": {
        "heading_color": RGBColor(0x0C, 0x25, 0x3C),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["8E3F9A", "002B48", "243A7E", "73AADB", "B00075", "DF7F7F", "7F7F7F", "9137AF"],
    },

    # --- Arvinas (ARVA) (1 deck) ---
    "ARVINAS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["DF7F7F", "4166A1", "F37116", "BFBFBF", "F1CE63", "59A14F", "4CAD4C", "499894"],
    },

    # --- Bausch (BAU) (1 deck) ---
    "BAUSCH": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["00A9EB", "4B34A2", "00B050", "222A35", "A6A6A6", "400286", "D9D9D9", "FFC000"],
    },

    # --- Biogen (1 deck) ---
    "BIOGEN": {
        "heading_color": RGBColor(0x00, 0x20, 0x60),
        "font_heading": "Arial (Body)",
        "font_body": "Arial (Body)",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["C00000", "01738C", "5CA136", "387F75", "A6A6A6", "80B9C5", "BFBEBB", "EF3D33"],
    },

    # --- Braeburn (1 deck) ---
    "BRAEBURN": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["3D66E3", "B80D48", "29C2A1", "BFBFBF", "01437E", "F29724", "701C6A", "6D6D6D"],
    },

    # --- Coherus BioSciences (1 deck) ---
    "COHERUS_BIOSCIENCES": {
        "heading_color": RGBColor(0x00, 0xB1, 0xD1),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["053C69", "F26322", "006C8E", "00A8DF", "BE89C9", "0070C0", "D39092", "A5A5A5"],
    },

    # --- Collegium _ Depomed (1 deck) ---
    "COLLEGIUM___DEPOMED": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial(Body)",
        "font_body": "Arial(Body)",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["30378E", "349442", "A20060", "6FBF4A", "ECDF7C", "EC6F6F", "FF8774", "08747C"],
    },

    # --- Insmed (INM) (1 deck) ---
    "INSMED": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Georgia",
        "font_body": "Georgia",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["440099", "00B050", "D9D9D9", "FB935C", "BFBFBF", "8439BD", "A66BD3", "92D050"],
    },

    # --- Kiniksa (KNS) (1 deck) ---
    "KINIKSA": {
        "heading_color": RGBColor(0x00, 0x20, 0x60),
        "font_heading": "Calibri (Body)",
        "font_body": "Calibri (Body)",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["002060", "00B0F0"],
    },

    # --- Lundbeck (1 deck) ---
    "LUNDBECK": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- Madrigal (MDL) (1 deck) ---
    "MADRIGAL": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- Phathom Pharmaceuticals (PHP) (1 deck) ---
    "PHATHOM_PHARMACEUTICALS": {
        "heading_color": RGBColor(0x00, 0x43, 0x70),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["8B8B8F", "A4CBEA", "2B7BBB", "205C8C", "60BFFF", "003254", "004370", "76B1E0"],
    },

    # --- Puma Biotechnology (1 deck) ---
    "PUMA_BIOTECHNOLOGY": {
        "heading_color": RGBColor(0x14, 0x31, 0x55),
        "font_heading": "Arial(Body)",
        "font_body": "Arial(Body)",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["0083BF", "FFC000", "48A7AA", "00833C", "7F7F7F", "7F53C1", "203864", "70AD47"],
    },

    # --- Servier (SRV) (1 deck) ---
    "SERVIER": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- Viatris - Mylan (1 deck) ---
    "VIATRIS___MYLAN": {
        "heading_color": RGBColor(0xFD, 0xFB, 0xFF),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- ViiV (VII) (1 deck) ---
    "VIIV": {
        "heading_color": RGBColor(0xEB, 0x18, 0x52),
        "font_heading": "Raleway",
        "font_body": "Raleway",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- Vir Biotech (1 deck) ---
    "VIR_BIOTECH": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Lato",
        "font_body": "Lato",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["5DB355", "FFD100", "016CC6", "FFFFFF"],
    },

}


def get_client(client_key: str) -> dict:
    """Look up client defaults by key. Raises KeyError with available keys."""
    key = client_key.upper().replace(' ', '_').replace('-', '_')
    if key not in CLIENT:
        raise KeyError(
            f"Unknown client {client_key!r}. Available: {sorted(CLIENT.keys())}"
        )
    return CLIENT[key]


# ============================================================================
# BRAND{} — Per-product brand entries (auto-extracted from deck filenames)
# ============================================================================

BRAND = {
    # --- UNMAPPED (463 decks, Unknown) ---
    "UNMAPPED": {
        "client": "VIR_BIOTECH",
        "primary_current": RGBColor(0xFF, 0xC0, 0x00),
        "primary_prior": RGBColor(0xD9, 0xD9, 0xD9),
        "competitor_current": RGBColor(0x92, 0xD0, 0x50),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Unknown",
        "_observed_palette": ["FFC000", "92D050", "D9D9D9", "7030A0", "0070C0", "00B0F0", "002060", "0063C3"],
    },

    # --- ENHERTU (23 decks, Oncology/HER2) ---
    "ENHERTU": {
        "client": "DSI_AZN",
        "primary_current": RGBColor(0xFF, 0x88, 0x13),
        "primary_prior": RGBColor(0xEE, 0x76, 0x23),
        "competitor_current": RGBColor(0x34, 0xA3, 0x55),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/HER2",
        "_observed_palette": ["FF8813", "34A355", "EE7623", "F5AD7B", "8F66A9", "33A4FF", "0070C0", "F6BC94"],
    },

    # --- OTEZLA (9 decks, Dermatology/Psoriasis) ---
    "OTEZLA": {
        "client": "AMGEN",
        "primary_current": RGBColor(0x1F, 0x49, 0x7D),
        "primary_prior": RGBColor(0xC4, 0x0E, 0x12),
        "competitor_current": RGBColor(0x00, 0xA3, 0xDF),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Dermatology/Psoriasis",
        "_observed_palette": ["1F497D", "00A3DF", "C40E12", "448D96", "009900", "F3C108", "4C35DE", "11A1A2"],
    },

    # --- UPLIZNA (7 decks, Neurology/NMOSD) ---
    "UPLIZNA": {
        "client": "AMGEN",
        "primary_current": RGBColor(0xD6, 0xD2, 0xD0),
        "primary_prior": RGBColor(0x00, 0x63, 0xC3),
        "competitor_current": RGBColor(0x76, 0xC2, 0x69),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Neurology/NMOSD",
        "_observed_palette": ["D6D2D0", "76C269", "0063C3", "F2BFBF", "2F6D5B", "E57F7F", "5B729B", "813F97"],
    },

    # --- RHAPSIDO (5 decks, Immunology) ---
    "RHAPSIDO": {
        "client": "NOVARTIS",
        "primary_current": RGBColor(0x83, 0x0F, 0x66),
        "primary_prior": RGBColor(0x01, 0x8E, 0x86),
        "competitor_current": RGBColor(0x00, 0x70, 0xFE),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Immunology",
        "_observed_palette": ["830F66", "0070FE", "018E86", "002068", "8F2DDE", "852065", "B56FA3", "50E2D0"],
    },

    # --- ABRYSVO (5 decks, Vaccines/Maternal) ---
    "ABRYSVO": {
        "client": "PFIZER",
        "primary_current": RGBColor(0x00, 0x00, 0xC9),
        "primary_prior": RGBColor(0x60, 0x2A, 0x81),
        "competitor_current": RGBColor(0x00, 0x95, 0xFF),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Vaccines/Maternal",
        "_observed_palette": ["0000C9", "0095FF", "602A81", "F49C34", "67BB6E", "43964A", "B9B9B9", "D95776"],
    },

    # --- LYNPARZA (4 decks, Oncology) ---
    "LYNPARZA": {
        "client": "AZN",
        "primary_current": RGBColor(0x00, 0x66, 0x80),
        "primary_prior": RGBColor(0x33, 0xCC, 0xCC),
        "competitor_current": RGBColor(0xB2, 0xD2, 0x34),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology",
        "_observed_palette": ["006680", "B2D234", "33CCCC", "D9D9D9", "675F53", "92D050", "94D448", "FFBA00"],
    },

    # --- BONE_HCP_TRACKER (3 decks, Bone) ---
    "BONE_HCP_TRACKER": {
        "client": "AMGEN",
        "primary_current": RGBColor(0xFF, 0x99, 0x33),
        "primary_prior": RGBColor(0x58, 0x19, 0x3D),
        "competitor_current": RGBColor(0x00, 0x92, 0x01),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Bone",
        "_observed_palette": ["FF9933", "009201", "58193D", "ADADAD", "12437E", "0070C0", "FFC000", "00B0F0"],
    },

    # --- CALQUENCE (3 decks, Oncology/CLL) ---
    "CALQUENCE": {
        "client": "AZN",
        "primary_current": RGBColor(0x00, 0x20, 0x60),
        "primary_prior": RGBColor(0x00, 0xA2, 0xE0),
        "competitor_current": RGBColor(0xEE, 0x3E, 0xDD),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/CLL",
        "_observed_palette": ["002060", "EE3EDD", "00A2E0", "DC4405", "FFC72C", "E44405", "001F58", "E97142"],
    },

    # --- OJJAARA (3 decks, Hematology/MF) ---
    "OJJAARA": {
        "client": "GSK",
        "primary_current": RGBColor(0xEE, 0x53, 0x40),
        "primary_prior": RGBColor(0x60, 0x14, 0xFF),
        "competitor_current": RGBColor(0x6B, 0xB2, 0x3E),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Hematology/MF",
        "_observed_palette": ["EE5340", "6BB23E", "6014FF", "6CB33E", "6114FF", "92D050", "DFD0FF", "1DC2E7"],
    },

    # --- LIBTAYO (3 decks, Oncology/NMSC) ---
    "LIBTAYO": {
        "client": "REGENERON",
        "primary_current": RGBColor(0xDC, 0x00, 0x77),
        "primary_prior": RGBColor(0xC8, 0x33, 0x33),
        "competitor_current": RGBColor(0x00, 0x4F, 0x6F),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/NMSC",
        "_observed_palette": ["DC0077", "004F6F", "C83333", "338533", "96BBBB", "C19875", "D34D4D", "DE0D7E"],
    },

    # --- PHYSICIANS_SFE_PET (2 decks, Cross-brand) ---
    "PHYSICIANS_SFE_PET": {
        "client": "ALNYLAM",
        "primary_current": RGBColor(0xE7, 0x00, 0x4C),
        "primary_prior": RGBColor(0x6B, 0xC9, 0xEF),
        "competitor_current": RGBColor(0xFF, 0x69, 0x69),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Cross-brand",
        "_observed_palette": ["E7004C", "FF6969", "6BC9EF", "006838", "00607C", "00A44A", "00CC00", "D9D9D9"],
    },

    # --- TEZSPIRE (2 decks, Respiratory) ---
    "TEZSPIRE": {
        "client": "AZN",
        "primary_current": RGBColor(0xF8, 0x9C, 0x0D),
        "primary_prior": RGBColor(0x6C, 0x51, 0x99),
        "competitor_current": RGBColor(0xFF, 0xC6, 0x00),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Respiratory",
        "_observed_palette": ["F89C0D", "FFC600", "6C5199", "EF426F", "9B37FF", "0063C3", "C688E8", "5D9D39"],
    },

    # --- TEPEZZA (2 decks, Endocrinology/TED) ---
    "TEPEZZA": {
        "client": "JJ",
        "primary_current": RGBColor(0x7F, 0xB1, 0xE1),
        "primary_prior": RGBColor(0x37, 0xCF, 0xCA),
        "competitor_current": RGBColor(0x00, 0x63, 0xC3),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Endocrinology/TED",
        "_observed_palette": ["7FB1E1", "0063C3", "37CFCA", "228D8A", "A0D0FF", "CBF3F1", "00A3C4", "ADADAD"],
    },

    # --- DATROWAY (2 decks, Oncology/NSCLC) ---
    "DATROWAY": {
        "client": "DSI",
        "primary_current": RGBColor(0x1E, 0x22, 0xAA),
        "primary_prior": RGBColor(0x00, 0xAB, 0xC7),
        "competitor_current": RGBColor(0xEE, 0x76, 0x23),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/NSCLC",
        "_observed_palette": ["1E22AA", "EE7623", "00ABC7", "018EAF", "F5AD7B", "664993", "250E62", "B0E4FF"],
    },

    # --- ONIVYDE (2 decks, Oncology/Pancreatic) ---
    "ONIVYDE": {
        "client": "IPSEN",
        "primary_current": RGBColor(0x54, 0xAC, 0x65),
        "primary_prior": RGBColor(0x38, 0x73, 0x43),
        "competitor_current": RGBColor(0xC8, 0x48, 0x74),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/Pancreatic",
        "_observed_palette": ["54AC65", "C84874", "387343", "FA9F1A", "C2EAFF", "8EC899", "BBDEC2", "D2E9D6"],
    },

    # --- ADBRY (2 decks, Dermatology/AD) ---
    "ADBRY": {
        "client": "LEO",
        "primary_current": RGBColor(0xD5, 0x68, 0x5F),
        "primary_prior": RGBColor(0x00, 0x9B, 0x77),
        "competitor_current": RGBColor(0x6F, 0x43, 0x9A),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Dermatology/AD",
        "_observed_palette": ["D5685F", "6F439A", "009B77", "C014A3", "FED006", "EAAFE0", "90117A", "D561C1"],
    },

    # --- CCA_PP_TRACKER (1 decks, Cross-brand) ---
    "CCA_PP_TRACKER": {
        "client": "CCA",
        "primary_current": RGBColor(0xF2, 0x8E, 0x2B),
        "primary_prior": RGBColor(0xBD, 0xD7, 0xEE),
        "competitor_current": RGBColor(0x8F, 0xAA, 0xDC),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Cross-brand",
        "_observed_palette": ["F28E2B", "8FAADC", "BDD7EE", "AFED5D", "FFABD5", "FF6700", "FF8F43", "FFBC8F"],
    },

    # --- ULTOMIRIS (1 decks, Hematology/PNH) ---
    "ULTOMIRIS": {
        "client": "ALEXION",
        "primary_current": RGBColor(0x0E, 0x87, 0x79),
        "primary_prior": RGBColor(0xE3, 0x51, 0x05),
        "competitor_current": RGBColor(0x70, 0x30, 0xA0),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Hematology/PNH",
        "_observed_palette": ["0E8779", "7030A0", "E35105", "799A01", "FFA300", "FF8181", "84CA68", "009886"],
    },

    # --- EMPAVELI (1 decks, Hematology) ---
    "EMPAVELI": {
        "client": "APELLIS",
        "primary_current": RGBColor(0xFC, 0x3B, 0x6E),
        "primary_prior": RGBColor(0x81, 0xE3, 0xE3),
        "competitor_current": RGBColor(0x30, 0xCF, 0xD0),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Hematology",
        "_observed_palette": ["FC3B6E", "30CFD0", "81E3E3", "5A4986", "FD7FA0", "FEB0C4", "FFC000", "163A6B"],
    },

    # --- TRUQAP (1 decks, Oncology/Breast) ---
    "TRUQAP": {
        "client": "AZN",
        "primary_current": RGBColor(0x25, 0x0E, 0x62),
        "primary_prior": RGBColor(0x00, 0x69, 0x37),
        "competitor_current": RGBColor(0x65, 0x66, 0x81),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/Breast",
        "_observed_palette": ["250E62", "656681", "006937", "FFCB05", "F5A899", "05CDCD", "006680", "664993"],
    },

    # --- LOKELMA (1 decks, Nephrology) ---
    "LOKELMA": {
        "client": "AZN",
        "primary_current": RGBColor(0x00, 0x94, 0x7F),
        "primary_prior": RGBColor(0x4A, 0x2C, 0x8C),
        "competitor_current": RGBColor(0xB2, 0xB2, 0xB2),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Nephrology",
        "_observed_palette": ["00947F", "B2B2B2", "4A2C8C", "AB162C", "00B0F0", "FF2916", "003965", "00EAC9"],
    },

    # --- MIEBO (1 decks, Ophthalmology/DED) ---
    "MIEBO": {
        "client": "BL",
        "primary_current": RGBColor(0x00, 0xA9, 0xEB),
        "primary_prior": RGBColor(0x22, 0x2A, 0x35),
        "competitor_current": RGBColor(0x4B, 0x34, 0xA2),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Ophthalmology/DED",
        "_observed_palette": ["00A9EB", "4B34A2", "222A35", "400286", "D9D9D9", "FFC000", "003865", "993366"],
    },

    # --- BAVENCIO (1 decks, Oncology) ---
    "BAVENCIO": {
        "client": "EMD",
        "primary_current": RGBColor(0x80, 0x80, 0x80),
        "primary_prior": RGBColor(0xCC, 0xCC, 0xCC),
        "competitor_current": RGBColor(0xA6, 0xA6, 0xA6),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology",
        "_observed_palette": [],
    },

    # --- BLENREP (1 decks, Oncology/MM) ---
    "BLENREP": {
        "client": "GSK",
        "primary_current": RGBColor(0x3A, 0xB5, 0x1D),
        "primary_prior": RGBColor(0x66, 0x58, 0xA6),
        "competitor_current": RGBColor(0x5C, 0x10, 0x3B),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/MM",
        "_observed_palette": ["3AB51D", "5C103B", "6658A6", "7030A0", "668EDD", "21B6B2", "F25E29", "183978"],
    },

    # --- JEMPERLI-ZEJULA (1 decks, Oncology/Endometrial) ---
    "JEMPERLI-ZEJULA": {
        "client": "GSK",
        "primary_current": RGBColor(0x00, 0x84, 0x7C),
        "primary_prior": RGBColor(0xFF, 0xC0, 0x00),
        "competitor_current": RGBColor(0xD5, 0x19, 0x00),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/Endometrial",
        "_observed_palette": ["00847C", "D51900", "FFC000", "4F8734", "C69900", "668EDD", "706352", "A9A197"],
    },

    # --- ABILIFY_MAINTENA (1 decks, Psychiatry) ---
    "ABILIFY_MAINTENA": {
        "client": "OTSUKA",
        "primary_current": RGBColor(0xFF, 0xC0, 0x00),
        "primary_prior": RGBColor(0x3B, 0x59, 0x98),
        "competitor_current": RGBColor(0x7F, 0x47, 0xAA),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Psychiatry",
        "_observed_palette": ["FFC000", "7F47AA", "3B5998", "7D44A9", "7E97CD", "2D5CA2", "B3C9EA", "0000FF"],
    },

    # --- DUPIXENT_EoE (1 decks, Immunology/EoE) ---
    "DUPIXENT_EoE": {
        "client": "REGENERON_SANOFI",
        "primary_current": RGBColor(0x80, 0x80, 0x80),
        "primary_prior": RGBColor(0xCC, 0xCC, 0xCC),
        "competitor_current": RGBColor(0xA6, 0xA6, 0xA6),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Immunology/EoE",
        "_observed_palette": [],
    },

}


def get_brand(brand_key: str) -> dict:
    """Look up brand by key. Falls back to CLIENT{} for client-level lookup."""
    key = brand_key.upper().replace(' ', '_').replace('-', '_')
    if key in BRAND:
        return BRAND[key]
    if key in CLIENT:
        return CLIENT[key]
    raise KeyError(
        f"Unknown brand {brand_key!r}. Available brands: {sorted(BRAND.keys())}"
    )


def get_competitor(brand_key: str) -> str | None:
    """Return the competitor brand name for a given brand, or None."""
    b = get_brand(brand_key)
    return b.get("competitor_name")
