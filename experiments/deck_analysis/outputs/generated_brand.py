"""
Client-level brand defaults — fonts, heading colors, observed palettes.

Generated from deck analysis of 905 decks across 96 clients.
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
    # --- Amgen (AMG) (76 decks) ---
    "AMGEN": {
        "heading_color": RGBColor(0x00, 0x3C, 0x71),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 76,
        "_observed_palette": ["0063C3", "D6D2D0", "76C269", "1F497D", "92D050", "E57F7F", "813F97", "00B050"],
    },

    # --- daiichi sankyo (55 decks) ---
    "DAIICHI_SANKYO": {
        "heading_color": RGBColor(0x40, 0x40, 0x40),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 55,
        "_observed_palette": ["FF8813", "D9D9D9", "7F7F7F", "EE7623", "FFBFFF", "33A4FF", "34A355", "8F66A9"],
    },

    # --- AbbVie (ABV) (46 decks) ---
    "ABBVIE": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 46,
        "_observed_palette": ["99CC00", "007B80", "2091AC", "00B2AD", "96D801", "D73749", "FFD100", "FFC000"],
    },

    # --- Gilead (GLD) (41 decks) ---
    "GILEAD": {
        "heading_color": RGBColor(0x54, 0x56, 0x5B),
        "font_heading": "Trebuchet MS",
        "font_body": "Trebuchet MS",
        "template_path": None,
        "deck_count": 41,
        "_observed_palette": ["E7751F", "D11241", "00B050", "A6A6A6", "002060", "929A92", "C61A1A", "D86161"],
    },

    # --- AstraZeneca (AZN) (40 decks) ---
    "ASTRAZENECA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 40,
        "_observed_palette": ["EE7623", "00B050", "BFBFBF", "1E22AA", "FF0000", "0070C0", "643466", "A6A6A6"],
    },

    # --- Deciphera (35 decks) ---
    "DECIPHERA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 35,
        "_observed_palette": ["5D2759", "EB8125", "C72750", "F1B914", "5DBEBF", "008A00", "C00000", "92D050"],
    },

    # --- Merck (MER) (33 decks) ---
    "MERCK": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 33,
        "_observed_palette": ["801851", "BFBFBF", "00B050", "E886BC", "228848", "0590A8", "004D74", "A4C4EE"],
    },

    # --- Sanofi - Genzyme (SAN) (29 decks) ---
    "SANOFI___GENZYME": {
        "heading_color": RGBColor(0x23, 0x00, 0x4C),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 29,
        "_observed_palette": ["219491", "BFBFBF", "00B050", "005498", "D9D9D9", "793F93", "C00000", "249592"],
    },

    # --- Alnylam (ALN) (26 decks) ---
    "ALNYLAM": {
        "heading_color": RGBColor(0x00, 0xA3, 0xDC),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 26,
        "_observed_palette": ["E7004C", "00ACB3", "FF6969", "002060", "003866", "00607C", "33B66E", "C00000"],
    },

    # --- Novartis (NVS) (26 decks) ---
    "NOVARTIS": {
        "heading_color": RGBColor(0x1F, 0x49, 0x7D),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 26,
        "_observed_palette": ["018E86", "852065", "0C68B0", "92D050", "830F66", "FFC000", "7C1C33", "0000C9"],
    },

    # --- Foundation Medicine (24 decks) ---
    "FOUNDATION_MEDICINE": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "GT Sectra Book",
        "font_body": "GT Sectra Book",
        "template_path": None,
        "deck_count": 24,
        "_observed_palette": ["435363", "A0D8B3", "267270", "38ACA8", "00B0F0", "FF9466", "64CCC9", "378551"],
    },

    # --- Agios _ Servier (21 decks) ---
    "AGIOS___SERVIER": {
        "heading_color": RGBColor(0x24, 0x22, 0x69),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 21,
        "_observed_palette": ["D9D9D9", "BFBFBF", "242269", "A4C4EE", "1A355D", "A6A6A6", "242168", "D3DEF2"],
    },

    # --- BridgeBio (18 decks) ---
    "BRIDGEBIO": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 18,
        "_observed_palette": ["00B050", "507892", "92D050", "DFFBFC", "C2DFE3", "5B2478", "AF1634", "F30050"],
    },

    # --- Genentech (GNE) (17 decks) ---
    "GENENTECH": {
        "heading_color": RGBColor(0x3F, 0x3F, 0x3F),
        "font_heading": "Gene Sans",
        "font_body": "Gene Sans",
        "template_path": None,
        "deck_count": 17,
        "_observed_palette": ["C00000", "0E326F", "10529D", "7B4E96", "9ECEEB", "00ADA6", "CCCDCD", "999B9B"],
    },

    # --- Jazz (17 decks) ---
    "JAZZ": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 17,
        "_observed_palette": ["0070C0", "8497B0", "D6EBD4", "FFC000", "FF69D6", "9A0072", "00B050", "7F7F7F"],
    },

    # --- Seagen (Pfizer Onc) (17 decks) ---
    "SEAGEN": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 17,
        "_observed_palette": ["BFBFBF", "92D050", "8275E9", "E36767", "ADADAD", "595959", "2AA4FF", "F7C99B"],
    },

    # --- BioMarin (16 decks) ---
    "BIOMARIN": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 16,
        "_observed_palette": ["FBD401", "28509C", "091A79", "ED1849", "ED037C", "A9208E", "081A79", "B60018"],
    },

    # --- Blueprint Medicines (BPM) (16 decks) ---
    "BLUEPRINT_MEDICINES": {
        "heading_color": RGBColor(0x00, 0x26, 0x3D),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 16,
        "_observed_palette": ["BBB5AF", "638326", "1E5271", "C2CA7E", "67ADD7", "00B050", "FF0000", "C6BFB6"],
    },

    # --- Ipsen (14 decks) ---
    "IPSEN": {
        "heading_color": RGBColor(0x00, 0x0E, 0x56),
        "font_heading": "Lato",
        "font_body": "Lato",
        "template_path": None,
        "deck_count": 14,
        "_observed_palette": ["00B050", "54AC65", "461C6F", "D38A4F", "C84874", "C00000", "BFBFBF", "0052B2"],
    },

    # --- Pfizer (14 decks) ---
    "PFIZER": {
        "heading_color": RGBColor(0x00, 0x63, 0xC3),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 14,
        "_observed_palette": ["0000C9", "602A81", "B3C9EA", "FFFFFF", "05BE66", "43964A", "6B0834", "E9EED9"],
    },

    # --- Takeda (TAK) (14 decks) ---
    "TAKEDA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 14,
        "_observed_palette": ["00B050", "BFBFBF", "00823B", "C00000", "731013", "E1242A", "00708A", "0EB2CD"],
    },

    # --- ArgenX (AGX) (12 decks) ---
    "ARGENX": {
        "heading_color": RGBColor(0x00, 0x2B, 0x48),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 12,
        "_observed_palette": ["002B48", "9137AF", "1F5FA0", "FF9300", "8E3F9A", "70D9D3", "780032", "00223C"],
    },

    # --- BMS (12 decks) ---
    "BMS": {
        "heading_color": RGBColor(0x59, 0x54, 0x54),
        "font_heading": "Trebuchet MS",
        "font_body": "Trebuchet MS",
        "template_path": None,
        "deck_count": 12,
        "_observed_palette": ["BFBFBF", "00B050", "0E2D75", "FB796E", "780081", "002140", "FFD186", "7F7F7F"],
    },

    # --- Sobi Pharma (SOB) (12 decks) ---
    "SOBI_PHARMA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 12,
        "_observed_palette": ["2D5E77", "F15A22", "4CA1A6", "7E82FF", "00B050", "003B5C", "F79C7A", "ABBFC9"],
    },

    # --- UCB (12 decks) ---
    "UCB": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Tahoma",
        "font_body": "Tahoma",
        "template_path": None,
        "deck_count": 12,
        "_observed_palette": ["7030A0", "92D050", "00535E", "008192", "646D6D", "A6A6A6", "C198E0", "6ECEB2"],
    },

    # --- GSK (11 decks) ---
    "GSK": {
        "heading_color": RGBColor(0xF3, 0x66, 0x33),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 11,
        "_observed_palette": ["00B050", "A6A6A6", "EE5340", "C00000", "0B2745", "D51900", "D0400C", "6014FF"],
    },

    # --- LEO Pharma (10 decks) ---
    "LEO_PHARMA": {
        "heading_color": RGBColor(0x20, 0x28, 0x2F),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 10,
        "_observed_palette": ["FF5067", "5387A6", "009B77", "C014A3", "D5685F", "9AE649", "00B050", "C017A2"],
    },

    # --- Phathom Pharmaceuticals (PHP) (10 decks) ---
    "PHATHOM_PHARMACEUTICALS": {
        "heading_color": RGBColor(0x00, 0x43, 0x70),
        "font_heading": "Arial(Body)",
        "font_body": "Arial(Body)",
        "template_path": None,
        "deck_count": 10,
        "_observed_palette": ["004370", "C5344A", "AEDFFB", "9A9B9D", "76B1E0", "9C9534", "8B8B8B", "60AB3F"],
    },

    # --- Arcutis (ARC) (9 decks) ---
    "ARCUTIS": {
        "heading_color": RGBColor(0xE0, 0xB4, 0x1C),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 9,
        "_observed_palette": ["FFC000", "229719", "7F7F7F", "D34C4C", "CB35DB", "B4C7E7", "DE5754", "8271E5"],
    },

    # --- Alexion Pharmaceuticals (8 decks) ---
    "ALEXION_PHARMACEUTICALS": {
        "heading_color": RGBColor(0x00, 0x1E, 0x60),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["009886", "415C79", "EC9DB8", "76C23A", "BC3666", "59B7C7", "8B2DAB", "DA4305"],
    },

    # --- Genmab (8 decks) ---
    "GENMAB": {
        "heading_color": RGBColor(0xFF, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["002060", "007935", "97C463", "F9B235", "E6AF00", "3B1A53", "00937A", "4E2B19"],
    },

    # --- ITF Therapeutics (8 decks) ---
    "ITF_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["0E8B37", "A6A6A6", "840B55", "004F9E", "2CB6B3", "0E8D38", "A7C978", "1F0064"],
    },

    # --- Natera (NAT) (8 decks) ---
    "NATERA": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["44C04F", "9E0000", "0070F2", "FF494B", "FA18FA", "A973FD", "008A00", "DF7F7F"],
    },

    # --- Neurocrine (8 decks) ---
    "NEUROCRINE": {
        "heading_color": RGBColor(0x2B, 0x5F, 0x51),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 8,
        "_observed_palette": ["439F5B", "D34D4D", "A6A6A6", "ED9C97", "B6465F", "B092B4", "ED7D31", "F2BAB6"],
    },

    # --- Celltrion (7 decks) ---
    "CELLTRION": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["00B050", "FF0000", "92D050", "7BD7D1", "0064A0", "FFC000", "AE178D", "532F88"],
    },

    # --- Corcept (7 decks) ---
    "CORCEPT": {
        "heading_color": RGBColor(0x23, 0x00, 0x4C),
        "font_heading": "Lato",
        "font_body": "Lato",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": ["28378E", "2A3C97", "EE7623", "1E40BE", "5BC67C", "43248C", "C00000", "FF9393"],
    },

    # --- Ionis (7 decks) ---
    "IONIS": {
        "heading_color": RGBColor(0xFF, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 7,
        "_observed_palette": [],
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

    # --- Alkermes (ALK) (6 decks) ---
    "ALKERMES": {
        "heading_color": RGBColor(0x01, 0x56, 0x8C),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 6,
        "_observed_palette": ["95054B", "088DD0", "00B050", "E65D02", "3C7E79", "BF85B6", "7030A0", "6096C8"],
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

    # --- Eisai (6 decks) ---
    "EISAI": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 6,
        "_observed_palette": ["880043", "7C164A", "FFA700", "FCBA3D", "006600", "A6A6A6", "FFC000", "B7849E"],
    },

    # --- Theravance (THV) (6 decks) ---
    "THERAVANCE": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 6,
        "_observed_palette": ["005EB8", "C00000", "002060", "3C7623", "D9900D", "EEDE12", "B2B5B6", "509E2F"],
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

    # --- EMD Serono (5 decks) ---
    "EMD_SERONO": {
        "heading_color": RGBColor(0x65, 0x18, 0x5A),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": ["503291", "672A76", "7F7F7F", "65185A", "43964A", "BA2A4C", "FF5500", "4EA9F0"],
    },

    # --- Horizon (Viela Bio) (5 decks) ---
    "HORIZON": {
        "heading_color": RGBColor(0x40, 0x47, 0x4E),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": ["D6D2D0", "76C269", "E57F7F", "0063C3", "5B729B", "F2BFBF", "2F6D5B", "92D050"],
    },

    # --- Novo Nordisk (5 decks) ---
    "NOVO_NORDISK": {
        "heading_color": RGBColor(0x00, 0x19, 0x65),
        "font_heading": "Apis For Office",
        "font_body": "Apis For Office",
        "template_path": None,
        "deck_count": 5,
        "_observed_palette": ["795DED", "EA4970", "001965", "D66E28", "8DD6F7", "219491", "2D79C1", "FFC000"],
    },

    # --- Aerovate Therapeutics (AOT) (4 decks) ---
    "AEROVATE_THERAPEUTICS": {
        "heading_color": RGBColor(0x59, 0x59, 0x59),
        "font_heading": "Arial(Body)",
        "font_body": "Arial(Body)",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["08708A", "DAC29E", "B9285C", "00B050", "A6A6A6", "FF0000", "76A6C6", "98BCD4"],
    },

    # --- Apellis (APL) (4 decks) ---
    "APELLIS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["FC3B6E", "3E797E", "E7E6E6", "30CFD0", "5B9BD5", "7F7F7F", "81E3E3", "5A4986"],
    },

    # --- Astellas (4 decks) ---
    "ASTELLAS": {
        "heading_color": RGBColor(0x00, 0x20, 0x60),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["7030A0", "CB73DF", "73ABDD", "960048", "0070C0", "00BED5", "E6882B", "92D050"],
    },

    # --- AVEO Oncology (4 decks) ---
    "AVEO_ONCOLOGY": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["395F79", "D6037D", "A6103D", "6F1A44", "333F8D", "6096C8", "99B3C1", "4D7A93"],
    },

    # --- Bayer (BAY) (4 decks) ---
    "BAYER": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Arial (Body)",
        "font_body": "Arial (Body)",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["54306A", "06AACD", "D21D00", "1F4288", "002060", "42A2FF", "F4B183", "92D050"],
    },

    # --- Braeburn (4 decks) ---
    "BRAEBURN": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["3D66E3", "B80D48", "29C2A1", "01437E", "BFBFBF", "701C6A", "F29724", "6D6D6D"],
    },

    # --- Coherus BioSciences (4 decks) ---
    "COHERUS_BIOSCIENCES": {
        "heading_color": RGBColor(0x7F, 0x7F, 0x7F),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["053C69", "A5A5A5", "00857C", "D39092", "F26322", "01B1D1", "A6A6A6", "00A8DF"],
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

    # --- Kyowa Kyrin (KKN) (4 decks) ---
    "KYOWA_KYRIN": {
        "heading_color": RGBColor(0xBF, 0xBF, 0xBF),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 4,
        "_observed_palette": ["191C51", "009866", "F74C4F", "4A2C8C", "A8AAA5", "009AC7", "BFBFBF", "0090B8"],
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

    # --- Arvinas (ARVA) (3 decks) ---
    "ARVINAS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["DF7F7F", "4166A1", "F37116", "BFBFBF", "F1CE63", "59A14F", "4CAD4C", "499894"],
    },

    # --- Cytokinetics (CYTO) (3 decks) ---
    "CYTOKINETICS": {
        "heading_color": RGBColor(0x75, 0xA7, 0x33),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["0A367D", "BE215E", "D9D9D9", "75A733", "859BBE", "125FDC", "80ACF4", "CADCFA"],
    },

    # --- Inhibikase Therapeutics (3 decks) ---
    "INHIBIKASE_THERAPEUTICS": {
        "heading_color": RGBColor(0x7F, 0x7F, 0x7F),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["303A8D", "FBE5D6", "F8CBAD", "F4B183", "C00000", "233616", "A97755", "C55A11"],
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

    # --- TG Therapeutics (3 decks) ---
    "TG_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Crimson Pro ExtraBold",
        "font_body": "Crimson Pro ExtraBold",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": ["00B050", "0E326F", "FF9393", "92D050", "85C226", "40BB9A", "83D977", "6990E7"],
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

    # --- ViiV (VII) (3 decks) ---
    "VIIV": {
        "heading_color": RGBColor(0xFF, 0x00, 0x80),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 3,
        "_observed_palette": [],
    },

    # --- Bausch (BAU) (2 decks) ---
    "BAUSCH": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Century Gothic",
        "font_body": "Century Gothic",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["4B34A2", "00A9EB", "00B050", "222A35", "400286", "A6A6A6", "D9D9D9", "FFC000"],
    },

    # --- Client Access - Sales Sim (2 decks) ---
    "CLIENT_ACCESS___SALES_SIM": {
        "heading_color": RGBColor(0x00, 0x9D, 0x94),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["2D3E85"],
    },

    # --- Denali Therapeutics (2 decks) ---
    "DENALI_THERAPEUTICS": {
        "heading_color": RGBColor(0x40, 0x40, 0x40),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["9F76C0", "415C77", "A2C745", "FFC000", "41A55E", "494949", "FF8C5E", "92D050"],
    },

    # --- PTC Therapeutics (PTC) (2 decks) ---
    "PTC_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["804080", "48BC6F", "BFBFBF", "FF7F7F", "224468", "346323", "43964A", "23466C"],
    },

    # --- Reata Pharmaceuticals (2 decks) ---
    "REATA_PHARMACEUTICALS": {
        "heading_color": RGBColor(0xBD, 0x28, 0x5E),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["7ABFBD", "72AAB3", "FFE3A3", "83C864", "FFAA34", "62B1C4", "219491", "6FBF4A"],
    },

    # --- Roche (2 decks) ---
    "ROCHE": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["10529D", "884199", "25B3BE", "ECAE8C", "BCCB01", "B41530", "F07D91", "F9BF7F"],
    },

    # --- Sarepta Therapeutics (2 decks) ---
    "SAREPTA_THERAPEUTICS": {
        "heading_color": RGBColor(0x66, 0x1B, 0x62),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["A6A6A6", "54004B", "E7A90A", "04197F", "661B62", "7D9D57", "FFC000"],
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

    # --- Vertex (2 decks) ---
    "VERTEX": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 2,
        "_observed_palette": ["C366A9", "3B7ADA", "EC002F", "FFC000", "7F7F7F", "92D050", "06BED6", "B6B6B6"],
    },

    # --- 01 Nirogacestat PET and HCP-Pt (1 deck) ---
    "01_NIROGACESTAT_PET_AND_HCP_PT": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- Alumis (1 deck) ---
    "ALUMIS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["002060", "0074BB", "6969FF", "FF742F", "2A5046", "66AF9C", "2FB0FF", "FFA74F"],
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

    # --- BeiGene (1 deck) ---
    "BEIGENE": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Poppins",
        "font_body": "Poppins",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["FFC000", "5DBA09", "92D050", "B0E2FF", "6C130E", "DC4405", "C00000", "27A7DB"],
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

    # --- Boston Scientific (1 deck) ---
    "BOSTON_SCIENTIFIC": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "GT Sectra Book",
        "font_body": "GT Sectra Book",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["3881C4", "C00000", "93C0E1", "83C937", "002D55", "118FFF", "60B5FF", "B0DAFF"],
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

    # --- FDA Data Extraction Exercise (1 deck) ---
    "FDA_DATA_EXTRACTION_EXERCISE": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Calibri",
        "font_body": "Calibri",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["0082BA", "FFC000", "7F7F7F", "6BBBAE", "004E70"],
    },

    # --- G1 Therapeutics (1 deck) ---
    "G1_THERAPEUTICS": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Arial",
        "font_body": "Arial",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
    },

    # --- Gossamerbio (1 deck) ---
    "GOSSAMERBIO": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
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

    # --- LIvaNova (1 deck) ---
    "LIVANOVA": {
        "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
        "font_heading": "Manrope",
        "font_body": "Manrope",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
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

    # --- Puma Biotechnology (1 deck) ---
    "PUMA_BIOTECHNOLOGY": {
        "heading_color": RGBColor(0x14, 0x31, 0x55),
        "font_heading": "Arial(Body)",
        "font_body": "Arial(Body)",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["0083BF", "FFC000", "48A7AA", "00833C", "7F7F7F", "7F53C1", "203864", "70AD47"],
    },

    # --- Sage Therapeutics (1 deck) ---
    "SAGE_THERAPEUTICS": {
        "heading_color": RGBColor(0x1C, 0x51, 0x4C),
        "font_heading": "Roboto",
        "font_body": "Roboto",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": ["FF6E14", "93A7A1", "1C514C", "7F7F7F"],
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

    # --- Verastem Oncology (1 deck) ---
    "VERASTEM_ONCOLOGY": {
        "heading_color": RGBColor(0xEE, 0xEC, 0xE1),
        "font_heading": "Roboto Light",
        "font_body": "Roboto Light",
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

    # --- Vir Biotech (1 deck) ---
    "VIR_BIOTECH": {
        "heading_color": RGBColor(0x00, 0x00, 0x00),
        "font_heading": "Lato Semibold",
        "font_body": "Lato Semibold",
        "template_path": None,
        "deck_count": 1,
        "_observed_palette": [],
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
    # --- UNMAPPED (785 decks, Unknown) ---
    "UNMAPPED": {
        "client": "VIR_BIOTECH",
        "primary_current": RGBColor(0xFF, 0xC0, 0x00),
        "primary_prior": RGBColor(0xD9, 0xD9, 0xD9),
        "competitor_current": RGBColor(0x92, 0xD0, 0x50),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Unknown",
        "_observed_palette": ["FFC000", "92D050", "D9D9D9", "7030A0", "00B0F0", "0070C0", "002060", "0063C3"],
    },

    # --- ENHERTU (32 decks, Oncology/HER2) ---
    "ENHERTU": {
        "client": "DSI_AZN",
        "primary_current": RGBColor(0xFF, 0x88, 0x13),
        "primary_prior": RGBColor(0x34, 0xA3, 0x55),
        "competitor_current": RGBColor(0xEE, 0x76, 0x23),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/HER2",
        "_observed_palette": ["FF8813", "EE7623", "34A355", "33A4FF", "0070C0", "8F66A9", "F5AD7B", "D9D9D9"],
    },

    # --- UPLIZNA (13 decks, Neurology/NMOSD) ---
    "UPLIZNA": {
        "client": "AMGEN",
        "primary_current": RGBColor(0xD6, 0xD2, 0xD0),
        "primary_prior": RGBColor(0x5B, 0x72, 0x9B),
        "competitor_current": RGBColor(0x76, 0xC2, 0x69),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Neurology/NMOSD",
        "_observed_palette": ["D6D2D0", "76C269", "5B729B", "E57F7F", "0063C3", "F2BFBF", "2F6D5B", "00BCE4"],
    },

    # --- OTEZLA (11 decks, Dermatology/Psoriasis) ---
    "OTEZLA": {
        "client": "AMGEN",
        "primary_current": RGBColor(0x1F, 0x49, 0x7D),
        "primary_prior": RGBColor(0xC4, 0x0E, 0x12),
        "competitor_current": RGBColor(0x00, 0xA3, 0xDF),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Dermatology/Psoriasis",
        "_observed_palette": ["1F497D", "00A3DF", "C40E12", "448D96", "009900", "BF9761", "F3C108", "AAA8AB"],
    },

    # --- LYNPARZA (6 decks, Oncology) ---
    "LYNPARZA": {
        "client": "AZN",
        "primary_current": RGBColor(0x92, 0xD0, 0x50),
        "primary_prior": RGBColor(0xD9, 0xD9, 0xD9),
        "competitor_current": RGBColor(0x00, 0x4D, 0x74),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology",
        "_observed_palette": ["92D050", "004D74", "D9D9D9", "94D448", "00B3FF", "006680", "B2D234", "33CCCC"],
    },

    # --- ABRYSVO (6 decks, Vaccines/Maternal) ---
    "ABRYSVO": {
        "client": "PFIZER",
        "primary_current": RGBColor(0x00, 0x00, 0xC9),
        "primary_prior": RGBColor(0xE3, 0x56, 0xDC),
        "competitor_current": RGBColor(0x00, 0x95, 0xFF),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Vaccines/Maternal",
        "_observed_palette": ["0000C9", "0095FF", "E356DC", "94D448", "BEF202", "0287A2", "004B70", "03075E"],
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

    # --- BONE_HCP_TRACKER (4 decks, Bone) ---
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

    # --- LIBTAYO (4 decks, Oncology/NMSC) ---
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

    # --- PHYSICIANS_SFE_PET (3 decks, Cross-brand) ---
    "PHYSICIANS_SFE_PET": {
        "client": "ALNYLAM",
        "primary_current": RGBColor(0xFF, 0x69, 0x69),
        "primary_prior": RGBColor(0x00, 0xCC, 0x00),
        "competitor_current": RGBColor(0x33, 0xB6, 0x6E),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Cross-brand",
        "_observed_palette": ["FF6969", "33B66E", "00CC00", "B2B2B2", "E7004C", "6BC9EF", "006838", "00607C"],
    },

    # --- TEPEZZA (3 decks, Endocrinology/TED) ---
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

    # --- TEZSPIRE (3 decks, Respiratory) ---
    "TEZSPIRE": {
        "client": "AZN",
        "primary_current": RGBColor(0x00, 0x63, 0xC3),
        "primary_prior": RGBColor(0xFF, 0xC6, 0x00),
        "competitor_current": RGBColor(0xF8, 0x9C, 0x0D),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Respiratory",
        "_observed_palette": ["0063C3", "F89C0D", "FFC600", "6C5199", "EF426F", "9B37FF", "C688E8", "5D9D39"],
    },

    # --- EMPAVELI (3 decks, Hematology) ---
    "EMPAVELI": {
        "client": "APELLIS",
        "primary_current": RGBColor(0xFC, 0x3B, 0x6E),
        "primary_prior": RGBColor(0x5B, 0x9B, 0xD5),
        "competitor_current": RGBColor(0x3E, 0x79, 0x7E),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Hematology",
        "_observed_palette": ["FC3B6E", "3E797E", "5B9BD5", "584D75", "433764", "8A8A8A", "66ADB4", "C9E2E5"],
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

    # --- ADBRY (3 decks, Dermatology/AD) ---
    "ADBRY": {
        "client": "LEO",
        "primary_current": RGBColor(0xD5, 0x68, 0x5F),
        "primary_prior": RGBColor(0x00, 0x9B, 0x77),
        "competitor_current": RGBColor(0x6F, 0x43, 0x9A),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Dermatology/AD",
        "_observed_palette": ["D5685F", "6F439A", "009B77", "C017A2", "C8DEED", "84A0B5", "D39AD0", "C014A3"],
    },

    # --- ULTOMIRIS (2 decks, Hematology/PNH) ---
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

    # --- TRUQAP (2 decks, Oncology/Breast) ---
    "TRUQAP": {
        "client": "AZN",
        "primary_current": RGBColor(0x25, 0x0E, 0x62),
        "primary_prior": RGBColor(0xFF, 0xCB, 0x05),
        "competitor_current": RGBColor(0x00, 0x69, 0x37),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Oncology/Breast",
        "_observed_palette": ["250E62", "006937", "FFCB05", "656681", "F5A899", "05CDCD", "006680", "664993"],
    },

    # --- MIEBO (2 decks, Ophthalmology/DED) ---
    "MIEBO": {
        "client": "BL",
        "primary_current": RGBColor(0x4B, 0x34, 0xA2),
        "primary_prior": RGBColor(0xFF, 0xC0, 0x00),
        "competitor_current": RGBColor(0xD9, 0xD9, 0xD9),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Ophthalmology/DED",
        "_observed_palette": ["4B34A2", "D9D9D9", "FFC000", "00A9EB", "300264", "A57394", "BED32E", "E4483C"],
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

    # --- DUPIXENT_EoE (2 decks, Immunology/EoE) ---
    "DUPIXENT_EoE": {
        "client": "REGENERON_SANOFI",
        "primary_current": RGBColor(0x21, 0x94, 0x91),
        "primary_prior": RGBColor(0xF7, 0x96, 0x46),
        "competitor_current": RGBColor(0xA7, 0xA7, 0xA7),
        "positive": POSITIVE_GREEN,
        "negative": NEGATIVE_RED,
        "therapy_area": "Immunology/EoE",
        "_observed_palette": ["219491", "A7A7A7", "F79646", "1BA2DA"],
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
