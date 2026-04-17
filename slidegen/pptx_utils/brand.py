"""
brand.py — Pharmaceutical brand definitions and slide constants.

BRAND{} is keyed by **brand** (one entry per pharmaceutical product), not by
client. A single client (e.g., J&J) typically has many brands (Rybrevant,
Tepezza, Darzalex, ...) each with distinct colors. Client-level defaults
(fonts, heading color, template path) live in CLIENT{} and are referenced by
each brand via its "client" field.

Brand entries were populated from real-deck analysis of 32 PET decks across
30 distinct brands. Series colors were extracted directly from chart XML
(<c:ser>/<c:spPr>/<a:solidFill>/<a:srgbClr>) with structural colors
(black, white, delta-green/red, text greys) filtered out, so the remaining
palette represents the brand's actual identity.

Lookup:
    from slidegen.pptx_utils.brand import get_brand
    b = get_brand("RYBREVANT")
    b["primary_current"]  # RGBColor orange — Rybrevant's brand color
    b["font_heading"]     # "Johnson Display" — inherited from CLIENT["JJ"]
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

# ── Universal colors observed across all 32 decks ────────────────────────────

POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)     # 494 occurrences
NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)       # 913 occurrences
NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)  #  61 occurrences

GREY_DARK = RGBColor(0x40, 0x40, 0x40)
GREY_MID = RGBColor(0x59, 0x59, 0x59)
GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)
GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)       # 229 occurrences — default table alt-row


# ============================================================================
# CLIENT{} — Client-level defaults (fonts, heading color, template path)
# ============================================================================

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
    "DSI_AZN": {          # Datroway / Enhertu co-commercialized DSI-AZN
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
    # ========================================================================
    # 551-deck grounding additions (Apr 2026) — 60+ pharma clients
    # ========================================================================
    # --- AbbVie (ABV) (28 decks) ---
    "ABV": {
    "heading_color": RGBColor(0x00, 0x00, 0x00),
    "font_heading": "Calibri",
    "font_body": "Calibri",
    "template_path": None,
    "deck_count": 28,
    "_observed_palette": ["99CC00", "007B80", "2091AC", "96D801", "D73749", "FFD100", "FFC000", "532F88"],
    },
    # --- Gilead (GLD) (22 decks) ---
    "GLD": {
    "heading_color": RGBColor(0x54, 0x56, 0x5B),
    "font_heading": "Trebuchet MS",
    "font_body": "Trebuchet MS",
    "template_path": None,
    "deck_count": 22,
    "_observed_palette": ["D11241", "A6A6A6", "E7751F", "002060", "00B050", "B9BBBF", "FF0000", "E86A6A"],
    },
    # --- Merck (MER) (19 decks) ---
    "MER": {
    "heading_color": RGBColor(0x00, 0x00, 0x00),
    "font_heading": "Arial",
    "font_body": "Arial",
    "template_path": None,
    "deck_count": 19,
    "_observed_palette": ["00B050", "801851", "E886BC", "014C35", "BFBFBF", "228848", "004D74", "00857C"],
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
    "AGIOS": {
    "heading_color": RGBColor(0x24, 0x22, 0x69),
    "font_heading": "Century Gothic",
    "font_body": "Century Gothic",
    "template_path": None,
    "deck_count": 16,
    "_observed_palette": ["242269", "D9D9D9", "BFBFBF", "A4C4EE", "1A355D", "A6A6A6", "242168", "D3DEF2"],
    },
    # --- Sanofi - Genzyme (SAN) (13 decks) ---
    "SAN": {
    "heading_color": RGBColor(0x7F, 0x7F, 0x7F),
    "font_heading": "Arial",
    "font_body": "Arial",
    "template_path": None,
    "deck_count": 13,
    "_observed_palette": ["00B050", "219491", "BFBFBF", "525CA3", "D9D9D9", "C00000", "249592", "2A3C97"],
    },
    # --- BioMarin (12 decks) ---
    "BMN": {
    "heading_color": RGBColor(0x00, 0x00, 0x00),
    "font_heading": "Arial ",
    "font_body": "Arial ",
    "template_path": None,
    "deck_count": 12,
    "_observed_palette": ["FBD401", "28509C", "091A79", "ED1849", "ED037C", "A9208E", "081A79", "B60018"],
    },
    # --- BridgeBio (12 decks) ---
    "BBIO": {
    "heading_color": RGBColor(0x00, 0x00, 0x00),
    "font_heading": "Arial",
    "font_body": "Arial",
    "template_path": None,
    "deck_count": 12,
    "_observed_palette": ["507892", "AF1634", "E62E4B", "2F6938", "5B2478", "DF7B17", "5AC2ED", "051D74"],
    },
    # --- Blueprint Medicines (BPM) (11 decks) ---
    "BPM": {
    "heading_color": RGBColor(0x00, 0x26, 0x3D),
    "font_heading": "Calibri",
    "font_body": "Calibri",
    "template_path": None,
    "deck_count": 11,
    "_observed_palette": ["1E5271", "BBB5AF", "638326", "C2CA7E", "67ADD7", "C6BFB6", "7030A0", "722E3E"],
    },
    # --- Genentech (GNE) (11 decks) ---
    "GNE": {
    "heading_color": RGBColor(0x00, 0x00, 0x00),
    "font_heading": "Gene Sans",
    "font_body": "Gene Sans",
    "template_path": None,
    "deck_count": 11,
    "_observed_palette": ["C00000", "0E326F", "7B4E96", "9ECEEB", "00ADA6", "CCCDCD", "999B9B", "0070C0"],
    },
    # --- Alnylam (ALN) (10 decks) ---
    "ALN": {
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
    "TAK": {
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
    # --- Seagen (Pfizer Onc) (9 decks) ---
    "SEAGEN_PFIZER_ONC": {
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
    "AGX": {
    "heading_color": RGBColor(0x00, 0x2B, 0x48),
    "font_heading": "Arial",
    "font_body": "Arial",
    "template_path": None,
    "deck_count": 8,
    "_observed_palette": ["9137AF", "002B48", "1F5FA0", "FF9300", "8E3F9A", "780032", "00223C", "B00075"],
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
    # --- Sobi Pharma (SOB) (7 decks) ---
    "SOB": {
    "heading_color": RGBColor(0xFF, 0xFF, 0xFF),
    "font_heading": "Calibri",
    "font_body": "Calibri",
    "template_path": None,
    "deck_count": 7,
    "_observed_palette": ["2D5E77", "003B5C", "F15A22", "00B050", "4CA1A6", "ABBFC9", "F79C7A", "6CB33E"],
    },
    # --- Arcutis (ARC) (6 decks) ---
    "ARC": {
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


# ============================================================================

# Long-form key aliases (map folder names to CLIENT keys)
CLIENT_ALIASES = {
    "ABBVIE": "ABV",
    "AGIOS___SERVIER": "AGIOS",
    "ALNYLAM": "ALN",
    "ARCUTIS": "ARC",
    "ARGENX": "AGX",
    "ASTRAZENECA": "AZN",
    "BIOMARIN": "BMN",
    "BLUEPRINT_MEDICINES": "BPM",
    "BRIDGEBIO": "BBIO",
    "DAIICHI_SANKYO": "DSI",
    "EMD_SERONO": "EMD",
    "GENENTECH": "GNE",
    "GILEAD": "GLD",
    "LEO_PHARMA": "LEO",
    "MERCK": "MER",
    "SANOFI___GENZYME": "SAN",
    "SEAGEN": "SEAGEN_PFIZER_ONC",
    "SOBI_PHARMA": "SOB",
    "TAKEDA": "TAK",
}

# BRAND{} — One entry per pharmaceutical brand
# ============================================================================
#
# Every brand has its own palette. Same client, different brands = different
# colors. Palettes were extracted from real chart XML across 32 PET decks and
# filtered to remove structural colors (black, white, delta greens/reds, text
# greys) so the remaining top colors represent brand identity.
#
# Fields:
#   client:           str — key into CLIENT{} for fonts, heading color
#   therapy_area:     str — helpful metadata for viz-selector decisions
#   primary_current:  RGBColor — current-wave bar/chart fill
#   primary_prior:    RGBColor — prior-wave tint
#   secondary:        RGBColor — accent / secondary comparison series
#   competitor_brand: str (optional) — key of competitor BRAND for automatic
#                     competitor color lookup
#   positive/negative: universal delta colors (imported)
#   _observed_palette: list[str] — top non-structural colors observed

BRAND = {
    # J&J portfolio
    "RYBREVANT": {
        "client":           "JJ",
        "therapy_area":     "Oncology/NSCLC",
        "primary_current":  RGBColor(0xF7, 0x58, 0x24),
        "primary_prior":    RGBColor(0xFB, 0xAB, 0x91),
        "secondary":        RGBColor(0xF9, 0x59, 0x24),
        "competitor_brand": "TAGRISSO",
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#F75824", "#FBAB91", "#F95924", "#7030A0", "#BAB0AC", "#1F698F", "#77BEE3", "#7F46AA"],
    },
    "TEPEZZA": {
        "client":           "JJ",
        "therapy_area":     "Endocrinology/TED",
        "primary_current":  RGBColor(0x7F, 0xB1, 0xE1),
        "primary_prior":    RGBColor(0x37, 0xCF, 0xCA),
        "secondary":        RGBColor(0x00, 0x63, 0xC3),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#7FB1E1", "#0063C3", "#37CFCA", "#228D8A", "#A0D0FF", "#CBF3F1", "#00A3C4", "#ADADAD"],
    },
    "DARZALEX": {
        "client":           "JJ",
        "therapy_area":     "Oncology/MM",
        "primary_current":  RGBColor(0x6F, 0xC6, 0xC1),
        "primary_prior":    RGBColor(0x2A, 0x41, 0xC1),
        "secondary":        RGBColor(0xBD, 0x05, 0xED),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#6FC6C1", "#BD05ED", "#2A41C1", "#FF6A5A", "#FFC000", "#0302C8", "#2D5BEF", "#AC4069"],
    },

    # AZN portfolio
    "CALQUENCE": {
        "client":           "AZN",
        "therapy_area":     "Oncology/CLL",
        "primary_current":  RGBColor(0x00, 0xA2, 0xE0),
        "primary_prior":    RGBColor(0xDC, 0x44, 0x05),
        "secondary":        RGBColor(0x00, 0x20, 0x60),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#00A2E0", "#002060", "#DC4405", "#FFC72C", "#EE3EDD", "#E44405", "#001F58", "#E97142"],
    },
    "LOKELMA": {
        "client":           "AZN",
        "therapy_area":     "Nephrology",
        "primary_current":  RGBColor(0x00, 0x94, 0x7F),
        "primary_prior":    RGBColor(0x4A, 0x2C, 0x8C),
        "secondary":        RGBColor(0x8F, 0x20, 0x64),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#00947F", "#8F2064", "#4A2C8C", "#00EAC9", "#B6FFF4", "#003965", "#00B0F0", "#D1C5ED"],
    },
    "LYNPARZA": {
        "client":           "AZN",
        "therapy_area":     "Oncology",
        "primary_current":  RGBColor(0x94, 0xD4, 0x48),
        "primary_prior":    RGBColor(0x00, 0xB3, 0xFF),
        "secondary":        RGBColor(0xFF, 0xBA, 0x00),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#94D448", "#FFBA00", "#00B3FF", "#34A355", "#D2CEC8", "#F74548", "#D5D5D5", "#9FD448"],
    },
    "TRUQAP": {
        "client":           "AZN",
        "therapy_area":     "Oncology/Breast",
        "primary_current":  RGBColor(0x25, 0x0E, 0x62),
        "primary_prior":    RGBColor(0x00, 0x69, 0x37),
        "secondary":        RGBColor(0x65, 0x66, 0x81),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#250E62", "#656681", "#006937", "#FFCB05", "#F5A899", "#EE7623", "#05CDCD", "#006680"],
    },
    "TEZSPIRE": {
        "client":           "AZN",
        "therapy_area":     "Respiratory",
        "primary_current":  RGBColor(0x00, 0x38, 0x65),
        "primary_prior":    RGBColor(0xFF, 0xD0, 0x5D),
        "secondary":        RGBColor(0x85, 0xBF, 0xFF),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#003865", "#85BFFF", "#FFD05D", "#00B5E2", "#C688E8", "#830051", "#9B37FF", "#EF426F"],
    },

    # DSI / DSI-AZN portfolio
    "DATROWAY": {
        "client":           "DSI_AZN",
        "therapy_area":     "Oncology/mBC",
        "primary_current":  RGBColor(0x1E, 0x22, 0xAA),
        "primary_prior":    RGBColor(0x64, 0x34, 0x66),
        "secondary":        RGBColor(0xEE, 0x76, 0x23),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#1E22AA", "#EE7623", "#643466", "#FBCCB1", "#018EAF", "#00ABC7", "#94A61F", "#F5AD7B"],
    },
    "ENHERTU": {
        "client":           "DSI_AZN",
        "therapy_area":     "Oncology/HER2",
        "primary_current":  RGBColor(0xFF, 0x88, 0x13),
        "primary_prior":    RGBColor(0x33, 0xA4, 0xFF),
        "secondary":        RGBColor(0xE1, 0x57, 0x59),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FF8813", "#E15759", "#33A4FF", "#59A14F", "#9467BD", "#1D7B87", "#91C3D5", "#F0FF4D"],
    },

    # GSK portfolio
    "BLENREP": {
        "client":           "GSK",
        "therapy_area":     "Oncology/MM",
        "primary_current":  RGBColor(0x3A, 0xB5, 0x1D),
        "primary_prior":    RGBColor(0x5C, 0x10, 0x3B),
        "secondary":        RGBColor(0x66, 0x58, 0xA6),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#3AB51D", "#6658A6", "#5C103B", "#668EDD", "#7030A0", "#17B3AF", "#1B3B7A", "#E21860"],
    },
    "JEMPERLI_ZEJULA": {   # Compound brand used in joint report
        "client":           "GSK",
        "therapy_area":     "Oncology/Endometrial",
        "primary_current":  RGBColor(0xD5, 0x19, 0x00),
        "primary_prior":    RGBColor(0x70, 0x63, 0x52),
        "secondary":        RGBColor(0x00, 0x84, 0x7C),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#D51900", "#00847C", "#706352", "#A9A197", "#3F348E", "#E7004E", "#668EDD", "#FFC000"],
    },
    "OJJAARA": {
        "client":           "GSK",
        "therapy_area":     "Hematology/MF",
        "primary_current":  RGBColor(0x6B, 0xB2, 0x3E),
        "primary_prior":    RGBColor(0xEE, 0x53, 0x40),
        "secondary":        RGBColor(0x61, 0x14, 0xFF),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#6BB23E", "#6114FF", "#EE5340", "#3A5C6D", "#6014FF", "#DDDDDD", "#FC6666", "#AB162C"],
    },

    # Pfizer / EMD
    "ABRYSVO": {
        "client":           "PFIZER",
        "therapy_area":     "Vaccines/Maternal",
        "primary_current":  RGBColor(0x00, 0x00, 0xC9),
        "primary_prior":    RGBColor(0x43, 0x96, 0x4A),
        "secondary":        RGBColor(0xE3, 0x56, 0xDC),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#0000C9", "#E356DC", "#43964A", "#F49C34", "#BA2A4C", "#0095FF", "#D3770B", "#0070BF"],
    },
    "BAVENCIO": {
        "client":           "EMD",
        "therapy_area":     "Oncology",
        "primary_current":  RGBColor(0x00, 0x00, 0xC9),
        "primary_prior":    RGBColor(0x40, 0x60, 0xAF),
        "secondary":        RGBColor(0x6D, 0x5A, 0x7A),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#0000C9", "#6D5A7A", "#4060AF", "#081E3D", "#808080", "#6FBF4A", "#E4B77F", "#43964A"],
    },

    # Regeneron (+ Sanofi)
    "LIBTAYO": {
        "client":           "REGENERON",
        "therapy_area":     "Oncology/NMSC",
        "primary_current":  RGBColor(0xDC, 0x00, 0x77),
        "primary_prior":    RGBColor(0xC8, 0x33, 0x33),
        "secondary":        RGBColor(0x00, 0x4F, 0x6F),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#DC0077", "#004F6F", "#C83333", "#96BBBB", "#C19875", "#338533", "#D34D4D", "#668998"],
    },
    "DUPIXENT_EOE": {
        "client":           "REGENERON_SANOFI",
        "therapy_area":     "Immunology/EoE",
        "primary_current":  RGBColor(0x21, 0x94, 0x91),
        "primary_prior":    RGBColor(0xE4, 0x6C, 0x0A),
        "secondary":        RGBColor(0xF7, 0x96, 0x46),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#219491", "#F79646", "#E46C0A", "#8EB4E3", "#00745A", "#7030A0", "#B370A4", "#7ABFBD"],
    },

    # Novartis
    "RHAPSIDO": {
        "client":           "NOVARTIS",
        "therapy_area":     "Immunology",
        "primary_current":  RGBColor(0x01, 0x8E, 0x86),
        "primary_prior":    RGBColor(0x00, 0x20, 0x68),
        "secondary":        RGBColor(0x00, 0x70, 0xFE),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#018E86", "#0070FE", "#002068", "#8F2DDE", "#852065", "#B56FA3", "#61A8FF", "#50E2D0"],
    },

    # Amgen
    "OTEZLA": {
        "client":           "AMGEN",
        "therapy_area":     "Dermatology/Psoriasis",
        "primary_current":  RGBColor(0x1F, 0x49, 0x7D),
        "primary_prior":    RGBColor(0x00, 0xA3, 0xDF),
        "secondary":        RGBColor(0x6F, 0x8D, 0xB4),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#1F497D", "#6F8DB4", "#00A3DF", "#4C35DE", "#A5A5A5", "#2E75B6", "#FFB81C", "#C303D4"],
    },
    "UPLIZNA": {
        "client":           "AMGEN",
        "therapy_area":     "Neurology/NMOSD",
        "primary_current":  RGBColor(0x81, 0x3F, 0x97),
        "primary_prior":    RGBColor(0x00, 0x3C, 0x71),
        "secondary":        RGBColor(0x00, 0x7D, 0x6D),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#813F97", "#007D6D", "#003C71", "#DA205F", "#00B0F0", "#E46C0A", "#D9D9D9", "#425563"],
    },

    # LEO
    "ADBRY": {
        "client":           "LEO",
        "therapy_area":     "Dermatology/AD",
        "primary_current":  RGBColor(0xC0, 0x14, 0xA3),
        "primary_prior":    RGBColor(0x6F, 0x43, 0x9A),
        "secondary":        RGBColor(0xD5, 0x68, 0x5F),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#C014A3", "#D5685F", "#6F439A", "#FED006", "#EAAFE0", "#009B77", "#D561C1", "#90117A"],
    },

    # Alexion
    "ULTOMIRIS": {
        "client":           "ALEXION",
        "therapy_area":     "Hematology/PNH",
        "primary_current":  RGBColor(0x0E, 0x87, 0x79),
        "primary_prior":    RGBColor(0xE3, 0x51, 0x05),
        "secondary":        RGBColor(0x70, 0x30, 0xA0),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#0E8779", "#7030A0", "#E35105", "#799A01", "#FFA300", "#FF8181", "#84CA68", "#009886"],
    },

    # B+L
    "MIEBO": {
        "client":           "BL",
        "therapy_area":     "Ophthalmology/DED",
        "primary_current":  RGBColor(0x00, 0xA9, 0xEB),
        "primary_prior":    RGBColor(0xA8, 0x10, 0x9D),
        "secondary":        RGBColor(0x40, 0x02, 0x86),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#00A9EB", "#400286", "#A8109D", "#FF1119", "#413A5F", "#000F9F", "#FFC000", "#F78626"],
    },

    # Otsuka
    "ABILIFY_MAINTENA": {
        "client":           "OTSUKA",
        "therapy_area":     "Psychiatry",
        "primary_current":  RGBColor(0xFF, 0xC0, 0x00),
        "primary_prior":    RGBColor(0x7F, 0x47, 0xAA),
        "secondary":        RGBColor(0x2D, 0x5C, 0xA2),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FFC000", "#2D5CA2", "#7F47AA", "#7D44A9", "#3B5998", "#7E97CD", "#B3C9EA", "#0000FF"],
    },

    # Apellis
    "EMPAVELI": {
        "client":           "APELLIS",
        "therapy_area":     "Hematology",
        "primary_current":  RGBColor(0xFC, 0x3B, 0x6E),
        "primary_prior":    RGBColor(0x81, 0xE3, 0xE3),
        "secondary":        RGBColor(0x30, 0xCF, 0xD0),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FC3B6E", "#30CFD0", "#81E3E3", "#5A4986", "#FD7FA0", "#FEB0C4", "#FFC000", "#163A6B"],
    },

    # Ipsen
    "ONIVYDE": {
        "client":           "IPSEN",
        "therapy_area":     "Oncology/Pancreatic",
        "primary_current":  RGBColor(0x54, 0xAC, 0x65),
        "primary_prior":    RGBColor(0x8E, 0xC8, 0x99),
        "secondary":        RGBColor(0xC8, 0x48, 0x74),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#54AC65", "#C84874", "#8EC899", "#BBDEC2", "#D2E9D6", "#A9D5B2", "#00B039", "#387343"],
    },

    # Cross-brand / aggregate trackers (no single brand)
    "CCA_PP_TRACKER": {
        "client":           "CCA",
        "therapy_area":     "Cross-brand",
        "primary_current":  RGBColor(0xF2, 0x8E, 0x2B),
        "primary_prior":    RGBColor(0xBD, 0xD7, 0xEE),
        "secondary":        RGBColor(0x8F, 0xAA, 0xDC),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#F28E2B", "#8FAADC", "#BDD7EE", "#AFED5D", "#FFABD5", "#FF6700", "#FF8F43", "#FFBC8F"],
    },
    "ALL_PET_KPI": {
        "client":           "UNKNOWN",
        "therapy_area":     "Cross-brand",
        "primary_current":  RGBColor(0x29, 0x49, 0x83),
        "primary_prior":    RGBColor(0x8F, 0xAA, 0xDC),
        "secondary":        RGBColor(0x5C, 0x84, 0xCC),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#294983", "#5C84CC", "#8FAADC", "#F0406E", "#82BC00", "#133053"],
    },
    "PHYSICIANS_SFE_PET": {
        "client":           "UNKNOWN",
        "therapy_area":     "Cross-brand",
        "primary_current":  RGBColor(0xE7, 0x00, 0x4C),
        "primary_prior":    RGBColor(0x6B, 0xC9, 0xEF),
        "secondary":        RGBColor(0xC7, 0xA0, 0x13),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#E7004C", "#C7A013", "#6BC9EF", "#006838", "#FF6969", "#00607C", "#4E79A7", "#00A44A"],
    },
    "BONE_HCP_TRACKER": {
        "client":           "UNKNOWN",
        "therapy_area":     "Bone",
        "primary_current":  RGBColor(0xFF, 0x99, 0x33),
        "primary_prior":    RGBColor(0x00, 0xB0, 0xF0),
        "secondary":        RGBColor(0x00, 0x92, 0x01),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#FF9933", "#009201", "#00B0F0", "#58193D", "#0063C3", "#4BACC6", "#E46C0A", "#425563"],
    },

    # ── Known competitor brands (placeholder entries — colors from prior work) ──
    # Competitors that appear as context in other brands' decks. Palette to be
    # confirmed when we analyze their own PET decks.
    "TAGRISSO": {
        "client":           "AZN",
        "therapy_area":     "Oncology/NSCLC",
        "primary_current":  RGBColor(0x70, 0x30, 0xA0),   # classic Tagrisso purple
        "primary_prior":    RGBColor(0xAD, 0x88, 0xC8),
        "secondary":        RGBColor(0xC8, 0xA4, 0xE2),
        "positive":         POSITIVE_GREEN,
        "negative":         NEGATIVE_RED,
        "_observed_palette": ["#7030A0", "#AD88C8", "#C8A4E2"],  # from Rybrevant deck cross-references
        "_source":          "observed as competitor in RYBREVANT deck",
    },

    # ── Legacy aliases (preserved for backward compat) ──
    "jnj": {
        "client":           "JJ",
        "therapy_area":     "Legacy/RybrevantAlias",
        "primary_current":  RGBColor(0xF7, 0x58, 0x24),  # Rybrevant orange (was hardcoded as "jnj")
        "primary_prior":    RGBColor(0xFF, 0xC1, 0x99),
        "secondary":        RGBColor(0x70, 0x30, 0xA0),  # Tagrisso competitor (legacy behavior)
        "competitor":       RGBColor(0x70, 0x30, 0xA0),
        "competitor_prior": RGBColor(0xAD, 0x88, 0xC8),
        "accent":           RGBColor(0xFF, 0x00, 0x00),  # J&J red (legacy)
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
        "client":           "UNKNOWN",
        "therapy_area":     "Generic",
        "primary_current":  RGBColor(0x44, 0x72, 0xC4),
        "primary_prior":    RGBColor(0xA9, 0xC5, 0xE8),
        "secondary":        RGBColor(0xED, 0x7D, 0x31),
        "competitor":       RGBColor(0xED, 0x7D, 0x31),
        "competitor_prior": RGBColor(0xF5, 0xBE, 0x97),
        "accent":           RGBColor(0x44, 0x72, 0xC4),
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
}


# ── Lookup helpers ──────────────────────────────────────────────────────────


def _normalize(name: str) -> str:
    return name.strip().replace(" ", "_").replace("-", "_").replace("+", "_").upper()


def get_brand(brand_name: str) -> dict:
    """Look up brand config. Merges BRAND entry with its CLIENT entry so callers
    get fonts/heading_color/template_path inline without a second lookup.

    Accepts case-insensitive brand names with spaces/hyphens/plus normalized.
    Preserves legacy lowercase aliases ("jnj", "default") unchanged.
    """
    # Try direct lookup first (preserves legacy lowercase keys like "jnj" / "default"
    # and exact matches for brand keys like "RYBREVANT")
    if brand_name in BRAND:
        key = brand_name
    else:
        key = _normalize(brand_name)
        # Check CLIENT_ALIASES for long-form folder names (e.g., "ASTRAZENECA" → "AZN")
        key = CLIENT_ALIASES.get(key, key)
        if key not in BRAND:
            raise KeyError(
                f"Unknown brand {brand_name!r}. Available: {sorted(BRAND.keys())}"
            )
    brand = BRAND[key].copy()
    client_key = brand.get("client")
    if client_key and client_key in CLIENT:
        # Merge client-level defaults. Brand entries can override via their own keys;
        # setdefault ensures we don't clobber explicit per-brand settings.
        for k, v in CLIENT[client_key].items():
            brand.setdefault(k, v)
    return brand


def get_color(brand: str, role: str) -> RGBColor:
    """Shortcut: `get_color("RYBREVANT", "primary_current")` → `RGBColor`."""
    return get_brand(brand)[role]


def get_competitor(brand_name: str) -> dict | None:
    """If the brand has a known competitor_brand, return that competitor's merged dict.

    Enables automatic competitor-color lookup:
        rybrevant = get_brand("RYBREVANT")
        tagrisso = get_competitor("RYBREVANT")  # auto-resolved
        # use rybrevant["primary_current"] for Rybrevant bars
        # use tagrisso["primary_current"] for Tagrisso bars
    """
    brand = get_brand(brand_name)
    comp_name = brand.get("competitor_brand")
    if comp_name and _normalize(comp_name) in BRAND:
        return get_brand(comp_name)
    return None


# ── Legacy color aliases (derived from BRAND["jnj"] / Rybrevant) ────────────
# Existing code imports these by name — preserved unchanged.

_jnj = BRAND["jnj"]

C_RYB_Q4   = _jnj["primary_current"]       # deep orange — Rybrevant Q4 bars
C_RYB_Q3   = _jnj["primary_prior"]         # pale orange — Rybrevant Q3 bars
C_TAG      = _jnj["competitor"]            # purple — Tagrisso
C_RED      = _jnj["accent"]                # J&J red
C_GREEN    = _jnj["positive"]
C_GREY     = _jnj["neutral_grey"]

# Structural colors (client-independent)
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_FTGREY   = RGBColor(0x7F, 0x7F, 0x7F)
C_LBGREY   = RGBColor(0xF4, 0xF4, 0xF4)
C_HDRGREY  = RGBColor(0x40, 0x40, 0x40)
C_LTGREY   = RGBColor(0xBF, 0xBF, 0xBF)

# ── COM colour equivalents (BGR order) ───────────────────────────────────────

COM_RED    = 0x0000FF
COM_ORANGE = 0x2458F7
COM_GREEN  = 0x50B000
COM_GREY   = 0x505050

# ── Fonts (derived from BRAND["jnj"]) ────────────────────────────────────────

FONT_DISPLAY = _jnj["font_display"]
FONT_TEXT    = _jnj["font_body"]
