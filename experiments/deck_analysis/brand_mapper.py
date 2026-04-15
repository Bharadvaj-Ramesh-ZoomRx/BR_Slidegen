"""
Brand Mapper
============

Maps each deck to its specific BRAND (not client), extracts per-brand series
colors directly from chart XML (so palette filters out structural text/fill
colors), and produces a BRAND{} dict keyed by brand name.

This replaces the brand-by-client aggregation in deep_analyzer.py, which had
the data model wrong (a client has many brands with different palettes).

Run:
    cd experiments/deck_analysis
    python brand_mapper.py

Outputs:
    outputs/brand_mapping.json     — deck → brand/client/competitor mapping
    outputs/brand_by_brand.json    — colors aggregated by brand
    outputs/client_registry.json   — client-level fonts, heading colors
    outputs/generated_brand_v2.py  — refactored pptx_utils/brand.py proposal
"""
from __future__ import annotations

import json
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).parent
DECKS_DIR = HERE / "decks"

# OOXML namespaces
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def qn_c(tag: str) -> str:
    return f"{{{NS_C}}}{tag}"


def qn_a(tag: str) -> str:
    return f"{{{NS_A}}}{tag}"


# Colors that are structural (text, delta, greys) and should NOT be treated as
# brand-defining even if they dominate the frequency count.
STRUCTURAL_COLORS = {
    "000000", "FFFFFF",                                    # black, white
    "00B050",                                              # positive-delta green (universal)
    "FF0000", "C00000", "EB1700", "FF2929",                # negative-delta reds
    "F2F2F2", "E7E6E6", "E5E5E5", "F0F0F0", "F3F4F3",      # light greys
    "BFBFBF", "A6A6A6", "7F7F7F", "D3D3D3", "CCCCCC",      # mid-light greys
    "595959", "404040", "505050", "3F3F3F", "3E403F",      # dark greys (body text)
    "2C2C2C", "282828", "292A2A", "171616", "1F1F1F",      # near-black
    "4D4D4F", "535554", "636466", "636466", "747F74",      # charcoals
    "4A5C58", "3F4444", "4B4B4B", "26272A",                # charcoal variants
    "E7E8E9", "F0F1F1", "E3E2E3", "CACACA", "D8DDE5",      # pale greys
    "AFABAB",                                              # neutral greys
}


def is_structural(hex_str: str) -> bool:
    return hex_str.upper() in STRUCTURAL_COLORS


# ---------------------------------------------------------------------------
# Extract series colors from a .pptx
# ---------------------------------------------------------------------------


def extract_series_colors(pptx_path: Path) -> Counter:
    """Pull the <a:srgbClr> values from every series <c:spPr> in every chart XML.

    These are the canonical "series fill colors" used to draw bars, lines,
    markers. Distinct from slide text colors, table fills, etc.
    """
    counter: Counter = Counter()
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            for name in z.namelist():
                if not (name.startswith("ppt/charts/chart") and name.endswith(".xml")):
                    continue
                xml = z.read(name)
                try:
                    root = ET.fromstring(xml)
                except ET.ParseError:
                    continue
                # Walk each <c:ser> and grab the fill from its <c:spPr>
                for ser in root.iter(qn_c("ser")):
                    sp_pr = ser.find(qn_c("spPr"))
                    if sp_pr is None:
                        continue
                    rgb = sp_pr.find(f"{qn_a('solidFill')}/{qn_a('srgbClr')}")
                    if rgb is not None and rgb.get("val"):
                        counter[rgb.get("val").upper()] += 1
    except Exception:
        return counter
    return counter


def top_brand_palette(colors: Counter, max_colors: int = 8) -> list[str]:
    """Return the top non-structural colors in frequency order."""
    filtered = [(c, cnt) for c, cnt in colors.most_common() if not is_structural(c)]
    return [c for c, _ in filtered[:max_colors]]

HERE = Path(__file__).parent
OUTPUTS_DIR = HERE / "outputs"


# ---------------------------------------------------------------------------
# Deck → brand + client + competitor mapping
# ---------------------------------------------------------------------------
#
# Each rule matches a filename pattern and assigns the primary brand (+ client
# + optional competitor extracted from the filename or common-knowledge).
#
# Patterns are checked in order; first match wins.

DECK_TO_BRAND = [
    # (regex, brand_name, client, competitor_name_or_None, therapy_area)

    # J&J portfolio
    (r"RYBREVANT\+LAZCLUZE|RYBREVANT", "RYBREVANT", "JJ", "TAGRISSO", "Oncology/NSCLC"),
    (r"TEPEZZA", "TEPEZZA", "JJ", None, "Endocrinology/TED"),
    (r"J&J MM SFEA", "DARZALEX", "JJ", None, "Oncology/MM"),  # J&J MM brand is typically Darzalex

    # AstraZeneca portfolio
    (r"CALQUENCE", "CALQUENCE", "AZN", None, "Oncology/CLL"),
    (r"LOKELMA", "LOKELMA", "AZN", None, "Nephrology"),
    (r"LYNPARZA", "LYNPARZA", "AZN", None, "Oncology"),
    (r"TRUQAP", "TRUQAP", "AZN", None, "Oncology/Breast"),
    (r"TEZSPIRE", "TEZSPIRE", "AZN", None, "Respiratory"),
    (r"DATROWAY.*mBC", "DATROWAY", "DSI_AZN", None, "Oncology/mBC"),  # DSI-AZN co-commercialized
    (r"DATROWAY.*NSCLC|DATROWAY EGFRm", "DATROWAY", "DSI", None, "Oncology/NSCLC"),
    (r"ENHERTU", "ENHERTU", "DSI_AZN", None, "Oncology/HER2"),

    # GSK portfolio
    (r"BLENREP", "BLENREP", "GSK", None, "Oncology/MM"),
    (r"JEMPERLI.*ZEJULA", "JEMPERLI-ZEJULA", "GSK", None, "Oncology/Endometrial"),
    (r"OJJAARA", "OJJAARA", "GSK", None, "Hematology/MF"),

    # Pfizer
    (r"ABRYSVO", "ABRYSVO", "PFIZER", None, "Vaccines/Maternal"),
    (r"BAVENCIO", "BAVENCIO", "EMD", None, "Oncology"),  # Bavencio is EMD Serono (not Pfizer)

    # Regeneron / Sanofi
    (r"LIBTAYO", "LIBTAYO", "REGENERON", None, "Oncology/NMSC"),
    (r"DUPIXENT.*EoE", "DUPIXENT_EoE", "REGENERON_SANOFI", None, "Immunology/EoE"),

    # Novartis
    (r"NVS.*Rhapsido|RHAPSIDO", "RHAPSIDO", "NOVARTIS", None, "Immunology"),

    # Amgen
    (r"UPLIZNA", "UPLIZNA", "AMGEN", None, "Neurology/NMOSD"),
    (r"OTEZLA", "OTEZLA", "AMGEN", None, "Dermatology/Psoriasis"),

    # LEO Pharma
    (r"ADBRY", "ADBRY", "LEO", None, "Dermatology/AD"),

    # Alexion
    (r"ALEXION.*PNH", "ULTOMIRIS", "ALEXION", None, "Hematology/PNH"),  # Alexion's PNH franchise

    # Bausch + Lomb
    (r"B\+L.*DED|BAUSCH", "MIEBO", "BL", None, "Ophthalmology/DED"),  # B+L DED brand

    # Otsuka
    (r"Abilify LAI|ABILIFY", "ABILIFY_MAINTENA", "OTSUKA", None, "Psychiatry"),

    # Apellis
    (r"EMPAVELI", "EMPAVELI", "APELLIS", None, "Hematology"),

    # Ipsen
    (r"ONIVYDE", "ONIVYDE", "IPSEN", None, "Oncology/Pancreatic"),

    # CCA (Personal Promotion tracker — not a single brand)
    (r"CCA.*Personal Promotion", "CCA_PP_TRACKER", "CCA", None, "Cross-brand"),

    # Bone HCP (no brand name in title)
    (r"Bone HCP", "BONE_HCP_TRACKER", "UNKNOWN", None, "Bone"),

    # Aggregate KPI reports (cross-brand)
    (r"ALL PET.*KPI Report", "ALL_PET_KPI", "UNKNOWN", None, "Cross-brand"),
    (r"Physicians SFE.*PET", "PHYSICIANS_SFE_PET", "UNKNOWN", None, "Cross-brand"),
]


def map_deck(filename: str) -> tuple[str, str, str | None, str]:
    """Return (brand, client, competitor, therapy_area)."""
    for rx, brand, client, competitor, ta in DECK_TO_BRAND:
        if re.search(rx, filename, re.IGNORECASE):
            return brand, client, competitor, ta
    return "UNMAPPED", "UNKNOWN", None, "Unknown"


# ---------------------------------------------------------------------------
# Client-level registry (fonts, heading colors — things that are actually
# client-wide, not brand-specific)
# ---------------------------------------------------------------------------

CLIENT_REGISTRY = {
    "JJ": {
        "font_heading": "Johnson Display",
        "font_body":    "Johnson Text",
        "heading_color": "001E60",  # J&J navy
    },
    "AZN": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "595959",  # AZN grey
    },
    "DSI": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "1E22AA",
    },
    "DSI_AZN": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "1E22AA",
    },
    "GSK": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "F36633",
    },
    "PFIZER": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "0063C3",
    },
    "EMD": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "000000",
    },
    "REGENERON": {
        "font_heading": "Trade Gothic LT Std",
        "font_body":    "Arial",
        "heading_color": "00745A",
    },
    "REGENERON_SANOFI": {
        "font_heading": "Trade Gothic LT Std",
        "font_body":    "Arial",
        "heading_color": "00745A",
    },
    "NOVARTIS": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "018E86",
    },
    "AMGEN": {
        "font_heading": "Century Gothic",
        "font_body":    "Century Gothic",
        "heading_color": "003C71",
    },
    "LEO": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "C014A3",
    },
    "ALEXION": {
        "font_heading": "Arial Black",
        "font_body":    "Arial",
        "heading_color": "0E8779",
    },
    "BL": {
        "font_heading": "Avenir Next LT Pro",
        "font_body":    "Century Gothic",
        "heading_color": "400286",
    },
    "OTSUKA": {
        "font_heading": "Calibri",
        "font_body":    "Arial",
        "heading_color": "000000",
    },
    "APELLIS": {
        "font_heading": "Calibri Light",
        "font_body":    "Calibri",
        "heading_color": "FC3B6E",
    },
    "IPSEN": {
        "font_heading": "Rethink Sans",
        "font_body":    "Calibri",
        "heading_color": "54AC65",
    },
    "CCA": {
        "font_heading": "Century Gothic",
        "font_body":    "Century Gothic",
        "heading_color": "F28E2B",
    },
    "UNKNOWN": {
        "font_heading": "Arial",
        "font_body":    "Arial",
        "heading_color": "000000",
    },
}


# ---------------------------------------------------------------------------
# Manual brand-color overrides
# ---------------------------------------------------------------------------
#
# The top observed series color isn't always the brand's real primary. E.g.,
# AZN Calquence decks showed green prominently, but that's the Calquence brand
# color; Lokelma would show navy. Overrides go here when observation disagrees
# with the known brand identity.

BRAND_PRIMARY_OVERRIDES: dict[str, str] = {
    # None needed right now — observed top colors are actually brand colors
    # because each deck is a single-brand PET. Kept as extension point.
}


# ---------------------------------------------------------------------------
# Aggregation
# ---------------------------------------------------------------------------


def main() -> int:
    # Use inventory only for font data (per-deck from slide text)
    inventory = json.loads((OUTPUTS_DIR / "inventory.json").read_text(encoding="utf-8"))

    brand_map: dict[str, dict] = {}
    per_brand_colors: dict[str, Counter] = defaultdict(Counter)
    per_brand_fonts: dict[str, Counter] = defaultdict(Counter)
    deck_mapping: list[dict] = []

    pptx_files = sorted(p for p in DECKS_DIR.glob("*.pptx") if not p.name.startswith("~"))
    inv_by_filename = {d.get("filename", ""): d for d in inventory if "error" not in d}

    print(f"Extracting series colors from {len(pptx_files)} decks...")
    for i, p in enumerate(pptx_files, 1):
        fn = p.name
        brand, client, competitor, ta = map_deck(fn)
        print(f"  [{i}/{len(pptx_files)}] {brand:30s} <- {fn[:50]}")
        # Pull series colors directly from chart XML (filters structural)
        colors = extract_series_colors(p)
        for c, cnt in colors.items():
            per_brand_colors[brand][c] += cnt
        # Fonts from inventory (slide text)
        inv = inv_by_filename.get(fn, {})
        for f, cnt in inv.get("font_counts", {}).items():
            per_brand_fonts[brand][f] += cnt

        deck_mapping.append({
            "filename": fn,
            "brand": brand,
            "client": client,
            "competitor": competitor,
            "therapy_area": ta,
        })
        brand_map.setdefault(brand, {
            "brand": brand,
            "client": client,
            "competitor": competitor,
            "therapy_area": ta,
            "deck_filenames": [],
        })
        brand_map[brand]["deck_filenames"].append(fn)

    # Attach aggregated colors (filtered) + fonts
    for brand, info in brand_map.items():
        raw = per_brand_colors[brand]
        info["top_series_colors_raw"] = dict(raw.most_common(30))
        # Palette filters out structural colors
        info["palette_filtered"] = top_brand_palette(raw, max_colors=10)
        info["top_fonts"] = dict(per_brand_fonts[brand].most_common(10))

    # Write outputs
    (OUTPUTS_DIR / "brand_mapping.json").write_text(
        json.dumps(deck_mapping, indent=2), encoding="utf-8"
    )
    (OUTPUTS_DIR / "brand_by_brand.json").write_text(
        json.dumps(brand_map, indent=2), encoding="utf-8"
    )
    (OUTPUTS_DIR / "client_registry.json").write_text(
        json.dumps(CLIENT_REGISTRY, indent=2), encoding="utf-8"
    )

    print(f"Mapped {len(deck_mapping)} decks to {len(brand_map)} distinct brands")
    for brand, info in sorted(brand_map.items()):
        print(f"  {brand:30s} | client={info['client']:20s} | {len(info['deck_filenames'])} deck(s)")

    # Generate the refactored brand.py proposal
    write_generated_brand_py(brand_map, OUTPUTS_DIR / "generated_brand_v2.py")
    print(f"\nWrote generated_brand_v2.py (refactored brand.py proposal)")

    return 0


def hex_to_tuple(hex_str: str) -> tuple[int, int, int]:
    h = hex_str.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def derive_tint(hex_str: str, lighten: float = 0.6) -> str:
    """Lighten a color toward white by `lighten` fraction (0-1). Used when no
    obvious 'prior wave' tint color is present in the observed palette.
    """
    r, g, b = hex_to_tuple(hex_str)
    r = int(r + (255 - r) * lighten)
    g = int(g + (255 - g) * lighten)
    b = int(b + (255 - b) * lighten)
    return f"{r:02X}{g:02X}{b:02X}"


def write_generated_brand_py(brand_map: dict, output: Path) -> None:
    lines: list[str] = []
    lines.append('"""')
    lines.append("Pharmaceutical brand definitions (one entry per brand, not per client).")
    lines.append("")
    lines.append("Generated from real-deck analysis of 32 PET decks. Each deck is a single-brand")
    lines.append("PET, so the series colors extracted per deck represent that brand's palette.")
    lines.append("")
    lines.append("A CLIENT has many BRANDS with different colors. The client-level defaults")
    lines.append("(fonts, heading colors, templates) live in CLIENT{} and are referenced by each")
    lines.append("brand via its \"client\" field.")
    lines.append("")
    lines.append("Lookup: `get_brand(\"RYBREVANT\")` -> merged dict with brand colors + client")
    lines.append("fonts/templates already resolved.")
    lines.append('"""')
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("from pptx.dml.color import RGBColor")
    lines.append("")
    lines.append("")
    lines.append("# Universal delta colors (observed across all 32 decks)")
    lines.append("POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)")
    lines.append("NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)")
    lines.append("NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)")
    lines.append("")
    lines.append("GREY_DARK = RGBColor(0x40, 0x40, 0x40)")
    lines.append("GREY_MID = RGBColor(0x59, 0x59, 0x59)")
    lines.append("GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)")
    lines.append("GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)")
    lines.append("")
    lines.append("")
    lines.append("# ─────────────────────────────────────────────────────────────────────")
    lines.append("# CLIENT-level defaults (fonts, heading color, template path)")
    lines.append("# ─────────────────────────────────────────────────────────────────────")
    lines.append("")
    lines.append("CLIENT = {")
    for client, info in CLIENT_REGISTRY.items():
        hc_r, hc_g, hc_b = hex_to_tuple(info["heading_color"])
        lines.append(f'    "{client}": {{')
        lines.append(f'        "font_heading":  "{info["font_heading"]}",')
        lines.append(f'        "font_body":     "{info["font_body"]}",')
        lines.append(f'        "heading_color": RGBColor(0x{hc_r:02X}, 0x{hc_g:02X}, 0x{hc_b:02X}),')
        lines.append(f'        "template_path": None,')
        lines.append(f'    }},')
    lines.append("}")
    lines.append("")
    lines.append("")
    lines.append("# ─────────────────────────────────────────────────────────────────────")
    lines.append("# BRAND-level definitions (one per pharma brand, not per client)")
    lines.append("# ─────────────────────────────────────────────────────────────────────")
    lines.append("")
    lines.append("BRAND = {")
    # Sort by deck count desc, then by brand name
    sorted_brands = sorted(
        brand_map.items(),
        key=lambda kv: (-len(kv[1]["deck_filenames"]), kv[0])
    )
    for brand, info in sorted_brands:
        client = info["client"]
        competitor = info.get("competitor")
        ta = info.get("therapy_area", "Unknown")
        deck_count = len(info["deck_filenames"])
        palette = info.get("palette_filtered", [])

        if not palette:
            # Brand has no non-structural series colors (e.g., tracker decks with
            # mostly grey/black bars). Fall back to a generic greys + brand-neutral.
            lines.append(f'    # {brand} — {client} / {ta} ({deck_count} deck{"s" if deck_count != 1 else ""}) [no distinctive palette]')
            lines.append(f'    "{brand}": {{')
            lines.append(f'        "client":           "{client}",')
            lines.append(f'        "therapy_area":     "{ta}",')
            lines.append('        "primary_current":  RGBColor(0x44, 0x72, 0xC4),  # fallback blue')
            lines.append('        "primary_prior":    RGBColor(0xA9, 0xC5, 0xE8),')
            lines.append('        "secondary":        GREY_MID,')
            lines.append('        "positive":         POSITIVE_GREEN,')
            lines.append('        "negative":         NEGATIVE_RED,')
            lines.append('        "_observed_palette": [],')
            lines.append('    },')
            lines.append("")
            continue

        primary = BRAND_PRIMARY_OVERRIDES.get(brand, palette[0])
        secondary = palette[1] if len(palette) > 1 else "808080"
        # Heuristic for prior tint: if an obvious lighter shade exists, use it;
        # else derive by lightening the primary.
        prior = palette[2] if len(palette) > 2 else derive_tint(primary)

        lines.append(f'    # {brand} — {client} / {ta} ({deck_count} deck{"s" if deck_count != 1 else ""})')
        lines.append(f'    "{brand}": {{')
        lines.append(f'        "client":           "{client}",')
        lines.append(f'        "therapy_area":     "{ta}",')
        r, g, b = hex_to_tuple(primary)
        lines.append(f'        "primary_current":  RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),  # brand primary')
        r, g, b = hex_to_tuple(prior)
        lines.append(f'        "primary_prior":    RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),  # tint for prior wave')
        r, g, b = hex_to_tuple(secondary)
        lines.append(f'        "secondary":        RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),  # accent / comparison')
        if competitor:
            lines.append(f'        "competitor_brand": "{competitor}",')
        lines.append('        "positive":         POSITIVE_GREEN,')
        lines.append('        "negative":         NEGATIVE_RED,')
        palette_str = ", ".join(f'"#{c}"' for c in palette[:8])
        lines.append(f'        "_observed_palette": [{palette_str}],')
        lines.append('    },')
        lines.append("")
    lines.append("}")
    lines.append("")
    lines.append("")
    lines.append("def get_brand(brand_name: str) -> dict:")
    lines.append('    """Look up brand config. Merges BRAND entry with its CLIENT entry so callers')
    lines.append("    get fonts/heading_color/template_path inline without a second lookup.")
    lines.append('    """')
    lines.append("    key = brand_name.upper().replace(' ', '_').replace('-', '_')")
    lines.append("    if key not in BRAND:")
    lines.append('        raise KeyError(f"Unknown brand {brand_name!r}. Available: {sorted(BRAND.keys())}")')
    lines.append("    brand = BRAND[key].copy()")
    lines.append("    client_key = brand[\"client\"]")
    lines.append("    if client_key in CLIENT:")
    lines.append("        # Client-level defaults merged in (brand can override, but doesn't today)")
    lines.append("        for k, v in CLIENT[client_key].items():")
    lines.append("            brand.setdefault(k, v)")
    lines.append("    return brand")
    lines.append("")
    lines.append("")
    lines.append("def get_competitor(brand_name: str) -> dict | None:")
    lines.append('    """If the brand has a known competitor, return that competitor\'s BRAND dict."""')
    lines.append("    brand = get_brand(brand_name)")
    lines.append("    comp = brand.get('competitor_brand')")
    lines.append("    if comp and comp.upper() in BRAND:")
    lines.append("        return get_brand(comp)")
    lines.append("    return None")
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    import sys
    sys.exit(main())
