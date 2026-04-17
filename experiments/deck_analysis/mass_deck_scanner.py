"""
Mass Deck Scanner — Regenerate pptx_utils from 400-500 client decks.
=====================================================================

Walks a directory tree of PPTX files (organized as client/project/deck.pptx),
extracts chart types, layout coordinates, brand colors, headline patterns, and
component compositions from every slide. Then **generates updated pptx_utils
Python source** directly from the full corpus:

  outputs/generated_brand.py          — BRAND{} per client (replaces current 33 entries)
  outputs/generated_layouts.py        — LAYOUTS{} from coordinate clusters
  outputs/generated_chart_patterns.py — CHART_PATTERNS{} with frequency stats
  outputs/mass_scan_inventory.json    — Raw per-deck metadata (for debugging)
  outputs/mass_scan_summary.md        — Human-readable summary of corpus

This is the same pattern as the original 40-deck generate_pptx_utils.py, but
grounded on the full ~500-deck corpus instead of a 40-deck sample.

Usage:
  python experiments/deck_analysis/mass_deck_scanner.py \\
    --decks-dir "C:/path/to/sharepoint-pptx-scanner/downloads" \\
    [--max-decks 50]       # limit for testing
    [--skip-errors]        # continue on corrupt PPTX files
    [--resume]             # skip already-scanned decks (reads mass_scan_inventory.json)

The script is resumable — if interrupted, re-run with --resume.
"""
from __future__ import annotations

import json
import statistics
import sys
import time
import traceback
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Optional

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

HERE = Path(__file__).parent
REPO = HERE.parents[1]

# Import deep OOXML extraction from the existing deep_analyzer (same directory)
sys.path.insert(0, str(HERE))
from deep_analyzer import analyze_chart_xml, extract_chart_xmls
OUTPUTS_DIR = HERE / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

# Map XL_CHART_TYPE enum → our pattern keys
_CHART_TYPE_MAP = {
    XL_CHART_TYPE.BAR_CLUSTERED: "bar_clustered_horizontal",
    XL_CHART_TYPE.BAR_STACKED: "bar_stacked_horizontal",
    XL_CHART_TYPE.BAR_STACKED_100: "bar_stacked_100_horizontal",
    XL_CHART_TYPE.COLUMN_CLUSTERED: "column_clustered_vertical",
    XL_CHART_TYPE.COLUMN_STACKED: "column_stacked_vertical",
    XL_CHART_TYPE.COLUMN_STACKED_100: "column_stacked_100_vertical",
    XL_CHART_TYPE.LINE: "line_markers_trended",
    XL_CHART_TYPE.LINE_MARKERS: "line_markers_trended",
    XL_CHART_TYPE.LINE_STACKED: "line_stacked",
    XL_CHART_TYPE.XY_SCATTER: "xy_scatter_abacus",
    XL_CHART_TYPE.XY_SCATTER_LINES: "xy_scatter_lines",
    XL_CHART_TYPE.DOUGHNUT: "doughnut_default",
    XL_CHART_TYPE.PIE: "pie_default",
    XL_CHART_TYPE.AREA_STACKED: "area_stacked",
    XL_CHART_TYPE.BUBBLE: "bubble",
    XL_CHART_TYPE.RADAR: "radar",
}

# Project type classification heuristics (from file path + deck name)
_PROJECT_TYPE_KEYWORDS = {
    "PET": ["PET", "Promotional Effectiveness", "Promo Effectiveness",
            "Message Recall", "Rep Performance"],
    "ATU": ["ATU", "Awareness Trial Usage", "Awareness, Trial",
            "Brand Tracking", "Brand Health"],
    "HCP-Pt": ["HCP-Pt", "HCP Patient", "HCP-Patient", "Patient Dialogue",
               "Conversation", "Patient Research"],
    "Digital Tracker": ["Digital Tracker", "Digital Track", "NPP",
                        "Non-Personal Promotion", "Omnichannel"],
    "PCA": ["PCA", "Patient Chart Audit", "Chart Review"],
    "MaxDiff": ["MaxDiff", "Max Diff", "Message Testing"],
    "SOV": ["SOV", "Share of Voice"],
    "Qualitative": ["Qual", "Qualitative", "Focus Group", "IDI"],
}

# Section/topic classification from headline text (feeds project-type skills)
_SECTION_KEYWORDS = {
    # PET sections
    "Message Recall": ["message recall", "msg recall", "unaided recall", "aided recall",
                       "recall of messages"],
    "Message Effectiveness": ["message effectiveness", "effectiveness of messages",
                              "message impact"],
    "Rep Performance": ["rep performance", "rep quality", "sales rep", "representative",
                        "rep rating", "SFE", "sales force"],
    "Prescription Intent": ["prescription intent", "intent to prescribe", "prescribing",
                            "likelihood to prescribe", "Rx intent"],
    "Call to Action": ["call to action", "CTA", "action taken"],
    "HII": ["high impact interaction", "HII", "high-impact"],
    # ATU sections
    "Awareness": ["awareness", "unaided awareness", "aided awareness", "brand awareness"],
    "Trial": ["trial", "ever tried", "ever prescribed"],
    "Usage": ["usage", "current use", "currently using", "share of", "SOV"],
    "Loyalty": ["loyalty", "switching", "switch intent", "brand switch"],
    "Satisfaction": ["satisfaction", "patient satisfaction", "HCP satisfaction"],
    # HCP-Pt sections
    "Patient Conversations": ["patient conversation", "patient dialogue",
                              "HCP-patient", "treatment discussion"],
    "Treatment Journey": ["treatment journey", "treatment path", "patient journey"],
    # Digital Tracker sections
    "Digital Engagement": ["digital engagement", "digital channel", "email",
                           "website", "banner ad", "online"],
    "Non-Personal Promotion": ["non-personal", "NPP", "omnichannel"],
    # Cross-cutting
    "Drivers and Barriers": ["driver", "barrier", "reason for", "reason not"],
    "Competitive Landscape": ["competitive", "competitor", "vs.", "versus", "comparison"],
    "Executive Summary": ["executive summary", "key findings", "summary"],
    "Methodology": ["methodology", "sample", "respondent profile", "study design"],
    "Recommendations": ["recommendation", "implication", "next steps"],
}


def _classify_section(headline: str) -> str:
    """Classify a headline into a section/topic."""
    hl = headline.lower()
    for section, keywords in _SECTION_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in hl:
                return section
    return "Other"


def _emu_to_inches(emu: int) -> float:
    return round(emu / 914400, 2) if emu else 0.0


def _classify_project_type(path: Path) -> str:
    full_text = str(path).upper()
    for ptype, keywords in _PROJECT_TYPE_KEYWORDS.items():
        for kw in keywords:
            if kw.upper() in full_text:
                return ptype
    return "Unknown"


def _extract_client_name(path: Path, base_dir: Path) -> str:
    try:
        rel = path.relative_to(base_dir)
        parts = rel.parts
        if len(parts) >= 2:
            return parts[0]
    except ValueError:
        pass
    return "Unknown"


def _extract_project_name(path: Path, base_dir: Path) -> str:
    try:
        rel = path.relative_to(base_dir)
        parts = rel.parts
        if len(parts) >= 3:
            return parts[1]
    except ValueError:
        pass
    return "Unknown"


def _client_key(name: str) -> str:
    """Normalize client name to a Python dict key."""
    # Strip parenthetical abbreviations: "AbbVie (ABV)" → "ABBVIE"
    import re
    name = re.sub(r'\s*\([^)]*\)\s*', '', name)
    return name.upper().strip().replace(" ", "_").replace("-", "_").replace("&", "AND")


# ---------------------------------------------------------------------------
# Per-deck analysis
# ---------------------------------------------------------------------------

def analyze_deck(pptx_path: Path, base_dir: Path) -> dict:
    """Analyze one PPTX file. Returns metadata dict."""
    prs = Presentation(str(pptx_path))
    client = _extract_client_name(pptx_path, base_dir)
    project = _extract_project_name(pptx_path, base_dir)
    project_type = _classify_project_type(pptx_path)

    deck = {
        "file": str(pptx_path.relative_to(base_dir)),
        "client": client,
        "client_key": _client_key(client),
        "project": project,
        "project_type": project_type,
        "total_slides": len(prs.slides),
        "total_charts": 0,
        "total_tables": 0,
        "chart_types": Counter(),
        "chart_patterns": Counter(),
        "series_colors": Counter(),       # per-deck for brand extraction
        "heading_colors": Counter(),
        "fonts": Counter(),
        "font_sizes": Counter(),
        "headline_lengths": [],
        "headline_font_sizes": [],
        "sections": Counter(),               # section/topic classification from headlines
        "headlines": [],                      # sample headlines (first 50)
        "composition_signatures": Counter(),
        "chart_positions": [],            # (left, top, width, height)
        "table_positions": [],
        "table_dimensions": Counter(),
        # Per-slide positions grouped by composition signature (for LAYOUTS{})
        "positions_by_sig": defaultdict(lambda: {"chart": [], "table": []}),
        # OOXML properties (from deep_analyzer.analyze_chart_xml)
        "ooxml": {
            "gap_widths": Counter(),
            "overlaps": Counter(),
            "data_label_positions": Counter(),
            "axis_orientations": Counter(),
            "tick_lbl_positions": Counter(),
            "num_formats": Counter(),
            "marker_types": Counter(),
            "line_widths": Counter(),
            "charts_with_legend": 0,
            "charts_with_gridlines": 0,
            "charts_with_title": 0,
        },
    }

    for slide_idx, slide in enumerate(prs.slides):
        chart_count = 0
        table_count = 0
        slide_chart_positions = []
        slide_table_positions = []

        for shape in slide.shapes:
            if shape.has_chart:
                chart_count += 1
                deck["total_charts"] += 1
                chart = shape.chart

                try:
                    ct = chart.chart_type
                    ct_str = f"{ct.name} ({ct.value})" if ct else "unknown"
                    pattern = _CHART_TYPE_MAP.get(ct, f"unmapped_{ct.name}" if ct else "unknown")
                except Exception:
                    ct_str = "unknown"
                    pattern = "unknown"

                deck["chart_types"][ct_str] += 1
                deck["chart_patterns"][pattern] += 1

                pos = (
                    _emu_to_inches(shape.left or 0),
                    _emu_to_inches(shape.top or 0),
                    _emu_to_inches(shape.width or 0),
                    _emu_to_inches(shape.height or 0),
                )
                deck["chart_positions"].append(pos)
                slide_chart_positions.append(pos)

                try:
                    for s in chart.plots[0].series:
                        try:
                            fill = s.format.fill
                            if fill.type is not None:
                                rgb = f"#{fill.fore_color.rgb}"
                                deck["series_colors"][rgb] += 1
                        except Exception:
                            pass
                except Exception:
                    pass

            elif shape.has_table:
                table_count += 1
                deck["total_tables"] += 1
                tbl = shape.table
                deck["table_dimensions"][f"{len(tbl.rows)}x{len(tbl.columns)}"] += 1
                pos = (
                    _emu_to_inches(shape.left or 0),
                    _emu_to_inches(shape.top or 0),
                    _emu_to_inches(shape.width or 0),
                    _emu_to_inches(shape.height or 0),
                )
                deck["table_positions"].append(pos)
                slide_table_positions.append(pos)

            elif shape.has_text_frame:
                text = shape.text_frame.text.strip()
                top = _emu_to_inches(shape.top or 0)

                if top < 1.5 and len(text) > 10:
                    deck["headline_lengths"].append(len(text))
                    # Section classification
                    section = _classify_section(text)
                    deck["sections"][section] += 1
                    if len(deck["headlines"]) < 50:
                        deck["headlines"].append(text[:150])
                    try:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.size:
                                    deck["headline_font_sizes"].append(round(run.font.size.pt, 1))
                                if run.font.color and run.font.color.rgb:
                                    deck["heading_colors"][f"#{run.font.color.rgb}"] += 1
                                break
                            break
                    except Exception:
                        pass

                try:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.name:
                                deck["fonts"][run.font.name] += 1
                except Exception:
                    pass

        sig = f"{chart_count}_chart_{table_count}_table"
        deck["composition_signatures"][sig] += 1
        # Track positions per composition signature (for LAYOUTS{})
        deck["positions_by_sig"][sig]["chart"].extend(slide_chart_positions)
        deck["positions_by_sig"][sig]["table"].extend(slide_table_positions)

    # Deep OOXML extraction (gapWidth, overlap, dLblPos, axis orientation, etc.)
    try:
        chart_xmls = extract_chart_xmls(pptx_path)
        for _name, xml_bytes in chart_xmls:
            info = analyze_chart_xml(xml_bytes)
            if "error" in info:
                continue
            ox = deck["ooxml"]
            for gw in info.get("gap_widths", []):
                ox["gap_widths"][str(gw)] += 1
            for ov in info.get("overlaps", []):
                ox["overlaps"][str(ov)] += 1
            for dlp in info.get("data_label_positions", []):
                ox["data_label_positions"][str(dlp)] += 1
            for ao in info.get("axis_orientations", []):
                key = f"{ao['axis']}:{ao['orientation']}"
                ox["axis_orientations"][key] += 1
            for tlp in info.get("tick_lbl_pos", []):
                key = f"{tlp['axis']}:{tlp['val']}"
                ox["tick_lbl_positions"][key] += 1
            for nf in info.get("num_formats_used", []):
                ox["num_formats"][str(nf)] += 1
            for nf in info.get("val_axis_number_formats", []):
                ox["num_formats"][str(nf)] += 1
            for mt in info.get("series_marker_types", []):
                ox["marker_types"][str(mt)] += 1
            for lw in info.get("series_line_widths", []):
                if lw:
                    ox["line_widths"][str(lw)] += 1
            if info.get("has_legend"):
                ox["charts_with_legend"] += 1
            if info.get("major_gridlines_present"):
                ox["charts_with_gridlines"] += 1
            if info.get("has_title"):
                ox["charts_with_title"] += 1
    except Exception:
        pass  # OOXML extraction is best-effort; don't fail the deck scan

    # Serialize Counters and defaultdicts for JSON
    for key in ["chart_types", "chart_patterns", "series_colors", "heading_colors",
                "fonts", "font_sizes", "composition_signatures", "table_dimensions",
                "sections"]:
        deck[key] = dict(Counter(deck[key]).most_common(50))
    deck["positions_by_sig"] = {k: dict(v) for k, v in deck["positions_by_sig"].items()}
    # Serialize OOXML counters
    for key in ["gap_widths", "overlaps", "data_label_positions", "axis_orientations",
                "tick_lbl_positions", "num_formats", "marker_types", "line_widths"]:
        deck["ooxml"][key] = dict(Counter(deck["ooxml"][key]).most_common(30))

    return deck


# ---------------------------------------------------------------------------
# Aggregation
# ---------------------------------------------------------------------------

def aggregate_for_codegen(inventories: list[dict]) -> dict:
    """Aggregate all deck data into structures ready for codegen."""

    # ── Per-client brand data ──
    # client_key → {series_colors: Counter, fonts: Counter, heading_colors: Counter, deck_count}
    by_client: dict[str, dict] = defaultdict(lambda: {
        "series_colors": Counter(),
        "fonts": Counter(),
        "heading_colors": Counter(),
        "deck_count": 0,
        "raw_name": "",
    })

    # ── Global aggregates ──
    all_chart_patterns = Counter()
    all_chart_types = Counter()
    all_composition_sigs = Counter()
    all_table_dims = Counter()
    all_chart_positions = []    # flat list of (left, top, width, height)
    all_table_positions = []
    all_headline_lengths = []
    all_headline_font_sizes = []
    project_types = Counter()
    clients = Counter()

    # ── Per-composition-signature position lists (for LAYOUTS{} clustering) ──
    # sig → {"chart_positions": [...], "table_positions": [...]}
    positions_by_sig: dict[str, dict] = defaultdict(lambda: {
        "chart_positions": [], "table_positions": [],
    })

    for d in inventories:
        ck = d.get("client_key", _client_key(d["client"]))
        by_client[ck]["deck_count"] += 1
        by_client[ck]["raw_name"] = d["client"]
        for color, count in d["series_colors"].items():
            by_client[ck]["series_colors"][color] += count
        for font, count in d["fonts"].items():
            by_client[ck]["fonts"][font] += count
        for hc, count in d.get("heading_colors", {}).items():
            by_client[ck]["heading_colors"][hc] += count

        for k, v in d["chart_patterns"].items():
            all_chart_patterns[k] += v
        for k, v in d["chart_types"].items():
            all_chart_types[k] += v
        for k, v in d["composition_signatures"].items():
            all_composition_sigs[k] += v
        for k, v in d["table_dimensions"].items():
            all_table_dims[k] += v

        all_chart_positions.extend(d.get("chart_positions", []))
        all_table_positions.extend(d.get("table_positions", []))
        all_headline_lengths.extend(d.get("headline_lengths", []))
        all_headline_font_sizes.extend(d.get("headline_font_sizes", []))
        project_types[d["project_type"]] += 1
        clients[d["client"]] += 1

        # Per-signature positions (tracked per-slide in analyze_deck)
        for sig, pos_data in d.get("positions_by_sig", {}).items():
            positions_by_sig[sig]["chart_positions"].extend(pos_data.get("chart", []))
            positions_by_sig[sig]["table_positions"].extend(pos_data.get("table", []))

    # ── OOXML aggregation ──
    ooxml_agg = {
        "gap_widths": Counter(),
        "overlaps": Counter(),
        "data_label_positions": Counter(),
        "axis_orientations": Counter(),
        "tick_lbl_positions": Counter(),
        "num_formats": Counter(),
        "marker_types": Counter(),
        "line_widths": Counter(),
        "charts_with_legend": 0,
        "charts_with_gridlines": 0,
        "charts_with_title": 0,
    }
    for d in inventories:
        ox = d.get("ooxml", {})
        for key in ["gap_widths", "overlaps", "data_label_positions", "axis_orientations",
                     "tick_lbl_positions", "num_formats", "marker_types", "line_widths"]:
            for k, v in ox.get(key, {}).items():
                ooxml_agg[key][k] += v
        for key in ["charts_with_legend", "charts_with_gridlines", "charts_with_title"]:
            ooxml_agg[key] += ox.get(key, 0)

    # Serialize OOXML counters
    for key in ["gap_widths", "overlaps", "data_label_positions", "axis_orientations",
                "tick_lbl_positions", "num_formats", "marker_types", "line_widths"]:
        ooxml_agg[key] = dict(ooxml_agg[key].most_common(30))

    # ── Per-project-type profiles (for project-type skills) ──
    by_project_type: dict[str, dict] = defaultdict(lambda: {
        "deck_count": 0,
        "total_slides": 0,
        "total_charts": 0,
        "total_tables": 0,
        "chart_patterns": Counter(),
        "composition_signatures": Counter(),
        "sections": Counter(),
        "avg_slides_per_deck": 0,
        "sample_headlines": [],
    })

    for d in inventories:
        pt = d["project_type"]
        bpt = by_project_type[pt]
        bpt["deck_count"] += 1
        bpt["total_slides"] += d["total_slides"]
        bpt["total_charts"] += d["total_charts"]
        bpt["total_tables"] += d["total_tables"]
        for k, v in d["chart_patterns"].items():
            bpt["chart_patterns"][k] += v
        for k, v in d["composition_signatures"].items():
            bpt["composition_signatures"][k] += v
        for k, v in d.get("sections", {}).items():
            bpt["sections"][k] += v
        if len(bpt["sample_headlines"]) < 20:
            bpt["sample_headlines"].extend(d.get("headlines", [])[:5])

    # Finalize per-type profiles
    for pt, bpt in by_project_type.items():
        bpt["avg_slides_per_deck"] = round(bpt["total_slides"] / max(bpt["deck_count"], 1), 1)
        bpt["chart_patterns"] = dict(Counter(bpt["chart_patterns"]).most_common(15))
        bpt["composition_signatures"] = dict(Counter(bpt["composition_signatures"]).most_common(15))
        bpt["sections"] = dict(Counter(bpt["sections"]).most_common(20))
        bpt["sample_headlines"] = bpt["sample_headlines"][:20]

    return {
        "by_client": dict(by_client),
        "by_project_type": dict(by_project_type),
        "chart_patterns": dict(all_chart_patterns.most_common(30)),
        "chart_types": dict(all_chart_types.most_common(30)),
        "composition_signatures": dict(all_composition_sigs.most_common(30)),
        "table_dimensions": dict(all_table_dims.most_common(30)),
        "chart_positions": all_chart_positions,
        "table_positions": all_table_positions,
        "positions_by_sig": {k: dict(v) for k, v in positions_by_sig.items()},
        "headline_lengths": all_headline_lengths,
        "headline_font_sizes": all_headline_font_sizes,
        "project_types": dict(project_types.most_common(20)),
        "clients": dict(clients.most_common(60)),
        "total_decks": len(inventories),
        "total_slides": sum(d["total_slides"] for d in inventories),
        "total_charts": sum(d["total_charts"] for d in inventories),
        "total_tables": sum(d["total_tables"] for d in inventories),
        "ooxml": ooxml_agg,
    }


# ---------------------------------------------------------------------------
# Codegen: BRAND{}
# ---------------------------------------------------------------------------

# Manual overrides for known brand primaries (from prior analysis + brand guides)
_BRAND_PRIMARY_OVERRIDES = {
    "JJ": "0063C3",
    "GSK": "F36633",
}

_BRAND_HEADING_OVERRIDES = {
    "JJ": "001E60",
    "AZN": "595959",
    "PFIZER": "0063C3",
    "AMGEN": "003C71",
    "REGENERON": "00745A",
    "GSK": "F36633",
}


def _hex_to_rgb_tuple(hex_str: str) -> tuple[int, int, int]:
    h = hex_str.lstrip("#")
    if len(h) < 6:
        h = h.ljust(6, "0")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def generate_brand_py(agg: dict) -> str:
    """Generate CLIENT{} Python source from aggregated per-client data.

    This produces client-level entries (fonts, heading colors, observed palettes).
    Per-brand entries (RYBREVANT, TAGRISSO, etc.) remain hand-curated in the
    existing BRAND{} — they require per-product color knowledge that can't be
    reliably auto-extracted from client-level aggregation.
    """
    by_client = agg["by_client"]
    n_decks = agg["total_decks"]

    lines = [
        '"""',
        "Client-level brand defaults — fonts, heading colors, observed palettes.",
        "",
        f"Generated from deck analysis of {n_decks} decks across {len(by_client)} clients.",
        "See experiments/deck_analysis/mass_deck_scanner.py for the analysis.",
        "",
        "CLIENT{} provides client-level defaults (shared across all brands for that client).",
        "Per-brand entries (product-level colors) live in BRAND{} and are hand-curated.",
        "",
        "Each CLIENT entry provides:",
        "  - font_heading:        headline font",
        "  - font_body:           body text font",
        "  - heading_color:       headline text color",
        "  - observed_palette:    top 8 series colors observed across all decks for this client",
        "  - deck_count:          number of decks analyzed",
        '"""',
        "from __future__ import annotations",
        "",
        "from pptx.dml.color import RGBColor",
        "",
        "",
        "# Universal delta colors — observed across all decks",
        "POSITIVE_GREEN = RGBColor(0x00, 0xB0, 0x50)",
        "NEGATIVE_RED = RGBColor(0xFF, 0x00, 0x00)",
        "NEGATIVE_DEEP_RED = RGBColor(0xC0, 0x00, 0x00)",
        "",
        "# Standard greys",
        "GREY_DARK = RGBColor(0x40, 0x40, 0x40)",
        "GREY_MID = RGBColor(0x59, 0x59, 0x59)",
        "GREY_LIGHT = RGBColor(0xBF, 0xBF, 0xBF)",
        "GREY_ALT_ROW = RGBColor(0xF2, 0xF2, 0xF2)",
        "",
        "",
        "CLIENT = {",
    ]

    # Sort clients by deck count descending
    sorted_clients = sorted(
        by_client.items(),
        key=lambda kv: -kv[1]["deck_count"],
    )

    for ck, info in sorted_clients:
        dc = info["deck_count"]
        raw = info.get("raw_name", ck)
        top_colors = [c for c, _ in Counter(info["series_colors"]).most_common(10)]
        top_fonts = [f for f, _ in Counter(info["fonts"]).most_common(5)]
        top_heading = [c for c, _ in Counter(info["heading_colors"]).most_common(3)]

        # Fonts: skip theme placeholders
        real_fonts = [f for f in top_fonts if not f.startswith("+")]
        font_body = real_fonts[0] if real_fonts else "Arial"
        font_heading = real_fonts[0] if real_fonts else "Arial"

        # Heading color
        heading_hex = _BRAND_HEADING_OVERRIDES.get(ck)
        if not heading_hex and top_heading:
            heading_hex = top_heading[0].lstrip("#")
        if not heading_hex:
            heading_hex = "000000"

        lines.append(f'    # --- {raw} ({dc} deck{"s" if dc != 1 else ""}) ---')
        lines.append(f'    "{ck}": {{')

        r, g, b = _hex_to_rgb_tuple(heading_hex)
        lines.append(f'        "heading_color": RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')
        lines.append(f'        "font_heading": "{font_heading}",')
        lines.append(f'        "font_body": "{font_body}",')
        lines.append(f'        "template_path": None,')
        lines.append(f'        "deck_count": {dc},')

        palette = ", ".join(f'"{c}"' for c in top_colors[:8])
        lines.append(f'        "_observed_palette": [{palette}],')
        lines.append("    },")
        lines.append("")

    lines += [
        "}",
        "",
        "",
        "def get_client(client_key: str) -> dict:",
        '    """Look up client defaults by key. Raises KeyError with available keys."""',
        "    key = client_key.upper().replace(' ', '_').replace('-', '_')",
        "    if key not in CLIENT:",
        "        raise KeyError(",
        '            f"Unknown client {client_key!r}. Available: {sorted(CLIENT.keys())}"',
        "        )",
        "    return CLIENT[key]",
        "",
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Codegen: CHART_PATTERNS{}
# ---------------------------------------------------------------------------

def generate_chart_patterns_py(agg: dict) -> str:
    """Generate CHART_PATTERNS{} Python source from chart type frequencies + OOXML defaults."""
    patterns = agg["chart_patterns"]
    total = sum(patterns.values())
    n_decks = agg["total_decks"]
    ox = agg.get("ooxml", {})

    # Most common OOXML defaults (mode values across the full corpus)
    def mode_val(counter_dict):
        if not counter_dict:
            return None
        return max(counter_dict, key=counter_dict.get)

    default_gap = mode_val(ox.get("gap_widths", {}))
    default_overlap = mode_val(ox.get("overlaps", {}))
    default_dlbl = mode_val(ox.get("data_label_positions", {}))
    default_numfmt = mode_val(ox.get("num_formats", {}))
    default_marker = mode_val(ox.get("marker_types", {}))

    lines = [
        '"""',
        "Chart pattern definitions — deterministic rendering configs per chart type.",
        "",
        f"Generated from deck analysis of {n_decks} decks ({total:,} charts total).",
        "Frequency-ranked. OOXML defaults from corpus-wide mode values.",
        "",
        "OOXML property summary (corpus-wide):",
        f"  gapWidth modes:     {dict(sorted(ox.get('gap_widths', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  overlap modes:      {dict(sorted(ox.get('overlaps', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  dLblPos modes:      {dict(sorted(ox.get('data_label_positions', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  numFmt modes:       {dict(sorted(ox.get('num_formats', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  axis orientations:  {dict(sorted(ox.get('axis_orientations', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  tickLblPos modes:   {dict(sorted(ox.get('tick_lbl_positions', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  marker types:       {dict(sorted(ox.get('marker_types', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  line widths (EMU):  {dict(sorted(ox.get('line_widths', {}).items(), key=lambda x: -x[1])[:5])}",
        f"  charts w/ legend:   {ox.get('charts_with_legend', 0)}/{total}",
        f"  charts w/ gridlines:{ox.get('charts_with_gridlines', 0)}/{total}",
        f"  charts w/ title:    {ox.get('charts_with_title', 0)}/{total}",
        '"""',
        "from __future__ import annotations",
        "",
        "",
        "# Corpus-wide OOXML defaults (mode values)",
        f"DEFAULT_GAP_WIDTH = {default_gap}" if default_gap else "DEFAULT_GAP_WIDTH = 80",
        f"DEFAULT_OVERLAP = {default_overlap}" if default_overlap else "DEFAULT_OVERLAP = 100",
        f'DEFAULT_DLBL_POS = "{default_dlbl}"' if default_dlbl else 'DEFAULT_DLBL_POS = "ctr"',
        f'DEFAULT_NUM_FORMAT = "{default_numfmt}"' if default_numfmt else 'DEFAULT_NUM_FORMAT = "0%"',
        f'DEFAULT_MARKER_TYPE = "{default_marker}"' if default_marker else 'DEFAULT_MARKER_TYPE = "circle"',
        "",
        "",
        "CHART_PATTERNS = {",
    ]

    for pattern, count in sorted(patterns.items(), key=lambda x: -x[1]):
        pct = round(count * 100 / total, 1) if total else 0
        lines.append(f'    "{pattern}": {{')
        lines.append(f'        "occurrences": {count},')
        lines.append(f'        "pct": {pct},')
        lines.append(f'        "gap_width": DEFAULT_GAP_WIDTH,')
        lines.append(f'        "overlap": DEFAULT_OVERLAP,')
        lines.append(f'        "dlbl_pos": DEFAULT_DLBL_POS,')
        lines.append(f'        "num_format": DEFAULT_NUM_FORMAT,')
        lines.append("    },")

    lines += [
        "}",
        "",
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Codegen: LAYOUTS{}
# ---------------------------------------------------------------------------

def _median(values: list[float]) -> float:
    return round(statistics.median(values), 2) if values else 0.0


def generate_layouts_py(agg: dict) -> str:
    """Generate LAYOUTS{} Python source from position clusters per composition signature."""
    sigs = agg["composition_signatures"]
    total_slides = agg["total_slides"]
    n_decks = agg["total_decks"]

    # Compute median positions across all charts / all tables
    all_cp = agg["chart_positions"]
    all_tp = agg["table_positions"]

    lines = [
        '"""',
        "Layout presets for slide compositions.",
        "",
        f"Generated from deck analysis of {n_decks} decks ({total_slides:,} slides).",
        "Coordinate medians from observed chart + table positions across the full corpus.",
        '"""',
        "from __future__ import annotations",
        "",
        "",
        "SLIDE_WIDTH = 13.333",
        "SLIDE_HEIGHT = 7.5",
        "",
        'HEADLINE_RECT = {"left": 0.2, "top": 0.3, "width": 12.8, "height": 0.9}',
        'FOOTER_RECT = {"left": 0.2, "top": 7.0, "width": 12.8, "height": 0.4}',
        "",
        "",
        "LAYOUTS = {",
        "",
    ]

    def rect_str(left, top, width, height):
        return f'{{"left": {left}, "top": {top}, "width": {width}, "height": {height}}}'

    pos_by_sig = agg.get("positions_by_sig", {})

    # Global delta column medians (narrow tables: <=1" wide, >=3" tall)
    delta_cols = [(l, t, w, h) for l, t, w, h in all_tp if w <= 1.0 and h >= 3.0]
    if delta_cols:
        dc_left = _median([p[0] for p in delta_cols])
        dc_top = _median([p[1] for p in delta_cols])
        dc_w = _median([p[2] for p in delta_cols])
        dc_h = _median([p[3] for p in delta_cols])
    else:
        dc_left, dc_top, dc_w, dc_h = 12.5, 2.0, 0.55, 4.5

    # Generate entries for top composition signatures using PER-SIGNATURE medians
    top_sigs = sorted(sigs.items(), key=lambda x: -x[1])

    for sig, count in top_sigs[:15]:
        if count < 10:
            continue
        pct = round(count * 100 / total_slides, 1) if total_slides else 0

        sig_data = pos_by_sig.get(sig, {"chart_positions": [], "table_positions": []})
        sig_cp = sig_data.get("chart_positions", [])
        sig_tp = sig_data.get("table_positions", [])

        # Per-signature chart medians
        if sig_cp:
            cl = _median([p[0] for p in sig_cp])
            ct = _median([p[1] for p in sig_cp])
            cw = _median([p[2] for p in sig_cp])
            ch = _median([p[3] for p in sig_cp])
        else:
            cl, ct, cw, ch = 3.5, 2.0, 6.0, 4.5

        # Per-signature table medians
        if sig_tp:
            tl = _median([p[0] for p in sig_tp])
            tt = _median([p[1] for p in sig_tp])
            tw = _median([p[2] for p in sig_tp])
            th = _median([p[3] for p in sig_tp])
        else:
            tl, tt, tw, th = 0.5, 2.0, 2.5, 4.5

        layout_key = f"observed_{sig}"

        lines.append(f'    # ---- {sig} ({count} slides, {pct}%) ----')
        lines.append(f'    #   chart positions: {len(sig_cp)}, table positions: {len(sig_tp)}')
        lines.append(f'    "{layout_key}": {{')
        if sig_cp:
            lines.append(f'        "chart_rect": {rect_str(cl, ct, cw, ch)},')
        if sig_tp:
            lines.append(f'        "table_rect": {rect_str(tl, tt, tw, th)},')
        lines.append(f'        "delta_col_rect": {rect_str(dc_left, dc_top, dc_w, dc_h)},')
        lines.append(f'        "_slides": {count},')
        lines.append("    },")
        lines.append("")

    lines += [
        "}",
        "",
        "",
        "def get_layout(layout_key: str) -> dict:",
        '    """Look up a layout preset by key."""',
        "    if layout_key not in LAYOUTS:",
        "        raise KeyError(",
        '            f"Unknown layout {layout_key!r}. Available: {sorted(LAYOUTS.keys())}"',
        "        )",
        "    return LAYOUTS[layout_key]",
        "",
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Codegen: Per-project-type profiles
# ---------------------------------------------------------------------------

def generate_project_type_profiles(agg: dict) -> dict[str, str]:
    """Generate a markdown profile per project type for project-type skills.

    Returns {project_type: markdown_string}.
    """
    by_pt = agg.get("by_project_type", {})
    profiles = {}

    for pt, data in sorted(by_pt.items(), key=lambda x: -x[1]["deck_count"]):
        dc = data["deck_count"]
        if dc < 2:
            continue  # skip singleton types

        total_charts = data["total_charts"] or 1
        lines = [
            f"# {pt} Project Type Profile",
            "",
            f"**Decks analyzed:** {dc}",
            f"**Total slides:** {data['total_slides']:,}",
            f"**Avg slides/deck:** {data['avg_slides_per_deck']}",
            f"**Total charts:** {data['total_charts']:,}",
            f"**Total tables:** {data['total_tables']:,}",
            "",
            "---",
            "",
            "## Chart Pattern Distribution",
            "",
            "| Pattern | Count | % |",
            "|---|---|---|",
        ]
        for p, count in sorted(data["chart_patterns"].items(), key=lambda x: -x[1]):
            pct = round(count * 100 / total_charts, 1)
            lines.append(f"| `{p}` | {count:,} | {pct}% |")

        lines += [
            "",
            "## Slide Composition Signatures",
            "",
            "| Signature | Slides |",
            "|---|---|",
        ]
        for sig, count in sorted(data["composition_signatures"].items(), key=lambda x: -x[1]):
            lines.append(f"| `{sig}` | {count:,} |")

        lines += [
            "",
            "## Section/Topic Distribution",
            "",
            "| Section | Slides |",
            "|---|---|",
        ]
        for sec, count in sorted(data["sections"].items(), key=lambda x: -x[1]):
            lines.append(f"| {sec} | {count:,} |")

        if data.get("sample_headlines"):
            lines += [
                "",
                "## Sample Headlines",
                "",
            ]
            for hl in data["sample_headlines"][:15]:
                lines.append(f"- {hl}")

        lines.append("")
        profiles[pt] = "\n".join(lines)

    return profiles


# ---------------------------------------------------------------------------
# Summary report
# ---------------------------------------------------------------------------

def write_summary(agg: dict, output_path: Path) -> None:
    """Write a human-readable summary of the full corpus."""
    lines = [
        "# Mass Deck Scan — Corpus Summary",
        "",
        f"**Decks scanned:** {agg['total_decks']}",
        f"**Total slides:** {agg['total_slides']:,}",
        f"**Total charts:** {agg['total_charts']:,}",
        f"**Total tables:** {agg['total_tables']:,}",
        f"**Clients:** {len(agg['clients'])}",
        "",
        "---",
        "",
        "## Project Type Distribution",
        "",
        "| Project Type | Decks | % |",
        "|---|---|---|",
    ]
    total = agg["total_decks"] or 1
    for pt, count in sorted(agg["project_types"].items(), key=lambda x: -x[1]):
        lines.append(f"| {pt} | {count} | {count*100//total}% |")

    lines += [
        "",
        "## Client Distribution (top 30)",
        "",
        "| Client | Decks |",
        "|---|---|",
    ]
    for client, count in sorted(agg["clients"].items(), key=lambda x: -x[1])[:30]:
        lines.append(f"| {client} | {count} |")

    lines += [
        "",
        "## Chart Patterns (frequency-ranked)",
        "",
        "| Pattern | Count | % |",
        "|---|---|---|",
    ]
    total_charts = agg["total_charts"] or 1
    for p, count in sorted(agg["chart_patterns"].items(), key=lambda x: -x[1]):
        pct = round(count * 100 / total_charts, 1)
        lines.append(f"| `{p}` | {count:,} | {pct}% |")

    lines += [
        "",
        "## Composition Signatures (top 15)",
        "",
        "| Signature | Slides |",
        "|---|---|",
    ]
    for sig, count in sorted(agg["composition_signatures"].items(), key=lambda x: -x[1])[:15]:
        lines.append(f"| `{sig}` | {count:,} |")

    lines += [
        "",
        "## Table Dimensions (top 15)",
        "",
        "| Rows x Cols | Count |",
        "|---|---|",
    ]
    for dim, count in sorted(agg["table_dimensions"].items(), key=lambda x: -x[1])[:15]:
        lines.append(f"| {dim} | {count:,} |")

    ox = agg.get("ooxml", {})
    lines += [
        "",
        "## OOXML Properties (across all charts)",
        "",
        "| Property | Top values |",
        "|---|---|",
    ]
    for prop in ["gap_widths", "overlaps", "data_label_positions", "axis_orientations",
                 "tick_lbl_positions", "num_formats", "marker_types", "line_widths"]:
        vals = ox.get(prop, {})
        top = ", ".join(f"{k}={v}" for k, v in sorted(vals.items(), key=lambda x: -x[1])[:5])
        lines.append(f"| {prop} | {top} |")
    lines.append(f"| charts_with_legend | {ox.get('charts_with_legend', 0)} |")
    lines.append(f"| charts_with_gridlines | {ox.get('charts_with_gridlines', 0)} |")
    lines.append(f"| charts_with_title | {ox.get('charts_with_title', 0)} |")

    hl = agg["headline_lengths"]
    hfs = agg["headline_font_sizes"]
    lines += [
        "",
        "## Headlines",
        f"- Average length: {round(sum(hl)/len(hl),1) if hl else 0} chars",
        f"- Median font size: {statistics.median(hfs) if hfs else 0}pt",
        "",
        "---",
        "",
        "## Generated Outputs",
        "",
        "| File | What it updates |",
        "|---|---|",
        "| `outputs/generated_brand.py` | `slidegen/pptx_utils/brand.py` — BRAND{} per client |",
        "| `outputs/generated_chart_patterns.py` | `slidegen/pptx_utils/charts.py` — CHART_PATTERNS{} |",
        "| `outputs/generated_layouts.py` | `slidegen/pptx_utils/layout.py` — LAYOUTS{} |",
        "",
        "Review the generated files, then copy into the corresponding `slidegen/pptx_utils/` modules.",
        "",
    ]

    output_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def find_pptx_files(base_dir: Path) -> list[Path]:
    """Recursively find all .pptx files (skip temp ~$ files)."""
    return sorted(p for p in base_dir.rglob("*.pptx") if not p.name.startswith("~$"))


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Mass Deck Scanner — regenerate pptx_utils from full deck corpus",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Full scan
  python experiments/deck_analysis/mass_deck_scanner.py \\
    --decks-dir "C:/path/to/sharepoint-pptx-scanner/downloads" \\
    --skip-errors

  # Test with first 10 decks
  python experiments/deck_analysis/mass_deck_scanner.py \\
    --decks-dir "C:/path/to/downloads" --max-decks 10 --skip-errors

  # Resume interrupted scan
  python experiments/deck_analysis/mass_deck_scanner.py \\
    --decks-dir "..." --skip-errors --resume
        """,
    )
    parser.add_argument("--decks-dir", required=True, help="Root dir of PPTX files")
    parser.add_argument("--max-decks", type=int, default=0, help="Limit (0 = all)")
    parser.add_argument("--skip-errors", action="store_true")
    parser.add_argument("--resume", action="store_true")
    args = parser.parse_args()

    base_dir = Path(args.decks_dir)
    if not base_dir.is_dir():
        print(f"ERROR: Not a directory: {base_dir}")
        sys.exit(1)

    pptx_files = find_pptx_files(base_dir)
    print(f"Found {len(pptx_files)} PPTX files in {base_dir}")

    if args.max_decks > 0:
        pptx_files = pptx_files[:args.max_decks]
        print(f"Limited to first {args.max_decks}")

    # Resume support
    inventories = []
    scanned_files = set()
    inventory_path = OUTPUTS_DIR / "mass_scan_inventory.json"
    if args.resume and inventory_path.exists():
        try:
            inventories = json.loads(inventory_path.read_text(encoding="utf-8"))
            scanned_files = {d["file"] for d in inventories}
            print(f"Resuming: {len(scanned_files)} already scanned")
        except Exception:
            pass

    # Scan
    errors = []
    t0 = time.time()
    for i, pptx_path in enumerate(pptx_files):
        rel = str(pptx_path.relative_to(base_dir))
        if rel in scanned_files:
            continue

        print(f"  [{i+1}/{len(pptx_files)}] {rel[:80]}...", end="", flush=True)
        try:
            meta = analyze_deck(pptx_path, base_dir)
            inventories.append(meta)
            print(f" {meta['total_slides']}s/{meta['total_charts']}c [{meta['project_type']}]")
        except Exception as exc:
            print(f" ERROR: {type(exc).__name__}: {str(exc)[:60]}")
            errors.append({"file": rel, "error": str(exc)[:200]})
            if not args.skip_errors:
                break

        # Periodic checkpoint
        if len(inventories) % 25 == 0:
            inventory_path.write_text(json.dumps(inventories, indent=2, default=str), encoding="utf-8")

    elapsed = time.time() - t0
    print(f"\nScanned {len(inventories)} decks in {elapsed:.1f}s ({len(errors)} errors)")

    # Save inventory
    inventory_path.write_text(json.dumps(inventories, indent=2, default=str), encoding="utf-8")

    if errors:
        (OUTPUTS_DIR / "mass_scan_errors.json").write_text(
            json.dumps(errors, indent=2), encoding="utf-8")

    # Aggregate
    print("Aggregating...")
    agg = aggregate_for_codegen(inventories)

    # Generate pptx_utils Python source
    print("Generating pptx_utils source...")

    brand_py = generate_brand_py(agg)
    brand_path = OUTPUTS_DIR / "generated_brand.py"
    brand_path.write_text(brand_py, encoding="utf-8")
    print(f"  {brand_path} ({len(agg['by_client'])} clients)")

    patterns_py = generate_chart_patterns_py(agg)
    patterns_path = OUTPUTS_DIR / "generated_chart_patterns.py"
    patterns_path.write_text(patterns_py, encoding="utf-8")
    print(f"  {patterns_path} ({len(agg['chart_patterns'])} patterns)")

    layouts_py = generate_layouts_py(agg)
    layouts_path = OUTPUTS_DIR / "generated_layouts.py"
    layouts_path.write_text(layouts_py, encoding="utf-8")
    print(f"  {layouts_path}")

    # Summary
    summary_path = OUTPUTS_DIR / "mass_scan_summary.md"
    write_summary(agg, summary_path)
    print(f"  {summary_path}")

    # Per-project-type profiles (for project-type skills)
    profiles = generate_project_type_profiles(agg)
    profiles_dir = OUTPUTS_DIR / "project_type_profiles"
    profiles_dir.mkdir(exist_ok=True)
    for pt, md_content in profiles.items():
        safe_name = pt.lower().replace(" ", "_").replace("-", "_")
        profile_path = profiles_dir / f"{safe_name}_profile.md"
        profile_path.write_text(md_content, encoding="utf-8")
    print(f"  {profiles_dir}/ ({len(profiles)} project types)")

    print(f"\n{'='*60}")
    print(f"DONE — {agg['total_decks']} decks, {agg['total_charts']:,} charts, "
          f"{len(agg['by_client'])} clients, {len(profiles)} project types")
    print(f"{'='*60}")
    print(f"Generated files:")
    print(f"  {brand_path}  — CLIENT{{}}")
    print(f"  {patterns_path}  — CHART_PATTERNS{{}}")
    print(f"  {layouts_path}  — LAYOUTS{{}}")
    for pt in sorted(profiles.keys()):
        safe_name = pt.lower().replace(" ", "_").replace("-", "_")
        print(f"  {profiles_dir}/{safe_name}_profile.md")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
