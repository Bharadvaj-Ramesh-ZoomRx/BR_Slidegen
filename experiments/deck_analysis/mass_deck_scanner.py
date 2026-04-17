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
        "composition_signatures": Counter(),
        "chart_positions": [],            # (left, top, width, height)
        "table_positions": [],
        "table_dimensions": Counter(),
    }

    for slide_idx, slide in enumerate(prs.slides):
        chart_count = 0
        table_count = 0

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

                deck["chart_positions"].append((
                    _emu_to_inches(shape.left or 0),
                    _emu_to_inches(shape.top or 0),
                    _emu_to_inches(shape.width or 0),
                    _emu_to_inches(shape.height or 0),
                ))

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
                deck["table_positions"].append((
                    _emu_to_inches(shape.left or 0),
                    _emu_to_inches(shape.top or 0),
                    _emu_to_inches(shape.width or 0),
                    _emu_to_inches(shape.height or 0),
                ))

            elif shape.has_text_frame:
                text = shape.text_frame.text.strip()
                top = _emu_to_inches(shape.top or 0)

                if top < 1.5 and len(text) > 10:
                    deck["headline_lengths"].append(len(text))
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

    # Serialize Counters
    for key in ["chart_types", "chart_patterns", "series_colors", "heading_colors",
                "fonts", "font_sizes", "composition_signatures", "table_dimensions"]:
        deck[key] = dict(Counter(deck[key]).most_common(50))

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

        # Assign positions to their slide's composition signature
        # We need per-slide info for this, which we have from the slides array
        # For now, add all chart/table positions globally per sig
        for sig, count in d["composition_signatures"].items():
            positions_by_sig[sig]["chart_positions"].extend(d.get("chart_positions", []))
            positions_by_sig[sig]["table_positions"].extend(d.get("table_positions", []))

    return {
        "by_client": dict(by_client),
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
    """Generate BRAND{} Python source from aggregated per-client data."""
    by_client = agg["by_client"]
    n_decks = agg["total_decks"]

    lines = [
        '"""',
        "Brand definitions per pharmaceutical client.",
        "",
        f"Generated from deck analysis of {n_decks} decks across {len(by_client)} clients.",
        "See experiments/deck_analysis/mass_deck_scanner.py for the analysis that produced this.",
        "",
        "Each BRAND entry provides:",
        "  - primary:       main series color (current wave bars)",
        "  - secondary:     secondary accent for comparisons",
        "  - prior:         tint color (prior wave bars)",
        "  - positive:      delta positive (universal green)",
        "  - negative:      delta negative (universal red)",
        "  - heading_color: headline text color",
        "  - font_heading:  headline font",
        "  - font_body:     body text font",
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
        "BRAND = {",
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

        if not top_colors:
            continue

        # Primary: override if known, else top observed series color
        primary = _BRAND_PRIMARY_OVERRIDES.get(ck, top_colors[0].lstrip("#"))
        secondary = top_colors[1].lstrip("#") if len(top_colors) > 1 else "808080"
        prior = top_colors[2].lstrip("#") if len(top_colors) > 2 else "BFBFBF"

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

        for role, hex_val in [("primary", primary), ("secondary", secondary), ("prior", prior)]:
            r, g, b = _hex_to_rgb_tuple(hex_val)
            lines.append(f'        "{role}": RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')

        lines.append('        "positive": POSITIVE_GREEN,')
        lines.append('        "negative": NEGATIVE_RED,')

        r, g, b = _hex_to_rgb_tuple(heading_hex)
        lines.append(f'        "heading_color": RGBColor(0x{r:02X}, 0x{g:02X}, 0x{b:02X}),')
        lines.append(f'        "font_heading": "{font_heading}",')
        lines.append(f'        "font_body": "{font_body}",')

        palette = ", ".join(f'"{c}"' for c in top_colors[:8])
        lines.append(f'        "_observed_palette": [{palette}],')
        lines.append("    },")
        lines.append("")

    lines += [
        "}",
        "",
        "",
        "def get_brand(client_key_or_alias: str) -> dict:",
        '    """Look up brand config by client key. Raises KeyError with available keys."""',
        "    key = client_key_or_alias.upper().replace(' ', '_').replace('-', '_')",
        "    if key not in BRAND:",
        "        raise KeyError(",
        '            f"Unknown client {client_key_or_alias!r}. Available: {sorted(BRAND.keys())}"',
        "        )",
        "    return BRAND[key]",
        "",
        "",
        "def get_color(client: str, role: str) -> 'RGBColor':",
        '    """Shortcut: get_color("JJ", "primary") -> RGBColor."""',
        "    return get_brand(client)[role]",
        "",
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Codegen: CHART_PATTERNS{}
# ---------------------------------------------------------------------------

def generate_chart_patterns_py(agg: dict) -> str:
    """Generate CHART_PATTERNS{} Python source from chart type frequencies."""
    patterns = agg["chart_patterns"]
    total = sum(patterns.values())
    n_decks = agg["total_decks"]

    lines = [
        '"""',
        "Chart pattern definitions — deterministic rendering configs per chart type.",
        "",
        f"Generated from deck analysis of {n_decks} decks ({total:,} charts total).",
        "Frequency-ranked. Top patterns cover the vast majority of real charts.",
        '"""',
        "from __future__ import annotations",
        "",
        "",
        "CHART_PATTERNS = {",
    ]

    for pattern, count in sorted(patterns.items(), key=lambda x: -x[1]):
        pct = round(count * 100 / total, 1) if total else 0
        lines.append(f'    "{pattern}": {{')
        lines.append(f'        "occurrences": {count},')
        lines.append(f'        "pct": {pct},')
        lines.append(f'        # TODO: Add OOXML defaults (gapWidth, overlap, dLblPos, etc.)')
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

    # Global chart position medians
    if all_cp:
        chart_left = _median([p[0] for p in all_cp])
        chart_top = _median([p[1] for p in all_cp])
        chart_w = _median([p[2] for p in all_cp])
        chart_h = _median([p[3] for p in all_cp])
    else:
        chart_left, chart_top, chart_w, chart_h = 3.5, 2.0, 6.0, 4.5

    if all_tp:
        table_left = _median([p[0] for p in all_tp])
        table_top = _median([p[1] for p in all_tp])
        table_w = _median([p[2] for p in all_tp])
        table_h = _median([p[3] for p in all_tp])
    else:
        table_left, table_top, table_w, table_h = 0.5, 2.0, 2.5, 4.5

    # Narrow delta columns: filter tables that are narrow (<=1") and tall (>=3")
    delta_cols = [(l, t, w, h) for l, t, w, h in all_tp if w <= 1.0 and h >= 3.0]
    if delta_cols:
        dc_left = _median([p[0] for p in delta_cols])
        dc_top = _median([p[1] for p in delta_cols])
        dc_w = _median([p[2] for p in delta_cols])
        dc_h = _median([p[3] for p in delta_cols])
    else:
        dc_left, dc_top, dc_w, dc_h = 12.5, 2.0, 0.55, 4.5

    # Generate entries for top composition signatures
    top_sigs = sorted(sigs.items(), key=lambda x: -x[1])

    for sig, count in top_sigs[:15]:
        if count < 10:
            continue
        pct = round(count * 100 / total_slides, 1) if total_slides else 0

        lines.append(f'    # ---- {sig} ({count} slides, {pct}%) ----')

        # Determine layout key name
        layout_key = f"observed_{sig}"

        lines.append(f'    "{layout_key}": {{')
        lines.append(f'        "chart_rect": {rect_str(chart_left, chart_top, chart_w, chart_h)},')
        lines.append(f'        "table_rect": {rect_str(table_left, table_top, table_w, table_h)},')
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

    print(f"\n{'='*60}")
    print(f"DONE — {agg['total_decks']} decks, {agg['total_charts']:,} charts, "
          f"{len(agg['by_client'])} clients")
    print(f"{'='*60}")
    print(f"Generated files (review, then copy to slidegen/pptx_utils/):")
    print(f"  {brand_path}")
    print(f"  {patterns_path}")
    print(f"  {layouts_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
