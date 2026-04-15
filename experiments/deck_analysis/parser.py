"""
Deck Analysis Parser
====================

Reverse-engineers PET slide decks across clients to inform pptx_utils refactor
and skill library improvements. Produces:

  outputs/inventory.json   — Structured per-deck metadata
  outputs/report.md        — Human-readable summary
  outputs/chart_types.json — Chart type frequency map
  outputs/shape_types.json — Shape type frequency map
  outputs/colors.json      — Color palette analysis
  outputs/fonts.json       — Font usage analysis
  outputs/gap_analysis.md  — Coverage gaps vs. current 22 renderers

Run:
    cd experiments/deck_analysis
    python parser.py

All decks under decks/ are parsed. Outputs land in outputs/.
"""
from __future__ import annotations

import json
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

HERE = Path(__file__).parent
DECKS_DIR = HERE / "decks"
OUTPUTS_DIR = HERE / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Current 22 renderer inventory (for gap analysis)
# ---------------------------------------------------------------------------

EXISTING_RENDERERS = {
    "cover",
    "executive_summary",
    "single_bar_with_delta",
    "dual_bar_with_delta",
    "dual_bar_qoq",
    "qoq_bar_with_delta",
    "two_section_bar",
    "clustered_compare",
    "dual_bar_compare",
    "stacked_order",
    "hii_scorecard",
    "dual_doughnut",
    "abacus",
    "dual_abacus",
    "followup_rep",
    "message_mbd",
    "trended_scorecard",
    "trended_activity",
    "quadrant_scatter",
    "heatmap_table",
    "qual_theme_analysis",
    "dual_brand_compare",  # backward-compat alias
}

# ---------------------------------------------------------------------------
# Shape type helpers
# ---------------------------------------------------------------------------

SHAPE_TYPE_NAMES = {
    MSO_SHAPE_TYPE.AUTO_SHAPE: "auto_shape",
    MSO_SHAPE_TYPE.CALLOUT: "callout",
    MSO_SHAPE_TYPE.CHART: "chart",
    MSO_SHAPE_TYPE.COMMENT: "comment",
    MSO_SHAPE_TYPE.FREEFORM: "freeform",
    MSO_SHAPE_TYPE.GROUP: "group",
    MSO_SHAPE_TYPE.LINE: "line",
    MSO_SHAPE_TYPE.MEDIA: "media",
    MSO_SHAPE_TYPE.OLE_CONTROL_OBJECT: "ole_control",
    MSO_SHAPE_TYPE.PICTURE: "picture",
    MSO_SHAPE_TYPE.PLACEHOLDER: "placeholder",
    MSO_SHAPE_TYPE.TABLE: "table",
    MSO_SHAPE_TYPE.TEXT_BOX: "text_box",
    MSO_SHAPE_TYPE.TEXT_EFFECT: "text_effect",
}


def shape_type_name(shape) -> str:
    st = shape.shape_type
    if st in SHAPE_TYPE_NAMES:
        return SHAPE_TYPE_NAMES[st]
    return f"other_{st}"


def chart_type_name(chart) -> str:
    """Best-effort chart type identification."""
    try:
        ct = chart.chart_type
        return str(ct).replace("XL_CHART_TYPE.", "").lower()
    except Exception:
        return "unknown"


# ---------------------------------------------------------------------------
# Color / font extractors
# ---------------------------------------------------------------------------


def iter_runs(shape):
    """Yield text runs across a shape (if it has text)."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            yield run


def extract_colors_and_fonts(shape) -> tuple[set[str], set[str], set[int]]:
    """Extract RGB hex colors, font names, and font sizes from a shape."""
    colors: set[str] = set()
    fonts: set[str] = set()
    sizes: set[int] = set()

    if shape.has_text_frame:
        for run in iter_runs(shape):
            try:
                if run.font.name:
                    fonts.add(run.font.name)
            except Exception:
                pass
            try:
                if run.font.size is not None:
                    sizes.add(int(run.font.size.pt))
            except Exception:
                pass
            try:
                color = run.font.color
                if color.type is not None:
                    rgb = getattr(color, "rgb", None)
                    if rgb is not None:
                        colors.add(str(rgb))
            except Exception:
                pass

    # Fill colors (solid fills only, best-effort)
    try:
        fill = shape.fill
        if fill.type is not None and str(fill.type) == "MSO_FILL_TYPE.SOLID (1)":
            rgb = fill.fore_color.rgb
            if rgb is not None:
                colors.add(str(rgb))
    except Exception:
        pass

    return colors, fonts, sizes


# ---------------------------------------------------------------------------
# Slide analysis
# ---------------------------------------------------------------------------


def analyze_chart(chart) -> dict:
    """Extract chart configuration."""
    info: dict[str, Any] = {
        "chart_type": chart_type_name(chart),
        "has_title": chart.has_title,
        "has_legend": chart.has_legend,
    }
    try:
        info["series_count"] = len(list(chart.series))
    except Exception:
        info["series_count"] = 0
    try:
        info["category_count"] = len(chart.plots[0].categories) if chart.plots else 0
    except Exception:
        info["category_count"] = 0
    try:
        info["plot_count"] = len(chart.plots)
    except Exception:
        info["plot_count"] = 0

    # Data label inspection
    try:
        plot = chart.plots[0]
        info["has_data_labels"] = plot.has_data_labels
    except Exception:
        info["has_data_labels"] = False

    return info


def analyze_table(table) -> dict:
    """Extract table structure."""
    try:
        rows = len(table.rows)
        cols = len(table.columns)
    except Exception:
        rows = cols = 0
    return {"rows": rows, "cols": cols}


def analyze_shape(shape) -> dict:
    """Analyze a single shape."""
    info: dict[str, Any] = {
        "name": shape.name,
        "type": shape_type_name(shape),
    }
    try:
        info["left_in"] = round(shape.left / 914400.0, 2) if shape.left is not None else None
        info["top_in"] = round(shape.top / 914400.0, 2) if shape.top is not None else None
        info["width_in"] = round(shape.width / 914400.0, 2) if shape.width is not None else None
        info["height_in"] = round(shape.height / 914400.0, 2) if shape.height is not None else None
    except Exception:
        pass

    colors, fonts, sizes = extract_colors_and_fonts(shape)
    info["colors"] = sorted(colors)
    info["fonts"] = sorted(fonts)
    info["font_sizes"] = sorted(sizes)

    if shape.has_chart:
        info["chart"] = analyze_chart(shape.chart)
    if shape.has_table:
        info["table"] = analyze_table(shape.table)
    if shape.has_text_frame:
        text = shape.text_frame.text[:200]  # first 200 chars
        info["text_sample"] = text.strip()
        info["char_count"] = len(shape.text_frame.text)

    # Recurse into groups
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        try:
            info["group_children"] = [analyze_shape(s) for s in shape.shapes]
        except Exception:
            info["group_children"] = []

    return info


def analyze_slide(slide, index: int) -> dict:
    """Analyze one slide."""
    info: dict[str, Any] = {
        "index": index,
        "layout_name": getattr(slide.slide_layout, "name", "unknown"),
    }

    shape_infos: list[dict] = []
    flat_colors: set[str] = set()
    flat_fonts: set[str] = set()
    flat_sizes: set[int] = set()
    type_counts: Counter = Counter()
    chart_types: list[str] = []
    tables: list[dict] = []

    def walk(shapes):
        for shape in shapes:
            sh = analyze_shape(shape)
            shape_infos.append(sh)
            type_counts[sh["type"]] += 1
            flat_colors.update(sh.get("colors", []))
            flat_fonts.update(sh.get("fonts", []))
            flat_sizes.update(sh.get("font_sizes", []))
            if "chart" in sh:
                chart_types.append(sh["chart"]["chart_type"])
            if "table" in sh:
                tables.append(sh["table"])
            # Walk group children
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    walk(shape.shapes)
                except Exception:
                    pass

    walk(slide.shapes)

    info["shape_count"] = len(shape_infos)
    info["shape_type_counts"] = dict(type_counts)
    info["chart_types_on_slide"] = chart_types
    info["table_count"] = len(tables)
    info["tables"] = tables
    info["colors"] = sorted(flat_colors)
    info["fonts"] = sorted(flat_fonts)
    info["font_sizes"] = sorted(flat_sizes)
    # Raw shape details (truncated if too large — keep under reasonable size)
    info["shapes"] = shape_infos

    # Title detection
    title_shape = next(
        (s for s in slide.shapes if s.has_text_frame and s.is_placeholder),
        None,
    )
    if title_shape is not None:
        info["has_title_placeholder"] = True
        info["title_text"] = title_shape.text_frame.text.strip()[:200]
    else:
        info["has_title_placeholder"] = False

    return info


# ---------------------------------------------------------------------------
# Deck-level analysis
# ---------------------------------------------------------------------------


def infer_client_from_filename(name: str) -> str:
    """Heuristic client inference from filename."""
    lower = name.lower()
    for tag, client in [
        ("jj", "JJ"), ("j&j", "JJ"), ("rybrevant", "JJ"), ("tepezza", "JJ"),
        ("gsk", "GSK"), ("blenrep", "GSK"), ("jemperli", "GSK"), ("ojjaara", "GSK"),
        ("az", "AZN"), ("azn", "AZN"), ("calquence", "AZN"), ("lokelma", "AZN"),
        ("datroway", "DSI"), ("dsi", "DSI"), ("enhertu", "DSI"),
        ("pfizer", "Pfizer"), ("bavencio", "EMD"),
        ("truqap", "AZN"),
        ("lynparza", "AZN"),
        ("libtayo", "Regeneron"),
        ("dupixent", "Regeneron"),
        ("tezspire", "AZN"),
        ("nvs", "Novartis"), ("rhapsido", "Novartis"),
        ("otezla", "Amgen"), ("uplizna", "Amgen"),
        ("adbry", "LEO"),
        ("alexion", "Alexion"),
        ("abrysvo", "Pfizer"),
        ("abilify", "Otsuka"),
        ("apellis", "Apellis"), ("empaveli", "Apellis"),
        ("b+l", "BL"), ("bausch", "BL"),
        ("onivyde", "Ipsen"),
        ("cca", "CCA"),
    ]:
        if tag in lower:
            return client
    return "Unknown"


def analyze_deck(pptx_path: Path) -> dict:
    """Analyze one deck."""
    try:
        prs = Presentation(str(pptx_path))
    except Exception as e:
        return {"filename": pptx_path.name, "error": str(e), "slides": []}

    deck_info: dict[str, Any] = {
        "filename": pptx_path.name,
        "file_size_bytes": pptx_path.stat().st_size,
        "client": infer_client_from_filename(pptx_path.name),
        "slide_count": len(prs.slides),
        "slide_width_in": round(prs.slide_width / 914400.0, 2),
        "slide_height_in": round(prs.slide_height / 914400.0, 2),
    }

    slides: list[dict] = []
    for i, slide in enumerate(prs.slides):
        try:
            slides.append(analyze_slide(slide, i))
        except Exception as e:
            slides.append({"index": i, "error": str(e)})

    deck_info["slides"] = slides

    # Deck-level aggregations
    all_chart_types: Counter = Counter()
    all_shape_types: Counter = Counter()
    all_layouts: Counter = Counter()
    all_colors: Counter = Counter()
    all_fonts: Counter = Counter()

    for s in slides:
        if "error" in s:
            continue
        all_layouts[s.get("layout_name", "unknown")] += 1
        for ct in s.get("chart_types_on_slide", []):
            all_chart_types[ct] += 1
        for stype, count in s.get("shape_type_counts", {}).items():
            all_shape_types[stype] += count
        for c in s.get("colors", []):
            all_colors[c] += 1
        for f in s.get("fonts", []):
            all_fonts[f] += 1

    deck_info["chart_type_counts"] = dict(all_chart_types)
    deck_info["shape_type_counts"] = dict(all_shape_types)
    deck_info["layout_counts"] = dict(all_layouts)
    deck_info["color_counts"] = dict(all_colors)
    deck_info["font_counts"] = dict(all_fonts)

    return deck_info


# ---------------------------------------------------------------------------
# Cross-deck aggregation
# ---------------------------------------------------------------------------


def aggregate(decks: list[dict]) -> dict:
    """Cross-deck aggregations."""
    agg: dict[str, Any] = {
        "deck_count": len(decks),
        "successful_decks": sum(1 for d in decks if "error" not in d),
    }

    chart_types: Counter = Counter()
    shape_types: Counter = Counter()
    layouts: Counter = Counter()
    colors: Counter = Counter()
    fonts: Counter = Counter()
    slide_counts: list[int] = []
    client_counts: Counter = Counter()
    slides_with_charts = 0
    slides_with_tables = 0
    slides_text_only = 0
    total_slides = 0

    # Patterns: shape signatures per slide (e.g., "1_chart,2_table,3_text_box")
    shape_signatures: Counter = Counter()

    for d in decks:
        if "error" in d:
            continue
        client_counts[d.get("client", "Unknown")] += 1
        slide_counts.append(d.get("slide_count", 0))

        for ct, cnt in d.get("chart_type_counts", {}).items():
            chart_types[ct] += cnt
        for st, cnt in d.get("shape_type_counts", {}).items():
            shape_types[st] += cnt
        for ly, cnt in d.get("layout_counts", {}).items():
            layouts[ly] += cnt
        for c, cnt in d.get("color_counts", {}).items():
            colors[c] += cnt
        for f, cnt in d.get("font_counts", {}).items():
            fonts[f] += cnt

        for s in d.get("slides", []):
            if "error" in s:
                continue
            total_slides += 1
            stcs = s.get("shape_type_counts", {})
            if stcs.get("chart", 0) > 0:
                slides_with_charts += 1
            if stcs.get("table", 0) > 0:
                slides_with_tables += 1
            if stcs.get("chart", 0) == 0 and stcs.get("table", 0) == 0:
                slides_text_only += 1

            # Signature: "chart_count_table_count_textbox_count"
            sig = f"charts={stcs.get('chart', 0)},tables={stcs.get('table', 0)},text_boxes={stcs.get('text_box', 0)},pictures={stcs.get('picture', 0)}"
            shape_signatures[sig] += 1

    agg["total_slides"] = total_slides
    agg["slides_with_charts"] = slides_with_charts
    agg["slides_with_tables"] = slides_with_tables
    agg["slides_text_only"] = slides_text_only
    agg["slide_count_distribution"] = {
        "min": min(slide_counts) if slide_counts else 0,
        "max": max(slide_counts) if slide_counts else 0,
        "avg": round(sum(slide_counts) / len(slide_counts), 1) if slide_counts else 0,
    }
    agg["client_counts"] = dict(client_counts.most_common())
    agg["chart_types"] = dict(chart_types.most_common())
    agg["shape_types"] = dict(shape_types.most_common())
    agg["layouts"] = dict(layouts.most_common(50))
    agg["colors_top_40"] = dict(colors.most_common(40))
    agg["fonts"] = dict(fonts.most_common(20))
    agg["shape_signatures_top_20"] = dict(shape_signatures.most_common(20))

    return agg


# ---------------------------------------------------------------------------
# Gap analysis vs. existing 22 renderers
# ---------------------------------------------------------------------------


def gap_analysis(aggregation: dict) -> dict:
    """Map deck chart types to existing renderers; flag gaps."""
    chart_types = aggregation.get("chart_types", {})

    # Rough mapping: deck chart_type -> likely renderer(s) that handle it
    # This is heuristic — a `bar_clustered` could be rendered by several renderers
    chart_to_renderer_candidates = {
        "bar_clustered": ["single_bar_with_delta", "clustered_compare", "dual_bar_with_delta"],
        "bar_stacked": ["stacked_order", "two_section_bar"],
        "bar_stacked_100": ["stacked_order"],
        "column_clustered": ["qoq_bar_with_delta", "dual_bar_qoq", "hii_scorecard"],
        "column_stacked": ["two_section_bar"],
        "column_stacked_100": ["stacked_order"],
        "line": ["trended_scorecard", "trended_activity"],
        "xy_scatter": ["abacus", "dual_abacus", "quadrant_scatter", "followup_rep", "message_mbd"],
        "doughnut": ["dual_doughnut"],
        "pie": ["dual_doughnut"],  # could be used for similar purposes
    }

    covered = {}
    uncovered = {}
    for ct, freq in chart_types.items():
        normalized = ct.replace(" ", "_")
        matched_renderers = []
        for key, renderers in chart_to_renderer_candidates.items():
            if key in normalized:
                matched_renderers = renderers
                break
        if matched_renderers:
            covered[ct] = {"frequency": freq, "candidate_renderers": matched_renderers}
        else:
            uncovered[ct] = freq

    return {
        "covered_chart_types": covered,
        "uncovered_chart_types": uncovered,
        "existing_renderers": sorted(EXISTING_RENDERERS),
        "note": "Mapping is heuristic. A chart of type X being 'covered' by renderer Y means Y uses charts of type X; does NOT mean Y produces the exact layout seen in the deck.",
    }


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------


def write_report_md(decks: list[dict], agg: dict, gap: dict, output: Path) -> None:
    lines: list[str] = []
    lines.append("# Deck Analysis Report")
    lines.append("")
    lines.append(f"**Decks analyzed:** {agg['successful_decks']} successful / {agg['deck_count']} total")
    lines.append(f"**Total slides:** {agg['total_slides']}")
    lines.append(f"**Slide count range:** {agg['slide_count_distribution']['min']}–{agg['slide_count_distribution']['max']} (avg {agg['slide_count_distribution']['avg']})")
    lines.append("")

    lines.append("## Decks Parsed")
    lines.append("")
    lines.append("| # | Filename | Client | Slides | Size (MB) | Status |")
    lines.append("|---|---|---|---|---|---|")
    for i, d in enumerate(decks, 1):
        if "error" in d:
            lines.append(f"| {i} | `{d['filename']}` | — | — | — | ERROR: {d['error']} |")
        else:
            size_mb = round(d.get("file_size_bytes", 0) / 1024 / 1024, 1)
            lines.append(f"| {i} | `{d['filename']}` | {d.get('client', '?')} | {d.get('slide_count', 0)} | {size_mb} | ✓ |")
    lines.append("")

    lines.append("## Clients Represented")
    lines.append("")
    for client, count in agg["client_counts"].items():
        lines.append(f"- **{client}:** {count} deck(s)")
    lines.append("")

    lines.append("## Chart Type Frequency")
    lines.append("")
    lines.append("| Chart Type | Count | Covered? | Candidate Renderers |")
    lines.append("|---|---|---|---|")
    covered = gap["covered_chart_types"]
    uncovered = gap["uncovered_chart_types"]
    for ct, freq in agg["chart_types"].items():
        if ct in covered:
            renderers = ", ".join(f"`{r}`" for r in covered[ct]["candidate_renderers"])
            lines.append(f"| `{ct}` | {freq} | ✓ | {renderers} |")
        else:
            lines.append(f"| `{ct}` | {freq} | ✗ | — |")
    lines.append("")

    if uncovered:
        lines.append("### Uncovered Chart Types (gaps)")
        lines.append("")
        for ct, freq in uncovered.items():
            lines.append(f"- `{ct}` — {freq} occurrences")
        lines.append("")

    lines.append("## Shape Type Frequency")
    lines.append("")
    lines.append("| Shape Type | Count |")
    lines.append("|---|---|")
    for st, freq in agg["shape_types"].items():
        lines.append(f"| `{st}` | {freq} |")
    lines.append("")

    lines.append("## Top Slide Shape Signatures")
    lines.append("")
    lines.append("Recurring shape compositions per slide (tells us what multi-element layouts look like).")
    lines.append("")
    lines.append("| Signature | Count |")
    lines.append("|---|---|")
    for sig, count in agg["shape_signatures_top_20"].items():
        lines.append(f"| `{sig}` | {count} |")
    lines.append("")

    lines.append("## Layout Distribution (top 20)")
    lines.append("")
    lines.append("| Layout Name | Count |")
    lines.append("|---|---|")
    for ly, count in list(agg["layouts"].items())[:20]:
        lines.append(f"| `{ly}` | {count} |")
    lines.append("")

    lines.append("## Colors (top 40 most common)")
    lines.append("")
    lines.append("| Hex | Count |")
    lines.append("|---|---|")
    for c, count in agg["colors_top_40"].items():
        lines.append(f"| `#{c}` | {count} |")
    lines.append("")

    lines.append("## Fonts")
    lines.append("")
    lines.append("| Font | Count |")
    lines.append("|---|---|")
    for f, count in agg["fonts"].items():
        lines.append(f"| `{f}` | {count} |")
    lines.append("")

    lines.append("## Slide Content Type Breakdown")
    lines.append("")
    lines.append(f"- Slides with charts: **{agg['slides_with_charts']}** ({round(agg['slides_with_charts'] / max(agg['total_slides'], 1) * 100)}%)")
    lines.append(f"- Slides with tables: **{agg['slides_with_tables']}** ({round(agg['slides_with_tables'] / max(agg['total_slides'], 1) * 100)}%)")
    lines.append(f"- Text-only slides: **{agg['slides_text_only']}** ({round(agg['slides_text_only'] / max(agg['total_slides'], 1) * 100)}%)")
    lines.append("")

    lines.append("## Current Renderer Inventory")
    lines.append("")
    lines.append(f"SlideGen currently has **{len(EXISTING_RENDERERS)}** renderers:")
    lines.append("")
    for r in sorted(EXISTING_RENDERERS):
        lines.append(f"- `{r}`")
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


def write_gap_md(gap: dict, agg: dict, output: Path) -> None:
    lines: list[str] = []
    lines.append("# Gap Analysis: Decks vs. Existing pptx_utils")
    lines.append("")
    lines.append("This document compares chart types found in real client decks against the 22 existing renderers.")
    lines.append("")
    lines.append("**Caveat:** \"Covered\" here means the renderer uses a chart of this type — it does NOT mean the renderer matches the exact layout seen in the deck. Deeper comparison of layout signatures, callouts, label placement, etc. is a follow-up pass.")
    lines.append("")

    lines.append("## Covered Chart Types")
    lines.append("")
    for ct, info in gap["covered_chart_types"].items():
        lines.append(f"- **`{ct}`** ({info['frequency']} occurrences) → candidates: {', '.join(f'`{r}`' for r in info['candidate_renderers'])}")
    lines.append("")

    lines.append("## Uncovered Chart Types")
    lines.append("")
    if gap["uncovered_chart_types"]:
        for ct, freq in gap["uncovered_chart_types"].items():
            lines.append(f"- **`{ct}`** ({freq} occurrences) — needs new renderer or lxml helper")
    else:
        lines.append("_None — all chart types found in decks map to an existing renderer candidate._")
    lines.append("")

    lines.append("## Recommended Next Steps")
    lines.append("")
    lines.append("1. **Manual review of high-frequency covered types** — do the existing renderers actually produce what the real decks show? If not, refactor them to be more general.")
    lines.append("2. **Uncovered types** — if frequency is high, build new renderers. If low, flag for inline lxml with later extraction.")
    lines.append("3. **Layout signature analysis** — the shape signature counts in `report.md` reveal multi-element compositions (e.g., 1 chart + 2 tables = a common clustered_compare pattern). Map these to `LAYOUTS{}` entries in `pptx_utils`.")
    lines.append("4. **Color palette** — the top-40 colors feed the `BRAND{}` dict. Client-specific hues (orange, purple, etc.) become named entries.")
    lines.append("5. **Font inventory** — if a few fonts dominate, encode defaults in `BRAND{}`.")
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    pptx_files = sorted(p for p in DECKS_DIR.glob("*.pptx") if not p.name.startswith("~"))
    if not pptx_files:
        print(f"No .pptx files found in {DECKS_DIR}", file=sys.stderr)
        return 1

    print(f"Parsing {len(pptx_files)} decks from {DECKS_DIR}...")
    decks: list[dict] = []
    for i, p in enumerate(pptx_files, 1):
        print(f"  [{i}/{len(pptx_files)}] {p.name}", flush=True)
        try:
            decks.append(analyze_deck(p))
        except Exception as e:
            decks.append({"filename": p.name, "error": str(e), "slides": []})
            print(f"    ERROR: {e}", file=sys.stderr)

    print("\nAggregating...")
    agg = aggregate(decks)
    gap = gap_analysis(agg)

    # Write outputs
    (OUTPUTS_DIR / "inventory.json").write_text(
        json.dumps(decks, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "aggregation.json").write_text(
        json.dumps(agg, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "gap_analysis.json").write_text(
        json.dumps(gap, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "chart_types.json").write_text(
        json.dumps(agg["chart_types"], indent=2), encoding="utf-8"
    )
    (OUTPUTS_DIR / "shape_types.json").write_text(
        json.dumps(agg["shape_types"], indent=2), encoding="utf-8"
    )
    (OUTPUTS_DIR / "colors.json").write_text(
        json.dumps(agg["colors_top_40"], indent=2), encoding="utf-8"
    )
    (OUTPUTS_DIR / "fonts.json").write_text(
        json.dumps(agg["fonts"], indent=2), encoding="utf-8"
    )

    write_report_md(decks, agg, gap, OUTPUTS_DIR / "report.md")
    write_gap_md(gap, agg, OUTPUTS_DIR / "gap_analysis.md")

    print(f"\nDone. Outputs in {OUTPUTS_DIR}/")
    print(f"  inventory.json       — {sum(len(d.get('slides', [])) for d in decks)} slides across {agg['successful_decks']} decks")
    print(f"  aggregation.json     — cross-deck aggregations")
    print(f"  gap_analysis.json    — chart coverage vs. 22 renderers")
    print(f"  chart_types.json     — {len(agg['chart_types'])} distinct chart types")
    print(f"  shape_types.json     — {len(agg['shape_types'])} distinct shape types")
    print(f"  colors.json          — top 40 colors")
    print(f"  fonts.json           — {len(agg['fonts'])} fonts")
    print(f"  report.md            — human-readable summary")
    print(f"  gap_analysis.md      — gaps + recommended next steps")

    return 0


if __name__ == "__main__":
    sys.exit(main())
