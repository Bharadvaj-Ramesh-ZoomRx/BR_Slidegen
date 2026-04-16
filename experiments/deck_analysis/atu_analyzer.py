"""
ATU Deck Analyzer
=================

Analyzes the 8 ATU (Awareness, Trial, Usage) decks in decks/ATU/ to inform
the ATU project skill. Produces:

  outputs/atu_inventory.json    — Per-deck structured metadata
  outputs/atu_aggregation.json  — Cross-deck aggregations
  outputs/atu_analysis.md       — Human-readable synthesis + PET comparison

Run:
    cd experiments/deck_analysis
    python atu_analyzer.py
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

HERE = Path(__file__).parent
DECKS_DIR = HERE / "decks" / "ATU"
OUTPUTS_DIR = HERE / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

# ---------------------------------------------------------------------------
# Shape helpers (reused from parser.py pattern)
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
    return SHAPE_TYPE_NAMES.get(st, f"other_{st}")


def chart_type_name(chart) -> str:
    try:
        ct = chart.chart_type
        return str(ct).replace("XL_CHART_TYPE.", "").lower()
    except Exception:
        return "unknown"


# ---------------------------------------------------------------------------
# ATU-specific client inference
# ---------------------------------------------------------------------------

def infer_client_brand(filename: str) -> tuple[str, str]:
    """Infer client and brand from ATU filename."""
    lower = filename.lower()
    mappings = [
        ("romvimza", "Daiichi Sankyo / Petrocelli", "Romvimza"),
        ("dpr", "Daiichi Sankyo / Petrocelli", "DPR"),
        ("filspari", "Travere / Otsuka", "Filspari"),
        ("igan", "Travere / Otsuka", "Filspari"),
        ("prevymis", "Merck", "Prevymis"),
        ("fmi", "Foundation Medicine", "FMI"),
        ("repatha", "Amgen", "Repatha"),
        ("enhertu", "AZ-DSI", "Enhertu"),
        ("otezla", "Amgen", "Otezla"),
        ("gamifant", "SOBI", "Gamifant"),
    ]
    for tag, client, brand in mappings:
        if tag in lower:
            return client, brand
    # Fallback: try to extract from filename patterns
    return "Unknown", "Unknown"


# ---------------------------------------------------------------------------
# Color extraction from chart XML (series colors)
# ---------------------------------------------------------------------------

def extract_chart_series_colors(chart) -> list[str]:
    """Extract series fill colors from chart XML."""
    colors = []
    try:
        chart_xml = chart._chartSpace.xml
        root = ET.fromstring(chart_xml)
        # Look for solidFill colors in series
        for srgb in root.iter(f"{{{NS['a']}}}srgbClr"):
            val = srgb.get("val")
            if val:
                colors.append(val.upper())
    except Exception:
        pass
    return colors


# ---------------------------------------------------------------------------
# Text / font extraction
# ---------------------------------------------------------------------------

def extract_colors_and_fonts(shape) -> tuple[set[str], set[str], set[int]]:
    """Extract RGB hex colors, font names, and font sizes from a shape."""
    colors: set[str] = set()
    fonts: set[str] = set()
    sizes: set[int] = set()

    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
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

    try:
        fill = shape.fill
        if fill.type is not None and "SOLID" in str(fill.type):
            rgb = fill.fore_color.rgb
            if rgb is not None:
                colors.add(str(rgb))
    except Exception:
        pass

    return colors, fonts, sizes


# ---------------------------------------------------------------------------
# Headline extraction
# ---------------------------------------------------------------------------

def extract_headline(slide) -> str | None:
    """Extract headline text from a slide (title placeholder or top text box)."""
    # Try placeholder first
    for shape in slide.shapes:
        if shape.is_placeholder and shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text and len(text) > 5:
                return text[:300]

    # Fallback: topmost text box with substantial text
    text_shapes = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text and len(text) > 10:
                try:
                    top = shape.top if shape.top else 999999999
                    text_shapes.append((top, text[:300]))
                except Exception:
                    pass
    if text_shapes:
        text_shapes.sort(key=lambda x: x[0])
        return text_shapes[0][1]
    return None


# ---------------------------------------------------------------------------
# ATU section detection
# ---------------------------------------------------------------------------

ATU_SECTION_KEYWORDS = {
    "awareness": ["awareness", "aware", "unaided", "aided", "familiarity", "know about"],
    "trial": ["trial", "ever tried", "ever used", "ever prescribed", "trialed"],
    "usage": ["usage", "use", "prescrib", "utiliz", "current use", "currently using", "past 6 months", "past 12 months"],
    "loyalty": ["loyalty", "loyal", "continue", "switch", "retain", "stay on", "adherence"],
    "brand_funnel": ["funnel", "brand funnel", "conversion", "patient journey", "treatment flow"],
    "patient_demographics": ["demographics", "patient profile", "patient characteristics", "age", "gender", "comorbid"],
    "competitive": ["competitive", "competitor", "versus", "vs.", "share", "market share", "competitive landscape"],
    "treatment_journey": ["treatment journey", "line of therapy", "1L", "2L", "3L+", "treatment sequence", "regimen"],
    "satisfaction": ["satisfaction", "satisfied", "dissatisf", "unmet need", "improvement"],
    "perception": ["perception", "perceive", "attitude", "opinion", "view of"],
    "barriers": ["barrier", "concern", "reason for not", "hesitat", "reluctan", "challenge"],
    "drivers": ["driver", "reason for", "motivat", "trigger", "why prescrib"],
    "executive_summary": ["executive summary", "key findings", "key takeaways", "summary of findings"],
    "methodology": ["methodology", "sample", "respondent", "n=", "survey design", "fielding"],
    "recommendations": ["recommend", "implications", "action", "next steps"],
    "cover": ["atu", "awareness trial usage", "full report", "hcp atu", "patient atu"],
}


def classify_slide_section(headline: str | None, all_text: str) -> str:
    """Classify a slide into an ATU section based on headline + body text."""
    if not headline and not all_text:
        return "unknown"

    combined = ((headline or "") + " " + all_text).lower()

    scores: dict[str, int] = {}
    for section, keywords in ATU_SECTION_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in combined)
        if score > 0:
            scores[section] = score

    if scores:
        return max(scores, key=scores.get)
    return "unknown"


# ---------------------------------------------------------------------------
# Slide-level analysis
# ---------------------------------------------------------------------------

def analyze_slide(slide, index: int) -> dict:
    """Analyze one slide comprehensively."""
    info: dict[str, Any] = {
        "index": index,
        "layout_name": getattr(slide.slide_layout, "name", "unknown"),
    }

    type_counts: Counter = Counter()
    chart_types: list[str] = []
    tables: list[dict] = []
    all_colors: set[str] = set()
    all_fonts: set[str] = set()
    all_sizes: set[int] = set()
    chart_series_colors: list[str] = []
    all_text_parts: list[str] = []
    chart_details: list[dict] = []

    def walk(shapes):
        for shape in shapes:
            stype = shape_type_name(shape)
            type_counts[stype] += 1

            colors, fonts, sizes = extract_colors_and_fonts(shape)
            all_colors.update(colors)
            all_fonts.update(fonts)
            all_sizes.update(sizes)

            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    all_text_parts.append(text[:200])

            if shape.has_chart:
                ct = chart_type_name(shape.chart)
                chart_types.append(ct)
                series_colors = extract_chart_series_colors(shape.chart)
                chart_series_colors.extend(series_colors)

                # Chart detail
                detail: dict[str, Any] = {"chart_type": ct}
                try:
                    detail["has_title"] = shape.chart.has_title
                    detail["has_legend"] = shape.chart.has_legend
                    detail["series_count"] = len(list(shape.chart.series))
                except Exception:
                    pass
                try:
                    detail["category_count"] = len(shape.chart.plots[0].categories) if shape.chart.plots else 0
                except Exception:
                    detail["category_count"] = 0
                try:
                    plot = shape.chart.plots[0]
                    detail["has_data_labels"] = plot.has_data_labels
                except Exception:
                    detail["has_data_labels"] = False
                chart_details.append(detail)

            if shape.has_table:
                try:
                    rows = len(shape.table.rows)
                    cols = len(shape.table.columns)
                    tables.append({"rows": rows, "cols": cols})
                except Exception:
                    tables.append({"rows": 0, "cols": 0})

            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    walk(shape.shapes)
                except Exception:
                    pass

    walk(slide.shapes)

    headline = extract_headline(slide)
    info["headline"] = headline
    info["shape_type_counts"] = dict(type_counts)
    info["chart_types"] = chart_types
    info["chart_count"] = len(chart_types)
    info["chart_details"] = chart_details
    info["table_count"] = len(tables)
    info["tables"] = tables
    info["colors"] = sorted(all_colors)
    info["fonts"] = sorted(all_fonts)
    info["font_sizes"] = sorted(all_sizes)
    info["chart_series_colors"] = sorted(set(chart_series_colors))

    # Slide composition signature
    sig = f"{len(chart_types)}_chart_{len(tables)}_table"
    info["composition_signature"] = sig

    # Section classification
    body_text = " ".join(all_text_parts[:10])  # first 10 text snippets
    info["section"] = classify_slide_section(headline, body_text)

    # Chrome: has title, has legend, has gridlines, has data labels
    info["has_title_placeholder"] = any(
        s.is_placeholder and s.has_text_frame
        for s in slide.shapes
    )

    return info


# ---------------------------------------------------------------------------
# Deck-level analysis
# ---------------------------------------------------------------------------

def analyze_deck(pptx_path: Path) -> dict:
    """Analyze one ATU deck."""
    try:
        prs = Presentation(str(pptx_path))
    except Exception as e:
        return {"filename": pptx_path.name, "error": str(e), "slides": []}

    client, brand = infer_client_brand(pptx_path.name)

    deck_info: dict[str, Any] = {
        "filename": pptx_path.name,
        "file_size_mb": round(pptx_path.stat().st_size / 1024 / 1024, 1),
        "client": client,
        "brand": brand,
        "slide_count": len(prs.slides),
        "slide_width_in": round(prs.slide_width / 914400.0, 2),
        "slide_height_in": round(prs.slide_height / 914400.0, 2),
    }

    slides: list[dict] = []
    headlines: list[str] = []
    for i, slide in enumerate(prs.slides):
        try:
            slide_info = analyze_slide(slide, i)
            slides.append(slide_info)
            if slide_info.get("headline"):
                headlines.append(slide_info["headline"])
        except Exception as e:
            slides.append({"index": i, "error": str(e)})

    deck_info["slides"] = slides
    deck_info["headlines_first_15"] = headlines[:15]

    # Deck-level aggregations
    chart_type_counts: Counter = Counter()
    shape_type_counts: Counter = Counter()
    font_counts: Counter = Counter()
    color_counts: Counter = Counter()
    series_color_counts: Counter = Counter()
    section_counts: Counter = Counter()
    composition_counts: Counter = Counter()
    layout_counts: Counter = Counter()
    total_charts = 0
    total_tables = 0

    for s in slides:
        if "error" in s:
            continue
        for ct in s.get("chart_types", []):
            chart_type_counts[ct] += 1
            total_charts += 1
        total_tables += s.get("table_count", 0)
        for st, cnt in s.get("shape_type_counts", {}).items():
            shape_type_counts[st] += cnt
        for f in s.get("fonts", []):
            font_counts[f] += 1
        for c in s.get("colors", []):
            color_counts[c] += 1
        for c in s.get("chart_series_colors", []):
            series_color_counts[c] += 1
        section_counts[s.get("section", "unknown")] += 1
        composition_counts[s.get("composition_signature", "unknown")] += 1
        layout_counts[s.get("layout_name", "unknown")] += 1

    deck_info["total_charts"] = total_charts
    deck_info["total_tables"] = total_tables
    deck_info["chart_type_counts"] = dict(chart_type_counts.most_common())
    deck_info["shape_type_counts"] = dict(shape_type_counts.most_common())
    deck_info["font_counts"] = dict(font_counts.most_common(15))
    deck_info["color_counts"] = dict(color_counts.most_common(20))
    deck_info["series_color_counts"] = dict(series_color_counts.most_common(20))
    deck_info["section_counts"] = dict(section_counts.most_common())
    deck_info["composition_counts"] = dict(composition_counts.most_common(15))
    deck_info["layout_counts"] = dict(layout_counts.most_common(10))

    return deck_info


# ---------------------------------------------------------------------------
# Cross-deck aggregation
# ---------------------------------------------------------------------------

def aggregate(decks: list[dict]) -> dict:
    """Cross-deck aggregations for ATU."""
    agg: dict[str, Any] = {
        "deck_count": len(decks),
        "successful_decks": sum(1 for d in decks if "error" not in d),
    }

    chart_types: Counter = Counter()
    shape_types: Counter = Counter()
    fonts: Counter = Counter()
    colors: Counter = Counter()
    series_colors: Counter = Counter()
    sections: Counter = Counter()
    compositions: Counter = Counter()
    layouts: Counter = Counter()
    clients: Counter = Counter()
    brands: list[dict] = []

    total_slides = 0
    total_charts = 0
    total_tables = 0
    slides_with_charts = 0
    slides_with_tables = 0
    slides_text_only = 0
    slide_counts: list[int] = []

    # Chrome stats
    slides_with_title = 0
    charts_with_legend = 0
    charts_with_data_labels = 0
    charts_with_title = 0
    total_chart_details = 0

    all_headlines: list[str] = []

    for d in decks:
        if "error" in d:
            continue

        clients[d.get("client", "Unknown")] += 1
        brands.append({"client": d.get("client", "?"), "brand": d.get("brand", "?"), "filename": d.get("filename", "?")})
        slide_counts.append(d.get("slide_count", 0))
        total_charts += d.get("total_charts", 0)
        total_tables += d.get("total_tables", 0)

        for ct, cnt in d.get("chart_type_counts", {}).items():
            chart_types[ct] += cnt
        for st, cnt in d.get("shape_type_counts", {}).items():
            shape_types[st] += cnt
        for f, cnt in d.get("font_counts", {}).items():
            fonts[f] += cnt
        for c, cnt in d.get("color_counts", {}).items():
            colors[c] += cnt
        for c, cnt in d.get("series_color_counts", {}).items():
            series_colors[c] += cnt
        for sec, cnt in d.get("section_counts", {}).items():
            sections[sec] += cnt
        for comp, cnt in d.get("composition_counts", {}).items():
            compositions[comp] += cnt
        for ly, cnt in d.get("layout_counts", {}).items():
            layouts[ly] += cnt

        all_headlines.extend(d.get("headlines_first_15", []))

        for s in d.get("slides", []):
            if "error" in s:
                continue
            total_slides += 1
            stcs = s.get("shape_type_counts", {})
            has_chart = stcs.get("chart", 0) > 0
            has_table = stcs.get("table", 0) > 0

            if has_chart:
                slides_with_charts += 1
            if has_table:
                slides_with_tables += 1
            if not has_chart and not has_table:
                slides_text_only += 1
            if s.get("has_title_placeholder"):
                slides_with_title += 1

            for cd in s.get("chart_details", []):
                total_chart_details += 1
                if cd.get("has_legend"):
                    charts_with_legend += 1
                if cd.get("has_data_labels"):
                    charts_with_data_labels += 1
                if cd.get("has_title"):
                    charts_with_title += 1

    agg["total_slides"] = total_slides
    agg["total_charts"] = total_charts
    agg["total_tables"] = total_tables
    agg["slides_with_charts"] = slides_with_charts
    agg["slides_with_tables"] = slides_with_tables
    agg["slides_text_only"] = slides_text_only
    agg["slide_count_distribution"] = {
        "min": min(slide_counts) if slide_counts else 0,
        "max": max(slide_counts) if slide_counts else 0,
        "avg": round(sum(slide_counts) / len(slide_counts), 1) if slide_counts else 0,
        "total": sum(slide_counts),
    }

    # Chrome percentages
    agg["chrome"] = {
        "title_pct": round(slides_with_title / max(total_slides, 1) * 100, 1),
        "legend_pct": round(charts_with_legend / max(total_chart_details, 1) * 100, 1),
        "data_label_pct": round(charts_with_data_labels / max(total_chart_details, 1) * 100, 1),
        "chart_title_pct": round(charts_with_title / max(total_chart_details, 1) * 100, 1),
    }

    agg["client_counts"] = dict(clients.most_common())
    agg["brands"] = brands
    agg["chart_types"] = dict(chart_types.most_common())
    agg["shape_types"] = dict(shape_types.most_common())
    agg["fonts_top_20"] = dict(fonts.most_common(20))
    agg["colors_top_40"] = dict(colors.most_common(40))
    agg["series_colors_top_30"] = dict(series_colors.most_common(30))
    agg["section_counts"] = dict(sections.most_common())
    agg["composition_signatures"] = dict(compositions.most_common(20))
    agg["layouts_top_20"] = dict(layouts.most_common(20))

    return agg


# ---------------------------------------------------------------------------
# PET comparison
# ---------------------------------------------------------------------------

PET_STATS = {
    "deck_count": 32,
    "total_slides": 2333,
    "avg_slides_per_deck": 72.9,
    "total_charts": 4350,
    "slides_with_charts_pct": 70,
    "slides_with_tables_pct": 71,
    "text_only_pct": 19,
    "top_chart_types": {
        "bar_clustered": 1511,
        "xy_scatter": 674,
        "line_markers": 563,
        "column_stacked_100": 486,
        "bar_stacked_100": 324,
        "bar_stacked": 305,
        "doughnut": 138,
        "column_clustered": 118,
        "column_stacked": 98,
    },
    "top_fonts": ["Arial", "Century Gothic", "Calibri", "Johnson Text"],
}


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def write_analysis_md(decks: list[dict], agg: dict, output: Path) -> None:
    """Write the comprehensive ATU analysis synthesis."""
    lines: list[str] = []

    # Header
    lines.append("# ATU Deck Analysis")
    lines.append("")
    lines.append(f"**Decks analyzed:** {agg['successful_decks']} / {agg['deck_count']}")
    lines.append(f"**Total slides:** {agg['total_slides']}")
    lines.append(f"**Total charts:** {agg['total_charts']}")
    lines.append(f"**Total tables:** {agg['total_tables']}")
    lines.append(f"**Slide count range:** {agg['slide_count_distribution']['min']}-{agg['slide_count_distribution']['max']} (avg {agg['slide_count_distribution']['avg']})")
    lines.append("")

    # Brand coverage
    lines.append("## Brand Coverage")
    lines.append("")
    lines.append("| # | Filename | Client | Brand | Slides | Size (MB) |")
    lines.append("|---|---|---|---|---|---|")
    for i, d in enumerate(decks, 1):
        if "error" in d:
            lines.append(f"| {i} | `{d['filename']}` | - | - | - | ERROR |")
        else:
            lines.append(f"| {i} | `{d['filename']}` | {d.get('client', '?')} | {d.get('brand', '?')} | {d.get('slide_count', 0)} | {d.get('file_size_mb', 0)} |")
    lines.append("")
    lines.append("**Clients represented:**")
    for client, count in agg["client_counts"].items():
        lines.append(f"- **{client}:** {count} deck(s)")
    lines.append("")

    # Key numbers
    lines.append("## Key Numbers")
    lines.append("")
    total_slides = agg["total_slides"]
    lines.append(f"- Total slides analyzed: **{total_slides}**")
    lines.append(f"- Total charts: **{agg['total_charts']}**")
    lines.append(f"- Total tables: **{agg['total_tables']}**")
    lines.append(f"- Charts per slide: **{round(agg['total_charts'] / max(total_slides, 1), 2)}**")
    lines.append(f"- Tables per slide: **{round(agg['total_tables'] / max(total_slides, 1), 2)}**")
    lines.append(f"- Slides with charts: **{agg['slides_with_charts']}** ({round(agg['slides_with_charts'] / max(total_slides, 1) * 100)}%)")
    lines.append(f"- Slides with tables: **{agg['slides_with_tables']}** ({round(agg['slides_with_tables'] / max(total_slides, 1) * 100)}%)")
    lines.append(f"- Text-only slides: **{agg['slides_text_only']}** ({round(agg['slides_text_only'] / max(total_slides, 1) * 100)}%)")
    lines.append("")

    # Chrome defaults
    chrome = agg.get("chrome", {})
    lines.append("### Chrome Defaults")
    lines.append("")
    lines.append(f"- Title placeholder present: **{chrome.get('title_pct', 0)}%** of slides")
    lines.append(f"- Chart has legend: **{chrome.get('legend_pct', 0)}%** of charts")
    lines.append(f"- Chart has data labels: **{chrome.get('data_label_pct', 0)}%** of charts")
    lines.append(f"- Chart has title: **{chrome.get('chart_title_pct', 0)}%** of charts")
    lines.append("")

    # Chart types
    lines.append("## Chart Type Frequency")
    lines.append("")
    total_charts = agg["total_charts"]
    lines.append("| Chart Type | Count | % of Total |")
    lines.append("|---|---|---|")
    for ct, count in agg["chart_types"].items():
        pct = round(count / max(total_charts, 1) * 100, 1)
        lines.append(f"| `{ct}` | {count} | {pct}% |")
    lines.append("")

    # Composition signatures
    lines.append("## Slide Composition Signatures")
    lines.append("")
    lines.append("| Signature | Count | % of Slides |")
    lines.append("|---|---|---|")
    for sig, count in agg["composition_signatures"].items():
        pct = round(count / max(total_slides, 1) * 100, 1)
        lines.append(f"| `{sig}` | {count} | {pct}% |")
    lines.append("")

    # Section distribution
    lines.append("## Section Distribution (auto-classified)")
    lines.append("")
    lines.append("| Section | Slide Count | % |")
    lines.append("|---|---|---|")
    for sec, count in agg["section_counts"].items():
        pct = round(count / max(total_slides, 1) * 100, 1)
        lines.append(f"| {sec} | {count} | {pct}% |")
    lines.append("")

    # Fonts
    lines.append("## Font Usage")
    lines.append("")
    lines.append("| Font | Occurrences |")
    lines.append("|---|---|")
    for f, count in agg["fonts_top_20"].items():
        lines.append(f"| `{f}` | {count} |")
    lines.append("")

    # Colors (text + fill)
    lines.append("## Colors (text + fill, top 30)")
    lines.append("")
    lines.append("| Hex | Occurrences |")
    lines.append("|---|---|")
    for c, count in list(agg["colors_top_40"].items())[:30]:
        lines.append(f"| `#{c}` | {count} |")
    lines.append("")

    # Series colors (from chart XML)
    lines.append("## Chart Series Colors (from chart XML, top 20)")
    lines.append("")
    lines.append("| Hex | Occurrences |")
    lines.append("|---|---|")
    for c, count in list(agg["series_colors_top_30"].items())[:20]:
        lines.append(f"| `#{c}` | {count} |")
    lines.append("")

    # Headlines
    lines.append("## Sample Headlines (first 10 per deck)")
    lines.append("")
    for d in decks:
        if "error" in d:
            continue
        lines.append(f"### {d.get('brand', '?')} ({d.get('client', '?')})")
        lines.append("")
        for h in d.get("headlines_first_15", [])[:10]:
            clean = h.replace("\n", " ").strip()[:150]
            lines.append(f"- {clean}")
        lines.append("")

    # Layout distribution
    lines.append("## Layout Distribution (top 20)")
    lines.append("")
    lines.append("| Layout Name | Count |")
    lines.append("|---|---|")
    for ly, count in agg["layouts_top_20"].items():
        lines.append(f"| `{ly}` | {count} |")
    lines.append("")

    # ===================================================================
    # ATU vs PET Comparison
    # ===================================================================
    lines.append("---")
    lines.append("")
    lines.append("# ATU vs PET Comparison")
    lines.append("")

    # Structural comparison
    atu_avg = agg["slide_count_distribution"]["avg"]
    pet_avg = PET_STATS["avg_slides_per_deck"]
    atu_charts_pct = round(agg["slides_with_charts"] / max(total_slides, 1) * 100)
    atu_tables_pct = round(agg["slides_with_tables"] / max(total_slides, 1) * 100)
    atu_text_pct = round(agg["slides_text_only"] / max(total_slides, 1) * 100)

    lines.append("## Structural Comparison")
    lines.append("")
    lines.append("| Metric | ATU (8 decks) | PET (32 decks) |")
    lines.append("|---|---|---|")
    lines.append(f"| Avg slides per deck | {atu_avg} | {pet_avg} |")
    lines.append(f"| Total slides | {total_slides} | {PET_STATS['total_slides']} |")
    lines.append(f"| Total charts | {agg['total_charts']} | {PET_STATS['total_charts']} |")
    lines.append(f"| Charts per slide | {round(agg['total_charts'] / max(total_slides, 1), 2)} | {round(PET_STATS['total_charts'] / PET_STATS['total_slides'], 2)} |")
    lines.append(f"| Slides with charts | {atu_charts_pct}% | {PET_STATS['slides_with_charts_pct']}% |")
    lines.append(f"| Slides with tables | {atu_tables_pct}% | {PET_STATS['slides_with_tables_pct']}% |")
    lines.append(f"| Text-only slides | {atu_text_pct}% | {PET_STATS['text_only_pct']}% |")
    lines.append("")

    # Chart type comparison
    lines.append("## Chart Type Comparison")
    lines.append("")
    lines.append("| Chart Type | ATU Count | ATU % | PET Count | PET % |")
    lines.append("|---|---|---|---|---|")
    all_ct_names = set()
    for ct in agg["chart_types"]:
        # Normalize to base type for comparison
        all_ct_names.add(ct)
    for ct in PET_STATS["top_chart_types"]:
        all_ct_names.add(ct)

    # Match ATU chart type names to PET names
    atu_ct = agg["chart_types"]
    pet_ct = PET_STATS["top_chart_types"]
    for ct_name in sorted(atu_ct.keys(), key=lambda x: atu_ct.get(x, 0), reverse=True):
        atu_count = atu_ct.get(ct_name, 0)
        atu_pct = round(atu_count / max(agg["total_charts"], 1) * 100, 1)
        # Try to find matching PET type
        pet_match = 0
        base = ct_name.split(" (")[0] if " (" in ct_name else ct_name
        for pk, pv in pet_ct.items():
            if base in pk or pk in base:
                pet_match = pv
                break
        pet_pct = round(pet_match / max(PET_STATS["total_charts"], 1) * 100, 1) if pet_match else "-"
        lines.append(f"| `{ct_name}` | {atu_count} | {atu_pct}% | {pet_match or '-'} | {pet_pct}% |")
    lines.append("")

    # ATU-universal patterns
    lines.append("## ATU-Universal Patterns")
    lines.append("")
    lines.append("Sections detected across all 8 ATU decks (auto-classified from headlines + body text):")
    lines.append("")

    # Check which sections appear in how many decks
    section_deck_coverage: dict[str, int] = defaultdict(int)
    for d in decks:
        if "error" in d:
            continue
        deck_sections = set()
        for s in d.get("slides", []):
            sec = s.get("section", "unknown")
            if sec != "unknown":
                deck_sections.add(sec)
        for sec in deck_sections:
            section_deck_coverage[sec] += 1

    lines.append("| Section | Decks Present In | Total Slides | Universal? |")
    lines.append("|---|---|---|---|")
    n_decks = agg["successful_decks"]
    for sec, deck_count in sorted(section_deck_coverage.items(), key=lambda x: -x[1]):
        slide_count = agg["section_counts"].get(sec, 0)
        universal = "Yes" if deck_count >= n_decks * 0.75 else "Partial" if deck_count >= n_decks * 0.5 else "No"
        lines.append(f"| **{sec}** | {deck_count}/{n_decks} | {slide_count} | {universal} |")
    lines.append("")

    # ATU-specific chart patterns
    lines.append("## ATU-Specific Chart Patterns")
    lines.append("")
    lines.append("Chart types in ATU decks that differ in frequency or presence from PET:")
    lines.append("")

    # Identify ATU-specific patterns
    atu_dominant = sorted(atu_ct.items(), key=lambda x: -x[1])
    for ct_name, count in atu_dominant[:5]:
        pct = round(count / max(agg["total_charts"], 1) * 100, 1)
        lines.append(f"- **`{ct_name}`**: {count} ({pct}%) - ", )
        # Check if this is higher or lower than PET
        base = ct_name.split(" (")[0]
        pet_match_pct = 0
        for pk, pv in pet_ct.items():
            if base in pk:
                pet_match_pct = round(pv / PET_STATS["total_charts"] * 100, 1)
                break
        if pet_match_pct:
            if pct > pet_match_pct + 5:
                lines[-1] += f"**higher than PET** ({pet_match_pct}%)"
            elif pct < pet_match_pct - 5:
                lines[-1] += f"**lower than PET** ({pet_match_pct}%)"
            else:
                lines[-1] += f"similar to PET ({pet_match_pct}%)"
        else:
            lines[-1] += "not prominent in PET"
    lines.append("")

    # Per-deck summary table
    lines.append("## Per-Deck Summary")
    lines.append("")
    lines.append("| Deck | Slides | Charts | Tables | Top Chart Type | Top Section |")
    lines.append("|---|---|---|---|---|---|")
    for d in decks:
        if "error" in d:
            continue
        top_ct = max(d.get("chart_type_counts", {"none": 0}).items(), key=lambda x: x[1], default=("none", 0))
        top_sec = max(d.get("section_counts", {"unknown": 0}).items(), key=lambda x: x[1], default=("unknown", 0))
        lines.append(f"| {d.get('brand', '?')} | {d.get('slide_count', 0)} | {d.get('total_charts', 0)} | {d.get('total_tables', 0)} | `{top_ct[0]}` ({top_ct[1]}) | {top_sec[0]} ({top_sec[1]}) |")
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    pptx_files = sorted(
        p for p in DECKS_DIR.glob("*.pptx") if not p.name.startswith("~")
    )
    if not pptx_files:
        print(f"No .pptx files found in {DECKS_DIR}", file=sys.stderr)
        return 1

    print(f"Analyzing {len(pptx_files)} ATU decks from {DECKS_DIR}...")
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

    # Write outputs
    inv_path = OUTPUTS_DIR / "atu_inventory.json"
    agg_path = OUTPUTS_DIR / "atu_aggregation.json"
    md_path = OUTPUTS_DIR / "atu_analysis.md"

    inv_path.write_text(
        json.dumps(decks, indent=2, default=str), encoding="utf-8"
    )
    agg_path.write_text(
        json.dumps(agg, indent=2, default=str), encoding="utf-8"
    )
    write_analysis_md(decks, agg, md_path)

    print(f"\nDone. Outputs:")
    print(f"  {inv_path}  — per-deck inventory")
    print(f"  {agg_path}  — cross-deck aggregation")
    print(f"  {md_path}   — synthesis + PET comparison")
    print(f"\n  {agg['total_slides']} slides, {agg['total_charts']} charts, {agg['total_tables']} tables across {agg['successful_decks']} decks")

    return 0


if __name__ == "__main__":
    sys.exit(main())
