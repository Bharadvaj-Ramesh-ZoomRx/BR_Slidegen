"""
Deep Deck Analyzer
==================

Goes beyond the first-pass parser. Extracts:

  1. Per-chart OOXML detail — gap_width, overlap, axis orientation, label positions,
     series formatting (fill, border, markers), data label formats
  2. Non-python-pptx properties touched — the inventory for lxml_helpers.py
  3. Per-client brand extraction — colors + fonts grouped by client with proposals
     for BRAND{} entries
  4. Layout coordinate analysis — for top shape signatures, dump full coordinate
     grids to identify LAYOUTS{} patterns
  5. Talking header patterns — extract top-of-slide text_boxes, font sizes, colors,
     lengths — informs headline skill + pptx_utils text helpers
  6. Table structure patterns — row/column counts, cell fills, common dimensions
  7. Title text length distribution
  8. Position clustering — which grid positions recur most often

Run:
    cd experiments/deck_analysis
    python deep_analyzer.py

Outputs:
    outputs/deep_ooxml_properties.json      — every OOXML property touched
    outputs/deep_chart_details.json         — per-chart details across all decks
    outputs/deep_chart_summary.md           — human-readable chart config summary
    outputs/deep_brand_proposals.json       — BRAND{} proposals per client
    outputs/deep_brand_proposals.md         — brand proposals (readable)
    outputs/deep_layouts.json               — LAYOUTS{} candidates from coord clusters
    outputs/deep_layouts.md                 — layout patterns (readable)
    outputs/deep_headlines.json             — headline formatting patterns
    outputs/deep_tables.json                — table structure patterns
    outputs/deep_report.md                  — master human-readable deep report
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean, median, stdev
from typing import Any, Iterable
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

HERE = Path(__file__).parent
DECKS_DIR = HERE / "decks"
OUTPUTS_DIR = HERE / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

# OOXML namespaces
NS = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c":   "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p":   "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# ---------------------------------------------------------------------------
# Client inference (reused from parser.py, expanded)
# ---------------------------------------------------------------------------

CLIENT_RULES = [
    (re.compile(r"j&j|rybrevant|tepezza|ojjaara", re.I), "JJ"),
    (re.compile(r"jj\b", re.I), "JJ"),
    (re.compile(r"gsk|blenrep|jemperli", re.I), "GSK"),
    (re.compile(r"az[-_ ]|azn|calquence|lokelma|lynparza|truqap|tezspire", re.I), "AZN"),
    (re.compile(r"datroway|dsi\b|enhertu", re.I), "DSI"),
    (re.compile(r"pfizer|abrysvo|bavencio", re.I), "Pfizer"),
    (re.compile(r"libtayo|dupixent", re.I), "Regeneron"),
    (re.compile(r"nvs|rhapsido", re.I), "Novartis"),
    (re.compile(r"otezla|uplizna", re.I), "Amgen"),
    (re.compile(r"adbry", re.I), "LEO"),
    (re.compile(r"alexion", re.I), "Alexion"),
    (re.compile(r"abilify", re.I), "Otsuka"),
    (re.compile(r"apellis|empaveli", re.I), "Apellis"),
    (re.compile(r"b\+l|bausch", re.I), "BL"),
    (re.compile(r"onivyde", re.I), "Ipsen"),
    (re.compile(r"cca\b", re.I), "CCA"),
    (re.compile(r"bone hcp", re.I), "Bone-HCP"),
]


def infer_client(filename: str) -> str:
    for rx, client in CLIENT_RULES:
        if rx.search(filename):
            return client
    return "Unknown"


# ---------------------------------------------------------------------------
# OOXML property inventory — parse raw chart XML
# ---------------------------------------------------------------------------

# Properties we know python-pptx exposes (we EXCLUDE these from the gap list)
PYTHON_PPTX_COVERS = {
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}title",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}tx",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}numFmt",  # we use this via helper
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr",    # shape properties general
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}txPr",    # text properties general
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}chart",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}printSettings",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}externalData",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}lang",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}roundedCorners",
    "{http://schemas.openxmlformats.org/drawingml/2006/chart}autoTitleDeleted",
}


def strip_ns(tag: str) -> str:
    """Return local tag name without namespace."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def qn_c(local: str) -> str:
    return f"{{{NS['c']}}}{local}"


def qn_a(local: str) -> str:
    return f"{{{NS['a']}}}{local}"


# ---------------------------------------------------------------------------
# Chart deep analysis
# ---------------------------------------------------------------------------


def analyze_chart_xml(chart_xml: bytes) -> dict:
    """Extract deep chart properties from raw XML bytes."""
    try:
        root = ET.fromstring(chart_xml)
    except ET.ParseError as e:
        return {"error": f"xml_parse_error: {e}"}

    info: dict[str, Any] = {
        "chart_types": [],
        "gap_widths": [],
        "overlaps": [],
        "data_label_positions": [],
        "series_colors": [],
        "axis_orientations": [],
        "cat_axis_visible": None,
        "val_axis_number_formats": [],
        "has_title": False,
        "series_count": 0,
        "has_legend": False,
        "legend_position": None,
        "tick_lbl_pos": [],
        "major_gridlines_present": False,
        "minor_gridlines_present": False,
        "num_formats_used": [],
        "series_has_smooth": [],
        "series_marker_types": [],
        "series_line_widths": [],
        "series_border_settings": [],
        "axis_line_settings": [],
        "trendlines": [],
        "error_bars": [],
    }

    # chart type — look under plotArea
    plot_area = root.find(f".//{qn_c('plotArea')}")
    if plot_area is None:
        return info

    # Chart type tags
    for tag in ("barChart", "lineChart", "pieChart", "doughnutChart", "scatterChart",
                "bubbleChart", "areaChart", "radarChart", "surface3DChart", "stockChart"):
        elems = plot_area.findall(qn_c(tag))
        for el in elems:
            # Check barDir for bar vs column
            bar_dir = el.find(qn_c("barDir"))
            grouping = el.find(qn_c("grouping"))
            chart_type_info = {
                "tag": tag,
                "bar_dir": bar_dir.get("val") if bar_dir is not None else None,
                "grouping": grouping.get("val") if grouping is not None else None,
            }
            info["chart_types"].append(chart_type_info)

            # gapWidth + overlap (bar/column charts)
            gw = el.find(qn_c("gapWidth"))
            if gw is not None:
                info["gap_widths"].append(gw.get("val"))
            ov = el.find(qn_c("overlap"))
            if ov is not None:
                info["overlaps"].append(ov.get("val"))

            # dLbls (data labels at plot level)
            dlbls = el.find(qn_c("dLbls"))
            if dlbls is not None:
                pos = dlbls.find(qn_c("dLblPos"))
                if pos is not None:
                    info["data_label_positions"].append(pos.get("val"))

            # series
            for ser in el.findall(qn_c("ser")):
                info["series_count"] += 1
                # series color (fill)
                sp_pr = ser.find(qn_c("spPr"))
                if sp_pr is not None:
                    solid = sp_pr.find(f"{qn_a('solidFill')}/{qn_a('srgbClr')}")
                    if solid is not None:
                        info["series_colors"].append(solid.get("val"))
                    # line/border
                    ln = sp_pr.find(qn_a("ln"))
                    if ln is not None:
                        w = ln.get("w")
                        no_fill = ln.find(qn_a("noFill")) is not None
                        info["series_border_settings"].append({
                            "width_emu": w,
                            "no_fill": no_fill,
                        })
                        info["series_line_widths"].append(w)
                # series-level data labels
                ser_dlbls = ser.find(qn_c("dLbls"))
                if ser_dlbls is not None:
                    pos = ser_dlbls.find(qn_c("dLblPos"))
                    if pos is not None:
                        info["data_label_positions"].append(pos.get("val"))
                    nf = ser_dlbls.find(qn_c("numFmt"))
                    if nf is not None:
                        info["num_formats_used"].append(nf.get("formatCode"))
                # smooth
                smooth = ser.find(qn_c("smooth"))
                if smooth is not None:
                    info["series_has_smooth"].append(smooth.get("val"))
                # marker
                marker = ser.find(qn_c("marker"))
                if marker is not None:
                    sym = marker.find(qn_c("symbol"))
                    if sym is not None:
                        info["series_marker_types"].append(sym.get("val"))
                # trendline
                tl = ser.find(qn_c("trendline"))
                if tl is not None:
                    t_type = tl.find(qn_c("trendlineType"))
                    info["trendlines"].append(t_type.get("val") if t_type is not None else "unknown")
                # errBars
                eb = ser.find(qn_c("errBars"))
                if eb is not None:
                    info["error_bars"].append("present")

    # Axes
    for cat_ax in plot_area.findall(qn_c("catAx")):
        # orientation
        scaling = cat_ax.find(qn_c("scaling"))
        if scaling is not None:
            orient = scaling.find(qn_c("orientation"))
            if orient is not None:
                info["axis_orientations"].append({
                    "axis": "cat",
                    "orientation": orient.get("val"),
                })
        # tick label position
        tlp = cat_ax.find(qn_c("tickLblPos"))
        if tlp is not None:
            info["tick_lbl_pos"].append({"axis": "cat", "val": tlp.get("val")})
        # axis line
        sp_pr = cat_ax.find(qn_c("spPr"))
        if sp_pr is not None:
            ln = sp_pr.find(qn_a("ln"))
            if ln is not None:
                info["axis_line_settings"].append({
                    "axis": "cat",
                    "width_emu": ln.get("w"),
                    "no_fill": ln.find(qn_a("noFill")) is not None,
                })

    for val_ax in plot_area.findall(qn_c("valAx")):
        scaling = val_ax.find(qn_c("scaling"))
        if scaling is not None:
            orient = scaling.find(qn_c("orientation"))
            if orient is not None:
                info["axis_orientations"].append({
                    "axis": "val",
                    "orientation": orient.get("val"),
                })
        tlp = val_ax.find(qn_c("tickLblPos"))
        if tlp is not None:
            info["tick_lbl_pos"].append({"axis": "val", "val": tlp.get("val")})
        nf = val_ax.find(qn_c("numFmt"))
        if nf is not None:
            info["val_axis_number_formats"].append(nf.get("formatCode"))
        mg = val_ax.find(qn_c("majorGridlines"))
        if mg is not None:
            info["major_gridlines_present"] = True
        ming = val_ax.find(qn_c("minorGridlines"))
        if ming is not None:
            info["minor_gridlines_present"] = True

    # Title
    if root.find(f".//{qn_c('title')}") is not None:
        info["has_title"] = True

    # Legend
    legend = root.find(f".//{qn_c('legend')}")
    if legend is not None:
        info["has_legend"] = True
        pos = legend.find(qn_c("legendPos"))
        if pos is not None:
            info["legend_position"] = pos.get("val")

    return info


def walk_element_tags(element, counter: Counter):
    """Recursively count all (namespace, localname) tag occurrences."""
    counter[element.tag] += 1
    for child in element:
        walk_element_tags(child, counter)


# ---------------------------------------------------------------------------
# Extract all chart XML from a .pptx
# ---------------------------------------------------------------------------


def extract_chart_xmls(pptx_path: Path) -> list[tuple[str, bytes]]:
    """Return list of (chart_filename, xml_bytes) tuples."""
    results: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        for name in z.namelist():
            if name.startswith("ppt/charts/chart") and name.endswith(".xml"):
                results.append((name, z.read(name)))
    return results


def extract_theme_xml(pptx_path: Path) -> bytes | None:
    """Return the theme1.xml bytes if present (first theme only for brevity)."""
    with zipfile.ZipFile(pptx_path, "r") as z:
        for name in z.namelist():
            if name.startswith("ppt/theme/theme") and name.endswith(".xml"):
                return z.read(name)
    return None


def extract_slide_xmls(pptx_path: Path) -> list[tuple[str, bytes]]:
    """Return list of (slide_filename, xml_bytes) tuples."""
    results: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        for name in sorted(z.namelist()):
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                results.append((name, z.read(name)))
    return results


# ---------------------------------------------------------------------------
# Theme color/font extraction
# ---------------------------------------------------------------------------


def parse_theme(theme_xml: bytes) -> dict:
    """Parse theme1.xml for colors and fonts."""
    info: dict[str, Any] = {"colors": {}, "fonts": {"major": None, "minor": None}}
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError:
        return info

    scheme = root.find(f".//{qn_a('clrScheme')}")
    if scheme is not None:
        for child in scheme:
            color_name = strip_ns(child.tag)
            rgb = child.find(qn_a("srgbClr"))
            sys_clr = child.find(qn_a("sysClr"))
            if rgb is not None:
                info["colors"][color_name] = "#" + rgb.get("val", "").upper()
            elif sys_clr is not None:
                info["colors"][color_name] = "sys:" + sys_clr.get("lastClr", sys_clr.get("val", ""))

    font_scheme = root.find(f".//{qn_a('fontScheme')}")
    if font_scheme is not None:
        major = font_scheme.find(qn_a("majorFont"))
        minor = font_scheme.find(qn_a("minorFont"))
        if major is not None:
            latin = major.find(qn_a("latin"))
            if latin is not None:
                info["fonts"]["major"] = latin.get("typeface")
        if minor is not None:
            latin = minor.find(qn_a("latin"))
            if latin is not None:
                info["fonts"]["minor"] = latin.get("typeface")

    return info


# ---------------------------------------------------------------------------
# Shape coordinate extraction — for layout pattern mining
# ---------------------------------------------------------------------------

EMU_PER_INCH = 914400.0


def analyze_slide_coords(slide) -> list[dict]:
    """Extract shape coordinates + types + text samples per slide."""
    out: list[dict] = []

    def walk(shapes, parent="root"):
        for s in shapes:
            try:
                left = s.left / EMU_PER_INCH if s.left is not None else None
                top = s.top / EMU_PER_INCH if s.top is not None else None
                width = s.width / EMU_PER_INCH if s.width is not None else None
                height = s.height / EMU_PER_INCH if s.height is not None else None
            except Exception:
                left = top = width = height = None

            stype = str(s.shape_type).split(" ")[0].replace("XL_", "").lower() if s.shape_type else "none"
            sample_text = ""
            try:
                if s.has_text_frame:
                    sample_text = s.text_frame.text[:80].replace("\n", " ").strip()
            except Exception:
                pass

            out.append({
                "type": stype,
                "left": round(left, 2) if left is not None else None,
                "top": round(top, 2) if top is not None else None,
                "width": round(width, 2) if width is not None else None,
                "height": round(height, 2) if height is not None else None,
                "has_chart": s.has_chart,
                "has_table": s.has_table,
                "text_sample": sample_text,
                "parent": parent,
            })

            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    walk(s.shapes, parent="group")
                except Exception:
                    pass

    walk(slide.shapes)
    return out


def shape_signature(shapes: list[dict]) -> str:
    """Simplified signature used to cluster slides by composition."""
    counts: Counter = Counter()
    for s in shapes:
        counts[s["type"]] += 1
    keys = ["chart", "table", "text_box", "picture", "auto_shape"]
    parts = [f"{k}={counts.get(k, 0)}" for k in keys]
    return ",".join(parts)


# ---------------------------------------------------------------------------
# Headline / text pattern extraction
# ---------------------------------------------------------------------------


def extract_headlines(slide, slide_w_in: float) -> list[dict]:
    """Identify likely headline text boxes (top of slide, wide, larger font)."""
    results: list[dict] = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        try:
            top_in = shape.top / EMU_PER_INCH if shape.top is not None else 999
            width_in = shape.width / EMU_PER_INCH if shape.width is not None else 0
        except Exception:
            continue
        # Heuristic: headlines are in top 1.5" and span >40% of slide width
        if top_in > 1.5 or width_in < slide_w_in * 0.4:
            continue

        text = shape.text_frame.text.strip()
        if not text or len(text) < 8:
            continue

        # Font size of first run
        font_size = None
        font_name = None
        font_color = None
        bold = None
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                try:
                    font_size = int(run.font.size.pt) if run.font.size is not None else font_size
                except Exception:
                    pass
                try:
                    font_name = run.font.name or font_name
                except Exception:
                    pass
                try:
                    if run.font.color.type is not None:
                        rgb = run.font.color.rgb
                        if rgb is not None:
                            font_color = str(rgb)
                except Exception:
                    pass
                try:
                    bold = run.font.bold if run.font.bold is not None else bold
                except Exception:
                    pass
                if font_size is not None:
                    break
            if font_size is not None:
                break

        results.append({
            "text": text[:300],
            "char_count": len(text),
            "top_in": round(top_in, 2),
            "width_in": round(width_in, 2),
            "font_size_pt": font_size,
            "font_name": font_name,
            "font_color": font_color,
            "bold": bold,
        })
    return results


# ---------------------------------------------------------------------------
# Table structure extraction
# ---------------------------------------------------------------------------


def extract_tables(slide) -> list[dict]:
    results: list[dict] = []
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        try:
            tbl = shape.table
            rows = len(tbl.rows)
            cols = len(tbl.columns)
        except Exception:
            continue
        # Sample first cell text + a few middle cells
        try:
            top_left = tbl.rows[0].cells[0].text.strip()[:60] if rows and cols else ""
        except Exception:
            top_left = ""
        try:
            left_in = shape.left / EMU_PER_INCH
            top_in = shape.top / EMU_PER_INCH
            width_in = shape.width / EMU_PER_INCH
            height_in = shape.height / EMU_PER_INCH
        except Exception:
            left_in = top_in = width_in = height_in = None

        # Header cell color sample
        header_fill = None
        try:
            cell = tbl.rows[0].cells[0]
            fill = cell.fill
            if fill.type is not None:
                try:
                    rgb = fill.fore_color.rgb
                    header_fill = str(rgb) if rgb else None
                except Exception:
                    pass
        except Exception:
            pass

        results.append({
            "rows": rows,
            "cols": cols,
            "left_in": round(left_in, 2) if left_in is not None else None,
            "top_in": round(top_in, 2) if top_in is not None else None,
            "width_in": round(width_in, 2) if width_in is not None else None,
            "height_in": round(height_in, 2) if height_in is not None else None,
            "header_cell_sample": top_left,
            "header_fill": header_fill,
        })
    return results


# ---------------------------------------------------------------------------
# Master analyzer
# ---------------------------------------------------------------------------


def analyze_deck_deep(pptx_path: Path) -> dict:
    client = infer_client(pptx_path.name)

    deck_out: dict[str, Any] = {
        "filename": pptx_path.name,
        "client": client,
        "chart_configs": [],
        "theme": {},
        "headlines": [],
        "tables": [],
        "layouts": [],
        "brand_colors": Counter(),
        "brand_fonts": Counter(),
        "ooxml_tags": Counter(),
    }

    # Theme extraction
    theme_xml = extract_theme_xml(pptx_path)
    if theme_xml:
        deck_out["theme"] = parse_theme(theme_xml)

    # Chart XML deep analysis
    chart_xmls = extract_chart_xmls(pptx_path)
    for cname, cxml in chart_xmls:
        cfg = analyze_chart_xml(cxml)
        cfg["chart_file"] = cname
        deck_out["chart_configs"].append(cfg)
        # Tag counter
        try:
            walk_element_tags(ET.fromstring(cxml), deck_out["ooxml_tags"])
        except Exception:
            pass
        # Accumulate brand colors from series
        for col in cfg.get("series_colors", []):
            if col:
                deck_out["brand_colors"][col.upper()] += 1

    # Slide-level: coordinates, headlines, tables, layouts
    try:
        prs = Presentation(str(pptx_path))
        slide_w_in = prs.slide_width / EMU_PER_INCH
        for i, slide in enumerate(prs.slides):
            try:
                coords = analyze_slide_coords(slide)
                sig = shape_signature(coords)
                deck_out["layouts"].append({
                    "slide_index": i,
                    "signature": sig,
                    "layout_name": getattr(slide.slide_layout, "name", "unknown"),
                    "shapes": coords,
                })
                hdrs = extract_headlines(slide, slide_w_in)
                for h in hdrs:
                    h["slide_index"] = i
                    deck_out["headlines"].append(h)
                for t in extract_tables(slide):
                    t["slide_index"] = i
                    deck_out["tables"].append(t)
                # Font accumulator
                for shape in slide.shapes:
                    if not shape.has_text_frame:
                        continue
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            try:
                                if run.font.name:
                                    deck_out["brand_fonts"][run.font.name] += 1
                            except Exception:
                                pass
            except Exception:
                continue
    except Exception as e:
        deck_out["slide_error"] = str(e)

    # Convert Counters to dicts for JSON
    deck_out["brand_colors"] = dict(deck_out["brand_colors"])
    deck_out["brand_fonts"] = dict(deck_out["brand_fonts"])
    deck_out["ooxml_tags"] = {strip_ns(k): v for k, v in deck_out["ooxml_tags"].items()}
    return deck_out


# ---------------------------------------------------------------------------
# Cross-deck aggregations
# ---------------------------------------------------------------------------


def aggregate_ooxml_properties(all_decks: list[dict]) -> dict:
    """Aggregate OOXML tags seen across all decks."""
    tag_counter: Counter = Counter()
    for d in all_decks:
        for tag, cnt in d.get("ooxml_tags", {}).items():
            tag_counter[tag] += cnt
    return dict(tag_counter.most_common())


def aggregate_chart_properties(all_decks: list[dict]) -> dict:
    """Aggregate chart-level properties: gap_widths, overlaps, label positions, etc."""
    gap_widths: Counter = Counter()
    overlaps: Counter = Counter()
    data_label_positions: Counter = Counter()
    axis_orientations: Counter = Counter()
    tick_lbl_pos: Counter = Counter()
    num_formats: Counter = Counter()
    legend_positions: Counter = Counter()
    series_marker_types: Counter = Counter()
    series_line_widths: Counter = Counter()
    chart_tag_groupings: Counter = Counter()

    total_charts = 0
    charts_with_gridlines_major = 0
    charts_with_gridlines_minor = 0
    charts_with_title = 0
    charts_with_legend = 0
    charts_with_trendlines = 0
    charts_with_error_bars = 0

    for d in all_decks:
        for cfg in d.get("chart_configs", []):
            total_charts += 1
            for gw in cfg.get("gap_widths", []):
                gap_widths[gw] += 1
            for ov in cfg.get("overlaps", []):
                overlaps[ov] += 1
            for dlp in cfg.get("data_label_positions", []):
                data_label_positions[dlp] += 1
            for ao in cfg.get("axis_orientations", []):
                axis_orientations[f"{ao['axis']}:{ao['orientation']}"] += 1
            for tlp in cfg.get("tick_lbl_pos", []):
                tick_lbl_pos[f"{tlp['axis']}:{tlp['val']}"] += 1
            for nf in cfg.get("val_axis_number_formats", []):
                num_formats[nf] += 1
            for nf in cfg.get("num_formats_used", []):
                num_formats[nf] += 1
            if cfg.get("legend_position"):
                legend_positions[cfg["legend_position"]] += 1
            for mt in cfg.get("series_marker_types", []):
                series_marker_types[mt] += 1
            for lw in cfg.get("series_line_widths", []):
                if lw:
                    series_line_widths[lw] += 1
            for ct in cfg.get("chart_types", []):
                key = f"{ct['tag']}(bar_dir={ct.get('bar_dir')},grouping={ct.get('grouping')})"
                chart_tag_groupings[key] += 1
            if cfg.get("major_gridlines_present"):
                charts_with_gridlines_major += 1
            if cfg.get("minor_gridlines_present"):
                charts_with_gridlines_minor += 1
            if cfg.get("has_title"):
                charts_with_title += 1
            if cfg.get("has_legend"):
                charts_with_legend += 1
            if cfg.get("trendlines"):
                charts_with_trendlines += 1
            if cfg.get("error_bars"):
                charts_with_error_bars += 1

    return {
        "total_charts": total_charts,
        "chart_tag_groupings": dict(chart_tag_groupings.most_common()),
        "gap_widths": dict(gap_widths.most_common()),
        "overlaps": dict(overlaps.most_common()),
        "data_label_positions": dict(data_label_positions.most_common()),
        "axis_orientations": dict(axis_orientations.most_common()),
        "tick_lbl_positions": dict(tick_lbl_pos.most_common()),
        "num_formats": dict(num_formats.most_common()),
        "legend_positions": dict(legend_positions.most_common()),
        "series_marker_types": dict(series_marker_types.most_common()),
        "series_line_widths_emu": dict(series_line_widths.most_common()),
        "charts_with_gridlines_major": charts_with_gridlines_major,
        "charts_with_gridlines_minor": charts_with_gridlines_minor,
        "charts_with_title": charts_with_title,
        "charts_with_legend": charts_with_legend,
        "charts_with_trendlines": charts_with_trendlines,
        "charts_with_error_bars": charts_with_error_bars,
    }


def aggregate_brand_by_client(all_decks: list[dict]) -> dict:
    """Group colors + fonts + theme by client."""
    per_client: dict[str, dict] = defaultdict(lambda: {
        "decks": [],
        "colors": Counter(),
        "fonts": Counter(),
        "theme_colors": [],
        "theme_fonts_major": Counter(),
        "theme_fonts_minor": Counter(),
    })
    for d in all_decks:
        client = d.get("client", "Unknown")
        per_client[client]["decks"].append(d["filename"])
        for c, cnt in d.get("brand_colors", {}).items():
            per_client[client]["colors"][c] += cnt
        for f, cnt in d.get("brand_fonts", {}).items():
            per_client[client]["fonts"][f] += cnt
        theme = d.get("theme", {})
        if theme.get("colors"):
            per_client[client]["theme_colors"].append(theme["colors"])
        if theme.get("fonts", {}).get("major"):
            per_client[client]["theme_fonts_major"][theme["fonts"]["major"]] += 1
        if theme.get("fonts", {}).get("minor"):
            per_client[client]["theme_fonts_minor"][theme["fonts"]["minor"]] += 1

    out: dict[str, dict] = {}
    for client, info in per_client.items():
        out[client] = {
            "deck_count": len(info["decks"]),
            "decks": info["decks"],
            "top_series_colors": dict(info["colors"].most_common(20)),
            "top_fonts": dict(info["fonts"].most_common(10)),
            "theme_fonts_major": dict(info["theme_fonts_major"].most_common(5)),
            "theme_fonts_minor": dict(info["theme_fonts_minor"].most_common(5)),
            "theme_colors_samples": info["theme_colors"][:3],
        }
    return out


def aggregate_layouts(all_decks: list[dict]) -> dict:
    """Cluster slides by signature + extract coordinate statistics."""
    sig_groups: dict[str, list[dict]] = defaultdict(list)
    for d in all_decks:
        for slide in d.get("layouts", []):
            sig_groups[slide["signature"]].append(slide)

    out: dict[str, Any] = {"signatures": {}}
    for sig, slides in sig_groups.items():
        if len(slides) < 5:
            continue  # skip rare patterns

        # For this signature, analyze coordinates of each shape type
        coord_buckets: dict[str, dict] = defaultdict(lambda: {
            "lefts": [], "tops": [], "widths": [], "heights": [],
        })
        for slide in slides:
            for s in slide["shapes"]:
                stype = s["type"]
                if all(s[k] is not None for k in ("left", "top", "width", "height")):
                    coord_buckets[stype]["lefts"].append(s["left"])
                    coord_buckets[stype]["tops"].append(s["top"])
                    coord_buckets[stype]["widths"].append(s["width"])
                    coord_buckets[stype]["heights"].append(s["height"])

        type_stats: dict[str, dict] = {}
        for stype, coords in coord_buckets.items():
            if not coords["lefts"]:
                continue
            type_stats[stype] = {
                "count_observations": len(coords["lefts"]),
                "left_mean": round(mean(coords["lefts"]), 2),
                "left_median": round(median(coords["lefts"]), 2),
                "top_mean": round(mean(coords["tops"]), 2),
                "top_median": round(median(coords["tops"]), 2),
                "width_mean": round(mean(coords["widths"]), 2),
                "width_median": round(median(coords["widths"]), 2),
                "height_mean": round(mean(coords["heights"]), 2),
                "height_median": round(median(coords["heights"]), 2),
            }
            if len(coords["lefts"]) > 1:
                try:
                    type_stats[stype]["left_stdev"] = round(stdev(coords["lefts"]), 2)
                    type_stats[stype]["top_stdev"] = round(stdev(coords["tops"]), 2)
                except Exception:
                    pass

        out["signatures"][sig] = {
            "occurrences": len(slides),
            "example_slides": [(s["slide_index"], s["layout_name"]) for s in slides[:5]],
            "coord_stats_by_shape_type": type_stats,
        }

    # Sort signatures by occurrence
    out["signatures"] = dict(sorted(
        out["signatures"].items(), key=lambda kv: -kv[1]["occurrences"]
    ))
    return out


def aggregate_headlines(all_decks: list[dict]) -> dict:
    headlines = []
    for d in all_decks:
        for h in d.get("headlines", []):
            h_copy = dict(h)
            h_copy["client"] = d.get("client")
            headlines.append(h_copy)

    font_sizes: Counter = Counter()
    font_colors: Counter = Counter()
    font_names: Counter = Counter()
    char_counts = []
    top_positions: Counter = Counter()
    widths = []

    for h in headlines:
        if h.get("font_size_pt"):
            font_sizes[h["font_size_pt"]] += 1
        if h.get("font_color"):
            font_colors[h["font_color"]] += 1
        if h.get("font_name"):
            font_names[h["font_name"]] += 1
        if h.get("char_count"):
            char_counts.append(h["char_count"])
        if h.get("top_in") is not None:
            top_positions[round(h["top_in"], 1)] += 1
        if h.get("width_in") is not None:
            widths.append(h["width_in"])

    return {
        "headline_count": len(headlines),
        "font_sizes": dict(font_sizes.most_common(20)),
        "font_colors": dict(font_colors.most_common(20)),
        "font_names": dict(font_names.most_common(10)),
        "char_count_stats": {
            "min": min(char_counts) if char_counts else 0,
            "max": max(char_counts) if char_counts else 0,
            "mean": round(mean(char_counts), 1) if char_counts else 0,
            "median": round(median(char_counts), 1) if char_counts else 0,
        },
        "top_positions_in": dict(sorted(top_positions.items())),
        "width_stats_in": {
            "min": round(min(widths), 2) if widths else 0,
            "max": round(max(widths), 2) if widths else 0,
            "mean": round(mean(widths), 2) if widths else 0,
            "median": round(median(widths), 2) if widths else 0,
        },
    }


def aggregate_tables(all_decks: list[dict]) -> dict:
    row_counts: Counter = Counter()
    col_counts: Counter = Counter()
    dimensions: Counter = Counter()
    header_fills: Counter = Counter()
    widths = []
    heights = []

    for d in all_decks:
        for t in d.get("tables", []):
            row_counts[t["rows"]] += 1
            col_counts[t["cols"]] += 1
            dimensions[f"{t['rows']}x{t['cols']}"] += 1
            if t.get("header_fill"):
                header_fills[t["header_fill"]] += 1
            if t.get("width_in") is not None:
                widths.append(t["width_in"])
            if t.get("height_in") is not None:
                heights.append(t["height_in"])

    return {
        "table_count": sum(row_counts.values()),
        "row_count_distribution": dict(row_counts.most_common(20)),
        "col_count_distribution": dict(col_counts.most_common(20)),
        "dimensions_top_20": dict(dimensions.most_common(20)),
        "header_fills_top_20": dict(header_fills.most_common(20)),
        "width_stats_in": {
            "min": round(min(widths), 2) if widths else 0,
            "max": round(max(widths), 2) if widths else 0,
            "mean": round(mean(widths), 2) if widths else 0,
            "median": round(median(widths), 2) if widths else 0,
        } if widths else {},
        "height_stats_in": {
            "min": round(min(heights), 2) if heights else 0,
            "max": round(max(heights), 2) if heights else 0,
            "mean": round(mean(heights), 2) if heights else 0,
            "median": round(median(heights), 2) if heights else 0,
        } if heights else {},
    }


# ---------------------------------------------------------------------------
# Report writers
# ---------------------------------------------------------------------------


def write_deep_report(
    all_decks: list[dict],
    chart_agg: dict,
    ooxml_props: dict,
    brand_by_client: dict,
    layout_agg: dict,
    headline_agg: dict,
    table_agg: dict,
    output: Path,
) -> None:
    lines: list[str] = []
    lines.append("# Deep Deck Analysis Report")
    lines.append("")
    lines.append(f"**Decks analyzed:** {len(all_decks)}")
    lines.append(f"**Total charts:** {chart_agg['total_charts']}")
    lines.append(f"**Total headlines detected:** {headline_agg['headline_count']}")
    lines.append(f"**Total tables:** {table_agg['table_count']}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ===== Chart Config =====
    lines.append("## 1. Chart Configuration (OOXML-level)")
    lines.append("")
    lines.append("### 1.1 Chart Tag + Grouping Combinations (top 20)")
    lines.append("")
    lines.append("| Tag / Direction / Grouping | Count |")
    lines.append("|---|---|")
    for tag, cnt in list(chart_agg["chart_tag_groupings"].items())[:20]:
        lines.append(f"| `{tag}` | {cnt} |")
    lines.append("")

    lines.append("### 1.2 Bar/Column Gap Width")
    lines.append("")
    lines.append("The `<c:gapWidth>` property controls space between bars. python-pptx does NOT expose this — requires lxml_helpers.")
    lines.append("")
    lines.append("| Gap Width | Count |")
    lines.append("|---|---|")
    for gw, cnt in chart_agg["gap_widths"].items():
        lines.append(f"| `{gw}` | {cnt} |")
    lines.append("")

    lines.append("### 1.3 Bar/Column Overlap")
    lines.append("")
    lines.append("| Overlap | Count |")
    lines.append("|---|---|")
    for ov, cnt in chart_agg["overlaps"].items():
        lines.append(f"| `{ov}` | {cnt} |")
    lines.append("")

    lines.append("### 1.4 Data Label Positions")
    lines.append("")
    lines.append("| Position | Count |")
    lines.append("|---|---|")
    for dlp, cnt in chart_agg["data_label_positions"].items():
        lines.append(f"| `{dlp}` | {cnt} |")
    lines.append("")

    lines.append("### 1.5 Axis Orientation")
    lines.append("")
    lines.append("`maxMin` = inverted axis (top-down for category axis in bar charts).")
    lines.append("")
    lines.append("| Axis : Orientation | Count |")
    lines.append("|---|---|")
    for ao, cnt in chart_agg["axis_orientations"].items():
        lines.append(f"| `{ao}` | {cnt} |")
    lines.append("")

    lines.append("### 1.6 Tick Label Positions")
    lines.append("")
    lines.append("`none` = hidden axis labels (common when labels shown in separate table).")
    lines.append("")
    lines.append("| Axis : Position | Count |")
    lines.append("|---|---|")
    for tlp, cnt in chart_agg["tick_lbl_positions"].items():
        lines.append(f"| `{tlp}` | {cnt} |")
    lines.append("")

    lines.append("### 1.7 Number Formats (top 20)")
    lines.append("")
    lines.append("| Format | Count |")
    lines.append("|---|---|")
    for nf, cnt in list(chart_agg["num_formats"].items())[:20]:
        lines.append(f"| `{nf}` | {cnt} |")
    lines.append("")

    lines.append("### 1.8 Legend Positions")
    lines.append("")
    lines.append("| Position | Count |")
    lines.append("|---|---|")
    for lp, cnt in chart_agg["legend_positions"].items():
        lines.append(f"| `{lp}` | {cnt} |")
    lines.append("")

    lines.append("### 1.9 Series Marker Types (scatter/line charts)")
    lines.append("")
    lines.append("| Marker | Count |")
    lines.append("|---|---|")
    for mt, cnt in chart_agg["series_marker_types"].items():
        lines.append(f"| `{mt}` | {cnt} |")
    lines.append("")

    lines.append("### 1.10 Series Line Widths (EMU)")
    lines.append("")
    lines.append("12700 EMU = 1pt. 19050 EMU = 1.5pt. 25400 EMU = 2pt.")
    lines.append("")
    lines.append("| Width (EMU) | Count |")
    lines.append("|---|---|")
    for lw, cnt in list(chart_agg["series_line_widths_emu"].items())[:20]:
        lines.append(f"| `{lw}` | {cnt} |")
    lines.append("")

    lines.append("### 1.11 Chart Features Summary")
    lines.append("")
    total = chart_agg["total_charts"] or 1
    lines.append(f"- Charts with title: **{chart_agg['charts_with_title']}** ({round(chart_agg['charts_with_title']/total*100)}%)")
    lines.append(f"- Charts with legend: **{chart_agg['charts_with_legend']}** ({round(chart_agg['charts_with_legend']/total*100)}%)")
    lines.append(f"- Charts with major gridlines: **{chart_agg['charts_with_gridlines_major']}** ({round(chart_agg['charts_with_gridlines_major']/total*100)}%)")
    lines.append(f"- Charts with minor gridlines: **{chart_agg['charts_with_gridlines_minor']}** ({round(chart_agg['charts_with_gridlines_minor']/total*100)}%)")
    lines.append(f"- Charts with trendlines: **{chart_agg['charts_with_trendlines']}**")
    lines.append(f"- Charts with error bars: **{chart_agg['charts_with_error_bars']}**")
    lines.append("")

    lines.append("---")
    lines.append("")

    # ===== OOXML Property Inventory =====
    lines.append("## 2. OOXML Property Inventory")
    lines.append("")
    lines.append("Every chart-XML tag seen across all decks. Tags python-pptx covers natively are NOT flagged as gaps — but this list is exhaustive reference for which lxml helpers might be needed.")
    lines.append("")
    lines.append(f"**{len(ooxml_props)}** distinct chart XML tags encountered.")
    lines.append("")
    lines.append("### 2.1 Top 40 Most-Used Tags")
    lines.append("")
    lines.append("| Tag | Occurrences |")
    lines.append("|---|---|")
    for tag, cnt in list(ooxml_props.items())[:40]:
        lines.append(f"| `{tag}` | {cnt} |")
    lines.append("")

    lines.append("---")
    lines.append("")

    # ===== Brand by Client =====
    lines.append("## 3. Brand Extraction by Client")
    lines.append("")
    for client, info in sorted(brand_by_client.items(), key=lambda kv: -kv[1]["deck_count"]):
        lines.append(f"### {client} ({info['deck_count']} deck{'s' if info['deck_count'] != 1 else ''})")
        lines.append("")
        lines.append("**Decks:**")
        for fn in info["decks"]:
            lines.append(f"- `{fn}`")
        lines.append("")
        lines.append("**Top series colors (from chart XML):**")
        lines.append("")
        lines.append("| Hex | Count |")
        lines.append("|---|---|")
        for c, cnt in list(info["top_series_colors"].items())[:15]:
            lines.append(f"| `#{c}` | {cnt} |")
        lines.append("")
        lines.append("**Top fonts (from slides):**")
        lines.append("")
        lines.append("| Font | Count |")
        lines.append("|---|---|")
        for f, cnt in list(info["top_fonts"].items())[:8]:
            lines.append(f"| `{f}` | {cnt} |")
        lines.append("")
        if info["theme_fonts_major"]:
            lines.append("**Theme fonts (major/minor):**")
            lines.append("")
            lines.append(f"- Major: {list(info['theme_fonts_major'].keys())}")
            lines.append(f"- Minor: {list(info['theme_fonts_minor'].keys())}")
            lines.append("")
        lines.append("")

    lines.append("---")
    lines.append("")

    # ===== Layout Patterns =====
    lines.append("## 4. Layout Coordinate Analysis")
    lines.append("")
    lines.append("For recurring shape signatures (≥5 occurrences), coordinate statistics by shape type.")
    lines.append("")
    lines.append("The `left_median` and `width_median` values are directly usable as `LAYOUTS{}` entries in `pptx_utils/layout.py`.")
    lines.append("")

    for sig, details in list(layout_agg["signatures"].items())[:15]:
        lines.append(f"### Signature: `{sig}` ({details['occurrences']} slides)")
        lines.append("")
        lines.append("**Coordinate stats by shape type:**")
        lines.append("")
        lines.append("| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |")
        lines.append("|---|---|---|---|---|---|")
        for stype, stats in details["coord_stats_by_shape_type"].items():
            lines.append(
                f"| `{stype}` | {stats['count_observations']} | "
                f"{stats['left_median']}\" | {stats['top_median']}\" | "
                f"{stats['width_median']}\" | {stats['height_median']}\" |"
            )
        lines.append("")

    lines.append("---")
    lines.append("")

    # ===== Headlines =====
    lines.append("## 5. Headline Patterns")
    lines.append("")
    lines.append(f"**{headline_agg['headline_count']}** headline-like text boxes detected (top of slide, wide, ≥8 chars).")
    lines.append("")
    lines.append(f"**Char count:** min {headline_agg['char_count_stats']['min']}, median {headline_agg['char_count_stats']['median']}, max {headline_agg['char_count_stats']['max']}")
    lines.append("")
    lines.append(f"**Width (inches):** median {headline_agg['width_stats_in']['median']}, min {headline_agg['width_stats_in']['min']}, max {headline_agg['width_stats_in']['max']}")
    lines.append("")
    lines.append("### Font sizes (top 10)")
    lines.append("")
    lines.append("| Size (pt) | Count |")
    lines.append("|---|---|")
    for size, cnt in list(headline_agg["font_sizes"].items())[:10]:
        lines.append(f"| {size} | {cnt} |")
    lines.append("")
    lines.append("### Font colors (top 15)")
    lines.append("")
    lines.append("| Hex | Count |")
    lines.append("|---|---|")
    for c, cnt in list(headline_agg["font_colors"].items())[:15]:
        lines.append(f"| `#{c}` | {cnt} |")
    lines.append("")
    lines.append("### Top positions (top of slide, inches)")
    lines.append("")
    lines.append("| Top (in) | Count |")
    lines.append("|---|---|")
    for pos, cnt in list(headline_agg["top_positions_in"].items())[:15]:
        lines.append(f"| {pos} | {cnt} |")
    lines.append("")

    lines.append("---")
    lines.append("")

    # ===== Tables =====
    lines.append("## 6. Table Patterns")
    lines.append("")
    lines.append(f"**{table_agg['table_count']}** tables across all decks.")
    lines.append("")
    if table_agg.get("width_stats_in"):
        lines.append(f"**Width (inches):** median {table_agg['width_stats_in']['median']}, min {table_agg['width_stats_in']['min']}, max {table_agg['width_stats_in']['max']}")
        lines.append(f"**Height (inches):** median {table_agg['height_stats_in']['median']}, min {table_agg['height_stats_in']['min']}, max {table_agg['height_stats_in']['max']}")
        lines.append("")
    lines.append("### Top dimensions (rows × cols)")
    lines.append("")
    lines.append("| Dimensions | Count |")
    lines.append("|---|---|")
    for dim, cnt in list(table_agg["dimensions_top_20"].items())[:15]:
        lines.append(f"| `{dim}` | {cnt} |")
    lines.append("")

    lines.append("### Row count distribution (top 10)")
    lines.append("")
    lines.append("| Rows | Count |")
    lines.append("|---|---|")
    for r, cnt in list(table_agg["row_count_distribution"].items())[:10]:
        lines.append(f"| {r} | {cnt} |")
    lines.append("")

    lines.append("### Col count distribution (top 10)")
    lines.append("")
    lines.append("| Cols | Count |")
    lines.append("|---|---|")
    for c, cnt in list(table_agg["col_count_distribution"].items())[:10]:
        lines.append(f"| {c} | {cnt} |")
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


def write_brand_proposals_md(brand_by_client: dict, output: Path) -> None:
    lines: list[str] = []
    lines.append("# BRAND{} Proposal by Client")
    lines.append("")
    lines.append("Derived from real deck analysis. For each client, proposes a `BRAND{}` entry structure using observed series colors and fonts.")
    lines.append("")
    lines.append("Use as input for `pptx_utils/brand.py`.")
    lines.append("")

    for client, info in sorted(brand_by_client.items(), key=lambda kv: -kv[1]["deck_count"]):
        lines.append(f"## {client}")
        lines.append("")
        # Derive best guesses from top series colors
        top_colors = list(info["top_series_colors"].keys())
        top_fonts = list(info["top_fonts"].keys())
        theme_major = list(info["theme_fonts_major"].keys())

        # Simple heuristic: first distinctive color = primary, second = secondary
        primary_guess = top_colors[0] if top_colors else "000000"
        secondary_guess = top_colors[1] if len(top_colors) > 1 else "808080"

        lines.append("```python")
        lines.append(f'"{client}": {{')
        lines.append(f'    "primary":    RGBColor(0x{primary_guess[0:2]}, 0x{primary_guess[2:4]}, 0x{primary_guess[4:6]}),  # most-used series color')
        lines.append(f'    "secondary":  RGBColor(0x{secondary_guess[0:2]}, 0x{secondary_guess[2:4]}, 0x{secondary_guess[4:6]}),  # 2nd most-used series color')
        lines.append(f'    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta')
        lines.append(f'    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta')
        if theme_major:
            lines.append(f'    "font_heading": "{theme_major[0]}",  # from theme')
        else:
            lines.append(f'    "font_heading": "Arial",  # default')
        if top_fonts:
            best_font = next((f for f in top_fonts if not f.startswith("+")), "Arial")
            lines.append(f'    "font_body":    "{best_font}",  # most-used body font')
        lines.append("},")
        lines.append("```")
        lines.append("")
        lines.append("**Observed color palette (top 10):**")
        lines.append("")
        for c in list(info["top_series_colors"].keys())[:10]:
            lines.append(f"- `#{c}`")
        lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    pptx_files = sorted(p for p in DECKS_DIR.glob("*.pptx") if not p.name.startswith("~"))
    if not pptx_files:
        print(f"No .pptx in {DECKS_DIR}", file=sys.stderr)
        return 1

    print(f"Deep-analyzing {len(pptx_files)} decks...")
    all_decks: list[dict] = []
    for i, p in enumerate(pptx_files, 1):
        print(f"  [{i}/{len(pptx_files)}] {p.name}", flush=True)
        try:
            all_decks.append(analyze_deck_deep(p))
        except Exception as e:
            print(f"    ERROR: {e}", file=sys.stderr)
            all_decks.append({"filename": p.name, "error": str(e)})

    print("\nAggregating...")
    chart_agg = aggregate_chart_properties(all_decks)
    ooxml_props = aggregate_ooxml_properties(all_decks)
    brand_by_client = aggregate_brand_by_client(all_decks)
    layout_agg = aggregate_layouts(all_decks)
    headline_agg = aggregate_headlines(all_decks)
    table_agg = aggregate_tables(all_decks)

    # Write JSON outputs
    (OUTPUTS_DIR / "deep_chart_details.json").write_text(
        json.dumps(chart_agg, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "deep_ooxml_properties.json").write_text(
        json.dumps(ooxml_props, indent=2), encoding="utf-8"
    )
    (OUTPUTS_DIR / "deep_brand_proposals.json").write_text(
        json.dumps(brand_by_client, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "deep_layouts.json").write_text(
        json.dumps(layout_agg, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "deep_headlines.json").write_text(
        json.dumps(headline_agg, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "deep_tables.json").write_text(
        json.dumps(table_agg, indent=2, default=str), encoding="utf-8"
    )

    # Write readable reports
    write_deep_report(
        all_decks, chart_agg, ooxml_props, brand_by_client,
        layout_agg, headline_agg, table_agg,
        OUTPUTS_DIR / "deep_report.md",
    )
    write_brand_proposals_md(brand_by_client, OUTPUTS_DIR / "deep_brand_proposals.md")

    print(f"\nDone. Outputs in {OUTPUTS_DIR}/")
    print("  deep_report.md                — master readable report")
    print("  deep_brand_proposals.md       — BRAND{} proposals per client")
    print("  deep_chart_details.json       — all chart property aggregations")
    print("  deep_ooxml_properties.json    — every chart-XML tag touched")
    print("  deep_brand_proposals.json     — brand data (structured)")
    print("  deep_layouts.json             — coordinate stats by signature")
    print("  deep_headlines.json           — headline patterns")
    print("  deep_tables.json              — table patterns")

    return 0


if __name__ == "__main__":
    sys.exit(main())
