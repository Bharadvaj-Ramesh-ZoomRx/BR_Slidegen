"""
Visual Regression Harness
=========================

Measures a generated or existing PPTX against the golden coordinate clusters
from the Apr 15 deck analysis. Defines the acceptance bar for slide-creator.

**Axes checked** (per slide-creator SKILL.md):
  1. Layout   — shape bboxes within ±0.1" of cluster median
  2. Viz      — chart pattern OOXML flags match CHART_PATTERNS entry
  3. Data     — data-label strings match expected format
  4. Chrome   — title/legend/gridlines presence matches spec intent
  5. Color    — series colors match expected brand palette (if brand + BRAND{} entry provided)

## Modes

**Baseline mode** (run today against real decks):
    python visual_regression.py --deck decks/JJ_PET_RYBREVANT_Q1_26.pptx
    → Extracts every slide's shape signature, matches to nearest cluster,
      reports per-shape drift. No pass/fail; establishes what "good" looks like.

**Regression mode** (once slide-creator lands):
    python visual_regression.py --deck output/Q1_26/deck.pptx \\
                                --expected-cluster 1_chart_1_table \\
                                --brand RYBREVANT --strict
    → Every shape must be ≤0.1" from the target cluster median. Exits non-zero
      on failure. Used in CI.

**Spec mode** (render a spec and verify it):
    python visual_regression.py --spec slidegen/slide_spec/examples/bar_clustered_with_delta.json
    → Loads spec, calls slide-creator (when available), measures output.
      Currently reports WHAT WILL BE CHECKED — waits on slide-creator impl.

## Cluster matching

Clusters are identified by shape-type signature: e.g. "1_chart_1_table" =
"exactly one chart + exactly one table + any number of text boxes + no
pictures." See experiments/deck_analysis/outputs/layout_clusters.md for
the full signature list.

## Run

    cd experiments/deck_analysis
    python visual_regression.py --deck path/to/deck.pptx
"""
from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

HERE = Path(__file__).parent
OUTPUTS_DIR = HERE / "outputs"
CLUSTERS_FILE = OUTPUTS_DIR / "layout_clusters.json"
CHART_POSITIONS_FILE = OUTPUTS_DIR / "chart_positions.json"

# Slide dimensions (inches) — matches pptx_utils
SLIDE_W = 13.333
SLIDE_H = 7.500

# Default fidelity tolerances (see slide-creator SKILL.md §Acceptance Bar)
LAYOUT_TOL_IN = 0.1       # shape bbox drift tolerance, inches
CHROME_REQUIRED_FREQ = {
    "has_title": 0.01,     # 1% — no-title is default
    "has_legend": 0.01,    # 1%
    "has_major_gridlines": 0.16,
    "data_label_pct": 0.96,
}

# OOXML namespaces
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


EMU_PER_IN = 914400


def emu_to_in(emu: int | float | None) -> float | None:
    if emu is None:
        return None
    return round(float(emu) / EMU_PER_IN, 3)


# ---------------------------------------------------------------------------
# Shape bbox extraction
# ---------------------------------------------------------------------------


@dataclass
class ShapeBbox:
    slide_index: int
    shape_type: str                  # "chart", "table", "text_box", "picture", "other"
    left: float
    top: float
    width: float
    height: float
    name: Optional[str] = None

    def signature_type(self) -> str:
        # Cluster signatures group by shape type
        return self.shape_type


@dataclass
class SlideSignature:
    slide_index: int
    shapes: list[ShapeBbox] = field(default_factory=list)

    def signature(self) -> str:
        """Build a cluster signature string: e.g. '1_chart_1_table'."""
        counts = Counter(s.shape_type for s in self.shapes if s.shape_type != "text_box")
        if not counts:
            return "empty"
        parts = sorted(f"{n}_{t}" for t, n in counts.items() if t != "other")
        return "_".join(parts) if parts else "empty"


def extract_slide_signatures(pptx_path: Path) -> list[SlideSignature]:
    """Read every slide's shapes and return bboxes + cluster signatures.

    Uses a lightweight ZIP + XML parse (no python-pptx dependency for Track 1
    scaffold — we can upgrade to python-pptx when slide-creator lands).
    """
    signatures: list[SlideSignature] = []
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            slide_files = sorted(
                n for n in z.namelist()
                if n.startswith("ppt/slides/slide") and n.endswith(".xml")
            )
            # Sort numerically (slide1.xml < slide2.xml < slide10.xml)
            slide_files.sort(key=lambda n: int(n.replace("ppt/slides/slide", "").replace(".xml", "")))

            for idx, fname in enumerate(slide_files):
                sig = SlideSignature(slide_index=idx)
                try:
                    root = ET.fromstring(z.read(fname))
                except ET.ParseError:
                    signatures.append(sig)
                    continue

                # Walk spTree / graphicFrame / sp / pic
                for shape in root.iter():
                    tag = shape.tag.split("}")[-1]
                    if tag not in ("sp", "graphicFrame", "pic"):
                        continue

                    # Locate xfrm
                    xfrm = shape.find(f".//{qn(NS_P, 'spPr')}/{qn(NS_A, 'xfrm')}") \
                           or shape.find(f".//{qn(NS_A, 'xfrm')}")
                    if xfrm is None:
                        continue
                    off = xfrm.find(qn(NS_A, "off"))
                    ext = xfrm.find(qn(NS_A, "ext"))
                    if off is None or ext is None:
                        continue
                    left = emu_to_in(int(off.get("x", 0)))
                    top = emu_to_in(int(off.get("y", 0)))
                    width = emu_to_in(int(ext.get("cx", 0)))
                    height = emu_to_in(int(ext.get("cy", 0)))

                    # Classify shape type
                    if tag == "graphicFrame":
                        # Chart or table?
                        data_elem = shape.find(f".//{qn(NS_A, 'graphic')}/{qn(NS_A, 'graphicData')}")
                        uri = data_elem.get("uri") if data_elem is not None else ""
                        if "chart" in uri:
                            shape_type = "chart"
                        elif "table" in uri:
                            shape_type = "table"
                        else:
                            shape_type = "graphic_frame"
                    elif tag == "pic":
                        shape_type = "picture"
                    elif tag == "sp":
                        # Textbox if it has a txBody
                        txbody = shape.find(qn(NS_P, "txBody"))
                        shape_type = "text_box" if txbody is not None else "other"
                    else:
                        shape_type = "other"

                    sig.shapes.append(ShapeBbox(
                        slide_index=idx,
                        shape_type=shape_type,
                        left=left or 0.0,
                        top=top or 0.0,
                        width=width or 0.0,
                        height=height or 0.0,
                    ))
                signatures.append(sig)
    except Exception as exc:
        print(f"ERROR reading {pptx_path}: {exc}", file=sys.stderr)
        return []
    return signatures


# ---------------------------------------------------------------------------
# Cluster matching + drift reporting
# ---------------------------------------------------------------------------


def load_clusters() -> dict:
    if not CLUSTERS_FILE.exists():
        raise FileNotFoundError(f"Cluster file not found: {CLUSTERS_FILE}. "
                                f"Run deep_analyzer.py + layout_clusterer.py first.")
    return json.loads(CLUSTERS_FILE.read_text(encoding="utf-8"))


@dataclass
class DriftReport:
    slide_index: int
    cluster_matched: Optional[str]
    shape_type: str
    axis: str                        # "left" | "top" | "width" | "height"
    observed: float
    expected_median: float
    drift: float                     # observed - expected_median
    within_tolerance: bool


def measure_drift(sig: SlideSignature, clusters: dict,
                  tolerance_in: float = LAYOUT_TOL_IN) -> list[DriftReport]:
    """For a slide signature, find the matching cluster and measure per-shape drift.

    Returns a list of DriftReport entries — one per (shape × axis).
    """
    reports: list[DriftReport] = []
    cluster_key = sig.signature()

    if cluster_key not in clusters:
        # Try fuzzy match by primary shape type
        primary_types = {s.shape_type for s in sig.shapes if s.shape_type != "text_box"}
        best = None
        for k in clusters:
            if all(f"1_{t}" in k or f"2_{t}" in k for t in primary_types):
                best = k
                break
        cluster_key = best

    if cluster_key is None or cluster_key not in clusters:
        return [DriftReport(
            slide_index=sig.slide_index, cluster_matched=None,
            shape_type="(none)", axis="(none)", observed=0.0,
            expected_median=0.0, drift=0.0, within_tolerance=False,
        )]

    cluster = clusters[cluster_key]
    stats = cluster.get("stats", {})

    for shape in sig.shapes:
        per_type_stats = stats.get(shape.shape_type)
        if per_type_stats is None:
            continue
        for axis in ("left", "top", "width", "height"):
            expected_median = per_type_stats.get(axis, {}).get("median")
            if expected_median is None:
                continue
            observed = getattr(shape, axis)
            drift = round(observed - expected_median, 3)
            reports.append(DriftReport(
                slide_index=sig.slide_index,
                cluster_matched=cluster_key,
                shape_type=shape.shape_type,
                axis=axis,
                observed=observed,
                expected_median=expected_median,
                drift=drift,
                within_tolerance=abs(drift) <= tolerance_in,
            ))
    return reports


# ---------------------------------------------------------------------------
# Chart OOXML inspection (viz + chrome + data axes)
# ---------------------------------------------------------------------------


@dataclass
class ChartAudit:
    slide_index: int
    chart_file: str
    has_title: bool
    has_legend: bool
    has_major_gridlines: bool
    data_label_formats: list[str]    # seen number formats on dLbls
    bar_direction: Optional[str]     # "bar" | "col" | None
    grouping: Optional[str]          # "clustered" | "stacked" | "percentStacked" | "standard"
    cat_axis_inverted: bool          # orientation="maxMin"
    series_colors: list[str]         # hex, in series order
    axis_num_formats: dict[str, str] # {"val": "0%", "cat": ""}


def audit_chart_xml(xml_bytes: bytes, slide_index: int, chart_file: str) -> Optional[ChartAudit]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return None
    c = lambda tag: qn(NS_C, tag)
    a = lambda tag: qn(NS_A, tag)

    # Title
    title = root.find(f".//{c('title')}")
    auto_title_deleted = root.find(f".//{c('autoTitleDeleted')}")
    has_title = title is not None and (auto_title_deleted is None or auto_title_deleted.get("val") != "1")

    # Legend
    has_legend = root.find(f".//{c('legend')}") is not None

    # Gridlines
    has_major_gridlines = root.find(f".//{c('majorGridlines')}") is not None

    # Bar direction + grouping
    bar_chart = root.find(f".//{c('barChart')}")
    bar_direction = grouping = None
    if bar_chart is not None:
        d = bar_chart.find(c("barDir"))
        g = bar_chart.find(c("grouping"))
        bar_direction = d.get("val") if d is not None else None
        grouping = g.get("val") if g is not None else None

    # Category axis inversion
    cat_axis = root.find(f".//{c('catAx')}")
    cat_axis_inverted = False
    if cat_axis is not None:
        orient = cat_axis.find(f"{c('scaling')}/{c('orientation')}")
        cat_axis_inverted = orient is not None and orient.get("val") == "maxMin"

    # Data label formats
    data_label_formats: list[str] = []
    for num_fmt in root.iter(c("numFmt")):
        code = num_fmt.get("formatCode")
        if code and code not in data_label_formats:
            data_label_formats.append(code)

    # Series colors
    series_colors: list[str] = []
    for ser in root.iter(c("ser")):
        sp_pr = ser.find(c("spPr"))
        if sp_pr is None:
            continue
        rgb = sp_pr.find(f"{a('solidFill')}/{a('srgbClr')}")
        if rgb is not None and rgb.get("val"):
            series_colors.append(rgb.get("val").upper())

    # Axis number formats
    axis_num_formats: dict[str, str] = {}
    for axis_tag, key in (("valAx", "val"), ("catAx", "cat")):
        ax = root.find(f".//{c(axis_tag)}")
        if ax is not None:
            num = ax.find(c("numFmt"))
            if num is not None:
                axis_num_formats[key] = num.get("formatCode", "")

    return ChartAudit(
        slide_index=slide_index,
        chart_file=chart_file,
        has_title=has_title,
        has_legend=has_legend,
        has_major_gridlines=has_major_gridlines,
        data_label_formats=data_label_formats,
        bar_direction=bar_direction,
        grouping=grouping,
        cat_axis_inverted=cat_axis_inverted,
        series_colors=series_colors,
        axis_num_formats=axis_num_formats,
    )


def audit_all_charts(pptx_path: Path) -> list[ChartAudit]:
    """Audit every chart XML in the deck. Returns one ChartAudit per chart."""
    audits: list[ChartAudit] = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        chart_files = sorted(
            n for n in z.namelist()
            if n.startswith("ppt/charts/chart") and n.endswith(".xml")
        )
        for i, fname in enumerate(chart_files):
            audit = audit_chart_xml(z.read(fname), slide_index=-1, chart_file=fname)
            if audit:
                audits.append(audit)
    return audits


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------


def summarize_drift(all_reports: list[DriftReport], tolerance_in: float) -> dict:
    total = len(all_reports)
    within = sum(1 for r in all_reports if r.within_tolerance)
    by_axis = Counter((r.shape_type, r.axis, r.within_tolerance) for r in all_reports)

    per_slide_fail: dict[int, int] = {}
    for r in all_reports:
        if not r.within_tolerance:
            per_slide_fail[r.slide_index] = per_slide_fail.get(r.slide_index, 0) + 1

    return {
        "total_measurements": total,
        "within_tolerance": within,
        "pct_within": round(100 * within / total, 1) if total else 0.0,
        "tolerance_in": tolerance_in,
        "per_shape_axis_breakdown": {
            f"{shape}.{axis}.{'pass' if ok else 'fail'}": n
            for (shape, axis, ok), n in by_axis.most_common()
        },
        "slides_with_failures": len(per_slide_fail),
        "top_failing_slides": sorted(per_slide_fail.items(), key=lambda kv: -kv[1])[:10],
    }


def summarize_chart_audits(audits: list[ChartAudit]) -> dict:
    n = len(audits)
    if n == 0:
        return {"n": 0}
    return {
        "n": n,
        "has_title_pct": round(100 * sum(a.has_title for a in audits) / n, 1),
        "has_legend_pct": round(100 * sum(a.has_legend for a in audits) / n, 1),
        "has_major_gridlines_pct": round(100 * sum(a.has_major_gridlines for a in audits) / n, 1),
        "cat_axis_inverted_pct": round(100 * sum(a.cat_axis_inverted for a in audits) / n, 1),
        "pct_label_format_0pct": round(
            100 * sum(1 for a in audits if "0%" in a.data_label_formats) / n, 1
        ),
        "bar_direction_counts": dict(Counter(a.bar_direction for a in audits)),
        "grouping_counts": dict(Counter(a.grouping for a in audits)),
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="SlideGen visual regression harness")
    parser.add_argument("--deck", type=Path,
                        help="Path to a PPTX to measure against cluster reference")
    parser.add_argument("--spec", type=Path,
                        help="Path to a spec JSON — renders via slide-creator "
                             "(NOT YET IMPLEMENTED — waits on slide-creator landing)")
    parser.add_argument("--expected-cluster", type=str, default=None,
                        help="Cluster signature to enforce (regression mode)")
    parser.add_argument("--brand", type=str, default=None,
                        help="BRAND key for color fidelity checks")
    parser.add_argument("--tolerance", type=float, default=LAYOUT_TOL_IN,
                        help=f"Layout bbox drift tolerance in inches (default {LAYOUT_TOL_IN})")
    parser.add_argument("--strict", action="store_true",
                        help="Exit non-zero if any measurement fails")
    parser.add_argument("--json", action="store_true",
                        help="Emit machine-readable JSON summary instead of human text")
    args = parser.parse_args(argv)

    if args.spec and not args.deck:
        print("[spec mode] slide-creator not yet implemented — printing what WILL be checked:")
        print(f"  1. Layout fidelity against cluster clusters (tolerance {args.tolerance}\")")
        print(f"  2. Chart pattern OOXML flags match CHART_PATTERNS[spec.chart_pattern]")
        print(f"  3. Data label strings match spec's chrome.data_labels.format")
        print(f"  4. Chrome flags (has_title, has_legend, gridlines) match spec")
        print(f"  5. Series colors match BRAND[spec.brand] if BRAND token used")
        return 0

    if not args.deck:
        parser.print_help()
        return 2

    if not args.deck.exists():
        print(f"Deck not found: {args.deck}", file=sys.stderr)
        return 2

    clusters = load_clusters()
    print(f"Loaded {len(clusters)} cluster signatures from {CLUSTERS_FILE.name}")

    signatures = extract_slide_signatures(args.deck)
    print(f"Extracted {len(signatures)} slide signatures from {args.deck.name}")

    all_reports: list[DriftReport] = []
    for sig in signatures:
        reports = measure_drift(sig, clusters, tolerance_in=args.tolerance)
        # If --expected-cluster, override the cluster match
        if args.expected_cluster and sig.signature() != args.expected_cluster:
            for r in reports:
                r.cluster_matched = args.expected_cluster
        all_reports.extend(reports)

    drift_summary = summarize_drift(all_reports, args.tolerance)
    chart_audits = audit_all_charts(args.deck)
    chart_summary = summarize_chart_audits(chart_audits)

    out = {
        "deck": str(args.deck),
        "slides_measured": len(signatures),
        "layout_fidelity": drift_summary,
        "chart_fidelity": chart_summary,
    }

    if args.json:
        print(json.dumps(out, indent=2))
    else:
        print("\n=== Layout Fidelity ===")
        print(f"  Tolerance: {args.tolerance}\"")
        print(f"  Measurements: {drift_summary['total_measurements']}")
        print(f"  Within tolerance: {drift_summary['within_tolerance']} "
              f"({drift_summary['pct_within']}%)")
        print(f"  Slides with any failure: {drift_summary['slides_with_failures']}")
        if drift_summary["top_failing_slides"]:
            print("  Top failing slides (index, fail_count):")
            for idx, n in drift_summary["top_failing_slides"]:
                print(f"    slide {idx}: {n} measurements out of tolerance")

        print("\n=== Chart Fidelity ===")
        if chart_summary["n"] == 0:
            print("  No charts found.")
        else:
            print(f"  Charts: {chart_summary['n']}")
            print(f"  has_title:          {chart_summary['has_title_pct']}%  (real decks: 1%)")
            print(f"  has_legend:         {chart_summary['has_legend_pct']}%  (real decks: 1%)")
            print(f"  major_gridlines:    {chart_summary['has_major_gridlines_pct']}%  (real decks: 16%)")
            print(f"  label_format_0pct:  {chart_summary['pct_label_format_0pct']}%  (real decks: 96%)")
            print(f"  cat_axis_inverted:  {chart_summary['cat_axis_inverted_pct']}%  (real decks: 43%)")
            print(f"  bar_direction:      {chart_summary['bar_direction_counts']}")
            print(f"  grouping:           {chart_summary['grouping_counts']}")

    if args.strict and drift_summary["within_tolerance"] < drift_summary["total_measurements"]:
        print(f"\nFAIL — {drift_summary['total_measurements'] - drift_summary['within_tolerance']} "
              f"measurements out of tolerance", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
