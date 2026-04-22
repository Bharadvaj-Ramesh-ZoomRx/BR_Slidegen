"""
PS_slide_refresher.py — Fixed slide refresher with % format preservation.

Changes from slide_refresher.py:
  1. _update_chart_data: snapshots all formatCode elements from the chart XML
     BEFORE replace_data(), then restores them AFTER. python-pptx's replace_data()
     rewrites the numCache and resets every formatCode to "General", which turns
     values like 0.45 from displaying as "45%" into "0.45".

  2. next_output_version: helper to auto-increment output filenames
     (CREON_roundtrip_v1.pptx, v2.pptx, ...) so we never overwrite prior outputs.

  3. verify_round_trip: extended to also check formatCode preservation per chart,
     so format regressions are caught automatically.
"""
from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.chart.data import CategoryChartData

# OOXML chart namespace — used for all formatCode lookups
_NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_FC_TAG = f"{{{_NS_C}}}formatCode"


@dataclass
class RefreshResult:
    """Result of a deck refresh operation."""
    source_path: str
    output_path: str
    total_slides: int = 0
    charts_refreshed: int = 0
    tables_refreshed: int = 0
    charts_failed: int = 0
    tables_failed: int = 0
    errors: list[str] = field(default_factory=list)


def next_output_version(base_path: str | Path) -> Path:
    """Return the next versioned path for an output file.

    Given base_path="projects/CREON/output/CREON_roundtrip.pptx":
      - If no versioned files exist, returns .../CREON_roundtrip_v1.pptx
      - If v1 exists, returns v2; if v1+v2 exist, returns v3; etc.

    Versions are inserted before the file suffix.
    """
    base_path = Path(base_path)
    stem = base_path.stem
    suffix = base_path.suffix
    parent = base_path.parent

    # Find existing versioned files
    existing = list(parent.glob(f"{stem}_v*.{suffix.lstrip('.')}"))
    max_v = 0
    for p in existing:
        m = re.search(r"_v(\d+)$", p.stem)
        if m:
            max_v = max(max_v, int(m.group(1)))

    return parent / f"{stem}_v{max_v + 1}{suffix}"


def _snapshot_format_codes(chart) -> list[str]:
    """Capture all formatCode text values from a chart's OOXML, in document order.

    python-pptx's replace_data() regenerates numCache elements and resets
    every formatCode to "General". This snapshot lets us restore the originals.

    formatCode elements appear in:
      <c:numFmt> on data label specs (dLbls, dLbl) — NOT reset by replace_data
      <c:formatCode> inside <c:numCache> under each series — RESET by replace_data

    We snapshot ALL <c:formatCode> elements so the positional restore in
    _restore_format_codes() works regardless of chart type.
    """
    try:
        return [el.text or "General"
                for el in chart._chartSpace.iter(_FC_TAG)]
    except Exception:
        return []


def _restore_format_codes(chart, snapshot: list[str]) -> None:
    """Restore formatCode elements from a snapshot taken before replace_data().

    Matches elements positionally (same order as _snapshot_format_codes).
    If replace_data() added more elements than existed before (unlikely but
    possible), the extras are left as-is.
    """
    if not snapshot:
        return
    try:
        fc_els = list(chart._chartSpace.iter(_FC_TAG))
        for i, fc_el in enumerate(fc_els):
            if i < len(snapshot) and snapshot[i]:
                fc_el.text = snapshot[i]
    except Exception:
        pass


def _update_chart_data(
    chart,
    categories: list[str],
    series_data: list[tuple[str, list]],
) -> bool:
    """Replace chart data in place, preserving number formats. Returns True on success.

    Key fix: snapshot formatCode elements BEFORE replace_data(), restore AFTER.
    replace_data() rewrites numCache and resets all formatCodes to "General",
    which breaks % display (0.45 shows as "0.45" instead of "45%").
    """
    try:
        # 1. Capture formats before the data write
        fmt_snapshot = _snapshot_format_codes(chart)

        # 2. Write new data
        cd = CategoryChartData()
        cd.categories = categories
        for name, vals in series_data:
            cd.add_series(name, vals)
        chart.replace_data(cd)

        # 3. Restore formats
        _restore_format_codes(chart, fmt_snapshot)

        return True
    except Exception:
        return False


def _update_table_cells(table, rows: list[list[str]]) -> bool:
    """Update table cell text without touching formatting. Returns True on success."""
    try:
        n_rows = min(len(rows), len(table.rows))
        for r in range(n_rows):
            n_cols = min(len(rows[r]), len(table.columns))
            for c in range(n_cols):
                cell = table.cell(r, c)
                new_text = str(rows[r][c])
                if cell.text.strip() != new_text:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = ""
                    if (cell.text_frame.paragraphs and
                            cell.text_frame.paragraphs[0].runs):
                        cell.text_frame.paragraphs[0].runs[0].text = new_text
        return True
    except Exception:
        return False


def refresh_deck(
    source_path: str | Path,
    output_path: str | Path,
    new_data: Optional[dict] = None,
    specs: Optional[list] = None,
    auto_version: bool = False,
) -> RefreshResult:
    """Clone a deck and refresh chart/table data in place, preserving % formats.

    Args:
        source_path:  Path to the source PPTX.
        output_path:  Path for the output PPTX. If auto_version=True, the actual
                      path used is output_path with _vN suffix (see next_output_version).
        new_data:     Optional dict {slide_index: {shape_name: new_chart_data}}.
                      If None, performs a round-trip (writes back extracted data).
        specs:        Optional pre-computed specs from deck_reader.
                      If None, runs deck_reader on source_path.
        auto_version: If True, auto-increment output filename to avoid overwriting.

    Returns:
        RefreshResult with counts and the actual output_path used.
    """
    source_path = Path(source_path)
    output_path = Path(output_path)

    if auto_version:
        output_path = next_output_version(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Step 1: Extract specs if not provided
    if specs is None:
        from slidegen.deck_reader import read_deck
        specs, _summary = read_deck(str(source_path))

    # Step 2: Clone the source deck
    shutil.copy2(str(source_path), str(output_path))

    # Step 3: Open the clone and update data in place
    prs = Presentation(str(output_path))

    result = RefreshResult(
        source_path=str(source_path),
        output_path=str(output_path),
        total_slides=len(prs.slides),
    )

    spec_by_slide = {s.slide_index: s for s in specs}

    for slide_idx, slide in enumerate(prs.slides):
        spec = spec_by_slide.get(slide_idx)
        if spec is None:
            continue

        chart_components = [c for c in spec.components if c.type == "chart"]
        table_components = [c for c in spec.components
                            if c.type in ("value_table", "label_table")]

        slide_charts = [s for s in slide.shapes if s.has_chart]
        slide_tables = [s for s in slide.shapes if s.has_table]

        # Match charts by position (closest match within 0.5")
        for shape in slide_charts:
            shape_left = round(shape.left / 914400, 2) if shape.left else 0
            shape_top = round(shape.top / 914400, 2) if shape.top else 0

            best_match = None
            best_dist = float("inf")
            for comp in chart_components:
                if comp.position.left is None:
                    continue
                dist = (abs(comp.position.left - shape_left) +
                        abs(comp.position.top - shape_top))
                if dist < best_dist:
                    best_dist = dist
                    best_match = comp

            if best_match is None or best_dist > 0.5:
                continue

            if new_data and slide_idx in new_data:
                # TODO: wire Synapse refresh data here
                pass

            categories = best_match.data.categories
            series_data = [(s.name, list(s.values))
                           for s in best_match.data.series]

            if categories == ["unknown"]:
                continue

            if _update_chart_data(shape.chart, categories, series_data):
                result.charts_refreshed += 1
            else:
                result.charts_failed += 1
                result.errors.append(
                    f"Slide {slide_idx}: chart at ({shape_left},{shape_top}) failed")

    # Step 4: Save
    prs.save(str(output_path))

    return result


def verify_round_trip(
    source_path: str | Path,
    output_path: str | Path,
    check_formats: bool = True,
) -> dict:
    """Verify that a clone+refresh round-trip preserves shapes and formats.

    Checks:
      - Same number of slides
      - Same number of shapes per slide (by type)
      - formatCode preservation per chart (new: catches % regression)

    Args:
        source_path:   The original source PPTX.
        output_path:   The refreshed PPTX.
        check_formats: If True (default), also compare per-chart formatCodes
                       between source and output. Reports any chart where
                       replace_data() silently reset a "0%" format to "General".

    Returns dict with:
        slides_match, orig_slides, rend_slides, slide_issues, total_issues,
        format_issues (list of {slide, position, src_formats, out_formats})
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    orig = Presentation(str(source_path))
    rend = Presentation(str(output_path))

    results = {
        "slides_match": len(orig.slides) == len(rend.slides),
        "orig_slides": len(orig.slides),
        "rend_slides": len(rend.slides),
        "slide_issues": [],
        "format_issues": [],
        "total_issues": 0,
    }

    for idx in range(min(len(orig.slides), len(rend.slides))):
        os_ = orig.slides[idx]
        rs_ = rend.slides[idx]

        # ── Shape count check ──
        def count_by_type(slide):
            c = {"chart": 0, "table": 0, "text": 0,
                 "picture": 0, "group": 0, "other": 0}
            for s in slide.shapes:
                if s.has_chart:
                    c["chart"] += 1
                elif s.has_table:
                    c["table"] += 1
                elif s.has_text_frame:
                    c["text"] += 1
                elif s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    c["picture"] += 1
                elif s.shape_type == MSO_SHAPE_TYPE.GROUP:
                    c["group"] += 1
                else:
                    c["other"] += 1
            return c

        oc = count_by_type(os_)
        rc = count_by_type(rs_)
        shape_issues = [f"{k}: {oc[k]}→{rc[k]}"
                        for k in oc if oc[k] != rc[k]]
        if shape_issues:
            results["slide_issues"].append({"slide": idx, "issues": shape_issues})
            results["total_issues"] += len(shape_issues)

        # ── Format check ──
        if not check_formats:
            continue

        orig_charts = [s for s in os_.shapes if s.has_chart]
        rend_charts = [s for s in rs_.shapes if s.has_chart]

        for src_shape in orig_charts:
            src_left = round(src_shape.left / 914400, 2) if src_shape.left else 0
            src_top = round(src_shape.top / 914400, 2) if src_shape.top else 0

            # Find positional match in output
            best_ref = None
            best_dist = float("inf")
            for ref_shape in rend_charts:
                ref_left = round(ref_shape.left / 914400, 2) if ref_shape.left else 0
                ref_top = round(ref_shape.top / 914400, 2) if ref_shape.top else 0
                dist = abs(src_left - ref_left) + abs(src_top - ref_top)
                if dist < best_dist:
                    best_dist = dist
                    best_ref = ref_shape

            if best_ref is None or best_dist > 0.5:
                continue

            try:
                src_fmts = [el.text or "General"
                            for el in src_shape.chart._chartSpace.iter(_FC_TAG)]
                out_fmts = [el.text or "General"
                            for el in best_ref.chart._chartSpace.iter(_FC_TAG)]

                # Report if any formatCode changed (e.g. "0%" -> "General")
                if src_fmts != out_fmts:
                    results["format_issues"].append({
                        "slide": idx,
                        "position": f"({src_left},{src_top})",
                        "src_formats": src_fmts,
                        "out_formats": out_fmts,
                    })
                    results["total_issues"] += 1
            except Exception:
                pass

    return results
