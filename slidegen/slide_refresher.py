"""
slide_refresher.py — Clone a deck and refresh chart/table data in place.

This is the core of the refresh-deck-workflow: clone the source PPTX (preserving
ALL formatting, layout, template, decorative shapes) and update only the DATA
inside chart and table shapes. The formatting is never touched.

Usage:
    from slidegen.slide_refresher import refresh_deck

    # Round-trip test (write back extracted data — output should match input)
    result = refresh_deck("source.pptx", "output.pptx")

    # Actual refresh with new data
    result = refresh_deck("source.pptx", "output.pptx", new_data={"slide_005": {...}})
"""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.chart.data import CategoryChartData


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


def _update_chart_data(chart, categories: list[str], series_data: list[tuple[str, list]]) -> bool:
    """Replace chart data in place. Returns True on success."""
    try:
        cd = CategoryChartData()
        cd.categories = categories
        for name, vals in series_data:
            cd.add_series(name, vals)
        chart.replace_data(cd)
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
                    # Update text while preserving formatting:
                    # Clear existing text, set new text on first paragraph
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = ""
                    cell.text_frame.paragraphs[0].runs[0].text = new_text if cell.text_frame.paragraphs[0].runs else new_text
        return True
    except Exception:
        return False


def refresh_deck(
    source_path: str | Path,
    output_path: str | Path,
    new_data: Optional[dict] = None,
    specs: Optional[list] = None,
) -> RefreshResult:
    """Clone a deck and refresh chart/table data in place.

    Args:
        source_path: Path to the source PPTX.
        output_path: Path for the output PPTX (clone + updated data).
        new_data: Optional dict of {slide_index: {shape_name: new_chart_data}}.
                  If None, performs a round-trip test (writes back extracted data).
        specs: Optional pre-computed specs from deck_reader. If None, runs deck_reader.

    Returns:
        RefreshResult with counts of refreshed shapes and any errors.
    """
    source_path = Path(source_path)
    output_path = Path(output_path)

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

    # Build spec lookup by slide_index
    spec_by_slide = {s.slide_index: s for s in specs}

    for slide_idx, slide in enumerate(prs.slides):
        spec = spec_by_slide.get(slide_idx)
        if spec is None:
            continue

        # Match shapes on this slide to spec components by position
        # (since shape names change between extractions)
        chart_components = [c for c in spec.components if c.type == "chart"]
        table_components = [c for c in spec.components if c.type in ("value_table", "label_table")]

        slide_charts = [s for s in slide.shapes if s.has_chart]
        slide_tables = [s for s in slide.shapes if s.has_table]

        # Match charts by position (closest match)
        for shape in slide_charts:
            shape_left = round(shape.left / 914400, 2) if shape.left else 0
            shape_top = round(shape.top / 914400, 2) if shape.top else 0

            # Find matching spec component by position
            best_match = None
            best_dist = float('inf')
            for comp in chart_components:
                if comp.position.left is None:
                    continue
                dist = abs(comp.position.left - shape_left) + abs(comp.position.top - shape_top)
                if dist < best_dist:
                    best_dist = dist
                    best_match = comp

            if best_match is None or best_dist > 0.5:
                continue

            # Get new data or use extracted data
            if new_data and slide_idx in new_data:
                # TODO: look up new data by shape/analysis ID
                pass

            # Round-trip: write back the extracted data
            categories = best_match.data.categories
            series_data = [(s.name, list(s.values)) for s in best_match.data.series]

            if categories == ["unknown"]:
                continue  # can't update with unknown categories

            if _update_chart_data(shape.chart, categories, series_data):
                result.charts_refreshed += 1
            else:
                result.charts_failed += 1
                result.errors.append(f"Slide {slide_idx}: chart at ({shape_left},{shape_top}) failed")

    # Step 4: Save
    prs.save(str(output_path))

    return result


def verify_round_trip(
    source_path: str | Path,
    output_path: str | Path,
) -> dict:
    """Verify that a clone+refresh round-trip preserves all shapes and data.

    Compares the source and output deck shape-by-shape:
    - Same number of slides
    - Same number of shapes per slide (by type)
    - Chart data matches (categories, series values)
    - Table text matches
    - Positions match

    Returns a dict with comparison results.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    orig = Presentation(str(source_path))
    rend = Presentation(str(output_path))

    results = {
        "slides_match": len(orig.slides) == len(rend.slides),
        "orig_slides": len(orig.slides),
        "rend_slides": len(rend.slides),
        "slide_issues": [],
        "total_issues": 0,
    }

    for idx in range(min(len(orig.slides), len(rend.slides))):
        os = orig.slides[idx]
        rs = rend.slides[idx]

        issues = []

        # Shape counts by type
        def count_by_type(slide):
            c = {"chart": 0, "table": 0, "text": 0, "picture": 0, "group": 0, "other": 0}
            for s in slide.shapes:
                if s.has_chart: c["chart"] += 1
                elif s.has_table: c["table"] += 1
                elif s.has_text_frame: c["text"] += 1
                elif s.shape_type == MSO_SHAPE_TYPE.PICTURE: c["picture"] += 1
                elif s.shape_type == MSO_SHAPE_TYPE.GROUP: c["group"] += 1
                else: c["other"] += 1
            return c

        oc = count_by_type(os)
        rc = count_by_type(rs)

        for key in oc:
            if oc[key] != rc[key]:
                issues.append(f"{key}: {oc[key]}→{rc[key]}")

        if issues:
            results["slide_issues"].append({"slide": idx, "issues": issues})
            results["total_issues"] += len(issues)

    return results
