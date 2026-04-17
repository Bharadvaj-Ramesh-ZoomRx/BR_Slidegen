"""
inference.py — Tier 2: Structural inference for untagged PPTX shapes.

For shapes WITHOUT Connector tags, parses headline text, chart OOXML
(bar/scatter/line/stacked classification), category labels, and embedded
table content. Cross-references against optional config.yaml + source_data.json
to propose likely extraction_method + question_codes.

Emits best-effort DataLineage in legacy fields (data_source, extraction_method,
question_codes, source_file). Flags confidence (high/medium/low) in metadata.

Requires user confirmation: writes a review doc <deck_name>_inferred_specs.md.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE

from slidegen.slide_spec.schema import (
    SlideSpec, HeadlineSpec, FooterSpec, Position, DataLineage, DataLineageCandidate,
    SlideMetadata, ChartComponent, LabelTableComponent, ValueTableComponent,
    DeltaColumnComponent, TextboxComponent,
    ChartData, Series, ChartChrome, DataLabelsSpec,
    dump_spec, SPEC_VERSION,
)


# ── Chart type classification ─────────────────────────────────────────────────

# Map python-pptx XL_CHART_TYPE enum values to our chart_pattern keys
_CHART_TYPE_MAP = {
    XL_CHART_TYPE.BAR_CLUSTERED: "bar_clustered_horizontal",
    XL_CHART_TYPE.BAR_STACKED: "bar_stacked_100_horizontal",
    XL_CHART_TYPE.BAR_STACKED_100: "bar_stacked_100_horizontal",
    XL_CHART_TYPE.COLUMN_CLUSTERED: "column_clustered_vertical",
    XL_CHART_TYPE.COLUMN_STACKED: "column_stacked_100_vertical",
    XL_CHART_TYPE.COLUMN_STACKED_100: "column_stacked_100_vertical",
    XL_CHART_TYPE.LINE: "line_markers_trended",
    XL_CHART_TYPE.LINE_MARKERS: "line_markers_trended",
    XL_CHART_TYPE.LINE_STACKED: "line_markers_trended",
    XL_CHART_TYPE.XY_SCATTER: "xy_scatter_abacus",
    XL_CHART_TYPE.XY_SCATTER_LINES: "xy_scatter_abacus",
    XL_CHART_TYPE.DOUGHNUT: "doughnut_default",
    XL_CHART_TYPE.PIE: "doughnut_default",
}


@dataclass
class InferenceSummary:
    """Summary of Tier 2 inference."""
    total_untagged: int = 0
    charts_inferred: int = 0
    tables_inferred: int = 0
    textboxes_skipped: int = 0
    specs_produced: int = 0
    confidence_counts: dict = field(default_factory=lambda: {
        "high": 0, "medium": 0, "low": 0
    })


def _emu_to_inches(emu: int) -> float:
    return round(emu / 914400, 2) if emu else 0.0


def _classify_chart_type(chart) -> tuple[str, str]:
    """Classify a python-pptx chart object into a chart_pattern + confidence.

    Returns (chart_pattern, confidence).
    """
    try:
        chart_type = chart.chart_type
        pattern = _CHART_TYPE_MAP.get(chart_type, None)
        if pattern:
            return pattern, "high"
        # Fallback: try to determine from XML
        return "bar_clustered_horizontal", "low"
    except Exception:
        return "bar_clustered_horizontal", "low"


def _extract_chart_data(chart) -> tuple[ChartData, str]:
    """Extract ChartData from a python-pptx chart object.

    Returns (ChartData, confidence).
    """
    confidence = "medium"
    categories = []
    series_list = []

    try:
        # Extract categories
        plot = chart.plots[0]
        try:
            cats = plot.categories
            if cats is not None:
                categories = [str(c) for c in cats]
        except Exception:
            categories = []

        # Extract series
        for idx, s in enumerate(plot.series):
            name = str(s.name) if s.name else f"Series {idx}"
            values = []
            try:
                for v in s.values:
                    if v is not None:
                        values.append(float(v))
                    else:
                        values.append(0.0)
            except Exception:
                values = [0.0] * len(categories) if categories else [0.0]

            # Extract series color (best effort)
            color = "#999999"
            try:
                fill = s.format.fill
                if fill.type is not None:
                    rgb = fill.fore_color.rgb
                    color = f"#{rgb}"
            except Exception:
                pass

            series_list.append(Series(
                name=name,
                values=values,
                color=color,
            ))

        if categories and series_list:
            confidence = "high"
        elif series_list:
            confidence = "medium"
        else:
            confidence = "low"

    except Exception:
        confidence = "low"
        categories = ["unknown"]
        series_list = [Series(name="unknown", values=[0.0], color="#999999")]

    return ChartData(categories=categories, series=series_list), confidence


def _extract_table_content(shape) -> tuple[list[str], list[list[str]]]:
    """Extract headers and rows from a python-pptx table shape."""
    headers = []
    rows = []
    try:
        table = shape.table
        n_rows = len(table.rows)
        n_cols = len(table.columns)

        if n_rows == 0 or n_cols == 0:
            return [], []

        # First row as headers
        for col_idx in range(n_cols):
            cell = table.cell(0, col_idx)
            headers.append(cell.text.strip())

        # Remaining rows as data
        for row_idx in range(1, n_rows):
            row = []
            for col_idx in range(n_cols):
                cell = table.cell(row_idx, col_idx)
                row.append(cell.text.strip())
            rows.append(row)

    except Exception:
        pass

    return headers, rows


def _cross_reference_source_data(
    categories: list[str],
    headline: str,
    source_data: Optional[dict],
) -> tuple[str, list[str], str]:
    """Cross-reference chart categories against source_data.json to find question codes.

    Returns (extraction_method, question_codes, confidence).
    """
    if source_data is None:
        return "", [], "low"

    # Search _codes section for matching descriptions
    codes_section = source_data.get("_codes", {})
    matched_codes = []

    for sheet_name, sheet_codes in codes_section.items():
        if not isinstance(sheet_codes, dict):
            continue
        for code, code_info in sheet_codes.items():
            if not isinstance(code_info, dict):
                continue
            desc = code_info.get("desc", "").lower()
            # Check if any category matches a sub-row description
            for cat in categories:
                if cat.lower() in desc or desc in cat.lower():
                    matched_codes.append(code)
                    break

    if matched_codes:
        return "question_code", matched_codes[:5], "medium"

    return "", [], "low"


def infer_untagged_shapes(
    pptx_path: str | Path,
    untagged_shapes: list,
    config_path: Optional[str | Path] = None,
    source_data_path: Optional[str | Path] = None,
) -> tuple[list[SlideSpec], InferenceSummary]:
    """Tier 2: Infer SlideSpecs for untagged shapes.

    Args:
        pptx_path: Path to PPTX (re-opened for detailed parsing).
        untagged_shapes: list of UntaggedShape from Tier 1.
        config_path: Optional YAML config for inference hints.
        source_data_path: Optional source_data.json for cross-referencing.

    Returns:
        (specs, summary)
    """
    pptx_path = Path(pptx_path)
    summary = InferenceSummary(total_untagged=len(untagged_shapes))

    # Load optional source data for cross-referencing
    source_data = None
    if source_data_path:
        sd_path = Path(source_data_path)
        if sd_path.exists():
            try:
                source_data = json.loads(sd_path.read_text(encoding="utf-8"))
            except Exception:
                pass

    # Re-open PPTX for detailed parsing
    prs = Presentation(str(pptx_path))

    # Group untagged shapes by slide
    by_slide: dict[int, list] = {}
    for us in untagged_shapes:
        by_slide.setdefault(us.slide_index, []).append(us)

    specs: list[SlideSpec] = []

    for slide_idx, shapes_on_slide in sorted(by_slide.items()):
        if slide_idx >= len(prs.slides):
            continue
        slide = prs.slides[slide_idx]

        # Extract headline for this slide
        headline_text = _extract_headline_from_slide_obj(slide)

        # Process chart shapes on this slide
        chart_shapes = [s for s in shapes_on_slide if s.has_chart]
        table_shapes = [s for s in shapes_on_slide if s.has_table and not s.has_chart]

        if not chart_shapes and not table_shapes:
            summary.textboxes_skipped += len(shapes_on_slide)
            continue

        components = []
        overall_confidence = "medium"

        for us in chart_shapes:
            summary.charts_inferred += 1
            # Find the actual shape in the slide
            chart_shape = _find_shape_by_name(slide, us.shape_name)
            if chart_shape is None or not chart_shape.has_chart:
                continue

            chart = chart_shape.chart
            chart_pattern, pattern_confidence = _classify_chart_type(chart)
            chart_data, data_confidence = _extract_chart_data(chart)

            position = Position(
                left=_emu_to_inches(us.left_emu),
                top=_emu_to_inches(us.top_emu),
                width=_emu_to_inches(us.width_emu),
                height=_emu_to_inches(us.height_emu),
            )

            components.append(ChartComponent(
                position=position,
                chart_pattern=chart_pattern,
                data=chart_data,
                chrome=ChartChrome(),
            ))

            # Update overall confidence
            if data_confidence == "low" or pattern_confidence == "low":
                overall_confidence = "low"

        for us in table_shapes:
            summary.tables_inferred += 1
            table_shape = _find_shape_by_name(slide, us.shape_name)
            if table_shape is None or not table_shape.has_table:
                continue

            headers, rows = _extract_table_content(table_shape)

            position = Position(
                left=_emu_to_inches(us.left_emu),
                top=_emu_to_inches(us.top_emu),
                width=_emu_to_inches(us.width_emu),
                height=_emu_to_inches(us.height_emu),
            )

            if headers and rows:
                components.append(ValueTableComponent(
                    position=position,
                    headers=headers,
                    rows=rows,
                ))
            elif not headers and rows:
                # Likely a label table
                labels = [r[0] for r in rows if r]
                components.append(LabelTableComponent(
                    position=position,
                    labels=labels,
                ))

        if not components:
            continue

        # Cross-reference categories against source data → lineage candidates
        extraction_method = ""
        question_codes = []
        candidates = []
        for comp in components:
            if isinstance(comp, ChartComponent) and comp.data.categories:
                extraction_method, question_codes, xref_conf = \
                    _cross_reference_source_data(
                        comp.data.categories, headline_text, source_data
                    )
                if xref_conf == "low":
                    overall_confidence = "low"
                # Build candidate from cross-reference result
                if extraction_method and question_codes:
                    candidates.append(DataLineageCandidate(
                        method=extraction_method,
                        question_codes=question_codes,
                        source_description=f"Cross-referenced from chart categories vs source_data.json",
                        confidence={"high": 0.9, "medium": 0.6, "low": 0.3}.get(xref_conf, 0.3),
                        reason=f"Category labels matched {len(question_codes)} question code(s)",
                    ))
                break

        # If we found a confident match, populate lineage; otherwise leave empty
        has_lineage = bool(extraction_method and question_codes and overall_confidence != "low")
        lineage = DataLineage(
            data_source=pptx_path.stem,
            extraction_method=extraction_method if has_lineage else "",
            question_codes=question_codes if has_lineage else [],
            source_file=str(pptx_path.name),
        )

        # Determine spec_completeness
        if has_lineage:
            completeness = "complete"
        elif components:
            completeness = "layout_complete_data_missing"
        else:
            completeness = "partial"

        spec = SlideSpec(
            slide_id=f"inferred_{slide_idx:03d}",
            slide_index=slide_idx,
            layout="observed_1chart_1table",
            headline=HeadlineSpec(text=headline_text),
            components=components,
            data_lineage=lineage,
            metadata=SlideMetadata(
                created_by="deck-reader-tier2",
                created_at=None,
                tier="2",
                confidence=overall_confidence,
                speaker_notes=f"Inferred from {len(chart_shapes)} charts, "
                              f"{len(table_shapes)} tables on slide {slide_idx + 1}.",
            ),
            spec_completeness=completeness,
            data_lineage_candidates=candidates,
            spec_version=SPEC_VERSION,
        )

        summary.confidence_counts[overall_confidence] += 1
        summary.specs_produced += 1
        specs.append(spec)

    # Write review doc for user confirmation
    _write_review_doc(pptx_path, specs, summary)

    return specs, summary


def _extract_headline_from_slide_obj(slide) -> str:
    """Extract headline text from a slide object."""
    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text or len(text) < 5:
            continue
        top_inches = _emu_to_inches(shape.top)
        if top_inches < 1.5:
            candidates.append((top_inches, len(text), text))

    if not candidates:
        return "Untitled Slide"

    candidates.sort(key=lambda c: (c[0], -c[1]))
    return candidates[0][2]


def _find_shape_by_name(slide, name: str):
    """Find a shape on a slide by name."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def _write_review_doc(
    pptx_path: Path,
    specs: list[SlideSpec],
    summary: InferenceSummary,
) -> None:
    """Write a review markdown doc for user confirmation of Tier 2 inferences."""
    review_path = pptx_path.parent / f"{pptx_path.stem}_inferred_specs.md"

    lines = [
        f"# Tier 2 Inference Review: {pptx_path.name}",
        "",
        f"**Total untagged shapes:** {summary.total_untagged}",
        f"**Charts inferred:** {summary.charts_inferred}",
        f"**Tables inferred:** {summary.tables_inferred}",
        f"**Specs produced:** {summary.specs_produced}",
        f"**Confidence breakdown:** "
        f"High={summary.confidence_counts['high']}, "
        f"Medium={summary.confidence_counts['medium']}, "
        f"Low={summary.confidence_counts['low']}",
        "",
        "---",
        "",
    ]

    for spec in specs:
        confidence = "unknown"
        if spec.metadata and spec.metadata.speaker_notes:
            if "Confidence: high" in spec.metadata.speaker_notes:
                confidence = "high"
            elif "Confidence: medium" in spec.metadata.speaker_notes:
                confidence = "medium"
            elif "Confidence: low" in spec.metadata.speaker_notes:
                confidence = "low"

        lines.append(f"## Slide {spec.slide_index + 1}: {spec.headline.text[:80]}")
        lines.append("")
        lines.append(f"| Field | Inferred Value | Confidence |")
        lines.append(f"|-------|---------------|------------|")
        lines.append(f"| slide_id | `{spec.slide_id}` | - |")
        lines.append(f"| layout | `{spec.layout}` | {confidence} |")

        for i, comp in enumerate(spec.components):
            if isinstance(comp, ChartComponent):
                lines.append(
                    f"| component[{i}] chart_pattern | "
                    f"`{comp.chart_pattern}` | {confidence} |"
                )
                lines.append(
                    f"| component[{i}] categories | "
                    f"`{comp.data.categories[:5]}` | {confidence} |"
                )
                lines.append(
                    f"| component[{i}] series_count | "
                    f"`{len(comp.data.series)}` | {confidence} |"
                )

        if spec.data_lineage:
            dl = spec.data_lineage
            if dl.extraction_method:
                lines.append(
                    f"| extraction_method | `{dl.extraction_method}` | {confidence} |"
                )
            if dl.question_codes:
                lines.append(
                    f"| question_codes | `{dl.question_codes}` | {confidence} |"
                )

        lines.append("")
        lines.append("**Action needed:** Confirm or correct the inferred values above.")
        lines.append("")
        lines.append("---")
        lines.append("")

    try:
        review_path.write_text("\n".join(lines), encoding="utf-8")
    except Exception:
        pass  # Non-fatal: review doc is a convenience, not a gate
