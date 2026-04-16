"""
slide_plan_single.py — one ask → one SlideSpec.

Used by `add-slide-workflow` for client-followup + segment-comparison paths.
Composes viz-selector + layout-selector + headline-writer into a single
SlideSpec ready for spec-validator + slide-creator.

Contract:
    generate_single_slide(ask, data, brand, section, ...) -> SlideSpec
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from slidegen.slide_spec import (
    SlideSpec,
    HeadlineSpec,
    FooterSpec,
    ChartComponent,
    LabelTableComponent,
    DeltaColumnComponent,
    Position,
    ChartData,
    Series,
    ChartChrome,
    DataLabelsSpec,
    SlideMetadata,
    DataLineage,
    validate_spec,
)
from slidegen.viz_selector import select_chart_pattern
from slidegen.layout_selector import select_layout


@dataclass
class Ask:
    """Structured input to generate_single_slide."""
    question: str                       # the ask text (e.g. "How does recall differ by region?")
    metric_tag: Optional[str] = None    # e.g. "message_recall" — drives viz-selector
    q_type: Optional[str] = None        # fallback for viz-selector if no metric_tag
    segment: Optional[str] = None       # e.g. "Academic vs Community"
    hypothesis_refs: list[str] = field(default_factory=list)
    arc: Optional[str] = None
    suggested_headline: Optional[str] = None


def _build_chart_components(
    chart_pattern: str,
    data: dict,
    brand: Optional[str],
    layout: str,
    label_rect: tuple[float, float, float, float] | None = None,
    chart_rect: tuple[float, float, float, float] | None = None,
    delta_rect: tuple[float, float, float, float] | None = None,
):
    """Build (label_table?, chart, delta_column?) components with safe positions.

    Positions fall back to hardcoded canonical 1chart_1table coordinates if not
    provided (label left 0.30-3.10, chart 3.20-12.00, delta 12.10-13.10).
    """
    components = []
    cats = list(data.get("categories", []))
    series_list = data.get("series", [])

    # Default canonical positions (non-overlapping)
    label_rect = label_rect or (0.30, 1.85, 2.80, 4.50)
    chart_rect = chart_rect or (3.20, 1.85, 8.80, 4.50)
    delta_rect = delta_rect or (12.10, 1.85, 1.00, 4.50)

    # Label table (for bar_clustered_horizontal and other hide-cat-labels patterns)
    use_label_table = chart_pattern in (
        "bar_clustered_horizontal", "bar_stacked_100_horizontal",
        "clustered_bar", "stacked_bar", "single_bar",
    )
    if use_label_table and cats:
        components.append(LabelTableComponent(
            position=Position(left=label_rect[0], top=label_rect[1],
                              width=label_rect[2], height=label_rect[3]),
            labels=cats,
            alternating_rows=True,
            wrap=True,
        ))

    # Build chart component with spec's series shape
    spec_series = []
    primary_token = "{brand.primary_prior}" if brand else "#8FAADC"
    current_token = "{brand.primary_current}" if brand else "#F75824"
    secondary_token = "{brand.secondary}" if brand else "#7030A0"

    for i, s in enumerate(series_list):
        # Color: first=prior, last=current, middle=secondary
        if i == 0:
            color = primary_token
        elif i == len(series_list) - 1:
            color = current_token
        else:
            color = secondary_token
        spec_series.append(Series(
            name=s.get("name", f"Series {i+1}"),
            values=list(s.get("values", [])),
            color=color,
        ))

    chart_data_obj = ChartData(categories=cats, series=spec_series)

    chrome = ChartChrome(
        title=None,
        gridlines=False,
        hide_category_labels=use_label_table,
        data_labels=DataLabelsSpec(
            show=True,
            format="0%",
            position="inEnd" if use_label_table else "outEnd",
            font_color="#FFFFFF" if use_label_table else None,
        ),
    )
    # Line/scatter patterns want axis visible
    if chart_pattern in ("line_markers_trended", "xy_scatter_abacus"):
        chrome.hide_category_labels = False
        chrome.data_labels.position = "above"
        chrome.data_labels.font_color = None

    components.append(ChartComponent(
        position=Position(left=chart_rect[0], top=chart_rect[1],
                          width=chart_rect[2], height=chart_rect[3]),
        chart_pattern=chart_pattern,
        data=chart_data_obj,
        chrome=chrome,
    ))

    # Delta column if spec has 2 series (prior/current) + same-length values
    if (len(series_list) >= 2
        and use_label_table
        and len(series_list[0].get("values", [])) == len(series_list[-1].get("values", []))):
        prior_vals = series_list[0].get("values", [])
        current_vals = series_list[-1].get("values", [])
        delta_vals = [round(current_vals[i] - prior_vals[i], 4) for i in range(len(cats))]
        components.append(DeltaColumnComponent(
            position=Position(left=delta_rect[0], top=delta_rect[1],
                              width=delta_rect[2], height=delta_rect[3]),
            header="QoQ",
            values=delta_vals,
            format="delta_pp",
        ))

    return components


def generate_single_slide(
    ask: Ask,
    data: dict,
    brand: Optional[str] = None,
    section: Optional[str] = None,
    headline: Optional[str] = None,
    footer: Optional[str] = None,
    data_lineage: Optional[DataLineage] = None,
    slide_id: str = "zrx_new",
    slide_index: int = 0,
    positions: Optional[dict] = None,
) -> SlideSpec:
    """Build ONE SlideSpec from ONE ask.

    Args:
        ask: Ask dataclass with question + optional metric_tag / q_type / segment.
        data: {"categories": [...], "series": [{"name", "values"}, ...]} format.
        brand: BRAND key for color resolution.
        section: slide section (e.g. "Client Follow-up", "Segment Analysis").
        headline: optional pre-written headline; else headline-writer is used.
        footer: optional footer text.
        data_lineage: optional DataLineage for audit trail.
        slide_id, slide_index: deck positioning.
        positions: optional {"label_rect": (l,t,w,h), "chart_rect": ..., "delta_rect": ...}
            to override default component positions.

    Returns:
        A validated SlideSpec ready for slide-creator.
    """
    positions = positions or {}

    # 1. Pick chart pattern via viz-selector
    chart_pattern = select_chart_pattern(metric=ask.metric_tag, q_type=ask.q_type)
    if chart_pattern == "UNRESOLVED":
        raise ValueError(
            f"viz-selector couldn't resolve a chart pattern for ask={ask.question!r} "
            f"(metric={ask.metric_tag!r}, q_type={ask.q_type!r}). "
            f"Pass a recognized metric_tag or q_type."
        )

    # 2. Build chart components (label table + chart + delta column)
    components = _build_chart_components(
        chart_pattern=chart_pattern,
        data=data,
        brand=brand,
        layout="",  # set below
        label_rect=positions.get("label_rect"),
        chart_rect=positions.get("chart_rect"),
        delta_rect=positions.get("delta_rect"),
    )

    # 3. Pick layout based on actual component types
    component_types = [c.type for c in components]
    layout = select_layout(component_types)

    # 4. Headline — use provided, or ask.suggested_headline, or defer to headline-writer
    if headline is None:
        if ask.suggested_headline:
            headline = ask.suggested_headline
        else:
            # Lazy-import to avoid circular imports
            from slidegen.headline_writer import generate_headline
            # Build a minimal spec to feed headline-writer
            temp_spec = SlideSpec(
                slide_id=slide_id, slide_index=slide_index, layout=layout,
                brand=brand, section=section or ask.question,
                headline=HeadlineSpec(text="(pending)"),
                components=components,
            )
            try:
                headline = generate_headline(temp_spec, narrative_hint=ask.arc)
            except Exception:
                headline = ask.question  # fallback to the question itself

    # 5. Assemble final spec
    spec = SlideSpec(
        slide_id=slide_id,
        slide_index=slide_index,
        layout=layout,
        brand=brand,
        section=section,
        headline=HeadlineSpec(text=headline),
        footer=FooterSpec(text=footer) if footer else None,
        components=components,
        data_lineage=data_lineage,
        metadata=SlideMetadata(
            created_by="slide-plan-generator-single",
            hypothesis_refs=list(ask.hypothesis_refs),
            arc=ask.arc,
        ),
    )

    # 6. Validate on exit
    errors = validate_spec(spec)
    if errors:
        raise ValueError(
            f"generate_single_slide produced invalid spec: {errors}"
        )
    return spec
