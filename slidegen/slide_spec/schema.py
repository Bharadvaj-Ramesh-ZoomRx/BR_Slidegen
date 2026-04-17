"""
Slide Spec Schema (v1.0)
========================

Declarative description of one slide. Produced by intelligent skills, consumed
by deterministic renderers. Never contains code, lambdas, or tool references.

## Fidelity axes (what slide-creator must get right)
1. **Layout** — shape positions match observed real-deck clusters (<=0.1" drift)
2. **Visualization** — chart pattern matches one of 10 CHART_PATTERNS keys
3. **Data** — values render to the exact numeric labels expected
4. Brand colors — resolvable from 4 sources, not hardcoded

## Color resolution sources (in priority order)
1. Explicit hex:            "#F75824"
2. BRAND{} token:           "{brand.primary_current}"
3. Context file token:      "{context.brand_palette.primary}"
4. Deck-reader extraction:  "{deck.slide_4.series_0.color}"
Tokens are resolved at render time by slide-creator, not at spec creation.

## Position sources
1. Explicit bbox:   {"left": 1.62, "top": 2.02, "width": 5.41, "height": 4.29}
2. Layout preset:   {"preset": "chart_rect", "layout": "observed_1chart_1table"}

## Spec JSON structure (top-level)
    {
      "spec_version": "1.0",
      "slide_id": "zrx_005",
      "slide_index": 4,
      "brand": "RYBREVANT",
      "section": "Message Recall",
      "layout": "observed_1chart_1table",
      "headline": {...},
      "subheadline": {...},
      "footer": {...},
      "components": [ {chart}, {label_table}, {delta_column}, ... ],
      "data_lineage": {...},
      "metadata": {...}
    }
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Optional, Union


SPEC_VERSION = "1.1"  # v1.1: added spec_completeness, data_lineage_candidates, metadata.tier/confidence


# Top 6 chart patterns by real-deck frequency (88% coverage) — Apr 15 analysis.
# These MUST be supported by slide-creator at client-delivery fidelity.
# Keys match CHART_PATTERNS in slidegen/pptx_utils/charts.py.
SUPPORTED_CHART_PATTERNS: tuple[str, ...] = (
    "bar_clustered_horizontal",     # 35% of real charts
    "xy_scatter_abacus",            # 17%
    "line_markers_trended",         # 14%
    "column_stacked_100_vertical",  # 11%
    "bar_stacked_100_horizontal",   #  7%
    "column_clustered_vertical",    #  7% (treated as bar_stacked equivalent)
    # Long-tail patterns (also supported, lower priority for regression):
    "doughnut_default",
    "single_bar",
    "clustered_bar",
    "stacked_bar",
)


SUPPORTED_COMPONENT_TYPES: tuple[str, ...] = (
    "chart",
    "label_table",
    "value_table",
    "delta_column",
    "callout",
    "image",
    "textbox",
)


# ─────────────────────────────────────────────────────────────────────────────
# Position
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Position:
    """Shape placement — either an explicit bbox or a named layout-preset rect.

    Exactly one form must be populated (validated in validator.py).
    Coordinates are in inches (slide is 13.333 × 7.500).

    Explicit:
        Position(left=1.62, top=2.02, width=5.41, height=4.29)

    Preset reference:
        Position(preset="chart_rect", layout="observed_1chart_1table")
        -> resolved at render time via LAYOUTS["observed_1chart_1table"]["chart_rect"]
    """
    # Explicit bbox (inches)
    left: Optional[float] = None
    top: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None

    # Or — named preset from LAYOUTS{}
    preset: Optional[str] = None       # e.g. "chart_rect", "primary_table_rect"
    layout: Optional[str] = None       # e.g. "observed_1chart_1table" (required if preset set)

    def is_explicit(self) -> bool:
        return all(v is not None for v in (self.left, self.top, self.width, self.height))

    def is_preset(self) -> bool:
        return self.preset is not None


# ─────────────────────────────────────────────────────────────────────────────
# Color token (kept as string — resolver lives in slide-creator, not here)
# ─────────────────────────────────────────────────────────────────────────────
#
# A ColorToken is just `str`. We document the grammar here; resolution is
# pattern-matching on prefix at render time. Validator checks syntax only.
#
# Valid forms:
#   "#F75824"                                 — explicit hex (case-insensitive, 6 chars)
#   "{brand.primary_current}"                 — BRAND{} lookup, requires spec.brand
#   "{brand.positive}" / "{brand.negative}"   — universal delta colors
#   "{context.brand_palette.primary}"         — resolved from project context file
#   "{deck.slide_N.series_M.color}"           — extracted from prior wave deck
ColorToken = str


# ─────────────────────────────────────────────────────────────────────────────
# Chart data
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Series:
    """One data series in a chart."""
    name: str                           # legend label
    values: list[float]                 # one per category (or [x, y] pairs for scatter)
    color: ColorToken                   # see ColorToken grammar above
    # Optional per-series overrides
    line_style: Optional[str] = None    # "solid" | "dashed" | "dotted"
    marker: Optional[str] = None        # "circle" | "square" | "none"
    data_labels: Optional[bool] = None  # override chart-level data_labels for this series


@dataclass
class ChartData:
    """Category-oriented chart data (bar/column/line)."""
    categories: list[str]               # e.g. row labels, message codes, quarters
    series: list[Series]


# ─────────────────────────────────────────────────────────────────────────────
# Chart chrome (defaults validated against real-deck frequencies)
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class LegendSpec:
    show: bool = False                  # 99% of real charts: no legend
    position: str = "bottom"            # "top" | "bottom" | "left" | "right"


@dataclass
class DataLabelsSpec:
    show: bool = True                   # most PET charts show labels
    format: str = "0%"                  # 96% of labels use "0%" (Apr 15 finding)
    position: str = "inEnd"             # "inEnd" | "outEnd" | "ctr" | "above"
    font_color: Optional[ColorToken] = None  # None -> inherit; common: "#FFFFFF" for inEnd


@dataclass
class AxisSpec:
    """Value-axis configuration (category axis is derived from ChartData.categories)."""
    min: Optional[float] = None
    max: Optional[float] = None
    format: str = "0%"                  # "0%" | "0" | "0.0" | "#,##0"
    show: bool = False                  # most real decks hide the value axis entirely
    invert: bool = False                # for horizontal bars with negative values


@dataclass
class ChartChrome:
    """Visual chrome around the chart's data region. Defaults mirror real-deck norms.

    Observed frequencies (Apr 15, 32 decks, 4354 charts):
      - no title:     99%  -> title=None
      - no legend:    99%  -> legend.show=False
      - no gridlines: 84%  -> gridlines=False
      - "0%" labels:  96%  -> data_labels.format="0%"

    DO NOT flip these defaults without evidence. If your workflow needs a legend,
    set legend.show=True explicitly in the spec — don't change the default.
    """
    title: Optional[str] = None         # None = no title (default per real decks)
    legend: LegendSpec = field(default_factory=LegendSpec)
    gridlines: bool = False
    data_labels: DataLabelsSpec = field(default_factory=DataLabelsSpec)
    value_axis: AxisSpec = field(default_factory=AxisSpec)
    hide_category_labels: bool = True   # 99% of bar_clustered_horizontal uses a separate label table
    # Bar/column chart geometry — None = use CHART_PATTERNS default
    gap_width: Optional[int] = None     # gapWidth (0-500). Extracted from OOXML by deck-reader.
    overlap: Optional[int] = None       # overlap (-100 to 100). Extracted from OOXML by deck-reader.


# ─────────────────────────────────────────────────────────────────────────────
# Components
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class ComponentSpec:
    """Base class. Never instantiated directly — always one of the subclasses below.

    Subclasses override `type` in their `__post_init__` — so callers can construct
    a ChartComponent(position=..., chart_pattern=...) without passing `type`.
    """
    position: Position = field(default_factory=lambda: Position())
    type: str = ""                      # set by subclass __post_init__


@dataclass
class ChartComponent(ComponentSpec):
    """A native PPT chart (python-pptx + lxml_helpers)."""
    chart_pattern: str = ""             # key into CHART_PATTERNS (see SUPPORTED_CHART_PATTERNS)
    data: ChartData = field(default_factory=lambda: ChartData(categories=[], series=[]))
    chrome: ChartChrome = field(default_factory=ChartChrome)

    def __post_init__(self) -> None:
        self.type = "chart"


@dataclass
class LabelTableComponent(ComponentSpec):
    """Label table (leftmost column of table-based bar layouts).

    Rows match the chart's categories 1:1. Used when chart.hide_category_labels=True.

    Defaults (9.0pt regular) match real-deck observations: label cells and delta
    cells are rendered at similar sizes so neither visually dominates. Drop to
    7.5pt only for dense decks where labels would otherwise wrap past 2 lines.
    """
    labels: list[str] = field(default_factory=list)
    alternating_rows: bool = True       # grey first, then white (matches real decks)
    font_size_pt: float = 9.0           # was 7.5 — looked too small beside 10-11pt delta cells
    font_name: Optional[str] = None     # None = inherit brand.font_body
    wrap: bool = True

    def __post_init__(self) -> None:
        self.type = "label_table"


@dataclass
class ValueTableComponent(ComponentSpec):
    """Multi-column numeric table (e.g. abacus value columns)."""
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)  # cells are already formatted strings
    alternating_rows: bool = True

    def __post_init__(self) -> None:
        self.type = "value_table"


@dataclass
class DeltaColumnComponent(ComponentSpec):
    """Single-column delta display (percentage-point deltas with color coding).

    Font defaults (10pt bold) balance against the 9pt label table default.
    Real-deck delta cells are 9-10pt; the slight upweight is the emphasis that
    makes the sign + magnitude pop without dwarfing the labels.
    """
    header: str = "QoQ"
    values: list[float] = field(default_factory=list)   # percentage points, signed
    format: str = "delta_pp"            # "delta_pp" | "delta_pct" | "delta_abs"
    positive_color: ColorToken = "{brand.positive}"
    negative_color: ColorToken = "{brand.negative}"
    alternating_rows: bool = True
    font_size_pt: float = 10.0          # was hardcoded 11pt in add_delta_col
    header_font_size_pt: Optional[float] = None   # defaults to font_size_pt
    font_name: Optional[str] = None     # None = inherit brand.font_body

    def __post_init__(self) -> None:
        self.type = "delta_column"


@dataclass
class CalloutComponent(ComponentSpec):
    """Data-driven annotation: theme tag + verbatim + attribution."""
    text: str = ""
    style: str = "default"              # "default" | "dashed" | "insight"
    theme: Optional[str] = None
    attribution: Optional[str] = None

    def __post_init__(self) -> None:
        self.type = "callout"


@dataclass
class ImageComponent(ComponentSpec):
    """Logo, screenshot, diagram."""
    source: str = ""                    # path relative to project or {{placeholder}}

    def __post_init__(self) -> None:
        self.type = "image"


@dataclass
class TextboxComponent(ComponentSpec):
    """Free-form text block."""
    text: str = ""
    font_size_pt: float = 10.0
    font_color: Optional[ColorToken] = None
    bold: bool = False
    italic: bool = False
    alignment: str = "left"             # "left" | "center" | "right"

    def __post_init__(self) -> None:
        self.type = "textbox"


# ─────────────────────────────────────────────────────────────────────────────
# Slide-level fields
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class HeadlineSpec:
    """Main headline — positioned per layout preset's headline slot."""
    text: str                           # may contain template vars: {{brand.name}}, {{period_current}}
    style: str = "default"              # "default" | "red_accent" | "neutral"
    color: Optional[ColorToken] = None  # override style default


@dataclass
class FooterSpec:
    """Footer — source footnote at slide bottom."""
    text: str
    style: str = "default"


@dataclass
class DataLineage:
    """Data provenance for audit + refresh workflows.

    Mirrors the Galen-PowerPoint Synapse Connector ReportConfig fields so that
    specs produced by SlideGen and specs reconstructed by deck-reader from
    existing decks are interchangeable.

    **Primary identifiers (Synapse-canonical)** — populated when the slide was
    authored through galen-powerpoint or through SlideGen's synapse tracks:
      - project_id, reporting_plan_id, analysis_ids — uniquely identify the
        data source in Synapse
      - survey_id — respondent-level fetches
      - segment_ids — segmentation applied at query time
      - static_time_period_ids OR dynamic_latest_n / include_live_wave — deliverables
      - analysis_type — e.g. "Question Analysis"

    **Legacy / Excel-path identifiers** — populated for decks built from local
    Excel files (Track B without Synapse):
      - data_source, extraction_method, question_codes, segment_codes, source_file

    **Audit fields** — updated on every refresh:
      - last_data_pull (ISO 8601), last_refresh_error (if non-empty, the last
        refresh failed), config_hash (invalidation key)

    deck-reader populates the Synapse fields from the shape tags
    (ReportConfigHash → Custom XML Part lookup). slide-updater writes back
    last_data_pull + last_refresh_error after re-fetch.
    """
    # Synapse-canonical identifiers (preferred when available)
    project_id: Optional[int] = None
    project_name: Optional[str] = None               # e.g. "AMG [ATU]: Repatha"
    reporting_plan_id: Optional[int] = None
    reporting_plan_name: Optional[str] = None         # e.g. "Quarterly"
    analysis_ids: list[int] = field(default_factory=list)
    analysis_names: list[str] = field(default_factory=list)  # human-readable per analysis
    survey_id: Optional[int] = None
    segment_ids: list[int] = field(default_factory=list)
    segment_names: list[str] = field(default_factory=list)   # human-readable per segment
    # Deliverables — either static list or dynamic "latest N + live"
    static_time_period_ids: list[int] = field(default_factory=list)
    static_time_period_names: list[str] = field(default_factory=list)  # e.g. ["Q2 2026", "Q1 2026"]
    dynamic_latest_n: Optional[int] = None
    include_live_wave: Optional[bool] = None
    analysis_type: Optional[str] = None              # e.g. "SINGLE_QUESTION"
    question_text: Optional[str] = None              # full question text from survey
    column_key_label_map: Optional[dict[str, str]] = None  # from ColumnKeyLabelMap shape tag

    # Legacy / Excel-path identifiers (used when no Synapse connection)
    data_source: str = ""                            # e.g. "Rybrevant/Q2_10"
    extraction_method: str = ""                      # "question_code" | "synapse_raw" | ...
    question_codes: list[str] = field(default_factory=list)
    segment_codes: list[str] = field(default_factory=list)
    source_file: Optional[str] = None                # "input/wave/Q1 2026/source_data.xlsx"

    # Audit fields (updated on every refresh)
    last_data_pull: Optional[str] = None             # ISO 8601
    last_refresh_error: Optional[str] = None         # None if last refresh succeeded
    config_hash: Optional[str] = None                # cache-invalidation key


@dataclass
class DataLineageCandidate:
    """A proposed data source for a slide whose lineage is unresolved.

    Populated by deck-reader Tier 2 when structural inference finds possible
    matches but can't confirm. The analyst reviews candidates and picks one
    (or provides their own) before the spec is frozen into config.yaml.
    """
    method: str = ""                     # "question_code" | "synapse_report" | "raw_aggregate"
    question_codes: list[str] = field(default_factory=list)
    source_description: str = ""         # human-readable: "Q2_10Z — Message Recall Top2Box"
    confidence: float = 0.0              # 0.0-1.0 — how sure we are this is right
    reason: str = ""                     # why this candidate was proposed

    # Optional Synapse identifiers (when inference suggests a Synapse source)
    analysis_id: Optional[int] = None
    reporting_plan_id: Optional[int] = None


@dataclass
class SlideMetadata:
    """Provenance + orchestration hints. Read by workflows, not by renderers."""
    created_by: Optional[str] = None    # e.g. "slide-plan-generator-hypothesis" | "deck-reader-tier1"
    created_at: Optional[str] = None    # ISO 8601
    hypothesis_refs: list[str] = field(default_factory=list)
    arc: Optional[str] = None           # "ACT NOW: Efficacy drift"
    role_in_arc: Optional[str] = None   # "evidence" | "closure" | "convergence"
    speaker_notes: Optional[str] = None # overrides auto-generated notes if set

    # deck-reader provenance (two-pass model)
    tier: Optional[str] = None          # "1" (Connector tag) | "2" (structural inference)
    confidence: Optional[str] = None    # "high" | "medium" | "low"
    original_tag_lineage: Optional[dict] = None  # preserved tag JSON when tag fails health check
    slide_layout_name: Optional[str] = None  # original slide layout name from source PPTX (for template matching)


# ─────────────────────────────────────────────────────────────────────────────
# Top-level SlideSpec
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class SlideSpec:
    """The contract between intelligent planning skills and deterministic renderers.

    A valid SlideSpec is sufficient input for slide-creator to render one
    client-ready slide with zero additional inference.

    ## Two-pass model (deck-reader → analyst review → config.yaml)

    When deck-reader creates specs from an existing PPTX:

    **Pass 1 (automatic):** Layout, components, headline, chart data, positions,
    brand detection — everything extractable from the PPTX itself. This always
    completes. The spec is renderable after Pass 1.

    **Pass 2 (data lineage):** Where did the numbers come from? For Synapse-
    connected shapes, tags provide this. For unconnected shapes, deck-reader
    proposes candidates in `data_lineage_candidates` and sets
    `spec_completeness="layout_complete_data_missing"`. The analyst reviews,
    picks the right candidate (or provides their own), and the spec becomes
    "complete" — ready to freeze into config.yaml.

    `spec_completeness` values:
      - "complete" — layout + data lineage both resolved. Renderable AND refreshable.
      - "layout_complete_data_missing" — renderable (can re-create the slide) but NOT
        refreshable (don't know where to get new data). Analyst must resolve.
      - "partial" — some components could not be extracted. Render will be approximate.
    """
    slide_id: str                       # unique within deck, e.g. "zrx_005"
    slide_index: int                    # 0-based position in deck
    layout: str                         # key into LAYOUTS{} (e.g. "observed_1chart_1table")
    headline: HeadlineSpec
    components: list[ComponentSpec] = field(default_factory=list)

    # Optional but recommended
    brand: Optional[str] = None         # key into BRAND{}; resolves color tokens
    section: Optional[str] = None       # e.g. "Message Recall"
    subheadline: Optional[HeadlineSpec] = None
    footer: Optional[FooterSpec] = None
    data_lineage: Optional[DataLineage] = None
    metadata: Optional[SlideMetadata] = None

    # Two-pass deck-reader fields
    spec_completeness: str = "complete"  # "complete" | "layout_complete_data_missing" | "partial"
    data_lineage_candidates: list[DataLineageCandidate] = field(default_factory=list)

    # Versioning
    spec_version: str = SPEC_VERSION


# ─────────────────────────────────────────────────────────────────────────────
# Serde
# ─────────────────────────────────────────────────────────────────────────────


# Registry: component type string -> dataclass
_COMPONENT_CLASSES: dict[str, type] = {
    "chart": ChartComponent,
    "label_table": LabelTableComponent,
    "value_table": ValueTableComponent,
    "delta_column": DeltaColumnComponent,
    "callout": CalloutComponent,
    "image": ImageComponent,
    "textbox": TextboxComponent,
}


def _component_from_dict(d: dict) -> ComponentSpec:
    ctype = d.get("type")
    cls = _COMPONENT_CLASSES.get(ctype)
    if cls is None:
        raise ValueError(f"Unknown component type: {ctype!r}")
    # Extract position
    pos = d.get("position", {})
    position = Position(**pos) if pos else Position()
    # Chart needs nested deserialization
    if ctype == "chart":
        data_d = d.get("data", {})
        series_list = [Series(**s) for s in data_d.get("series", [])]
        data = ChartData(categories=data_d.get("categories", []), series=series_list)
        chrome_d = d.get("chrome", {})
        chrome = ChartChrome(
            title=chrome_d.get("title"),
            legend=LegendSpec(**chrome_d["legend"]) if "legend" in chrome_d else LegendSpec(),
            gridlines=chrome_d.get("gridlines", False),
            data_labels=DataLabelsSpec(**chrome_d["data_labels"]) if "data_labels" in chrome_d else DataLabelsSpec(),
            value_axis=AxisSpec(**chrome_d["value_axis"]) if "value_axis" in chrome_d else AxisSpec(),
            hide_category_labels=chrome_d.get("hide_category_labels", True),
        )
        return ChartComponent(
            type="chart",
            position=position,
            chart_pattern=d.get("chart_pattern", ""),
            data=data,
            chrome=chrome,
        )
    # Other component types — pass through plain fields
    kwargs = {k: v for k, v in d.items() if k not in ("type", "position")}
    return cls(type=ctype, position=position, **kwargs)


def _strip_underscore_keys(d: Any) -> Any:
    """Recursively drop keys starting with '_' — used so spec JSON files can
    carry inline `_comment` / `_comment_*` fields without polluting the dataclass.
    """
    if isinstance(d, dict):
        return {k: _strip_underscore_keys(v) for k, v in d.items() if not k.startswith("_")}
    if isinstance(d, list):
        return [_strip_underscore_keys(x) for x in d]
    return d


def load_spec(source: Union[str, Path, dict]) -> SlideSpec:
    """Load a SlideSpec from a JSON file path, JSON string, or dict.

    Strips `_comment`-style keys before parsing so human-edited spec files can
    carry inline notes. Does NOT validate — call validate_spec() separately.
    """
    if isinstance(source, (str, Path)) and Path(str(source)).exists():
        data = json.loads(Path(source).read_text(encoding="utf-8"))
    elif isinstance(source, str):
        data = json.loads(source)
    elif isinstance(source, dict):
        data = source
    else:
        raise TypeError(f"Cannot load spec from {type(source).__name__}")

    data = _strip_underscore_keys(data)

    # Headline
    hd = data.get("headline", {})
    headline = HeadlineSpec(**hd)

    # Subheadline (optional)
    subheadline = None
    if data.get("subheadline"):
        subheadline = HeadlineSpec(**data["subheadline"])

    # Footer (optional)
    footer = None
    if data.get("footer"):
        footer = FooterSpec(**data["footer"])

    # Components
    components = [_component_from_dict(c) for c in data.get("components", [])]

    # Lineage + metadata
    lineage = DataLineage(**data["data_lineage"]) if data.get("data_lineage") else None
    metadata = SlideMetadata(**data["metadata"]) if data.get("metadata") else None

    # Data lineage candidates (two-pass model)
    candidates = []
    for cand_d in data.get("data_lineage_candidates", []):
        candidates.append(DataLineageCandidate(**cand_d))

    return SlideSpec(
        slide_id=data["slide_id"],
        slide_index=data["slide_index"],
        layout=data["layout"],
        headline=headline,
        components=components,
        brand=data.get("brand"),
        section=data.get("section"),
        subheadline=subheadline,
        footer=footer,
        data_lineage=lineage,
        metadata=metadata,
        spec_completeness=data.get("spec_completeness", "complete"),
        data_lineage_candidates=candidates,
        spec_version=data.get("spec_version", SPEC_VERSION),
    )


def dump_spec(spec: SlideSpec, path: Optional[Union[str, Path]] = None) -> str:
    """Serialize a SlideSpec to JSON. Writes to path if given; always returns the JSON string."""
    d = asdict(spec)
    # asdict recursively converts dataclasses — no post-processing needed for v1.0
    text = json.dumps(d, indent=2, ensure_ascii=False)
    if path is not None:
        Path(path).write_text(text, encoding="utf-8")
    return text
