"""
project_config.py — ProjectConfig schema and YAML loader.

Each project (J&J Rybrevant, Pfizer Ibrance, etc.) is described by a YAML file.
This module defines the schema and loads it into typed dataclasses.
"""

from __future__ import annotations
import os
import yaml
from dataclasses import dataclass, field
from typing import Optional
from pptx.dml.color import RGBColor


# ── Color parsing ────────────────────────────────────────────────────────────

def parse_color(hex_str: str) -> RGBColor:
    """Parse '#RRGGBB' or 'RRGGBB' → RGBColor."""
    h = hex_str.lstrip("#")
    if len(h) != 6 or not all(c in "0123456789abcdefABCDEF" for c in h):
        raise ValueError(f"Invalid hex color '{hex_str}' — expected '#RRGGBB' or 'RRGGBB'")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ── Dataclasses ──────────────────────────────────────────────────────────────

@dataclass
class BrandConfig:
    name: str                    # short label ("RYB+LAZ")
    full_name: str               # display name ("Rybrevant + Lazcluze")
    color_current: RGBColor      # Q4 / current wave color
    color_prior: RGBColor        # Q3 / prior wave color

    @classmethod
    def from_dict(cls, d: dict) -> BrandConfig:
        return cls(
            name=d["name"],
            full_name=d["full_name"],
            color_current=parse_color(d["color_current"]),
            color_prior=parse_color(d["color_prior"]),
        )


@dataclass
class SheetConfig:
    name: str                    # Excel sheet name
    q_prior_col: int             # 0-indexed column for prior wave total
    q_current_col: int           # 0-indexed column for current wave total
    code_col: int = 0            # column with question codes
    desc_col: int = 1            # column with descriptions


@dataclass
class SampleSize:
    prior: int
    current: int


@dataclass
class LabelShortcut:
    """Keyword-based label shortener: if kw1 in label (and kw2 in label), use short."""
    keywords: list[str]          # 1 or 2 keywords to match
    short: str                   # replacement label


@dataclass
class SegmentCutConfig:
    """Segment cut configuration: groupby splits data by segment, filter restricts it."""
    id: int
    name: str = ""
    mode: str = "groupby"                 # "groupby" | "filter"
    values: list[str] = field(default_factory=list)  # for filter: which values to keep


@dataclass
class SynapseConfig:
    """Synapse API connection parameters for data fetching."""
    api_url: str                              # e.g. "https://synapse.zoomrx.com/api"
    project_id: int
    survey_ids: list[int]
    deliverable_ids: list[int]
    segment_ids: list[int]
    reporting_plan_id: int = 0                # for time period / wave resolution
    wave_ids: list[int] = field(default_factory=list)  # explicit wave IDs (optional)
    multi_question_analysis_ids: list[int] = field(default_factory=list)
    virtual_question_analysis_ids: list[int] = field(default_factory=list)
    segments: list[SegmentCutConfig] = field(default_factory=list)  # segment cut configs
    poll_interval: int = 10                   # seconds between status checks
    max_wait: int = 300                       # max seconds to wait for banner plan generation


@dataclass
class AskConfig:
    """One ask = one slide (or slide group) to generate."""
    id: str                      # unique key ("ryb_mr_me", "rep_perf", ...)
    slide_type: str              # renderer name ("single_bar", "dual_bar", "clustered_compare", ...)
    headline: str                # slide headline (supports {{brand.name}} templates)
    section: str                 # section header bar text
    source_text: str             # footer source citation
    data_key: str                # key in extracted data dict
    brand: str = "primary"       # which brand this ask is for
    sort_by: Optional[str] = None
    sort_desc: bool = True
    template_slide: Optional[int] = None  # 1-based template slide to clone (future use)
    extra: dict = field(default_factory=dict)  # slide-type-specific params


@dataclass
class DataExtractionConfig:
    """How to extract one data block from the source."""
    id: str                      # maps to AskConfig.data_key
    method: str                  # "question_code" | "row_range" | "nested_ordinal" | "mock"
    sheet: str = ""              # "primary" | "competitor" | "analysis" (optional for mock)
    params: dict = field(default_factory=dict)  # method-specific params


@dataclass
class ProjectConfig:
    """Top-level project configuration."""
    name: str
    client: str
    period_current: str          # "Q4'25"
    period_prior: str            # "Q3'25"

    brands: dict[str, BrandConfig]  # "primary", "competitor"
    fonts: dict[str, str]        # "display", "body"
    sheets: dict[str, SheetConfig]  # "primary", "competitor", "analysis"
    sample_sizes: dict[str, SampleSize]

    extractions: list[DataExtractionConfig]
    asks: list[AskConfig]

    # Optional
    label_shortcuts: list[LabelShortcut] = field(default_factory=list)
    sections: list[dict] = field(default_factory=list)  # PPT sections: [{"name": ..., "start": ask_id}]
    wave: str = ""               # wave identifier (e.g. "PET_Q3Q4_2025")
    output_path: str = ""
    template_path: str = ""
    data_source_path: str = ""
    context_path: str = ""       # wave-versioned context folder (system-generated files)
    section_icon_path: str = ""  # small icon for section header bars
    raw_data_source_path: str = ""  # respondent-level raw data Excel (source_raw_data.xlsx)
    staleness_hours: int = 24   # JSON cache staleness warning threshold (hours)
    synapse_api_url: str = ""    # deprecated — use synapse.api_url instead
    synapse: Optional[SynapseConfig] = None  # Synapse API config for banner plan download

    @property
    def primary(self) -> BrandConfig:
        return self.brands["primary"]

    @property
    def competitor(self) -> BrandConfig:
        return self.brands["competitor"]

    def validate(self) -> list[str]:
        """Validate config consistency. Returns list of error messages (empty = valid).

        Checks:
        - ask.data_key references a known extraction.id (or mock/raw_aggregate)
        - ask.slide_type exists in RENDERERS
        - extraction.method is a known method
        - extraction.sheet exists in config.sheets (unless mock)
        - Required params per extraction method are present
        - Extra data_key references (left/right, primary_key/comp_key) are valid
        """
        from slidegen.pipeline.slide_renderers import RENDERERS

        errors = []

        # Check for duplicate extraction IDs
        seen_ids: dict[str, int] = {}
        for i, ex in enumerate(self.extractions):
            if ex.id in seen_ids:
                errors.append(
                    f"extractions[{i}] ({ex.id}): duplicate extraction ID "
                    f"(first defined at extractions[{seen_ids[ex.id]}])"
                )
            else:
                seen_ids[ex.id] = i

        extraction_ids = set(seen_ids.keys())
        known_methods = {
            "question_code", "multi_question_code", "row_range",
            "question_code_multi_col", "nested_ordinal", "mock",
            "raw_aggregate", "synapse_report", "synapse_raw",
        }
        sheet_keys = set(self.sheets.keys())

        # Validate extractions
        for i, ex in enumerate(self.extractions):
            if ex.method not in known_methods:
                errors.append(
                    f"extractions[{i}] ({ex.id}): unknown method '{ex.method}'"
                )
            if ex.method not in ("mock", "raw_aggregate", "synapse_report", "synapse_raw") and ex.sheet and ex.sheet not in sheet_keys:
                errors.append(
                    f"extractions[{i}] ({ex.id}): sheet '{ex.sheet}' not in config.sheets {sorted(sheet_keys)}"
                )
            # Required params per method
            params = ex.params
            if ex.method == "question_code" and "code" not in params:
                errors.append(f"extractions[{i}] ({ex.id}): method 'question_code' requires params.code")
            if ex.method == "multi_question_code" and "codes" not in params:
                errors.append(f"extractions[{i}] ({ex.id}): method 'multi_question_code' requires params.codes")
            if ex.method == "row_range":
                for rk in ("row_start", "row_end", "col_map"):
                    if rk not in params:
                        errors.append(f"extractions[{i}] ({ex.id}): method 'row_range' requires params.{rk}")
            if ex.method == "question_code_multi_col":
                if "code" not in params:
                    errors.append(f"extractions[{i}] ({ex.id}): method 'question_code_multi_col' requires params.code")
                if "columns" not in params:
                    errors.append(f"extractions[{i}] ({ex.id}): method 'question_code_multi_col' requires params.columns")
            if ex.method == "nested_ordinal":
                for rk in ("row_start", "row_end"):
                    if rk not in params:
                        errors.append(f"extractions[{i}] ({ex.id}): method 'nested_ordinal' requires params.{rk}")
            if ex.method == "synapse_report":
                for rk in ("analysis_id", "reporting_plan_id"):
                    if rk not in params:
                        errors.append(f"extractions[{i}] ({ex.id}): method 'synapse_report' requires params.{rk}")

        # Validate brand colors
        for brand_key, brand in self.brands.items():
            for color_field in ("color_current", "color_prior"):
                c = getattr(brand, color_field, None)
                if c is None:
                    errors.append(f"brands.{brand_key}: missing {color_field}")

        # Validate sample sizes are positive
        for ss_key, ss in self.sample_sizes.items():
            if ss.prior <= 0:
                errors.append(f"sample_sizes.{ss_key}: prior must be positive, got {ss.prior}")
            if ss.current <= 0:
                errors.append(f"sample_sizes.{ss_key}: current must be positive, got {ss.current}")

        # Validate asks
        for i, ask in enumerate(self.asks):
            if ask.slide_type not in RENDERERS:
                errors.append(
                    f"asks[{i}] ({ask.id}): unknown slide_type '{ask.slide_type}'"
                )
            # Check sort_by references a plausible data field
            if ask.sort_by:
                valid_suffixes = ("current", "prior", "desc", "delta", "total", "gap")
                parts = ask.sort_by.rsplit("_", 1)
                base_ok = ask.sort_by in valid_suffixes or (
                    len(parts) == 2 and parts[-1] in valid_suffixes
                )
                if not base_ok:
                    errors.append(
                        f"asks[{i}] ({ask.id}): sort_by '{ask.sort_by}' doesn't end with a "
                        f"recognized field suffix ({', '.join(valid_suffixes)})"
                    )
            # Check brand reference
            if ask.brand and ask.brand not in self.brands:
                errors.append(
                    f"asks[{i}] ({ask.id}): brand '{ask.brand}' not in config.brands {sorted(self.brands.keys())}"
                )
            # Check primary data_key
            if ask.data_key and ask.data_key not in extraction_ids:
                errors.append(
                    f"asks[{i}] ({ask.id}): data_key '{ask.data_key}' not found in extractions"
                )
            # Check extra data_key references
            extra = ask.extra or {}
            for nested_key in ("left", "right", "top", "bottom"):
                nested = extra.get(nested_key, {})
                if isinstance(nested, dict) and nested.get("data_key"):
                    dk = nested["data_key"]
                    if dk not in extraction_ids:
                        errors.append(
                            f"asks[{i}] ({ask.id}): extra.{nested_key}.data_key '{dk}' not found in extractions"
                        )
            for ek in ("primary_key", "comp_key"):
                dk = extra.get(ek)
                if dk and dk not in extraction_ids:
                    errors.append(
                        f"asks[{i}] ({ask.id}): extra.{ek} '{dk}' not found in extractions"
                    )
            # Validate renderer-specific required extra fields
            _RENDERER_REQUIRED_EXTRA = {
                "dual_bar_with_delta": ["left", "right"],
                "dual_bar_compare": ["left", "right"],
                "dual_bar_qoq": ["left", "right"],
                "dual_abacus": ["left", "right"],
                "hii_scorecard": ["categories"],
                "heatmap_table": ["columns"],
                "dual_doughnut": ["sections"],
                "trended_scorecard": ["panels"],
                "trended_activity": ["panels"],
                "two_section_bar": ["top", "bottom"],
            }
            required_extra = _RENDERER_REQUIRED_EXTRA.get(ask.slide_type, [])
            for req_key in required_extra:
                if req_key not in extra:
                    errors.append(
                        f"asks[{i}] ({ask.id}): slide_type '{ask.slide_type}' "
                        f"requires extra.{req_key}"
                    )

        return errors

    @property
    def font_display(self) -> str:
        return self.fonts.get("display", "Calibri")

    @property
    def font_body(self) -> str:
        return self.fonts.get("body", "Calibri")


# ── YAML Loader ──────────────────────────────────────────────────────────────

def _require(raw: dict, key: str, yaml_path: str, parent: str = ""):
    """Raise a clear error if a required key is missing from the config."""
    if key not in raw:
        context = f" in '{parent}'" if parent else ""
        raise ValueError(
            f"Missing required key '{key}'{context} in {yaml_path}"
        )
    return raw[key]


def _validate_ask(ask: dict, idx: int, yaml_path: str):
    """Validate required fields in an ask entry."""
    required = ("id", "slide_type", "headline", "section", "source_text", "data_key")
    for field_name in required:
        if field_name not in ask:
            raise ValueError(
                f"Missing required field '{field_name}' in asks[{idx}] "
                f"(id={ask.get('id', '?')}) in {yaml_path}"
            )


def _validate_extraction(ex: dict, idx: int, yaml_path: str):
    """Validate required fields in an extraction entry."""
    for field_name in ("id", "method"):
        if field_name not in ex:
            raise ValueError(
                f"Missing required field '{field_name}' in extractions[{idx}] "
                f"(id={ex.get('id', '?')}) in {yaml_path}"
            )
    if ex.get("method") not in ("mock", "synapse_report", "synapse_raw") and "sheet" not in ex:
        raise ValueError(
            f"Missing required field 'sheet' in extractions[{idx}] "
            f"(id={ex.get('id', '?')}) in {yaml_path}"
        )


def load_project_config(yaml_path: str) -> ProjectConfig:
    """Load a project YAML file into a ProjectConfig."""
    try:
        with open(yaml_path, "r", encoding="utf-8") as f:
            raw = yaml.safe_load(f)
    except yaml.YAMLError as e:
        raise ValueError(f"Malformed YAML in {yaml_path}: {e}") from e

    if not isinstance(raw, dict):
        raise ValueError(f"Config file must be a YAML mapping, got {type(raw).__name__}: {yaml_path}")

    project_dir = os.path.dirname(os.path.abspath(yaml_path))

    # Validate top-level required sections
    project = _require(raw, "project", yaml_path)
    for key in ("name", "client", "period_current", "period_prior"):
        _require(project, key, yaml_path, parent="project")

    # Wave identifier — interpolated into paths as {{wave}}
    wave = project.get("wave", "")

    # Brands
    brands_raw = _require(raw, "brands", yaml_path)
    brands = {}
    for key, bd in brands_raw.items():
        brands[key] = BrandConfig.from_dict(bd)

    # Sheets
    sheets_raw = _require(raw, "sheets", yaml_path)
    sheets = {}
    for key, sd in sheets_raw.items():
        sheets[key] = SheetConfig(
            name=sd["name"],
            q_prior_col=sd["q_prior_col"],
            q_current_col=sd["q_current_col"],
            code_col=sd.get("code_col", 0),
            desc_col=sd.get("desc_col", 1),
        )

    # Sample sizes
    sample_sizes = {}
    for key, ss in raw.get("sample_sizes", {}).items():
        sample_sizes[key] = SampleSize(prior=ss["prior"], current=ss["current"])

    # Label shortcuts
    label_shortcuts = []
    for ls in raw.get("label_shortcuts", []):
        label_shortcuts.append(LabelShortcut(
            keywords=ls["keywords"],
            short=ls["short"],
        ))

    # Extractions
    extractions = []
    for idx, ex in enumerate(raw.get("extractions", [])):
        _validate_extraction(ex, idx, yaml_path)
        extractions.append(DataExtractionConfig(
            id=ex["id"],
            method=ex["method"],
            sheet=ex.get("sheet", ""),
            params=ex.get("params", {}),
        ))

    # Asks
    asks = []
    for idx, ask in enumerate(raw.get("asks", [])):
        _validate_ask(ask, idx, yaml_path)
        asks.append(AskConfig(
            id=ask["id"],
            slide_type=ask["slide_type"],
            headline=ask["headline"],
            section=ask["section"],
            source_text=ask["source_text"],
            data_key=ask["data_key"],
            brand=ask.get("brand", "primary"),
            sort_by=ask.get("sort_by"),
            sort_desc=ask.get("sort_desc", True),
            template_slide=ask.get("template_slide"),
            extra=ask.get("extra", {}),
        ))

    # Resolve paths relative to project root, interpolating {{wave}}
    def _resolve_path(p: str) -> str:
        if not p:
            return p
        if wave:
            p = p.replace("{{wave}}", wave)
        if not os.path.isabs(p):
            p = os.path.join(project_dir, p)
        # Prevent path traversal outside project directory
        resolved = os.path.realpath(p)
        if not resolved.startswith(os.path.realpath(project_dir)):
            raise ValueError(
                f"Path '{p}' resolves outside project directory "
                f"(resolved: {resolved}, project: {project_dir})"
            )
        return p

    data_path = _resolve_path(raw.get("data_source_path", ""))
    tmpl_path = _resolve_path(raw.get("template_path", ""))
    out_path = _resolve_path(raw.get("output_path", ""))
    ctx_path = _resolve_path(raw.get("context_path", ""))
    icon_path = _resolve_path(raw.get("section_icon_path", ""))
    raw_data_path = _resolve_path(raw.get("raw_data_source_path", ""))

    # Sections (optional): [{"name": "...", "start": "ask_id"}, ...]
    sections = raw.get("sections", [])

    # Synapse config (optional)
    synapse_cfg = None
    synapse_raw = raw.get("synapse")
    if synapse_raw and isinstance(synapse_raw, dict):
        # Parse segment cut configs
        segment_cuts = []
        for sc in synapse_raw.get("segments", []):
            segment_cuts.append(SegmentCutConfig(
                id=sc["id"],
                name=sc.get("name", ""),
                mode=sc.get("mode", "groupby"),
                values=sc.get("values", []),
            ))

        synapse_cfg = SynapseConfig(
            api_url=synapse_raw["api_url"],
            project_id=synapse_raw["project_id"],
            survey_ids=synapse_raw["survey_ids"],
            deliverable_ids=synapse_raw["deliverable_ids"],
            segment_ids=synapse_raw.get("segment_ids", []),
            reporting_plan_id=synapse_raw.get("reporting_plan_id", 0),
            wave_ids=synapse_raw.get("wave_ids", []),
            multi_question_analysis_ids=synapse_raw.get("multi_question_analysis_ids", []),
            virtual_question_analysis_ids=synapse_raw.get("virtual_question_analysis_ids", []),
            segments=segment_cuts,
            poll_interval=synapse_raw.get("poll_interval", 10),
            max_wait=synapse_raw.get("max_wait", 300),
        )

    return ProjectConfig(
        name=project["name"],
        client=project["client"],
        period_current=project["period_current"],
        period_prior=project["period_prior"],
        brands=brands,
        fonts=raw.get("fonts", {"display": "Calibri", "body": "Calibri"}),
        sheets=sheets,
        sample_sizes=sample_sizes,
        extractions=extractions,
        asks=asks,
        label_shortcuts=label_shortcuts,
        sections=sections,
        wave=wave,
        output_path=out_path,
        template_path=tmpl_path,
        data_source_path=data_path,
        context_path=ctx_path,
        section_icon_path=icon_path,
        raw_data_source_path=raw_data_path,
        staleness_hours=project.get("staleness_hours", 24),
        synapse_api_url=raw.get("synapse_api_url", ""),
        synapse=synapse_cfg,
    )
