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
    method: str                  # "question_code" | "row_range" | "nested_ordinal"
    sheet: str                   # "primary" | "competitor" | "analysis"
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
    section_icon_path: str = ""  # small icon for section header bars

    @property
    def primary(self) -> BrandConfig:
        return self.brands["primary"]

    @property
    def competitor(self) -> BrandConfig:
        return self.brands["competitor"]

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
    for field_name in ("id", "method", "sheet"):
        if field_name not in ex:
            raise ValueError(
                f"Missing required field '{field_name}' in extractions[{idx}] "
                f"(id={ex.get('id', '?')}) in {yaml_path}"
            )


def load_project_config(yaml_path: str) -> ProjectConfig:
    """Load a project YAML file into a ProjectConfig."""
    with open(yaml_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

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
            sheet=ex["sheet"],
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
        return p

    data_path = _resolve_path(raw.get("data_source_path", ""))
    tmpl_path = _resolve_path(raw.get("template_path", ""))
    out_path = _resolve_path(raw.get("output_path", ""))
    icon_path = _resolve_path(raw.get("section_icon_path", ""))

    # Sections (optional): [{"name": "...", "start": "ask_id"}, ...]
    sections = raw.get("sections", [])

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
        section_icon_path=icon_path,
    )
