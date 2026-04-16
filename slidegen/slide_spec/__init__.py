"""
slide_spec — the slide specification contract.

A *spec* is a complete, validated description of one slide. It is produced by
intelligent skills (viz-selector, headline-writer, slide-plan-generator-*) and
consumed by deterministic skills (slide-creator, slide-updater, slide-editor).

The spec is intentionally declarative and tool-agnostic:
- Stored as JSON on disk (reviewable, diffable, version-controllable)
- Serialized to / from Python dataclasses for type safety
- Validated by `validate_spec()` before any renderer touches it

See `schema.py` for dataclass definitions and `validator.py` for semantic checks.
"""
from __future__ import annotations

from .schema import (
    SlideSpec,
    HeadlineSpec,
    FooterSpec,
    ComponentSpec,
    ChartComponent,
    LabelTableComponent,
    ValueTableComponent,
    DeltaColumnComponent,
    CalloutComponent,
    ImageComponent,
    TextboxComponent,
    Position,
    Series,
    ChartData,
    ChartChrome,
    LegendSpec,
    DataLabelsSpec,
    AxisSpec,
    DataLineage,
    SlideMetadata,
    SPEC_VERSION,
    SUPPORTED_CHART_PATTERNS,
    SUPPORTED_COMPONENT_TYPES,
    load_spec,
    dump_spec,
)
from .validator import validate_spec, SpecValidationError

__all__ = [
    "SlideSpec",
    "HeadlineSpec",
    "FooterSpec",
    "ComponentSpec",
    "ChartComponent",
    "LabelTableComponent",
    "ValueTableComponent",
    "DeltaColumnComponent",
    "CalloutComponent",
    "ImageComponent",
    "TextboxComponent",
    "Position",
    "Series",
    "ChartData",
    "ChartChrome",
    "LegendSpec",
    "DataLabelsSpec",
    "AxisSpec",
    "DataLineage",
    "SlideMetadata",
    "SPEC_VERSION",
    "SUPPORTED_CHART_PATTERNS",
    "SUPPORTED_COMPONENT_TYPES",
    "load_spec",
    "dump_spec",
    "validate_spec",
    "SpecValidationError",
]
