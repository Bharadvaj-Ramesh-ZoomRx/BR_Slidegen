"""
Slide Spec Validator
====================

Rejects incomplete specs loudly, before any renderer runs.

`validate_spec(spec)` returns a list of error strings (empty = valid).
For strict mode, use `SpecValidationError` — raise on first problem.

Checks performed:
  1. Required top-level fields (slide_id, slide_index, layout, headline, components)
  2. `layout` key exists in LAYOUTS{}
  3. Every component's `type` is in SUPPORTED_COMPONENT_TYPES
  4. Every chart component's `chart_pattern` is in SUPPORTED_CHART_PATTERNS
  5. Every Position has either explicit bbox OR a preset reference (not both, not neither)
  6. Preset references resolve against LAYOUTS{layout}{preset}
  7. Brand tokens resolve against BRAND{} if spec.brand is set
  8. ChartData series lengths match categories (except scatter patterns)
  9. DeltaColumn values length aligns with sibling label/chart categories
 10. Color tokens use valid syntax
"""
from __future__ import annotations

import re
from typing import Any

from .schema import (
    SlideSpec,
    ComponentSpec,
    ChartComponent,
    LabelTableComponent,
    DeltaColumnComponent,
    Position,
    SUPPORTED_CHART_PATTERNS,
    SUPPORTED_COMPONENT_TYPES,
)


# Color token syntax: either "#RRGGBB" or "{namespace.path.field}"
_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
_TOKEN_RE = re.compile(r"^\{[a-z][a-z_0-9]*(\.[a-z_0-9]+)+\}$", re.IGNORECASE)

# Patterns that don't use category-oriented data (Series.values is [(x, y), ...])
_SCATTER_PATTERNS = {"xy_scatter_abacus"}


class SpecValidationError(ValueError):
    """Raised when a spec fails validation in strict mode."""
    def __init__(self, errors: list[str]):
        self.errors = errors
        super().__init__(f"Spec validation failed with {len(errors)} error(s):\n  - " +
                         "\n  - ".join(errors))


def _is_valid_color_token(color: Any) -> bool:
    if not isinstance(color, str) or not color:
        return False
    return bool(_HEX_RE.match(color) or _TOKEN_RE.match(color))


def _validate_position(pos: Position, path: str, errors: list[str],
                       layouts: dict | None = None, layout_name: str | None = None) -> None:
    """A Position must be EITHER explicit bbox OR preset reference — never both, never neither."""
    explicit = pos.is_explicit()
    preset = pos.is_preset()
    if explicit and preset:
        errors.append(f"{path}: position has both explicit bbox AND preset — choose one")
    elif not explicit and not preset:
        errors.append(f"{path}: position has neither explicit bbox nor preset")
    elif explicit:
        # Bounds sanity (allow small negative bleed — real decks have shapes at -0.03")
        if pos.left is not None and (pos.left < -0.5 or pos.left > 14.0):
            errors.append(f"{path}: position.left={pos.left} outside slide bounds [-0.5, 14.0]")
        if pos.top is not None and (pos.top < -0.5 or pos.top > 8.0):
            errors.append(f"{path}: position.top={pos.top} outside slide bounds [-0.5, 8.0]")
        if pos.width is not None and pos.width <= 0:
            errors.append(f"{path}: position.width={pos.width} must be positive")
        if pos.height is not None and pos.height <= 0:
            errors.append(f"{path}: position.height={pos.height} must be positive")
    elif preset:
        # Preset resolution
        target_layout = pos.layout or layout_name
        if target_layout is None:
            errors.append(f"{path}: position.preset={pos.preset!r} set but no layout context")
        elif layouts is not None:
            if target_layout not in layouts:
                errors.append(f"{path}: position.layout={target_layout!r} not in LAYOUTS{{}}")
            elif pos.preset not in layouts[target_layout]:
                available = sorted(k for k in layouts[target_layout].keys() if "rect" in k or "left" in k)
                errors.append(
                    f"{path}: position.preset={pos.preset!r} not in LAYOUTS[{target_layout!r}]. "
                    f"Available: {available}"
                )


def _validate_chart_component(c: ChartComponent, path: str, errors: list[str]) -> None:
    if not c.chart_pattern:
        errors.append(f"{path}: chart_pattern is required")
    elif c.chart_pattern not in SUPPORTED_CHART_PATTERNS:
        errors.append(
            f"{path}: chart_pattern={c.chart_pattern!r} not in SUPPORTED_CHART_PATTERNS. "
            f"Supported: {SUPPORTED_CHART_PATTERNS}"
        )

    # Data shape
    if not c.data.categories:
        errors.append(f"{path}.data.categories is empty — chart needs at least one category")
    if not c.data.series:
        errors.append(f"{path}.data.series is empty — chart needs at least one series")

    n_cats = len(c.data.categories)
    is_scatter = c.chart_pattern in _SCATTER_PATTERNS
    for i, series in enumerate(c.data.series):
        s_path = f"{path}.data.series[{i}]"
        if not series.name:
            errors.append(f"{s_path}.name is empty")
        if not series.values:
            errors.append(f"{s_path}.values is empty")
        # Length check (skip for scatter — values are [x, y] pairs with different semantics)
        # Also skip when categories is a fallback placeholder from deck-reader extraction
        cats_are_placeholder = (c.data.categories == ["unknown"] or
                                c.data.categories == ["placeholder"])
        if not is_scatter and n_cats and not cats_are_placeholder and len(series.values) != n_cats:
            errors.append(
                f"{s_path}.values has {len(series.values)} entries; "
                f"expected {n_cats} to match categories"
            )
        if not _is_valid_color_token(series.color):
            errors.append(
                f"{s_path}.color={series.color!r} not a valid color token "
                f"(must be #RRGGBB or {{namespace.path}})"
            )

    # Chrome color tokens
    dl_color = c.chrome.data_labels.font_color
    if dl_color is not None and not _is_valid_color_token(dl_color):
        errors.append(f"{path}.chrome.data_labels.font_color={dl_color!r} invalid")


def _validate_delta_component(c: DeltaColumnComponent, path: str, errors: list[str]) -> None:
    if not c.values:
        errors.append(f"{path}.values is empty")
    if not _is_valid_color_token(c.positive_color):
        errors.append(f"{path}.positive_color={c.positive_color!r} invalid")
    if not _is_valid_color_token(c.negative_color):
        errors.append(f"{path}.negative_color={c.negative_color!r} invalid")
    if c.format not in ("delta_pp", "delta_pct", "delta_abs"):
        errors.append(f"{path}.format={c.format!r} must be one of delta_pp|delta_pct|delta_abs")


def _validate_label_table(c: LabelTableComponent, path: str, errors: list[str]) -> None:
    if not c.labels:
        errors.append(f"{path}.labels is empty")


def _cross_component_length_check(spec: SlideSpec, errors: list[str]) -> None:
    """If a slide has a chart + label_table + delta_column, their row counts should match.

    Skip for deck-reader specs — extracted components are independent shapes on the
    same slide and may have intentionally different row counts (e.g., label table
    includes header/footer rows that don't correspond to chart categories).
    """
    if spec.metadata and spec.metadata.created_by and "deck-reader" in spec.metadata.created_by:
        return
    charts = [c for c in spec.components if isinstance(c, ChartComponent)]
    label_tables = [c for c in spec.components if isinstance(c, LabelTableComponent)]
    deltas = [c for c in spec.components if isinstance(c, DeltaColumnComponent)]

    if not charts:
        return
    # Take the "main" chart — first one — as the reference for row count
    main = charts[0]
    n_rows = len(main.data.categories)
    if not n_rows:
        return

    for i, lt in enumerate(label_tables):
        if len(lt.labels) != n_rows:
            errors.append(
                f"components.label_table[{i}]: {len(lt.labels)} labels but main chart has "
                f"{n_rows} categories — lengths must match for table-based layouts"
            )
    for i, d in enumerate(deltas):
        if len(d.values) != n_rows:
            errors.append(
                f"components.delta_column[{i}]: {len(d.values)} values but main chart has "
                f"{n_rows} categories — lengths must match"
            )


def _try_load_layouts() -> dict | None:
    """Lazy import to avoid circular deps when spec_spec is used standalone."""
    try:
        from slidegen.pptx_utils.layout import LAYOUTS
        return LAYOUTS
    except Exception:
        return None


def _try_load_brand_keys() -> set[str] | None:
    try:
        from slidegen.pptx_utils.brand import BRAND
        return set(BRAND.keys())
    except Exception:
        return None


def validate_spec(spec: SlideSpec, strict: bool = False) -> list[str]:
    """Validate a SlideSpec. Returns list of error strings (empty = valid).

    If strict=True, raises SpecValidationError on any error.
    """
    errors: list[str] = []

    # Top-level required fields
    if not spec.slide_id:
        errors.append("slide_id is required")
    if spec.slide_index is None or spec.slide_index < 0:
        errors.append(f"slide_index={spec.slide_index} must be a non-negative int")
    if not spec.layout:
        errors.append("layout is required")
    if spec.headline is None or not spec.headline.text:
        errors.append("headline.text is required")
    if not spec.components:
        errors.append("components is empty — slide must have at least one component")

    # Layout existence — skip for deck-reader specs where all components have explicit positions
    # (the layout key is informational, not needed for rendering)
    all_explicit = all(c.position.is_explicit() for c in spec.components) if spec.components else False
    layouts = _try_load_layouts()
    if layouts is not None and spec.layout and spec.layout not in layouts and not all_explicit:
        errors.append(
            f"layout={spec.layout!r} not in LAYOUTS{{}}. "
            f"Available: {sorted(layouts.keys())}"
        )

    # Brand existence (only check if brand is provided and spec uses brand tokens)
    brand_keys = _try_load_brand_keys()
    if spec.brand and brand_keys is not None and spec.brand not in brand_keys:
        # Try normalized form before failing
        norm = spec.brand.upper().replace(" ", "_").replace("-", "_")
        if norm not in brand_keys:
            errors.append(
                f"brand={spec.brand!r} not in BRAND{{}}. "
                f"Nearest: {[b for b in brand_keys if norm[:3] in b][:5] or 'none'}"
            )

    # Each component
    for i, comp in enumerate(spec.components):
        path = f"components[{i}]"
        if comp.type not in SUPPORTED_COMPONENT_TYPES:
            errors.append(
                f"{path}.type={comp.type!r} not in {SUPPORTED_COMPONENT_TYPES}"
            )
            continue

        _validate_position(comp.position, f"{path}.position", errors,
                           layouts=layouts, layout_name=spec.layout)

        if isinstance(comp, ChartComponent):
            _validate_chart_component(comp, path, errors)
        elif isinstance(comp, DeltaColumnComponent):
            _validate_delta_component(comp, path, errors)
        elif isinstance(comp, LabelTableComponent):
            _validate_label_table(comp, path, errors)
        # Other component types: type + position is enough for v1.0

    # Cross-component coherence
    _cross_component_length_check(spec, errors)

    if strict and errors:
        raise SpecValidationError(errors)
    return errors
