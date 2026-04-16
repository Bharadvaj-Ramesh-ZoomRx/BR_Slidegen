"""
slide_updater.py — Data refresh primitive for SlideSpec.

Takes a SlideSpec + new data (dict keyed by data_lineage identifiers),
produces an updated SlideSpec where chart series values + delta_column
values + data_labels are refreshed. Preserves layout, headline style,
brand colors, metadata. Writes back last_data_pull + clears last_refresh_error.
Spec-validates before returning.

Usage:
    from slidegen.slide_updater import update_slide_data

    updated_spec = update_slide_data(spec, new_data)
"""
from __future__ import annotations

import copy
from datetime import datetime, timezone
from typing import Any, Optional

from slidegen.slide_spec.schema import (
    SlideSpec, ChartComponent, DeltaColumnComponent,
    ValueTableComponent, LabelTableComponent,
    DataLineage, Series, ChartData,
    dump_spec,
)
from slidegen.slide_spec.validator import validate_spec, SpecValidationError


class SlideUpdateError(Exception):
    """Raised when a slide update fails."""
    pass


def update_slide_data(
    spec: SlideSpec,
    new_data: dict[str, Any],
    timestamp: Optional[str] = None,
) -> SlideSpec:
    """Refresh a SlideSpec with new data values.

    The new_data dict supports these top-level keys:

    For chart components:
        "categories": list[str]          — new category labels (optional, preserves if absent)
        "series": list[dict]             — new series data, each dict has:
            "name": str                  — series name (must match existing or be positional)
            "values": list[float]        — new values
            "color": str (optional)      — color override (if absent, preserves original)

    For delta_column components:
        "delta_values": list[float]      — new delta values

    For value_table components:
        "table_headers": list[str]       — new headers (optional)
        "table_rows": list[list[str]]    — new row data

    For label_table components:
        "labels": list[str]              — new label text

    Args:
        spec: The source SlideSpec to update.
        new_data: Dict with new data values.
        timestamp: ISO 8601 timestamp for last_data_pull (default: now).

    Returns:
        Updated SlideSpec (deep copy — original is not mutated).

    Raises:
        SpecValidationError: If the updated spec fails validation.
        SlideUpdateError: If the update logic encounters an error.
    """
    # Validate input spec
    errors = validate_spec(spec)
    if errors:
        raise SpecValidationError(errors)

    # Deep copy to avoid mutating the original
    updated = copy.deepcopy(spec)

    # Track what was updated for audit
    changes = []

    # Update chart components
    new_categories = new_data.get("categories")
    new_series = new_data.get("series")

    for i, comp in enumerate(updated.components):
        if isinstance(comp, ChartComponent):
            if new_categories is not None:
                if len(new_categories) == 0:
                    raise SlideUpdateError(
                        f"components[{i}]: new categories list is empty"
                    )
                comp.data.categories = list(new_categories)
                changes.append(f"components[{i}].data.categories")

            if new_series is not None:
                _update_chart_series(comp, new_series, i)
                changes.append(f"components[{i}].data.series")

        elif isinstance(comp, DeltaColumnComponent):
            new_deltas = new_data.get("delta_values")
            if new_deltas is not None:
                comp.values = [float(v) for v in new_deltas]
                changes.append(f"components[{i}].values")

        elif isinstance(comp, ValueTableComponent):
            new_headers = new_data.get("table_headers")
            new_rows = new_data.get("table_rows")
            if new_headers is not None:
                comp.headers = list(new_headers)
                changes.append(f"components[{i}].headers")
            if new_rows is not None:
                comp.rows = [list(row) for row in new_rows]
                changes.append(f"components[{i}].rows")

        elif isinstance(comp, LabelTableComponent):
            new_labels = new_data.get("labels")
            if new_labels is not None:
                comp.labels = list(new_labels)
                changes.append(f"components[{i}].labels")

    # Update data lineage audit fields
    if updated.data_lineage is None:
        updated.data_lineage = DataLineage()

    ts = timestamp or datetime.now(timezone.utc).isoformat()
    updated.data_lineage.last_data_pull = ts
    updated.data_lineage.last_refresh_error = None  # Clear on successful refresh

    # Validate the updated spec
    validate_spec(updated, strict=True)

    return updated


def _update_chart_series(
    comp: ChartComponent,
    new_series: list[dict],
    comp_index: int,
) -> None:
    """Update chart series data.

    Matching strategy:
    1. If new_series[i] has "name", match by name against existing series.
    2. If no name match, use positional matching (index-based).
    3. If new_series has more items than existing, append new series.
    4. If new_series has fewer, keep the existing extras unchanged.
    """
    existing_by_name = {s.name: (idx, s) for idx, s in enumerate(comp.data.series)}

    for j, ns in enumerate(new_series):
        if not isinstance(ns, dict):
            raise SlideUpdateError(
                f"components[{comp_index}].data.series[{j}]: "
                f"expected dict, got {type(ns).__name__}"
            )

        name = ns.get("name")
        values = ns.get("values")
        color = ns.get("color")

        # Find target series
        target = None
        if name and name in existing_by_name:
            _, target = existing_by_name[name]
        elif j < len(comp.data.series):
            target = comp.data.series[j]
        else:
            # Append new series
            if values is None:
                raise SlideUpdateError(
                    f"components[{comp_index}].data.series[{j}]: "
                    f"new series must have 'values'"
                )
            new_s = Series(
                name=name or f"Series {j}",
                values=[float(v) for v in values],
                color=color or "#999999",
            )
            comp.data.series.append(new_s)
            continue

        # Update existing series
        if values is not None:
            target.values = [float(v) for v in values]
        if name is not None:
            target.name = name
        if color is not None:
            target.color = color


def refresh_slide_error(
    spec: SlideSpec,
    error_message: str,
    timestamp: Optional[str] = None,
) -> SlideSpec:
    """Mark a spec as having a refresh error.

    Used when the data fetch fails — records the error without changing data.

    Args:
        spec: The source SlideSpec.
        error_message: Human-readable error description.
        timestamp: ISO 8601 timestamp (default: now).

    Returns:
        Updated SlideSpec with last_refresh_error set.
    """
    updated = copy.deepcopy(spec)
    if updated.data_lineage is None:
        updated.data_lineage = DataLineage()

    ts = timestamp or datetime.now(timezone.utc).isoformat()
    updated.data_lineage.last_data_pull = ts
    updated.data_lineage.last_refresh_error = error_message

    return updated
