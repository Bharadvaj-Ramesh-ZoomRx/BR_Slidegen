"""
slide_plan_refresh.py — diff plan generator for Workflow 2 (refresh deck).

Takes prior-wave specs (from deck-reader) + new wave data → a structured diff
plan telling refresh-deck-workflow which slides to update, add, or delete.

Contract:
    generate_refresh_plan(prior_specs, new_data, ...) -> RefreshPlan

Where RefreshPlan has:
    - update: list[SlideSpec] — slides whose data changed; updated specs ready
              for slide-updater + slide-creator.
    - add:    list[SlideSpec] — net-new slides where new-wave data warrants it.
    - delete: list[str]       — slide_ids to drop (deprecated metrics, removed
              segments).
    - rationale: dict[str, str] — one-line reason per changed slide_id.

Design notes:
- This module does NOT itself render slides or fetch data. It operates on
  already-extracted specs and pre-fetched data.
- It does NOT invoke headline-writer; that's slide-updater's job when it
  applies the plan to each slide. This module just produces the plan.
- Structural "add"s are conservative — only add a slide when the new wave
  introduces data that has no home in any prior slide.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from slidegen.slide_spec import SlideSpec, ChartComponent


# ─────────────────────────────────────────────────────────────────────────────
# Output types
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class RefreshPlan:
    """Structured diff plan — consumed by refresh-deck-workflow."""
    update: list[SlideSpec] = field(default_factory=list)
    add: list[SlideSpec] = field(default_factory=list)
    delete: list[str] = field(default_factory=list)   # slide_ids
    rationale: dict[str, str] = field(default_factory=dict)

    def summary(self) -> str:
        return (
            f"RefreshPlan(update={len(self.update)}, "
            f"add={len(self.add)}, delete={len(self.delete)})"
        )


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────


def _lineage_key(spec: SlideSpec) -> Optional[str]:
    """Build a stable key to match a spec's data_lineage against incoming data.

    Preference order:
      1. Synapse canonical: `reporting_plan_id + analysis_ids + segment_ids + deliverables`
      2. Excel-path: `data_source + extraction_method + question_codes`
      3. None if lineage is empty
    """
    if spec.data_lineage is None:
        return None
    dl = spec.data_lineage

    if dl.reporting_plan_id is not None and dl.analysis_ids:
        segs = ",".join(str(s) for s in sorted(dl.segment_ids or []))
        delivs = ",".join(str(d) for d in sorted(dl.static_time_period_ids or []))
        return (
            f"synapse:{dl.reporting_plan_id}:"
            f"{','.join(str(a) for a in sorted(dl.analysis_ids))}:"
            f"{segs}:{delivs}"
        )

    if dl.data_source or dl.extraction_method:
        codes = ",".join(sorted(dl.question_codes or []))
        return f"excel:{dl.data_source}:{dl.extraction_method}:{codes}"

    return None


def _get_chart(spec: SlideSpec) -> Optional[ChartComponent]:
    return next((c for c in spec.components if isinstance(c, ChartComponent)), None)


def _has_meaningful_change(
    prior_spec: SlideSpec,
    new_data: dict,
    min_delta_pp: float = 0.5,
) -> bool:
    """Return True if the new data differs meaningfully from the spec's existing values.

    Used to decide whether a slide's update warrants narrative reprocessing.
    """
    chart = _get_chart(prior_spec)
    if chart is None or not chart.data.series:
        return True  # no chart to compare — assume update

    # Compare current series (last) against new_data's current
    new_series = new_data.get("series") or []
    if not new_series:
        return False
    new_current_vals = new_series[-1].get("values") or []

    # Spec's current values (last series)
    old_current_vals = list(chart.data.series[-1].values)

    if len(new_current_vals) != len(old_current_vals):
        return True  # shape changed → definitely meaningful

    max_delta_pp = max(
        abs((a - b) * 100)
        for a, b in zip(new_current_vals, old_current_vals)
    )
    return max_delta_pp >= min_delta_pp


def _apply_new_data_to_spec(
    spec: SlideSpec, new_data: dict,
) -> SlideSpec:
    """Produce an updated spec with new_data applied to its chart.

    Mutates a copy; preserves layout, brand, section, colors, chrome.
    Recomputes deltas if there's a delta_column.
    Does NOT regenerate the headline — that's slide-updater's job per
    the headline-refresh policy (PRD §6.9).
    """
    import copy
    updated = copy.deepcopy(spec)

    chart = _get_chart(updated)
    if chart is None:
        return updated

    new_categories = new_data.get("categories") or chart.data.categories
    new_series_data = new_data.get("series") or []

    # Update categories (order may shift per new wave)
    chart.data.categories = list(new_categories)

    # Update series values — preserve series colors + names from spec (brand/display),
    # only swap values.
    for i, existing_series in enumerate(chart.data.series):
        if i < len(new_series_data):
            existing_series.values = list(new_series_data[i].get("values", existing_series.values))
            # Series name updates (e.g. "Q4 '25" → "Q1 '26") come from new_data
            if "name" in new_series_data[i]:
                existing_series.name = new_series_data[i]["name"]

    # Recompute delta column if present (current - prior)
    from slidegen.slide_spec import DeltaColumnComponent
    for component in updated.components:
        if isinstance(component, DeltaColumnComponent):
            if len(chart.data.series) >= 2:
                prior = chart.data.series[0].values
                current = chart.data.series[-1].values
                n = min(len(prior), len(current))
                component.values = [round(current[i] - prior[i], 4) for i in range(n)]

    return updated


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────


def generate_refresh_plan(
    prior_specs: list[SlideSpec],
    new_data: dict[str, dict],
    feedback: Optional[dict] = None,
    narrative_threads: Optional[dict] = None,
    new_period_prior: Optional[str] = None,
    new_period_current: Optional[str] = None,
) -> RefreshPlan:
    """Generate a structured refresh plan: update / add / delete.

    Args:
        prior_specs: list of SlideSpec from the prior wave's deck (via deck-reader).
        new_data: dict mapping a lineage key (from `_lineage_key()`) to the
            new wave's data payload `{"categories": [...], "series": [{"name", "values"}, ...]}`.
        feedback: optional client feedback that should inform deletions/additions
            (e.g. {"remove_slides": [slide_id, ...], "add_topics": [str, ...]}).
        narrative_threads: optional pre-built narrative arcs from sfea-insight-writer —
            used to inform which NEW slides to propose.
        new_period_prior, new_period_current: period labels for the new wave, used
            to annotate rationale messages.

    Returns:
        A RefreshPlan.
    """
    plan = RefreshPlan()
    period_note = ""
    if new_period_current and new_period_prior:
        period_note = f" ({new_period_prior} → {new_period_current})"

    # Map prior specs by lineage key for O(1) lookup
    prior_by_key: dict[str, SlideSpec] = {}
    for spec in prior_specs:
        key = _lineage_key(spec)
        if key:
            prior_by_key[key] = spec

    matched_keys: set[str] = set()
    feedback = feedback or {}
    removal_ids = set(feedback.get("remove_slides", []))

    # 1. UPDATE pass — for each prior spec with matching new data, apply update
    for spec in prior_specs:
        if spec.slide_id in removal_ids:
            plan.delete.append(spec.slide_id)
            plan.rationale[spec.slide_id] = f"client feedback requested removal{period_note}"
            continue

        key = _lineage_key(spec)
        if key is None:
            # No lineage to match — preserve as-is (non-data slide: cover, ES, etc.)
            plan.update.append(spec)
            plan.rationale[spec.slide_id] = f"structural slide (cover/divider/ES) — carried forward"
            continue

        incoming = new_data.get(key)
        if incoming is None:
            # Lineage was present but data is missing — flag for user
            plan.delete.append(spec.slide_id)
            plan.rationale[spec.slide_id] = (
                f"data for this slide's lineage not available in new wave{period_note} "
                f"— confirm deletion or set aside for manual update"
            )
            continue

        matched_keys.add(key)
        if _has_meaningful_change(spec, incoming):
            updated = _apply_new_data_to_spec(spec, incoming)
            plan.update.append(updated)
            plan.rationale[spec.slide_id] = f"data refreshed with new wave{period_note}"
        else:
            # Values are substantially unchanged — still update for audit stamp, but note
            updated = _apply_new_data_to_spec(spec, incoming)
            plan.update.append(updated)
            plan.rationale[spec.slide_id] = (
                f"minimal change vs prior wave{period_note} — refreshed for audit trail"
            )

    # 2. ADD pass — new wave data with lineage keys not matched by any prior spec
    unmatched_keys = set(new_data.keys()) - matched_keys
    for key in unmatched_keys:
        # Future: use narrative_threads to decide whether to propose a new slide.
        # For now: flag in rationale; caller (workflow) can decide to build the spec.
        placeholder_id = f"new_{key.replace(':', '_').replace(',', '-')[:32]}"
        plan.rationale[placeholder_id] = (
            f"new wave introduces lineage {key!r} with no prior-wave slide — "
            f"consider adding via viz-selector + layout-selector + slide-creator"
        )
        # Note: we don't build a full SlideSpec here — that requires viz/layout/headline
        # generation which lives upstream in refresh-deck-workflow's add branch.

    # 3. DELETE pass — already handled inline above (removal + data-not-available cases)

    return plan
