"""Deck-level comparison for Eval #3 (refresh execution).

Compares a source PPTX against a refreshed PPTX at the component level.
For every chart and table that can be matched by (slide_index, position),
records three independent match signals:

    categories_match       — row labels (e.g. message names) are identical
    series_names_match     — series labels (e.g. brand names) are identical
    values_match           — series values are numerically equal (within tolerance)

Where series_names/values are unavailable (e.g. tables), the corresponding
signals are None and excluded from totals.

This deepens Vijay's Stage 3 comparison (which is categories-only). Categories-
only can silently miss corruption — a chart whose wave filter was dropped may
keep its row labels but show wrong values. The values_match signal catches that.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

VALUE_TOLERANCE = 1e-6  # absolute float tolerance for value comparison
POSITION_TOLERANCE = 0.1  # inches — max drift for source/refreshed shape matching


@dataclass
class ComponentMatch:
    slide_idx: int
    position: str  # "(left_in, top_in)" rounded to 2 decimals
    kind: str  # "chart" or "table"
    categories_match: bool | None = None
    series_names_match: bool | None = None
    values_match: bool | None = None
    note: str = ""  # e.g. "no matching shape in refreshed" or "extraction error"


@dataclass
class DeckMatchReport:
    source_deck: str
    refreshed_deck: str
    components: list[ComponentMatch] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "source_deck": self.source_deck,
            "refreshed_deck": self.refreshed_deck,
            "totals": self.totals(),
            "components": [c.__dict__ for c in self.components],
        }

    def totals(self) -> dict:
        total = len(self.components)
        cat_ok = sum(1 for c in self.components if c.categories_match is True)
        srs_ok = sum(1 for c in self.components if c.series_names_match is True)
        val_ok = sum(1 for c in self.components if c.values_match is True)
        errors = sum(1 for c in self.components if c.note)
        return {
            "components_compared": total,
            "categories_match": cat_ok,
            "series_names_match": srs_ok,
            "values_match": val_ok,
            "components_with_errors": errors,
        }


def _pos(shape) -> tuple[float, float]:
    left_in = round((shape.left or 0) / 914400, 2)
    top_in = round((shape.top or 0) / 914400, 2)
    return (left_in, top_in)


def _find_match(src_shape, candidates) -> object | None:
    """Find the candidate shape closest to src_shape by position, within tolerance."""
    sl, st = _pos(src_shape)
    best, best_dist = None, float("inf")
    for c in candidates:
        cl, ct = _pos(c)
        dist = abs(sl - cl) + abs(st - ct)
        if dist < best_dist:
            best_dist, best = dist, c
    if best is not None and best_dist <= POSITION_TOLERANCE:
        return best
    return None


def _extract_chart_data(shape) -> tuple[list[str], list[tuple[str, list[float | None]]]]:
    """Return (categories, [(series_name, values), ...])."""
    plot = shape.chart.plots[0]
    categories = [str(c) for c in (plot.categories or [])]
    series = []
    for s in plot.series:
        name = s.name if s.name else ""
        values = []
        for v in s.values:
            values.append(None if v is None else float(v))
        series.append((name, values))
    return categories, series


def _values_equal(a: list[float | None], b: list[float | None]) -> bool:
    if len(a) != len(b):
        return False
    for x, y in zip(a, b):
        if x is None and y is None:
            continue
        if x is None or y is None:
            return False
        if abs(x - y) > VALUE_TOLERANCE:
            return False
    return True


def _extract_table_rows(shape) -> list[list[str]]:
    tbl = shape.table
    rows = []
    for row in tbl.rows:
        rows.append([cell.text_frame.text.strip() for cell in row.cells])
    return rows


def compare_decks(
    source_path: Path,
    refreshed_path: Path,
    connected_slide_indices: set[int] | None = None,
    chart_modes: dict[tuple[int, str], str] | None = None,
) -> DeckMatchReport:
    """Run component-level comparison between source and refreshed decks.

    `connected_slide_indices`: when provided, only slides whose 0-based index is
    in this set are compared. Pass the indices from the Step 1 spec golden to
    restrict comparison to connected (Connector-tagged) slides only — non-connected
    slides are never refreshed and would trivially inflate match counts.

    `chart_modes`: maps (slide_idx, chart_shape_name) -> 'static' | 'dynamic'.
    Static charts are scored with strict identity (source values == refreshed
    values). Dynamic charts (those tagged with dynamic_latest_n>0 and no static
    pin) are expected to auto-roll forward — their wave labels legitimately
    differ after refresh, so we score on STRUCTURE only: same cat count and
    series count = match. When `chart_modes` is None, all charts are treated
    as static (current strict behavior).
    """
    source = Presentation(str(source_path))
    refreshed = Presentation(str(refreshed_path))

    report = DeckMatchReport(
        source_deck=source_path.name,
        refreshed_deck=refreshed_path.name,
    )

    n_slides = min(len(source.slides), len(refreshed.slides))
    for slide_idx in range(n_slides):
        if connected_slide_indices is not None and slide_idx not in connected_slide_indices:
            continue
        src_slide = source.slides[slide_idx]
        ref_slide = refreshed.slides[slide_idx]

        # ── Charts ──
        src_charts = [s for s in src_slide.shapes if s.has_chart]
        ref_charts = [s for s in ref_slide.shapes if s.has_chart]

        for src in src_charts:
            sl, st = _pos(src)
            match = _find_match(src, ref_charts)
            comp = ComponentMatch(
                slide_idx=slide_idx,
                position=f"({sl},{st})",
                kind="chart",
            )
            if match is None:
                comp.note = "no matching refreshed chart at this position"
                report.components.append(comp)
                continue
            try:
                src_cats, src_series = _extract_chart_data(src)
                ref_cats, ref_series = _extract_chart_data(match)

                src_names = [name for name, _ in src_series]
                ref_names = [name for name, _ in ref_series]

                mode = (
                    chart_modes.get((slide_idx, src.name or ""), "static")
                    if chart_modes is not None else "static"
                )

                if mode == "dynamic":
                    # Dynamic charts auto-roll forward — wave labels in
                    # cats/series legitimately differ after refresh. Score
                    # each dimension on STRUCTURE only:
                    #   categories_match  = same cat count
                    #   series_names_match = same series count
                    #   values_match      = both counts preserved
                    cats_ok = len(src_cats) == len(ref_cats)
                    series_ok = len(src_series) == len(ref_series)
                    comp.categories_match = cats_ok
                    comp.series_names_match = series_ok
                    comp.values_match = cats_ok and series_ok
                else:
                    # Static charts: strict identity (current behavior).
                    comp.categories_match = src_cats == ref_cats
                    comp.series_names_match = src_names == ref_names
                    if len(src_series) == len(ref_series):
                        comp.values_match = all(
                            _values_equal(sv, rv)
                            for (_, sv), (_, rv) in zip(src_series, ref_series)
                        )
                    else:
                        comp.values_match = False
            except Exception as e:
                comp.note = f"chart extraction error: {type(e).__name__}: {e}"
            report.components.append(comp)

        # ── Tables ──
        src_tables = [s for s in src_slide.shapes if s.has_table]
        ref_tables = [s for s in ref_slide.shapes if s.has_table]

        for src in src_tables:
            sl, st = _pos(src)
            match = _find_match(src, ref_tables)
            comp = ComponentMatch(
                slide_idx=slide_idx,
                position=f"({sl},{st})",
                kind="table",
            )
            if match is None:
                comp.note = "no matching refreshed table at this position"
                report.components.append(comp)
                continue
            try:
                src_rows = _extract_table_rows(src)
                ref_rows = _extract_table_rows(match)
                # Treat table cell-text as "values" for match tracking
                comp.values_match = src_rows == ref_rows
                # Categories/series don't apply to tables; leave as None
            except Exception as e:
                comp.note = f"table extraction error: {type(e).__name__}: {e}"
            report.components.append(comp)

    return report


def compare_reports(actual: dict, expected: dict) -> list[str]:
    """Diff two match reports and return drift messages.

    Checks:
    - totals did not regress (each "*_match" count must be >= expected)
    - every component that was True in expected is still True in actual
      (catches silent per-component regressions even when totals hold)
    """
    errors: list[str] = []

    # Totals regression
    a_tot = actual["totals"]
    e_tot = expected["totals"]
    for key in ("categories_match", "series_names_match", "values_match"):
        if a_tot.get(key, 0) < e_tot.get(key, 0):
            errors.append(
                f"totals.{key}: regressed from {e_tot[key]} → {a_tot[key]}"
            )

    # Per-component regression (keyed by slide_idx + position + kind)
    def key(c):
        return (c["slide_idx"], c["position"], c["kind"])

    actual_by_key = {key(c): c for c in actual["components"]}
    expected_by_key = {key(c): c for c in expected["components"]}

    for k, e_comp in expected_by_key.items():
        a_comp = actual_by_key.get(k)
        if a_comp is None:
            errors.append(f"component {k}: MISSING in actual run")
            continue
        for signal in ("categories_match", "series_names_match", "values_match"):
            if e_comp.get(signal) is True and a_comp.get(signal) is not True:
                errors.append(
                    f"component {k}: {signal} regressed from True → {a_comp.get(signal)}"
                )

    return errors
