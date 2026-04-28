"""For every chart in REFRESHED · WRONG, dump:
  - the chart's series values
  - the full API value pool (decimal, share, percentage/100, count, etc.)
  - which specific chart values are missing from the pool
  - a tag suggesting the likely cause:
      ALL_MISSING       — none of the chart's values appear in API
      FEW_MISSING       — ≤ 2 values missing (likely tolerance edge)
      AGG_MISMATCH      — chart values look aggregated (sum to 1, share-of-row)
      NO_NUMERIC        — chart has all None values

CSV output per deck:
  output/_forward/<deck>/refreshed_wrong_triage.csv
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from pptx import Presentation
from slidegen.intelligent_refresh import fetch_synapse_data
from tests.evals.end_to_end.compare_decks import _extract_chart_data, _find_match


DECKS = {
    "atu_q1_26": REPO / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
    "creon_pet_w33": REPO / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
}

VALUE_TOLERANCE = 1e-3
NUMERIC_FIELDS = (
    ("decimal", 1.0),
    ("share", 1.0),
    ("penetration", 1.0),
    ("percentage", 0.01),
    ("count", 1.0),
    ("sum", 1.0),
    ("penetration_num", 1.0),
    ("value", 1.0),
)


def _values_eq(a, b, tol=1e-3):
    if len(a) != len(b):
        return False
    for x, y in zip(a, b):
        if x is None and y is None:
            continue
        if x is None or y is None:
            return False
        if abs(x - y) > tol:
            return False
    return True


def _series_eq(a_series, b_series):
    if len(a_series) != len(b_series):
        return False
    for (an, av), (bn, bv) in zip(a_series, b_series):
        if an != bn:
            return False
        if not _values_eq(av, bv):
            return False
    return True


def triage_deck(deck_key: str) -> None:
    src_path = DECKS[deck_key]
    ref_path = REPO / "output" / "step2_test_connected" / f"{deck_key}.pptx"
    spec_path = REPO / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"

    spec = json.loads(spec_path.read_text(encoding="utf-8"))

    comp_ds_by_name: dict[tuple, str] = {}
    comp_ds_by_pos: dict[int, list] = {}
    for slide in spec.get("slides", []):
        s_idx = slide["slide_index"]
        slide_default = slide.get("data_source")
        for comp in slide.get("components", []):
            if comp.get("type") != "chart":
                continue
            ds = comp.get("data_source") or slide_default
            name = comp.get("name") or ""
            if name:
                comp_ds_by_name[(s_idx, name)] = ds
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_ds_by_pos.setdefault(s_idx, []).append((cl, ct, ds))

    static_ds = {
        ds_key for ds_key, d in spec.get("data_sources", {}).items()
        if d.get("static_time_period_ids") and not d.get("include_live_wave")
    }

    api_pool_cache: dict[str, list[float]] = {}

    def _api_pool(ds_key: str) -> list[float]:
        if ds_key in api_pool_cache:
            return api_pool_cache[ds_key]
        d = spec["data_sources"].get(ds_key, {})
        try:
            records, _ = fetch_synapse_data({
                "project_id": d.get("project_id"),
                "reporting_plan_id": d.get("reporting_plan_id"),
                "analysis_ids": list(d.get("analysis_ids") or []),
                "segment_ids": list(d.get("segment_ids") or []),
                "dynamic_latest_n": d.get("dynamic_latest_n") or 0,
                "static_time_period_ids": d.get("static_time_period_ids") or [],
                "static_time_period_names": [],
                "include_live_wave": bool(d.get("include_live_wave", False)),
            })
        except Exception:
            records = []
        pool: list[float] = []
        for r in records:
            for key, scale in NUMERIC_FIELDS:
                v = r.get(key)
                if v is None:
                    continue
                try:
                    pool.append(round(float(v) * scale, 4))
                except (TypeError, ValueError):
                    continue
        api_pool_cache[ds_key] = pool
        return pool

    def _resolve_ds(s_idx, name, pos):
        if name and (s_idx, name) in comp_ds_by_name:
            return comp_ds_by_name[(s_idx, name)]
        for cl, ct, ds in comp_ds_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return ds
        return None

    src = Presentation(str(src_path))
    ref = Presentation(str(ref_path))

    rows = []
    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        src_charts = [s for s in s_slide.shapes if s.has_chart]
        ref_charts = [s for s in r_slide.shapes if s.has_chart]
        for src_shape in src_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            ds_key = _resolve_ds(s_idx, shape_name, (sl, st))
            if ds_key is None or ds_key in static_ds:
                continue
            ref_shape = _find_match(src_shape, ref_charts)
            if ref_shape is None:
                continue
            try:
                _, src_series = _extract_chart_data(src_shape)
                _, ref_series = _extract_chart_data(ref_shape)
            except Exception:
                continue
            # Skip welded
            if _series_eq(src_series, ref_series):
                continue
            # Skip structural drift
            if (src_shape.chart.chart_type != ref_shape.chart.chart_type
                    or len(src_series) != len(ref_series)):
                continue

            ref_values = []
            for _n, vals in ref_series:
                for v in vals:
                    if v is not None:
                        ref_values.append(round(float(v), 4))

            if not ref_values:
                tag = "NO_NUMERIC"
                missing_count = 0
                missing_sample = ""
                api_pool_size = -1
            else:
                pool = _api_pool(ds_key)
                pool_remain = list(pool)
                missing = []
                for v in ref_values:
                    found = False
                    for i, p in enumerate(pool_remain):
                        if abs(v - p) <= VALUE_TOLERANCE:
                            pool_remain.pop(i)
                            found = True
                            break
                    if not found:
                        missing.append(v)
                if not missing:
                    continue   # Actually OK — skip this row
                api_pool_size = len(pool)
                missing_count = len(missing)
                # Heuristic: AGG_MISMATCH if all missing values look like
                # share-of-row aggregates (each series sums to ~1 across cats)
                series_sums = []
                for n, vs in ref_series:
                    nonn = [v for v in vs if v is not None]
                    if nonn:
                        series_sums.append(round(sum(nonn), 2))
                if (series_sums and
                        all(abs(s - 1.0) <= 0.05 or abs(s - 100) <= 5
                            for s in series_sums)):
                    tag = "AGG_MISMATCH"
                elif missing_count == len(ref_values):
                    tag = "ALL_MISSING"
                elif missing_count <= 2:
                    tag = "FEW_MISSING"
                else:
                    tag = "SOME_MISSING"
                missing_sample = ", ".join(str(m) for m in missing[:6])

            rows.append({
                "deck": deck_key,
                "slide_idx_1based": s_idx + 1,
                "chart_name": shape_name,
                "ds_key": ds_key,
                "tag": tag,
                "n_series": len(ref_series),
                "n_chart_values": len(ref_values),
                "n_missing_in_api": missing_count,
                "missing_sample": missing_sample,
                "api_pool_size": api_pool_size,
                "first_series_name": (ref_series[0][0] if ref_series else ""),
                "first_series_vals": ", ".join(str(round(v, 4)) if v is not None else "—"
                                                for v in (ref_series[0][1] if ref_series else [])[:6]),
            })

    out = REPO / "output" / "_forward" / deck_key / "refreshed_wrong_triage.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8", newline="") as f:
        if rows:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    # Summary
    from collections import Counter
    tag_counts = Counter(r["tag"] for r in rows)
    print(f"=== {deck_key}: {len(rows)} refreshed-wrong charts ===")
    for tag, n in tag_counts.most_common():
        print(f"  {tag}: {n}")
    print(f"  -> {out.relative_to(REPO)}")


if __name__ == "__main__":
    for k in DECKS:
        triage_deck(k)
