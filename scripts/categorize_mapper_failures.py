"""For every chart on a dynamic data source where forward refresh welded
the chart AND the API has new wave data available (i.e., WELDED · MAPPER
FAILED — case 2), categorize by selectedColumns pattern.

This identifies the dominant failure shapes so we know where to focus
mapper fixes. Modeled on Fix L pattern triage — pick the largest pattern
and write a targeted handler.

Patterns:
  P_STANDALONE_WAVE          — entry like 'Project Wave 13' (Fix L should handle)
  P_AT_AT_COMPOUND_WAVE      — 'X @:@ Project Wave 13' (Fix L should handle)
  P_DASH_COMPOUND_WAVE       — 'X - HIT - Project Wave 13' (Fix L extension should handle)
  P_W_SHORT_FORM             — 'W28' or 'W28 - Speciality' (Fix L should handle now)
  P_QUARTER_LABEL            — 'Q1'26', 'Jan'26' (Fix L should handle)
  P_ALIASED_TO_QUARTER       — selectedColumns has 'Project Wave 13' but
                                source chart shows 'Q1'26' (alias mismatch)
  P_NO_WAVE_LABEL_IN_SC      — selectedColumns has no wave-like entries
                                — failure must be from RowFields/ColumnFields
                                or PivotConfig mismatch
  P_OTHER                    — none of the above

Run: python scripts/categorize_mapper_failures.py
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from pptx import Presentation
from tests.evals.end_to_end.compare_decks import _extract_chart_data, _find_match
from slidegen.intelligent_refresh import fetch_synapse_data


WAVE_PATTERNS = (
    re.compile(r"^(?:Project )?Wave \d+$"),
    re.compile(r"^W\d+$"),
    re.compile(r"^[A-Z][a-z]{2}'\d{2}$"),
    re.compile(r"^Q[1-4]'\d{2}$"),
    re.compile(r"^Q[1-4] \d{4}$"),
)
EMBEDDED_HINT = re.compile(r"\b(?:W|Wave|Q|Project Wave)[\s_]*\d+", re.IGNORECASE)


def is_wave_label(s: str) -> bool:
    return any(p.match(s) for p in WAVE_PATTERNS)


def has_embedded_wave_hint(s: str) -> bool:
    if is_wave_label(s):
        return False
    return bool(EMBEDDED_HINT.search(s))


def categorize_sc(sc: list, src_labels_str: str) -> str:
    """Categorize a single chart's failure pattern based on selectedColumns
    AND the source chart's actual displayed labels (cats + series names).
    The displayed-labels signal lets us catch the "aliased to quarter" case
    where selectedColumns has 'Project Wave 13' but the chart shows 'Q1'26'.
    """
    sc_strs = [s for s in sc if isinstance(s, str)]
    has_standalone_wave = any(is_wave_label(s) for s in sc_strs)
    has_at_compound_wave = any(
        " @:@ " in s and any(is_wave_label(p) for p in s.split(" @:@ "))
        for s in sc_strs
    )
    has_dash_compound_wave = any(
        " - " in s and any(is_wave_label(p) for p in s.split(" - "))
        for s in sc_strs
    )
    has_w_short = any(
        "W" in s and any(re.match(r"^W\d+$", p) for p in re.split(r"\s+|[-,]", s))
        for s in sc_strs
    )
    has_embedded = any(has_embedded_wave_hint(s) for s in sc_strs)

    # Detect aliased-to-quarter: source labels contain Q-style or month-style
    # waves not literally present in selectedColumns
    src_has_q_label = bool(re.search(r"\b(?:Q[1-4]'\d{2}|[A-Z][a-z]{2}'\d{2})\b", src_labels_str))
    sc_concat = " | ".join(sc_strs)
    sc_has_q_label = bool(re.search(r"\b(?:Q[1-4]'\d{2}|[A-Z][a-z]{2}'\d{2})\b", sc_concat))

    if src_has_q_label and not sc_has_q_label and (has_standalone_wave or has_at_compound_wave or has_dash_compound_wave):
        return "P_ALIASED_TO_QUARTER"

    if has_at_compound_wave:
        return "P_AT_AT_COMPOUND_WAVE"
    if has_dash_compound_wave:
        return "P_DASH_COMPOUND_WAVE"
    if has_w_short:
        return "P_W_SHORT_FORM"
    if has_standalone_wave:
        return "P_STANDALONE_WAVE"
    if has_embedded:
        return "P_OTHER_EMBEDDED"
    return "P_NO_WAVE_LABEL_IN_SC"


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


def categorize_deck(deck_key: str) -> dict:
    src_path_map = {
        "atu_q1_26": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
        "creon_pet_w33": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
    }
    src_path = src_path_map[deck_key]
    ref_path = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}.pptx"
    spec_path = REPO_ROOT / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"

    spec = json.loads(spec_path.read_text(encoding="utf-8"))

    comp_info_by_name: dict[tuple, tuple] = {}
    comp_info_by_pos: dict[int, list] = {}
    for slide in spec.get("slides", []):
        s_idx = slide["slide_index"]
        slide_default = slide.get("data_source")
        for comp in slide.get("components", []):
            if comp.get("type") != "chart":
                continue
            ds = comp.get("data_source") or slide_default
            sc = comp.get("raw_mapping_config", {}).get("selectedColumns", [])
            rpc = comp.get("raw_pivot_config", {})
            name = comp.get("name") or ""
            if name:
                comp_info_by_name[(s_idx, name)] = (sc, ds, rpc)
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_info_by_pos.setdefault(s_idx, []).append((cl, ct, sc, ds, rpc))

    static_ds = {
        ds_key for ds_key, ds in spec.get("data_sources", {}).items()
        if ds.get("static_time_period_ids") and not ds.get("include_live_wave")
    }

    def _resolve(s_idx, name, pos):
        if name and (s_idx, name) in comp_info_by_name:
            return comp_info_by_name[(s_idx, name)]
        for cl, ct, sc, ds, rpc in comp_info_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return (sc, ds, rpc)
        return None

    api_wave_cache: dict[str, set] = {}

    def _api_waves(ds_key: str) -> set:
        if ds_key in api_wave_cache:
            return api_wave_cache[ds_key]
        ds = spec["data_sources"].get(ds_key, {})
        try:
            records, _ = fetch_synapse_data({
                "project_id": ds.get("project_id"),
                "reporting_plan_id": ds.get("reporting_plan_id"),
                "analysis_ids": list(ds.get("analysis_ids") or []),
                "segment_ids": list(ds.get("segment_ids") or []),
                "dynamic_latest_n": ds.get("dynamic_latest_n") or 0,
                "static_time_period_ids": ds.get("static_time_period_ids") or [],
                "static_time_period_names": [],
                "include_live_wave": bool(ds.get("include_live_wave", False)),
            })
        except Exception:
            records = []
        names: set = set()
        for r in records:
            n = r.get("time_period_name")
            if n is not None:
                names.add(str(n))
                stripped = str(n).replace("Project Wave ", "Wave ")
                if stripped != str(n):
                    names.add(stripped)
        api_wave_cache[ds_key] = names
        return names

    src = Presentation(str(src_path))
    ref = Presentation(str(ref_path))

    by_pattern: dict[str, list[dict]] = defaultdict(list)

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        s_charts = [s for s in s_slide.shapes if s.has_chart]
        r_charts = [s for s in r_slide.shapes if s.has_chart]
        for src_shape in s_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            info = _resolve(s_idx, shape_name, (sl, st))
            if info is None:
                continue
            sc, ds_key, rpc = info
            if ds_key in static_ds:
                continue  # not a mapper failure — static by design
            ref_shape = _find_match(src_shape, r_charts)
            if ref_shape is None:
                continue
            try:
                src_cats, src_series = _extract_chart_data(src_shape)
                _, ref_series = _extract_chart_data(ref_shape)
            except Exception:
                continue
            # We only care about welded charts (src == ref)
            if not _series_eq(src_series, ref_series):
                continue
            # Of those welded, only case 2 (API has new waves not in source).
            # Empty API → also case 1 (no records to refresh to — correct welding).
            api_waves = _api_waves(ds_key)
            src_labels = list(src_cats) + [n for n, _ in src_series]
            haystack = " | ".join(src_labels)
            if not api_waves:
                continue  # case 1 — API returned nothing, no data available
            if all(w in haystack for w in api_waves):
                continue  # case 1 — API waves already in source
            # Case 2: mapper failed
            cat = categorize_sc(sc, haystack)
            by_pattern[cat].append({
                "slide": s_idx,
                "name": shape_name,
                "ds": ds_key,
                "selectedColumns": sc[:5],
                "src_labels_sample": src_labels[:6],
                "api_waves": sorted(api_waves)[:5],
                "RowFields": rpc.get("RowFields"),
                "ColumnFields": rpc.get("ColumnFields"),
            })

    return {"deck": deck_key, "by_pattern": dict(by_pattern)}


def print_report(rpt: dict) -> None:
    print(f"\n=== {rpt['deck']} | mapper-failed pattern breakdown ===")
    by_pat = rpt["by_pattern"]
    total = sum(len(v) for v in by_pat.values())
    print(f"  total mapper-failed charts: {total}")
    # Sort by count desc
    for cat in sorted(by_pat, key=lambda k: -len(by_pat[k])):
        rows = by_pat[cat]
        print(f"\n  {cat}: {len(rows)} charts")
        for r in rows[:3]:
            print(f"    slide {r['slide']:>2} {r['name']!r}  ds={r['ds']}")
            print(f"      sc: {r['selectedColumns']}")
            print(f"      src labels: {r['src_labels_sample']}")
            print(f"      api waves: {r['api_waves']}")
            print(f"      Row/Col: {r['RowFields']} / {r['ColumnFields']}")
        if len(rows) > 3:
            print(f"    ... +{len(rows)-3} more")


if __name__ == "__main__":
    for k in ("atu_q1_26", "creon_pet_w33"):
        rpt = categorize_deck(k)
        print_report(rpt)
        out = REPO_ROOT / "output" / "_forward" / k / "mapper_failed_patterns.json"
        out.parent.mkdir(parents=True, exist_ok=True)
        # Compact full report
        out.write_text(
            json.dumps(rpt, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        print(f"\n  full report -> {out}")
