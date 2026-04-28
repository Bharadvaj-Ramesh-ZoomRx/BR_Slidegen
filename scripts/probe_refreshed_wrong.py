"""For each chart classified REFRESHED · WRONG, look at:
  - chart's actual cats + series values
  - what the API returned for the chart's ds (raw records)
  - whether each chart value can be found in the API records as decimal,
    percentage/100, or as a derived share

Identifies whether the truth check is at fault or refresh is producing
wrong values.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import json
from pptx import Presentation
from slidegen.intelligent_refresh import fetch_synapse_data
from tests.evals.end_to_end.compare_decks import _extract_chart_data, _find_match

REPO = Path(__file__).resolve().parents[1]

# Pick one ATU chart and one CREON chart from the wrong list
TARGETS = [
    ("atu_q1_26", "ZoomRx_UC_ATU_Report_Q1_'26.pptx"),
    ("creon_pet_w33", "CREON Share of Voice Study - W33.pptx"),
]

for deck_key, src_name in TARGETS:
    src_path = REPO / "projects" / "J&J Rybrevant PET" / "Template" / src_name
    ref_path = REPO / "output" / "step2_test_connected" / f"{deck_key}.pptx"
    spec_path = REPO / "output" / "step2_test_connected" / f"{deck_key}_full_spec.json"
    summary_path = REPO / "output" / "_forward" / deck_key / "summary.json"

    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    wrong = summary.get("refreshed_wrong_detail", [])[:2]  # first 2

    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    src = Presentation(str(src_path))
    ref = Presentation(str(ref_path))

    # Index spec components for ds lookup
    comp_ds_by_name: dict[tuple, str] = {}
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

    for w in wrong:
        s_idx = w["slide"]
        name = w["name"]
        ds_key = w.get("ds") or comp_ds_by_name.get((s_idx, name))
        if ds_key is None:
            continue
        ds = spec["data_sources"][ds_key]

        # Find ref shape on slide s_idx with name == name
        ref_charts = [s for s in ref.slides[s_idx].shapes if s.has_chart]
        ref_shape = next((s for s in ref_charts if s.name == name), None)
        if ref_shape is None:
            continue
        cats, series = _extract_chart_data(ref_shape)

        print(f"\n========== {deck_key} | slide {s_idx} {name!r} ==========")
        print(f"  ds: {ds_key}")
        print(f"  ds config: dyn_n={ds.get('dynamic_latest_n')} static={ds.get('static_time_period_ids')} "
              f"live={ds.get('include_live_wave')} segs={ds.get('segment_ids')}")
        print(f"  chart cats: {cats[:6]}")
        print(f"  chart series: {[(n, vals[:4]) for n, vals in series[:4]]}")
        print(f"  test reason: {w.get('reason', '')}")

        # Fetch what API returned
        records, df = fetch_synapse_data(ds)
        print(f"  API records: {len(records)}")
        if not df.empty:
            print(f"    columns: {list(df.columns)[:12]}")
            print(f"    distinct waves: {df['time_period_name'].unique().tolist() if 'time_period_name' in df.columns else 'n/a'}")
            # Distinct decimal values
            if "decimal" in df.columns:
                dec_vals = sorted({round(float(v), 4) for v in df["decimal"].dropna().tolist()})[:12]
                print(f"    distinct decimals (first 12): {dec_vals}")
            if "count" in df.columns:
                cnt_vals = sorted({int(v) for v in df["count"].dropna().tolist()})[:12]
                print(f"    distinct counts (first 12): {cnt_vals}")
        # Refreshed values that supposedly don't match
        ref_vals = []
        for _n, vals in series:
            for v in vals:
                if v is not None:
                    ref_vals.append(round(float(v), 4))
        print(f"    refreshed values (first 12): {sorted(set(ref_vals))[:12]}")
