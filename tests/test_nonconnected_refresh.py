"""
test_nonconnected_refresh.py — End-to-end test of non-connected slide refresh.

Tests the inference-based flow:
  1. Extract series names + chart pattern from OOXML (no Connector tags)
  2. User provides project_id + analysis_id
  3. Fetch Synapse records
  4. infer_data_transform() matches series names against data columns
  5. Generate complete spec with DataTransform
  6. Refresh chart using same pipeline as connected slides

Usage:
    python -m slidegen.synapse_auth --update "Bearer eyJ..."
    python tests/test_nonconnected_refresh.py
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData

REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_DIR = REPO_ROOT.parent / "Sample Decks"
SOURCE_PPTX = SAMPLE_DIR / "Repatha ATU Slide 6.pptx"
OUTPUT_PPTX = SAMPLE_DIR / "Repatha_ATU_Slide6_nonconnected_refresh.pptx"

# User-provided data lineage (they know the Synapse project even without tags)
DATA_LINEAGE = {
    "project_id": 1428,
    "project_name": "Amgen [ATU]: Repatha",
    "reporting_plan_id": 574,
    "analysis_ids": [545991],
    "segment_ids": [14633],
    "dynamic_latest_n": 2,
}


def main():
    sys.path.insert(0, str(REPO_ROOT))
    from dotenv import load_dotenv
    load_dotenv(REPO_ROOT / ".env")
    import os
    import requests
    import pandas as pd
    from slidegen.slide_spec.data_inference import infer_data_transform
    from slidegen.synapse_chart_mapper import (
        pivot_records_to_chart_data, ChartRefreshData,
    )

    print("=" * 60)
    print("Non-Connected Slide Refresh Test")
    print("=" * 60)

    assert SOURCE_PPTX.exists(), f"Source not found: {SOURCE_PPTX}"

    # ── Step 1: Extract chart info from OOXML (no tags needed) ──
    prs = Presentation(str(SOURCE_PPTX))
    slide = prs.slides[0]  # Slide 0 = non-connected version

    print("\nStep 1: Extract chart info from OOXML")
    charts_info = []
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        chart = shape.chart
        p = chart.plots[0]
        cats = [str(c) for c in p.categories] if p.categories else []
        series_names = [str(s.name) for s in p.series]
        series_colors = []
        for s in p.series:
            try:
                rgb = s.format.fill.fore_color.rgb
                series_colors.append(f"#{rgb}")
            except Exception:
                series_colors.append("#999999")

        from pptx.enum.chart import XL_CHART_TYPE
        chart_type_map = {
            XL_CHART_TYPE.BAR_STACKED_100: "bar_stacked_100_horizontal",
            XL_CHART_TYPE.BAR_CLUSTERED: "bar_clustered_horizontal",
            XL_CHART_TYPE.LINE_MARKERS: "line_markers_trended",
        }
        chart_pattern = chart_type_map.get(chart.chart_type, "bar_stacked_100_horizontal")

        pos = (round(shape.left / 914400, 2), round(shape.top / 914400, 2))
        info = {
            "position": pos,
            "chart_pattern": chart_pattern,
            "categories": cats,
            "series_names": series_names,
            "series_colors": series_colors,
            "category_count": len(cats),
        }
        charts_info.append(info)
        print(f"  Chart @{pos}: {chart_pattern}")
        print(f"    series: {series_names}")
        print(f"    colors: {series_colors}")
        print(f"    cats[{len(cats)}]: {cats[:4]}")

    # ── Step 2: Fetch Synapse data ──
    print("\nStep 2: Fetch Synapse data")
    token = os.environ.get("SYNAPSE_API_TOKEN", "")
    base_url = os.environ.get("SYNAPSE_API_BASE_URL", "https://synapse-api.zoomrx.com")

    if not token:
        print("  ERROR: No token. Run: python -m slidegen.synapse_auth --update 'Bearer eyJ...'")
        return

    # Check token validity
    try:
        import base64, time
        raw = token.replace("Bearer ", "").strip()
        payload_b64 = raw.split(".")[1] + "=="
        claims = json.loads(base64.urlsafe_b64decode(payload_b64))
        remaining = (claims["exp"] - time.time()) / 60
        if remaining < 1:
            print(f"  ERROR: Token expired {abs(remaining):.0f} min ago")
            return
        print(f"  Token valid for {remaining:.0f} min")
    except Exception:
        pass

    headers = {
        "Authorization": token if token.startswith("Bearer") else f"Bearer {token}",
        "accept": "application/json",
        "Content-Type": "application/json",
    }
    payload = {
        "project_id": DATA_LINEAGE["project_id"],
        "reporting_plan_id": DATA_LINEAGE["reporting_plan_id"],
        "analysis_ids": DATA_LINEAGE["analysis_ids"],
        "segment_ids": DATA_LINEAGE["segment_ids"],
        "setup_type": "DYNAMIC",
        "dynamic_time_period": {
            "latest_n_deliverables": DATA_LINEAGE["dynamic_latest_n"],
            "include_live_wave": True,
        },
    }
    resp = requests.post(f"{base_url}/api/reports/generate",
                         headers=headers, json=payload, timeout=60)
    if resp.status_code not in (200, 201):
        print(f"  API error {resp.status_code}: {resp.text[:200]}")
        return
    records = resp.json().get("records", [])
    print(f"  Fetched {len(records)} records")

    df = pd.DataFrame(records)
    print(f"  Columns: {sorted(df.columns.tolist())[:10]}")

    # ── Step 3: Infer DataTransform from series names ──
    print("\nStep 3: Infer DataTransform")

    # Build column values map for inference
    record_columns = {}
    for col in df.columns:
        try:
            unique = df[col].dropna().unique()
            # Only include non-numeric columns for series matching
            if df[col].dtype == object or len(unique) < 50:
                record_columns[col] = [str(v) for v in unique]
        except Exception:
            pass

    print(f"  Available columns: {list(record_columns.keys())[:8]}")

    for ci, info in enumerate(charts_info):
        print(f"\n  Chart {ci} @{info['position']}:")
        print(f"    Series names: {info['series_names']}")

        try:
            transform, series_config = infer_data_transform(
                series_names=info["series_names"],
                chart_pattern=info["chart_pattern"],
                record_columns=record_columns,
                category_count=info["category_count"],
            )
            print(f"    Inferred transform:")
            print(f"      row_field: {transform.row_field}")
            print(f"      column_field: {transform.column_field}")
            print(f"      value_field: {transform.value_field}")
            print(f"      filters: {[(f.field, f.value) for f in transform.filters]}")
            print(f"    Series config: {[(s.role, s.color) for s in series_config]}")

            # Store for refresh
            info["transform"] = transform
            info["series_config"] = series_config
            # Store the name-to-value mapping for series role matching
            # (e.g. "IDK" -> "0" in the data)
            info["name_to_value"] = {}
            # Get it from the inference results
            from slidegen.slide_spec.data_inference import match_series_to_column
            for col_name, col_vals in record_columns.items():
                cm = match_series_to_column(info["series_names"], col_name, col_vals)
                if cm and cm.column_name == transform.column_field:
                    info["name_to_value"] = cm.name_to_value
                    break
        except Exception as e:
            print(f"    INFERENCE FAILED: {e}")
            info["transform"] = None

    # ── Step 4: Refresh using inferred transforms ──
    print("\nStep 4: Refresh charts")
    shutil.copy2(str(SOURCE_PPTX), str(OUTPUT_PPTX))
    out_prs = Presentation(str(OUTPUT_PPTX))
    out_slide = out_prs.slides[0]

    out_charts = sorted(
        [s for s in out_slide.shapes if s.has_chart],
        key=lambda x: (x.left or 0, x.top or 0),
    )

    for ci, (shape, info) in enumerate(zip(out_charts, charts_info)):
        if not info.get("transform"):
            print(f"  Chart {ci}: skipped (no transform)")
            continue

        transform = info["transform"]
        series_config = info["series_config"]

        # Use the non-connected path: DataTransform → pivot via pandas
        df_pivot = pd.DataFrame(records)
        # Apply filters
        for filt in transform.filters:
            if filt.field in df_pivot.columns:
                if filt.operator == "eq":
                    df_pivot = df_pivot[df_pivot[filt.field] == filt.value]
                elif filt.operator == "ne":
                    df_pivot = df_pivot[df_pivot[filt.field] != filt.value]
        # Pivot
        if transform.column_field and transform.column_field in df_pivot.columns:
            pivot = df_pivot.pivot_table(
                index=transform.row_field,
                columns=transform.column_field,
                values=transform.value_field,
                aggfunc="first",
            )
            cats_list = list(pivot.index.astype(str))
            series_list = []
            for sc in series_config:
                # Match role to pivot column
                matched = None
                for pc in pivot.columns:
                    if str(pc) == sc.role or sc.role.lower() == str(pc).lower():
                        matched = pc
                        break
                # Try name_to_value mapping
                if matched is None:
                    for k, v in info.get("name_to_value", {}).items():
                        if k == sc.role:
                            for pc in pivot.columns:
                                if str(pc) == v:
                                    matched = pc; break
                if matched is not None:
                    vals = [0.0 if pd.isna(v) else float(v) for v in pivot[matched]]
                    series_list.append((sc.role, vals))
            chart_data = type('D', (), {'success': True, 'categories': cats_list,
                                        'series': series_list, 'error': ''})()

        if not chart_data.success or not chart_data.categories:
            print(f"  Chart {ci}: transform failed: {chart_data.error}")
            continue

        # Read source formatCode before replace_data
        ns_c = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        src_fmt = None
        for fc_el in shape.chart._chartSpace.iter(f'{{{ns_c}}}formatCode'):
            src_fmt = fc_el.text
            break

        # Write chart data
        try:
            cd = CategoryChartData()
            cd.categories = chart_data.categories
            for name, vals in chart_data.series:
                cd.add_series(name, vals)
            shape.chart.replace_data(cd)

            # Restore formatCode
            if src_fmt:
                for fc_el in shape.chart._chartSpace.iter(f'{{{ns_c}}}formatCode'):
                    fc_el.text = src_fmt

            print(f"  Chart {ci}: OK - {len(chart_data.categories)} cats, {len(chart_data.series)} series")
            print(f"    cats: {chart_data.categories[:4]}")
            print(f"    series: {[s[0] for s in chart_data.series]}")
            print(f"    vals[0]: {[round(v,3) for v in chart_data.series[0][1]][:4]}")
        except Exception as e:
            print(f"  Chart {ci}: FAILED: {e}")

    out_prs.save(str(OUTPUT_PPTX))

    # ── Step 5: Compare ──
    print("\nStep 5: Compare source vs refreshed")
    src = Presentation(str(SOURCE_PPTX))
    ref = Presentation(str(OUTPUT_PPTX))

    src_charts = sorted([s for s in src.slides[0].shapes if s.has_chart],
                        key=lambda x: (x.left or 0, x.top or 0))
    ref_charts = sorted([s for s in ref.slides[0].shapes if s.has_chart],
                        key=lambda x: (x.left or 0, x.top or 0))

    for ci in range(min(len(src_charts), len(ref_charts))):
        sp = src_charts[ci].chart.plots[0]
        rp = ref_charts[ci].chart.plots[0]
        s_cats = [str(c) for c in sp.categories]
        r_cats = [str(c) for c in rp.categories]
        s_ns = len(sp.series)
        r_ns = len(rp.series)
        s_v = [round(float(v), 3) if v else 0 for v in sp.series[0].values][:4]
        r_v = [round(float(v), 3) if v else 0 for v in rp.series[0].values][:4]

        cats_match = s_cats == r_cats
        series_match = s_ns == r_ns
        print(f"  Chart {ci}: cats={'OK' if cats_match else 'DIFF'} series={'OK' if series_match else 'DIFF'}")
        if not cats_match:
            print(f"    src: {s_cats[:4]}")
            print(f"    ref: {r_cats[:4]}")
        print(f"    src vals: {s_v}")
        print(f"    ref vals: {r_v}")

    print(f"\nOutput: {OUTPUT_PPTX}")


if __name__ == "__main__":
    main()
