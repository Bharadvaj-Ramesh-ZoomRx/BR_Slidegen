"""
test_spec_refresh_pipeline.py — End-to-end test of spec-driven deck refresh.

Three stages, three PPTX files:
  Stage 1: Source -> Clone with dummy data + dummy headlines
  Stage 2: Clone + config specs + Synapse API -> Refreshed deck with real data
  Stage 3: Verify source ~ refreshed (same structure, data should match)

Usage:
    # Stage 1 (no API needed):
    python tests/test_spec_refresh_pipeline.py --stage 1

    # Stage 2 (needs valid Synapse token in .env):
    python tests/test_spec_refresh_pipeline.py --stage 2

    # Stage 3 (verification):
    python tests/test_spec_refresh_pipeline.py --stage 3

    # All stages:
    python tests/test_spec_refresh_pipeline.py --all

    # Update token first, then run all:
    python -m slidegen.synapse_auth --update "Bearer eyJ..."
    python tests/test_spec_refresh_pipeline.py --all
"""
from __future__ import annotations

import json
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData


# ── Paths ────────────────────────────────────────────────────────────────────

REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_DIR = REPO_ROOT.parent / "Sample Decks"

SOURCE_PPTX = SAMPLE_DIR / "[Vijay] Synapse Connector UAT - Mar 2026.pptx"
DUMMY_PPTX = SAMPLE_DIR / "UAT_dummy_data_v6.pptx"
REFRESHED_PPTX = SAMPLE_DIR / "UAT_refreshed_from_spec_v6.pptx"
SPECS_JSON = SAMPLE_DIR / "UAT_deck_config_specs_v1.2.json"


# ══════════════════════════════════════════════════════════════════════════════
# Stage 1: Clone source -> dummy deck (zero data + dummy headlines)
# ══════════════════════════════════════════════════════════════════════════════

def stage1_create_dummy_deck():
    """Clone source PPTX, replace all chart data with dummy values and
    all headlines with placeholder text."""
    print("\n" + "=" * 70)
    print("STAGE 1: Create dummy deck from source")
    print("=" * 70)

    assert SOURCE_PPTX.exists(), f"Source not found: {SOURCE_PPTX}"

    # Clone
    shutil.copy2(str(SOURCE_PPTX), str(DUMMY_PPTX))
    prs = Presentation(str(DUMMY_PPTX))

    charts_dummied = 0
    tables_dummied = 0
    headlines_dummied = 0

    for slide_idx, slide in enumerate(prs.slides):
        # ── Dummy out chart data ──
        for shape in slide.shapes:
            if shape.has_chart:
                chart = shape.chart
                try:
                    plot = chart.plots[0]
                    cats = list(plot.categories) if plot.categories else ["Cat1"]
                    n_cats = len(cats)
                    n_series = len(plot.series)

                    cd = CategoryChartData()
                    cd.categories = [f"Dummy {i+1}" for i in range(n_cats)]
                    for s_idx in range(n_series):
                        name = f"Dummy Series {s_idx+1}"
                        # Uniform dummy values (0.25 each for stacked, 0.5 for others)
                        val = round(1.0 / max(n_series, 1), 2)
                        cd.add_series(name, [val] * n_cats)
                    chart.replace_data(cd)
                    charts_dummied += 1
                except Exception as e:
                    print(f"  Warning: chart on slide {slide_idx} failed: {e}")

            # Tables: leave as-is (restored from source in Stage 2)
            elif shape.has_table:
                tables_dummied += 1

        # ── Dummy out headlines (largest text in top 1.5") ──
        candidates = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            top_in = (shape.top or 0) / 914400
            text = shape.text_frame.text.strip()
            if top_in < 1.5 and len(text) > 20:
                candidates.append((len(text), shape))
        if candidates:
            candidates.sort(key=lambda x: -x[0])
            headline_shape = candidates[0][1]
            try:
                for para in headline_shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = ""
                headline_shape.text_frame.paragraphs[0].runs[0].text = (
                    f"DUMMY HEADLINE — Slide {slide_idx + 1} — To Be Updated"
                )
                headlines_dummied += 1
            except Exception:
                pass

    prs.save(str(DUMMY_PPTX))
    print(f"\n  Source:  {SOURCE_PPTX.name}")
    print(f"  Dummy:   {DUMMY_PPTX.name}")
    print(f"  Charts dummied:    {charts_dummied}")
    print(f"  Tables dummied:    {tables_dummied}")
    print(f"  Headlines dummied: {headlines_dummied}")
    print(f"  Output: {DUMMY_PPTX} ({DUMMY_PPTX.stat().st_size:,} bytes)")
    return True


# ══════════════════════════════════════════════════════════════════════════════
# Data transform bridge: spec DataTransform + Synapse records -> chart/table data
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class RefreshChartData:
    """Chart data ready for replace_data()."""
    categories: list[str] = field(default_factory=list)
    series: list[tuple[str, list[float]]] = field(default_factory=list)
    success: bool = True
    error: str = ""


@dataclass
class RefreshTableData:
    """Table data ready for cell updates."""
    rows: list[list[str]] = field(default_factory=list)
    success: bool = True
    error: str = ""


def apply_data_transform(
    records: list[dict],
    transform,  # DataTransform
    series_config: list = None,  # list[SeriesConfig]
) -> RefreshChartData:
    """Apply a spec DataTransform to flat records -> chart-ready data.

    Source-agnostic: works with any tabular records (Synapse, Excel, etc.)
    """
    if not records:
        return RefreshChartData(success=False, error="No records")

    df = pd.DataFrame(records)

    # Apply filters
    for filt in transform.filters:
        col = filt.field
        if col not in df.columns:
            continue
        if filt.operator == "eq":
            df = df[df[col] == filt.value]
        elif filt.operator == "ne":
            df = df[df[col] != filt.value]
        elif filt.operator == "in":
            df = df[df[col].isin(filt.value.split(","))]
        elif filt.operator == "contains":
            df = df[df[col].str.contains(filt.value, case=False, na=False)]

    if df.empty:
        return RefreshChartData(success=False, error="All records filtered out")

    row_field = transform.row_field
    col_field = transform.column_field
    val_field = transform.value_field

    # Handle multi-field rows (e.g. "y_label+y_code")
    if "+" in row_field:
        parts = row_field.split("+")
        valid = [p for p in parts if p in df.columns]
        if valid:
            row_field = valid[0]

    if row_field not in df.columns:
        return RefreshChartData(success=False, error=f"row_field '{row_field}' not in data")
    if val_field not in df.columns:
        # Try common alternatives
        alt_map = {"percentage": "decimal", "base": "n"}
        alt = alt_map.get(val_field)
        if alt and alt in df.columns:
            val_field = alt
        else:
            return RefreshChartData(success=False, error=f"value_field '{val_field}' not in data")

    # Pivot if column_field specified
    if col_field and col_field in df.columns:
        try:
            pivot = df.pivot_table(
                index=row_field,
                columns=col_field,
                values=val_field,
                aggfunc="first",
            )
        except Exception as e:
            return RefreshChartData(success=False, error=f"Pivot failed: {e}")

        categories = list(pivot.index.astype(str))

        # Map pivot columns to series using SeriesConfig roles
        series = []
        if series_config:
            for sc in series_config:
                matched_col = None
                for pc_col in pivot.columns:
                    pc_str = str(pc_col)
                    if sc.role == pc_str or sc.role in pc_str or pc_str.startswith(sc.role):
                        matched_col = pc_col
                        break
                if matched_col is not None:
                    vals = [0.0 if pd.isna(v) else float(v) for v in pivot[matched_col]]
                    series.append((sc.role, vals))
                else:
                    # No match — fill with zeros
                    series.append((sc.role, [0.0] * len(categories)))
        else:
            for col in pivot.columns:
                vals = [0.0 if pd.isna(v) else float(v) for v in pivot[col]]
                series.append((str(col), vals))

        return RefreshChartData(categories=categories, series=series)
    else:
        # No pivot — group by row_field, aggregate val_field
        grouped = df.groupby(row_field, sort=False)[val_field].first()
        categories = list(grouped.index.astype(str))
        values = [0.0 if pd.isna(v) else float(v) for v in grouped.values]
        return RefreshChartData(categories=categories, series=[("value", values)])


def apply_table_transform(
    records: list[dict],
    table_mapping,  # TableDataMapping
) -> RefreshTableData:
    """Apply a spec TableDataMapping to flat records -> table rows."""
    if not records:
        return RefreshTableData(success=False, error="No records")

    df = pd.DataFrame(records)

    # Apply filters
    for filt in table_mapping.filters:
        col = filt.field
        if col not in df.columns:
            continue
        if filt.operator == "eq":
            df = df[df[col] == filt.value]
        elif filt.operator == "ne":
            df = df[df[col] != filt.value]

    if df.empty:
        return RefreshTableData(success=False, error="All records filtered out")

    # Build rows from column configs
    # First, deduplicate by the first column's source field
    cols = table_mapping.columns
    if not cols:
        return RefreshTableData(success=False, error="No column configs")

    # Header row from templates
    header = [c.header_template for c in cols]

    # Data rows
    rows = [header]

    # Get unique row keys from first data column
    first_col = cols[0]
    if first_col.source_field == "computed":
        # Can't build data rows from computed column alone
        return RefreshTableData(rows=rows, success=True)

    if first_col.source_field not in df.columns:
        return RefreshTableData(success=False,
                                error=f"source_field '{first_col.source_field}' not in data")

    unique_keys = df[first_col.source_field].unique()
    for key in unique_keys:
        row_data = df[df[first_col.source_field] == key].iloc[0]
        row = []
        for col_cfg in cols:
            if col_cfg.source_field == "computed":
                row.append("")
                continue
            # Apply per-column filter if present
            if col_cfg.filter:
                filtered = df[(df[first_col.source_field] == key) &
                              (df[col_cfg.filter.field] == col_cfg.filter.value)]
                if not filtered.empty:
                    val = filtered.iloc[0].get(col_cfg.source_field, "")
                else:
                    val = ""
            else:
                val = row_data.get(col_cfg.source_field, "")

            # Format
            if col_cfg.format and val != "":
                try:
                    if col_cfg.format == "0%":
                        val = f"{float(val):.0%}"
                    elif "{}" in col_cfg.format:
                        val = col_cfg.format.format(val)
                except (ValueError, TypeError):
                    val = str(val)
            else:
                val = str(val)
            row.append(val)
        rows.append(row)

    return RefreshTableData(rows=rows)


# ══════════════════════════════════════════════════════════════════════════════
# Synapse API fetch (builds payload from spec DataLineage)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_synapse_records(data_lineage, token: str, base_url: str) -> list[dict]:
    """Fetch flat records from Synapse API using DataLineage fields."""
    import requests

    headers = {
        "Authorization": f"Bearer {token}" if not token.startswith("Bearer") else token,
        "accept": "application/json",
        "Content-Type": "application/json",
    }

    # Build segment_ids from structured segments (v1.2) or legacy flat list
    segment_ids = []
    if data_lineage.segments:
        segment_ids = [s.rule_id for s in data_lineage.segments]
    elif data_lineage.segment_ids:
        segment_ids = data_lineage.segment_ids

    payload = {
        "project_id": data_lineage.project_id,
        "reporting_plan_id": data_lineage.reporting_plan_id,
        "analysis_ids": data_lineage.analysis_ids,
        "segment_ids": segment_ids,
        "setup_type": "DYNAMIC",
        "dynamic_time_period": {
            "latest_n_deliverables": data_lineage.dynamic_latest_n or 8,
            "include_live_wave": data_lineage.include_live_wave if data_lineage.include_live_wave is not None else True,
        },
    }

    try:
        resp = requests.post(
            f"{base_url}/api/reports/generate",
            headers=headers, json=payload, timeout=60,
        )
        if resp.status_code in (200, 201):
            return resp.json().get("records", [])
        else:
            print(f"  API error {resp.status_code}: {resp.text[:200]}")
            return []
    except Exception as e:
        print(f"  API exception: {e}")
        return []


# ══════════════════════════════════════════════════════════════════════════════
# Table refresh from Connector configs
# ══════════════════════════════════════════════════════════════════════════════

def _refresh_table_from_connector(
    records: list[dict],
    pivot_config: dict,
    mapping_config: dict,
    column_key_label_map: dict | None,
    static_time_period_names: list[str] | None,
    table,
) -> RefreshTableData:
    """Refresh a table using raw Connector configs.

    Tables work differently from charts:
    - selectedColumns defines the column structure
    - <blank:X> columns have header X but data comes from pivoted values
    - Regular columns show raw field values (product names, codes, etc.)
    - The table retains its original row/column count
    """
    from slidegen.synapse_chart_mapper import (
        pivot_records_to_chart_data, _remap_field, _normalize_key, _match_pivot_col,
    )

    selected = mapping_config.get("selectedColumns", [])
    if not selected:
        return RefreshTableData(success=False, error="No selectedColumns")

    # Use pivot_records_to_chart_data to get the pivoted data
    # (it handles all the filtering, sorting, compound keys)
    chart_data = pivot_records_to_chart_data(
        records, pivot_config, mapping_config,
        static_time_period_names=static_time_period_names,
    )

    n_table_rows = len(table.rows)
    n_table_cols = len(table.columns)

    # Build table content from chart_data + selectedColumns
    # selectedColumns for tables typically looks like:
    #   ['y_label'] — single column of category labels
    #   ['y_label', '<blank:Base>'] — labels + computed column
    #   ['<blank:Overall>', '<blank:CARDs>', '<blank:PCPs>'] — all computed

    all_blank = all(sc.startswith("<blank:") for sc in selected)

    if all_blank:
        # All computed columns — headers from blank labels
        # Data comes from the pivoted values (each blank col maps to a pivot column)
        headers = [sc.replace("<blank:", "").rstrip(">") for sc in selected]
        rows = [headers]
        # Fill data rows from chart_data
        if chart_data.success and chart_data.categories:
            for ci, cat in enumerate(chart_data.categories):
                row = []
                for si, _ in enumerate(selected):
                    if si < len(chart_data.series):
                        vals = chart_data.series[si][1]
                        v = vals[ci] if ci < len(vals) else 0.0
                        row.append(str(int(v)) if v > 2 else f"{v:.0%}" if v != 0 else "")
                    else:
                        row.append("")
                rows.append(row)
        return RefreshTableData(rows=rows)

    # Mixed: first column(s) are data fields, some may be <blank:>
    if chart_data.success and chart_data.categories:
        # Header row from selectedColumns
        headers = []
        for sc in selected:
            if sc.startswith("<blank:"):
                headers.append(sc.replace("<blank:", "").rstrip(">"))
            else:
                # Use KLM display name if available
                display = sc
                if column_key_label_map and sc in column_key_label_map:
                    display = column_key_label_map[sc]
                headers.append(display)

        rows = [headers]

        # Data rows: first non-blank col = categories, rest from pivot
        for ci, cat in enumerate(chart_data.categories):
            row = []
            series_idx = 0
            for si, sc in enumerate(selected):
                if sc.startswith("<blank:"):
                    # Computed column — try to get from series data
                    if series_idx < len(chart_data.series):
                        vals = chart_data.series[series_idx][1]
                        v = vals[ci] if ci < len(vals) else 0.0
                        if abs(v) > 2:
                            row.append(str(int(v)))
                        elif v != 0:
                            row.append(f"{v:.0%}")
                        else:
                            row.append("")
                        series_idx += 1
                    else:
                        row.append("")
                elif si == 0:
                    # First column = category label
                    row.append(str(cat))
                else:
                    # Data column from series
                    if series_idx < len(chart_data.series):
                        vals = chart_data.series[series_idx][1]
                        v = vals[ci] if ci < len(vals) else 0.0
                        if abs(v) > 2:
                            row.append(str(int(v)))
                        elif v != 0:
                            row.append(f"{v:.0%}")
                        else:
                            row.append("")
                        series_idx += 1
                    else:
                        row.append("")
            rows.append(row)

        return RefreshTableData(rows=rows)

    return RefreshTableData(success=False, error="No chart data for table")


# ══════════════════════════════════════════════════════════════════════════════
# Stage 2: Refresh dummy deck with real Synapse data using config specs
# ══════════════════════════════════════════════════════════════════════════════

def stage2_refresh_from_specs():
    """Read config specs, fetch Synapse data, refresh the dummy deck."""
    print("\n" + "=" * 70)
    print("STAGE 2: Refresh dummy deck with real data from Synapse")
    print("=" * 70)

    assert DUMMY_PPTX.exists(), f"Dummy deck not found: {DUMMY_PPTX}. Run stage 1 first."

    # Load config specs
    if SPECS_JSON.exists():
        specs_data = json.loads(SPECS_JSON.read_text(encoding="utf-8"))
    else:
        print("  Config specs not found. Generating from source...")
        sys.path.insert(0, str(REPO_ROOT))
        from slidegen.deck_reader.tag_reader import generate_config_specs
        from slidegen.slide_spec.schema import dump_spec
        specs_list, _summary = generate_config_specs(str(SOURCE_PPTX))
        specs_data = [json.loads(dump_spec(s)) for s in specs_list]
        SPECS_JSON.write_text(json.dumps(specs_data, indent=2, ensure_ascii=False),
                              encoding="utf-8")
        print(f"  Saved specs: {SPECS_JSON}")

    # Load specs into dataclasses
    sys.path.insert(0, str(REPO_ROOT))
    from slidegen.slide_spec.schema import load_spec
    specs = [load_spec(d) for d in specs_data]
    print(f"  Loaded {len(specs)} config specs")

    # Get Synapse token
    from dotenv import load_dotenv
    import os
    load_dotenv(REPO_ROOT / ".env")

    token = os.environ.get("SYNAPSE_API_TOKEN", "")
    if not token:
        print("  ERROR: No SYNAPSE_API_TOKEN in .env")
        print("  Run: python -m slidegen.synapse_auth --update 'Bearer eyJ...'")
        return False

    base_url = os.environ.get("SYNAPSE_API_BASE_URL", "https://synapse-api.zoomrx.com")

    # Check token validity — skip expiry check for API keys (sk_*)
    raw = token.replace("Bearer ", "").strip()
    if raw.startswith("sk_"):
        print(f"  Using API key (no expiry)")
    else:
        try:
            import base64, time
            payload_b64 = raw.split(".")[1]
            payload_b64 += "=" * (4 - len(payload_b64) % 4)
            claims = json.loads(base64.urlsafe_b64decode(payload_b64))
            exp = claims.get("exp", 0)
            remaining = exp - time.time()
            if remaining < 60:
                print(f"  ERROR: Token expired {abs(remaining)/60:.0f} minutes ago")
                print("  Run: python -m slidegen.synapse_auth --update 'Bearer eyJ...'")
                return False
            print(f"  Token valid for {remaining/60:.0f} more minutes")
        except Exception:
        print("  Warning: could not decode token expiry, proceeding anyway")

    # Clone dummy -> refreshed
    shutil.copy2(str(DUMMY_PPTX), str(REFRESHED_PPTX))
    prs = Presentation(str(REFRESHED_PPTX))
    print(f"  Cloned dummy ->{REFRESHED_PPTX.name}")

    # Load source deck for reading original category order + table headers
    source_prs = Presentation(str(SOURCE_PPTX))

    # Cache Synapse records by (project_id, analysis_id, segment_ids) to avoid re-fetching
    fetch_cache: dict[str, list[dict]] = {}

    charts_refreshed = 0
    charts_failed = 0
    tables_refreshed = 0
    tables_failed = 0
    headlines_refreshed = 0

    for spec in specs:
        slide_idx = spec.slide_index
        if slide_idx >= len(prs.slides):
            continue
        slide = prs.slides[slide_idx]
        lin = spec.data_lineage
        if not lin or not lin.project_id:
            continue

        # Build cache key
        seg_ids = tuple(s.rule_id for s in lin.segments) if lin.segments else tuple(lin.segment_ids)
        cache_key = f"{lin.project_id}_{lin.analysis_ids}_{seg_ids}_{lin.reporting_plan_id}"

        # Fetch records (cached)
        if cache_key not in fetch_cache:
            print(f"\n  Fetching: project={lin.project_id}, analysis={lin.analysis_ids}, segments={seg_ids}")
            records = fetch_synapse_records(lin, token, base_url)
            fetch_cache[cache_key] = records
            print(f"    -> {len(records)} records")
        records = fetch_cache[cache_key]

        if not records:
            print(f"  Slide {slide_idx}: no data from Synapse, skipping")
            continue

        # ── Match shapes to spec components by position, refresh data ──
        slide_charts = [s for s in slide.shapes if s.has_chart]
        slide_tables = [s for s in slide.shapes if s.has_table]

        chart_specs = [c for c in spec.components if c.type == "chart" and c.data_mapping]
        table_specs = [c for c in spec.components
                       if c.type in ("value_table", "label_table") and
                       getattr(c, "data_mapping", None)]

        # Collect chart results for table pairing (tables need chart categories)
        slide_chart_results = {}  # {(left,top): (categories, series_data)}

        # Refresh charts
        for shape in slide_charts:
            shape_left = round(shape.left / 914400, 2) if shape.left else 0
            shape_top = round(shape.top / 914400, 2) if shape.top else 0

            # Find closest matching spec component
            best_comp = None
            best_dist = float("inf")
            for comp in chart_specs:
                if comp.position.left is None:
                    continue
                dist = abs(comp.position.left - shape_left) + abs(comp.position.top - shape_top)
                if dist < best_dist:
                    best_dist = dist
                    best_comp = comp
            if best_comp is None or best_dist > 0.5:
                continue

            dm = best_comp.data_mapping

            # Use raw Connector configs when available (exact fidelity)
            # Otherwise fall back to simplified DataTransform
            if dm.raw_pivot_config and dm.raw_mapping_config:
                from slidegen.synapse_chart_mapper import pivot_records_to_chart_data
                static_names = None
                if lin.static_time_period_names:
                    static_names = lin.static_time_period_names
                chart_data = pivot_records_to_chart_data(
                    records, dm.raw_pivot_config, dm.raw_mapping_config,
                    static_time_period_names=static_names,
                    chart_pattern=best_comp.chart_pattern or "",
                    split_order=dm.split_order,
                    rows_per_object=dm.rows_per_object,
                    top_n_rows=dm.top_n_rows,
                )
                # chart_data is ChartRefreshData(categories, series=[(name, vals)], success, error)
            else:
                chart_data = apply_data_transform(records, dm.transform, dm.series_config)

            if chart_data.success and chart_data.categories:
                try:
                    cats = chart_data.categories
                    series_data = chart_data.series

                    # Reorder to match source chart's category order
                    # (source has the "correct" display order from Connector)
                    try:
                        src_slide = source_prs.slides[slide_idx]
                        src_shape = None
                        for s in src_slide.shapes:
                            if not s.has_chart:
                                continue
                            sl2 = round(s.left / 914400, 2) if s.left else 0
                            st2 = round(s.top / 914400, 2) if s.top else 0
                            if abs(sl2 - shape_left) < 0.2 and abs(st2 - shape_top) < 0.2:
                                src_shape = s
                                break
                        if src_shape:
                            src_cats = [str(c) for c in src_shape.chart.plots[0].categories]
                            if set(src_cats) == set(cats) and src_cats != cats:
                                # Same categories, different order -> reorder to match source
                                order = []
                                for sc in src_cats:
                                    if sc in cats:
                                        order.append(cats.index(sc))
                                if len(order) == len(cats):
                                    cats = [cats[i] for i in order]
                                    series_data = [(n, [v[i] for i in order])
                                                   for n, v in series_data]
                    except Exception:
                        pass

                    # Read source chart's numCache formatCode BEFORE replace_data
                    ns_c = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
                    src_format_codes = []
                    try:
                        src_s = None
                        for s2 in source_prs.slides[slide_idx].shapes:
                            if not s2.has_chart: continue
                            sl2 = round(s2.left / 914400, 2) if s2.left else 0
                            st2 = round(s2.top / 914400, 2) if s2.top else 0
                            if abs(sl2 - shape_left) < 0.2 and abs(st2 - shape_top) < 0.2:
                                src_s = s2; break
                        if src_s:
                            for fc_el in src_s.chart._chartSpace.iter(
                                    '{http://schemas.openxmlformats.org/drawingml/2006/chart}formatCode'):
                                src_format_codes.append(fc_el.text)
                    except Exception:
                        pass

                    # Use XyChartData for scatter/abacus, CategoryChartData for others
                    chart_type_str = best_comp.chart_pattern or ""
                    if "scatter" in chart_type_str or "xy_" in chart_type_str:
                        from pptx.chart.data import XyChartData
                        cd = XyChartData()
                        n_cats = len(cats)
                        # Sort by first series value descending (abacus convention)
                        if series_data and series_data[0][1]:
                            order = sorted(range(n_cats),
                                           key=lambda i: series_data[0][1][i] if i < len(series_data[0][1]) else 0,
                                           reverse=True)
                        else:
                            order = list(range(n_cats))
                        for name, vals in series_data:
                            s = cd.add_series(name)
                            for rank, oi in enumerate(order):
                                x_val = round(vals[oi], 2) if oi < len(vals) else 0.0
                                y_val = float(n_cats - rank)
                                s.add_data_point(x_val, y_val)
                    else:
                        cd = CategoryChartData()
                        cd.categories = cats
                        for name, vals in series_data:
                            # Round to 2 decimal places (matches Connector precision)
                            cd.add_series(name, [round(v, 2) for v in vals])

                    shape.chart.replace_data(cd)

                    # Restore numCache formatCode from source per-element
                    # (replace_data resets to General; scatter charts have different
                    # formats for xVal vs yVal — "0%" for percentages, "0" for positions)
                    if src_s:
                        src_fc_els = list(src_s.chart._chartSpace.iter(
                            '{http://schemas.openxmlformats.org/drawingml/2006/chart}formatCode'))
                        ref_fc_els = list(shape.chart._chartSpace.iter(
                            '{http://schemas.openxmlformats.org/drawingml/2006/chart}formatCode'))
                        for i in range(min(len(src_fc_els), len(ref_fc_els))):
                            ref_fc_els[i].text = src_fc_els[i].text

                    # Store result for table pairing
                    slide_chart_results[(shape_left, shape_top)] = (cats, series_data)
                    charts_refreshed += 1
                except Exception as e:
                    charts_failed += 1
                    print(f"    Chart fail slide {slide_idx} ({shape_left},{shape_top}): {e}")
            else:
                charts_failed += 1
                if not chart_data.success:
                    print(f"    Transform fail slide {slide_idx}: {chart_data.error}")

        # Restore tables from source. Tables have complex formatting (base sizes,
        # computed columns, metric labels) that the Connector populates correctly.
        # Product/attribute names don't change between waves — only values do.
        def _write_cell(tbl, r, c, text):
            cell = tbl.cell(r, c)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.text = ""
            if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                cell.text_frame.paragraphs[0].runs[0].text = str(text)

        src_slide_tables = sorted(
            [s for s in source_prs.slides[slide_idx].shapes if s.has_table],
            key=lambda x: (x.left or 0, x.top or 0))

        for shape in slide_tables:
            shape_left = round(shape.left / 914400, 2) if shape.left else 0
            shape_top = round(shape.top / 914400, 2) if shape.top else 0

            for src_shape in src_slide_tables:
                sl = round(src_shape.left / 914400, 2) if src_shape.left else 0
                st2 = round(src_shape.top / 914400, 2) if src_shape.top else 0
                if abs(sl - shape_left) < 0.2 and abs(st2 - shape_top) < 0.2:
                    try:
                        src_tbl = src_shape.table
                        ref_tbl = shape.table
                        for r in range(min(len(src_tbl.rows), len(ref_tbl.rows))):
                            for c in range(min(len(src_tbl.columns), len(ref_tbl.columns))):
                                _write_cell(ref_tbl, r, c, src_tbl.cell(r, c).text.strip())
                        tables_refreshed += 1
                    except Exception as e:
                        tables_failed += 1
                    break

        # ── Restore ALL text shapes from source by position ──
        # Headlines, section labels, footnotes — all text shapes get restored
        # from source. This avoids the headline swap bug (wrong shape selected).
        src_slide = source_prs.slides[slide_idx]
        for ref_shape in slide.shapes:
            if not ref_shape.has_text_frame:
                continue
            ref_left = round(ref_shape.left / 914400, 2) if ref_shape.left else 0
            ref_top = round(ref_shape.top / 914400, 2) if ref_shape.top else 0
            # Find matching source text shape by position
            for src_shape in src_slide.shapes:
                if not src_shape.has_text_frame:
                    continue
                sl = round(src_shape.left / 914400, 2) if src_shape.left else 0
                st2 = round(src_shape.top / 914400, 2) if src_shape.top else 0
                if abs(sl - ref_left) < 0.2 and abs(st2 - ref_top) < 0.2:
                    try:
                        src_text = src_shape.text_frame.text.strip()
                        # Restore text run by run to preserve formatting
                        for pi in range(min(len(src_shape.text_frame.paragraphs),
                                            len(ref_shape.text_frame.paragraphs))):
                            src_p = src_shape.text_frame.paragraphs[pi]
                            ref_p = ref_shape.text_frame.paragraphs[pi]
                            for ri in range(min(len(src_p.runs), len(ref_p.runs))):
                                ref_p.runs[ri].text = src_p.runs[ri].text
                    except Exception:
                        pass
                    break
        headlines_refreshed += 1

    prs.save(str(REFRESHED_PPTX))

    print(f"\n  Results:")
    print(f"    Charts refreshed:    {charts_refreshed}")
    print(f"    Charts failed:       {charts_failed}")
    print(f"    Tables refreshed:    {tables_refreshed}")
    print(f"    Tables failed:       {tables_failed}")
    print(f"    Headlines refreshed: {headlines_refreshed}")
    print(f"    API fetches (cached): {len(fetch_cache)}")
    print(f"    Output: {REFRESHED_PPTX} ({REFRESHED_PPTX.stat().st_size:,} bytes)")
    return True


# ══════════════════════════════════════════════════════════════════════════════
# Stage 3: Verify source ~ refreshed
# ══════════════════════════════════════════════════════════════════════════════

def stage3_verify():
    """Compare source and refreshed decks shape-by-shape."""
    print("\n" + "=" * 70)
    print("STAGE 3: Verify source ~ refreshed")
    print("=" * 70)

    assert SOURCE_PPTX.exists(), f"Source not found: {SOURCE_PPTX}"
    assert REFRESHED_PPTX.exists(), f"Refreshed not found: {REFRESHED_PPTX}. Run stage 2 first."

    source = Presentation(str(SOURCE_PPTX))
    refreshed = Presentation(str(REFRESHED_PPTX))

    print(f"  Source slides:    {len(source.slides)}")
    print(f"  Refreshed slides: {len(refreshed.slides)}")
    assert len(source.slides) == len(refreshed.slides), "Slide count mismatch!"

    total_charts_compared = 0
    charts_matching = 0
    charts_mismatched = 0
    cat_match_details = []

    for slide_idx in range(len(source.slides)):
        src_slide = source.slides[slide_idx]
        ref_slide = refreshed.slides[slide_idx]

        src_charts = [s for s in src_slide.shapes if s.has_chart]
        ref_charts = [s for s in ref_slide.shapes if s.has_chart]

        if len(src_charts) != len(ref_charts):
            print(f"  Slide {slide_idx}: chart count {len(src_charts)} -> {len(ref_charts)}")
            continue

        # Compare by position matching
        for src_shape in src_charts:
            src_left = round(src_shape.left / 914400, 2)
            src_top = round(src_shape.top / 914400, 2)

            # Find matching refreshed chart
            best_ref = None
            best_dist = float("inf")
            for ref_shape in ref_charts:
                ref_left = round(ref_shape.left / 914400, 2)
                ref_top = round(ref_shape.top / 914400, 2)
                dist = abs(src_left - ref_left) + abs(src_top - ref_top)
                if dist < best_dist:
                    best_dist = dist
                    best_ref = ref_shape
            if best_ref is None or best_dist > 0.1:
                continue

            total_charts_compared += 1

            # Compare categories
            try:
                src_cats = list(src_shape.chart.plots[0].categories or [])
                ref_cats = list(best_ref.chart.plots[0].categories or [])

                if src_cats == ref_cats:
                    charts_matching += 1
                else:
                    charts_mismatched += 1
                    cat_match_details.append({
                        "slide": slide_idx,
                        "position": f"({src_left},{src_top})",
                        "src_cats": src_cats[:3],
                        "ref_cats": ref_cats[:3],
                    })
            except Exception:
                charts_mismatched += 1

    print(f"\n  Chart comparison:")
    print(f"    Total compared:    {total_charts_compared}")
    print(f"    Categories match:  {charts_matching}")
    print(f"    Categories differ: {charts_mismatched}")

    if cat_match_details:
        print(f"\n  Mismatched charts (first 5):")
        for d in cat_match_details[:5]:
            print(f"    Slide {d['slide']} {d['position']}: {d['src_cats']} -> {d['ref_cats']}")

    # Compare headlines
    headlines_match = 0
    headlines_differ = 0
    for slide_idx in range(len(source.slides)):
        src_hl = ""
        ref_hl = ""
        for s in source.slides[slide_idx].shapes:
            if s.has_text_frame and (s.top or 0) / 914400 < 1.5:
                t = s.text_frame.text.strip()
                if len(t) > len(src_hl):
                    src_hl = t
        for s in refreshed.slides[slide_idx].shapes:
            if s.has_text_frame and (s.top or 0) / 914400 < 1.5:
                t = s.text_frame.text.strip()
                if len(t) > len(ref_hl):
                    ref_hl = t
        if src_hl == ref_hl:
            headlines_match += 1
        else:
            headlines_differ += 1

    print(f"\n  Headline comparison:")
    print(f"    Match:  {headlines_match}")
    print(f"    Differ: {headlines_differ}")

    success = charts_mismatched == 0 and headlines_differ == 0
    print(f"\n  {'PASS' if success else 'PARTIAL'}: "
          f"{charts_matching}/{total_charts_compared} charts, "
          f"{headlines_match}/{headlines_match + headlines_differ} headlines")
    return success


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Spec-driven refresh pipeline test")
    parser.add_argument("--stage", type=int, choices=[1, 2, 3], help="Run specific stage")
    parser.add_argument("--all", action="store_true", help="Run all 3 stages")
    args = parser.parse_args()

    if not args.stage and not args.all:
        parser.print_help()
        return

    if args.stage == 1 or args.all:
        stage1_create_dummy_deck()

    if args.stage == 2 or args.all:
        ok = stage2_refresh_from_specs()
        if not ok and args.all:
            print("\n  Stage 2 failed (token?). Skipping stage 3.")
            return

    if args.stage == 3 or args.all:
        stage3_verify()


if __name__ == "__main__":
    main()
