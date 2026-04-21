"""
test_nonconnected_refresh.py — End-to-end test of non-connected slide refresh.

Tests the full pipeline for slides WITHOUT Connector tags:
  Step 1: generate_nonconnected_spec() — extract layout from OOXML, create spec template
  Step 2: complete_spec_from_user_input() — user provides data lineage, spec gets completed
  Step 3: refresh_nonconnected_deck() — refresh the deck using the completed spec
  Step 4: Compare source vs refreshed

Usage:
    python -m slidegen.synapse_auth --update "Bearer eyJ..."
    python tests/test_nonconnected_refresh.py
"""
from __future__ import annotations

import json
import shutil
import sys
from dataclasses import asdict
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

SAMPLE_DIR = REPO_ROOT.parent / "Sample Decks"
SOURCE_PPTX = SAMPLE_DIR / "Repatha ATU Slide 6.pptx"
OUTPUT_PPTX = SAMPLE_DIR / "Repatha_ATU_Slide6_nonconnected_refresh.pptx"
SPEC_TEMPLATE_JSON = SAMPLE_DIR / "Repatha_ATU_Slide6_spec_template.json"
SPEC_COMPLETE_JSON = SAMPLE_DIR / "Repatha_ATU_Slide6_spec_complete.json"

# User-provided data lineage (they know the Synapse project even without tags)
DATA_LINEAGE = {
    "project_id": 1428,
    "project_name": "Amgen [ATU]: Repatha",
    "reporting_plan_id": 574,
    "analysis_ids": [545991],
    "segment_ids": [14633],
    "dynamic_latest_n": 2,
}

# Map python-pptx chart types to our chart_pattern keys
_CHART_TYPE_MAP = {
    XL_CHART_TYPE.BAR_STACKED_100: "bar_stacked_100_horizontal",
    XL_CHART_TYPE.BAR_CLUSTERED: "bar_clustered_horizontal",
    XL_CHART_TYPE.COLUMN_CLUSTERED: "column_clustered_vertical",
    XL_CHART_TYPE.COLUMN_STACKED_100: "column_stacked_100_vertical",
    XL_CHART_TYPE.LINE_MARKERS: "line_markers_trended",
    XL_CHART_TYPE.LINE: "line_markers_trended",
    XL_CHART_TYPE.DOUGHNUT: "doughnut_default",
    XL_CHART_TYPE.XY_SCATTER: "xy_scatter_abacus",
}


# ══════════════════════════════════════════════════════════════════════════════
# Step 1: Generate spec template from OOXML (no Synapse data needed)
# ══════════════════════════════════════════════════════════════════════════════

def generate_nonconnected_spec(pptx_path: str) -> "SlideSpec":
    """Extract chart/table layout from OOXML and create a spec template.

    For each chart: extracts chart_pattern, series_names, series_colors,
    category_count from the OOXML structure.
    For each table: extracts row_count, col_count, headers (first row).

    Creates a SlideSpec with:
      - ChartComponent with chrome + position but data=None, data_mapping=None
      - Table components with position + dimensions but no data
      - DataLineage with all fields empty (user fills in later)
      - spec_completeness="layout_complete_data_missing"
    """
    from slidegen.slide_spec.schema import (
        SlideSpec, HeadlineSpec, Position, DataLineage,
        ChartComponent, ChartChrome, DataLabelsSpec, AxisSpec, LegendSpec,
        LabelTableComponent, ValueTableComponent,
        dump_spec,
    )

    print("\n  Extracting layout from OOXML...")
    prs = Presentation(pptx_path)
    slide = prs.slides[0]

    components = []
    chart_index = 0
    table_index = 0

    # ── Extract headline ──
    headline_text = "Non-Connected Slide"
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        top_in = (shape.top or 0) / 914400
        text = shape.text_frame.text.strip()
        if top_in < 1.5 and len(text) > 10:
            headline_text = text
            break

    # ── Extract charts ──
    for shape in slide.shapes:
        if not shape.has_chart:
            continue

        chart = shape.chart
        plot = chart.plots[0]
        cats = [str(c) for c in plot.categories] if plot.categories else []
        series_names = [str(s.name) for s in plot.series]

        # Extract series colors from OOXML
        series_colors = []
        for s in plot.series:
            try:
                rgb = s.format.fill.fore_color.rgb
                series_colors.append(f"#{rgb}")
            except Exception:
                series_colors.append("#999999")

        chart_pattern = _CHART_TYPE_MAP.get(chart.chart_type, "bar_stacked_100_horizontal")

        # Extract position in inches
        pos = Position(
            left=round(shape.left / 914400, 2),
            top=round(shape.top / 914400, 2),
            width=round(shape.width / 914400, 2),
            height=round(shape.height / 914400, 2),
        )

        # Extract data label format from OOXML
        ns_c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        fmt_code = "0%"
        for fc_el in chart._chartSpace.iter(f"{{{ns_c}}}formatCode"):
            fmt_code = fc_el.text
            break

        # Extract gap_width and overlap from chart XML
        gap_width = None
        overlap = None
        try:
            for gw in chart._chartSpace.iter(f"{{{ns_c}}}gapWidth"):
                gap_width = int(gw.get("val", 150))
            for ov in chart._chartSpace.iter(f"{{{ns_c}}}overlap"):
                overlap = int(ov.get("val", 0))
        except Exception:
            pass

        chrome = ChartChrome(
            title=None,
            legend=LegendSpec(show=False),
            gridlines=False,
            data_labels=DataLabelsSpec(show=True, format=fmt_code, position="inEnd"),
            value_axis=AxisSpec(format=fmt_code, show=False),
            hide_category_labels=True,
            gap_width=gap_width,
            overlap=overlap,
        )

        comp = ChartComponent(
            position=pos,
            chart_pattern=chart_pattern,
            data=None,           # no data in template
            chrome=chrome,
            data_mapping=None,   # no mapping yet — user must provide lineage first
        )
        components.append(comp)

        print(f"    Chart {chart_index} @({pos.left}, {pos.top}): {chart_pattern}")
        print(f"      series: {series_names}")
        print(f"      colors: {series_colors}")
        print(f"      categories[{len(cats)}]: {cats[:4]}")
        chart_index += 1

    # ── Extract tables ──
    for shape in slide.shapes:
        if not shape.has_table:
            continue

        tbl = shape.table
        row_count = len(tbl.rows)
        col_count = len(tbl.columns)

        # Extract headers (first row text)
        headers = []
        if row_count > 0:
            for ci in range(col_count):
                cell_text = tbl.cell(0, ci).text.strip()
                headers.append(cell_text)

        pos = Position(
            left=round(shape.left / 914400, 2),
            top=round(shape.top / 914400, 2),
            width=round(shape.width / 914400, 2),
            height=round(shape.height / 914400, 2),
        )

        # Heuristic: narrow single-column tables are label tables
        if col_count == 1 or (col_count <= 2 and shape.width / 914400 < 3.0):
            comp = LabelTableComponent(
                position=pos,
                labels=[],  # empty in template
                alternating_rows=True,
                data_mapping=None,
            )
            print(f"    LabelTable {table_index} @({pos.left}, {pos.top}): {row_count}r x {col_count}c")
        else:
            comp = ValueTableComponent(
                position=pos,
                headers=headers,
                rows=[],
                row_count=row_count,
                col_count=col_count,
                data_mapping=None,
            )
            print(f"    ValueTable {table_index} @({pos.left}, {pos.top}): {row_count}r x {col_count}c, headers={headers[:3]}")

        components.append(comp)
        table_index += 1

    # ── Build spec template ──
    spec = SlideSpec(
        slide_id="nonconnected_001",
        slide_index=0,
        layout="observed_custom",
        headline=HeadlineSpec(text=headline_text),
        components=components,
        brand=None,
        section=None,
        data_lineage=DataLineage(),  # empty — user fills in
        spec_completeness="layout_complete_data_missing",
    )

    # Save as JSON template with instructions
    spec_dict = asdict(spec)
    spec_dict["_instructions"] = (
        "This is a spec template for a non-connected slide. "
        "Fill in data_lineage with your Synapse project_id, reporting_plan_id, "
        "analysis_ids, segment_ids, and dynamic_latest_n. Then run "
        "complete_spec_from_user_input() to populate data mappings."
    )
    spec_dict["data_lineage"]["_instructions"] = (
        "Fill in: project_id, project_name, reporting_plan_id, analysis_ids, "
        "segment_ids (or segments[]), dynamic_latest_n. "
        "Example: project_id=1428, analysis_ids=[545991], segment_ids=[14633], dynamic_latest_n=2"
    )

    template_path = Path(pptx_path).with_suffix(".spec_template.json")
    with open(template_path, "w", encoding="utf-8") as f:
        json.dump(spec_dict, f, indent=2, ensure_ascii=False)
    print(f"    Spec template saved: {template_path.name}")

    return spec


# ══════════════════════════════════════════════════════════════════════════════
# Step 2: Complete spec from user input (data lineage + inference)
# ══════════════════════════════════════════════════════════════════════════════

def complete_spec_from_user_input(
    spec: "SlideSpec",
    user_input: dict,
    pptx_path: str,
) -> "SlideSpec":
    """Takes a partial spec + user-provided data lineage and completes it.

    Steps:
      a. Populate spec.data_lineage from user_input
      b. Fetch Synapse records using the data lineage
      c. For each chart: extract series_names from OOXML, run infer_data_transform(),
         create ChartDataMapping with inferred transform + series_config
      d. For each table: infer column mapping based on proximity to charts
      e. Set spec_completeness="complete"

    user_input format:
        {
            "project_id": 1428,
            "project_name": "Amgen [ATU]: Repatha",
            "reporting_plan_id": 574,
            "analysis_ids": [545991],
            "segment_ids": [14633],
            "dynamic_latest_n": 2,
        }
    """
    import os
    import requests
    import pandas as pd
    from dotenv import load_dotenv

    from slidegen.slide_spec.schema import (
        DataLineage, SegmentRule,
        ChartDataMapping, DataTransform, SeriesConfig, DataFilter,
        TableDataMapping, TableColumnConfig,
        ChartComponent, LabelTableComponent, ValueTableComponent,
        dump_spec,
    )
    from slidegen.slide_spec.data_inference import infer_data_transform, match_series_to_column

    load_dotenv(REPO_ROOT / ".env")

    # ── (a) Populate data_lineage ──
    print("\n  Populating data lineage from user input...")
    lineage = DataLineage(
        project_id=user_input.get("project_id"),
        project_name=user_input.get("project_name"),
        reporting_plan_id=user_input.get("reporting_plan_id"),
        analysis_ids=user_input.get("analysis_ids", []),
        segment_ids=user_input.get("segment_ids", []),
        dynamic_latest_n=user_input.get("dynamic_latest_n"),
        include_live_wave=True,
    )
    spec.data_lineage = lineage
    print(f"    project_id={lineage.project_id}, analysis_ids={lineage.analysis_ids}")

    # ── (b) Fetch Synapse records ──
    print("\n  Fetching Synapse records...")
    token = os.environ.get("SYNAPSE_API_TOKEN", "")
    base_url = os.environ.get("SYNAPSE_API_BASE_URL", "https://synapse-api.zoomrx.com")

    if not token:
        raise RuntimeError(
            "No SYNAPSE_API_TOKEN. Run: python -m slidegen.synapse_auth --update 'Bearer eyJ...'"
        )

    # Validate token expiry
    try:
        import base64
        import time as _time
        raw = token.replace("Bearer ", "").strip()
        payload_b64 = raw.split(".")[1] + "=="
        claims = json.loads(base64.urlsafe_b64decode(payload_b64))
        remaining = (claims["exp"] - _time.time()) / 60
        if remaining < 1:
            raise RuntimeError(f"Token expired {abs(remaining):.0f} min ago")
        print(f"    Token valid for {remaining:.0f} min")
    except (KeyError, IndexError, json.JSONDecodeError):
        pass  # non-JWT token, skip validation

    headers = {
        "Authorization": token if token.startswith("Bearer") else f"Bearer {token}",
        "accept": "application/json",
        "Content-Type": "application/json",
    }

    # Build segment_ids for API call
    segment_ids = lineage.segment_ids if lineage.segment_ids else []
    if lineage.segments:
        segment_ids = [s.rule_id for s in lineage.segments]

    payload = {
        "project_id": lineage.project_id,
        "reporting_plan_id": lineage.reporting_plan_id,
        "analysis_ids": lineage.analysis_ids,
        "segment_ids": segment_ids,
        "setup_type": "DYNAMIC",
        "dynamic_time_period": {
            "latest_n_deliverables": lineage.dynamic_latest_n or 2,
            "include_live_wave": True,
        },
    }

    resp = requests.post(
        f"{base_url}/api/reports/generate",
        headers=headers, json=payload, timeout=60,
    )
    if resp.status_code not in (200, 201):
        raise RuntimeError(f"Synapse API error {resp.status_code}: {resp.text[:300]}")

    records = resp.json().get("records", [])
    print(f"    Fetched {len(records)} records")
    if not records:
        raise RuntimeError("Synapse returned 0 records — check analysis_ids and segment_ids")

    df = pd.DataFrame(records)
    print(f"    Columns: {sorted(df.columns.tolist())[:10]}")

    # Build column values map for inference
    record_columns = {}
    for col in df.columns:
        try:
            unique = df[col].dropna().unique()
            if df[col].dtype == object or len(unique) < 50:
                record_columns[col] = [str(v) for v in unique]
        except Exception:
            pass

    # ── (c) Re-read OOXML to get series_names per chart ──
    print("\n  Matching chart series against Synapse columns...")
    prs = Presentation(pptx_path)
    slide = prs.slides[0]

    # Collect chart shapes sorted by position (same order as components)
    chart_shapes = sorted(
        [s for s in slide.shapes if s.has_chart],
        key=lambda x: (x.left or 0, x.top or 0),
    )

    # Collect chart components from spec (same order)
    chart_components = [c for c in spec.components if isinstance(c, ChartComponent)]

    for ci, (shape, comp) in enumerate(zip(chart_shapes, chart_components)):
        chart = shape.chart
        plot = chart.plots[0]
        series_names = [str(s.name) for s in plot.series]
        cats = [str(c) for c in plot.categories] if plot.categories else []

        # Extract series colors from OOXML for SeriesConfig
        series_colors = []
        for s in plot.series:
            try:
                rgb = s.format.fill.fore_color.rgb
                series_colors.append(f"#{rgb}")
            except Exception:
                series_colors.append("#999999")

        print(f"\n    Chart {ci}: series={series_names}")

        try:
            transform, inferred_series_config = infer_data_transform(
                series_names=series_names,
                chart_pattern=comp.chart_pattern,
                record_columns=record_columns,
                category_count=len(cats),
            )

            # Merge OOXML colors into inferred SeriesConfig
            for si, sc in enumerate(inferred_series_config):
                if si < len(series_colors) and not sc.color:
                    sc.color = series_colors[si]

            print(f"      transform: row={transform.row_field}, col={transform.column_field}, val={transform.value_field}")
            print(f"      filters: {[(f.field, f.value) for f in transform.filters]}")
            print(f"      series_config: {[(s.role, s.color) for s in inferred_series_config]}")

            comp.data_mapping = ChartDataMapping(
                transform=transform,
                series_config=inferred_series_config,
            )

        except ValueError as e:
            print(f"      INFERENCE FAILED: {e}")
            # Leave data_mapping as None — this chart won't be refreshed

    # ── (d) Infer table mappings based on proximity to charts ──
    print("\n  Inferring table data mappings...")
    table_shapes = sorted(
        [s for s in slide.shapes if s.has_table],
        key=lambda x: (x.left or 0, x.top or 0),
    )
    table_components = [
        c for c in spec.components
        if isinstance(c, (LabelTableComponent, ValueTableComponent))
    ]

    for ti, comp in enumerate(table_components):
        if isinstance(comp, LabelTableComponent):
            # Label tables get categories from the nearest chart
            # Find the closest chart component by vertical position
            nearest_chart = None
            min_dist = float("inf")
            for cc in chart_components:
                if cc.data_mapping and cc.data_mapping.transform.row_field:
                    dist = abs((comp.position.top or 0) - (cc.position.top or 0))
                    if dist < min_dist:
                        min_dist = dist
                        nearest_chart = cc

            if nearest_chart:
                tx = nearest_chart.data_mapping.transform
                comp.data_mapping = TableDataMapping(
                    filters=list(tx.filters),
                    columns=[
                        TableColumnConfig(
                            source_field=tx.row_field,
                            header_template="",
                        )
                    ],
                )
                print(f"    LabelTable {ti}: mapped to row_field={tx.row_field} from nearest chart")
            else:
                print(f"    LabelTable {ti}: no nearby chart found, skipping")

        elif isinstance(comp, ValueTableComponent):
            # Value tables are harder — try to map to nearest chart's values
            nearest_chart = None
            min_dist = float("inf")
            for cc in chart_components:
                if cc.data_mapping and cc.data_mapping.transform.value_field:
                    dist = abs((comp.position.top or 0) - (cc.position.top or 0))
                    if dist < min_dist:
                        min_dist = dist
                        nearest_chart = cc

            if nearest_chart and comp.headers:
                tx = nearest_chart.data_mapping.transform
                table_cols = []
                for header in comp.headers:
                    # First column is usually the label column
                    if not table_cols:
                        table_cols.append(TableColumnConfig(
                            source_field=tx.row_field,
                            header_template=header,
                        ))
                    else:
                        table_cols.append(TableColumnConfig(
                            source_field=tx.value_field,
                            header_template=header,
                        ))
                comp.data_mapping = TableDataMapping(
                    filters=list(tx.filters),
                    columns=table_cols,
                )
                print(f"    ValueTable {ti}: mapped {len(table_cols)} columns from nearest chart")
            else:
                print(f"    ValueTable {ti}: no mapping inferred, skipping")

    # ── (e) Mark spec complete ──
    spec.spec_completeness = "complete"

    # Save completed spec
    complete_path = Path(pptx_path).with_suffix(".spec_complete.json")
    from slidegen.slide_spec.schema import dump_spec
    dump_spec(spec, complete_path)
    print(f"\n    Completed spec saved: {complete_path.name}")

    return spec


# ══════════════════════════════════════════════════════════════════════════════
# Step 3: Refresh deck using completed spec
# ══════════════════════════════════════════════════════════════════════════════

def refresh_nonconnected_deck(pptx_path: str, spec: "SlideSpec") -> str:
    """Refresh a non-connected deck using a completed spec.

    For each chart component with data_mapping:
      - Fetches Synapse data from spec.data_lineage
      - Applies data_transform via apply_data_transform()
      - Uses CategoryChartData to replace chart data
      - Restores formatCode from source

    For each label table:
      - Uses chart categories from the nearest refreshed chart

    Returns the output file path.
    """
    import os
    import requests
    import pandas as pd
    from dotenv import load_dotenv

    from slidegen.slide_spec.schema import (
        ChartComponent, LabelTableComponent, ValueTableComponent,
    )

    # Import apply_data_transform from test_spec_refresh_pipeline (same directory)
    import importlib.util
    _spec_refresh_path = Path(__file__).parent / "test_spec_refresh_pipeline.py"
    _mod_spec = importlib.util.spec_from_file_location(
        "test_spec_refresh_pipeline", _spec_refresh_path,
    )
    _mod = importlib.util.module_from_spec(_mod_spec)
    sys.modules["test_spec_refresh_pipeline"] = _mod  # register so @dataclass works
    _mod_spec.loader.exec_module(_mod)
    apply_data_transform = _mod.apply_data_transform

    load_dotenv(REPO_ROOT / ".env")

    assert spec.spec_completeness == "complete", (
        f"Spec is not complete: {spec.spec_completeness}"
    )
    assert spec.data_lineage and spec.data_lineage.project_id, "Spec has no data lineage"

    # ── (a) Clone source PPTX ──
    output_path = str(OUTPUT_PPTX)
    shutil.copy2(pptx_path, output_path)
    print(f"\n  Cloned source -> {Path(output_path).name}")

    # ── (b) Fetch Synapse records ──
    print("  Fetching Synapse data for refresh...")
    token = os.environ.get("SYNAPSE_API_TOKEN", "")
    base_url = os.environ.get("SYNAPSE_API_BASE_URL", "https://synapse-api.zoomrx.com")

    if not token:
        raise RuntimeError("No SYNAPSE_API_TOKEN")

    lineage = spec.data_lineage
    headers = {
        "Authorization": token if token.startswith("Bearer") else f"Bearer {token}",
        "accept": "application/json",
        "Content-Type": "application/json",
    }

    segment_ids = lineage.segment_ids if lineage.segment_ids else []
    if lineage.segments:
        segment_ids = [s.rule_id for s in lineage.segments]

    payload = {
        "project_id": lineage.project_id,
        "reporting_plan_id": lineage.reporting_plan_id,
        "analysis_ids": lineage.analysis_ids,
        "segment_ids": segment_ids,
        "setup_type": "DYNAMIC",
        "dynamic_time_period": {
            "latest_n_deliverables": lineage.dynamic_latest_n or 2,
            "include_live_wave": True,
        },
    }

    resp = requests.post(
        f"{base_url}/api/reports/generate",
        headers=headers, json=payload, timeout=60,
    )
    if resp.status_code not in (200, 201):
        raise RuntimeError(f"Synapse API error {resp.status_code}: {resp.text[:300]}")

    records = resp.json().get("records", [])
    print(f"    Fetched {len(records)} records")

    # ── (c) Refresh charts ──
    print("\n  Refreshing charts...")
    out_prs = Presentation(output_path)
    out_slide = out_prs.slides[0]

    # Sort chart shapes by position (matches component order)
    out_chart_shapes = sorted(
        [s for s in out_slide.shapes if s.has_chart],
        key=lambda x: (x.left or 0, x.top or 0),
    )
    chart_components = [c for c in spec.components if isinstance(c, ChartComponent)]

    # Track refreshed chart data for label table updates
    refreshed_chart_data = {}  # chart_index -> RefreshChartData

    ns_c = "http://schemas.openxmlformats.org/drawingml/2006/chart"

    for ci, (shape, comp) in enumerate(zip(out_chart_shapes, chart_components)):
        if not comp.data_mapping:
            print(f"    Chart {ci}: skipped (no data_mapping)")
            continue

        transform = comp.data_mapping.transform
        series_config = comp.data_mapping.series_config

        # Apply data transform
        chart_data = apply_data_transform(records, transform, series_config)

        if not chart_data.success or not chart_data.categories:
            print(f"    Chart {ci}: transform failed: {chart_data.error}")
            continue

        refreshed_chart_data[ci] = chart_data

        # Read source formatCode before replace_data
        src_fmt = None
        for fc_el in shape.chart._chartSpace.iter(f"{{{ns_c}}}formatCode"):
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
                for fc_el in shape.chart._chartSpace.iter(f"{{{ns_c}}}formatCode"):
                    fc_el.text = src_fmt

            print(f"    Chart {ci}: OK — {len(chart_data.categories)} cats, {len(chart_data.series)} series")
            print(f"      cats: {chart_data.categories[:4]}")
            print(f"      series: {[s[0] for s in chart_data.series]}")
            if chart_data.series:
                print(f"      vals[0]: {[round(v, 3) for v in chart_data.series[0][1]][:4]}")
        except Exception as e:
            print(f"    Chart {ci}: FAILED: {e}")

    # ── (d) Refresh label tables ──
    print("\n  Refreshing tables...")
    out_table_shapes = sorted(
        [s for s in out_slide.shapes if s.has_table],
        key=lambda x: (x.left or 0, x.top or 0),
    )
    table_components = [
        c for c in spec.components
        if isinstance(c, (LabelTableComponent, ValueTableComponent))
    ]

    for ti, (shape, comp) in enumerate(zip(out_table_shapes, table_components)):
        if isinstance(comp, LabelTableComponent):
            # Find nearest refreshed chart to get categories
            nearest_ci = None
            min_dist = float("inf")
            for ci, cd in refreshed_chart_data.items():
                chart_comp = chart_components[ci]
                dist = abs((comp.position.top or 0) - (chart_comp.position.top or 0))
                if dist < min_dist:
                    min_dist = dist
                    nearest_ci = ci

            if nearest_ci is not None:
                cats = refreshed_chart_data[nearest_ci].categories
                tbl = shape.table
                n_data_rows = len(tbl.rows) - 1  # exclude header row if present
                n_cats = len(cats)

                # Update label cells — skip header row (row 0)
                updated = 0
                for ri in range(min(n_data_rows, n_cats)):
                    try:
                        cell = tbl.cell(ri + 1, 0) if len(tbl.rows) > ri + 1 else None
                        if cell is None:
                            cell = tbl.cell(ri, 0)
                        # Preserve formatting: update only the text
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                run.text = ""
                        if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                            cell.text_frame.paragraphs[0].runs[0].text = cats[ri]
                        else:
                            cell.text_frame.text = cats[ri]
                        updated += 1
                    except Exception:
                        pass

                print(f"    LabelTable {ti}: updated {updated}/{n_cats} labels from chart {nearest_ci}")
            else:
                print(f"    LabelTable {ti}: no refreshed chart data available, skipped")

        elif isinstance(comp, ValueTableComponent):
            print(f"    ValueTable {ti}: skipped (value table refresh not yet implemented)")

    # ── Save ──
    out_prs.save(output_path)
    print(f"\n  Output saved: {output_path}")
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# Main orchestrator
# ══════════════════════════════════════════════════════════════════════════════

def main():
    from dotenv import load_dotenv
    load_dotenv(REPO_ROOT / ".env")

    print("=" * 60)
    print("Non-Connected Slide Refresh Pipeline")
    print("=" * 60)

    assert SOURCE_PPTX.exists(), f"Source not found: {SOURCE_PPTX}"

    # ── Step 1: Generate spec template ──
    print("\nStep 1: Generate spec template from OOXML")
    spec = generate_nonconnected_spec(str(SOURCE_PPTX))
    print(f"  Spec template: {spec.spec_completeness}")
    print(f"  Components: {len(spec.components)}")
    for c in spec.components:
        print(f"    - {c.type} @({c.position.left}, {c.position.top})")

    # ── Step 2: Complete spec with user input ──
    print("\n" + "=" * 60)
    print("Step 2: Complete spec from user input")
    print("=" * 60)
    print(f"  User input: project_id={DATA_LINEAGE['project_id']}, "
          f"analysis_ids={DATA_LINEAGE['analysis_ids']}")

    try:
        spec = complete_spec_from_user_input(spec, DATA_LINEAGE, str(SOURCE_PPTX))
        print(f"\n  Spec completeness: {spec.spec_completeness}")

        # Show data mapping summary
        from slidegen.slide_spec.schema import ChartComponent
        for i, c in enumerate(spec.components):
            if isinstance(c, ChartComponent) and c.data_mapping:
                tx = c.data_mapping.transform
                sc = c.data_mapping.series_config
                print(f"  Chart {i}: row={tx.row_field}, col={tx.column_field}, "
                      f"val={tx.value_field}, series={[s.role for s in sc]}")
    except RuntimeError as e:
        print(f"\n  ERROR: {e}")
        return

    # ── Step 3: Refresh deck ──
    print("\n" + "=" * 60)
    print("Step 3: Refresh deck using completed spec")
    print("=" * 60)

    try:
        output_path = refresh_nonconnected_deck(str(SOURCE_PPTX), spec)
    except RuntimeError as e:
        print(f"\n  ERROR: {e}")
        return

    # ── Step 4: Compare source vs refreshed ──
    print("\n" + "=" * 60)
    print("Step 4: Compare source vs refreshed")
    print("=" * 60)

    src_prs = Presentation(str(SOURCE_PPTX))
    ref_prs = Presentation(output_path)

    src_charts = sorted(
        [s for s in src_prs.slides[0].shapes if s.has_chart],
        key=lambda x: (x.left or 0, x.top or 0),
    )
    ref_charts = sorted(
        [s for s in ref_prs.slides[0].shapes if s.has_chart],
        key=lambda x: (x.left or 0, x.top or 0),
    )

    for ci in range(min(len(src_charts), len(ref_charts))):
        sp = src_charts[ci].chart.plots[0]
        rp = ref_charts[ci].chart.plots[0]

        s_cats = [str(c) for c in sp.categories]
        r_cats = [str(c) for c in rp.categories]
        s_ns = len(sp.series)
        r_ns = len(rp.series)

        cats_match = s_cats == r_cats
        series_match = s_ns == r_ns

        print(f"  Chart {ci}: cats={'OK' if cats_match else 'DIFF'} "
              f"series={'OK' if series_match else 'DIFF'}")

        if not cats_match:
            print(f"    src cats: {s_cats[:4]}")
            print(f"    ref cats: {r_cats[:4]}")
        if not series_match:
            print(f"    src series: {s_ns}, ref series: {r_ns}")

        # Compare first series values
        s_v = [round(float(v), 3) if v else 0 for v in sp.series[0].values][:4]
        r_v = [round(float(v), 3) if v else 0 for v in rp.series[0].values][:4]
        vals_match = s_v == r_v
        print(f"    vals[0]: src={s_v} ref={r_v} {'OK' if vals_match else 'DIFF'}")

    print(f"\n  Source:   {SOURCE_PPTX}")
    print(f"  Output:   {output_path}")
    print(f"\nDone.")


if __name__ == "__main__":
    main()
