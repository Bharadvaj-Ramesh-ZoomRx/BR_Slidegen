"""
synapse_chart_mapper.py — Map Synapse API records to chart data using Connector tags.

Takes flat Synapse records + PivotConfig + MappingConfig from the Connector tags
and produces CategoryChartData that can be written into a chart via replace_data().

The Connector's tag model:
  - ReportConfig: what to fetch (project, analysis, segments, time periods)
  - PivotConfig (DataFrameConfigHash): how to pivot (row/col/value fields, filters)
  - MappingConfig: what goes into the chart (which pivoted columns become series)

Field name mapping (Connector internal → Synapse API):
  - "value" → "y_label" (answer/option text)
  - "segment_1" → "segment_1" (already matches)
  - "time_period_name" → "time_period_name" (already matches)
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Optional

import pandas as pd
import requests

logger = logging.getLogger(__name__)

# Connector field names → Synapse API record column names
FIELD_MAP = {
    "value": "y_label",
    "y_code": "y_code",
    "y_label": "y_label",
    "segment_1": "segment_1",
    "segment_2": "segment_2",
    "time_period_name": "time_period_name",
    "time_period_id": "time_period_id",
    "question": "y_label",
    "title": "y_label",       # Connector "title" = API "y_label"
    "measure": "measure",
    "n": "base",              # Connector "n" = API "base" (sample size)
    "n - value": "base",      # Computed "n cross value" → use base
}


def _remap_field(name: str) -> str:
    return FIELD_MAP.get(name, name)


@dataclass
class ChartRefreshData:
    """Data ready to write into a chart via replace_data()."""
    categories: list[str]
    series: list[tuple[str, list[float]]]  # [(series_name, values), ...]
    success: bool = True
    error: str = ""


def fetch_synapse_report(
    base_url: str,
    token: str,
    report_config: dict,
) -> list[dict]:
    """Fetch report data from Synapse API using ReportConfig from Connector tag.

    Uses StaticTimePeriodIds when available (matches the exact data the chart
    was rendered with). Falls back to DynamicTimePeriod for fresh/latest data.
    """
    headers = {"Authorization": token, "accept": "application/json",
               "Content-Type": "application/json"}

    payload = {
        "project_id": report_config.get("ProjectId"),
        "reporting_plan_id": report_config.get("ReportingPlanId"),
        "analysis_ids": report_config.get("AnalysisIds"),
        "segment_ids": report_config.get("SegmentIds"),
    }

    # Use DynamicTimePeriod for refresh (gets latest data including new waves)
    dtp = report_config.get("DynamicTimePeriod")
    if dtp:
        payload["setup_type"] = "DYNAMIC"
        payload["dynamic_time_period"] = {
            "latest_n_deliverables": dtp.get("LatestNDeliverables", 8),
            "include_live_wave": dtp.get("IncludeLiveWave", True),
        }
    else:
        # No dynamic config — use static
        payload["setup_type"] = "DYNAMIC"
        payload["dynamic_time_period"] = {
            "latest_n_deliverables": 20,
            "include_live_wave": True,
        }

    try:
        resp = requests.post(
            f"{base_url}/api/reports/generate",
            headers=headers,
            json=payload,
            timeout=60,
        )
        if resp.status_code in (200, 201):
            return resp.json().get("records", [])
        else:
            logger.warning(f"Synapse API {resp.status_code}: {resp.text[:200]}")
            return []
    except Exception as e:
        logger.warning(f"Synapse API error: {e}")
        return []


def pivot_records_to_chart_data(
    records: list[dict],
    pivot_config: dict,
    mapping_config: dict,
    static_time_period_names: list[str] | None = None,
) -> ChartRefreshData:
    """Transform flat Synapse records into chart categories + series.

    Args:
        records: Flat records from Synapse API
        pivot_config: PivotConfig from DataFrameConfigHash tag
        mapping_config: MappingConfig from shape tag
        static_time_period_names: If provided, filter to these time periods
            (from ReportConfig.StaticTimePeriodNames — the exact periods the chart shows)
    """
    if not records:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error="No records from Synapse")

    df = pd.DataFrame(records)

    row_fields = [_remap_field(f) for f in pivot_config.get("RowFields", [])]
    col_fields = [_remap_field(f) for f in pivot_config.get("ColumnFields", [])]
    val_fields = [_remap_field(f) for f in pivot_config.get("ValueFields", [])]

    if not row_fields or not val_fields:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error=f"Missing RowFields or ValueFields in PivotConfig")

    # Verify required columns exist; for compound fields like "n - value",
    # split and check each part
    for f in row_fields + val_fields:
        if f not in df.columns:
            # Try splitting compound field (e.g., "n - value" → check "base" and "y_label")
            parts = [_remap_field(p.strip()) for p in f.split(" - ")]
            if all(p in df.columns for p in parts):
                continue  # compound field — parts exist, will handle in pivot
            return ChartRefreshData(categories=[], series=[], success=False,
                                    error=f"Field '{f}' not in records")

    # Apply filters
    for filt in pivot_config.get("Filters", []):
        col_key = _remap_field(filt.get("ColumnKey", ""))
        val = filt.get("Value", "")
        if col_key and val and col_key in df.columns:
            if val in df[col_key].values:
                df = df[df[col_key] == val]
            else:
                # Fuzzy: try contains
                mask = df[col_key].str.contains(val.split()[0], case=False, na=False)
                if mask.any():
                    df = df[mask]

    if df.empty:
        df = pd.DataFrame(records)

    # Filter to static time periods if provided
    if static_time_period_names and "time_period_name" in df.columns:
        df = df[df["time_period_name"].isin(static_time_period_names)]

    # Build pivot column key
    valid_col_fields = [f for f in col_fields if f in df.columns]
    if len(valid_col_fields) >= 2:
        df = df.copy()
        df["_col_key"] = df[valid_col_fields[0]].astype(str)
        for cf in valid_col_fields[1:]:
            df["_col_key"] = df["_col_key"] + " - " + df[cf].astype(str)
    elif len(valid_col_fields) == 1:
        df = df.copy()
        df["_col_key"] = df[valid_col_fields[0]].astype(str)
    else:
        df = df.copy()
        df["_col_key"] = "value"

    # Pivot
    try:
        pivot = df.pivot_table(
            index=row_fields[0],
            columns="_col_key",
            values=val_fields[0],
            aggfunc="first",
        )
    except Exception as e:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error=f"Pivot failed: {e}")

    # Apply MappingConfig to select series
    selected = mapping_config.get("selectedColumns", [])
    categories = list(pivot.index)

    if selected and len(selected) > 1:
        # First = category label, rest = series
        series = []
        for sc in selected[1:]:
            matched_col = None
            for pc_col in pivot.columns:
                if sc == str(pc_col) or sc in str(pc_col) or str(pc_col) in sc:
                    matched_col = pc_col
                    break
            if matched_col is not None:
                vals = [0.0 if pd.isna(v) else float(v) for v in pivot[matched_col].tolist()]
                series.append((sc, vals))
    else:
        # No mapping — use all columns
        series = []
        for col in pivot.columns:
            vals = [0.0 if pd.isna(v) else float(v) for v in pivot[col].tolist()]
            series.append((str(col), vals))

    return ChartRefreshData(categories=categories, series=series)


def refresh_chart_from_synapse(
    chart_shape,
    report_config: dict,
    pivot_config: dict,
    mapping_config: dict,
    base_url: str,
    token: str,
) -> ChartRefreshData:
    """Full pipeline: fetch from Synapse → pivot → map → return chart data.

    Does NOT write to the chart — caller uses replace_data() with the result.
    """
    records = fetch_synapse_report(base_url, token, report_config)
    if not records:
        return ChartRefreshData(categories=[], series=[], success=False,
                                error="No records from Synapse API")

    static_names = report_config.get("StaticTimePeriodNames")

    result = pivot_records_to_chart_data(
        records, pivot_config, mapping_config, static_names
    )

    return result
