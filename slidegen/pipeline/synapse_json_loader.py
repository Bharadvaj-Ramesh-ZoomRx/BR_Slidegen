"""
synapse_json_loader.py — Fetch aggregated data directly as JSON from Synapse.

Calls POST /reports/generate per analysis_id and restructures the response
into the same {desc, prior, current} format that Excel extraction produces.
This bypasses Excel entirely for Tier 1 data — no pandas, no openpyxl.

Usage:
    from slidegen.pipeline.synapse_json_loader import fetch_data_as_json
    data = fetch_data_as_json(config, api_key="...")
    # data has same structure as load_all_data() output — renderers unchanged

Config:
    extractions:
      - id: ryb_mr
        method: synapse_report
        params:
          analysis_id: 301
          reporting_plan_id: 50
          time_period_map: { prior: 501, current: 502 }
          segment_ids: [1]
"""

from __future__ import annotations

import json
import logging
import os
import time
from datetime import datetime
from typing import Optional

from slidegen.pipeline.project_config import ProjectConfig

logger = logging.getLogger(__name__)

# Prefer requests; fall back to urllib
try:
    import requests as _requests
    _USE_REQUESTS = True
except ImportError:
    import urllib.request as _urllib_req
    import urllib.error as _urllib_err
    _USE_REQUESTS = False

_ENV_KEY_NAME = "SYNAPSE_API_KEY"
_POLL_INTERVAL = 5
_MAX_POLL_WAIT = 120


# ── Public API ──────────────────────────────────────────────────────────────

def fetch_data_as_json(
    config: ProjectConfig,
    api_key: Optional[str] = None,
    synapse_url: Optional[str] = None,
) -> dict:
    """Fetch all synapse_report extractions as JSON and return slidegen-format data.

    For each extraction with method='synapse_report', calls POST /reports/generate
    on the Synapse API and restructures the response records into
    [{desc, prior, current, code}] — same format as Excel extractors.

    Also builds a _sheets index from the question codes in responses so
    downstream stages can discover available codes.

    Args:
        config: Loaded ProjectConfig instance.
        api_key: Synapse Bearer token. Falls back to SYNAPSE_API_KEY env var.
        synapse_url: API base URL override.

    Returns:
        dict mapping extraction_id → list[dict], plus _sheets and _sample_sizes.
    """
    synapse = config.synapse
    resolved_key = api_key or os.environ.get(_ENV_KEY_NAME, "")
    if not resolved_key:
        raise ValueError(
            f"API key not provided. Pass api_key= or set {_ENV_KEY_NAME} env var."
        )

    resolved_url = (
        synapse_url
        or (synapse.api_url if synapse else "")
        or getattr(config, "synapse_api_url", "")
    ).rstrip("/")
    if not resolved_url:
        raise ValueError("No Synapse API URL available.")

    headers = {
        "Authorization": f"Bearer {resolved_key}",
        "Content-Type": "application/json",
    }

    # Collect synapse_report extractions
    synapse_extractions = [
        ex for ex in config.extractions if ex.method == "synapse_report"
    ]
    if not synapse_extractions:
        logger.info("No synapse_report extractions in config — nothing to fetch")
        return {}

    data = {}
    sheets_index = {}  # question_code → desc for _sheets rebuilding

    for ex in synapse_extractions:
        params = ex.params
        analysis_id = params.get("analysis_id")
        if not analysis_id:
            logger.warning("[%s] synapse_report missing params.analysis_id — skipping", ex.id)
            continue

        reporting_plan_id = params.get("reporting_plan_id")
        if not reporting_plan_id:
            logger.warning("[%s] synapse_report missing params.reporting_plan_id — skipping", ex.id)
            continue

        time_period_map = params.get("time_period_map", {})
        # time_period_map: {"prior": 501, "current": 502} — deliverable/time_period IDs
        tp_ids = list(time_period_map.values())
        seg_ids = params.get("segment_ids", synapse.segment_ids if synapse else [])

        print(f"  [{ex.id}] Fetching analysis_id={analysis_id} from Synapse...")

        report = _fetch_report(
            resolved_url, headers,
            analysis_id=analysis_id,
            project_id=synapse.project_id if synapse else params.get("project_id"),
            reporting_plan_id=reporting_plan_id,
            time_period_ids=tp_ids,
            segment_ids=seg_ids,
            setup_type=params.get("setup_type", "STATIC"),
            time_period_mode=params.get("time_period_mode", "DELIVERABLE"),
        )

        if report is None:
            logger.warning("[%s] Report fetch returned None — skipping", ex.id)
            data[ex.id] = []
            continue

        # Restructure records into slidegen format
        rows = _restructure_records(report, time_period_map, params)
        data[ex.id] = rows
        print(f"    → {len(rows)} rows extracted")

        # Collect for _sheets index
        q_code = report.get("question_code", "")
        q_text = report.get("question_text", "")
        if q_code:
            sheets_index[q_code] = q_text

    # Build a minimal _sheets structure for downstream discovery
    if sheets_index:
        data["_sheets"] = {
            "synapse": [
                {"row": i, "code": code, "desc": desc}
                for i, (code, desc) in enumerate(sheets_index.items())
            ]
        }

    data["_sample_sizes"] = config.sample_sizes
    return data


# ── Report fetching ────────────────────────────────────────────────────────

def _fetch_report(
    base_url: str,
    headers: dict,
    *,
    analysis_id: int,
    project_id: int,
    reporting_plan_id: int,
    time_period_ids: list[int],
    segment_ids: list[int],
    setup_type: str = "STATIC",
    time_period_mode: str = "DELIVERABLE",
) -> dict | None:
    """Call POST /reports/generate and return the report response.

    Handles both synchronous (light) and cached (heavy) reports.
    For heavy reports, polls the cached report endpoint until ready.
    """
    url = f"{base_url}/reports/generate"
    payload = {
        "analysis_ids": [analysis_id],
        "project_id": project_id,
        "reporting_plan_id": reporting_plan_id,
        "setup_type": setup_type,
        "time_period_ids": time_period_ids,
        "segment_ids": segment_ids,
        "time_period_mode": time_period_mode,
        "include_overall": False,
        "rollup_time_periods": False,
        "include_live_wave": False,
    }

    report = _post_json(url, headers, payload)

    # Handle cached/processing reports — poll until ready
    if report.get("is_cached_report") and report.get("cache_status") == "PROCESSING":
        cached_id = report.get("cached_report_id")
        if cached_id:
            print(f"    Report is processing (cached_report_id={cached_id}), polling...")
            report = _poll_cached_report(base_url, headers, cached_id)

    # Verify we have records
    if not report or not report.get("records"):
        return report

    return report


def _poll_cached_report(
    base_url: str, headers: dict, cached_report_id: int
) -> dict | None:
    """Poll GET /reports/cached/{id} until the report is ready."""
    url = f"{base_url}/reports/cached/{cached_report_id}"
    elapsed = 0

    while elapsed < _MAX_POLL_WAIT:
        time.sleep(_POLL_INTERVAL)
        elapsed += _POLL_INTERVAL

        report = _get_json(url, headers)
        status = report.get("cache_status", "")

        if status == "PROCESSED":
            return report
        elif status == "PROCESSING":
            print(f"    [{elapsed}s] Still processing...")
            continue
        else:
            logger.warning("Unexpected cache_status: %s", status)
            return report

    logger.warning("Cached report %d not ready after %ds", cached_report_id, _MAX_POLL_WAIT)
    return None


# ── Record restructuring ──────────────────────────────────────────────────

def _restructure_records(
    report: dict,
    time_period_map: dict,
    params: dict,
) -> list[dict]:
    """Convert Synapse report records into slidegen [{desc, prior, current}] format.

    The report.records are flat dicts with dimension keys (code, option, segment_1, ...)
    and metric keys (count, percentage, ...). Time periods appear as separate columns
    or as repeated records with a time_period dimension.

    time_period_map: {"prior": tp_id, "current": tp_id} maps our labels to time_period IDs.
    """
    records = report.get("records", [])
    if not records:
        return []

    time_periods = report.get("time_periods", [])
    # Build tp_id → our label mapping
    tp_id_to_label = {}
    for label, tp_id in time_period_map.items():
        tp_id_to_label[tp_id] = label
        tp_id_to_label[str(tp_id)] = label

    # Build tp_name → our label mapping (fallback)
    tp_name_to_label = {}
    for tp in time_periods:
        tp_id = tp.get("id")
        tp_name = tp.get("name", "")
        if tp_id in tp_id_to_label:
            tp_name_to_label[tp_name] = tp_id_to_label[tp_id]

    # Determine the metric key to use for percentage values
    metric_key = params.get("metric_key", "percentage")
    pct_mode = params.get("pct_mode", "pct")

    # Detect structure: do records have time_period dimension or separate columns?
    # Check if records have time_period_1, time_period_2, etc. keys
    sample = records[0] if records else {}
    has_tp_columns = any(k.startswith("time_period_") for k in sample.keys())

    if has_tp_columns:
        return _restructure_tp_columns(records, time_periods, tp_id_to_label, metric_key, pct_mode)
    else:
        # Records may be grouped by a time_period dimension in the data
        return _restructure_flat_records(records, tp_name_to_label, metric_key, pct_mode)


def _restructure_tp_columns(
    records: list[dict],
    time_periods: list[dict],
    tp_id_to_label: dict,
    metric_key: str,
    pct_mode: str,
) -> list[dict]:
    """Restructure when time periods are separate columns (time_period_1, time_period_2, ...)."""
    # Map time_period_N column → our label (prior/current)
    tp_col_to_label = {}
    for i, tp in enumerate(time_periods, 1):
        tp_id = tp.get("id")
        label = tp_id_to_label.get(tp_id, tp_id_to_label.get(str(tp_id)))
        if label:
            tp_col_to_label[f"time_period_{i}"] = label

    results = []
    for rec in records:
        desc = rec.get("option", rec.get("code", ""))
        if not desc or str(desc).lower() in ("base", "total", ""):
            continue

        row = {"desc": str(desc).strip(), "code": rec.get("code", "")}

        for tp_col, label in tp_col_to_label.items():
            val = rec.get(tp_col)
            if val is not None:
                row[label] = _convert_value(val, pct_mode)

        # Only include rows that have at least one period value
        if "current" in row or "prior" in row:
            results.append(row)

    return results


def _restructure_flat_records(
    records: list[dict],
    tp_name_to_label: dict,
    metric_key: str,
    pct_mode: str,
) -> list[dict]:
    """Restructure when records are flat (one row per option, metric columns inline)."""
    results = []
    for rec in records:
        desc = rec.get("option", rec.get("code", ""))
        if not desc or str(desc).lower() in ("base", "total", ""):
            continue

        row = {"desc": str(desc).strip(), "code": rec.get("code", "")}

        # Look for percentage/metric value
        val = rec.get(metric_key)
        if val is not None:
            row["current"] = _convert_value(val, pct_mode)

        if "current" in row:
            results.append(row)

    return results


def _convert_value(val, pct_mode: str) -> float | None:
    """Convert a value to percentage based on pct_mode."""
    if val is None:
        return None
    try:
        v = float(val)
    except (ValueError, TypeError):
        return None

    if pct_mode == "straight":
        return round(v, 1)
    else:
        # Default: decimal → percentage (0.45 → 45.0)
        return round(v * 100, 1)


# ── HTTP helpers ──────────────────────────────────────────────────────────

def _post_json(url: str, headers: dict, payload: dict) -> dict:
    """POST JSON and return parsed response."""
    body = json.dumps(payload).encode("utf-8")

    if _USE_REQUESTS:
        resp = _requests.post(url, data=body, headers=headers, timeout=60)
        if not resp.ok:
            detail = resp.text
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                pass
            raise RuntimeError(f"Synapse API error {resp.status_code}: {detail}")
        return resp.json()
    else:
        req = _urllib_req.Request(url, data=body, headers=headers, method="POST")
        try:
            with _urllib_req.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e


def _get_json(url: str, headers: dict) -> dict:
    """GET and return parsed JSON."""
    if _USE_REQUESTS:
        resp = _requests.get(url, headers=headers, timeout=30)
        if not resp.ok:
            raise RuntimeError(f"Synapse API error {resp.status_code}: {resp.text}")
        return resp.json()
    else:
        req = _urllib_req.Request(url, headers=headers)
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e
