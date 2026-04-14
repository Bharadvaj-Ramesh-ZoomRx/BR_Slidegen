"""
raw_data_first.py — Track D: Raw-Data-First integration via synapse-cli.

Downloads ALL survey responses + VQs + segment assignments in 1-2 API calls,
caches locally as JSON, then aggregates per-extraction on demand.

This is an additive track — existing Tracks A/B/C are untouched.

Requires: pip install -e /path/to/synapse-cli
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
from datetime import datetime, timezone
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from slidegen.pipeline.project_config import ProjectConfig, DataExtractionConfig

logger = logging.getLogger(__name__)


# ── Cache helpers ─────────────────────────────────────────────────────────

def _cache_path(config: "ProjectConfig") -> str:
    """Return the path for the raw-data-first JSON cache file."""
    context = config.context_path or "context"
    return os.path.join(context, "raw_data_first.json")


def _config_hash(config: "ProjectConfig") -> str:
    """Hash the synapse config fields that affect the download.

    Cache invalidates if any of these change.
    """
    syn = config.synapse
    if not syn:
        return ""
    blob = json.dumps({
        "project_id": syn.project_id,
        "survey_ids": syn.survey_ids,
        "deliverable_ids": syn.deliverable_ids,
        "reporting_plan_id": syn.reporting_plan_id,
        "segment_ids": syn.segment_ids,
        "wave_ids": syn.wave_ids,
    }, sort_keys=True)
    return hashlib.md5(blob.encode()).hexdigest()[:12]


def _load_cache(path: str, expected_hash: str) -> Optional[dict]:
    """Load cached raw data if the config hash matches."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            cached = json.load(f)
        if cached.get("_meta", {}).get("config_hash") == expected_hash:
            logger.info("  raw_data_first: using cached data (%s)", path)
            return cached
        logger.info("  raw_data_first: cache stale (config changed), re-downloading")
    except (json.JSONDecodeError, KeyError):
        logger.warning("  raw_data_first: cache corrupt, re-downloading")
    return None


def _save_cache(path: str, raw_data: dict, config_hash: str) -> None:
    """Persist raw data with metadata for cache validation."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    raw_data.setdefault("_meta", {})
    raw_data["_meta"]["config_hash"] = config_hash
    raw_data["_meta"]["fetched_at"] = datetime.now(timezone.utc).isoformat()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(raw_data, f)
    logger.info("  raw_data_first: cached → %s", path)


# ── Download ──────────────────────────────────────────────────────────────

def fetch_raw_data_first(config: "ProjectConfig", api_key: Optional[str] = None) -> dict:
    """Download ALL survey responses + VQs + segments via raw-data-first.

    Returns the synapse-cli aggregator's columnar structure::

        {
            "columns": {idx: {"code": str, "text": str}},
            "code_map": {code: [idx, ...]},
            "respondents": [{id, user_id, quarter, values, segments}, ...],
            "quarters": [str, ...],
            "_meta": {config_hash, fetched_at, respondent_count, code_count, ...}
        }

    Cached at ``context/{wave}/raw_data_first.json``.  Invalidated when
    the synapse config section changes.

    Args:
        config: Loaded ProjectConfig with a ``synapse`` block.
        api_key: Optional explicit Bearer token.

    Returns:
        Parsed columnar raw data structure.

    Raises:
        ImportError: If synapse-cli is not installed.
        ValueError: If config is missing required synapse fields.
        RuntimeError: If the API call fails.
    """
    from slidegen.pipeline.synapse_client_factory import get_client

    try:
        from synapse_cli.aggregator import parse_ndjson_stream
        from synapse_cli.workflows.fetch_for_plan import (
            _download_ndjson, resolve_deliverable_ids,
            _fetch_and_join_segments,
        )
    except ImportError:
        raise ImportError(
            "synapse-cli is not installed.  Run:\n"
            "  pip install -e /path/to/synapse-cli"
        )

    # Check cache
    cache = _cache_path(config)
    cfg_hash = _config_hash(config)
    cached = _load_cache(cache, cfg_hash)
    if cached is not None:
        return cached

    # Build client
    client = get_client(config, api_key=api_key)
    syn = config.synapse
    api_calls = 0

    # Resolve deliverable IDs from reporting plan + current/prior quarters
    deliverable_ids = syn.deliverable_ids
    if not deliverable_ids and syn.reporting_plan_id:
        deliverable_ids = resolve_deliverable_ids(
            client,
            reporting_plan_id=syn.reporting_plan_id,
            quarter=config.period_current,
            include_prior=config.period_prior,
        )

    if not deliverable_ids:
        raise ValueError(
            "No deliverable_ids resolved — set synapse.deliverable_ids in config "
            "or ensure synapse.reporting_plan_id has deliverables matching "
            f"'{config.period_current}' / '{config.period_prior}'"
        )

    # Download ALL responses + VQs in one NDJSON call
    survey_id = syn.survey_ids[0] if syn.survey_ids else None
    if not survey_id:
        raise ValueError("synapse.survey_ids must have at least one entry")

    logger.info("  raw_data_first: downloading NDJSON (survey=%d, deliverables=%s, include_vqs=True)",
                survey_id, deliverable_ids)
    ndjson_text = _download_ndjson(
        client,
        survey_id=survey_id,
        deliverable_ids=deliverable_ids,
        include_vqs=True,
    )
    api_calls += 1

    # Parse into columnar structure
    raw_data = parse_ndjson_stream(ndjson_text)
    logger.info("  raw_data_first: parsed %d respondents, %d codes, quarters=%s",
                len(raw_data.get("respondents", [])),
                len(raw_data.get("code_map", {})),
                raw_data.get("quarters", []))

    # Fetch + join segment assignments (separate call)
    segment_ids = syn.segment_ids or []
    if segment_ids:
        try:
            _fetch_and_join_segments(
                client,
                raw_data=raw_data,
                project_id=syn.project_id,
                reporting_plan_id=syn.reporting_plan_id,
                segment_ids=segment_ids,
            )
            api_calls += 1
            logger.info("  raw_data_first: joined segments (%d segment IDs)", len(segment_ids))
        except Exception as e:
            logger.warning("  raw_data_first: segment fetch failed (non-fatal): %s", e)

    # Store metadata
    raw_data["_meta"] = {
        "api_calls": api_calls,
        "respondent_count": len(raw_data.get("respondents", [])),
        "code_count": len(raw_data.get("code_map", {})),
        "quarters": raw_data.get("quarters", []),
    }

    # Cache
    _save_cache(cache, raw_data, cfg_hash)
    return raw_data


# ── Aggregation ───────────────────────────────────────────────────────────

def aggregate_extraction(
    raw_data: dict,
    extraction: "DataExtractionConfig",
    config: "ProjectConfig",
) -> list[dict]:
    """Aggregate a single extraction config against raw data.

    Reads ``extraction.params`` for question codes, aggregation function,
    optional segment cuts, and mode.  Returns the standard
    ``[{desc, prior, current, delta, code}]`` list that all renderers consume.

    Supported modes (mirrors ``raw_aggregate``):

    - ``single``: One code → one row with prior/current/delta.
    - ``multi_code``: Multiple codes → one row each.
    - ``cross_brand``: Same code aggregated with two different segment
      filters (primary vs competitor brand proxy).
    - ``by_segment``: One code broken out by segment values.

    Args:
        raw_data: Output of ``fetch_raw_data_first()``.
        extraction: A ``DataExtractionConfig`` with ``method="raw_data_first"``.
        config: ProjectConfig (for period labels).

    Returns:
        List of dicts with ``desc``, ``prior``, ``current``, ``delta`` keys.
    """
    from synapse_cli.aggregator import aggregate_code

    params = extraction.params
    agg_fn = params.get("agg_fn", params.get("agg", "recall_pct"))
    quarter_current = params.get("quarter_current", config.period_current)
    quarter_prior = params.get("quarter_prior", config.period_prior)
    segment_column = params.get("segment_column")
    segment_value = params.get("segment_value")
    mode = params.get("mode", "single")

    if mode == "multi_code":
        return _aggregate_multi_code(
            raw_data, params["codes"], agg_fn,
            quarter_current, quarter_prior,
            segment_column, segment_value,
        )
    elif mode == "cross_brand":
        return _aggregate_cross_brand(
            raw_data, params, agg_fn,
            quarter_current, quarter_prior,
        )
    elif mode == "by_segment":
        return _aggregate_by_segment(
            raw_data, params, agg_fn,
            quarter_current,
        )
    else:  # single
        return _aggregate_single(
            raw_data, params, agg_fn,
            quarter_current, quarter_prior,
            segment_column, segment_value,
        )


def _aggregate_single(
    raw_data, params, agg_fn,
    quarter_current, quarter_prior,
    segment_column, segment_value,
) -> list[dict]:
    """Single code → [{desc, prior, current, delta}]."""
    from synapse_cli.aggregator import aggregate_code

    code = params.get("code", "")
    if not code:
        codes = params.get("codes", [])
        if codes:
            return _aggregate_multi_code(
                raw_data, codes, agg_fn,
                quarter_current, quarter_prior,
                segment_column, segment_value,
            )
        return []

    row = aggregate_code(
        raw_data, code=code, agg_fn=agg_fn,
        quarter_current=quarter_current,
        quarter_prior=quarter_prior,
        segment_column=segment_column,
        segment_value=segment_value,
    )
    return [row] if row.get("current") is not None or row.get("prior") is not None else []


def _aggregate_multi_code(
    raw_data, codes, agg_fn,
    quarter_current, quarter_prior,
    segment_column=None, segment_value=None,
) -> list[dict]:
    """Multiple codes → one row each."""
    from synapse_cli.aggregator import aggregate_code

    rows = []
    for code in codes:
        row = aggregate_code(
            raw_data, code=code, agg_fn=agg_fn,
            quarter_current=quarter_current,
            quarter_prior=quarter_prior,
            segment_column=segment_column,
            segment_value=segment_value,
        )
        if row.get("current") is not None or row.get("prior") is not None:
            rows.append(row)
    return rows


def _aggregate_cross_brand(
    raw_data, params, agg_fn,
    quarter_current, quarter_prior,
) -> list[dict]:
    """Same code with two segment filters (primary vs competitor)."""
    from synapse_cli.aggregator import aggregate_code

    code = params["code"]
    primary_seg_col = params.get("primary_segment_column")
    primary_seg_val = params.get("primary_segment_value")
    comp_seg_col = params.get("comp_segment_column")
    comp_seg_val = params.get("comp_segment_value")

    rows = []
    for label, seg_col, seg_val in [
        ("primary", primary_seg_col, primary_seg_val),
        ("competitor", comp_seg_col, comp_seg_val),
    ]:
        row = aggregate_code(
            raw_data, code=code, agg_fn=agg_fn,
            quarter_current=quarter_current,
            quarter_prior=quarter_prior,
            segment_column=seg_col,
            segment_value=seg_val,
        )
        row["brand"] = label
        rows.append(row)
    return rows


def _aggregate_by_segment(
    raw_data, params, agg_fn,
    quarter_current,
) -> list[dict]:
    """One code broken out by segment values (current wave only)."""
    from synapse_cli.aggregator import aggregate_code

    code = params["code"]
    segment_column = params["segment_column"]
    segment_values = params.get("segment_values", [])

    rows = []
    for seg_val in segment_values:
        row = aggregate_code(
            raw_data, code=code, agg_fn=agg_fn,
            quarter_current=quarter_current,
            quarter_prior=quarter_current,  # same quarter — segment comparison, not QoQ
            segment_column=segment_column,
            segment_value=seg_val,
        )
        row["desc"] = seg_val  # segment value becomes the label
        rows.append(row)
    return rows


# ── Discovery (for skills / config generation) ───────────────────────────

def get_available_codes(config: "ProjectConfig", api_key: Optional[str] = None) -> list[dict]:
    """List all question codes available in the raw data.

    Useful for skills and config generators to discover what codes exist
    without needing to know the Excel layout.

    Returns:
        List of {code, text, column_count, sample_values} dicts.
    """
    from synapse_cli.aggregator import get_available_codes as _get_codes

    raw_data = fetch_raw_data_first(config, api_key=api_key)
    return _get_codes(raw_data)
