"""
synapse_raw_fetcher.py — Fetch raw respondent-level data from the Synapse API.

Downloads raw survey responses as JSON (via NDJSON streaming endpoint),
converts to pandas DataFrames, and caches as pkl for fast reuse and validation.

Also fetches:
  - Segment definitions for the project
  - Virtual question responses (auto-discovered per project)

Usage:
    from slidegen.pipeline.synapse_raw_fetcher import fetch_all_raw
    result = fetch_all_raw(config)
    # result: {"pkl_path": str, "dataframes": {sheet: DataFrame}, ...}

Config (synapse section):
    synapse:
      api_url: "https://synapse.zoomrx.com/api"
      project_id: 123
      survey_ids: [456]
      reporting_plan_id: 50
      deliverable_ids: [501, 502]
      segment_ids: [1, 2]
      segments:
        - id: 1
          name: "Practice Setting"
          mode: "groupby"
        - id: 2
          name: "Region"
          mode: "filter"
          values: ["Northeast"]
"""

from __future__ import annotations

import csv
import io
import json
import logging
import os
import pickle
import shutil
from datetime import datetime
from typing import Optional

from slidegen.pipeline.project_config import ProjectConfig, SynapseConfig, SegmentCutConfig

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


# ── Public API ──────────────────────────────────────────────────────────────

def fetch_all_raw(
    config: ProjectConfig,
    api_key: Optional[str] = None,
    synapse_url: Optional[str] = None,
) -> dict:
    """Fetch all raw data from Synapse: survey responses + segments + VQs.

    Downloads responses as JSON (NDJSON stream), converts to pandas DataFrames,
    and caches everything as a pkl file for fast reuse.

    Args:
        config: Loaded ProjectConfig (must have synapse section).
        api_key: Bearer token. Falls back to SYNAPSE_API_KEY env var.
        synapse_url: API base URL override.

    Returns:
        {
            "pkl_path": str,                    # path to cached pkl file
            "dataframes": {sheet_name: DataFrame},  # per-sheet DataFrames
            "segments": [SegmentInfo],           # segment definitions from API
            "vq_data": {qid: [{...}]},           # VQ responses
            "vq_questions": [{qid, title, ...}], # VQ metadata
        }
    """
    synapse, resolved_url, headers = _resolve_params(config, api_key, synapse_url)

    result = {}

    # 1. Download raw survey responses as JSON → DataFrames → pkl
    print("Fetching raw survey responses from Synapse (JSON)...")
    dataframes, pkl_path = fetch_survey_responses_json(config, synapse, resolved_url, headers)
    result["dataframes"] = dataframes
    result["pkl_path"] = pkl_path

    # 2. Fetch segment definitions
    print("Fetching segment definitions...")
    segments = fetch_segments(synapse, resolved_url, headers)
    result["segments"] = segments

    # 3. Auto-discover and fetch VQ data
    print("Fetching virtual question data...")
    vq_questions, vq_data = fetch_virtual_questions(synapse, resolved_url, headers)
    result["vq_questions"] = vq_questions
    result["vq_data"] = vq_data

    return result


def fetch_survey_responses_json(
    config: ProjectConfig,
    synapse: SynapseConfig,
    resolved_url: str,
    headers: dict,
) -> tuple[dict, str]:
    """Download raw survey responses as JSON and convert to DataFrames.

    Calls POST /surveys/download-responses-json which returns NDJSON.
    Each survey sheet becomes a separate DataFrame. The result is cached
    as a pkl file for fast reloading.

    Returns:
        (dataframes, pkl_path) — dict of {sheet_name: DataFrame} and path to pkl cache.
    """
    import pandas as pd

    url = f"{resolved_url}/surveys/download-responses-json"
    payload = {
        "survey_ids": synapse.survey_ids,
        "deliverable_ids": synapse.deliverable_ids,
        "partial_response": False,
        "screened_out": False,
        "reset": False,
        "ta_test_users": False,
        "simulated_panelist": False,
    }

    print(f"  Downloading JSON responses for surveys {synapse.survey_ids}...")
    ndjson_lines = _post_ndjson(url, headers, payload)

    # Parse NDJSON into per-sheet DataFrames
    # Format: header lines have type="header", data lines have type="row"
    dataframes = {}
    current_columns = []
    current_sheet = ""
    current_rows = []
    header_metadata = {}  # sheet_name → header_rows for validation

    for line_str in ndjson_lines:
        line_str = line_str.strip()
        if not line_str:
            continue
        obj = json.loads(line_str)

        if obj.get("type") == "header":
            # Flush previous sheet
            if current_sheet and current_rows:
                df = pd.DataFrame(current_rows, columns=current_columns)
                dataframes[current_sheet] = df
                print(f"    {current_sheet}: {len(df)} respondents, {len(current_columns)} columns")

            current_sheet = obj.get("sheet_name", f"survey_{obj.get('survey_id', 'unknown')}")
            current_columns = obj.get("columns", [])
            current_rows = []
            header_metadata[current_sheet] = obj.get("header_rows", [])

        elif obj.get("type") == "row":
            values = obj.get("values", [])
            # Pad or trim to match column count
            if len(values) < len(current_columns):
                values.extend([None] * (len(current_columns) - len(values)))
            elif len(values) > len(current_columns):
                values = values[:len(current_columns)]
            current_rows.append(values)

    # Flush last sheet
    if current_sheet and current_rows:
        df = pd.DataFrame(current_rows, columns=current_columns)
        dataframes[current_sheet] = df
        print(f"    {current_sheet}: {len(df)} respondents, {len(current_columns)} columns")

    if not dataframes:
        print("  [WARN] No response data received from Synapse JSON endpoint")

    # Save as pkl cache
    pkl_path = _raw_pkl_path(config)
    os.makedirs(os.path.dirname(pkl_path), exist_ok=True)

    # Backup existing pkl
    if os.path.exists(pkl_path):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = pkl_path.replace(".pkl", f"_backup_{ts}.pkl")
        shutil.copy2(pkl_path, backup_path)
        print(f"  Backed up existing pkl → {os.path.basename(backup_path)}")

    cache_data = {
        "_meta": {
            "fetched_at": datetime.now().isoformat(),
            "survey_ids": synapse.survey_ids,
            "deliverable_ids": synapse.deliverable_ids,
            "source": "synapse_json_api",
        },
        "_header_metadata": header_metadata,
        "dataframes": dataframes,
    }

    with open(pkl_path, "wb") as f:
        pickle.dump(cache_data, f, protocol=pickle.HIGHEST_PROTOCOL)

    size_mb = os.path.getsize(pkl_path) / (1024 * 1024)
    print(f"  Saved pkl cache: {pkl_path} ({size_mb:.1f} MB)")

    return dataframes, pkl_path


def load_cached_pkl(config: ProjectConfig) -> dict | None:
    """Load cached raw data from pkl file if it exists.

    Returns the full cache dict with 'dataframes', '_meta', etc., or None.
    Useful for validation and reuse without re-fetching from Synapse.
    """
    pkl_path = _raw_pkl_path(config)
    if not os.path.exists(pkl_path):
        return None

    try:
        with open(pkl_path, "rb") as f:
            cached = pickle.load(f)
        meta = cached.get("_meta", {})
        dfs = cached.get("dataframes", {})
        total_rows = sum(len(df) for df in dfs.values())
        print(f"  Loaded pkl cache: {pkl_path}")
        print(f"    {len(dfs)} sheets, {total_rows} total rows")
        print(f"    Fetched: {meta.get('fetched_at', 'unknown')}")
        return cached
    except (pickle.UnpicklingError, EOFError, Exception) as e:
        logger.warning("Failed to load pkl cache %s: %s", pkl_path, e)
        return None


def fetch_segments(
    synapse: SynapseConfig,
    resolved_url: str,
    headers: dict,
) -> list[dict]:
    """Fetch segment definitions via GET /projects/{id}/segments/list.

    Returns list of {"segment_id": int, "segment_name": str, "type": str}.
    """
    url = f"{resolved_url}/projects/{synapse.project_id}/segments/list"
    response = _get_json(url, headers)

    segments = response.get("segments", [])
    print(f"  Found {len(segments)} segments for project {synapse.project_id}")
    for seg in segments:
        print(f"    [{seg['segment_id']}] {seg['segment_name']} ({seg['type']})")

    return segments


def fetch_virtual_questions(
    synapse: SynapseConfig,
    resolved_url: str,
    headers: dict,
) -> tuple[list[dict], dict]:
    """Auto-discover and export all VQs for the project's surveys.

    For each survey_id:
      1. GET /surveys/virtual-questions/{survey_id} → list VQ metadata
      2. POST /virtual-questions/export → download VQ response CSV

    Returns:
        (vq_questions, vq_data)
        - vq_questions: [{qid, title, question, type, survey_id}]
        - vq_data: {qid: [{"users_wave_id": int, "value": str, ...}]}
    """
    all_vq_questions = []
    all_vq_data = {}

    for survey_id in synapse.survey_ids:
        # List VQs for this survey
        vq_list = _list_virtual_questions(resolved_url, headers, survey_id)
        if not vq_list:
            print(f"  No virtual questions found for survey {survey_id}")
            continue

        print(f"  Survey {survey_id}: {len(vq_list)} virtual questions")
        for vq in vq_list:
            vq["survey_id"] = survey_id
        all_vq_questions.extend(vq_list)

        # Export VQ responses
        qids = [vq["qid"] for vq in vq_list]
        wave_ids = _resolve_wave_ids(synapse, resolved_url, headers)

        if qids and wave_ids:
            vq_responses = _export_virtual_questions(
                resolved_url, headers, qids, survey_id, wave_ids
            )
            all_vq_data.update(vq_responses)
            total_rows = sum(len(rows) for rows in vq_responses.values())
            print(f"    Exported {total_rows} VQ response rows across {len(vq_responses)} questions")

    return all_vq_questions, all_vq_data


# ── Internal helpers ──────────────────────────────────────────────────────

def _resolve_params(
    config: ProjectConfig,
    api_key: Optional[str],
    synapse_url: Optional[str],
) -> tuple[SynapseConfig, str, dict]:
    """Validate and resolve synapse config, URL, and headers."""
    synapse = config.synapse
    if not synapse:
        raise ValueError(
            "No 'synapse' section in config.yaml. "
            "Add synapse config to enable raw data fetching."
        )

    from slidegen.pipeline.synapse_auth import resolve_api_key
    resolved_key = resolve_api_key(api_key)

    resolved_url = (
        synapse_url or synapse.api_url or ""
    ).rstrip("/")
    if not resolved_url:
        raise ValueError("No Synapse API URL available.")

    headers = {
        "Authorization": f"Bearer {resolved_key}",
        "Content-Type": "application/json",
    }
    return synapse, resolved_url, headers


def _raw_pkl_path(config: ProjectConfig) -> str:
    """Determine path for the pkl cache file."""
    if config.context_path:
        return os.path.join(config.context_path, "source_raw_data.pkl")
    base_dir = os.path.dirname(
        config.raw_data_source_path or config.data_source_path
    )
    return os.path.join(base_dir, "source_raw_data.pkl")


def _list_virtual_questions(
    base_url: str, headers: dict, survey_id: int
) -> list[dict]:
    """GET /surveys/virtual-questions/{survey_id} with pagination."""
    all_questions = []
    page = 1
    page_size = 100

    while True:
        url = f"{base_url}/surveys/virtual-questions/{survey_id}?page={page}&page_size={page_size}"
        response = _get_json(url, headers)
        questions = response.get("questions", [])
        all_questions.extend(questions)

        meta = response.get("meta", {})
        total_pages = meta.get("total_pages", 1)
        if page >= total_pages:
            break
        page += 1

    return all_questions


def _export_virtual_questions(
    base_url: str, headers: dict,
    qids: list[int], survey_id: int, wave_ids: list[int],
) -> dict[int, list[dict]]:
    """POST /virtual-questions/export → parse CSV into per-qid response dicts."""
    url = f"{base_url}/virtual-questions/export"
    payload = {
        "qids": qids,
        "survey_id": survey_id,
        "wave_ids": wave_ids,
    }

    csv_text = _post_csv(url, headers, payload)
    if not csv_text:
        return {}

    # Parse CSV: columns vary by VQ type, but typically include
    # users_wave_id and response columns per question
    reader = csv.DictReader(io.StringIO(csv_text))
    vq_data = {}

    for row in reader:
        # Each row is a respondent; columns are question responses
        resp_id = row.get("users_wave_id") or row.get("id")
        if not resp_id:
            continue

        for qid in qids:
            qid_str = str(qid)
            # VQ export may use qid or question title as column header
            value = row.get(qid_str)
            if value is None:
                # Try matching by partial key
                for col_name, col_val in row.items():
                    if qid_str in col_name:
                        value = col_val
                        break

            if value is not None:
                if qid not in vq_data:
                    vq_data[qid] = []
                vq_data[qid].append({
                    "users_wave_id": resp_id,
                    "value": value,
                })

    return vq_data


def _resolve_wave_ids(
    synapse: SynapseConfig,
    resolved_url: str,
    headers: dict,
) -> list[int]:
    """Resolve wave_ids from deliverable_ids via reporting plan.

    If synapse config already has wave_ids, use those directly.
    Otherwise, query the reporting plan to find waves associated with deliverables.
    """
    # Direct wave_ids if configured
    wave_ids = getattr(synapse, "wave_ids", None)
    if wave_ids:
        return wave_ids

    # Resolve from reporting plan deliverables
    reporting_plan_id = getattr(synapse, "reporting_plan_id", None)
    if not reporting_plan_id:
        logger.warning("No wave_ids or reporting_plan_id — cannot export VQ data")
        return []

    try:
        url = f"{resolved_url}/reporting-plans/{reporting_plan_id}"
        plan_detail = _get_json(url, headers)

        wave_ids = []
        for deliverable in plan_detail.get("deliverables", []):
            del_id = deliverable.get("id")
            if del_id in synapse.deliverable_ids:
                # Get waves from survey_deliverables
                for sd in deliverable.get("survey_deliverables", []):
                    for dsw in sd.get("waves", []):
                        wid = dsw.get("wave_id")
                        if wid and wid not in wave_ids:
                            wave_ids.append(wid)

        if wave_ids:
            print(f"  Resolved {len(wave_ids)} wave_ids from reporting plan {reporting_plan_id}")
        return wave_ids

    except Exception as e:
        logger.warning("Failed to resolve wave_ids from reporting plan: %s", e)
        return []


def _invalidate_raw_cache(config: ProjectConfig) -> None:
    """Delete source_raw_data.pkl cache so next run re-fetches."""
    pkl_path = _raw_pkl_path(config)
    if os.path.exists(pkl_path):
        os.remove(pkl_path)
        print(f"  Invalidated raw data cache: {pkl_path}")

    # Also clean up legacy JSON cache if present
    if config.context_path:
        json_path = os.path.join(config.context_path, "source_raw_data.json")
    else:
        json_path = os.path.join(
            os.path.dirname(config.raw_data_source_path or config.data_source_path),
            "source_raw_data.json",
        )
    if os.path.exists(json_path):
        os.remove(json_path)


# ── HTTP helpers ──────────────────────────────────────────────────────────

def _get_json(url: str, headers: dict) -> dict:
    """GET and return parsed JSON."""
    if _USE_REQUESTS:
        resp = _requests.get(url, headers=headers, timeout=30)
        if not resp.ok:
            detail = resp.text
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                pass
            raise RuntimeError(f"Synapse API error {resp.status_code}: {detail}")
        return resp.json()
    else:
        req = _urllib_req.Request(url, headers=headers)
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e


def _post_json(url: str, headers: dict, payload: dict) -> dict:
    """POST JSON and return parsed JSON response."""
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


def _post_ndjson(url: str, headers: dict, payload: dict) -> list[str]:
    """POST JSON and return NDJSON response as list of lines."""
    body = json.dumps(payload).encode("utf-8")
    req_headers = dict(headers)
    req_headers["Accept"] = "application/x-ndjson"

    if _USE_REQUESTS:
        resp = _requests.post(url, data=body, headers=req_headers, timeout=300)
        if not resp.ok:
            detail = resp.text[:500]
            try:
                detail = resp.json().get("detail", resp.text[:500])
            except Exception:
                pass
            raise RuntimeError(f"Synapse API error {resp.status_code}: {detail}")
        return resp.text.strip().split("\n")
    else:
        req = _urllib_req.Request(url, data=body, headers=req_headers, method="POST")
        try:
            with _urllib_req.urlopen(req, timeout=300) as r:
                text = r.read().decode("utf-8")
                return text.strip().split("\n")
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e


def _post_csv(url: str, headers: dict, payload: dict) -> str:
    """POST JSON and return CSV response as text."""
    body = json.dumps(payload).encode("utf-8")
    req_headers = dict(headers)
    req_headers["Accept"] = "text/csv"

    if _USE_REQUESTS:
        resp = _requests.post(url, data=body, headers=req_headers, timeout=120)
        if not resp.ok:
            detail = resp.text
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                pass
            raise RuntimeError(f"Synapse API error {resp.status_code}: {detail}")
        return resp.text
    else:
        req = _urllib_req.Request(url, data=body, headers=req_headers, method="POST")
        try:
            with _urllib_req.urlopen(req, timeout=120) as r:
                return r.read().decode("utf-8")
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e


# ── CLI entrypoint ────────────────────────────────────────────────────────

def _cli() -> None:
    """CLI handler for: python -m slidegen fetch-raw <yaml_path> [options]"""
    import argparse

    parser = argparse.ArgumentParser(
        prog="python -m slidegen fetch-raw",
        description="Fetch raw respondent data from the Synapse API (JSON → DataFrame → pkl).",
    )
    parser.add_argument("yaml_path", help="Path to project config YAML")
    parser.add_argument(
        "--api-key", default=None, metavar="KEY",
        help=f"Synapse Bearer token (default: ${_ENV_KEY_NAME} env var)",
    )
    parser.add_argument(
        "--url", default=None, dest="synapse_url", metavar="URL",
        help="Synapse API base URL override",
    )
    parser.add_argument(
        "--and-generate", action="store_true",
        help="Run generate_deck() after fetching (full pipeline)",
    )
    parser.add_argument(
        "--load-cached", action="store_true",
        help="Load from pkl cache instead of fetching fresh data",
    )
    args = parser.parse_args()

    from slidegen.pipeline.project_config import load_project_config
    config = load_project_config(args.yaml_path)

    if args.load_cached:
        cached = load_cached_pkl(config)
        if cached:
            dfs = cached.get("dataframes", {})
            for name, df in dfs.items():
                print(f"  {name}: {df.shape[0]} rows × {df.shape[1]} cols")
        else:
            print("No cached pkl found. Run without --load-cached to fetch.")
        return

    result = fetch_all_raw(config, api_key=args.api_key, synapse_url=args.synapse_url)

    print(f"\nDone. Pkl cache: {result['pkl_path']}")
    print(f"  DataFrames: {len(result['dataframes'])} sheets")
    for name, df in result["dataframes"].items():
        print(f"    {name}: {df.shape[0]} rows × {df.shape[1]} cols")
    print(f"  Segments: {len(result['segments'])}")
    print(f"  VQ questions: {len(result['vq_questions'])}")

    if args.and_generate:
        from slidegen.pipeline.orchestrator import generate_deck
        print("\nRunning generate_deck()...")
        generate_deck(args.yaml_path)
