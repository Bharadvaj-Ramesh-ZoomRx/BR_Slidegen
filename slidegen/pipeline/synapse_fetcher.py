"""
synapse_fetcher.py — Download banner plan data from the Synapse API.

Submits a banner plan generation request to the Synapse server, polls for
completion, downloads the resulting Excel file, backs up any existing
source_data.xlsx, saves the new file in its place, and invalidates the JSON
cache so the next pipeline run re-extracts from the fresh data.

Usage (CLI):
    python -m slidegen fetch-synapse projects/jnj_rybrevant/config.yaml

Usage (Python):
    from slidegen.pipeline import fetch_synapse_data
    from slidegen.pipeline.project_config import load_project_config
    config = load_project_config("projects/jnj_rybrevant/config.yaml")
    fetch_synapse_data(config)

Authentication:
    Set env var SYNAPSE_API_KEY or pass api_key= directly.
    The key is used as a Bearer token (Azure AD access token).
    Never store the key in config.yaml.
"""

import json
import os
import shutil
import time
from datetime import datetime
from typing import Optional

# Prefer requests (already in env) for cleaner error handling; fall back to stdlib.
try:
    import requests as _requests
    _USE_REQUESTS = True
except ImportError:
    import urllib.request as _urllib_req
    import urllib.error as _urllib_err
    _USE_REQUESTS = False

from slidegen.pipeline.project_config import ProjectConfig, SynapseConfig, load_project_config


# ── Constants ─────────────────────────────────────────────────────────────────

_DEFAULT_POLL_INTERVAL = 10   # seconds between status checks
_DEFAULT_MAX_WAIT = 300       # 5-minute hard timeout
_ENV_KEY_NAME = "SYNAPSE_API_KEY"

# Banner plan history statuses (mirrors BannerPlanHistoryStatus enum)
_STATUS_DONE = "processed"
_STATUS_FAILED = "failed"


# ── Public API ────────────────────────────────────────────────────────────────

def trigger_generation(
    config: ProjectConfig,
    api_key: Optional[str] = None,
    synapse_url: Optional[str] = None,
) -> int:
    """Submit a banner plan generation job (non-blocking).

    Returns the history_id immediately — call wait_and_download() later
    to poll for completion and download the Excel file.

    This lets other pipeline stages (context building, hypotheses) run
    while the banner plan generates in the background.

    Args:
        config:      Loaded ProjectConfig (must have synapse section).
        api_key:     Bearer token. Falls back to SYNAPSE_API_KEY env var.
        synapse_url: API base URL override.

    Returns:
        history_id (int) for the queued generation job.
    """
    synapse, resolved_url, headers = _resolve_synapse_params(config, api_key, synapse_url)

    print(f"Submitting banner plan generation for project_id={synapse.project_id}...")
    history_id = _trigger_generation(resolved_url, synapse, headers)
    print(f"  Generation queued: history_id={history_id}")
    return history_id


def wait_and_download(
    config: ProjectConfig,
    history_id: int,
    api_key: Optional[str] = None,
    synapse_url: Optional[str] = None,
    poll_interval: int = _DEFAULT_POLL_INTERVAL,
    max_wait: int = _DEFAULT_MAX_WAIT,
) -> str:
    """Poll for completion and download the banner plan Excel (blocking).

    Call this after trigger_generation() when you're ready to use the data.
    Handles polling, downloading, backup, and cache invalidation.

    Args:
        config:        Loaded ProjectConfig.
        history_id:    Job ID from trigger_generation().
        api_key:       Bearer token.
        synapse_url:   API base URL override.
        poll_interval: Seconds between polls. Default 10.
        max_wait:      Maximum wait seconds. Default 300.

    Returns:
        Absolute path to the saved source_data.xlsx.
    """
    _, resolved_url, headers = _resolve_synapse_params(config, api_key, synapse_url)

    # Poll for completion
    print(f"Polling for completion (max {max_wait}s, interval {poll_interval}s)...")
    _poll_until_done(resolved_url, headers, history_id, poll_interval, max_wait)
    print("  Banner plan generation complete.")

    # Download and save
    return _download_and_save(config, resolved_url, headers, history_id)


def fetch_synapse_data(
    config: ProjectConfig,
    api_key: Optional[str] = None,
    synapse_url: Optional[str] = None,
    poll_interval: int = _DEFAULT_POLL_INTERVAL,
    max_wait: int = _DEFAULT_MAX_WAIT,
) -> str:
    """Fetch fresh banner plan data from the Synapse API (blocking convenience wrapper).

    Combines trigger_generation() + wait_and_download() into a single call.
    Use the split functions when you want to run other work in parallel.

    Steps:
      1. POST /banner-plans/generate to submit the job
      2. Poll until done or timeout
      3. Download Excel, backup existing, invalidate JSON cache

    Returns:
        Absolute path to the saved source_data.xlsx.
    """
    history_id = trigger_generation(config, api_key, synapse_url)
    return wait_and_download(config, history_id, api_key, synapse_url,
                             poll_interval, max_wait)


# ── Internal helpers ──────────────────────────────────────────────────────────

def _resolve_synapse_params(
    config: ProjectConfig,
    api_key: Optional[str],
    synapse_url: Optional[str],
) -> tuple[SynapseConfig, str, dict]:
    """Validate and resolve synapse config, URL, and headers."""
    synapse = config.synapse
    if not synapse:
        raise ValueError(
            "No 'synapse' section in config.yaml. "
            "Add synapse.api_url, synapse.project_id, etc. to enable banner plan download."
        )
    resolved_key = _resolve_api_key(api_key)
    resolved_url = (synapse_url or synapse.api_url or getattr(config, "synapse_api_url", "")).rstrip("/")
    if not resolved_url:
        raise ValueError(
            "synapse_url not provided and config has no synapse.api_url field. "
            "Pass --url on the CLI or add synapse.api_url to config.yaml."
        )
    headers = _build_headers(resolved_key)
    return synapse, resolved_url, headers


def _download_and_save(
    config: ProjectConfig,
    resolved_url: str,
    headers: dict,
    history_id: int,
) -> str:
    """Download Excel, backup existing, save, and invalidate JSON cache."""
    excel_dest = config.data_source_path
    os.makedirs(os.path.dirname(excel_dest), exist_ok=True)

    _backup_existing_excel(excel_dest)

    tmp_path = excel_dest + ".tmp"
    try:
        print(f"  Downloading Excel → {os.path.basename(excel_dest)}...")
        _download_banner_plan(resolved_url, headers, history_id, tmp_path)
        if os.path.exists(excel_dest):
            os.remove(excel_dest)
        os.rename(tmp_path, excel_dest)
        print(f"  Saved: {excel_dest}")
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

    _invalidate_json_cache(config)
    return excel_dest


def _resolve_api_key(api_key: Optional[str]) -> str:
    from slidegen.pipeline.synapse_auth import resolve_api_key
    return resolve_api_key(api_key)


def _build_headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }


def _trigger_generation(base_url: str, synapse: SynapseConfig, headers: dict) -> int:
    """POST /banner-plans/generate and return the history_id."""
    url = f"{base_url}/banner-plans/generate"
    payload = {
        "project_id": synapse.project_id,
        "survey_ids": synapse.survey_ids,
        "deliverable_ids": synapse.deliverable_ids,
        "segment_ids": synapse.segment_ids,
        "multi_question_analysis_ids": synapse.multi_question_analysis_ids,
        "virtual_question_analysis_ids": synapse.virtual_question_analysis_ids,
    }
    body = json.dumps(payload).encode("utf-8")

    if _USE_REQUESTS:
        resp = _requests.post(url, data=body, headers=headers, timeout=30)
        if not resp.ok:
            detail = resp.text
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                pass
            raise RuntimeError(f"Synapse API error {resp.status_code}: {detail}")
        data = resp.json()
    else:
        req = _urllib_req.Request(url, data=body, headers=headers, method="POST")
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                data = json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e

    history_id = data.get("history_id")
    if not history_id:
        raise ValueError(f"Generate response missing 'history_id'. Got: {data}")
    return history_id


def _poll_until_done(
    base_url: str,
    headers: dict,
    history_id: int,
    poll_interval: int,
    max_wait: int,
) -> None:
    """Poll GET /banner-plans/histories/{id}/status until done or timeout."""
    url = f"{base_url}/banner-plans/histories/{history_id}/status"
    elapsed = 0
    last_status = "unknown"

    while elapsed < max_wait:
        response = _get_json(url, headers)
        last_status = response.get("status", "").lower()

        if last_status == _STATUS_DONE:
            return
        elif last_status == _STATUS_FAILED:
            raise RuntimeError(
                f"Banner plan generation failed (history_id={history_id})"
            )
        else:
            print(f"  [{elapsed}s] status={last_status!r} — waiting {poll_interval}s...")
            time.sleep(poll_interval)
            elapsed += poll_interval

    raise TimeoutError(
        f"Banner plan generation not complete after {max_wait}s "
        f"(history_id={history_id}, last status: {last_status!r})"
    )


def _download_banner_plan(
    base_url: str, headers: dict, history_id: int, dest_path: str
) -> None:
    """Stream GET /banner-plans/histories/{id}/download into dest_path."""
    url = f"{base_url}/banner-plans/histories/{history_id}/download"

    if _USE_REQUESTS:
        resp = _requests.get(url, headers=headers, stream=True, timeout=120)
        if not resp.ok:
            detail = resp.text
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                pass
            raise RuntimeError(f"Download error {resp.status_code}: {detail}")
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=131072):  # 128 KB chunks
                f.write(chunk)
    else:
        req = _urllib_req.Request(url, headers=headers)
        try:
            with _urllib_req.urlopen(req, timeout=120) as r:
                with open(dest_path, "wb") as f:
                    while True:
                        chunk = r.read(131072)
                        if not chunk:
                            break
                        f.write(chunk)
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Download error {e.code}: {e.reason}") from e


def _get_json(url: str, headers: dict) -> dict:
    """GET a URL and return parsed JSON."""
    if _USE_REQUESTS:
        resp = _requests.get(url, headers=headers, timeout=30)
        if not resp.ok:
            detail = resp.text
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                pass
            raise RuntimeError(f"Status check error {resp.status_code}: {detail}")
        return resp.json()
    else:
        req = _urllib_req.Request(url, headers=headers)
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Status check error {e.code}: {e.reason}") from e


def _backup_existing_excel(excel_path: str) -> Optional[str]:
    """Copy excel_path to source_data_backup_{ts}.xlsx in the same directory.

    Returns the backup path, or None if no file existed.
    """
    if not os.path.exists(excel_path):
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"source_data_backup_{ts}.xlsx"
    backup_path = os.path.join(os.path.dirname(excel_path), backup_name)
    shutil.copy2(excel_path, backup_path)
    print(f"  Backed up existing Excel → {backup_name}")
    return backup_path


def _invalidate_json_cache(config: ProjectConfig) -> None:
    """Delete source_data.json so the next pipeline run re-extracts from Excel."""
    if config.context_path:
        json_path = os.path.join(config.context_path, "source_data.json")
    else:
        json_path = os.path.join(os.path.dirname(config.data_source_path), "source_data.json")

    if os.path.exists(json_path):
        os.remove(json_path)
        print(f"  Invalidated JSON cache: {json_path}")
    else:
        print("  No JSON cache to invalidate.")


# ── CLI entrypoint ────────────────────────────────────────────────────────────

def _cli() -> None:
    """CLI handler for: python -m slidegen fetch-synapse <yaml_path> [options]"""
    import argparse

    parser = argparse.ArgumentParser(
        prog="python -m slidegen fetch-synapse",
        description="Download banner plan data from the Synapse API.",
    )
    parser.add_argument("yaml_path", help="Path to project config YAML")
    parser.add_argument(
        "--api-key",
        default=None,
        metavar="KEY",
        help=f"Synapse Bearer token (default: ${_ENV_KEY_NAME} env var)",
    )
    parser.add_argument(
        "--url",
        default=None,
        dest="synapse_url",
        metavar="URL",
        help="Synapse API base URL override (default: config.synapse.api_url)",
    )
    parser.add_argument(
        "--poll-interval",
        default=_DEFAULT_POLL_INTERVAL,
        type=int,
        metavar="N",
        help=f"Seconds between status polls (default: {_DEFAULT_POLL_INTERVAL})",
    )
    parser.add_argument(
        "--max-wait",
        default=_DEFAULT_MAX_WAIT,
        type=int,
        metavar="N",
        help=f"Max total seconds to wait (default: {_DEFAULT_MAX_WAIT})",
    )
    parser.add_argument(
        "--and-generate",
        action="store_true",
        help="Run generate_deck() after fetching (full pipeline)",
    )
    args = parser.parse_args()

    config = load_project_config(args.yaml_path)
    excel_path = fetch_synapse_data(
        config,
        api_key=args.api_key,
        synapse_url=args.synapse_url,
        poll_interval=args.poll_interval,
        max_wait=args.max_wait,
    )
    print(f"\nDone. Excel saved: {excel_path}")

    if args.and_generate:
        from slidegen.pipeline.orchestrator import generate_deck
        print("\nRunning generate_deck()...")
        generate_deck(args.yaml_path)
