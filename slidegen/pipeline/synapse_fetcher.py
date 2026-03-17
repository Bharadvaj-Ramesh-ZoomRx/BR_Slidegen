"""
synapse_fetcher.py — Fetch survey data from the Synapse async API.

Submits a job to the Synapse portal API, polls for completion, downloads the
resulting Excel from S3, backs up any existing source_data.xlsx, saves the new
file in its place, and invalidates the JSON cache so the next pipeline run
re-extracts from the fresh data.

Usage (CLI):
    python -m slidegen fetch-synapse projects/jnj_rybrevant/config.yaml \\
        --url https://api.synapse.example.com/v1

Usage (Python):
    from slidegen.pipeline import fetch_synapse_data
    from slidegen.pipeline.project_config import load_project_config
    config = load_project_config("projects/jnj_rybrevant/config.yaml")
    fetch_synapse_data(config, synapse_url="https://api.synapse.example.com/v1")

Authentication:
    Set env var SYNAPSE_API_KEY or pass api_key= directly.
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

from slidegen.pipeline.project_config import ProjectConfig, load_project_config


# ── Constants ─────────────────────────────────────────────────────────────────

_DEFAULT_POLL_INTERVAL = 10   # seconds between status checks
_DEFAULT_MAX_WAIT = 300       # 5-minute hard timeout
_ENV_KEY_NAME = "SYNAPSE_API_KEY"

# Terminal status sets — handle variant spellings from different API versions
_STATUS_DONE_SET   = {"done", "complete", "completed", "success"}
_STATUS_FAILED_SET = {"failed", "error", "cancelled", "canceled"}


# ── Public entry point ────────────────────────────────────────────────────────

def fetch_synapse_data(
    config: ProjectConfig,
    api_key: Optional[str] = None,
    synapse_url: Optional[str] = None,
    poll_interval: int = _DEFAULT_POLL_INTERVAL,
    max_wait: int = _DEFAULT_MAX_WAIT,
) -> str:
    """Fetch fresh survey data from the Synapse async API.

    Steps:
      1. Resolve API key (param → SYNAPSE_API_KEY env var)
      2. Resolve API base URL (param → config.synapse_api_url)
      3. POST to {url}/jobs to submit the job; get back a job_id
      4. Poll {url}/jobs/{job_id} until done or timeout
      5. Download Excel from the S3 pre-signed URL in the response
      6. Backup existing source_data.xlsx (timestamped rename, same dir)
      7. Atomically replace source_data.xlsx with the new file
      8. Invalidate source_data.json cache so next run re-extracts

    Args:
        config:         Loaded ProjectConfig instance.
        api_key:        Bearer token. Falls back to SYNAPSE_API_KEY env var.
        synapse_url:    API base URL. Falls back to config.synapse_api_url.
        poll_interval:  Seconds between status polls. Default 10.
        max_wait:       Maximum total seconds to wait. Default 300.

    Returns:
        Absolute path to the saved source_data.xlsx.

    Raises:
        ValueError:    Missing API key, URL, or unexpected API response format.
        TimeoutError:  Job not complete within max_wait seconds.
        RuntimeError:  Job failed or API returned an error status.
        OSError:       File I/O errors during backup or save.
    """
    # 1. Resolve credentials and URL
    resolved_key = _resolve_api_key(api_key)
    resolved_url = (synapse_url or getattr(config, "synapse_api_url", "")).rstrip("/")
    if not resolved_url:
        raise ValueError(
            "synapse_url not provided and config has no synapse_api_url field. "
            "Pass --url on the CLI or add synapse_api_url to config.yaml."
        )

    headers = _build_headers(resolved_key)
    payload = _build_job_payload(config)

    # 2. Submit job
    print(f"Submitting Synapse job for '{config.name}' / wave='{config.wave}'...")
    job_response = _post_job(resolved_url, payload, headers)
    job_id = job_response.get("job_id")
    if not job_id:
        raise ValueError(
            f"Synapse API response missing 'job_id'. Got: {job_response}"
        )
    print(f"  Job submitted: {job_id}")

    # 3. Determine status URL (API may return an explicit URL or we construct it)
    status_url = job_response.get("status_url") or f"{resolved_url}/jobs"

    # 4. Poll for completion
    print(f"Polling for completion (max {max_wait}s, interval {poll_interval}s)...")
    result = _poll_status(status_url, headers, job_id, poll_interval, max_wait)

    s3_url = (
        result.get("result_url")
        or result.get("download_url")
        or result.get("url")
    )
    if not s3_url:
        raise ValueError(
            f"Completed job response missing download URL. Got: {result}"
        )
    print("  Job complete. Download URL received.")

    # 5. Ensure destination directory exists
    excel_dest = config.data_source_path
    os.makedirs(os.path.dirname(excel_dest), exist_ok=True)

    # 6. Backup existing Excel
    _backup_existing_excel(excel_dest)

    # 7. Download to .tmp then rename atomically (never leaves a partial file)
    tmp_path = excel_dest + ".tmp"
    try:
        print(f"  Downloading Excel → {os.path.basename(excel_dest)}...")
        _download_excel(s3_url, tmp_path)
        if os.path.exists(excel_dest):
            os.remove(excel_dest)
        os.rename(tmp_path, excel_dest)
        print(f"  Saved: {excel_dest}")
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

    # 8. Invalidate JSON cache
    _invalidate_json_cache(config)

    return excel_dest


# ── Internal helpers ──────────────────────────────────────────────────────────

def _resolve_api_key(api_key: Optional[str]) -> str:
    if api_key:
        return api_key
    key = os.environ.get(_ENV_KEY_NAME, "")
    if key:
        return key
    raise ValueError(
        f"API key not provided. Pass --api-key on the CLI or "
        f"set the {_ENV_KEY_NAME} environment variable."
    )


def _build_headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }


def _build_job_payload(config: ProjectConfig) -> dict:
    return {
        "project": config.name,
        "wave": config.wave,
        "period_current": config.period_current,
        "period_prior": config.period_prior,
        "client": config.client,
    }


def _post_job(base_url: str, payload: dict, headers: dict) -> dict:
    """POST to {base_url}/jobs and return parsed JSON response."""
    url = f"{base_url}/jobs"
    body = json.dumps(payload).encode("utf-8")

    if _USE_REQUESTS:
        resp = _requests.post(url, data=body, headers=headers, timeout=30)
        resp.raise_for_status()
        return resp.json()
    else:
        req = _urllib_req.Request(url, data=body, headers=headers, method="POST")
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e


def _poll_status(
    status_url: str,
    headers: dict,
    job_id: str,
    poll_interval: int,
    max_wait: int,
) -> dict:
    """Poll {status_url}/{job_id} until the job is done or timeout."""
    elapsed = 0
    last_status = "unknown"

    while elapsed < max_wait:
        response = _get_json(f"{status_url}/{job_id}", headers)
        last_status = response.get("status", "").lower()

        if last_status in _STATUS_DONE_SET:
            return response
        elif last_status in _STATUS_FAILED_SET:
            msg = response.get("message") or response.get("error") or last_status
            raise RuntimeError(f"Synapse job '{job_id}' failed: {msg}")
        else:
            # Includes pending, running, queued, processing, and unknown statuses.
            # Unknown statuses are treated as "still running" — safe default if
            # the API introduces new transient status strings in the future.
            print(f"  [{elapsed}s] status={last_status!r} — waiting {poll_interval}s...")
            time.sleep(poll_interval)
            elapsed += poll_interval

    raise TimeoutError(
        f"Synapse job '{job_id}' not complete after {max_wait}s "
        f"(last status: {last_status!r})"
    )


def _get_json(url: str, headers: dict) -> dict:
    """GET a URL and return parsed JSON."""
    if _USE_REQUESTS:
        resp = _requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        return resp.json()
    else:
        req = _urllib_req.Request(url, headers=headers)
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            raise RuntimeError(f"Status check error {e.code}: {e.reason}") from e


def _download_excel(s3_url: str, dest_path: str) -> None:
    """Stream-download from a pre-signed S3 URL into dest_path."""
    if _USE_REQUESTS:
        resp = _requests.get(s3_url, stream=True, timeout=120)
        resp.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=131072):  # 128 KB chunks
                f.write(chunk)
    else:
        _urllib_req.urlretrieve(s3_url, dest_path)


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
    # Mirror the fallback logic in data_loaders._json_path_for()
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
        description="Fetch fresh survey data from the Synapse async API.",
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
        help="Synapse API base URL (default: config.synapse_api_url)",
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
