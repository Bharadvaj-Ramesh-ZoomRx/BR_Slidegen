"""
http_utils.py — Shared HTTP helpers with retry logic for Synapse API calls.

Consolidates the duplicated _get_json/_post_json pattern from synapse_fetcher,
synapse_json_loader, and synapse_raw_fetcher into a single module with
exponential backoff for transient failures (5xx, timeouts).
"""

from __future__ import annotations

import json
import logging
import time

logger = logging.getLogger(__name__)

import urllib.request as _urllib_req
import urllib.error as _urllib_err

try:
    import requests as _requests
    _USE_REQUESTS = True
except ImportError:
    _USE_REQUESTS = False

# Retry config
_MAX_RETRIES = 3
_BACKOFF_BASE = 1.0  # seconds: 1, 2, 4
_RETRYABLE_STATUS_CODES = {500, 502, 503, 504, 429}


def _is_retryable(status_code: int) -> bool:
    return status_code in _RETRYABLE_STATUS_CODES


def _extract_detail(resp) -> str:
    """Extract error detail from a requests Response object."""
    detail = resp.text
    try:
        detail = resp.json().get("detail", resp.text)
    except (ValueError, KeyError):
        pass
    return detail


def get_json(url: str, headers: dict, timeout: int = 30) -> dict:
    """GET a URL and return parsed JSON, with retry on transient errors."""
    last_err = None
    for attempt in range(_MAX_RETRIES):
        try:
            if _USE_REQUESTS:
                resp = _requests.get(url, headers=headers, timeout=timeout)
                if resp.ok:
                    return resp.json()
                if _is_retryable(resp.status_code) and attempt < _MAX_RETRIES - 1:
                    wait = _BACKOFF_BASE * (2 ** attempt)
                    logger.warning("GET %s returned %d — retry %d/%d in %.1fs",
                                   url, resp.status_code, attempt + 1, _MAX_RETRIES, wait)
                    time.sleep(wait)
                    continue
                raise RuntimeError(f"Synapse API error {resp.status_code}: {_extract_detail(resp)}")
            else:
                req = _urllib_req.Request(url, headers=headers)
                with _urllib_req.urlopen(req, timeout=timeout) as r:
                    return json.loads(r.read().decode("utf-8"))
        except (ConnectionError, TimeoutError, OSError) as e:
            last_err = e
            if attempt < _MAX_RETRIES - 1:
                wait = _BACKOFF_BASE * (2 ** attempt)
                logger.warning("GET %s failed (%s) — retry %d/%d in %.1fs",
                               url, e, attempt + 1, _MAX_RETRIES, wait)
                time.sleep(wait)
                continue
            raise RuntimeError(f"Synapse API request failed after {_MAX_RETRIES} retries: {e}") from e
        except _urllib_err.HTTPError as e:
            if _is_retryable(e.code) and attempt < _MAX_RETRIES - 1:
                wait = _BACKOFF_BASE * (2 ** attempt)
                logger.warning("GET %s returned %d — retry %d/%d in %.1fs",
                               url, e.code, attempt + 1, _MAX_RETRIES, wait)
                time.sleep(wait)
                continue
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e

    raise RuntimeError(f"Synapse API request failed after {_MAX_RETRIES} retries: {last_err}")


def post_json(url: str, headers: dict, payload: dict, timeout: int = 60) -> dict:
    """POST JSON and return parsed response, with retry on transient errors."""
    body = json.dumps(payload).encode("utf-8")
    last_err = None
    for attempt in range(_MAX_RETRIES):
        try:
            if _USE_REQUESTS:
                resp = _requests.post(url, data=body, headers=headers, timeout=timeout)
                if resp.ok:
                    return resp.json()
                if _is_retryable(resp.status_code) and attempt < _MAX_RETRIES - 1:
                    wait = _BACKOFF_BASE * (2 ** attempt)
                    logger.warning("POST %s returned %d — retry %d/%d in %.1fs",
                                   url, resp.status_code, attempt + 1, _MAX_RETRIES, wait)
                    time.sleep(wait)
                    continue
                raise RuntimeError(f"Synapse API error {resp.status_code}: {_extract_detail(resp)}")
            else:
                req = _urllib_req.Request(url, data=body, headers=headers, method="POST")
                with _urllib_req.urlopen(req, timeout=timeout) as r:
                    return json.loads(r.read().decode("utf-8"))
        except (ConnectionError, TimeoutError, OSError) as e:
            last_err = e
            if attempt < _MAX_RETRIES - 1:
                wait = _BACKOFF_BASE * (2 ** attempt)
                logger.warning("POST %s failed (%s) — retry %d/%d in %.1fs",
                               url, e, attempt + 1, _MAX_RETRIES, wait)
                time.sleep(wait)
                continue
            raise RuntimeError(f"Synapse API request failed after {_MAX_RETRIES} retries: {e}") from e
        except _urllib_err.HTTPError as e:
            if _is_retryable(e.code) and attempt < _MAX_RETRIES - 1:
                wait = _BACKOFF_BASE * (2 ** attempt)
                logger.warning("POST %s returned %d — retry %d/%d in %.1fs",
                               url, e.code, attempt + 1, _MAX_RETRIES, wait)
                time.sleep(wait)
                continue
            raise RuntimeError(f"Synapse API error {e.code}: {e.reason}") from e

    raise RuntimeError(f"Synapse API request failed after {_MAX_RETRIES} retries: {last_err}")
