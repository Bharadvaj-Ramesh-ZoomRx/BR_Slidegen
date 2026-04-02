"""
synapse_auth.py — Automatic Synapse API token acquisition and caching.

Resolves a Bearer token through this fallback chain:
  1. Explicit api_key= parameter
  2. SYNAPSE_API_KEY environment variable
  3. Azure AD client credentials flow (auto-acquires + caches until expiry)

Azure AD credentials are read from environment variables:
  AZURE_TENANT_ID     — Azure AD tenant ID
  AZURE_CLIENT_ID     — App registration client ID
  AZURE_CLIENT_SECRET — App registration client secret
  AZURE_SCOPE         — API scope (e.g. "api://<client_id>/.default")

For local dev (when the Synapse server has auth mocked), set:
  SYNAPSE_API_KEY=dev
and skip Azure config entirely.
"""

from __future__ import annotations

import json
import logging
import os
import time
from typing import Optional

logger = logging.getLogger(__name__)

# ── Auto-load .env on first import ────────────────────────────────────────
_dotenv_loaded = False

def _ensure_dotenv() -> None:
    """Load .env file once (walks up from cwd to find it)."""
    global _dotenv_loaded
    if _dotenv_loaded:
        return
    _dotenv_loaded = True
    try:
        from dotenv import load_dotenv, find_dotenv
        env_path = find_dotenv(usecwd=True)
        if env_path:
            load_dotenv(env_path, override=False)
            logger.debug("Loaded .env from %s", env_path)
    except ImportError:
        pass  # python-dotenv not installed — rely on shell env

# Prefer requests; fall back to urllib
try:
    import requests as _requests
    _USE_REQUESTS = True
except ImportError:
    import urllib.request as _urllib_req
    import urllib.error as _urllib_err
    _USE_REQUESTS = False

_ENV_KEY_NAME = "SYNAPSE_API_KEY"

# ── Token cache ────────────────────────────────────────────────────────────

_cached_token: Optional[str] = None
_cached_token_expiry: float = 0  # epoch seconds


def _is_cache_valid() -> bool:
    """Check if cached token exists and has >60s before expiry."""
    return bool(_cached_token) and time.time() < (_cached_token_expiry - 60)


# ── Public API ─────────────────────────────────────────────────────────────

def resolve_api_key(api_key: Optional[str] = None) -> str:
    """Resolve a Synapse Bearer token through the fallback chain.

    Args:
        api_key: Explicit token (highest priority).

    Returns:
        A valid Bearer token string.

    Raises:
        ValueError: If no token can be resolved.
    """
    _ensure_dotenv()

    # 1. Explicit parameter
    if api_key:
        return api_key

    # 2. Environment variable
    env_key = os.environ.get(_ENV_KEY_NAME, "").strip()
    if env_key:
        return env_key

    # 3. Azure AD client credentials flow
    token = _acquire_azure_token()
    if token:
        return token

    raise ValueError(
        f"No Synapse API token available. Options:\n"
        f"  1. Set {_ENV_KEY_NAME} environment variable\n"
        f"  2. Set AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET + AZURE_SCOPE\n"
        f"     for automatic Azure AD token acquisition\n"
        f"  3. Pass api_key= directly to the function"
    )


# ── Azure AD client credentials ───────────────────────────────────────────

def _acquire_azure_token() -> Optional[str]:
    """Acquire a token via Azure AD client credentials flow.

    Returns the access_token string, or None if Azure config is missing.
    """
    global _cached_token, _cached_token_expiry

    # Return cached token if still valid
    if _is_cache_valid():
        logger.debug("Using cached Azure AD token (expires in %.0fs)",
                      _cached_token_expiry - time.time())
        return _cached_token

    tenant_id = os.environ.get("AZURE_TENANT_ID", "").strip()
    client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()
    client_secret = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
    scope = os.environ.get("AZURE_SCOPE", "").strip()

    if not all([tenant_id, client_id, client_secret]):
        return None

    # Default scope: api://<client_id>/.default (standard for client credentials)
    if not scope:
        scope = f"api://{client_id}/.default"

    token_url = (
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    )
    form_data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": scope,
    }

    logger.debug("Acquiring Azure AD token (client credentials flow)...")

    try:
        data = _post_form(token_url, form_data)
    except (OSError, ValueError, KeyError) as exc:
        logger.warning("Azure AD token acquisition failed: %s", exc)
        return None

    access_token = data.get("access_token")
    if not access_token:
        logger.warning("Azure AD response missing access_token: %s",
                        list(data.keys()))
        return None

    # Cache with expiry (Azure tokens are typically valid for 1 hour)
    expires_in = int(data.get("expires_in", 3600))
    _cached_token = access_token
    _cached_token_expiry = time.time() + expires_in

    logger.debug("Azure AD token acquired (expires in %ds)", expires_in)
    return access_token


def _post_form(url: str, form_data: dict) -> dict:
    """POST application/x-www-form-urlencoded and return parsed JSON."""
    if _USE_REQUESTS:
        resp = _requests.post(url, data=form_data, timeout=30)
        if not resp.ok:
            raise RuntimeError(
                f"Token request failed ({resp.status_code}): {resp.text[:200]}"
            )
        return resp.json()
    else:
        import urllib.parse
        encoded = urllib.parse.urlencode(form_data).encode("utf-8")
        req = _urllib_req.Request(
            url,
            data=encoded,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            with _urllib_req.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except _urllib_err.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")[:200]
            raise RuntimeError(
                f"Token request failed ({e.code}): {body}"
            ) from e