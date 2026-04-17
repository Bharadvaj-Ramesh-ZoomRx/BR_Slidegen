"""
synapse_auth.py — Azure AD token acquisition for Synapse API.

Uses MSAL (Microsoft Authentication Library) with persistent token cache.
Authenticates once via browser, then silently refreshes forever.

Usage:
    from slidegen.synapse_auth import get_synapse_token

    token = get_synapse_token()
    headers = {"Authorization": f"Bearer {token}"}
    requests.get("https://synapse-api.zoomrx.com/api/projects", headers=headers)

First call opens a browser for Azure AD login (instant if already logged in).
Subsequent calls return cached/refreshed tokens silently — no user interaction.

Token cache: ~/.synapse_token_cache.json (persists across sessions)
"""
from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from typing import Optional

import msal

logger = logging.getLogger(__name__)

# Azure AD app registration for Synapse
CLIENT_ID = "62944583-a3ea-41dd-bd0e-ecc010809016"
TENANT_ID = "33e5c76b-0eec-495e-9f32-93fad67196d2"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [f"api://{CLIENT_ID}/synapse-dev-api.default"]

# Persistent token cache location
CACHE_PATH = Path.home() / ".synapse_token_cache.json"

# Synapse API base URL
SYNAPSE_API_BASE = "https://synapse-api.zoomrx.com"


def _load_cache() -> msal.SerializableTokenCache:
    """Load the persistent token cache from disk."""
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.exists():
        try:
            cache.deserialize(CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass  # corrupted cache — start fresh
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    """Save the token cache to disk (only if changed)."""
    if cache.has_state_changed:
        try:
            CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
        except Exception as e:
            logger.warning(f"Could not save token cache: {e}")


def get_synapse_token(force_interactive: bool = False) -> str:
    """Get a valid Synapse API access token.

    Resolution order:
    1. SYNAPSE_API_TOKEN env var (from .env — check if still valid)
    2. MSAL cache (silent refresh from prior login)
    3. Device code flow (prompts user once, then cached)

    Returns the access_token string (without "Bearer " prefix).
    """
    import time

    # 1. Try env var first
    env_token = os.environ.get("SYNAPSE_API_TOKEN", "")
    if env_token and not force_interactive:
        # Strip "Bearer " prefix if present
        raw = env_token.replace("Bearer ", "").strip()
        # Check if expired by decoding JWT
        try:
            import base64
            payload = raw.split(".")[1]
            payload += "=" * (4 - len(payload) % 4)
            claims = json.loads(base64.urlsafe_b64decode(payload))
            exp = claims.get("exp", 0)
            if exp > time.time() + 60:  # valid for at least 60 more seconds
                return raw
            else:
                logger.info("Env token expired, trying other methods...")
        except Exception:
            # Can't decode — try using it anyway
            return raw

    # 2. Try MSAL cache (silent refresh)
    cache = _load_cache()
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    accounts = app.get_accounts()
    if accounts and not force_interactive:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # 3. Device code flow (user enters code at microsoft.com/devicelogin)
    # NOTE: requires the Azure AD app to be configured as a public client,
    # or have a client_secret. If this fails with AADSTS7000218, ask the
    # Synapse team to either:
    #   a) Add http://localhost as a redirect URI (enables interactive flow)
    #   b) Share the client_secret (enables confidential client flow)
    #   c) Mark the app as "Allow public client flows" in Azure portal
    logger.info("No valid token available. Starting device code login...")
    print("\nNo valid Synapse token. Starting device code login...", file=sys.stderr)

    try:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow failed: {flow.get('error_description', flow)}")
        print(flow["message"], file=sys.stderr)
        result = app.acquire_token_by_device_flow(flow)
    except Exception as e:
        raise RuntimeError(
            f"Login failed: {e}\n\n"
            "To fix permanently, ask Synapse team to enable 'Allow public client flows' "
            "on app 62944583-a3ea-41dd-bd0e-ecc010809016 in Azure portal.\n"
            "Or update SYNAPSE_API_TOKEN in .env with a fresh token from your browser."
        )

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Azure AD login failed: {error}")

    _save_cache(cache)
    return result["access_token"]


def get_synapse_headers() -> dict:
    """Get HTTP headers with a valid Bearer token for Synapse API calls."""
    token = get_synapse_token()
    return {
        "Authorization": f"Bearer {token}",
        "accept": "application/json",
        "Content-Type": "application/json",
    }


def get_synapse_base_url() -> str:
    """Return the Synapse API base URL."""
    return SYNAPSE_API_BASE


def update_env_token(token: str, env_path: Optional[Path] = None) -> None:
    """Update SYNAPSE_API_TOKEN in .env file with a fresh token.

    Usage: python -m slidegen.synapse_auth --update "Bearer eyJ..."
    """
    if env_path is None:
        # Find .env in repo root
        env_path = Path(__file__).resolve().parents[1] / ".env"

    if not token.startswith("Bearer "):
        token = f"Bearer {token}"

    if env_path.exists():
        content = env_path.read_text(encoding="utf-8")
        import re
        if "SYNAPSE_API_TOKEN=" in content:
            content = re.sub(
                r"SYNAPSE_API_TOKEN=.*",
                f"SYNAPSE_API_TOKEN={token}",
                content,
            )
        else:
            content += f"\nSYNAPSE_API_TOKEN={token}\n"
        env_path.write_text(content, encoding="utf-8")
    else:
        env_path.write_text(f"SYNAPSE_API_TOKEN={token}\n", encoding="utf-8")

    print(f"Token updated in {env_path}")


# CLI: run this module directly to test auth or update token
if __name__ == "__main__":
    import requests

    if "--update" in sys.argv:
        # Update .env with a new token
        idx = sys.argv.index("--update")
        if idx + 1 < len(sys.argv):
            update_env_token(sys.argv[idx + 1])
        else:
            print("Usage: python -m slidegen.synapse_auth --update 'Bearer eyJ...'")
        sys.exit(0)

    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parents[1] / ".env")

    token = get_synapse_token(force_interactive="--login" in sys.argv)
    print(f"Token: {token[:30]}...")
    print(f"Cache: {CACHE_PATH}")

    # Quick test
    headers = {"Authorization": f"Bearer {token}", "accept": "application/json"}
    resp = requests.get(f"{SYNAPSE_API_BASE}/api/projects?sort_by=name&sort_order=asc",
                        headers=headers, timeout=15)
    if resp.status_code == 200:
        data = resp.json()
        projects = data.get("data", {}).get("assigned", [])
        print(f"Projects: {len(projects)}")
        for p in projects[:3]:
            print(f"  {p.get('id')}: {p.get('name')}")
    else:
        print(f"API test: {resp.status_code} {resp.text[:200]}")
