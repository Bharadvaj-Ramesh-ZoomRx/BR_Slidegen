"""
synapse_client_factory.py — Create a SynapseClient from ProjectConfig.

Single integration point between SlideGen's config system and the
synapse-cli library.  All raw-data-first code imports this instead of
constructing clients directly.

Requires: pip install -e /path/to/synapse-cli
"""

from __future__ import annotations

import logging
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from slidegen.pipeline.project_config import ProjectConfig

logger = logging.getLogger(__name__)


def get_client(config: "ProjectConfig", api_key: Optional[str] = None,
               synapse_url: Optional[str] = None):
    """Create a SynapseClient from the project's ``synapse`` config section.

    Args:
        config: Loaded ProjectConfig with a ``synapse`` block.
        api_key: Explicit Bearer token (highest priority).  Falls back to
            ``SYNAPSE_API_KEY`` env → cached ``~/.synapse/token.json`` →
            Azure AD client credentials when *None*.
        synapse_url: Override base URL (falls back to ``config.synapse.api_url``).

    Returns:
        A ready-to-use ``SynapseClient`` instance.

    Raises:
        ImportError: If ``synapse-cli`` is not installed.
        ValueError: If no base URL or synapse config is available.
    """
    try:
        from synapse_cli import SynapseClient
    except ImportError:
        raise ImportError(
            "synapse-cli is not installed.  Run:\n"
            "  pip install -e /path/to/synapse-cli"
        )

    synapse = config.synapse
    if not synapse:
        raise ValueError("No 'synapse' section in config.yaml")

    base_url = (synapse_url or synapse.api_url or "").rstrip("/")
    if not base_url:
        raise ValueError(
            "synapse_url not provided and config has no synapse.api_url"
        )

    return SynapseClient(base_url=base_url, api_key=api_key)
