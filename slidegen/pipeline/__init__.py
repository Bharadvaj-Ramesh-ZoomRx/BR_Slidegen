"""
slidegen.pipeline — Generic slide deck generation pipeline.

Usage:
    from slidegen.pipeline import generate_deck
    generate_deck("projects/jnj_rybrevant/config.yaml")
"""

from slidegen.pipeline.orchestrator import generate_deck, regenerate_slide, refresh_deck
from slidegen.pipeline.data_loaders import index_excel
from slidegen.pipeline.synapse_fetcher import fetch_synapse_data, trigger_generation, wait_and_download
from slidegen.pipeline.synapse_json_loader import fetch_data_as_json
from slidegen.pipeline.synapse_raw_fetcher import fetch_all_raw, load_cached_pkl
from slidegen.pipeline.qual_data_loader import index_qualitative

__all__ = [
    "generate_deck", "regenerate_slide", "refresh_deck", "index_excel",
    "fetch_synapse_data", "trigger_generation", "wait_and_download",
    "fetch_data_as_json", "fetch_all_raw", "load_cached_pkl",
    "index_qualitative",
]
