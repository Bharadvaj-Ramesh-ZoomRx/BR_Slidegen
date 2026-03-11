"""
slidegen.pipeline — Generic slide deck generation pipeline.

Usage:
    from slidegen.pipeline import generate_deck
    generate_deck("projects/jnj_rybrevant/config.yaml")
"""

from slidegen.pipeline.orchestrator import generate_deck, regenerate_slide

__all__ = ["generate_deck", "regenerate_slide"]
