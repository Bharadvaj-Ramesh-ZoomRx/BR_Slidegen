"""
slidegen.pipeline — Generic slide deck generation pipeline.

Usage:
    from slidegen.pipeline import generate_deck
    generate_deck("projects/jnj_rybrevant/config.yaml")
"""

from slidegen.pipeline.orchestrator import generate_deck

__all__ = ["generate_deck"]
