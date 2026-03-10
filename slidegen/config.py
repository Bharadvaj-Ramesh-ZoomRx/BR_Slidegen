"""
Centralized configuration for the SlideGen system.

All path resolution, defaults, and environment-specific settings live here.
Scripts import from this module instead of computing paths themselves.
"""

import os

# ── Paths ────────────────────────────────────────────────────────────────────
SLIDEGEN_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SLIDEGEN_DIR)
OUTPUT_DIR = os.path.join(SLIDEGEN_DIR, "output")
REGISTRY_PATH = os.path.join(SLIDEGEN_DIR, "slide_registry.json")
EDIT_LOG_PATH = os.path.join(SLIDEGEN_DIR, "edit_log.json")

# ── Client template (override per-project in CLAUDE.md or env var) ───────────
TEMPLATE_PATH = os.environ.get(
    "SLIDEGEN_TEMPLATE",
    os.path.join(PROJECT_ROOT, "J_and_J_project", "templates",
                 "JJ PET RYBREVANT+LAZCLUZE Q4'25 Report 1.pptx"),
)

# ── Slide dimensions (widescreen 13.333" x 7.5") ────────────────────────────
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.500
SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000

# ── Shape naming ─────────────────────────────────────────────────────────────
SHAPE_PREFIX = "zrx_"

# ── COM constants ────────────────────────────────────────────────────────────
PTS_PER_INCH = 72  # 1 inch = 72 points (COM coordinate system)


def ensure_output_dir():
    """Create the output directory if it doesn't exist."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
