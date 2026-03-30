"""
images.py — Image and logo placement utilities.

Centralized functions for placing images, logos, and icons on slides.
Per PRD §4.2.
"""

from __future__ import annotations

import os
from typing import Optional

from pptx.util import Inches


def insert_image(slide, img_path: str, left: float, top: float,
                 width: float, height: float, name: Optional[str] = None):
    """Place an external PNG or JPG at specified inch coordinates.

    Args:
        slide: python-pptx Slide object.
        img_path: Path to image file.
        left, top, width, height: Position and size in inches.
        name: Optional shape name for registry tracking.

    Returns:
        The picture shape, or None if img_path does not exist.
    """
    if not os.path.exists(img_path):
        return None
    pic = slide.shapes.add_picture(
        img_path,
        Inches(left), Inches(top), Inches(width), Inches(height))
    if name:
        pic.name = name
    return pic


def add_logo(slide, logo_path: str, position: str = "bottom_left",
             width: float = 0.8, margin: float = 0.2,
             slide_w: float = 13.333, slide_h: float = 7.5,
             name: Optional[str] = None):
    """Place a logo at a standard position on the slide.

    Args:
        slide: python-pptx Slide object.
        logo_path: Path to logo image file.
        position: One of "bottom_left", "bottom_right", "top_left", "top_right".
        width: Logo width in inches (height auto-calculated to maintain aspect).
        margin: Distance from slide edge in inches.
        slide_w, slide_h: Slide dimensions in inches.
        name: Optional shape name.

    Returns:
        The picture shape, or None if logo_path does not exist.
    """
    if not os.path.exists(logo_path):
        return None

    # Estimate height as square (actual aspect ratio maintained by python-pptx)
    height = width

    positions = {
        "bottom_left":  (margin, slide_h - height - margin),
        "bottom_right": (slide_w - width - margin, slide_h - height - margin),
        "top_left":     (margin, margin),
        "top_right":    (slide_w - width - margin, margin),
    }
    left, top = positions.get(position, positions["bottom_left"])
    return insert_image(slide, logo_path, left, top, width, height, name=name)
