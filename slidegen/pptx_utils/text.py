"""
text.py — Text formatting helpers for PowerPoint generation.

Provides run-level formatting, paragraph alignment, and text utilities
that complement the shape primitives in shapes.py. Per PRD §4.2.
"""

from __future__ import annotations

from typing import Optional

from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


def format_run(run, font_name: str = "", font_size: float = 0,
               bold: Optional[bool] = None, italic: Optional[bool] = None,
               color: Optional[RGBColor] = None, underline: Optional[bool] = None):
    """Apply formatting to a single text run.

    Only sets properties that are explicitly provided (non-default).
    """
    font = run.font
    if font_name:
        font.name = font_name
    if font_size:
        font.size = Pt(font_size)
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if color is not None:
        font.color.rgb = color
    if underline is not None:
        font.underline = underline


def add_run(paragraph, text: str, font_name: str = "", font_size: float = 0,
            bold: Optional[bool] = None, italic: Optional[bool] = None,
            color: Optional[RGBColor] = None) -> object:
    """Add a formatted text run to an existing paragraph.

    Returns the created run object.
    """
    run = paragraph.add_run()
    run.text = text
    format_run(run, font_name=font_name, font_size=font_size,
               bold=bold, italic=italic, color=color)
    return run


def set_paragraph_alignment(paragraph, alignment: str = "left"):
    """Set paragraph alignment from a string value.

    Args:
        alignment: "left", "center", "right", or "justify"
    """
    align_map = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }
    paragraph.alignment = align_map.get(alignment, PP_ALIGN.LEFT)


def truncate_text(text: str, max_chars: int = 50, suffix: str = "...") -> str:
    """Truncate text to max_chars, adding suffix if truncated."""
    if len(text) <= max_chars:
        return text
    return text[:max_chars - len(suffix)] + suffix


def wrap_label(text: str, max_line_chars: int = 30) -> str:
    """Wrap a label into multiple lines for chart/table display.

    Splits at word boundaries, respecting max line length.
    """
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        if current_line and len(current_line) + 1 + len(word) > max_line_chars:
            lines.append(current_line)
            current_line = word
        else:
            current_line = f"{current_line} {word}" if current_line else word
    if current_line:
        lines.append(current_line)
    return "\n".join(lines)


def delta_format(value: float, suffix: str = " pp") -> tuple[str, str]:
    """Format a delta value with sign and determine color category.

    Args:
        value: Delta in percentage points.
        suffix: Unit suffix (default: " pp" for percentage points).

    Returns:
        (formatted_text, color_key) where color_key is "positive", "negative", or "neutral".
    """
    if abs(value) < 0.05:
        return f"—{suffix}", "neutral"
    sign = "+" if value > 0 else ""
    text = f"{sign}{value:.1f}{suffix}"
    color_key = "positive" if value > 0 else "negative"
    return text, color_key
