"""
com.py — COM helpers for live editing via win32com.

All positions in inches. Requires PowerPoint to be running on Windows.
"""

from __future__ import annotations

from typing import Optional

from .brand import IN


def com_connect(target_filename: str):
    """Connect to a running PowerPoint instance and return the named presentation.

    Args:
        target_filename: bare filename, e.g. 'phase1_test.pptx' (not full path)
    Returns: win32com Presentation object
    Raises: RuntimeError if PowerPoint is not open or file not found
    """
    import win32com.client
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        raise RuntimeError(f"Could not connect to PowerPoint: {e}")

    for i in range(1, ppt_app.Presentations.Count + 1):
        p = ppt_app.Presentations(i)
        if target_filename in p.Name:
            return p

    open_files = [ppt_app.Presentations(i).Name
                  for i in range(1, ppt_app.Presentations.Count + 1)]
    raise RuntimeError(
        f"'{target_filename}' not found in open presentations.\n"
        f"  Open: {open_files}"
    )


def com_find_shape(com_slide, name: str):
    """Find a shape on a COM slide by its zrx_ name.

    Args:
        com_slide: win32com Slide object (1-indexed)
        name: shape name string, e.g. 'zrx_001'
    Returns: win32com Shape object
    Raises: RuntimeError with clear message if not found
    """
    for i in range(1, com_slide.Shapes.Count + 1):
        sh = com_slide.Shapes(i)
        if sh.Name == name:
            return sh
    raise RuntimeError(
        f"Shape '{name}' not found on slide. "
        f"Run 'python -m slidegen reconcile' and check for renames."
    )


def com_set_text(shape, text: str, color_bgr: Optional[int] = None,
                 size_pt: Optional[float] = None, bold: Optional[bool] = None) -> None:
    """Set text content and optional formatting on a COM shape.

    Args:
        shape: win32com Shape object with a TextFrame
        text: new text string
        color_bgr: int in BGR order (COM convention), e.g. 0x0000FF for red
        size_pt: font size in points
        bold: True/False
    """
    tr = shape.TextFrame.TextRange
    tr.Text = text
    if color_bgr is not None:
        tr.Font.Color.RGB = color_bgr
    if size_pt is not None:
        tr.Font.Size = size_pt
    if bold is not None:
        tr.Font.Bold = bold


def com_set_fill(shape, color_bgr: int) -> None:
    """Set fill colour on a COM shape (BGR int, e.g. 0x2458F7 for orange)."""
    shape.Fill.ForeColor.RGB = color_bgr


def com_move(shape, left_in: float, top_in: float) -> None:
    """Move a COM shape to a new position (inches)."""
    shape.Left = left_in * IN
    shape.Top  = top_in  * IN


def com_resize(shape, width_in: float, height_in: float) -> None:
    """Resize a COM shape (inches)."""
    shape.Width  = width_in  * IN
    shape.Height = height_in * IN


def com_get_position(shape) -> tuple[float, float, float, float]:
    """Return current (left, top, width, height) in inches from a COM shape."""
    return (
        round(shape.Left   / IN, 4),
        round(shape.Top    / IN, 4),
        round(shape.Width  / IN, 4),
        round(shape.Height / IN, 4),
    )
