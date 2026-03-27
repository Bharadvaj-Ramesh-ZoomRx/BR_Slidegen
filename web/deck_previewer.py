"""
deck_previewer.py — Render PPTX slides as PNG images.

Flow:  PPTX → PDF (LibreOffice/soffice) → PNG (pypdfium2)

Falls back to Pillow placeholder images if soffice is not available.
Font rendering note: J&J custom fonts (Johnson Display/Text) will be
substituted in preview — the real deck requires PowerPoint.
"""

from __future__ import annotations

import io
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# pypdfium2 is optional — fall back gracefully if missing
try:
    import pypdfium2
    _PYPDFIUM2_AVAILABLE = True
except ImportError:
    _PYPDFIUM2_AVAILABLE = False


# ── Constants ─────────────────────────────────────────────────────────────────

THUMBNAIL_DPI = 150
HIRES_DPI = 300
SOFFICE_TIMEOUT = 90  # seconds

# Slide aspect ratio (widescreen 13.33 x 7.5 in)
SLIDE_ASPECT = 7.5 / 13.33


# ── Public API ────────────────────────────────────────────────────────────────

def render_thumbnails(pptx_path: str | Path, dpi: int = THUMBNAIL_DPI) -> list[bytes]:
    """
    Render all slides in a PPTX as PNG bytes.

    Returns list of PNG bytes, one per slide, in presentation order.
    Falls back to placeholder images if soffice/pypdfium2 is unavailable.
    """
    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        return []

    slide_count = _count_slides(pptx_path)

    try:
        return _render_via_soffice(pptx_path, dpi, slide_count)
    except Exception as e:
        # Graceful fallback — show numbered placeholders
        return _make_placeholders(slide_count, note=str(e)[:120])


def render_slide_hires(
    pptx_path: str | Path,
    slide_index: int,
    dpi: int = HIRES_DPI,
) -> bytes:
    """
    Render a single slide at high resolution.

    slide_index is 0-based.
    Returns PNG bytes. Falls back to placeholder on error.
    """
    pptx_path = Path(pptx_path)
    try:
        all_slides = _render_via_soffice(pptx_path, dpi, None)
        if 0 <= slide_index < len(all_slides):
            return all_slides[slide_index]
        return _make_placeholder(slide_index + 1, "Slide index out of range")
    except Exception as e:
        return _make_placeholder(slide_index + 1, str(e)[:120])


def soffice_available() -> bool:
    """Check if LibreOffice soffice is on PATH."""
    try:
        result = subprocess.run(
            ["soffice", "--version"],
            capture_output=True,
            timeout=10,
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


# ── Internal rendering ────────────────────────────────────────────────────────

def _render_via_soffice(
    pptx_path: Path,
    dpi: int,
    expected_count: int | None,
) -> list[bytes]:
    """Convert PPTX → PDF via soffice, then render pages via pypdfium2."""
    if not _PYPDFIUM2_AVAILABLE:
        raise RuntimeError("pypdfium2 not installed")

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        pdf_path = _soffice_convert(pptx_path, tmp)
        return _pdf_to_pngs(pdf_path, dpi)


def _soffice_convert(pptx_path: Path, out_dir: Path) -> Path:
    """Run soffice --headless --convert-to pdf."""
    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
        timeout=SOFFICE_TIMEOUT,
    )
    pdf_path = out_dir / f"{pptx_path.stem}.pdf"
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            f"soffice conversion failed (rc={result.returncode}): {result.stderr[:200]}"
        )
    return pdf_path


def _pdf_to_pngs(pdf_path: Path, dpi: int) -> list[bytes]:
    """Render each PDF page to PNG bytes using pypdfium2."""
    doc = pypdfium2.PdfDocument(str(pdf_path))
    scale = dpi / 72.0
    images = []
    try:
        for i in range(len(doc)):
            page = doc[i]
            bitmap = page.render(scale=scale, rotation=0)
            pil_img = bitmap.to_pil()
            buf = io.BytesIO()
            pil_img.save(buf, format="PNG", optimize=True)
            images.append(buf.getvalue())
    finally:
        doc.close()
    return images


# ── Slide counting ────────────────────────────────────────────────────────────

def _count_slides(pptx_path: Path) -> int:
    """Count slides in PPTX without loading the whole file."""
    try:
        import zipfile
        with zipfile.ZipFile(pptx_path) as z:
            names = z.namelist()
        return sum(1 for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
    except Exception:
        return 0


# ── Placeholder fallback ──────────────────────────────────────────────────────

def _make_placeholders(count: int, note: str = "") -> list[bytes]:
    """Generate placeholder PNG images for all slides."""
    return [_make_placeholder(i + 1, note if i == 0 else "") for i in range(count)]


def _make_placeholder(slide_num: int, note: str = "") -> bytes:
    """Create a simple grey placeholder PNG for a slide."""
    width, height = 800, 450  # 16:9 proportions
    img = Image.new("RGB", (width, height), color="#E8E8E8")
    draw = ImageDraw.Draw(img)

    # Border
    draw.rectangle([2, 2, width - 3, height - 3], outline="#BBBBBB", width=2)

    # Slide number
    try:
        font_large = ImageFont.load_default(size=48)
        font_small = ImageFont.load_default(size=18)
    except TypeError:
        font_large = ImageFont.load_default()
        font_small = ImageFont.load_default()

    label = f"Slide {slide_num}"
    bbox = draw.textbbox((0, 0), label, font=font_large)
    tw = bbox[2] - bbox[0]
    draw.text(((width - tw) // 2, height // 2 - 40), label, fill="#555555", font=font_large)

    if note:
        note_lines = _wrap_text(note, 80)
        y = height // 2 + 30
        for line in note_lines[:3]:
            bbox = draw.textbbox((0, 0), line, font=font_small)
            tw = bbox[2] - bbox[0]
            draw.text(((width - tw) // 2, y), line, fill="#999999", font=font_small)
            y += 24

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _wrap_text(text: str, max_chars: int) -> list[str]:
    """Simple word-wrap."""
    words = text.split()
    lines, current = [], ""
    for word in words:
        if len(current) + len(word) + 1 <= max_chars:
            current = (current + " " + word).strip()
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines
