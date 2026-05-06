"""Sync wave-period labels in text frames adjacent to refreshed charts.

After a refresh, chart category axes reflect the new wave window
(Nov'25-Apr'26) but the period banners next to each chart can stay frozen
on the source's old window (Oct'25-Mar'26). This module compares source
chart cats to refreshed chart cats per shape, builds a `{old_wave: new_wave}`
shift map, and rewrites only:

  - text frames (or table cells) whose bbox is immediately above or below a
    chart shape — vertically within 0.4" of the chart's top or bottom edge,
    horizontally overlapping the chart by >= 30%.
  - run-level period tokens (MMM'YY, Q[1-4]'YY, Q[1-4] YYYY, Wave N, W N).
    Only the wave token portion is touched — segment suffixes like
    "Nov'25 - Gastro" never appear outside the chart object.

Title bars at the top of slides, footnotes at the bottom, headlines, and
unrelated text are untouched. Conflicting per-chart shifts on the same
slide drop the conflicting key (other consistent keys still apply).
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

from pptx import Presentation

from slidegen.synapse_chart_mapper import _extract_wave_token

# Adjacency tuning — text frame must be within this vertical band of the
# chart's top OR bottom edge, and horizontally cover at least this fraction
# of the smaller of (text frame width, chart width).
ADJACENCY_VERTICAL_TOLERANCE_EMU = 365_760  # 0.4 inch
ADJACENCY_HORIZONTAL_OVERLAP_RATIO = 0.30

# Quote variants that may appear in chart cats vs slide text.
_QUOTE_NORM = {"‘": "'", "’": "'"}


def _norm_quotes(s: str) -> str:
    if not s:
        return s
    for k, v in _QUOTE_NORM.items():
        s = s.replace(k, v)
    return s


def _bbox(shape) -> tuple[int, int, int, int]:
    """Return (left, top, right, bottom) in EMU; (-1,-1,-1,-1) if not positioned."""
    if shape.left is None or shape.top is None:
        return (-1, -1, -1, -1)
    return (shape.left, shape.top,
            shape.left + (shape.width or 0),
            shape.top + (shape.height or 0))


def _is_adjacent(text_bbox, chart_bbox) -> bool:
    """True if text_bbox sits immediately above or below chart_bbox."""
    tl, tt, tr, tb = text_bbox
    cl, ct, cr, cb = chart_bbox
    if tl < 0 or cl < 0:
        return False
    h_overlap = max(0, min(tr, cr) - max(tl, cl))
    min_width = min(tr - tl, cr - cl)
    if min_width <= 0:
        return False
    if h_overlap / min_width < ADJACENCY_HORIZONTAL_OVERLAP_RATIO:
        return False
    text_center_y = (tt + tb) // 2
    above_dist = ct - text_center_y
    below_dist = text_center_y - cb
    if 0 <= above_dist <= ADJACENCY_VERTICAL_TOLERANCE_EMU:
        return True
    if 0 <= below_dist <= ADJACENCY_VERTICAL_TOLERANCE_EMU:
        return True
    return False


def _chart_categories(shape) -> list[str]:
    """Read category strings from a chart shape's first plot."""
    try:
        return [str(c) for c in shape.chart.plots[0].categories]
    except Exception:
        return []


def _build_shift_map(old_cats: list[str], new_cats: list[str]) -> dict[str, str]:
    """Per-chart {old_wave: new_wave}, positionally zipped, quote-normalized.

    Only emits an entry where both sides parse to a wave token AND the tokens
    differ. If either side is non-temporal or lengths don't match, returns {}.
    """
    if not old_cats or not new_cats or len(old_cats) != len(new_cats):
        return {}
    shift: dict[str, str] = {}
    for old, new in zip(old_cats, new_cats):
        old_tok = _extract_wave_token(_norm_quotes(old))
        new_tok = _extract_wave_token(_norm_quotes(new))
        if not (old_tok and new_tok):
            return {}  # any non-temporal cat → not a wave-axis chart
        if old_tok != new_tok:
            shift[_norm_quotes(old_tok)] = _norm_quotes(new_tok)
    return shift


def _merge_slide_shift(
    per_chart: list[dict[str, str]],
) -> tuple[dict[str, str], list[str]]:
    """Union per-chart maps. Drop keys with conflicting values.

    Returns (merged, conflicts). When charts on a slide have inconsistent
    shifts (e.g. mixed cadences), the conflicting keys are returned so the
    caller can log/skip them; consistent keys still ship.
    """
    accumulated: dict[str, set[str]] = defaultdict(set)
    for m in per_chart:
        for k, v in m.items():
            accumulated[k].add(v)
    merged: dict[str, str] = {}
    conflicts: list[str] = []
    for k, vs in accumulated.items():
        if len(vs) == 1:
            merged[k] = next(iter(vs))
        else:
            conflicts.append(k)
    return merged, conflicts


# Label-run guard — keeps the rewriter out of full sentences. A run is
# considered a "label" iff:
#   - stripped length <= LABEL_MAX_CHARS
#   - after removing every wave-shaped token, the residue is only whitespace,
#     punctuation, and chart-connector words (vs / and / to / through)
LABEL_MAX_CHARS = 30
_WAVE_LIKE_BROAD = re.compile(
    r"[A-Za-z]{3}['’‘]\d{2}\s*[-–—]\s*[A-Za-z]{3}['’‘]\d{2}|"  # rolling
    r"[A-Za-z]{3}['’‘]\d{2}|"
    r"Q[1-4]['’‘]\d{2}|Q[1-4]\s+\d{4}|"
    r"(?:Project\s+)?Wave\s+\d+|W\d+",
    re.IGNORECASE,
)
_LABEL_RESIDUE = re.compile(
    r"^[\s\-–—|,&+:.()\/]*(?:vs|and|to|through)?"
    r"[\s\-–—|,&+:.()\/]*(?:vs|and|to|through)?[\s\-–—|,&+:.()\/]*$",
    re.IGNORECASE,
)


def _is_label_run(text: str) -> bool:
    """True iff `text` reads like a short period-banner label.

    Excludes long sentences and notes. Empty / whitespace-only runs return
    False (no wave tokens to replace anyway).
    """
    s = _norm_quotes(text or "").strip()
    if not s or len(s) > LABEL_MAX_CHARS:
        return False
    residue = _WAVE_LIKE_BROAD.sub(" ", s).strip()
    if not residue:
        return True
    return bool(_LABEL_RESIDUE.fullmatch(residue))


def _apply_shift_to_text(text: str, shift: dict[str, str]) -> tuple[str, int]:
    """Single-pass token replacement using regex alternation.

    Quote variants in the run text are matched against the (already
    quote-normalized) shift_map keys. When no token matches, the original
    text is returned unchanged. When a token matches, the run is rewritten
    with the new token; quote style of the rewritten token follows the
    shift_map's value (ASCII apostrophe, matches refreshed chart cats).
    """
    if not shift or not text:
        return text, 0
    norm = _norm_quotes(text)
    pattern = re.compile(
        r"(?<!\w)(" + "|".join(re.escape(k) for k in shift) + r")(?!\w)"
    )
    new_text, count = pattern.subn(lambda m: shift[m.group(1)], norm)
    if count == 0:
        return text, 0
    return new_text, count


def _shape_text_frames(shape):
    """Yield (text_frame, locator_str) for tables (per cell) and text frames."""
    if shape.has_table:
        for row_idx, row in enumerate(shape.table.rows):
            for col_idx, cell in enumerate(row.cells):
                yield cell.text_frame, f"{shape.name}[{row_idx},{col_idx}]"
    elif shape.has_text_frame:
        yield shape.text_frame, shape.name


def _walk_all_shapes(shapes):
    """Yield top-level shapes plus shapes nested inside groups."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_all_shapes(s.shapes)
        else:
            yield s


def _pair_charts_by_position(src_shapes, ref_shapes):
    """Pair source/refreshed chart shapes by (left, top) — refresh doesn't
    move shapes, so positions are stable across the two decks. Falls back
    to name match when positions don't line up exactly."""
    src_sorted = sorted(
        [s for s in src_shapes if s.has_chart],
        key=lambda s: ((s.top or 0), (s.left or 0)),
    )
    ref_sorted = sorted(
        [s for s in ref_shapes if s.has_chart],
        key=lambda s: ((s.top or 0), (s.left or 0)),
    )
    pairs = []
    if len(src_sorted) == len(ref_sorted):
        for src, ref in zip(src_sorted, ref_sorted):
            pairs.append((src, ref))
    else:
        # Length mismatch — pair by name where unique, else skip.
        ref_by_name = {}
        for r in ref_sorted:
            ref_by_name.setdefault(r.name, []).append(r)
        for src in src_sorted:
            cands = ref_by_name.get(src.name, [])
            if len(cands) == 1:
                pairs.append((src, cands[0]))
    return pairs


def sync_period_labels(
    source_pptx: str | Path,
    refreshed_pptx: str | Path,
    *,
    out_pptx: str | Path | None = None,
    sidecar_path: str | Path | None = None,
) -> dict:
    """Rewrite period tokens in text frames adjacent to refreshed charts.

    Reads source chart cats and refreshed chart cats; per slide, builds a
    union shift map; rewrites text-frame and table-cell runs that sit
    immediately above or below a chart and contain a matching old wave
    token. Saves to `out_pptx` (defaults to in-place over `refreshed_pptx`).

    Returns a diagnostic dict with per-slide shift maps, edits, and
    conflicts. Also writes it to `sidecar_path` if given.
    """
    src_prs = Presentation(str(source_pptx))
    ref_prs = Presentation(str(refreshed_pptx))

    diag = {
        "source": str(source_pptx),
        "refreshed": str(refreshed_pptx),
        "slides": [],
        "totals": {"slides_with_shift": 0, "edits": 0, "conflicts": 0},
    }

    n_slides = min(len(src_prs.slides), len(ref_prs.slides))
    for s_idx in range(n_slides):
        src_slide = src_prs.slides[s_idx]
        ref_slide = ref_prs.slides[s_idx]

        src_shapes_flat = list(_walk_all_shapes(src_slide.shapes))
        ref_shapes_flat = list(_walk_all_shapes(ref_slide.shapes))
        pairs = _pair_charts_by_position(src_shapes_flat, ref_shapes_flat)

        per_chart_shifts: list[dict[str, str]] = []
        chart_bboxes: list[tuple[int, int, int, int]] = []
        for src_sh, ref_sh in pairs:
            shift = _build_shift_map(
                _chart_categories(src_sh), _chart_categories(ref_sh)
            )
            if shift:
                per_chart_shifts.append(shift)
                chart_bboxes.append(_bbox(ref_sh))

        if not per_chart_shifts:
            continue

        merged, conflicts = _merge_slide_shift(per_chart_shifts)
        slide_diag = {
            "slide_index": s_idx,
            "shift": merged,
            "conflicts": conflicts,
            "edits": [],
        }
        diag["slides"].append(slide_diag)

        if not merged:
            continue
        diag["totals"]["slides_with_shift"] += 1
        diag["totals"]["conflicts"] += len(conflicts)

        for sh in ref_shapes_flat:
            if not (sh.has_table or sh.has_text_frame):
                continue
            tb = _bbox(sh)
            if tb[0] < 0:
                continue
            if not any(_is_adjacent(tb, cb) for cb in chart_bboxes):
                continue
            for tf, locator in _shape_text_frames(sh):
                for para in tf.paragraphs:
                    for run in para.runs:
                        if not _is_label_run(run.text):
                            continue
                        new_text, n = _apply_shift_to_text(run.text, merged)
                        if n > 0:
                            slide_diag["edits"].append({
                                "shape": locator,
                                "before": run.text,
                                "after": new_text,
                                "n_replacements": n,
                            })
                            run.text = new_text

        diag["totals"]["edits"] += len(slide_diag["edits"])

    out = Path(out_pptx) if out_pptx else Path(refreshed_pptx)
    ref_prs.save(str(out))

    if sidecar_path:
        Path(sidecar_path).write_text(
            json.dumps(diag, indent=2, ensure_ascii=False), encoding="utf-8",
        )

    return diag
