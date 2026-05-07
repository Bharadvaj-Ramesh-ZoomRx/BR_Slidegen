"""Headliner full workflow — talking-header rewrite for any deck.

Walks every slide in a deck and rewrites the talking header (top-left
long-sentence text frame) using Claude, given:
  - The current data on that slide (every chart + table on the slide)
  - The existing talking header (for voice / tone)
  - Optionally the source-side data so Claude can describe MOVEMENT
    (when the slide is in a connected refresh spec)

Layout convention this module enforces:
  - top-LEFT  → talking header   (REWRITE — narrative claim, full sentence)
  - top-RIGHT → roadmap          (LEAVE — narrow nav strip)
  - below talking header → slide-section title (LEAVE — short label)

Distinct from `headline_writer.py` — that's a template-based generator
used in Pradeep's SlideSpec / Path B workflow. This module is the
LLM-driven post-refresh / on-demand talking-header rewriter.

When this runs:
  - As Step 3 of `slidegen.refresh_pipeline` (after refresh_deck_from_spec
    + stamp_refresh_notes/dynamic_tags); see `refresh_pipeline.py`.
  - Standalone via this module's CLI (run-headliner-only on any deck —
    useful for testing the rewriter without going through the full
    refresh pipeline). See the `if __name__ == "__main__"` block.

Coverage rules (changed 2026-05-08):
  - Every slide is processed — connected and non-connected. Mixed
    slides (some connected, some non-connected components) are
    rewritten because the non-connected piece may have manual edits.
  - Tables count as "data shapes" — table-only slides get headlines.
  - Headlines are rewritten even when chart values are unchanged
    pre/post (a non-connected component on the slide may have moved).
  - Only EMPTY/null slides preserve the original headline (no data
    for Claude to ground a verdict against).

Usage (Python):
    from slidegen.headliner_full_workflow import refresh_headlines
    refresh_headlines(
        source_pptx="path/to/source.pptx",
        refreshed_pptx="path/to/refreshed.pptx",
        spec_path="path/to/full_spec.json",   # optional
        out_pptx="path/to/output.pptx",
    )

Usage (CLI standalone — for testing the rewriter on any deck):
    python -m slidegen.headliner_full_workflow <deck.pptx>
    python -m slidegen.headliner_full_workflow <deck.pptx> \
        --out /tmp/with_headlines.pptx --slide 4 --slide 7 --dry-run

Requires: pip install anthropic python-dotenv. Reads LLM_API_KEY and
LLM_API_BASE from docs/.env (LiteLLM gateway, Anthropic-compatible).
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

REPO_ROOT = Path(__file__).resolve().parents[1]

from dotenv import load_dotenv
load_dotenv(REPO_ROOT / "docs" / ".env")


@dataclass
class HeadlineUpdate:
    slide_idx: int
    old_headline: str
    new_headline: str
    chart_summary: str
    # status:
    #   'updated'         — LLM rewrote, deck edited
    #   'unchanged'       — preserved (all data shapes empty/null)
    #   'no_chart'        — no chart or table on the slide
    #   'no_headline'     — couldn't find or create a headline shape
    #   'skipped_label'   — existing talking header reads as a slide-section
    #                       label (e.g., "Message Recall CREON") — NOT
    #                       rewritten. Empty / created shapes still get
    #                       written; only EXISTING label-style headers skip.
    #   'api_error'       — LLM call failed
    status: str


_NARRATIVE_VERBS = {
    # Movement
    "rose", "fell", "grew", "dropped", "increased", "decreased", "declined",
    "rebounded", "improved", "worsened", "lost", "gained", "expanded",
    "contracted", "lifted", "trended",
    # Comparison / position
    "outperformed", "underperformed", "led", "topped", "exceeded", "lagged",
    "trailed", "matched", "dominated", "ranked",
    # Stability
    "remained", "held", "retained", "continued", "stayed",
    # Change
    "shifted", "moved",
    # Reporting / observation
    "showed", "saw", "reported", "reached", "captured",
    # Perception / association
    "preferred", "favored", "associated", "linked", "perceived", "viewed",
    "described", "characterized", "considered", "rated", "regarded",
}

_CLAIM_KEYWORDS = [r"\bhighest\b", r"\blowest\b"]

# Real headlines live near the top of the slide; footnote/disclaimer text
# lives near the bottom. ~2 inches tolerates banner-then-headline stacking.
_HEADLINE_TOP_LIMIT_EMU = 2 * 914400  # 2 inches in English Metric Units


def _looks_like_data_narrative(text: str) -> bool:
    """Return True only if text reads like a data-driven claim.

    A narrative either uses a comparison/movement verb (rose, declined,
    outperformed, retained...) or contains an explicit highest/lowest
    claim. Chart subtitles, section labels, and methodology footnotes
    typically have none.
    """
    if len(text) < 40:
        return False
    lower = text.lower()
    if any(re.search(rf"\b{v}\b", lower) for v in _NARRATIVE_VERBS):
        return True
    if any(re.search(p, lower) for p in _CLAIM_KEYWORDS):
        return True
    return False


# Talking-header vs slide-section-title distinction.
# Layout convention:
#   top-LEFT  → talking header (narrative claim, full sentence; this is
#               what we rewrite)
#   top-RIGHT → roadmap (narrow nav strip; never rewrite — excluded by
#               width filter)
#   below talking header → slide-section title (short label like
#                          "Message Recall"; never rewrite — excluded
#                          by text-length filter)
#
# Talking headers are sentences (typically 60-200 chars). Section titles
# are 1-4 word labels (typically 5-30 chars). The 30-char threshold cleanly
# separates them in every deck we've inspected (Testing Deck, CREON, ATU,
# AVEO, Repatha, ILAI, Datroway).
_HEADLINE_MIN_CHARS = 30


# Words / patterns that suggest "section title" or "data label" rather
# than a narrative claim. Used by the narrative scorer below to penalise
# label-style candidates so the actual talking header wins on slides
# where multiple wide-tall-long shapes coexist.
_LABEL_PATTERNS = (
    re.compile(r"^[A-Z][A-Za-z\s&/–-]{4,40}\s*[-–]\s*[A-Z]", re.M),  # "X – Y"
    re.compile(r"\bRespondent\s+Profile\b", re.I),
    re.compile(r"\bInteraction\s+Details\b", re.I),
    re.compile(r"\bSlide\s+\d+\b", re.I),
    re.compile(r"\bSection\s+\d+\b", re.I),
    re.compile(r"^[A-Z][A-Z\s]+$"),  # ALL CAPS line
)
_DATE_RANGE_RE = re.compile(
    r"\b("
    r"[A-Z][a-z]{2}['’‘]\d{2}\s*[-–]\s*[A-Z][a-z]{2}['’‘]\d{2}"
    r"|Q[1-4]['’‘]?\s*\d{2,4}"
    r"|Wave\s+\d+"
    r"|[A-Z][a-z]{2}-[A-Z][a-z]{2}\s*['’‘]?\d{2,4}"
    r")\b"
)
_NUMERIC_RE = re.compile(r"\d+%?")
_COMPARISON_RE = re.compile(
    r"\b(vs\.?|versus|compared|against|while|whereas|whilst|than)\b", re.I
)


def _score_narrative(text: str) -> float:
    """Score how narrative-like a text frame reads.

    Higher score = more like a data-driven talking header.
    Lower (or negative) = more like a section title / label.

    Cheap heuristic — used as a first pass to pick a talking header
    when multiple candidates pass the layout filter. Reserved as input
    to LLM arbitration on close calls.
    """
    if not text:
        return -10.0
    score = 0.0
    lower = text.lower()
    # Narrative verbs (rose, fell, retained, dominated, …)
    score += 2.0 * sum(1 for v in _NARRATIVE_VERBS
                       if re.search(rf"\b{v}\b", lower))
    # Claim keywords (highest, lowest)
    score += 2.0 * sum(1 for p in _CLAIM_KEYWORDS if re.search(p, lower))
    # Numeric content (a real claim usually cites at least one number)
    score += min(5, len(_NUMERIC_RE.findall(text))) * 1.0
    # Comparison structures (vs, compared, while)
    score += 1.5 * len(_COMPARISON_RE.findall(text))
    # Length bonus (genuine narratives are usually longer)
    L = len(text)
    if L >= 80:
        score += 2.0
    if L >= 120:
        score += 1.5
    if L >= 180:
        score += 1.0
    # Penalties for label-shape patterns
    for pat in _LABEL_PATTERNS:
        if pat.search(text):
            score -= 3.0
            break
    # Pure-date-range slabs are section/banner labels, not insights
    date_hits = len(_DATE_RANGE_RE.findall(text))
    if date_hits and L < 60:
        score -= 4.0
    return score


def _llm_pick_talking_header(candidates_text: list[str], llm_arbiter) -> int | None:
    """Ask Claude which candidate reads most like a data-driven
    narrative claim (the talking header) vs a section title or label.

    Args:
      candidates_text: 1-N candidate text strings.
      llm_arbiter: callable (prompt: str) -> str. Used to get Claude's
        verdict. Pass `_call_claude` from this module for production.

    Returns:
      0-based index of the chosen candidate, or None on parse failure.
    """
    if not candidates_text:
        return None
    if len(candidates_text) == 1:
        return 0
    listing = "\n".join(
        f"({i+1}) {t.replace(chr(10), ' ')[:300]}"
        for i, t in enumerate(candidates_text)
    )
    prompt = (
        "On a PowerPoint slide, one text frame is the TALKING HEADER — a "
        "1-3 sentence data-driven INSIGHT or VERDICT about what the chart/"
        "table data shows (movement, comparison, claim). The others are "
        "section titles, slide labels, or banners (short, descriptive, "
        "not a claim).\n\n"
        "Pick the talking header. Output ONLY a single integer (the "
        "candidate number, 1-based). No explanation, no punctuation.\n\n"
        f"{listing}"
    )
    try:
        out = llm_arbiter(prompt).strip()
    except Exception:
        return None
    # Accept the first integer found in the output
    m = re.search(r"\d+", out)
    if not m:
        return None
    idx = int(m.group(0)) - 1
    if 0 <= idx < len(candidates_text):
        return idx
    return None


def _find_headline_shape(slide, llm_arbiter=None):
    """Pick the slide's TALKING HEADER — the top-left long-sentence text frame.

    Layered decision:

      1. Layout filter (cheap, geometry-only):
         * width >= 40% of slide width   (excludes narrow roadmap)
         * top < 2 inches from slide top (excludes body content)
         * text length >= 30 chars       (excludes most section titles)

      2. Narrative scoring (cheap, language-only):
         * narrative verbs / claim keywords / numbers / comparisons → +
         * label patterns ("Respondent Profile", "Interaction Details",
           "MMM'YY – MMM'YY") → −
         If one candidate has a decisive lead (>= 3 points over runner-
         up), pick it without LLM.

      3. LLM arbitration (when scores are close AND llm_arbiter given):
         Send candidate texts to Claude — "which is the talking header
         vs section title?" — and use Claude's pick.

      4. Fallback (no LLM available, scores still tied):
         Pick topmost-leftmost (the layout convention).

    `llm_arbiter` is optional — `refresh_headlines` passes `_call_claude`
    so production gets Claude judgment on edge cases. Tests / callers
    that want pure-heuristic behavior can omit it.

    SlideGen's `zrx_<slide:03d>_001` naming convention takes priority
    when present (and only if the shape passes the length+position
    filters — guards against `zrx_001_001` being a banner).

    Returns None when no qualifying shape exists (e.g., dividers).
    """
    slide_w_emu = (slide.part.package.presentation_part.presentation
                   .slide_width if hasattr(slide.part.package, "presentation_part")
                   else 9144000)
    min_width_emu = int(slide_w_emu * 0.4)
    tiebreak_top_band_emu = 200000  # ~0.2 inches
    DECISIVE_SCORE_GAP = 3.0

    candidates = []  # (top, left, shape, text)
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        top = shape.top or 0
        left = shape.left or 0
        width = shape.width or 0
        name = (shape.name or "").lower().strip()
        # SlideGen first-shape convention — only if it also passes filters.
        if name.startswith("zrx_") and name.endswith("_001"):
            if (width >= min_width_emu
                    and top < _HEADLINE_TOP_LIMIT_EMU
                    and len(text) >= _HEADLINE_MIN_CHARS):
                return shape
        if (width >= min_width_emu
                and top < _HEADLINE_TOP_LIMIT_EMU
                and len(text) >= _HEADLINE_MIN_CHARS):
            candidates.append((top, left, shape, text))

    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0][2]

    # Score each candidate by narrative quality.
    scored = [(c[0], c[1], c[2], c[3], _score_narrative(c[3])) for c in candidates]
    scored.sort(key=lambda x: x[4], reverse=True)
    top_score = scored[0][4]
    runner_up = scored[1][4]

    if (top_score - runner_up) >= DECISIVE_SCORE_GAP:
        return scored[0][2]

    # Scores are close — try LLM arbitration on the tied/near-tied set.
    near_tied_idx = [i for i, s in enumerate(scored)
                     if (top_score - s[4]) < DECISIVE_SCORE_GAP]
    if llm_arbiter and len(near_tied_idx) >= 2:
        texts = [scored[i][3] for i in near_tied_idx]
        pick = _llm_pick_talking_header(texts, llm_arbiter)
        if pick is not None:
            return scored[near_tied_idx[pick]][2]

    # Fallback: top-most, left-most among the near-tied set.
    near_tied = [scored[i] for i in near_tied_idx]
    near_tied.sort(key=lambda x: (x[0], x[1]))
    topmost_y = near_tied[0][0]
    top_band = [t for t in near_tied if t[0] - topmost_y <= tiebreak_top_band_emu]
    top_band.sort(key=lambda x: x[1])
    return top_band[0][2]


# Default talking-header zone used when creating a fresh textbox.
# Top-left, ~85% of slide width, ~1in tall — generous enough for a
# 2-3 line claim. Caller can resize after writing if needed.
_DEFAULT_TALKING_HEADER_LEFT_EMU = 274320      # 0.3 in
_DEFAULT_TALKING_HEADER_TOP_EMU = 182880       # 0.2 in
_DEFAULT_TALKING_HEADER_WIDTH_EMU = 7772400    # 8.5 in
_DEFAULT_TALKING_HEADER_HEIGHT_EMU = 914400    # 1.0 in


def _find_empty_headline_zone_shape(slide):
    """Find an empty-text shape sitting in the talking-header zone.

    Used when the user clears the talking-header text but leaves the
    shape on the slide. Returns the topmost-leftmost qualifying empty
    shape, or None.
    """
    slide_w_emu = (slide.part.package.presentation_part.presentation
                   .slide_width if hasattr(slide.part.package, "presentation_part")
                   else 9144000)
    min_width_emu = int(slide_w_emu * 0.4)

    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = (shape.text_frame.text or "").strip()
        if text:
            continue  # non-empty: handled by the regular finder
        top = shape.top or 0
        left = shape.left or 0
        width = shape.width or 0
        if width >= min_width_emu and top < _HEADLINE_TOP_LIMIT_EMU:
            candidates.append((top, left, shape))
    if not candidates:
        return None
    candidates.sort(key=lambda c: (c[0], c[1]))
    return candidates[0][2]


def _create_talking_header_shape(slide):
    """Create a fresh textbox in the slide's talking-header zone and
    return it. Used when neither a text-bearing nor an empty shape
    exists in the zone (i.e. the user deleted the shape entirely)."""
    return slide.shapes.add_textbox(
        _DEFAULT_TALKING_HEADER_LEFT_EMU,
        _DEFAULT_TALKING_HEADER_TOP_EMU,
        _DEFAULT_TALKING_HEADER_WIDTH_EMU,
        _DEFAULT_TALKING_HEADER_HEIGHT_EMU,
    )


def _ensure_headline_shape(slide, llm_arbiter=None):
    """Find the talking-header shape OR create one if missing.

    Three-pass discovery:
      Pass 1: existing text-bearing shape (`_find_headline_shape`).
              Handles the standard case + LLM arbitration on close calls.
      Pass 2: empty-text shape in the talking-header zone.
              Handles "user cleared the text but kept the shape".
      Pass 3: no qualifying shape exists — create a new textbox at the
              default top-left position. Handles "user deleted the shape".

    Returns: (shape, source) where source is one of:
      'existing'  — Pass 1 hit
      'empty'     — Pass 2 hit
      'created'   — Pass 3 created a new shape
    """
    shape = _find_headline_shape(slide, llm_arbiter=llm_arbiter)
    if shape is not None:
        return shape, "existing"
    shape = _find_empty_headline_zone_shape(slide)
    if shape is not None:
        return shape, "empty"
    shape = _create_talking_header_shape(slide)
    return shape, "created"


def _largest_chart(slide):
    """The chart shape on this slide with the largest area.

    Kept for backwards compatibility with existing callers / tests.
    The new headline flow uses _all_data_shapes() instead.
    """
    best, best_area = None, 0
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        area = (shape.width or 0) * (shape.height or 0)
        if area > best_area:
            best_area, best = area, shape
    return best


def _summarize_chart(shape) -> str:
    """Multi-line text summary of a chart's cats + series for the LLM prompt."""
    if not shape.has_chart:
        return ""
    try:
        plot = shape.chart.plots[0]
    except Exception:
        return ""
    cats = [str(c) for c in plot.categories]
    lines = [f"Categories: {cats}"]
    for s in plot.series:
        try:
            vals = [round(float(v), 4) if v is not None else None
                    for v in s.values]
        except Exception:
            vals = list(s.values) if hasattr(s, "values") else []
        lines.append(f"Series '{s.name or ''}': {vals}")
    return "\n".join(lines)


def _summarize_table(shape, max_rows: int = 25, max_cols: int = 12) -> str:
    """Text summary of a table's cells for the LLM prompt.

    Caps at max_rows × max_cols so very large tables don't blow up the
    prompt. Cell text is stripped; empty rows are dropped.
    """
    if not shape.has_table:
        return ""
    table = shape.table
    n_rows = min(len(table.rows), max_rows)
    n_cols = min(len(table.columns), max_cols)
    if n_rows == 0 or n_cols == 0:
        return ""
    lines = []
    for r in range(n_rows):
        cells = []
        for c in range(n_cols):
            try:
                txt = (table.cell(r, c).text_frame.text or "").strip()
            except Exception:
                txt = ""
            # Collapse whitespace inside cell
            txt = " ".join(txt.split())
            cells.append(txt)
        if any(cells):
            lines.append(" | ".join(cells))
    truncation = ""
    if len(table.rows) > max_rows:
        truncation += f"  (truncated: {len(table.rows) - max_rows} more rows)"
    if len(table.columns) > max_cols:
        truncation += f"  (truncated: {len(table.columns) - max_cols} more cols)"
    return "\n".join(lines) + (("\n" + truncation) if truncation else "")


def _all_data_shapes(slide):
    """Return every chart + table shape on the slide as
    [(kind, shape, name, summary), ...] sorted by area descending so the
    largest data shape leads the prompt.

    kind ∈ {'chart', 'table'}. Skips shapes with empty summaries.
    """
    out = []
    for shape in slide.shapes:
        kind = None
        summary = ""
        if shape.has_chart:
            kind = "chart"
            summary = _summarize_chart(shape)
        elif shape.has_table:
            kind = "table"
            summary = _summarize_table(shape)
        else:
            continue
        if not summary.strip():
            continue
        area = (shape.width or 0) * (shape.height or 0)
        out.append((area, kind, shape, shape.name or "", summary))
    out.sort(key=lambda t: -t[0])  # largest first
    return [(kind, shp, name, sm) for (_a, kind, shp, name, sm) in out]


def _build_prompt(old_headline: str, data_summaries: list[tuple[str, str, str]],
                  src_data_summaries: list[tuple[str, str, str]] | None = None,
                  slide_context: str = "") -> str:
    """Build the LLM prompt for talking-header rewrite.

    Args:
      old_headline: the existing talking header text.
      data_summaries: [(kind, name, summary), ...] for EVERY chart and
        table on the slide post-refresh, ordered largest-first. Claude
        sees them all and decides which to call out (could be all,
        could be one).
      src_data_summaries: optional [(kind, name, summary), ...] from
        SOURCE deck — when provided, Claude can compare to call out
        what moved. None for non-connected slides where we don't have a
        meaningful before/after split.
      slide_context: optional extra context (slide section, project, etc.)
    """
    is_narrative = _looks_like_data_narrative(old_headline)
    voice_instruction = (
        "Rewrite the headline to reflect the latest values, preserving the "
        "original's voice, structure, and tone where possible."
        if is_narrative
        else "The original may be a label or fragment — write a narrative-style "
             "VERDICT that captures what matters in the data. Natural-sounding "
             "2-3 line claim, not a label."
    )

    def _fmt_shapes(shapes):
        if not shapes:
            return "(none)"
        out = []
        for kind, name, summary in shapes:
            header = f"--- {kind.upper()}: {name or '(unnamed)'} ---"
            out.append(header + "\n" + summary)
        return "\n\n".join(out)

    src_block = ""
    if src_data_summaries:
        src_block = (
            f"\n\nSOURCE DATA (what the headline was written about — for comparison)\n"
            f"{_fmt_shapes(src_data_summaries)}\n"
        )

    return f"""You are updating a PowerPoint slide's TALKING HEADER. The output you produce will be written verbatim into the talking-header shape (top-left of the slide) — there is no editor in between.

{voice_instruction} A talking header is a VERDICT — a 2-3 line claim about what mattered — not a data inventory.

ORIGINAL TALKING HEADER
{old_headline}

CURRENT DATA ON THIS SLIDE (every chart and table, largest first)
{_fmt_shapes(data_summaries)}{src_block}

CONTEXT
{slide_context or '(none)'}

INSTRUCTIONS — STRICT
- Output ONLY the headline text. It will be written verbatim into the talking-header shape.
- LENGTH: 2-3 lines maximum. Aim for 80-180 characters total.
- A headline is a VERDICT — what mattered, expressed as a claim. Not a data inventory.
- The slide may have MULTIPLE charts/tables. You decide which numbers are worth calling out — sometimes one of them carries the story; sometimes two or three combine. Don't force every shape into the headline.
- Reflect actual values from the data with concrete movement words ("rose 6 points", "dropped 4 points", "held steady") when SOURCE DATA is provided so you can see movement. When SOURCE is not provided, describe the current state crisply ("X leads at 47%, Y trails at 12%").
- Treat moves under 1 point as flat.
- If every chart/table is empty / all-None / clearly corrupted, output the ORIGINAL talking header unchanged.
- Don't invent context that wasn't in the original.

DO NOT include any of the following in your output:
- "Looking at the new data:" or any analytical preamble
- Bullet points, **bold**, or markdown formatting
- "Key shift:", "Summary:", or section headers
- "---" separators or dividers
- Per-series breakdowns (e.g. "RINVOQ: 23%, Tremfya: 6%, No difference: 41%")
- Meta-commentary about the rewrite ("here's the updated headline...")
- Disclaimers, "[updated]" tags, or attribution

EXAMPLE OF WRONG OUTPUT:
Looking at the new data:
**RINVOQ:** 23% significantly better
**Tremfya:** 6% significantly better
Key shift: RINVOQ's lead has narrowed.
---
HCPs see most UC attributes as differentiated between RINVOQ and Tremfya.

EXAMPLE OF CORRECT OUTPUT:
HCPs see most UC attributes as differentiated between RINVOQ and Tremfya, with RINVOQ retaining an edge on rapid symptom relief and bio-experienced efficacy.
"""


def _call_claude(prompt: str, model: str = "anthropic/claude-sonnet-4-6") -> str:
    """Single Claude API call via LiteLLM gateway. Reads LLM_API_KEY + LLM_API_BASE."""
    import anthropic
    client = anthropic.Anthropic(
        api_key=os.environ["LLM_API_KEY"],
        base_url=os.environ["LLM_API_BASE"],
    )
    resp = client.messages.create(
        model=model,
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}],
    )
    text = ""
    for block in resp.content:
        if hasattr(block, "text"):
            text += block.text
    return text.strip()


def _all_none(series) -> bool:
    """True if every value in every series is None — refresh produced empty data."""
    if not series:
        return False
    for _name, vals in series:
        for v in vals:
            if v is not None:
                return False
    return True


def _values_eq(src_series, ref_series, tol=1e-3) -> bool:
    if len(src_series) != len(ref_series):
        return False
    for (sn, sv), (rn, rv) in zip(src_series, ref_series):
        if len(sv) != len(rv):
            return False
        for x, y in zip(sv, rv):
            if x is None and y is None:
                continue
            if x is None or y is None:
                return False
            if abs(x - y) > tol:
                return False
    return True


def refresh_headlines(
    source_pptx: str,
    refreshed_pptx: str,
    spec_path: str,
    out_pptx: str,
    only_slides: Optional[list[int]] = None,
    model: str = "anthropic/claude-sonnet-4-6",
    dry_run: bool = False,
) -> list[HeadlineUpdate]:
    """Rewrite the talking header on EVERY slide that has one + at least
    one chart or table.

    Behavior changes from the previous "Step 7 connected-only" flow:
      * Walks every slide, not just the spec-listed connected ones.
      * Reads ALL chart + table shapes on the slide (not just the
        largest chart) and sends every summary to Claude. Claude decides
        which numbers are worth calling out.
      * Doesn't skip when chart values are unchanged. Slides with mixed
        connected + non-connected components might have new manual edits
        on the non-connected side that the headline should reflect.
      * `spec_path` is now optional context — when present, source data is
        included in the prompt so Claude can describe MOVEMENT
        ("rose 6 points"); when None or the slide isn't in the spec, the
        prompt asks for a STATE description ("X leads at 47%").

    Args:
      source_pptx: pre-refresh deck. Used to read source-side chart/table
        values for movement-aware prompts on connected slides.
      refreshed_pptx: post-refresh deck — the file we read current data
        from and write headlines into.
      spec_path: optional spec JSON path. Used to flag which slides had
        connected refresh; doesn't gate processing.
      out_pptx: output path for the deck with rewritten talking headers.
      only_slides: optional 0-based slide-index whitelist (None = all).
      dry_run: if True, don't call the LLM or write the deck — just
        return planned updates.
    """
    from pptx import Presentation

    if not dry_run and not os.environ.get("LLM_API_KEY"):
        raise RuntimeError(
            "LLM_API_KEY not set. Add it to docs/.env or export it."
        )

    spec_slides: set[int] = set()
    if spec_path and Path(spec_path).exists():
        try:
            spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
            spec_slides = {s["slide_index"] for s in spec.get("slides", [])}
        except Exception:
            spec_slides = set()

    src = Presentation(source_pptx)
    ref = Presentation(refreshed_pptx)
    updates: list[HeadlineUpdate] = []

    # LLM arbiter — passed to _find_headline_shape so close-call cases
    # (multiple wide-tall-long candidates with similar narrative scores)
    # get Claude's judgment instead of falling to a topmost-leftmost
    # heuristic. Only fires when the heuristic alone can't decide.
    if dry_run:
        llm_arbiter = None
    else:
        def llm_arbiter(prompt: str) -> str:
            return _call_claude(prompt, model=model)

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        if only_slides is not None and s_idx not in only_slides:
            continue

        # 1. Gather every chart + table FIRST. Skip the slide entirely if
        #    there's nothing to headline (cover, divider, etc.) — this
        #    runs before _ensure_headline_shape so we don't leave an
        #    orphaned empty textbox on data-less slides.
        ref_data = _all_data_shapes(r_slide)
        if not ref_data:
            updates.append(HeadlineUpdate(
                s_idx, "", "", "no chart or table on slide", "no_chart"))
            continue
        ref_summaries = [(kind, name, summary) for (kind, _shp, name, summary) in ref_data]

        # 2. Find OR create the talking-header shape. Three pathways:
        #    - existing: shape exists with narrative text
        #    - empty:    shape exists in the zone but text was cleared
        #    - created:  no qualifying shape — new textbox added at top-left
        headline_shape, head_source = _ensure_headline_shape(
            r_slide, llm_arbiter=llm_arbiter)
        old_headline = headline_shape.text_frame.text.strip()

        # 2b. If the existing talking-header text reads like a slide-section
        #     title / banner ("Message Recall CREON", "Interaction Details
        #     - All Products", etc.) — the deck author intentionally didn't
        #     put a narrative claim in this slot, so don't overwrite it
        #     with one. Only skips the EXISTING-text path; empty and
        #     created shapes still get written.
        LABEL_SCORE_THRESHOLD = 2.0
        if head_source == "existing" and old_headline:
            old_score = _score_narrative(old_headline)
            if old_score < LABEL_SCORE_THRESHOLD:
                updates.append(HeadlineUpdate(
                    s_idx, old_headline, "",
                    f"existing header reads as slide label "
                    f"(score={old_score:.1f}); not rewriting",
                    "skipped_label"))
                continue

        # 3. Source summaries are optional — included only when the slide
        #    was connected so Claude can describe MOVEMENT. Non-connected
        #    slides get a state-description prompt instead.
        src_summaries = None
        if s_idx in spec_slides:
            src_data = _all_data_shapes(s_slide)
            if src_data:
                src_summaries = [(kind, name, summary)
                                 for (kind, _shp, name, summary) in src_data]

        # 4. Guard: if every refreshed shape is completely empty / null,
        #    don't ask the LLM to invent. Preserve the original (or, when
        #    we just created a placeholder shape, leave it empty rather
        #    than letting the LLM hallucinate).
        all_empty = True
        for _kind, _name, summary in ref_summaries:
            if any(ch.isdigit() for ch in summary):
                all_empty = False
                break
        if all_empty:
            # Don't keep an orphaned empty shape we just added.
            if head_source == "created":
                try:
                    headline_shape._element.getparent().remove(headline_shape._element)
                except Exception:
                    pass
            updates.append(HeadlineUpdate(
                s_idx, old_headline, "",
                "all data shapes empty/null — preserved", "unchanged"))
            continue

        # 5. Build the prompt and call Claude.
        prompt = _build_prompt(
            old_headline=old_headline,
            data_summaries=ref_summaries,
            src_data_summaries=src_summaries,
        )
        if dry_run:
            new_headline = ""
        else:
            try:
                new_headline = _call_claude(prompt, model=model)
            except Exception as exc:
                updates.append(HeadlineUpdate(
                    s_idx, old_headline, "", f"api error: {exc}", "api_error"))
                continue

        # 6. Write the new talking header into the shape (in place).
        if not dry_run:
            tf = headline_shape.text_frame
            p = tf.paragraphs[0]
            for run in list(p.runs):
                run.text = ""
            for extra_p in tf.paragraphs[1:]:
                extra_p._p.getparent().remove(extra_p._p)
            if not p.runs:
                run = p.add_run()
            else:
                run = p.runs[0]
            run.text = new_headline

        n_shapes = len(ref_summaries)
        kinds = ", ".join(sorted({k for k, _n, _s in ref_summaries}))
        connected_flag = "connected" if s_idx in spec_slides else "non-connected"
        updates.append(HeadlineUpdate(
            s_idx, old_headline, new_headline,
            f"{connected_flag} | {n_shapes} data shape(s) [{kinds}] | "
            f"shape: {head_source}",
            "updated"))

    if not dry_run:
        ref.save(out_pptx)

    return updates


if __name__ == "__main__":
    import argparse
    from collections import Counter

    parser = argparse.ArgumentParser(
        description=(
            "Run the talking-header rewriter on a deck. Standalone — does "
            "NOT need a refresh_status sidecar. Use this to test the "
            "rewriter on any deck (e.g. delete the headlines from a "
            "test deck, then point this CLI at it and Claude will "
            "regenerate them based on the data shapes on each slide)."
        )
    )
    parser.add_argument("deck", help="Path to the source PPTX")
    parser.add_argument(
        "--out", default=None,
        help="Path to write the deck with new headlines. "
             "Default: <deck>_with_headlines.pptx next to the source.",
    )
    parser.add_argument(
        "--spec", default=None,
        help="Optional spec JSON. When provided, slides listed in the "
             "spec get a connected-style prompt with source-vs-current "
             "comparison; slides not in the spec get a state-description "
             "prompt. Without --spec, every slide is treated as "
             "non-connected (state-description only).",
    )
    parser.add_argument(
        "--source", default=None,
        help="Optional second PPTX path. When the deck has been "
             "refreshed and you want Claude to compare source vs "
             "refreshed values for connected slides, pass the pre-refresh "
             "deck here. Default: same as `deck` (rewrite against current "
             "state only).",
    )
    parser.add_argument(
        "--slide", type=int, action="append", default=None,
        help="0-based slide indices to process (repeatable). Default: all.",
    )
    parser.add_argument(
        "--model", default="anthropic/claude-sonnet-4-6",
        help="LLM model id (LiteLLM-routed Anthropic).",
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Don't call the LLM or write the deck.")
    args = parser.parse_args()

    deck_path = Path(args.deck)
    if not deck_path.exists():
        parser.error(f"Deck not found: {deck_path}")

    out_path = Path(args.out) if args.out else \
        deck_path.with_name(deck_path.stem + "_with_headlines.pptx")
    src_path = Path(args.source) if args.source else deck_path

    updates = refresh_headlines(
        source_pptx=str(src_path),
        refreshed_pptx=str(deck_path),
        spec_path=str(args.spec) if args.spec else None,
        out_pptx=str(out_path),
        only_slides=args.slide,
        model=args.model,
        dry_run=args.dry_run,
    )

    counts = Counter(u.status for u in updates)
    print(f"=== Headliner | {deck_path.name} ===")
    for k, n in counts.most_common():
        print(f"  {k}: {n}")
    if not args.dry_run:
        print(f"\nWrote: {out_path}")
        sidecar = out_path.with_name(out_path.stem + "_status.json")
        sidecar.write_text(json.dumps(
            {"updates": [{"slide_idx": u.slide_idx, "status": u.status,
                          "old_headline": u.old_headline,
                          "new_headline": u.new_headline,
                          "summary": u.chart_summary} for u in updates]},
            indent=2,
        ), encoding="utf-8")
        print(f"Wrote: {sidecar}")

    def _safe(s: str) -> str:
        return (s or "").encode("ascii", "replace").decode("ascii")

    print("\nSamples (first 5 updated):")
    n = 0
    for u in updates:
        if u.status == "updated":
            print(f"\nSlide {u.slide_idx + 1}:")
            print(f"  OLD: {_safe(u.old_headline)}")
            print(f"  NEW: {_safe(u.new_headline)}")
            n += 1
            if n >= 5:
                break
