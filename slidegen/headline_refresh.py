"""Step 7 — LLM-driven post-refresh headline rewrite.

For each connected slide whose chart values changed during refresh, ask
Claude to rewrite the slide's headline to reflect the new wave's data
in the same voice and structure as the original.

Distinct from `headline_writer.py` — that's a template-based generator
used in Pradeep's SlideSpec / Path B workflow. This module is purely
for the connected-slide refresh use case (Step 7 of the 7-step
framework).

Usage:
    from slidegen.headline_refresh import refresh_headlines
    refresh_headlines(
        source_pptx="path/to/source.pptx",
        refreshed_pptx="path/to/refreshed.pptx",
        spec_path="path/to/full_spec.json",
        out_pptx="path/to/refreshed_with_new_headlines.pptx",
    )

Requires: pip install anthropic python-dotenv. Reads LLM_API_KEY and
LLM_API_BASE from docs/.env (LiteLLM gateway, Anthropic-compatible).

Approach:
  1. Iterate connected slides (slides in the spec).
  2. Pick the slide's main chart (largest by area).
  3. Compare source vs refreshed chart values; skip if unchanged.
  4. Find the slide's headline (top-most text frame).
  5. Send (old headline + old chart values + new chart values) to Claude.
  6. Write the refreshed headline back into the headline shape.
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
    status: str  # 'updated' | 'unchanged' | 'no_chart' | 'no_headline' | 'api_error'


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


def _find_headline_shape(slide):
    """Pick the slide's TALKING HEADER — the top-left long-sentence text frame.

    Strict heuristic to avoid confusion with the slide-section title
    (short label that sits BELOW the talking header) and the roadmap
    (narrow strip at top-right):

      1. Width >= 40% of slide width      (excludes roadmap)
      2. Top < 2 inches from slide top    (excludes body content)
      3. Text length >= 30 chars          (excludes section-title labels)
      4. Among survivors: pick TOPMOST.   (talking header is above section title)
      5. Tiebreak (same top, within ~0.2in): pick LEFTMOST.
                                            (talking header is top-left)

    SlideGen's own naming convention `zrx_<slide:03d>_001` for the
    first shape on a slide takes priority when present (and only if it
    also passes the length filter — guards against `zrx_001_001` being
    a banner).

    Returns None when no talking header exists (e.g., dividers, covers).
    """
    slide_w_emu = (slide.part.package.presentation_part.presentation
                   .slide_width if hasattr(slide.part.package, "presentation_part")
                   else 9144000)  # fallback ~10in
    min_width_emu = int(slide_w_emu * 0.4)
    tiebreak_top_band_emu = 200000  # ~0.2 inches

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
        # SlideGen first-shape convention — accept only if it also passes
        # length + position filters (banner-only `zrx_*_001` is excluded).
        name = (shape.name or "").lower().strip()
        if name.startswith("zrx_") and name.endswith("_001"):
            if (width >= min_width_emu
                    and top < _HEADLINE_TOP_LIMIT_EMU
                    and len(text) >= _HEADLINE_MIN_CHARS):
                return shape
            # Otherwise fall through to the general filter
        if (width >= min_width_emu
                and top < _HEADLINE_TOP_LIMIT_EMU
                and len(text) >= _HEADLINE_MIN_CHARS):
            candidates.append((top, left, shape, text))

    if not candidates:
        return None
    # Sort by top, then by left (talking header is topmost; ties resolved
    # leftmost since talking header is top-LEFT and section title may be
    # top-CENTERED).
    candidates.sort(key=lambda c: (c[0], c[1]))
    # Bucket all candidates within the tiebreak band of the topmost into a
    # group, then pick leftmost within that group.
    topmost_y = candidates[0][0]
    top_band = [c for c in candidates if c[0] - topmost_y <= tiebreak_top_band_emu]
    top_band.sort(key=lambda c: c[1])  # leftmost wins
    return top_band[0][2]


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

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        if only_slides is not None and s_idx not in only_slides:
            continue

        # 1. Find the talking header. No talking header → skip silently.
        headline_shape = _find_headline_shape(r_slide)
        if headline_shape is None:
            updates.append(HeadlineUpdate(s_idx, "", "", "", "no_headline"))
            continue
        old_headline = headline_shape.text_frame.text.strip()

        # 2. Gather every chart + table on the refreshed slide. Skip the
        #    slide entirely if there's nothing to summarise (cover, divider).
        ref_data = _all_data_shapes(r_slide)
        if not ref_data:
            updates.append(HeadlineUpdate(
                s_idx, old_headline, "", "no chart or table on slide", "no_chart"))
            continue
        ref_summaries = [(kind, name, summary) for (kind, _shp, name, summary) in ref_data]

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
        #    don't ask the LLM to invent. Preserve the original.
        all_empty = True
        for _kind, _name, summary in ref_summaries:
            if any(ch.isdigit() for ch in summary):
                all_empty = False
                break
        if all_empty:
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
            f"{connected_flag}: {n_shapes} data shape(s) [{kinds}]",
            "updated"))

    if not dry_run:
        ref.save(out_pptx)

    return updates


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("deck_key", choices=["atu_q1_26", "creon_pet_w33"])
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--slide", type=int, action="append", default=None,
                        help="0-based slide indices to process (repeatable).")
    parser.add_argument("--model", default="anthropic/claude-sonnet-4-6")
    args = parser.parse_args()

    src_path_map = {
        "atu_q1_26": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
        "creon_pet_w33": REPO_ROOT / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
    }
    src = src_path_map[args.deck_key]
    ref = REPO_ROOT / "output" / "step2_test_connected" / f"{args.deck_key}.pptx"
    spec = REPO_ROOT / "output" / "step2_test_connected" / f"{args.deck_key}_full_spec.json"
    out = REPO_ROOT / "output" / "step2_test_connected" / f"{args.deck_key}_with_headlines.pptx"

    updates = refresh_headlines(
        source_pptx=str(src),
        refreshed_pptx=str(ref),
        spec_path=str(spec),
        out_pptx=str(out),
        only_slides=args.slide,
        model=args.model,
        dry_run=args.dry_run,
    )

    from collections import Counter
    counts = Counter(u.status for u in updates)
    print(f"=== {args.deck_key} | headline refresh ===")
    for k, n in counts.most_common():
        print(f"  {k}: {n}")
    if not args.dry_run:
        print(f"\nWrote: {out.relative_to(REPO_ROOT)}")
        sidecar = out.with_name(out.stem + "_status.json")
        sidecar.write_text(json.dumps(
            {"updates": [{"slide_idx": u.slide_idx, "status": u.status,
                          "old_headline": u.old_headline,
                          "new_headline": u.new_headline,
                          "summary": u.chart_summary} for u in updates]},
            indent=2,
        ), encoding="utf-8")
        print(f"Wrote: {sidecar.relative_to(REPO_ROOT)}")
    def _safe(s: str) -> str:
        return (s or "").encode("ascii", "replace").decode("ascii")

    print(f"\nSamples (first 5 updated):")
    n = 0
    for u in updates:
        if u.status == "updated":
            print(f"\nSlide {u.slide_idx + 1}:")
            print(f"  OLD: {_safe(u.old_headline)}")
            print(f"  NEW: {_safe(u.new_headline)}")
            n += 1
            if n >= 5:
                break
