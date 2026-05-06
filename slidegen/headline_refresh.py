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


def _find_headline_shape(slide):
    """Pick the slide's headline text frame — topmost wide title shape.

    Updated heuristic (slide 73 fix on Repatha ATU): the previous logic
    required a narrative-verb match (rose / declined / trended / ...),
    which on multi-Title slides skipped the real headline at the top
    when its wording happened to lack those verbs (e.g. "Overall, HCPs
    exhibit a growing trend in their intent..."). It then fell to the
    next Title shape — a section header below ("Expected Increase in
    Prescription – Primary Prevention - Trended") — which IS lower on
    the slide but happens to contain a verb.

    The user's contract for "headliner": top-most wide text frame with
    a 1-3 line sentence, NOT a roadmap or section-header strip. Reading
    "narrative-shaped" gates BUILDING a fresh narrative (caller's job),
    not which shape is the headline.

    Heuristic order:
      1. zrx_<slide:03d>_001 — SlideGen pipeline's first-shape convention.
      2. Title*/Headline*-named shapes that are reasonably wide (>= 40%
         of slide width) and have substantive text — pick the topmost.
      3. Any wide text frame in the upper portion of the slide with
         substantive text (10+ chars) — pick the topmost.

    Returns None only when no title-like shape exists at all.
    """
    slide_w_emu = (slide.part.package.presentation_part.presentation
                   .slide_width if hasattr(slide.part.package, "presentation_part")
                   else 9144000)  # fallback ~10in
    min_width_emu = int(slide_w_emu * 0.4)

    title_candidates = []
    other_candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        name = (shape.name or "").lower().strip()
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        if name.startswith("zrx_") and name.endswith("_001"):
            return shape
        top = shape.top or 0
        width = shape.width or 0
        if name.startswith("title") or name.startswith("headline"):
            if width >= min_width_emu and len(text) >= 10:
                title_candidates.append((top, shape, text))
            continue
        if (len(text) > 10
                and top < _HEADLINE_TOP_LIMIT_EMU
                and width >= min_width_emu):
            other_candidates.append((top, shape, text))

    if title_candidates:
        title_candidates.sort(key=lambda t: t[0])
        return title_candidates[0][1]
    if other_candidates:
        other_candidates.sort(key=lambda t: t[0])
        return other_candidates[0][1]
    return None


def _largest_chart(slide):
    """The chart shape on this slide with the largest area."""
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
    plot = shape.chart.plots[0]
    cats = [str(c) for c in plot.categories]
    lines = [f"Categories: {cats}"]
    for s in plot.series:
        vals = [round(float(v), 4) if v is not None else None
                for v in s.values]
        lines.append(f"Series '{s.name or ''}': {vals}")
    return "\n".join(lines)


def _build_prompt(old_headline: str, src_chart_summary: str,
                  ref_chart_summary: str, slide_context: str = "") -> str:
    is_narrative = _looks_like_data_narrative(old_headline)
    voice_instruction = (
        "Rewrite the headline to reflect the new values, preserving the "
        "original's voice, structure, and tone."
        if is_narrative
        else "The original is a category/section label — replace it with a "
             "narrative-style verdict that reflects the new data. Use a "
             "natural-sounding 2-3 line claim, not a label."
    )
    return f"""You are updating a PowerPoint slide headline after a data refresh. The output you produce will be written verbatim into the slide's title shape — there is no editor in between.

{voice_instruction} The headline is a VERDICT — a 2-3 line claim about what mattered — not a data summary or analysis.

ORIGINAL HEADLINE
{old_headline}

ORIGINAL CHART DATA (what the headline was written about)
{src_chart_summary}

NEW CHART DATA (after refresh)
{ref_chart_summary}

CONTEXT
{slide_context or '(none)'}

INSTRUCTIONS — STRICT
- Output ONLY the headline text. It will be written verbatim into the title shape.
- LENGTH: 2-3 lines maximum. Aim for 80-160 characters total.
- A headline is a VERDICT — what mattered, expressed as a claim. Not a data inventory.
- Reflect actual values from the NEW data with concrete movement words ("rose 6 points", "dropped 4 points", "held steady"). Treat moves under 1 point as flat.
- If the new data is empty / all-None / clearly corrupted, output the ORIGINAL headline unchanged.
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
    """Rewrite headlines for connected slides whose chart values moved.

    only_slides: optional 0-based slide-index whitelist (None = all connected).
    dry_run: if True, don't write the output deck — just return planned updates.
    """
    from pptx import Presentation
    from tests.evals.end_to_end.compare_decks import _extract_chart_data

    if not dry_run and not os.environ.get("LLM_API_KEY"):
        raise RuntimeError(
            "LLM_API_KEY not set. Add it to docs/.env or export it."
        )

    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    spec_slides = {s["slide_index"] for s in spec.get("slides", [])}

    src = Presentation(source_pptx)
    ref = Presentation(refreshed_pptx)
    updates: list[HeadlineUpdate] = []

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        if only_slides is not None and s_idx not in only_slides:
            continue
        if s_idx not in spec_slides:
            continue
        src_chart = _largest_chart(s_slide)
        ref_chart = _largest_chart(r_slide)
        if src_chart is None or ref_chart is None:
            updates.append(HeadlineUpdate(s_idx, "", "", "", "no_chart"))
            continue
        try:
            _, src_series = _extract_chart_data(src_chart)
            _, ref_series = _extract_chart_data(ref_chart)
        except Exception as exc:
            updates.append(HeadlineUpdate(
                s_idx, "", "", f"extract error: {exc}", "no_chart"))
            continue
        if _values_eq(src_series, ref_series):
            updates.append(HeadlineUpdate(
                s_idx, "", "", "values unchanged", "unchanged"))
            continue
        if _all_none(ref_series) and not _all_none(src_series):
            # Refresh corrupted the chart (all values None). Don't ask the LLM
            # to write a headline against empty data — preserve the original.
            updates.append(HeadlineUpdate(
                s_idx, "", "", "refreshed values all-None (corrupted)", "unchanged"))
            continue

        headline_shape = _find_headline_shape(r_slide)
        if headline_shape is None:
            updates.append(HeadlineUpdate(
                s_idx, "", "", "", "no_headline"))
            continue
        old_headline = headline_shape.text_frame.text.strip()

        src_summary = _summarize_chart(src_chart)
        ref_summary = _summarize_chart(ref_chart)
        prompt = _build_prompt(old_headline, src_summary, ref_summary)

        if dry_run:
            new_headline = ""   # placeholder; not calling the API
        else:
            try:
                new_headline = _call_claude(prompt, model=model)
            except Exception as exc:
                updates.append(HeadlineUpdate(
                    s_idx, old_headline, "", f"api error: {exc}", "api_error"))
                continue

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

        updates.append(HeadlineUpdate(
            s_idx, old_headline, new_headline,
            f"src: {src_summary[:80]}... | ref: {ref_summary[:80]}...",
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
