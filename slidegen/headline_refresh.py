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

Requires: pip install anthropic, env var ANTHROPIC_API_KEY.

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
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

REPO_ROOT = Path(__file__).resolve().parents[1]


@dataclass
class HeadlineUpdate:
    slide_idx: int
    old_headline: str
    new_headline: str
    chart_summary: str
    status: str  # 'updated' | 'unchanged' | 'no_chart' | 'no_headline' | 'api_error'


def _find_headline_shape(slide):
    """Pick the slide's headline text frame.

    Heuristic order:
      1. Shape named 'Title' or 'Headline' (case-insensitive)
      2. zrx_<slide:03d>_001 (SlideGen's first-shape convention)
      3. Top-most shape with a non-empty text frame and >10 chars of text
    """
    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        name = (shape.name or "").lower()
        if name in ("title", "headline"):
            return shape
        if name.startswith("zrx_") and name.endswith("_001"):
            return shape
        text = (shape.text_frame.text or "").strip()
        if len(text) > 10:
            top = shape.top or 0
            candidates.append((top, shape))
    if not candidates:
        return None
    candidates.sort(key=lambda t: t[0])
    return candidates[0][1]


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
    return f"""You are updating a PowerPoint slide headline after a data refresh.

The chart on this slide has been updated with new survey-wave data. Rewrite the headline to reflect the new values, while preserving the original headline's voice, structure, length, and tone.

ORIGINAL HEADLINE
{old_headline}

ORIGINAL CHART DATA (what the headline was written about)
{src_chart_summary}

NEW CHART DATA (after refresh)
{ref_chart_summary}

CONTEXT
{slide_context or '(none)'}

INSTRUCTIONS
- Output ONE headline.
- Match the original's word count within ~20% and structural style.
- Reflect actual values from the NEW data — be specific and accurate.
- If a value moved meaningfully vs original, describe the change with a concrete number ("rose 6 points", "dropped 4 points", "held steady"). Treat moves under 1 point as flat.
- Don't invent context that wasn't in the original.
- Don't add disclaimers, commentary, or "[updated]" tags.
- Output ONLY the headline text. No quotes, no preamble, no trailing notes.
"""


def _call_claude(prompt: str, model: str = "claude-sonnet-4-6") -> str:
    """Single Claude API call. Caller must have ANTHROPIC_API_KEY set."""
    import anthropic
    client = anthropic.Anthropic()
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
    model: str = "claude-sonnet-4-6",
    dry_run: bool = False,
) -> list[HeadlineUpdate]:
    """Rewrite headlines for connected slides whose chart values moved.

    only_slides: optional 0-based slide-index whitelist (None = all connected).
    dry_run: if True, don't write the output deck — just return planned updates.
    """
    from pptx import Presentation
    from tests.evals.end_to_end.compare_decks import _extract_chart_data

    if not dry_run and not os.environ.get("ANTHROPIC_API_KEY"):
        raise RuntimeError(
            "ANTHROPIC_API_KEY not set. Add it to .env or export it."
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
    parser.add_argument("--model", default="claude-sonnet-4-6")
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
    print(f"\nSamples (first 5 updated):")
    n = 0
    for u in updates:
        if u.status == "updated":
            print(f"\nSlide {u.slide_idx + 1}:")
            print(f"  OLD: {u.old_headline}")
            print(f"  NEW: {u.new_headline}")
            n += 1
            if n >= 5:
                break
