"""Per-deck refresh quality eval.

After a refresh runs, the pipeline writes a `*_refresh_status.json`
sidecar listing every connected component on every slide and its
status. This module reads that sidecar and produces a single
verdict: what % of components refreshed cleanly, what % were correctly
skipped (static-pinned), what % need user review, what % failed
outright.

Usage (CLI):
    python -m slidegen.refresh_eval <refresh_status.json>
    python -m slidegen.refresh_eval --deck <deck.pptx>     # auto-locates status

Usage (Python):
    from slidegen.refresh_eval import eval_refresh_status, format_report
    report = eval_refresh_status("output_testing/json_output/My_Deck_refresh_status.json")
    print(format_report(report))

The pipeline itself calls eval_refresh_status() at the end of every
run and prints the verdict alongside the slide-bucket counts, so
running the eval separately is for cases where you want to re-check a
prior run without rerunning the refresh.

Status taxonomy (component-level, mirrors what `refresh_deck_from_spec`
writes into the sidecar):

  PASSED      ok, ok_with_dynamic_added, ok_with_partial
              -> the component received refreshed data with no notes
                 worth surfacing, OR with API extras flowed in (tag
                 was dynamic), OR with a partial-alignment hint that
                 doesn't block usability.

  PRESERVED   static_pinned_skipped
              -> tag explicitly says "frozen, don't refresh." Correct
                 by design; counted separately so the eval doesn't
                 penalize decks that are static by intent.

  REVIEW      tag_mismatch, ambiguous_tag, alignment_failed
              -> data was preserved (source values intact) but the
                 tag points at the wrong analysis or can't disambiguate
                 multiple sibling charts. User-side fix.

  ERROR       not_found, empty, error, missing
              -> something genuinely broken — shape couldn't be
                 located, API returned nothing, or an exception
                 happened mid-refresh.
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path


PASSED = {"ok", "ok_with_dynamic_added", "ok_with_partial"}
# `no_data_for_shifted_window` is the connector's correct outcome when the
# API returns no records in the source's wave window — source preserved,
# nothing to refresh. Semantically identical to `static_pinned_skipped`
# (both are "correct by design — leave the chart alone"); counted as
# PRESERVED so a deck whose window has rolled past available data isn't
# penalised. Mirrors Connector spec §16: refresh leaves cells untouched
# when there is no new data to apply.
PRESERVED = {"static_pinned_skipped", "no_data_for_shifted_window"}
REVIEW = {"tag_mismatch", "ambiguous_tag", "alignment_failed"}
ERROR = {"not_found", "empty", "error", "missing"}

# Note-level review signals: a component can have status="ok" but carry
# a `tag_mismatch` / `selectedColumns_drift` note that downgrades it to
# REVIEW. The mapper records these by writing API data faithfully + a
# note, so the status alone misrepresents quality.
REVIEW_NOTE_KINDS = {"tag_mismatch", "selectedColumns_drift"}


def eval_refresh_status(status_path: str | Path) -> dict:
    """Read a refresh_status.json sidecar and produce a quality report.

    Returns a dict with:
      - total_components: int
      - by_status: dict[str, int]  (raw status counts)
      - passed / preserved / review / error: int
      - pass_rate_pct: float (passed / (total - preserved))
      - inclusive_pass_rate_pct: float (passed+preserved / total)
      - review_rate_pct: float
      - error_rate_pct: float
      - slides_total / slides_with_components: int
      - by_review_kind: dict (which review buckets contributed, slide list per bucket)
    """
    status = json.loads(Path(status_path).read_text(encoding="utf-8"))
    by_status: dict[str, int] = defaultdict(int)
    by_review_kind: dict[str, list[int]] = defaultdict(list)
    total = 0
    slides = status.get("slides", [])
    slides_with_comps = 0
    for s in slides:
        comps = (s.get("charts") or []) + (s.get("tables") or [])
        if comps:
            slides_with_comps += 1
        for c in comps:
            total += 1
            st = c.get("status", "missing")
            # Note-level downgrade: if a component carries a review-kind
            # note (tag_mismatch / selectedColumns_drift), classify it
            # as REVIEW even when status is "ok". The mapper writes API
            # data faithfully + a note in those cases — counting them
            # as PASSED would inflate the score.
            note_kinds = {n.get("kind") for n in (c.get("notes") or [])}
            review_note = note_kinds & REVIEW_NOTE_KINDS
            if review_note:
                # Use the note kind as the bucketed status so the
                # review breakdown surfaces it cleanly.
                effective = next(iter(review_note))
                by_status[effective] += 1
                by_review_kind[effective].append(s.get("slide_index"))
            else:
                by_status[st] += 1
                if st in REVIEW:
                    by_review_kind[st].append(s.get("slide_index"))

    n_passed = sum(by_status.get(s, 0) for s in PASSED)
    n_preserved = sum(by_status.get(s, 0) for s in PRESERVED)
    n_review = sum(by_status.get(s, 0)
                   for s in REVIEW | REVIEW_NOTE_KINDS)
    n_error = sum(by_status.get(s, 0) for s in ERROR)

    refreshable = max(0, total - n_preserved)
    pass_rate = (n_passed / refreshable * 100) if refreshable else 0.0
    inclusive_pass = (
        (n_passed + n_preserved) / total * 100
    ) if total else 0.0
    review_rate = (n_review / total * 100) if total else 0.0
    error_rate = (n_error / total * 100) if total else 0.0

    return {
        "total_components": total,
        "slides_total": len(slides),
        "slides_with_components": slides_with_comps,
        "by_status": dict(by_status),
        "passed": n_passed,
        "preserved": n_preserved,
        "review": n_review,
        "error": n_error,
        "pass_rate_pct": round(pass_rate, 1),
        "inclusive_pass_rate_pct": round(inclusive_pass, 1),
        "review_rate_pct": round(review_rate, 1),
        "error_rate_pct": round(error_rate, 1),
        "by_review_kind": {k: sorted(set(v))
                           for k, v in by_review_kind.items()},
    }


def format_report(report: dict, *, verbose: bool = False) -> str:
    """Render an eval report as human-readable text."""
    lines = [
        "-" * 60,
        f"Refresh quality report",
        "-" * 60,
        f"  Slides processed:        {report['slides_with_components']} / {report['slides_total']}",
        f"  Total components:        {report['total_components']}",
        "",
        f"  PASSED (clean refresh):  {report['passed']:4d}  ({report['pass_rate_pct']}% of refreshable)",
        f"  PRESERVED (static-pin):  {report['preserved']:4d}",
        f"  REVIEW (tag fix needed): {report['review']:4d}  ({report['review_rate_pct']}% of total)",
        f"  ERROR (broken):          {report['error']:4d}  ({report['error_rate_pct']}% of total)",
        "",
        f"  Inclusive pass rate:     {report['inclusive_pass_rate_pct']}%   "
        f"(passed + preserved / total)",
    ]
    if report["by_review_kind"]:
        lines.append("")
        lines.append("  Review breakdown (1-based slide numbers):")
        for kind, slide_idxs in sorted(report["by_review_kind"].items()):
            slides_str = ", ".join(str(i + 1) for i in slide_idxs[:15])
            extra = f" (+{len(slide_idxs) - 15} more)" if len(slide_idxs) > 15 else ""
            lines.append(f"    {kind:20s} on slides {slides_str}{extra}")
    if verbose:
        lines.append("")
        lines.append("  All status counts:")
        for k, v in sorted(report["by_status"].items(), key=lambda kv: -kv[1]):
            lines.append(f"    {k:30s} {v}")
    lines.append("-" * 60)
    return "\n".join(lines)


def _resolve_status_path(arg: str) -> Path:
    """Accept either a status JSON path directly, or a deck path
    (we'll auto-locate the matching status sidecar)."""
    p = Path(arg)
    if p.suffix.lower() == ".json":
        return p
    if p.suffix.lower() == ".pptx":
        # Find the sibling status sidecar in output_testing/json_output/
        stem = p.stem
        candidates = [
            p.parent.parent / "json_output" / f"{stem}_refresh_status.json",
            p.parent / f"{stem}_refresh_status.json",
        ]
        for c in candidates:
            if c.exists():
                return c
        raise FileNotFoundError(
            f"No refresh_status.json found for {p.name}. "
            f"Tried: {[str(c) for c in candidates]}"
        )
    return p  # last resort, let json.load complain


def main():
    parser = argparse.ArgumentParser(
        prog="python -m slidegen.refresh_eval",
        description=(
            "Read a refresh_status.json sidecar and report what % of "
            "components refreshed cleanly. Pass either the JSON path "
            "directly, or a deck path (the matching sidecar is auto-"
            "located in output_testing/json_output/)."
        ),
    )
    parser.add_argument(
        "target",
        help=(
            "Refresh status JSON path, OR a .pptx deck path "
            "(sidecar auto-located)."
        ),
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Show all raw status counts in addition to the summary.",
    )
    parser.add_argument(
        "--threshold", type=float, default=0.0,
        help=(
            "Exit non-zero if pass_rate_pct is below this threshold. "
            "Default 0 (always exit 0). Useful for CI gates: --threshold 80."
        ),
    )
    args = parser.parse_args()

    status_path = _resolve_status_path(args.target)
    report = eval_refresh_status(status_path)
    print(format_report(report, verbose=args.verbose))

    if report["pass_rate_pct"] < args.threshold:
        sys.exit(1)


if __name__ == "__main__":
    main()
