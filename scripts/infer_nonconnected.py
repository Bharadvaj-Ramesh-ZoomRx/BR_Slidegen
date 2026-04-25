"""infer_nonconnected.py — batch inferring for a whole deck.

Wraps Pradeep's per-slide inferring primitives in a deck-level loop. For each
Tier 2 (untagged) chart on a deck, scores it against every Tier 1 analysis_id
already in use on the same deck, picks the best match above threshold.

Two modes:
  --dry-run (default) : score + report only. No PPTX or spec writes.
  --apply             : stamp Connector tags into a copy of the PPTX and write
                        raw configs to a spec JSON.

Usage:
    # Dry-run on UC ATU
    python scripts/infer_nonconnected.py \
        --source "projects/J&J Rybrevant PET/Template/ZoomRx_UC_ATU_Report_Q1_'26.pptx" \
        --report output/ucatu_infer_report.json

    # Apply — write tagged PPTX + spec
    python scripts/infer_nonconnected.py \
        --source <src.pptx> \
        --output output/UC_ATU_inferred.pptx \
        --spec   output/UC_ATU_inferred_spec.json \
        --apply
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

# Windows cp1252 console can't encode Unicode box-drawing / arrows.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Windows cp1252 console can't encode Unicode like '>=', arrows, or box-draws.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from dotenv import load_dotenv
load_dotenv(override=True)

from slidegen.deck_reader import read_deck
from slidegen.intelligent_refresh import (
    _DEFAULT_LABEL_COL_PRIORITY,
    _normalize_label,
    fetch_synapse_data,
    propose_raw_configs,
    read_slide_context,
    save_proposed_configs_to_spec,
    write_connector_tags,
)


def harvest_tier1_universe(specs: list) -> tuple[set[int], int | None, int | None, int | None]:
    """Extract all unique Tier 1 analysis_ids from a deck + project metadata."""
    analysis_ids: set[int] = set()
    project_id = reporting_plan_id = survey_id = None
    for s in specs:
        dl = getattr(s, "data_lineage", None)
        if not dl:
            continue
        for aid in getattr(dl, "analysis_ids", []) or []:
            analysis_ids.add(aid)
        if project_id is None and getattr(dl, "project_id", None):
            project_id = dl.project_id
        if reporting_plan_id is None and getattr(dl, "reporting_plan_id", None):
            reporting_plan_id = dl.reporting_plan_id
        if survey_id is None and getattr(dl, "survey_id", None):
            survey_id = dl.survey_id
    return analysis_ids, project_id, reporting_plan_id, survey_id


def fetch_candidate_samples(
    analysis_ids: set[int],
    project_id: int,
    reporting_plan_id: int,
) -> dict[int, dict]:
    """Fetch dynamic_latest_n=1 sample for each candidate analysis_id.

    Returns dict: analysis_id -> {df, label_col, normalized_labels}
    """
    cache: dict[int, dict] = {}
    total = len(analysis_ids)
    for i, aid in enumerate(sorted(analysis_ids), 1):
        t0 = time.time()
        lineage = {
            "project_id": project_id,
            "reporting_plan_id": reporting_plan_id,
            "analysis_ids": [aid],
            "segment_ids": [],
            "dynamic_latest_n": 1,
        }
        try:
            _, df = fetch_synapse_data(lineage)
        except Exception as e:
            print(f"  [{i}/{total}] {aid}: FETCH ERROR — {type(e).__name__}: {e}")
            cache[aid] = {"df": None, "label_col": None, "normalized_labels": set(), "error": str(e)}
            continue

        label_col = next(
            (c for c in _DEFAULT_LABEL_COL_PRIORITY if c in df.columns),
            None,
        )
        if label_col is None or df.empty:
            cache[aid] = {
                "df": df,
                "label_col": None,
                "normalized_labels": set(),
                "raw_labels": [],
            }
        else:
            raw = df[label_col].dropna().astype(str).unique().tolist()
            normed = {_normalize_label(v) for v in raw}
            cache[aid] = {
                "df": df,
                "label_col": label_col,
                "normalized_labels": normed,
                "raw_labels": raw,
            }
        dt = time.time() - t0
        n_labels = len(cache[aid]["normalized_labels"])
        lc = cache[aid]["label_col"] or "?"
        print(f"  [{i}/{total}] {aid}: label_col={lc}  labels={n_labels}  ({dt:.1f}s)")
    return cache


def score_chart_against_candidates(
    chart_series: list[str],
    candidates: dict[int, dict],
) -> list[dict]:
    """Score one chart's series against every candidate; return sorted list."""
    chart_norm_map = {_normalize_label(s): s for s in chart_series}
    chart_norm = set(chart_norm_map)
    results = []
    for aid, info in candidates.items():
        api_norm = info.get("normalized_labels", set())
        if not api_norm or not chart_norm:
            continue
        matched_keys = chart_norm & api_norm
        score = len(matched_keys) / len(chart_norm)
        results.append({
            "analysis_id": aid,
            "score": round(score, 3),
            "matched": [chart_norm_map[k] for k in matched_keys],
            "label_col": info.get("label_col"),
        })
    results.sort(key=lambda r: r["score"], reverse=True)
    return results


def enumerate_tier2_charts(
    source_pptx: str,
    specs: list,
) -> list[dict]:
    """List every Tier 2 chart shape across the deck with series+categories."""
    tier2_charts = []
    for s in specs:
        if s.spec_completeness != "layout_complete_data_missing":
            continue
        try:
            ctx = read_slide_context(source_pptx, slide_index=s.slide_index)
        except Exception as e:
            print(f"  slide {s.slide_index}: read_slide_context error — {e}")
            continue
        for sh in ctx.get("shapes", []):
            if sh.get("type") != "chart":
                continue
            series = sh.get("series_names") or []
            if not series:
                continue
            tier2_charts.append({
                "slide_index": s.slide_index,
                "shape_name": sh.get("name"),
                "series_names": series,
                "categories": sh.get("categories", []),
            })
    return tier2_charts


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--source", required=True, help="Source PPTX path")
    ap.add_argument("--output", default=None, help="Output PPTX (apply mode)")
    ap.add_argument("--spec",   default=None, help="Output spec JSON (apply mode)")
    ap.add_argument("--report", default=None, help="Report JSON path (default: <output>_report.json)")
    ap.add_argument("--threshold", type=float, default=0.6, help="Min overlap score (default: 0.6)")
    ap.add_argument("--apply", action="store_true", help="Actually stamp tags + write spec (default: dry-run)")
    args = ap.parse_args()

    source_pptx = Path(args.source).resolve()
    if not source_pptx.exists():
        print(f"ERROR: source PPTX not found: {source_pptx}")
        sys.exit(1)

    print(f"=== Batch inferring — {source_pptx.name} ===")
    print(f"Mode: {'APPLY' if args.apply else 'dry-run'}    threshold: {args.threshold}")
    print()

    # ── 1. Read deck ──────────────────────────────────────────────────────
    print("[1/4] Reading deck...")
    specs, summary = read_deck(source_pptx)
    t1 = summary["tier1"]
    t2 = summary["tier2"]
    print(f"  slides: {summary['total_slides']}   "
          f"shapes: {summary['total_shapes_scanned']}")
    print(f"  tier 1 tagged: {t1.tagged_shapes}   "
          f"tier 2 inferred: charts={t2.charts_inferred} tables={t2.tables_inferred}")

    # ── 2. Harvest Tier 1 candidates ──────────────────────────────────────
    print("\n[2/4] Harvesting Tier 1 analysis_id universe...")
    analysis_ids, project_id, reporting_plan_id, survey_id = harvest_tier1_universe(specs)
    print(f"  unique analysis_ids: {len(analysis_ids)}")
    print(f"  project_id: {project_id}   reporting_plan_id: {reporting_plan_id}   survey_id: {survey_id}")
    if not analysis_ids:
        print("  No Tier 1 analysis_ids found — cannot infer. Exiting.")
        sys.exit(2)

    # ── 3. Fetch each candidate once + enumerate Tier 2 charts ───────────
    print(f"\n[3/4] Fetching samples for {len(analysis_ids)} candidates...")
    candidates = fetch_candidate_samples(analysis_ids, project_id, reporting_plan_id)

    print("\n      Enumerating Tier 2 charts...")
    tier2_charts = enumerate_tier2_charts(source_pptx, specs)
    print(f"      {len(tier2_charts)} tier-2 chart shapes to score")

    # ── 4. Score each tier-2 chart ────────────────────────────────────────
    print(f"\n[4/4] Scoring...")
    verdicts = []
    for ch in tier2_charts:
        ranked = score_chart_against_candidates(ch["series_names"], candidates)
        top = ranked[0] if ranked else None
        confirmed = bool(top and top["score"] >= args.threshold)
        ambiguous = bool(
            confirmed
            and len(ranked) > 1
            and ranked[1]["score"] >= args.threshold
            and ranked[1]["score"] >= top["score"] - 0.01
        )
        verdicts.append({
            "slide_index":  ch["slide_index"],
            "shape_name":   ch["shape_name"],
            "series_names": ch["series_names"],
            "confirmed":    confirmed,
            "ambiguous":    ambiguous,
            "top_match":    top,
            "runner_ups":   ranked[1:4],   # for manual review
        })

    # ── Summary ────────────────────────────────────────────────────────────
    total = len(verdicts)
    confirmed = sum(1 for v in verdicts if v["confirmed"])
    ambiguous = sum(1 for v in verdicts if v["ambiguous"])
    rejected = total - confirmed
    per_slide = Counter(v["slide_index"] for v in verdicts if v["confirmed"])

    print()
    print("-" * 60)
    print(f"  total tier-2 charts:   {total}")
    print(f"  confirmed (≥ {args.threshold}):      {confirmed}  ({confirmed/total*100 if total else 0:.0f}%)")
    print(f"  ambiguous (top ties):  {ambiguous}")
    print(f"  rejected:              {rejected}  ({rejected/total*100 if total else 0:.0f}%)")
    print(f"  slides with matches:   {len(per_slide)}")
    score_hist = Counter()
    for v in verdicts:
        if v["top_match"]:
            bucket = round(v["top_match"]["score"] * 10) / 10
            score_hist[bucket] += 1
    print(f"  top-score distribution: {dict(sorted(score_hist.items(), reverse=True))}")
    print("-" * 60)

    # ── Write report ──────────────────────────────────────────────────────
    report_path = Path(args.report) if args.report else Path(source_pptx.stem + "_infer_report.json").resolve()
    report = {
        "source_pptx": str(source_pptx),
        "threshold": args.threshold,
        "project_id": project_id,
        "reporting_plan_id": reporting_plan_id,
        "survey_id": survey_id,
        "tier1_analysis_ids": sorted(analysis_ids),
        "summary": {
            "total": total,
            "confirmed": confirmed,
            "ambiguous": ambiguous,
            "rejected": rejected,
            "slides_with_matches": len(per_slide),
        },
        "verdicts": verdicts,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    print(f"\nReport written to: {report_path}")

    # ── Apply mode: stamp tags + write spec ───────────────────────────────
    if not args.apply:
        print("\n(dry-run — no PPTX or spec modifications. Use --apply to stamp tags.)")
        return

    if not args.output or not args.spec:
        print("\nERROR: --apply requires both --output and --spec")
        sys.exit(3)

    print(f"\n=== APPLY MODE — stamping tags + writing spec ===")
    out_pptx = Path(args.output).resolve()
    out_spec = Path(args.spec).resolve()
    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    out_spec.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_pptx, out_pptx)
    print(f"  copied source → {out_pptx}")

    # Seed a minimal spec JSON so save_proposed_configs_to_spec has a home.
    # Uses the same shape as tests/CREON_1slide_test.json.
    initial_spec = {
        "source_deck": str(source_pptx.name),
        "data_sources": {},
        "slides": [
            {"slide_index": v["slide_index"], "components": []}
            for v in verdicts
            if v["confirmed"]
        ],
    }
    # Dedupe + sort slides
    seen = set()
    initial_spec["slides"] = [
        s for s in initial_spec["slides"]
        if not (s["slide_index"] in seen or seen.add(s["slide_index"]))
    ]
    initial_spec["slides"].sort(key=lambda s: s["slide_index"])
    out_spec.write_text(json.dumps(initial_spec, indent=2), encoding="utf-8")
    print(f"  seeded spec → {out_spec}")

    # Group confirmed verdicts by slide_index, then apply per-slide.
    by_slide: dict[int, list[dict]] = {}
    for v in verdicts:
        if v["confirmed"]:
            by_slide.setdefault(v["slide_index"], []).append(v)

    tagged_total = 0
    errors: list[str] = []
    for slide_idx, items in sorted(by_slide.items()):
        updates, tag_inputs = [], []
        for v in items:
            aid = v["top_match"]["analysis_id"]
            lineage = {
                "project_id": project_id,
                "reporting_plan_id": reporting_plan_id,
                "segment_ids": [],
                "dynamic_latest_n": 5,
            }
            # Re-fetch with a larger window to derive raw_configs
            _, df = fetch_synapse_data({**lineage, "analysis_ids": [aid]})
            ctx = read_slide_context(str(out_pptx), slide_index=slide_idx)
            chart_shape = next(
                (s for s in ctx["shapes"] if s.get("name") == v["shape_name"]),
                None,
            )
            if chart_shape is None:
                errors.append(f"slide {slide_idx} / {v['shape_name']}: shape not found")
                continue
            proposal = propose_raw_configs(chart_shape, df)
            if proposal.get("error"):
                errors.append(f"slide {slide_idx} / {v['shape_name']}: propose: {proposal['error']}")
                continue
            updates.append({
                "shape_name":         v["shape_name"],
                "raw_pivot_config":   proposal["raw_pivot_config"],
                "raw_mapping_config": proposal["raw_mapping_config"],
                "analysis_id":        aid,
                "data_lineage":       lineage,
            })
            tag_inputs.append({
                "shape_name":         v["shape_name"],
                "raw_pivot_config":   proposal["raw_pivot_config"],
                "raw_mapping_config": proposal["raw_mapping_config"],
                "analysis_id":        aid,
            })
        if updates:
            save_proposed_configs_to_spec(str(out_spec), slide_idx, updates)
        if tag_inputs:
            tag_result = write_connector_tags(
                str(out_pptx), slide_idx, tag_inputs,
                {"project_id": project_id, "reporting_plan_id": reporting_plan_id,
                 "segment_ids": [], "dynamic_latest_n": 5},
                survey_id=survey_id,
            )
            tagged_total += len(tag_result)
            print(f"  slide {slide_idx}: stamped {len(tag_result)} tags")

    print()
    print(f"  total tags stamped: {tagged_total}")
    if errors:
        print(f"  errors ({len(errors)}):")
        for e in errors[:10]:
            print(f"    {e}")
        if len(errors) > 10:
            print(f"    ... and {len(errors) - 10} more")
    print()
    print(f"Output PPTX: {out_pptx}")
    print(f"Output spec: {out_spec}")


if __name__ == "__main__":
    main()
