"""Full CREON refresh pipeline — single command end-to-end.

Refreshes tests/Output_CREON.pptx using tests/CREON.json against current
Synapse API data, then runs in sequence:
  1. refresh_deck_from_spec  -> chart values + table cell_values
  2. stamp_refresh_notes     -> per-shape REFRESH_NOTE Connector tags
  3. refresh_headlines       -> LLM-driven post-refresh narrative rewrite
  4. annotate slide badges   -> REFRESH + HEADLINE rounded-rect badges

Output (everything under output_testing/, deck name + date for ID):
    output_testing/deck_output/CREON_{YYYY-MM-DD}.pptx
    output_testing/json_output/CREON_{YYYY-MM-DD}_refresh_status.json
    output_testing/json_output/CREON_{YYYY-MM-DD}_headline_status.json

Loads Synapse creds from the shared env file at:
    ../galen-consulting-r3m-report/synapse env/.env

Run from repo root:
    python scripts/refresh_creon_today.py
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from dotenv import load_dotenv  # noqa: E402

# Load LLM creds from local .env, Synapse creds from the shared file
load_dotenv(REPO_ROOT / ".env", override=False)
_DOCS_ENV = REPO_ROOT / "docs" / ".env"
if _DOCS_ENV.exists():
    load_dotenv(_DOCS_ENV, override=False)
_SHARED_SYNAPSE_ENV = REPO_ROOT / "synapse env" / ".env"
if _SHARED_SYNAPSE_ENV.exists():
    load_dotenv(_SHARED_SYNAPSE_ENV, override=False)

from slidegen.intelligent_refresh import (  # noqa: E402
    refresh_deck_from_spec,
    stamp_refresh_notes,
)


def _regen_spec(source_pptx: Path, out_spec: Path) -> None:
    """Regenerate full spec from a deck via tag_reader → build_full_spec."""
    import json
    from slidegen.deck_reader.tag_reader import generate_config_specs
    from slidegen.slide_spec.schema import dump_spec
    from tests.build_full_spec import build_data_sources, build_connected_slide, _lineage_ds_key
    from pptx import Presentation

    specs, summary = generate_config_specs(str(source_pptx))
    connector_specs = [json.loads(dump_spec(s)) for s in specs]
    data_sources = build_data_sources(connector_specs)
    slide_to_ds = {
        s["slide_index"]: _lineage_ds_key(s.get("data_lineage", {}))
        for s in connector_specs
        if _lineage_ds_key(s.get("data_lineage", {}))
    }

    prs = Presentation(str(source_pptx))
    slide_entries = []
    for spec in connector_specs:
        si = spec["slide_index"]
        ds_key = slide_to_ds.get(si)
        if not ds_key:
            continue
        slide_entry = build_connected_slide(spec, ds_key)
        if si < len(prs.slides):
            pptx_slide = prs.slides[si]
            chart_shapes = sorted(
                [s for s in pptx_slide.shapes if s.has_chart],
                key=lambda x: (x.left or 0, x.top or 0),
            )
            table_shapes = sorted(
                [s for s in pptx_slide.shapes if s.has_table],
                key=lambda x: (x.left or 0, x.top or 0),
            )
            for comp in slide_entry["components"]:
                pos = comp.get("position", {})
                cl = pos.get("left", 0); ct = pos.get("top", 0)
                cands = chart_shapes if comp["type"] == "chart" else table_shapes
                for shape in cands:
                    sl_pos = round(shape.left / 914400, 2) if shape.left else 0
                    st_pos = round(shape.top / 914400, 2) if shape.top else 0
                    if abs(sl_pos - cl) < 0.3 and abs(st_pos - ct) < 0.3:
                        comp["name"] = shape.name
                        break
        slide_entries.append(slide_entry)

    full_spec = {
        "source_deck": str(source_pptx),
        "data_sources": data_sources,
        "slides": slide_entries,
    }
    out_spec.write_text(
        json.dumps(full_spec, indent=2, ensure_ascii=False), encoding="utf-8",
    )


def _check_synapse_creds() -> tuple[bool, str]:
    if not os.getenv("SYNAPSE_API_URL"):
        return False, "SYNAPSE_API_URL not set in .env"
    if not os.getenv("SYNAPSE_API_KEY"):
        # synapse-cli has a JWT/Azure-AD fallback chain; allow it
        # but warn so the run isn't silently using a stale cached token.
        return True, "SYNAPSE_API_KEY not in .env — will try synapse-cli fallback chain"
    return True, "ok"


def main():
    today = datetime.now().strftime("%Y-%m-%d")
    print(f"=== CREON refresh: {today} ===\n")

    # ── Pre-flight ──
    creds_ok, creds_msg = _check_synapse_creds()
    if not creds_ok:
        print(f"[ERROR] {creds_msg}")
        print("        Add SYNAPSE_API_URL and SYNAPSE_API_KEY to .env, or run")
        print("        `synapse login` to populate the JWT cache.")
        sys.exit(2)
    if "fallback" in creds_msg:
        print(f"[warn] {creds_msg}")
    if not os.getenv("LLM_API_KEY"):
        print("[warn] LLM_API_KEY not set — Step 7 headlines will skip")

    # Production CREON deck (the real W33 deliverable, not the test copy).
    source_pptx = (REPO_ROOT / "output_testing" / "deck_output"
                   / "CREON Share of Voice Study - W33 updated source deck.pptx")
    if not source_pptx.exists():
        print(f"[ERROR] Source deck missing: {source_pptx}")
        sys.exit(2)

    # Regenerate the spec from this deck so static_time_period_ids,
    # include_live_wave overrides, and split-viz positions are all
    # extracted from the current deck (not from a stale tests/CREON.json).
    out_root = REPO_ROOT / "output_testing"
    json_out_dir = out_root / "json_output"
    json_out_dir.mkdir(parents=True, exist_ok=True)
    spec_path = json_out_dir / f"CREON_{datetime.now().strftime('%Y-%m-%d')}_v4_full_spec.json"

    print(f"  Regenerating spec from production deck...")
    _regen_spec(source_pptx, spec_path)
    print(f"  spec -> {spec_path.relative_to(REPO_ROOT)}")

    # ── Output paths ──
    out_root = REPO_ROOT / "output_testing"
    deck_out_dir = out_root / "deck_output"
    json_out_dir = out_root / "json_output"
    deck_out_dir.mkdir(parents=True, exist_ok=True)
    json_out_dir.mkdir(parents=True, exist_ok=True)

    base = f"CREON_{today}_v8"
    final_pptx = deck_out_dir / f"{base}.pptx"
    refresh_status_path = json_out_dir / f"{base}_refresh_status.json"
    headline_status_path = json_out_dir / f"{base}_headline_status.json"
    label_sync_status_path = json_out_dir / f"{base}_label_sync_status.json"

    # Intermediates stay in a temp dir — we only ship the final annotated deck.
    work_tmp = tempfile.mkdtemp(prefix="creon_refresh_")
    work_dir = Path(work_tmp)
    refreshed_pptx = work_dir / f"{base}_step1_refresh.pptx"

    print(f"  source:     {source_pptx.relative_to(REPO_ROOT)}")
    print(f"  spec:       {spec_path.relative_to(REPO_ROOT)}")
    print(f"  final deck: {final_pptx.relative_to(REPO_ROOT)}")
    print(f"  json out:   {json_out_dir.relative_to(REPO_ROOT)}")
    print()

    # ── Step 1: Refresh ──
    print("[1/5] refresh_deck_from_spec...")
    result = refresh_deck_from_spec(
        spec_path=str(spec_path),
        pptx_path=str(source_pptx),
        output_path=str(refreshed_pptx),
    )
    refresh_status_path.write_text(
        json.dumps(result, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    n_slides = len(result.get("slides", []))
    n_ok = sum(
        1 for s in result.get("slides", [])
        for c in (s.get("charts") or []) + (s.get("tables") or [])
        if c.get("status") == "ok"
    )
    n_with_notes = sum(
        1 for s in result.get("slides", [])
        for c in (s.get("charts") or []) + (s.get("tables") or [])
        if c.get("notes")
    )
    print(f"      {n_slides} slides processed, {n_ok} components ok, {n_with_notes} with notes")
    print(f"      status -> {refresh_status_path.relative_to(REPO_ROOT)}")

    # ── Step 2: Stamp REFRESH_NOTE Connector tags ──
    print("\n[2/5] stamp_refresh_notes (per-shape Connector tags)...")
    stamp_results = stamp_refresh_notes(str(refreshed_pptx), result)
    n_stamped = sum(1 for v in stamp_results.values() if v == "stamped")
    n_skipped = sum(1 for v in stamp_results.values() if v != "stamped")
    print(f"      {n_stamped} shapes stamped, {n_skipped} skipped")

    # ── Step 3: Headlines ──
    if os.getenv("LLM_API_KEY"):
        print("\n[3/5] refresh_headlines (LLM rewrite)...")
        from slidegen.headline_refresh import refresh_headlines
        try:
            with_headlines_pptx = work_dir / f"{base}_step3_headlines.pptx"
            updates = refresh_headlines(
                source_pptx=str(source_pptx),
                refreshed_pptx=str(refreshed_pptx),
                spec_path=str(spec_path),
                out_pptx=str(with_headlines_pptx),
            )
            # refresh_headlines returns list[HeadlineUpdate] dataclasses;
            # serialize to dict shape that the annotator + contract eval expect.
            updates_payload = {
                "updates": [
                    {"slide_idx": u.slide_idx, "old_headline": u.old_headline,
                     "new_headline": u.new_headline, "chart_summary": u.chart_summary,
                     "status": u.status}
                    for u in updates
                ]
            }
            headline_status_path.write_text(
                json.dumps(updates_payload, indent=2, ensure_ascii=False, default=str),
                encoding="utf-8",
            )
            n_updates = len(updates)
            n_updated = sum(1 for u in updates if u.status == "updated")
            print(f"      {n_updates} slides analyzed, {n_updated} headlines rewritten")
            print(f"      status -> {headline_status_path.relative_to(REPO_ROOT)}")
            after_headlines = with_headlines_pptx
        except Exception as exc:
            print(f"      [warn] headline rewrite failed: {exc}")
            after_headlines = refreshed_pptx
    else:
        print("\n[3/5] skipped (LLM_API_KEY not set)")
        after_headlines = refreshed_pptx

    # ── Step 4: Sync period labels in adjacent text frames ──
    print("\n[4/5] sync_period_labels (text frames adjacent to charts)...")
    from slidegen.label_sync import sync_period_labels
    after_label_sync = work_dir / f"{base}_step4_label_sync.pptx"
    label_sync_diag = sync_period_labels(
        source_pptx=str(source_pptx),
        refreshed_pptx=str(after_headlines),
        out_pptx=str(after_label_sync),
        sidecar_path=str(label_sync_status_path),
    )
    n_shifted = label_sync_diag["totals"]["slides_with_shift"]
    n_edits = label_sync_diag["totals"]["edits"]
    n_conflicts = label_sync_diag["totals"]["conflicts"]
    print(f"      {n_shifted} slides shifted, {n_edits} run-level edits, "
          f"{n_conflicts} conflicts")
    print(f"      status -> {label_sync_status_path.relative_to(REPO_ROOT)}")
    after_headlines = after_label_sync  # feeds the badge step below

    # ── Step 5: Annotate badges ──
    print("\n[5/5] slide-level annotation badges...")
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    sys.path.insert(0, str(REPO_ROOT / "scripts"))
    from annotate_refresh_and_headline import (  # noqa: E402
        _refresh_badge, _headline_badge, _add_badge,
        GREEN, BLUE, AMBER, RED, GREY, DARK,
    )

    refresh_status = json.loads(refresh_status_path.read_text(encoding="utf-8"))
    refresh_by_idx = {s["slide_index"]: s for s in refresh_status.get("slides", [])}

    headline_by_idx: dict[int, dict] = {}
    if headline_status_path.exists():
        hl_status = json.loads(headline_status_path.read_text(encoding="utf-8"))
        headline_by_idx = {u["slide_idx"]: u for u in hl_status.get("updates", [])}

    pres = Presentation(str(after_headlines))
    slide_w = pres.slide_width

    counts = {"non_connected": 0, "tag_mismatch": 0, "static_review": 0,
              "ok": 0, "ok_with_partial": 0, "ok_with_dynamic_added": 0,
              "alignment_failed": 0, "no_data": 0, "other": 0}

    for s_idx, slide in enumerate(pres.slides):
        rs = refresh_by_idx.get(s_idx)
        hl = headline_by_idx.get(s_idx)
        r_color, r_label = _refresh_badge(rs)
        h_color, h_label = _headline_badge(hl)
        _add_badge(slide, r_color, r_label, slide_w, row=0)
        _add_badge(slide, h_color, h_label, slide_w, row=1)

        # Bucket counts using the new badge labels
        if "NON-CONNECTED" in r_label:
            counts["non_connected"] += 1
        elif "tag-mismatch" in r_label:
            counts["tag_mismatch"] += 1
        elif "STATIC" in r_label or "static-pinned" in r_label:
            counts["static_review"] += 1
        elif "alignment-failed" in r_label:
            counts["alignment_failed"] += 1
        elif "(partial)" in r_label:
            counts["ok_with_partial"] += 1
        elif "dynamic-added" in r_label:
            counts["ok_with_dynamic_added"] += 1
        elif r_label.startswith("REFRESH ok"):
            counts["ok"] += 1
        elif "no new data" in r_label:
            counts["no_data"] += 1
        else:
            counts["other"] += 1

    pres.save(str(final_pptx))

    # Clean up intermediate decks; we only keep the final annotated one.
    try:
        shutil.rmtree(work_dir)
    except Exception:
        pass

    print(f"      annotated -> {final_pptx.relative_to(REPO_ROOT)}")
    print()
    print("=== Summary ===")
    print(f"  output deck: {final_pptx.relative_to(REPO_ROOT)}")
    print(f"  size:        {final_pptx.stat().st_size / 1024 / 1024:.1f} MB")
    print(f"  refresh JSON: {refresh_status_path.relative_to(REPO_ROOT)}")
    if headline_status_path.exists():
        print(f"  headline JSON: {headline_status_path.relative_to(REPO_ROOT)}")
    print(f"  slide buckets:")
    for k, v in counts.items():
        if v:
            print(f"    {k:25s} {v}")

    # Surface tag_mismatch slides for user review (Issue #4)
    tag_mismatch_slides = [
        s_idx for s_idx, _ in enumerate(pres.slides)
        if any(
            n.get("kind") == "tag_mismatch"
            for c in (refresh_by_idx.get(s_idx, {}).get("charts") or [])
            + (refresh_by_idx.get(s_idx, {}).get("tables") or [])
            for n in (c.get("notes") or [])
        )
    ]
    if tag_mismatch_slides:
        print()
        print(f"  [REVIEW NEEDED] {len(tag_mismatch_slides)} slide(s) with tag_mismatch:")
        for si in tag_mismatch_slides[:20]:
            shapes = [
                f"{c.get('name')} ({n.get('detail', '')[:80]})"
                for c in (refresh_by_idx.get(si, {}).get("charts") or [])
                for n in (c.get("notes") or [])
                if n.get("kind") == "tag_mismatch"
            ]
            print(f"    slide {si}: {shapes}")
        if len(tag_mismatch_slides) > 20:
            print(f"    ... and {len(tag_mismatch_slides) - 20} more")
        print(f"  Per slide: fix the Connector tag, or move to non-connected path.")


if __name__ == "__main__":
    main()
