"""Build the full spec.json for intelligent_refresh from Connector config specs.

For connected slides: includes raw_pivot_config + raw_mapping_config per component
(used by pivot_records_to_chart_data for exact Connector fidelity).

For non-connected slides: includes interpreted data_mapping
(used by the intelligent refresh mapping engine).

Usage:
    # CREON (default)
    python tests/build_full_spec.py

    # Any deck
    python tests/build_full_spec.py --specs tests/CREON_deck_config_specs.json \
        --pptx projects/CREON/CREON.pptx \
        --output tests/CREON.json

    # UAT deck (Vijay)
    python tests/build_full_spec.py \
        --specs "tests/UAT_deck_config_specs.json" \
        --pptx "tests/[Vijay] Synapse Connector UAT - Mar 2026.pptx" \
        --output "tests/[Vijay] Synapse Connector UAT - Mar 2026.json" \
        --manual-spec "tests/[Vijay] Synapse Connector UAT - Mar 2026_manual.json" \
        --manual-slides 1,4,7
"""
import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

# ── Defaults (CREON) ──────────────────────────────────────────────────────────
TESTS_DIR = REPO_ROOT / "tests"
DEFAULT_SPECS   = TESTS_DIR / "CREON_deck_config_specs.json"
DEFAULT_PPTX    = REPO_ROOT / "projects" / "CREON" / "CREON.pptx"
DEFAULT_OUTPUT  = TESTS_DIR / "CREON.json"


# ══════════════════════════════════════════════════════════════════════════════

def _datasource_from_lineage(lin: dict, sources: dict) -> str | None:
    """Add a data_source entry for one lineage dict and return its key.

    Handles union of static_time_period_ids/names across calls that hit the
    same key (multiple slides or components can share an analysis).
    """
    pid = lin.get("project_id")
    if not pid:
        return None
    rpid = lin.get("reporting_plan_id")
    aids = tuple(lin.get("analysis_ids", []))
    sids = tuple(s.get("rule_id", 0) for s in lin.get("segments", []))
    if not sids:
        sids = tuple(lin.get("segment_ids", []))

    key = f"p{pid}_rp{rpid}_a{'_'.join(str(a) for a in aids)}"
    if key not in sources:
        sources[key] = {
            "project_id": pid,
            "reporting_plan_id": rpid,
            "analysis_ids": list(aids),
            "segment_ids": list(sids),
            "dynamic_latest_n": lin.get("dynamic_latest_n") or 5,
            "static_time_period_ids": [],
            "static_time_period_names": [],
            "include_live_wave": lin.get("include_live_wave"),
        }
    # Union static IDs/names across all lineages sharing this key.
    existing_ids = set(sources[key]["static_time_period_ids"])
    new_ids = set(lin.get("static_time_period_ids") or [])
    sources[key]["static_time_period_ids"] = sorted(existing_ids | new_ids)
    existing_names = set(sources[key]["static_time_period_names"])
    new_names = set(lin.get("static_time_period_names") or [])
    sources[key]["static_time_period_names"] = sorted(existing_names | new_names)
    if lin.get("include_live_wave") is True:
        sources[key]["include_live_wave"] = True
    return key


def build_data_sources(connector_specs: list[dict]) -> dict:
    """Extract unique data sources from Connector specs.

    Walks both slide-level data_lineage AND per-component raw_data_lineage
    so every analysis any chart/table needs becomes its own data_source.
    """
    sources: dict = {}
    for spec in connector_specs:
        slide_lin = spec.get("data_lineage", {})
        _datasource_from_lineage(slide_lin, sources)
        # Component-level lineages (e.g. a slide with multiple charts pointing
        # at different analyses — each chart's tag has its own ReportConfig).
        for comp in spec.get("components", []):
            dm = comp.get("data_mapping") or {}
            comp_lin = dm.get("raw_data_lineage")
            if comp_lin:
                _datasource_from_lineage(comp_lin, sources)
    return sources


def _lineage_ds_key(lin: dict) -> str | None:
    pid = lin.get("project_id")
    if not pid:
        return None
    rpid = lin.get("reporting_plan_id")
    aids = lin.get("analysis_ids", [])
    return f"p{pid}_rp{rpid}_a{'_'.join(str(a) for a in aids)}"


def build_connected_slide(spec: dict, ds_key: str) -> dict:
    """Build a spec.json slide entry from a Connector config spec.

    When a component's raw_data_lineage has a different analysis set than the
    slide default, the component gets its own data_source key — so refresh
    fetches the right analysis for that specific chart/table.
    """
    components = []
    slide_lin = spec.get("data_lineage", {})
    slide_key = _lineage_ds_key(slide_lin)
    for comp in spec.get("components", []):
        ctype = comp.get("type")
        if ctype not in ("chart", "value_table", "label_table"):
            continue

        dm = comp.get("data_mapping")
        if not dm:
            continue

        entry = {
            "type": ctype,
            "name": "",  # filled from PPTX shape names below
            "position": comp.get("position", {}),
        }

        # Per-component data_source override — only when analyses diverge from slide
        comp_lin = dm.get("raw_data_lineage")
        if comp_lin:
            comp_key = _lineage_ds_key(comp_lin)
            if comp_key and comp_key != slide_key:
                entry["data_source"] = comp_key
            # Carry slide-level static IDs from the component's own lineage
            # (chart tag may have its own static_time_period_ids independent
            # of the slide's default lineage).
            cs_ids = comp_lin.get("static_time_period_ids")
            if cs_ids:
                entry["static_time_period_ids"] = list(cs_ids)
            cs_names = comp_lin.get("static_time_period_names")
            if cs_names:
                entry["static_time_period_names"] = list(cs_names)

        if dm.get("raw_pivot_config") and dm.get("raw_mapping_config"):
            entry["raw_pivot_config"] = dm["raw_pivot_config"]
            entry["raw_mapping_config"] = dm["raw_mapping_config"]
            if dm.get("raw_column_key_label_map"):
                entry["raw_column_key_label_map"] = dm["raw_column_key_label_map"]
            if dm.get("split_order") is not None:
                entry["split_order"] = dm["split_order"]
            if dm.get("rows_per_object"):
                entry["rows_per_object"] = dm["rows_per_object"]
            if dm.get("top_n_rows"):
                entry["top_n_rows"] = dm["top_n_rows"]

        # Human-readable mapping for reference / fallback
        tx = dm.get("transform", {})
        if tx:
            entry["data_mapping"] = {
                "row_field": tx.get("row_field", ""),
                "column_field": tx.get("column_field"),
                "value_field": tx.get("value_field", ""),
            }

        if ctype == "chart":
            entry["chart_pattern"] = comp.get("chart_pattern", "")

        components.append(entry)

    slide_entry = {
        "slide_index": spec["slide_index"],
        "slide_id": spec.get("slide_id", f"slide_{spec['slide_index']:03d}"),
        "data_source": ds_key,
        "components": components,
    }

    lin = spec.get("data_lineage", {})
    stn = lin.get("static_time_period_names", [])
    if stn:
        slide_entry["static_time_period_names"] = stn
    sti = lin.get("static_time_period_ids", [])
    if sti:
        slide_entry["static_time_period_ids"] = list(sti)

    return slide_entry


def main():
    parser = argparse.ArgumentParser(description="Build spec.json for intelligent_refresh")
    parser.add_argument("--specs",        default=str(DEFAULT_SPECS),
                        help="Connector config specs JSON (default: CREON_deck_config_specs.json)")
    parser.add_argument("--pptx",         default=str(DEFAULT_PPTX),
                        help="Source PPTX (used to resolve shape names)")
    parser.add_argument("--output",       default=str(DEFAULT_OUTPUT),
                        help="Output spec JSON path (default: tests/CREON.json)")
    parser.add_argument("--manual-spec",  default=None,
                        help="Optional JSON with hand-crafted specs for non-connected slides")
    parser.add_argument("--manual-slides", default="",
                        help="Comma-separated slide indices to take from manual-spec (e.g. 1,4,7)")
    args = parser.parse_args()

    pptx_path   = Path(args.pptx)
    output_path = Path(args.output)
    specs_path  = Path(args.specs)

    assert pptx_path.exists(), f"PPTX not found: {pptx_path}"

    # ── Load connector specs ─────────────────────────────────────────────
    # Prefer: read tags directly from the PPTX via generate_config_specs().
    # Fallback: load from a pre-generated JSON (legacy path / offline mode).
    if not specs_path.exists():
        print(f"Specs JSON not found — reading tags directly from PPTX ...")
        from slidegen.deck_reader.tag_reader import generate_config_specs
        from slidegen.slide_spec.schema import dump_spec
        specs_list, summary = generate_config_specs(str(pptx_path))
        connector_specs = [json.loads(dump_spec(s)) for s in specs_list]
        print(f"  {len(connector_specs)} slide specs extracted  "
              f"(tagged={summary.tagged_shapes}, untagged={summary.untagged_shapes})")
    else:
        connector_specs = json.loads(specs_path.read_text(encoding="utf-8"))

    # Optional manual specs for non-connected slides
    manual_slides: dict[int, dict] = {}
    manual_indices: set[int] = set()
    if args.manual_spec and Path(args.manual_spec).exists():
        manual_data = json.loads(Path(args.manual_spec).read_text(encoding="utf-8"))
        if args.manual_slides:
            manual_indices = {int(x) for x in args.manual_slides.split(",") if x.strip()}
        manual_slides = {s["slide_index"]: s for s in manual_data.get("slides", [])
                         if s["slide_index"] in manual_indices}

    # Build data sources index
    sources = build_data_sources(connector_specs)

    # Map slide_index → data_source key
    slide_to_ds: dict[int, str] = {}
    for spec in connector_specs:
        si = spec["slide_index"]
        lin = spec.get("data_lineage", {})
        pid = lin.get("project_id")
        if not pid:
            continue
        rpid = lin.get("reporting_plan_id")
        aids = tuple(lin.get("analysis_ids", []))
        key = f"p{pid}_rp{rpid}_a{'_'.join(str(a) for a in aids)}"
        slide_to_ds[si] = key

    # Load PPTX once to resolve shape names from positions
    from pptx import Presentation
    prs = Presentation(str(pptx_path))

    slides = []
    for spec in connector_specs:
        si = spec["slide_index"]

        if si in manual_slides:
            slides.append(manual_slides[si])
            continue

        ds_key = slide_to_ds.get(si)
        if not ds_key:
            continue

        slide_entry = build_connected_slide(spec, ds_key)

        # Fill in shape names from PPTX (intelligent_refresh matches by name)
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
                cl, ct = pos.get("left", 0), pos.get("top", 0)
                pool = chart_shapes if comp["type"] == "chart" else table_shapes
                for shape in pool:
                    sl = round(shape.left / 914400, 2) if shape.left else 0
                    st = round(shape.top / 914400, 2) if shape.top else 0
                    if abs(sl - cl) < 0.3 and abs(st - ct) < 0.3:
                        comp["name"] = shape.name
                        break

        slides.append(slide_entry)

    # source_deck path relative to output spec's directory
    rel_pptx = pptx_path.resolve().relative_to(output_path.parent.resolve()) \
        if pptx_path.resolve().is_relative_to(output_path.parent.resolve()) \
        else pptx_path.resolve()

    final = {
        "source_deck": str(rel_pptx).replace("\\", "/"),
        "data_sources": sources,
        "slides": slides,
    }

    output_path.write_text(
        json.dumps(final, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    # Summary
    n_connected = sum(1 for s in slides
                      if any(c.get("raw_pivot_config") for c in s.get("components", [])))
    n_manual = len(manual_slides)
    n_raw_comps = sum(1 for s in slides
                      for c in s.get("components", []) if c.get("raw_pivot_config"))
    n_unnamed = sum(1 for s in slides
                    for c in s.get("components", []) if not c.get("name"))

    print(f"Spec built: {len(slides)} slides")
    print(f"  Connected (raw configs): {n_connected}")
    print(f"  Manual (interpreted):    {n_manual}")
    print(f"  Components with raw configs: {n_raw_comps}")
    print(f"  Unnamed components (no shape match): {n_unnamed}")
    print(f"  Data sources: {len(sources)}")
    print(f"  Output: {output_path}")


if __name__ == "__main__":
    main()
