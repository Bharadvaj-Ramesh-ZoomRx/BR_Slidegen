"""Build the full spec.json from Connector config specs + manual specs.

For connected slides: includes raw_pivot_config + raw_mapping_config per component
(used by pivot_records_to_chart_data for exact Connector fidelity).

For non-connected slides: includes interpreted data_mapping
(used by the intelligent refresh mapping engine).

Both live in the same spec.json structure. The refresh engine checks for
raw configs first, falls back to mapping format.
"""
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

TESTS_DIR = REPO_ROOT / "tests"
CONNECTOR_SPECS = TESTS_DIR / "UAT_deck_config_specs.json"
MANUAL_SPEC = TESTS_DIR / "[Vijay] Synapse Connector UAT - Mar 2026.json"
OUTPUT_SPEC = TESTS_DIR / "[Vijay] Synapse Connector UAT - Mar 2026.json"

# Manual specs for non-connected slides (already hand-crafted and working)
MANUAL_SLIDE_INDICES = {1, 4, 7}


def build_data_sources(connector_specs: list[dict]) -> dict:
    """Extract unique data sources from Connector specs."""
    sources = {}
    for spec in connector_specs:
        lin = spec.get("data_lineage", {})
        pid = lin.get("project_id")
        if not pid:
            continue
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
            }
    return sources


def build_connected_slide(spec: dict, ds_key: str) -> dict:
    """Build a spec.json slide entry from a Connector config spec.

    Includes raw_pivot_config + raw_mapping_config for each component.
    """
    components = []
    for comp in spec.get("components", []):
        ctype = comp.get("type")
        if ctype not in ("chart", "value_table", "label_table"):
            continue

        dm = comp.get("data_mapping")
        if not dm:
            continue

        entry = {
            "type": ctype,
            "name": "",  # filled from PPTX shape name
            "position": comp.get("position", {}),
        }

        if dm.get("raw_pivot_config") and dm.get("raw_mapping_config"):
            # Connected mode: store raw configs for pivot_records_to_chart_data()
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

        # Also include the human-readable mapping for reference
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

    entry = {
        "slide_index": spec["slide_index"],
        "slide_id": spec.get("slide_id", f"slide_{spec['slide_index']:03d}"),
        "data_source": ds_key,
        "components": components,
    }

    # Include static_time_period_names if present (controls which waves are used)
    lin = spec.get("data_lineage", {})
    stn = lin.get("static_time_period_names", [])
    if stn:
        entry["static_time_period_names"] = stn

    return entry


def main():
    from pptx import Presentation

    # Load existing manual spec for non-connected slides
    manual_spec = json.loads(MANUAL_SPEC.read_text(encoding="utf-8"))
    manual_slides = {s["slide_index"]: s for s in manual_spec.get("slides", [])
                     if s["slide_index"] in MANUAL_SLIDE_INDICES}

    # Load Connector specs
    connector_specs = json.loads(CONNECTOR_SPECS.read_text(encoding="utf-8"))

    # Build data sources
    sources = build_data_sources(connector_specs)

    # Map slide_index -> data_source key
    slide_to_ds = {}
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

    # Read PPTX for shape names (Connector specs use position, we need names)
    pptx_path = str(TESTS_DIR / "[Vijay] Synapse Connector UAT - Mar 2026.pptx")
    prs = Presentation(pptx_path)

    # Build full spec
    slides = []
    for spec in connector_specs:
        si = spec["slide_index"]

        if si in manual_slides:
            # Use hand-crafted spec for non-connected slides
            slides.append(manual_slides[si])
            continue

        ds_key = slide_to_ds.get(si)
        if not ds_key:
            continue

        slide_entry = build_connected_slide(spec, ds_key)

        # Fill in shape names from PPTX
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

            # Match components to shapes by position
            for comp in slide_entry["components"]:
                pos = comp.get("position", {})
                cl = pos.get("left", 0)
                ct = pos.get("top", 0)

                if comp["type"] == "chart":
                    for shape in chart_shapes:
                        sl = round(shape.left / 914400, 2) if shape.left else 0
                        st = round(shape.top / 914400, 2) if shape.top else 0
                        if abs(sl - cl) < 0.3 and abs(st - ct) < 0.3:
                            comp["name"] = shape.name
                            break
                else:
                    for shape in table_shapes:
                        sl = round(shape.left / 914400, 2) if shape.left else 0
                        st = round(shape.top / 914400, 2) if shape.top else 0
                        if abs(sl - cl) < 0.3 and abs(st - ct) < 0.3:
                            comp["name"] = shape.name
                            break

        slides.append(slide_entry)

    # Assemble final spec
    final = {
        "source_deck": "[Vijay] Synapse Connector UAT - Mar 2026.pptx",
        "data_sources": {**sources, **manual_spec.get("data_sources", {})},
        "slides": slides,
    }

    OUTPUT_SPEC.write_text(json.dumps(final, indent=2, ensure_ascii=False), encoding="utf-8")

    # Summary
    n_connected = sum(1 for s in slides if any(
        c.get("raw_pivot_config") for c in s.get("components", [])))
    n_manual = sum(1 for s in slides if s["slide_index"] in MANUAL_SLIDE_INDICES)
    n_raw_comps = sum(
        1 for s in slides for c in s.get("components", [])
        if c.get("raw_pivot_config"))

    print(f"Spec built: {len(slides)} slides")
    print(f"  Connected (raw configs): {n_connected}")
    print(f"  Manual (interpreted): {n_manual}")
    print(f"  Components with raw configs: {n_raw_comps}")
    print(f"  Data sources: {len(final['data_sources'])}")
    print(f"  Output: {OUTPUT_SPEC}")


if __name__ == "__main__":
    main()
