"""
Layout Clusterer
================

Relaxes the shape signature to focus on the core content types (chart/table/picture)
and ignores decorative elements (auto_shape, line, freeform) that vary per client.

Also extracts position+size *per chart* and *per table* across all decks to identify
common grid positions that become LAYOUTS{} entries.

Run:
    cd experiments/deck_analysis
    python layout_clusterer.py

Outputs:
    outputs/layout_clusters.json   — relaxed signature coordinate stats
    outputs/layout_clusters.md     — readable
    outputs/chart_positions.json   — raw position/size data per chart
    outputs/table_positions.json   — raw position/size data per table
    outputs/position_clusters.md   — clusters of similar chart/table positions
"""
from __future__ import annotations

import json
import sys
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean, median, stdev
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

HERE = Path(__file__).parent
DECKS_DIR = HERE / "decks"
OUTPUTS_DIR = HERE / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

EMU = 914400.0


def emu_in(v) -> float | None:
    if v is None:
        return None
    return round(v / EMU, 2)


# ---------------------------------------------------------------------------
# Shape walker
# ---------------------------------------------------------------------------


def shape_bbox(shape) -> dict | None:
    try:
        return {
            "type": shape_type_name(shape),
            "left": emu_in(shape.left),
            "top": emu_in(shape.top),
            "width": emu_in(shape.width),
            "height": emu_in(shape.height),
            "has_chart": shape.has_chart,
            "has_table": shape.has_table,
        }
    except Exception:
        return None


def shape_type_name(shape) -> str:
    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.CHART:
        return "chart"
    if st == MSO_SHAPE_TYPE.TABLE:
        return "table"
    if st == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if st == MSO_SHAPE_TYPE.TEXT_BOX:
        return "text_box"
    if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
        return "auto_shape"
    if st == MSO_SHAPE_TYPE.PLACEHOLDER:
        return "placeholder"
    if st == MSO_SHAPE_TYPE.LINE:
        return "line"
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if st == MSO_SHAPE_TYPE.FREEFORM:
        return "freeform"
    return "other"


def walk_shapes(slide) -> list[dict]:
    """Walk all shapes including into groups."""
    out: list[dict] = []

    def recurse(shapes):
        for s in shapes:
            bb = shape_bbox(s)
            if bb:
                out.append(bb)
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    recurse(s.shapes)
                except Exception:
                    pass

    walk_shapes_errors_ignored(slide.shapes, out)
    return out


def walk_shapes_errors_ignored(shapes, out: list[dict]):
    for s in shapes:
        bb = shape_bbox(s)
        if bb:
            out.append(bb)
        try:
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk_shapes_errors_ignored(s.shapes, out)
        except Exception:
            pass


def relaxed_signature(shapes: list[dict]) -> str:
    """Signature based on CONTENT types only (chart, table, picture)."""
    counts: Counter = Counter()
    for s in shapes:
        if s["type"] in ("chart", "table", "picture"):
            counts[s["type"]] += 1
    if not counts:
        return "empty"
    parts = []
    for k in ("chart", "table", "picture"):
        if counts.get(k):
            parts.append(f"{counts[k]}_{k}")
    return "_".join(parts)


# ---------------------------------------------------------------------------
# Analyze all slides
# ---------------------------------------------------------------------------


def analyze_all() -> tuple[dict, list, list]:
    """Returns (signature_clusters, chart_positions, table_positions)."""
    pptx_files = sorted(p for p in DECKS_DIR.glob("*.pptx") if not p.name.startswith("~"))
    sig_groups: dict[str, list[list[dict]]] = defaultdict(list)
    chart_positions: list[dict] = []
    table_positions: list[dict] = []

    for p in pptx_files:
        try:
            prs = Presentation(str(p))
            for i, slide in enumerate(prs.slides):
                try:
                    shapes = walk_shapes(slide)
                    sig = relaxed_signature(shapes)
                    sig_groups[sig].append(shapes)
                    # Per-chart / per-table positions
                    for s in shapes:
                        if s["type"] == "chart" and all(s[k] is not None for k in ("left", "top", "width", "height")):
                            chart_positions.append({
                                **{k: s[k] for k in ("left", "top", "width", "height")},
                                "deck": p.name,
                                "slide_index": i,
                                "signature": sig,
                            })
                        if s["type"] == "table" and all(s[k] is not None for k in ("left", "top", "width", "height")):
                            table_positions.append({
                                **{k: s[k] for k in ("left", "top", "width", "height")},
                                "deck": p.name,
                                "slide_index": i,
                                "signature": sig,
                            })
                except Exception:
                    continue
        except Exception:
            continue

    return sig_groups, chart_positions, table_positions


# ---------------------------------------------------------------------------
# Stats helpers
# ---------------------------------------------------------------------------


def coord_stats(values: list[float]) -> dict:
    if not values:
        return {}
    out = {
        "n": len(values),
        "min": round(min(values), 2),
        "max": round(max(values), 2),
        "mean": round(mean(values), 2),
        "median": round(median(values), 2),
    }
    if len(values) > 1:
        try:
            out["stdev"] = round(stdev(values), 2)
        except Exception:
            pass
    return out


def per_type_stats(shape_groups: list[list[dict]]) -> dict:
    """For a signature group, compute stats per shape type across all slides."""
    buckets: dict[str, dict[str, list]] = defaultdict(
        lambda: {"lefts": [], "tops": [], "widths": [], "heights": []}
    )
    for slide_shapes in shape_groups:
        # For non-content shapes, only first N kept to avoid skew
        for s in slide_shapes:
            if s["type"] in ("chart", "table", "picture", "text_box"):
                if all(s[k] is not None for k in ("left", "top", "width", "height")):
                    buckets[s["type"]]["lefts"].append(s["left"])
                    buckets[s["type"]]["tops"].append(s["top"])
                    buckets[s["type"]]["widths"].append(s["width"])
                    buckets[s["type"]]["heights"].append(s["height"])

    result: dict[str, dict] = {}
    for stype, coords in buckets.items():
        result[stype] = {
            "n": len(coords["lefts"]),
            "left": coord_stats(coords["lefts"]),
            "top": coord_stats(coords["tops"]),
            "width": coord_stats(coords["widths"]),
            "height": coord_stats(coords["heights"]),
        }
    return result


# ---------------------------------------------------------------------------
# Position clustering — group similar positions together
# ---------------------------------------------------------------------------


def cluster_positions(positions: list[dict], tolerance: float = 0.5) -> list[dict]:
    """Fast clustering via rounded bucket keys.

    Rounds each coordinate to the nearest `tolerance` boundary and uses the tuple
    as a dict key. O(N) instead of O(N^2).
    """
    def bucket(v: float) -> float:
        return round(v / tolerance) * tolerance

    buckets: dict[tuple, list[dict]] = defaultdict(list)
    for p in positions:
        key = (bucket(p["left"]), bucket(p["top"]), bucket(p["width"]), bucket(p["height"]))
        buckets[key].append(p)

    clusters: list[dict] = []
    for key, items in buckets.items():
        if len(items) < 3:
            continue
        clusters.append({
            "n": len(items),
            "left_median": round(median(x["left"] for x in items), 2),
            "top_median": round(median(x["top"] for x in items), 2),
            "width_median": round(median(x["width"] for x in items), 2),
            "height_median": round(median(x["height"] for x in items), 2),
            "example_decks": list({x["deck"] for x in items})[:5],
        })

    clusters.sort(key=lambda c: -c["n"])
    return clusters


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------


def write_clusters_md(
    sig_groups: dict,
    chart_clusters: list[dict],
    table_clusters: list[dict],
    output: Path,
) -> None:
    lines: list[str] = []
    lines.append("# Layout Cluster Analysis")
    lines.append("")
    lines.append("Shape signatures relaxed to count only content shapes (chart / table / picture). Decorative elements (auto_shape, line, freeform) ignored — they vary per client but don't define layout.")
    lines.append("")

    lines.append("## Content Signature Distribution (top 20)")
    lines.append("")
    lines.append("| Signature | Slides |")
    lines.append("|---|---|")
    sorted_sigs = sorted(sig_groups.items(), key=lambda kv: -len(kv[1]))
    for sig, slides in sorted_sigs[:20]:
        lines.append(f"| `{sig}` | {len(slides)} |")
    lines.append("")

    lines.append("## Per-Signature Coordinate Stats (top 12)")
    lines.append("")

    for sig, slides in sorted_sigs[:12]:
        stats = per_type_stats(slides)
        if not any(t in stats for t in ("chart", "table", "picture")):
            continue  # skip text-only/empty
        lines.append(f"### `{sig}` — {len(slides)} slides")
        lines.append("")
        lines.append("| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |")
        lines.append("|---|---|---|---|---|---|")
        for stype in ("chart", "table", "picture", "text_box"):
            if stype not in stats:
                continue
            s = stats[stype]
            lines.append(
                f"| `{stype}` | {s['n']} | "
                f"{s['left']['median']}\" | {s['top']['median']}\" | "
                f"{s['width']['median']}\" | {s['height']['median']}\" |"
            )
        lines.append("")

    lines.append("---")
    lines.append("")

    lines.append("## Chart Position Clusters")
    lines.append("")
    lines.append(f"Clustered {len(chart_clusters)} distinct chart positions (grouping similar-position charts across all decks, tolerance 0.5 inch).")
    lines.append("")
    lines.append("Each row here is a candidate `LAYOUTS{}` preset — a recurring chart position in client decks.")
    lines.append("")
    lines.append("| # | Occurrences | Left | Top | Width | Height | Decks |")
    lines.append("|---|---|---|---|---|---|---|")
    for i, c in enumerate(chart_clusters[:30], 1):
        decks = ", ".join(d[:25] for d in c["example_decks"][:3])
        lines.append(
            f"| {i} | {c['n']} | {c['left_median']}\" | {c['top_median']}\" | "
            f"{c['width_median']}\" | {c['height_median']}\" | {decks} |"
        )
    lines.append("")

    lines.append("## Table Position Clusters")
    lines.append("")
    lines.append(f"Clustered {len(table_clusters)} distinct table positions.")
    lines.append("")
    lines.append("| # | Occurrences | Left | Top | Width | Height | Decks |")
    lines.append("|---|---|---|---|---|---|---|")
    for i, c in enumerate(table_clusters[:30], 1):
        decks = ", ".join(d[:25] for d in c["example_decks"][:3])
        lines.append(
            f"| {i} | {c['n']} | {c['left_median']}\" | {c['top_median']}\" | "
            f"{c['width_median']}\" | {c['height_median']}\" | {decks} |"
        )
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    print("Analyzing layouts...")
    sig_groups, chart_positions, table_positions = analyze_all()
    print(f"  Signatures found: {len(sig_groups)}")
    print(f"  Charts: {len(chart_positions)}")
    print(f"  Tables: {len(table_positions)}")

    print("Clustering chart positions...")
    chart_clusters = cluster_positions(chart_positions, tolerance=0.5)
    print(f"  Chart clusters (>=3 slides): {len(chart_clusters)}")

    print("Clustering table positions...")
    table_clusters = cluster_positions(table_positions, tolerance=0.5)
    print(f"  Table clusters (>=3 slides): {len(table_clusters)}")

    # Write outputs
    sig_out = {
        sig: {
            "slides": len(slides),
            "stats": per_type_stats(slides),
        }
        for sig, slides in sig_groups.items()
    }
    (OUTPUTS_DIR / "layout_clusters.json").write_text(
        json.dumps(sig_out, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "chart_positions.json").write_text(
        json.dumps(chart_clusters, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUTS_DIR / "table_positions.json").write_text(
        json.dumps(table_clusters, indent=2, default=str), encoding="utf-8"
    )

    write_clusters_md(
        sig_groups, chart_clusters, table_clusters,
        OUTPUTS_DIR / "layout_clusters.md",
    )

    print(f"\nDone. Outputs in {OUTPUTS_DIR}/")
    print("  layout_clusters.md    — readable cluster report")
    print("  layout_clusters.json  — signature stats")
    print("  chart_positions.json  — chart position clusters")
    print("  table_positions.json  — table position clusters")

    return 0


if __name__ == "__main__":
    sys.exit(main())
