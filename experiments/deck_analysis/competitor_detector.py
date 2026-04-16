"""
Competitor Detector
===================

Auto-detects competitor brands for each PET brand by cross-referencing series
colors across decks.

Hypothesis: When Brand A's deck contains charts comparing to Brand B, Brand B's
signature color shows up as a minor color (palette rank 2+) in A's deck. If
Brand B exists as its own primary brand in another deck, we have a match.

Signals used:
  1. Per-deck series colors (which colors appear together in ONE deck)
  2. Primary color registry (what each brand claims as its brand color)
  3. Therapy-area overlap (competitors usually share therapy areas)
  4. RGB distance for near-matches (printer variance, slight shade drift)

Outputs review-first suggestions — does NOT auto-populate BRAND{}.

Run:
    cd experiments/deck_analysis
    python competitor_detector.py

Outputs:
    outputs/competitor_map.json   — ranked competitor suggestions per brand
    outputs/competitor_map.md     — human-readable review doc
"""
from __future__ import annotations

import json
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).parent
DECKS_DIR = HERE / "decks"
OUTPUTS_DIR = HERE / "outputs"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def qn_c(tag: str) -> str:
    return f"{{{NS_C}}}{tag}"


def qn_a(tag: str) -> str:
    return f"{{{NS_A}}}{tag}"


# Structural colors (same list as brand_mapper) — never treated as brand signals
STRUCTURAL_COLORS = {
    "000000", "FFFFFF",
    "00B050",
    "FF0000", "C00000", "EB1700", "FF2929",
    "F2F2F2", "E7E6E6", "E5E5E5", "F0F0F0", "F3F4F3",
    "BFBFBF", "A6A6A6", "7F7F7F", "D3D3D3", "CCCCCC",
    "595959", "404040", "505050", "3F3F3F", "3E403F",
    "2C2C2C", "282828", "292A2A", "171616", "1F1F1F",
    "4D4D4F", "535554", "636466", "747F74",
    "4A5C58", "3F4444", "4B4B4B", "26272A",
    "E7E8E9", "F0F1F1", "E3E2E3", "CACACA", "D8DDE5",
    "AFABAB",
}


def is_structural(hex_str: str) -> bool:
    return hex_str.upper() in STRUCTURAL_COLORS


def hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    h = hex_str.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rgb_distance(a: str, b: str) -> float:
    """Euclidean RGB distance. Range 0-441. <20 = near-identical, <50 = same-ish."""
    r1, g1, b1 = hex_to_rgb(a)
    r2, g2, b2 = hex_to_rgb(b)
    return ((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2) ** 0.5


# ---------------------------------------------------------------------------
# Per-deck color extraction
# ---------------------------------------------------------------------------


def extract_deck_colors(pptx_path: Path) -> Counter:
    """Pull series fill colors from every chart XML in one deck."""
    counter: Counter = Counter()
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            for name in z.namelist():
                if not (name.startswith("ppt/charts/chart") and name.endswith(".xml")):
                    continue
                try:
                    root = ET.fromstring(z.read(name))
                except ET.ParseError:
                    continue
                for ser in root.iter(qn_c("ser")):
                    sp_pr = ser.find(qn_c("spPr"))
                    if sp_pr is None:
                        continue
                    rgb = sp_pr.find(f"{qn_a('solidFill')}/{qn_a('srgbClr')}")
                    if rgb is not None and rgb.get("val"):
                        counter[rgb.get("val").upper()] += 1
    except Exception:
        return counter
    return counter


# ---------------------------------------------------------------------------
# Primary color registry
# ---------------------------------------------------------------------------


def build_primary_registry(brand_by_brand: dict) -> dict[str, list[dict]]:
    """Map each primary-palette color → list of brands claiming it.

    A brand's "primary" is palette_filtered[0]. We also register palette_filtered[1]
    (secondary) with lower weight, because some brands' secondary is still a
    strong identifier.
    """
    registry: dict[str, list[dict]] = defaultdict(list)
    for brand, info in brand_by_brand.items():
        palette = info.get("palette_filtered", [])
        if not palette:
            continue
        # Primary — strongest claim on this color
        registry[palette[0].upper()].append({
            "brand": brand,
            "client": info.get("client"),
            "therapy_area": info.get("therapy_area"),
            "rank": 0,  # 0 = primary
        })
        # Secondary — weaker claim
        if len(palette) > 1:
            registry[palette[1].upper()].append({
                "brand": brand,
                "client": info.get("client"),
                "therapy_area": info.get("therapy_area"),
                "rank": 1,  # 1 = secondary
            })
    return registry


# ---------------------------------------------------------------------------
# Therapy area matching
# ---------------------------------------------------------------------------


def therapy_area_overlap(ta1: str, ta2: str) -> float:
    """Return 1.0 for exact match, 0.5 for parent overlap (same root), 0.0 otherwise."""
    if not ta1 or not ta2 or ta1 == "Unknown" or ta2 == "Unknown":
        return 0.0
    if ta1 == ta2:
        return 1.0
    # Split on "/" — e.g. "Oncology/NSCLC" → ["Oncology", "NSCLC"]
    root1 = ta1.split("/")[0].strip().lower()
    root2 = ta2.split("/")[0].strip().lower()
    if root1 == root2:
        return 0.5
    return 0.0


# ---------------------------------------------------------------------------
# Candidate scoring
# ---------------------------------------------------------------------------

EXACT_MATCH_THRESHOLD = 5    # RGB distance <5 = exact (printer variance)
NEAR_MATCH_THRESHOLD = 30    # <30 = close shade, likely same brand
LOOSE_MATCH_THRESHOLD = 60   # <60 = possible, needs review


def score_candidate(
    deck_color: str,
    deck_color_count: int,
    total_deck_colors: int,
    candidate_brand: dict,
    candidate_color: str,
    source_brand_ta: str,
) -> tuple[float, str]:
    """Return (confidence_score, reason_string) for a potential competitor match.

    Score factors:
      - Color match precision (exact > near > loose)
      - Source rank (primary match stronger than secondary match)
      - Therapy-area overlap
      - Frequency of the minor color in the source deck (more charts = more signal)
    """
    dist = rgb_distance(deck_color, candidate_color)
    if dist > LOOSE_MATCH_THRESHOLD:
        return 0.0, f"color too distant ({dist:.0f})"

    if dist < EXACT_MATCH_THRESHOLD:
        color_score = 1.0
        match_label = "exact"
    elif dist < NEAR_MATCH_THRESHOLD:
        color_score = 0.7
        match_label = f"near ({dist:.0f})"
    else:
        color_score = 0.4
        match_label = f"loose ({dist:.0f})"

    # Rank penalty — matching a primary is stronger than matching a secondary
    rank_score = 1.0 if candidate_brand["rank"] == 0 else 0.5

    # Therapy area overlap
    ta_score = therapy_area_overlap(source_brand_ta, candidate_brand["therapy_area"])
    if ta_score == 0.0:
        # No TA overlap is a strong negative — competitors almost always share TA
        ta_multiplier = 0.2
    elif ta_score == 0.5:
        ta_multiplier = 0.7
    else:
        ta_multiplier = 1.0

    # Frequency — minor color's share of this deck's series
    freq_share = deck_color_count / total_deck_colors if total_deck_colors else 0.0
    # Cap frequency boost — even a 50% minor color shouldn't dominate the score
    freq_score = min(freq_share * 2, 1.0)

    confidence = color_score * rank_score * ta_multiplier * (0.5 + 0.5 * freq_score)

    reason = (
        f"color={match_label}, "
        f"candidate_rank={'primary' if candidate_brand['rank'] == 0 else 'secondary'}, "
        f"ta_overlap={ta_score:.1f}, "
        f"minor_freq={freq_share:.0%} ({deck_color_count}/{total_deck_colors})"
    )
    return confidence, reason


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    # Load prior brand analysis
    brand_by_brand = json.loads((OUTPUTS_DIR / "brand_by_brand.json").read_text(encoding="utf-8"))
    deck_mapping = json.loads((OUTPUTS_DIR / "brand_mapping.json").read_text(encoding="utf-8"))

    # Build primary color registry across all brands
    registry = build_primary_registry(brand_by_brand)
    print(f"Built primary color registry: {len(registry)} distinct colors from {len(brand_by_brand)} brands")

    # Group decks by brand so we can aggregate candidate scores
    decks_by_brand: dict[str, list[dict]] = defaultdict(list)
    for d in deck_mapping:
        decks_by_brand[d["brand"]].append(d)

    # For each deck, extract per-deck colors and look up candidates
    pptx_files = sorted(p for p in DECKS_DIR.glob("*.pptx") if not p.name.startswith("~"))
    deck_colors: dict[str, Counter] = {}

    print(f"\nExtracting per-deck colors from {len(pptx_files)} decks...")
    for i, p in enumerate(pptx_files, 1):
        print(f"  [{i}/{len(pptx_files)}] {p.name[:60]}")
        deck_colors[p.name] = extract_deck_colors(p)

    # Aggregate candidates per source brand
    # Structure: {source_brand: {candidate_brand: {score, reasons, matched_colors, decks}}}
    candidates: dict[str, dict[str, dict]] = defaultdict(lambda: defaultdict(
        lambda: {"score_sum": 0.0, "reasons": [], "matched_colors": set(), "decks": set()}
    ))

    for deck_info in deck_mapping:
        filename = deck_info["filename"]
        source_brand = deck_info["brand"]
        source_ta = deck_info["therapy_area"]
        colors = deck_colors.get(filename, Counter())

        # Total series-color count in this deck (non-structural only)
        filtered_colors = {c: n for c, n in colors.items() if not is_structural(c)}
        total = sum(filtered_colors.values())
        if total == 0:
            continue

        # The source brand's own palette — don't flag these as competitors
        source_palette = {c.upper() for c in brand_by_brand.get(source_brand, {}).get("palette_filtered", [])}

        for deck_color, count in filtered_colors.items():
            if deck_color in source_palette:
                # It's the source brand's own color — skip
                continue
            if is_structural(deck_color):
                continue

            # Find candidates in the registry (within rgb threshold)
            for registry_color, candidate_list in registry.items():
                if rgb_distance(deck_color, registry_color) > LOOSE_MATCH_THRESHOLD:
                    continue
                for candidate in candidate_list:
                    if candidate["brand"] == source_brand:
                        # Don't match self (shouldn't happen due to palette skip, but safety)
                        continue
                    score, reason = score_candidate(
                        deck_color, count, total,
                        candidate, registry_color, source_ta,
                    )
                    if score <= 0:
                        continue
                    entry = candidates[source_brand][candidate["brand"]]
                    entry["score_sum"] += score
                    entry["reasons"].append(f"#{deck_color} → #{registry_color} ({reason})")
                    entry["matched_colors"].add(deck_color)
                    entry["decks"].add(filename)
                    # Record metadata once
                    entry["candidate_client"] = candidate["client"]
                    entry["candidate_therapy_area"] = candidate["therapy_area"]

    # Build final output: sorted candidates per brand with normalized scores
    competitor_map: dict[str, dict] = {}
    for source_brand, cand_dict in candidates.items():
        ranked = []
        for cand_brand, data in cand_dict.items():
            ranked.append({
                "candidate_brand": cand_brand,
                "candidate_client": data.get("candidate_client"),
                "candidate_therapy_area": data.get("candidate_therapy_area"),
                "score": round(data["score_sum"], 3),
                "matched_colors": sorted(data["matched_colors"]),
                "evidence_deck_count": len(data["decks"]),
                "top_reasons": data["reasons"][:3],  # cap
            })
        ranked.sort(key=lambda x: -x["score"])
        if not ranked:
            continue
        competitor_map[source_brand] = {
            "therapy_area": brand_by_brand.get(source_brand, {}).get("therapy_area", "Unknown"),
            "client": brand_by_brand.get(source_brand, {}).get("client", "UNKNOWN"),
            "deck_count": len(brand_by_brand.get(source_brand, {}).get("deck_filenames", [])),
            "candidates": ranked,
        }

    # Write JSON
    out_json = OUTPUTS_DIR / "competitor_map.json"
    out_json.write_text(json.dumps(competitor_map, indent=2), encoding="utf-8")
    print(f"\nWrote {out_json}")

    # Write human-readable review doc
    write_review_md(competitor_map, OUTPUTS_DIR / "competitor_map.md")
    print(f"Wrote {OUTPUTS_DIR / 'competitor_map.md'}")

    # Print a summary to console
    print("\n=== Competitor Detection Summary ===")
    print(f"Brands with >=1 candidate: {len(competitor_map)}")
    high_confidence = sum(1 for info in competitor_map.values()
                          if info["candidates"] and info["candidates"][0]["score"] >= 1.0)
    print(f"Brands with a HIGH-confidence top candidate (score >=1.0): {high_confidence}")
    print("\nTop candidates per brand (score >= 0.5 only):")
    for brand, info in sorted(competitor_map.items()):
        strong = [c for c in info["candidates"] if c["score"] >= 0.5]
        if not strong:
            continue
        top = strong[0]
        print(f"  {brand:25s} | TA: {info['therapy_area']:25s} -> "
              f"{top['candidate_brand']:20s} (score={top['score']:.2f}, "
              f"colors={','.join('#' + c for c in top['matched_colors'][:2])})")

    return 0


def write_review_md(competitor_map: dict, out_path: Path) -> None:
    lines = []
    lines.append("# Competitor Detection — Review Doc")
    lines.append("")
    lines.append("Auto-generated candidate competitors for each PET brand.")
    lines.append("Review and cherry-pick valid matches before merging into `BRAND{}`.")
    lines.append("")
    lines.append("**Score interpretation:**")
    lines.append("- `>= 1.0` — exact color match + same therapy area + frequent minor color")
    lines.append("- `0.5 - 1.0` — strong signal (near match or secondary rank)")
    lines.append("- `< 0.5` — weak / coincidental, usually ignore")
    lines.append("")
    lines.append("---")
    lines.append("")

    for brand in sorted(competitor_map.keys()):
        info = competitor_map[brand]
        lines.append(f"## {brand}")
        lines.append(f"- **Client:** {info['client']}")
        lines.append(f"- **Therapy area:** {info['therapy_area']}")
        lines.append(f"- **Decks analyzed:** {info['deck_count']}")
        lines.append("")
        if not info["candidates"]:
            lines.append("_No candidates detected._")
            lines.append("")
            continue
        lines.append("| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |")
        lines.append("|------|-----------|--------|----|-------|-----------------|-------|")
        for i, c in enumerate(info["candidates"][:5], 1):
            colors_str = ", ".join(f"`#{x}`" for x in c["matched_colors"][:3])
            lines.append(
                f"| {i} | **{c['candidate_brand']}** | {c['candidate_client']} | "
                f"{c['candidate_therapy_area']} | {c['score']:.2f} | {colors_str} | "
                f"{c['evidence_deck_count']} |"
            )
        lines.append("")
        # Show reasons for the top candidate
        top = info["candidates"][0]
        if top["score"] >= 0.5:
            lines.append(f"**Top match reasoning ({top['candidate_brand']}):**")
            for r in top["top_reasons"]:
                lines.append(f"- {r}")
            lines.append("")
        lines.append("---")
        lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    import sys
    sys.exit(main())
