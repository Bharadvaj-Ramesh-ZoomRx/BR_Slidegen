"""For every chart on a shifted ds whose refreshed values == source (welded),
inspect its selectedColumns and categorize the pattern. Helps us enumerate
the families Fix L needs to handle next.

Categories:
  P0_NO_WAVE_LABEL — selectedColumns has no wave labels at all (welding is
                     not Fix L territory; mapper failed for some other reason)
  P1_AT_AT_COMPOUND — uses ` @:@ ` compound separator (Fix L should rewrite)
  P2_DASH_COMPOUND  — uses ` - ` compound separator (Fix L extension should
                     rewrite)
  P3_STANDALONE     — entry is a bare wave label like 'Project Wave 13'
                     (Fix L should rewrite)
  P4_EMBEDDED_WAVE  — wave digits embedded inside a non-standard string
                     (e.g., 'PET_W33', 'Q1 2026 brand metrics'). Fix L's
                     pattern doesn't match these.
  P5_UNUSUAL_LABEL  — looks like a wave but format outside the regex
                     (e.g., 'Wave - 8949', 'Final Wave', 'Wave: 12')
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx import Presentation

REPO = Path(__file__).resolve().parents[1]

# Same patterns Fix L uses
WAVE_PATTERNS = (
    re.compile(r"^(?:Project )?Wave \d+$"),
    re.compile(r"^[A-Z][a-z]{2}'\d{2}$"),
    re.compile(r"^Q[1-4]'\d{2}$"),
    re.compile(r"^Q[1-4] \d{4}$"),
)
EMBEDDED_WAVE_HINT = re.compile(r"\b(?:W|Wave|Q|Project Wave)[\s_]*\d+", re.IGNORECASE)


def is_fix_l_wave_label(s: str) -> bool:
    return any(p.match(s) for p in WAVE_PATTERNS)


def has_embedded_wave_hint(s: str) -> bool:
    """Looks like it MIGHT be a wave label but doesn't match Fix L's strict regex."""
    if is_fix_l_wave_label(s):
        return False
    return bool(EMBEDDED_WAVE_HINT.search(s))


def categorize(sc: list) -> str:
    """Categorize one selectedColumns list."""
    has_at_compound_with_wave = False
    has_dash_compound_with_wave = False
    has_standalone_wave = False
    has_embedded_hint = False
    any_wave = False
    for entry in sc:
        if not isinstance(entry, str):
            continue
        # Standalone wave
        if is_fix_l_wave_label(entry):
            has_standalone_wave = True
            any_wave = True
            continue
        # @:@ compound
        if " @:@ " in entry:
            parts = entry.split(" @:@ ")
            if any(is_fix_l_wave_label(p) for p in parts):
                has_at_compound_with_wave = True
                any_wave = True
                continue
        # - compound
        if " - " in entry:
            parts = entry.split(" - ")
            if any(is_fix_l_wave_label(p) for p in parts):
                has_dash_compound_with_wave = True
                any_wave = True
                continue
        # Embedded hint (Fix L can't catch)
        if has_embedded_wave_hint(entry):
            has_embedded_hint = True
            continue
    if not any_wave and not has_embedded_hint:
        return "P0_NO_WAVE_LABEL"
    if has_at_compound_with_wave:
        return "P1_AT_AT_COMPOUND"
    if has_dash_compound_with_wave:
        return "P2_DASH_COMPOUND"
    if has_standalone_wave:
        return "P3_STANDALONE"
    if has_embedded_hint:
        return "P4_EMBEDDED_WAVE"
    return "P5_UNUSUAL_LABEL"


def categorize_deck(deck_key: str) -> dict:
    spec_path = REPO / "output" / "_step6" / deck_key / "spec_shift.json"
    refreshed_path = REPO / "output" / "_step6" / deck_key / "refreshed_shift.pptx"
    overrides_path = REPO / "output" / "_step6" / deck_key / "shift_overrides.json"
    source_path_map = {
        "atu_q1_26": REPO / "projects" / "J&J Rybrevant PET" / "Template" / "ZoomRx_UC_ATU_Report_Q1_'26.pptx",
        "creon_pet_w33": REPO / "projects" / "J&J Rybrevant PET" / "Template" / "CREON Share of Voice Study - W33.pptx",
    }
    source_path = source_path_map[deck_key]

    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    overrides = json.loads(overrides_path.read_text(encoding="utf-8"))["overrides"]

    # Index spec components on shifted ds → (selectedColumns, ds_key)
    comp_info_by_name: dict[tuple, tuple[list, str]] = {}
    comp_info_by_pos: dict[int, list[tuple[float, float, list, str]]] = {}
    for slide in spec.get("slides", []):
        s_idx = slide["slide_index"]
        slide_default_ds = slide.get("data_source")
        for comp in slide.get("components", []):
            ds = comp.get("data_source") or slide_default_ds
            if ds not in overrides:
                continue
            if comp.get("type") != "chart":
                continue
            sc = comp.get("raw_mapping_config", {}).get("selectedColumns", [])
            name = comp.get("name") or ""
            if name:
                comp_info_by_name[(s_idx, name)] = (sc, ds)
            cpos = comp.get("position", {}) or {}
            cl = float(cpos.get("left", 0) or 0)
            ct = float(cpos.get("top", 0) or 0)
            comp_info_by_pos.setdefault(s_idx, []).append((cl, ct, sc, ds))

    def _resolve(s_idx, name, pos):
        if name and (s_idx, name) in comp_info_by_name:
            return comp_info_by_name[(s_idx, name)]
        for cl, ct, sc, ds in comp_info_by_pos.get(s_idx, []):
            if abs(cl - pos[0]) <= 0.05 and abs(ct - pos[1]) <= 0.05:
                return sc, ds
        return None

    # Walk source vs refreshed; for charts where values are identical → welded
    src = Presentation(str(source_path))
    ref = Presentation(str(refreshed_path))

    by_category: dict[str, list[dict]] = defaultdict(list)
    welded_count = 0

    for s_idx, (s_slide, r_slide) in enumerate(zip(src.slides, ref.slides)):
        src_charts = [s for s in s_slide.shapes if s.has_chart]
        ref_charts_by_pos = {
            (round((s.left or 0) / 914400, 2), round((s.top or 0) / 914400, 2)): s
            for s in r_slide.shapes if s.has_chart
        }
        for src_shape in src_charts:
            shape_name = src_shape.name or ""
            sl = round((src_shape.left or 0) / 914400, 2)
            st = round((src_shape.top or 0) / 914400, 2)
            info = _resolve(s_idx, shape_name, (sl, st))
            if info is None:
                continue
            sc, ds_key = info
            # find ref shape
            ref_shape = ref_charts_by_pos.get((sl, st))
            if ref_shape is None:
                # nearest within tolerance
                for (rl, rt), rs in ref_charts_by_pos.items():
                    if abs(rl - sl) <= 0.05 and abs(rt - st) <= 0.05:
                        ref_shape = rs
                        break
            if ref_shape is None:
                continue
            try:
                src_vals = tuple(round(v, 4) if v is not None else None
                                 for s in src_shape.chart.plots[0].series for v in s.values)
                ref_vals = tuple(round(v, 4) if v is not None else None
                                 for s in ref_shape.chart.plots[0].series for v in s.values)
            except Exception:
                continue
            if src_vals != ref_vals:
                continue  # not welded
            welded_count += 1
            cat = categorize(sc)
            by_category[cat].append({
                "slide": s_idx,
                "name": shape_name,
                "ds": ds_key,
                "selectedColumns": sc,
            })

    return {"deck": deck_key, "welded_total": welded_count, "by_category": by_category}


def print_report(rpt: dict) -> None:
    print(f"\n=== {rpt['deck']} | welded_total = {rpt['welded_total']} ===")
    for cat in sorted(rpt["by_category"]):
        rows = rpt["by_category"][cat]
        print(f"\n  {cat}: {len(rows)} charts")
        for r in rows[:3]:
            sc_preview = r["selectedColumns"][:5]
            if len(r["selectedColumns"]) > 5:
                sc_preview = sc_preview + [f"... +{len(r['selectedColumns'])-5} more"]
            print(f"    slide {r['slide']:>2} {r['name']!r}  ds={r['ds']}")
            print(f"      selectedColumns ({len(r['selectedColumns'])}): {sc_preview}")
        if len(rows) > 3:
            print(f"    ... +{len(rows)-3} more in this category")


if __name__ == "__main__":
    for k in ("atu_q1_26", "creon_pet_w33"):
        rpt = categorize_deck(k)
        print_report(rpt)
        # Also dump full JSON for reference
        out = REPO / "output" / "_step6" / k / "welded_patterns.json"
        out.write_text(
            json.dumps(rpt, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        print(f"\n  full report -> {out}")
