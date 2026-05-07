"""Tests for the talking-header refresh flow.

Covers the four behavioural changes shipped this session:
  1. Drop spec-membership gate — non-connected slides ALSO get rewritten.
  2. Find tables when no chart is present — table-only slides supported.
  3. Always rewrite when a talking header exists, even if values are
     unchanged (mixed-component slides may have manual edits).
  4. Distinguish the talking header (top-left, full-sentence) from the
     slide-section title (short label below) and the roadmap (narrow,
     top-right).

Tests stub out the LLM call so they run offline.
"""
from io import BytesIO
import sys
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.headline_refresh import (
    _all_data_shapes,
    _build_prompt,
    _find_headline_shape,
    _summarize_table,
    refresh_headlines,
)


# ────────────────────────────────────────────────────────────────────────
# _find_headline_shape — talking header vs section title vs roadmap
# ────────────────────────────────────────────────────────────────────────


def _make_slide(width_in=10.0):
    """Build an empty 1-slide presentation for layout tests."""
    prs = Presentation()
    prs.slide_width = Inches(width_in)
    prs.slide_height = Inches(7.5)
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank


def _add_textbox(slide, text, *, left_in, top_in, width_in, height_in=0.5,
                 name=None):
    tb = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in))
    tb.text_frame.text = text
    if name:
        tb.name = name
    return tb


def test_picks_top_left_talking_header_over_section_title():
    """Slide has a talking header (top-LEFT, full sentence) AND a slide-
    section title (centered, short label) BELOW it. Pick the talking
    header — it's topmost AND has long text."""
    slide = _make_slide()
    _add_textbox(
        slide,
        "RINVOQ retained perceived edge on rapid symptom relief, with leadership "
        "at 47% vs 33% for Tremfya across UC HCPs in Q1.",
        left_in=0.3, top_in=0.2, width_in=8.0, height_in=1.0)
    _add_textbox(slide, "Message Recall",  # short section title BELOW
                 left_in=2.0, top_in=1.6, width_in=6.0, height_in=0.5,
                 name="Title 1")
    pick = _find_headline_shape(slide)
    assert pick is not None
    assert pick.text_frame.text.startswith("RINVOQ retained")


def test_excludes_short_section_title_label():
    """A short label (< 30 chars) like 'Message Recall' must NOT be
    chosen even if it's the only wide top-region text frame."""
    slide = _make_slide()
    _add_textbox(slide, "Message Recall",  # 14 chars — too short
                 left_in=0.5, top_in=0.5, width_in=9.0)
    pick = _find_headline_shape(slide)
    assert pick is None  # no talking header on this slide


def test_excludes_narrow_roadmap_strip():
    """The roadmap is narrow (< 40% of slide width) at top-right. Must
    not be picked even if its text happens to be > 30 chars."""
    slide = _make_slide(width_in=10.0)
    # roadmap ~3 inches = 30% width
    _add_textbox(slide,
                 "Section 1 / Section 2 / Section 3 / Section 4 active",
                 left_in=6.5, top_in=0.1, width_in=3.0)
    pick = _find_headline_shape(slide)
    assert pick is None  # roadmap excluded by width filter


def test_picks_topmost_when_multiple_talking_headers():
    """Two qualifying shapes — pick the topmost (talking header sits
    above any other narrative-shaped text)."""
    slide = _make_slide()
    _add_textbox(
        slide,
        "Top story: HCPs view RINVOQ as differentiated on rapid symptom relief.",
        left_in=0.3, top_in=0.2, width_in=8.0)
    _add_textbox(
        slide,
        "Secondary story: Tremfya holds steady on long-term remission perceptions.",
        left_in=0.3, top_in=1.5, width_in=8.0)
    pick = _find_headline_shape(slide)
    assert pick is not None
    assert pick.text_frame.text.startswith("Top story")


def test_tiebreak_left_when_tops_match():
    """When two qualifying shapes share the same top, prefer the
    LEFTMOST (talking header is top-left convention)."""
    slide = _make_slide()
    _add_textbox(
        slide,
        "Left talking header: claim about the leftmost data narrative here.",
        left_in=0.3, top_in=0.2, width_in=4.5)
    _add_textbox(
        slide,
        "Right side label that happens to also be a long sentence.",
        left_in=5.0, top_in=0.2, width_in=4.5)
    pick = _find_headline_shape(slide)
    assert pick is not None
    assert pick.text_frame.text.startswith("Left talking header")


def test_returns_none_for_cover_or_divider():
    """Cover slides / dividers have no talking header — return None."""
    slide = _make_slide()
    _add_textbox(slide, "PROJECT TITLE",  # short
                 left_in=2.0, top_in=3.0, width_in=6.0)  # centred, not at top
    pick = _find_headline_shape(slide)
    assert pick is None


# ────────────────────────────────────────────────────────────────────────
# _summarize_table + _all_data_shapes
# ────────────────────────────────────────────────────────────────────────


def test_summarize_table_extracts_cell_text():
    slide = _make_slide()
    tbl = slide.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(6), Inches(2))
    tbl.table.cell(0, 0).text = "ID"
    tbl.table.cell(0, 1).text = "Message"
    tbl.table.cell(0, 2).text = "%"
    tbl.table.cell(1, 0).text = "C6"
    tbl.table.cell(1, 1).text = "Take CREON every meal..."
    tbl.table.cell(1, 2).text = "57%"
    tbl.table.cell(2, 0).text = "C10"
    tbl.table.cell(2, 1).text = "CREON is the #1 PERT prescribed..."
    tbl.table.cell(2, 2).text = "55%"

    summary = _summarize_table(tbl)
    assert "ID | Message | %" in summary
    assert "C6" in summary
    assert "Take CREON every meal..." in summary
    assert "57%" in summary


def test_all_data_shapes_orders_by_area():
    slide = _make_slide()
    # Small table (1×1)
    small = slide.shapes.add_table(1, 1, Inches(0), Inches(0),
                                   Inches(2), Inches(0.5))
    small.name = "small_table"
    small.table.cell(0, 0).text = "n=120"
    # Larger table (3×3)
    big = slide.shapes.add_table(3, 3, Inches(2), Inches(2),
                                 Inches(6), Inches(3))
    big.name = "big_table"
    for r in range(3):
        for c in range(3):
            big.table.cell(r, c).text = f"r{r}c{c}"

    out = _all_data_shapes(slide)
    assert len(out) == 2
    # Largest first (by name, since shape `is` comparison fails for
    # python-pptx wrapper objects)
    assert out[0][2] == "big_table"
    assert out[1][2] == "small_table"
    # Each entry: (kind, shape, name, summary)
    assert out[0][0] == "table"
    assert out[1][0] == "table"


def test_all_data_shapes_skips_empty_table():
    slide = _make_slide()
    # Add a table but leave all cells blank
    tbl = slide.shapes.add_table(2, 2, Inches(1), Inches(1),
                                 Inches(4), Inches(2))
    out = _all_data_shapes(slide)
    # Empty table → no summary → omitted
    assert all(shp is not tbl for (_kind, shp, _name, _sm) in out)


# ────────────────────────────────────────────────────────────────────────
# _build_prompt — voice adaptation + multi-shape support
# ────────────────────────────────────────────────────────────────────────


def test_prompt_describes_state_when_no_source():
    """Non-connected case (no source data passed) — prompt should still
    work and instruct LLM to describe the current state."""
    p = _build_prompt(
        old_headline="Some existing headline about HCP perceptions.",
        data_summaries=[("chart", "Chart 1", "Categories: ['NA']\nSeries 'X': [0.5]")],
        src_data_summaries=None,
    )
    assert "TALKING HEADER" in p
    assert "Chart 1" in p
    # The SOURCE DATA *block* is absent (the literal string "SOURCE DATA"
    # appears in the instructions but not as a section header).
    assert "SOURCE DATA (what the headline was written about" not in p


def test_prompt_includes_source_when_supplied():
    """Connected case (source supplied) — prompt includes SOURCE DATA block
    so Claude can describe movement."""
    p = _build_prompt(
        old_headline="HCP perceptions trended upward in Q4.",
        data_summaries=[("chart", "Chart 1", "Categories: ['NA']\nSeries 'X': [0.55]")],
        src_data_summaries=[("chart", "Chart 1", "Categories: ['NA']\nSeries 'X': [0.50]")],
    )
    assert "SOURCE DATA" in p
    assert "[0.55]" in p
    assert "[0.50]" in p


def test_prompt_handles_multiple_data_shapes():
    """Multi-shape case — prompt presents every chart + table; Claude
    decides which to call out."""
    p = _build_prompt(
        old_headline="Mixed-component slide narrative.",
        data_summaries=[
            ("chart", "Chart A", "Categories: ['Q1']\nSeries 'X': [0.5]"),
            ("chart", "Chart B", "Categories: ['Q1']\nSeries 'Y': [0.3]"),
            ("table", "Summary Table", "Header | A | B\nRow1 | 47% | 33%"),
        ],
    )
    assert "Chart A" in p
    assert "Chart B" in p
    assert "Summary Table" in p
    # Multi-shape instruction surfaces
    assert "MULTIPLE charts/tables" in p


# ────────────────────────────────────────────────────────────────────────
# refresh_headlines — high-level orchestration
# ────────────────────────────────────────────────────────────────────────


def test_refresh_headlines_processes_non_connected_slides(tmp_path, monkeypatch):
    """Non-connected slide (slide_index NOT in spec) MUST still get
    its talking header rewritten — that's the new contract."""
    # Build a 1-slide deck with a talking header + a chart
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_textbox(slide,
                 "Original talking header that needs to be rewritten by Claude.",
                 left_in=0.3, top_in=0.2, width_in=9.0)
    # Add a small chart so _all_data_shapes returns something
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    cd = CategoryChartData()
    cd.categories = ["A", "B"]
    cd.add_series("S1", [1.0, 2.0])
    slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1), Inches(2), Inches(8), Inches(4), cd)

    src_path = tmp_path / "src.pptx"
    prs.save(src_path)
    # The refreshed deck = same as source for this test
    ref_path = tmp_path / "ref.pptx"
    prs.save(ref_path)
    out_path = tmp_path / "out.pptx"
    # EMPTY spec — slide 0 is "non-connected"
    spec_path = tmp_path / "spec.json"
    spec_path.write_text('{"slides": []}', encoding="utf-8")

    # Stub the LLM call
    monkeypatch.setattr(
        "slidegen.headline_refresh._call_claude",
        lambda prompt, model=None: "REWRITTEN headline from non-connected path."
    )
    monkeypatch.setenv("LLM_API_KEY", "stub")

    updates = refresh_headlines(
        source_pptx=str(src_path),
        refreshed_pptx=str(ref_path),
        spec_path=str(spec_path),
        out_pptx=str(out_path),
    )
    assert len(updates) == 1
    assert updates[0].status == "updated"
    assert updates[0].new_headline == "REWRITTEN headline from non-connected path."
    # The summary should flag this as non-connected
    assert "non-connected" in updates[0].chart_summary


def test_refresh_headlines_processes_table_only_slide(tmp_path, monkeypatch):
    """Slide with a talking header + a TABLE (no chart) — must still
    write a headline."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_textbox(slide,
                 "Talking header for a table-only slide that should still rewrite.",
                 left_in=0.3, top_in=0.2, width_in=9.0)
    tbl = slide.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(6), Inches(3))
    for r in range(3):
        for c in range(3):
            tbl.table.cell(r, c).text = f"v{r}{c}"

    src_path = tmp_path / "src.pptx"; prs.save(src_path)
    ref_path = tmp_path / "ref.pptx"; prs.save(ref_path)
    out_path = tmp_path / "out.pptx"
    spec_path = tmp_path / "spec.json"
    spec_path.write_text('{"slides": [{"slide_index": 0}]}', encoding="utf-8")

    monkeypatch.setattr(
        "slidegen.headline_refresh._call_claude",
        lambda prompt, model=None: "Table-derived rewrite."
    )
    monkeypatch.setenv("LLM_API_KEY", "stub")

    updates = refresh_headlines(
        source_pptx=str(src_path),
        refreshed_pptx=str(ref_path),
        spec_path=str(spec_path),
        out_pptx=str(out_path),
    )
    assert updates[0].status == "updated"
    assert "table" in updates[0].chart_summary


def test_refresh_headlines_skips_when_no_talking_header(tmp_path, monkeypatch):
    """Cover/divider slide with no qualifying talking header — skipped
    silently with status no_headline."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Only a centred short title — fails the > 30 chars filter
    _add_textbox(slide, "PROJECT COVER",
                 left_in=2.0, top_in=3.0, width_in=6.0)

    src_path = tmp_path / "src.pptx"; prs.save(src_path)
    ref_path = tmp_path / "ref.pptx"; prs.save(ref_path)
    out_path = tmp_path / "out.pptx"

    monkeypatch.setattr(
        "slidegen.headline_refresh._call_claude",
        lambda prompt, model=None: "should not be called"
    )
    monkeypatch.setenv("LLM_API_KEY", "stub")

    updates = refresh_headlines(
        source_pptx=str(src_path),
        refreshed_pptx=str(ref_path),
        spec_path=None,
        out_pptx=str(out_path),
    )
    assert updates[0].status == "no_headline"


def test_refresh_headlines_rewrites_even_when_values_unchanged(tmp_path, monkeypatch):
    """Connected slide where chart values are identical pre/post refresh
    — must STILL rewrite (mixed-component slides may have non-connected
    edits we can't see). Old behaviour skipped these as 'unchanged'."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_textbox(slide,
                 "Talking header that should be rewritten even with stable data.",
                 left_in=0.3, top_in=0.2, width_in=9.0)
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    cd = CategoryChartData()
    cd.categories = ["A", "B"]
    cd.add_series("S1", [1.0, 2.0])
    slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(1), Inches(2), Inches(8), Inches(4), cd)

    src_path = tmp_path / "src.pptx"; prs.save(src_path)
    ref_path = tmp_path / "ref.pptx"; prs.save(ref_path)  # same file = same values
    out_path = tmp_path / "out.pptx"

    monkeypatch.setattr(
        "slidegen.headline_refresh._call_claude",
        lambda prompt, model=None: "Rewritten despite identical values."
    )
    monkeypatch.setenv("LLM_API_KEY", "stub")

    updates = refresh_headlines(
        source_pptx=str(src_path),
        refreshed_pptx=str(ref_path),
        spec_path=None,
        out_pptx=str(out_path),
    )
    assert updates[0].status == "updated"
    assert updates[0].new_headline == "Rewritten despite identical values."
