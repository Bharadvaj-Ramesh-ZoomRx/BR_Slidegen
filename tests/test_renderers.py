"""
Unit tests for slide renderers.

Tests each renderer with minimal config + synthetic data to verify:
- No crashes on valid input
- Shapes are actually added to the slide
- Empty/missing data produces error placeholder (not crash)
"""

import os
import sys
import pytest
from dataclasses import dataclass, field
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches

BASE = os.path.dirname(__file__)
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)

from slidegen.pipeline.project_config import (
    ProjectConfig, BrandConfig, SheetConfig, SampleSize, AskConfig, parse_color,
)
from slidegen.pipeline.slide_renderers import RENDERERS


# ── Test fixtures ───────────────────────────────────────────────────────────

@pytest.fixture
def minimal_config():
    """Create a minimal ProjectConfig for renderer tests."""
    return ProjectConfig(
        name="Test Project",
        client="TestCorp",
        period_current="Q1'26",
        period_prior="Q4'25",
        brands={
            "primary": BrandConfig(
                name="RYB", full_name="Rybrevant",
                color_current=parse_color("#F75824"),
                color_prior=parse_color("#FFC199"),
            ),
            "competitor": BrandConfig(
                name="TAG", full_name="Tagrisso",
                color_current=parse_color("#7030A0"),
                color_prior=parse_color("#AD88C8"),
            ),
        },
        fonts={"display": "Calibri", "body": "Calibri"},
        sheets={
            "primary": SheetConfig(name="Sheet1", q_prior_col=7, q_current_col=13),
        },
        sample_sizes={
            "primary": SampleSize(prior=100, current=100),
        },
        extractions=[],
        asks=[],
    )


@pytest.fixture
def blank_slide():
    """Create a blank Presentation + slide for renderer testing."""
    prs = Presentation()
    # Set widescreen dimensions (13.33" x 7.5")
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(layout)
    return prs, slide


def _make_ask(slide_type, data_key="test_data", headline="Test Headline",
              section="Test Section", extra=None):
    """Create a minimal AskConfig for testing."""
    return AskConfig(
        id=f"test_{slide_type}",
        slide_type=slide_type,
        headline=headline,
        section=section,
        source_text="Source: Test",
        data_key=data_key,
        extra=extra or {},
    )


def _count_shapes(slide):
    """Count non-placeholder shapes on a slide."""
    return len(slide.shapes)


# ── Sample data for different renderer types ────────────────────────────────

SINGLE_BAR_DATA = {
    "test_data": [
        {"desc": "Message A", "current": 75.0, "prior": 70.0},
        {"desc": "Message B", "current": 60.0, "prior": 55.0},
        {"desc": "Message C", "current": 45.0, "prior": 50.0},
    ],
}

DUAL_BAR_DATA = {
    "test_data": [
        {"desc": "Msg A", "mr_current": 75.0, "mr_prior": 70.0,
         "me_current": 60.0, "me_prior": 55.0},
        {"desc": "Msg B", "mr_current": 50.0, "mr_prior": 45.0,
         "me_current": 40.0, "me_prior": 35.0},
    ],
}

CLUSTERED_DATA = {
    "primary_data": [
        {"desc": "Item A", "current": 80.0, "prior": 75.0},
        {"desc": "Item B", "current": 60.0, "prior": 55.0},
    ],
    "comp_data": [
        {"desc": "Item A", "current": 70.0, "prior": 65.0},
        {"desc": "Item B", "current": 50.0, "prior": 45.0},
    ],
}

STACKED_DATA = {
    "test_data": [
        {"desc": "Msg A", "code": "A1",
         "1st_current": 30.0, "1st_prior": 25.0,
         "2nd_current": 20.0, "2nd_prior": 18.0,
         "3rd_current": 10.0, "3rd_prior": 12.0,
         "total_current": 60.0, "total_prior": 55.0},
        {"desc": "Msg B", "code": "B1",
         "1st_current": 25.0, "1st_prior": 20.0,
         "2nd_current": 15.0, "2nd_prior": 13.0,
         "3rd_current": 8.0, "3rd_prior": 10.0,
         "total_current": 48.0, "total_prior": 43.0},
    ],
}

ABACUS_DATA = {
    "test_data": [
        {"desc": "Attribute A", "current": 75.0, "prior": 70.0},
        {"desc": "Attribute B", "current": 60.0, "prior": 55.0},
        {"desc": "Attribute C", "current": 45.0, "prior": 50.0},
    ],
}

COVER_DATA = {"test_data": []}

ES_DATA = {"test_data": []}

HEATMAP_DATA = {
    "test_data": [
        {"label": "Row A", "values": {"Col1": 80, "Col2": 60}, "deltas": {"Col1": 5.0, "Col2": -3.0}},
        {"label": "Row B", "values": {"Col1": 50, "Col2": 70}, "deltas": {"Col1": -2.0, "Col2": 8.0}},
    ],
}


# ── Renderer tests ──────────────────────────────────────────────────────────

class TestCoverRenderer:
    def test_cover_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("cover", extra={
            "subtitle": "Test Subtitle",
            "date": "Q1 2026",
            "client": "TestCorp",
        })
        render = RENDERERS["cover"]
        render(slide, minimal_config, ask, COVER_DATA)
        assert _count_shapes(slide) > 0

    def test_cover_with_templates(self, blank_slide, minimal_config):
        """Cover should resolve {{primary.name}} templates."""
        prs, slide = blank_slide
        ask = _make_ask("cover", headline="{{primary.name}} Report",
                        extra={"subtitle": "Wave {{period_current}}", "date": "", "client": ""})
        RENDERERS["cover"](slide, minimal_config, ask, COVER_DATA)
        assert _count_shapes(slide) > 0


class TestExecutiveSummaryRenderer:
    def test_es_with_insights(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("executive_summary", extra={
            "insights": [
                "First key finding about the market.",
                "Second finding with data support.",
                "Third insight for recommendations.",
            ],
        })
        RENDERERS["executive_summary"](slide, minimal_config, ask, ES_DATA)
        assert _count_shapes(slide) >= 3  # header + card elements

    def test_es_no_insights_shows_error(self, blank_slide, minimal_config):
        """Empty insights should show error placeholder, not crash."""
        prs, slide = blank_slide
        ask = _make_ask("executive_summary", extra={})
        RENDERERS["executive_summary"](slide, minimal_config, ask, ES_DATA)
        # Should have at least the header + error textbox
        assert _count_shapes(slide) >= 1


class TestSingleBarRenderer:
    def test_single_bar_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("single_bar_with_delta")
        RENDERERS["single_bar_with_delta"](slide, minimal_config, ask, SINGLE_BAR_DATA)
        # Should have: header, section bar, footer, label table, chart, delta table, legend
        assert _count_shapes(slide) >= 5

    def test_single_bar_empty_data(self, blank_slide, minimal_config):
        """Empty data should show error placeholder."""
        prs, slide = blank_slide
        ask = _make_ask("single_bar_with_delta")
        RENDERERS["single_bar_with_delta"](slide, minimal_config, ask, {"test_data": []})
        assert _count_shapes(slide) >= 1

    def test_single_bar_missing_key(self, blank_slide, minimal_config):
        """Missing data key should show error placeholder."""
        prs, slide = blank_slide
        ask = _make_ask("single_bar_with_delta", data_key="nonexistent")
        RENDERERS["single_bar_with_delta"](slide, minimal_config, ask, SINGLE_BAR_DATA)
        assert _count_shapes(slide) >= 1

    def test_single_bar_sort(self, blank_slide, minimal_config):
        """Verify sort_by works without crashing."""
        prs, slide = blank_slide
        ask = _make_ask("single_bar_with_delta")
        ask.sort_by = "current"
        ask.sort_desc = True
        RENDERERS["single_bar_with_delta"](slide, minimal_config, ask, SINGLE_BAR_DATA)
        assert _count_shapes(slide) >= 5


class TestDualBarRenderer:
    def test_dual_bar_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("dual_bar_with_delta", extra={
            "left": {"field_prefix": "mr", "label": "Recall"},
            "right": {"field_prefix": "me", "label": "Effectiveness"},
        })
        RENDERERS["dual_bar_with_delta"](slide, minimal_config, ask, DUAL_BAR_DATA)
        assert _count_shapes(slide) >= 5

    def test_dual_bar_empty_data(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("dual_bar_with_delta", extra={
            "left": {"field_prefix": "mr"},
            "right": {"field_prefix": "me"},
        })
        RENDERERS["dual_bar_with_delta"](slide, minimal_config, ask, {"test_data": []})
        assert _count_shapes(slide) >= 1


class TestClusteredCompareRenderer:
    def test_clustered_compare_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("clustered_compare", extra={
            "primary_key": "primary_data",
            "comp_key": "comp_data",
        })
        RENDERERS["clustered_compare"](slide, minimal_config, ask, CLUSTERED_DATA)
        assert _count_shapes(slide) >= 5


class TestStackedOrderRenderer:
    def test_stacked_order_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("stacked_order", extra={
            "ordinals": ["1st", "2nd", "3rd"],
        })
        RENDERERS["stacked_order"](slide, minimal_config, ask, STACKED_DATA)
        assert _count_shapes(slide) >= 5


class TestAbacusRenderer:
    def test_abacus_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("abacus")
        RENDERERS["abacus"](slide, minimal_config, ask, ABACUS_DATA)
        assert _count_shapes(slide) >= 5

    def test_abacus_empty_data(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("abacus")
        RENDERERS["abacus"](slide, minimal_config, ask, {"test_data": []})
        assert _count_shapes(slide) >= 1


class TestHeatmapRenderer:
    def test_heatmap_renders(self, blank_slide, minimal_config):
        prs, slide = blank_slide
        ask = _make_ask("heatmap_table", extra={
            "columns": ["Col1", "Col2"],
        })
        RENDERERS["heatmap_table"](slide, minimal_config, ask, HEATMAP_DATA)
        assert _count_shapes(slide) >= 3


class TestRendererRegistry:
    """Verify the RENDERERS registry is complete and consistent."""

    def test_all_renderers_are_callable(self):
        for name, fn in RENDERERS.items():
            assert callable(fn), f"RENDERERS['{name}'] is not callable"

    def test_no_duplicate_functions(self):
        """Each renderer function should be unique (except backward-compat aliases)."""
        seen = {}
        aliases = {"dual_brand_compare"}  # known alias
        for name, fn in RENDERERS.items():
            if name in aliases:
                continue
            fn_id = id(fn)
            assert fn_id not in seen, f"RENDERERS['{name}'] is a duplicate of '{seen[fn_id]}'"
            seen[fn_id] = name

    def test_renderer_count(self):
        """Verify we have 20 unique renderers + 1 alias = 21 total."""
        assert len(RENDERERS) == 21


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
