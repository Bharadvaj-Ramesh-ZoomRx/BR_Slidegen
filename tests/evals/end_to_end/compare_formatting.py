"""Step 3 formatting comparison: source vs refreshed decks.

Reports per-chart whether chart_type and per-series fill colors are
preserved under refresh. Step 3 of Bharadvaj's 7-step refresh framework
("Formatting intact: colors, shapes, chart types") — independent of the
data-level signals captured in compare_decks.py (Step 2).

Each ComponentMatch records:
    chart_type_match     — bool: identical pptx ChartType enum value
    series_colors_match  — bool: per-series color tokens (rgb hex,
                           theme color, or "INHERIT") match positionally

Tables are skipped — formatting comparison is chart-only for now.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

from .compare_decks import _find_match, _pos


@dataclass
class FormatMatch:
    slide_idx: int
    position: str
    chart_type_match: bool | None = None
    series_colors_match: bool | None = None
    note: str = ""


@dataclass
class FormatReport:
    source_deck: str
    refreshed_deck: str
    components: list[FormatMatch] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "source_deck": self.source_deck,
            "refreshed_deck": self.refreshed_deck,
            "totals": self.totals(),
            "components": [c.__dict__ for c in self.components],
        }

    def totals(self) -> dict:
        total = len(self.components)
        type_ok = sum(1 for c in self.components if c.chart_type_match is True)
        col_ok = sum(1 for c in self.components if c.series_colors_match is True)
        errors = sum(1 for c in self.components if c.note)
        return {
            "components_compared": total,
            "chart_type_match": type_ok,
            "series_colors_match": col_ok,
            "components_with_errors": errors,
        }


def _color_token(series) -> str:
    """Return a stable string token for a series' fill color.

    Falls back from explicit rgb to theme color to a generic INHERIT
    sentinel so two series with the same effective styling compare
    equal regardless of which OOXML representation the source used.
    """
    try:
        fill = series.format.fill
        if fill.type is None:
            return "INHERIT"
        try:
            rgb = fill.fore_color.rgb
            return f"rgb:{rgb}"
        except Exception:
            pass
        try:
            tc = fill.fore_color.theme_color
            return f"theme:{tc}"
        except Exception:
            return "UNKNOWN"
    except Exception:
        return "NOFILL"


def compare_formatting(source_path: Path, refreshed_path: Path) -> FormatReport:
    """Run chart-formatting comparison between source and refreshed decks."""
    source = Presentation(str(source_path))
    refreshed = Presentation(str(refreshed_path))

    report = FormatReport(
        source_deck=source_path.name,
        refreshed_deck=refreshed_path.name,
    )

    n_slides = min(len(source.slides), len(refreshed.slides))
    for slide_idx in range(n_slides):
        src_charts = [s for s in source.slides[slide_idx].shapes if s.has_chart]
        ref_charts = [s for s in refreshed.slides[slide_idx].shapes if s.has_chart]

        for src in src_charts:
            sl, st = _pos(src)
            match = _find_match(src, ref_charts)
            comp = FormatMatch(
                slide_idx=slide_idx,
                position=f"({sl},{st})",
            )
            if match is None:
                comp.note = "no matching refreshed chart at this position"
                report.components.append(comp)
                continue

            try:
                comp.chart_type_match = (
                    src.chart.chart_type == match.chart.chart_type
                )
                src_cols = [_color_token(s) for s in src.chart.plots[0].series]
                ref_cols = [_color_token(s) for s in match.chart.plots[0].series]
                comp.series_colors_match = src_cols == ref_cols
            except Exception as e:
                comp.note = f"format extraction error: {type(e).__name__}: {e}"

            report.components.append(comp)

    return report
