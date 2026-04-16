"""
Incremental-construction debug for the line-chart Repair bug.

Produces 3 PPTX files, each applying slide_creator operations up to a specific
checkpoint. User opens each in PowerPoint and reports whether it triggers the
Repair prompt. The first file that breaks identifies the culprit operation.

Checkpoints:
  v06 — through legend manipulation (_hide_legend_but_keep_element)
  v09 — through _apply_data_labels (redundant second pass over series labels)
  v11 — through _enforce_ser_child_order (plotArea/chart-type/ser reorder)

Outputs: experiments/deck_analysis/outputs/generated/debug_line_v06.pptx, v09.pptx, v11.pptx
"""
from __future__ import annotations
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

from slidegen.pptx_utils.lxml_helpers import (
    set_series_color,
    set_series_marker,
    set_series_line_style,
    set_data_label_color,
    set_datalabel_format,
    set_datalabel_pos_top,
    _get_or_add,
)
from slidegen.pptx_utils.charts import enable_data_labels

OUT_DIR = REPO / "experiments" / "deck_analysis" / "outputs" / "generated"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def _build_baseline_chart():
    """Steps 1-5: raw add_chart + series color/marker/line + data labels + axis scale.

    This matches what pptx_utils.add_line_chart does, without the later
    slide_creator chrome/reorder passes.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.500)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Step 1 — baseline add_chart (LINE_MARKERS = line with circle markers)
    cats = ["Q1 '26", "Q2 '26", "Q3 '26", "Q4 '26"]
    data = CategoryChartData()
    data.categories = cats
    data.add_series("Rybrevant", [0.30, 0.35, 0.38, 0.42])
    data.add_series("Tagrisso", [0.50, 0.48, 0.46, 0.44])

    cf = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(1), Inches(1.5), Inches(11), Inches(5),
        data,
    )
    chart = cf.chart

    # Step 2 — set series colors
    colors = [
        RGBColor(0xF7, 0x58, 0x24),  # Rybrevant orange
        RGBColor(0x70, 0x30, 0xA0),  # Tagrisso purple
    ]
    for i, series in enumerate(chart.series):
        set_series_color(series, colors[i])

    # Step 3 — markers + line styles
    for i, series in enumerate(chart.series):
        set_series_marker(series, "circle", 7, fill_color=colors[i], line_color=colors[i])
        set_series_line_style(series, width_pt=2.25)

    # Step 4 — data labels (enable_data_labels creates series-level dLbls)
    for i, series in enumerate(chart.series):
        pos = "t" if i == 0 else "b"
        enable_data_labels(series, colors[i], fsize=9, num_fmt='0"%"', pos=pos)

    # Step 5 — value axis scale
    ax = chart.value_axis._element
    scaling = _get_or_add(ax, "c:scaling")
    _get_or_add(scaling, "c:min").set("val", "0.0")
    _get_or_add(scaling, "c:max").set("val", "0.6")

    return prs, slide, chart


# ─────────────────────────────────────────────────────────────────────────────
# Operations to test
# ─────────────────────────────────────────────────────────────────────────────


def _op_06_legend_hide(chart):
    """Step 6 — slide-creator's _hide_legend_but_keep_element.

    Sets has_legend=True, adds <c:legendPos val="b"/>, injects zero-size layout.
    This is what replaced the original has_legend=False call.
    """
    from slidegen.slide_creator import _hide_legend_but_keep_element
    chart.has_legend = True
    _hide_legend_but_keep_element(chart)


def _op_07_title_off(chart):
    """Step 7 — chart.has_title = False."""
    chart.has_title = False


def _op_08_gridlines_off(chart):
    """Step 8 — disable major gridlines on both axes."""
    for ax_name in ("value_axis", "category_axis"):
        ax = getattr(chart, ax_name, None)
        if ax is not None:
            try:
                ax.has_major_gridlines = False
            except Exception:
                pass


def _op_09_apply_data_labels(chart):
    """Step 9 — slide-creator's _apply_data_labels (redundant second pass).

    Touches series.data_labels.show_value, set_datalabel_format, position,
    and color. This runs AFTER enable_data_labels already created the dLbls.
    """
    label_color = RGBColor(0xFF, 0xFF, 0xFF)  # example; real code uses spec
    for series in chart.series:
        try:
            series.data_labels.show_value = True
            set_datalabel_format(series, "0%")
            set_datalabel_pos_top(series)
            set_data_label_color(series, label_color)
        except Exception:
            pass


def _op_10_remove_auto_dLbls(chart):
    """Step 10 — _remove_auto_dLbls_on_chart_type.

    Strips chart-type-level <c:dLbls> (line charts don't want this per raw output).
    """
    from slidegen.slide_creator import _remove_auto_dLbls_on_chart_type
    _remove_auto_dLbls_on_chart_type(chart)


def _op_11_enforce_order(chart):
    """Step 11 — _enforce_ser_child_order.

    Reorders children at plotArea + chart-type + ser levels to OOXML schema.
    """
    from slidegen.slide_creator import _enforce_ser_child_order
    _enforce_ser_child_order(chart)


# ─────────────────────────────────────────────────────────────────────────────
# Build 3 checkpoint files
# ─────────────────────────────────────────────────────────────────────────────


def build_v06():
    """Apply through step 6 (legend manipulation). Stop there."""
    prs, slide, chart = _build_baseline_chart()
    _op_06_legend_hide(chart)
    path = OUT_DIR / "debug_line_v06_legend.pptx"
    prs.save(str(path))
    return path


def build_v09():
    """Apply through step 9 (after _apply_data_labels redundant pass)."""
    prs, slide, chart = _build_baseline_chart()
    _op_06_legend_hide(chart)
    _op_07_title_off(chart)
    _op_08_gridlines_off(chart)
    _op_09_apply_data_labels(chart)
    path = OUT_DIR / "debug_line_v09_data_labels.pptx"
    prs.save(str(path))
    return path


def build_v11():
    """Apply through step 11 (after full reorder). Almost complete slide_creator output."""
    prs, slide, chart = _build_baseline_chart()
    _op_06_legend_hide(chart)
    _op_07_title_off(chart)
    _op_08_gridlines_off(chart)
    _op_09_apply_data_labels(chart)
    _op_10_remove_auto_dLbls(chart)
    _op_11_enforce_order(chart)
    path = OUT_DIR / "debug_line_v11_reorder.pptx"
    prs.save(str(path))
    return path


def build_v12():
    """Apply through step 12 (adds chrome text boxes: headline, subheadline, footer, section bar).

    This is the full slide_creator flow for a line chart spec. If v12 triggers
    Repair but v11 does not, the chrome text-box insertions are the remaining
    culprit.
    """
    prs, slide, chart = _build_baseline_chart()
    _op_06_legend_hide(chart)
    _op_07_title_off(chart)
    _op_08_gridlines_off(chart)
    _op_09_apply_data_labels(chart)
    _op_10_remove_auto_dLbls(chart)
    _op_11_enforce_order(chart)

    # Step 12 — chrome text boxes (what slide_creator does in _apply_headline /
    # _apply_subheadline / _apply_footer / _apply_section_bar)
    from slidegen.pptx_utils.layout import slide_header, slide_footer, section_header_bar
    from slidegen.pptx_utils.shapes import textbox
    from pptx.enum.text import PP_ALIGN

    section_header_bar(slide, "Awareness & Recall Trend", top=1.40)
    slide_header(slide, "Rybrevant unaided recall has grown 12pp over the past four waves")
    textbox(slide,
            "Unaided brand recall, rolling three-month window — Q2 '25 through Q1 '26",
            0.30, 0.80, 12.70, 0.40,
            fsize=10)
    slide_footer(slide,
                 "Source: Rybrevant PET, R3M rolling among NSCLC prescribers. Base: all respondents with ≥1 rep interaction.")

    path = OUT_DIR / "debug_line_v12_full_chrome.pptx"
    prs.save(str(path))
    return path


def main() -> int:
    print("Generating 3 incremental debug files for line-chart Repair bug.")
    print()
    v06 = build_v06()
    print(f"  [v06 — legend hide ONLY]              {v06.name}  ({v06.stat().st_size} bytes)")
    v09 = build_v09()
    print(f"  [v09 — + apply_data_labels pass]      {v09.name}  ({v09.stat().st_size} bytes)")
    v11 = build_v11()
    print(f"  [v11 -- remove auto dLbls + reorder] {v11.name}  ({v11.stat().st_size} bytes)")
    v12 = build_v12()
    print(f"  [v12 -- + chrome text boxes (full)]  {v12.name}  ({v12.stat().st_size} bytes)")
    print()
    print("Open v12 and report whether it triggers the Repair prompt.")
    print("  v12 broken + v11 clean -> chrome text boxes are the culprit")
    print("  v12 clean -> bug is elsewhere (e.g., real render_slide path adds something debug doesn't)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
