"""
Slide creation engine.

Provides reusable functions for building slides from data using python-pptx.
Every shape is named with a zrx_ prefix and registered in slide_registry.json.

Usage as module:
    from slidegen.create import SlideBuilder
    builder = SlideBuilder(template="path/to/template.pptx")
    slide = builder.add_blank_slide()
    builder.add_chart(slide, chart_type, data, ...)
    builder.save("output.pptx")

Usage as CLI (demo with sample data):
    python -m slidegen.create
"""

import os
import sys
import json
from datetime import datetime

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Emu

from slidegen.config import (
    OUTPUT_DIR, REGISTRY_PATH, TEMPLATE_PATH,
    SLIDE_W_EMU, SLIDE_H_EMU, SHAPE_PREFIX, ensure_output_dir,
)
from slidegen.pptx_utils import (
    C_RYB_Q4, C_RYB_Q3, C_RED, C_GREEN, C_GREY, C_WHITE, C_FTGREY,
    FONT_TEXT, SLIDE_W_IN,
    textbox, solidrect, horiz_line, slide_header, slide_footer,
    add_delta_col, manual_legend,
    invert_cat_axis, set_datalabel_pos_outside_end,
    set_series_no_border, set_val_axis_number_format, set_plot_area_gap,
    set_series_color,
    save_registry, Pt,
)


class SlideBuilder:
    """Builds slides with automatic shape naming and registry tracking."""

    def __init__(self, template=None):
        """Initialize a new presentation.

        Args:
            template: Path to a .pptx template. If None, creates a blank
                      widescreen presentation.
        """
        if template and os.path.exists(template):
            self.prs = Presentation(template)
        else:
            self.prs = Presentation()
            self.prs.slide_width = Emu(SLIDE_W_EMU)
            self.prs.slide_height = Emu(SLIDE_H_EMU)

        self._counter = 0
        self._shapes = {}  # zrx_NNN -> metadata dict

    def _next_name(self):
        """Generate the next sequential zrx_ shape name."""
        self._counter += 1
        return f"{SHAPE_PREFIX}{self._counter:03d}"

    def register(self, shape, label, shape_type, **extra):
        """Name a shape and register it. Returns the assigned name."""
        name = self._next_name()
        shape.name = name
        record = {
            "label": label,
            "type": shape_type,
            "left": round(shape.left / 914400, 4),
            "top": round(shape.top / 914400, 4),
            "width": round(shape.width / 914400, 4),
            "height": round(shape.height / 914400, 4),
        }
        record.update(extra)
        self._shapes[name] = record
        return name

    def add_blank_slide(self, strip_existing=False):
        """Add a blank slide. Optionally strip all template slides."""
        blank_layout = None
        for layout in self.prs.slide_layouts:
            if "blank" in layout.name.lower():
                blank_layout = layout
                break
        if blank_layout is None:
            blank_layout = self.prs.slide_layouts[-1]

        slide = self.prs.slides.add_slide(blank_layout)

        if strip_existing:
            sldIdLst = self.prs.slides._sldIdLst
            for sldId in list(sldIdLst)[:-1]:
                sldIdLst.remove(sldId)

        return slide

    def add_clustered_bar(self, slide, categories, series_data, left, top,
                          width, height, label="Clustered bar chart"):
        """Add a horizontal clustered bar chart.

        Args:
            categories: list of category labels
            series_data: list of (series_name, values, color) tuples
            left, top, width, height: position in inches
            label: semantic label for registry

        Returns:
            (chart_shape, chart) tuple
        """
        chart_data = CategoryChartData()
        chart_data.categories = categories
        for name, values, _ in series_data:
            chart_data.add_series(name, values)

        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(left), Inches(top), Inches(width), Inches(height),
            chart_data,
        )
        self.register(chart_shape, label, "chart",
                      metric=label.lower().replace(" ", "_"))

        chart = chart_shape.chart
        chart.has_legend = False
        invert_cat_axis(chart)
        set_plot_area_gap(chart, gap_pct=60)

        for i, (_, _, color) in enumerate(series_data):
            series = chart.series[i]
            set_series_color(series, color)
            set_series_no_border(series)
            if i == 0:  # Data labels on primary series only
                series.has_data_labels = True
                series.data_labels.font.size = Pt(8)
                series.data_labels.font.color.rgb = C_GREY
                set_datalabel_pos_outside_end(series)

        set_val_axis_number_format(chart.value_axis, "0")

        return chart_shape, chart

    def add_delta_column(self, slide, deltas, left, top, width, height,
                         header="Delta", label="Delta table"):
        """Add a delta table aligned with a chart.

        Args:
            deltas: list of float|None
            left, top, width, height: position in inches
            header: column header text
            label: semantic label for registry

        Returns:
            table object
        """
        # Track shapes before to find the new table shape after
        before = {sh.shape_id for sh in slide.shapes}
        tbl = add_delta_col(slide, deltas, left, top, width, height, header)

        for sh in slide.shapes:
            if sh.shape_id not in before and sh.has_table:
                self.register(sh, label, "table")
                break

        return tbl

    def add_textbox(self, slide, text, left, top, width, height,
                    label="Text box", **kwargs):
        """Add a registered text box.

        Accepts all keyword arguments that pptx_utils.textbox() accepts
        (fsize, bold, color, align, italic, wrap, font).

        Returns:
            shape object
        """
        shape = textbox(slide, text, left, top, width, height, **kwargs)
        self.register(shape, label, "textbox",
                      text=text[:80] if text else "")
        return shape

    def add_rect(self, slide, left, top, width, height, fill,
                 label="Rectangle", **kwargs):
        """Add a registered filled rectangle.

        Returns:
            shape object
        """
        shape = solidrect(slide, left, top, width, height, fill, **kwargs)
        self.register(shape, label, "rect")
        return shape

    def save(self, output_path):
        """Save the presentation and write the registry.

        Args:
            output_path: path for the .pptx file

        Returns:
            dict with keys: pptx_path, registry_path, shape_count
        """
        ensure_output_dir()
        self.prs.save(output_path)

        registry = {
            "meta": {
                "created": datetime.now().isoformat(timespec="seconds"),
                "pptx_file": os.path.abspath(output_path),
                "shape_count": len(self._shapes),
            },
            "shapes": self._shapes,
        }
        save_registry(registry)

        return {
            "pptx_path": os.path.abspath(output_path),
            "registry_path": REGISTRY_PATH,
            "shape_count": len(self._shapes),
        }

    @property
    def shapes(self):
        """Return a copy of the current shape registry."""
        return dict(self._shapes)


# ── CLI demo ─────────────────────────────────────────────────────────────────

def _demo():
    """Create a demo slide with sample data to verify the pipeline works."""
    messages = [
        "RYBREVANT + LAZCLUZE is the first-in-class combo",
        "Demonstrated OS benefit in 1L EGFRm NSCLC",
        "Statistically significant PFS improvement",
        "Well-managed safety profile",
        "MARIPOSA trial results",
    ]
    q4 = [62.1, 55.3, 48.7, 41.2, 38.9]
    q3 = [58.4, 50.1, 45.2, 39.8, 35.0]
    deltas = [round(a - b, 1) for a, b in zip(q4, q3)]

    builder = SlideBuilder()
    slide = builder.add_blank_slide()

    slide_header(slide, "Message Recall \u2014 Rybrevant + Lazcluze")

    builder.add_clustered_bar(
        slide, messages,
        series_data=[("Q4'25", q4, C_RYB_Q4), ("Q3'25", q3, C_RYB_Q3)],
        left=0.20, top=1.50, width=8.00, height=4.80,
        label="MR clustered bar chart",
    )

    builder.add_delta_column(
        slide, deltas,
        left=8.30, top=1.50, width=0.62, height=4.80,
        header="MR\nDelta", label="MR delta table",
    )

    builder.add_textbox(
        slide,
        "Key Insight: First-in-class messaging shows strongest recall "
        "improvement (+3.7pp) driven by MARIPOSA data readout.",
        left=9.20, top=1.50, width=3.80, height=1.20,
        label="Key insight callout", fsize=10, color=C_GREY,
    )

    builder.add_rect(
        slide, 9.20, 3.00, 3.80, 0.06, C_RYB_Q4,
        label="Orange accent divider",
    )

    builder.add_textbox(
        slide,
        "Sample sizes:\n  Q4'25: n=132\n  Q3'25: n=128\n\n"
        "Source: Lung SFEA Survey",
        left=9.20, top=3.20, width=3.80, height=0.80,
        label="Sample size and source", fsize=8, color=C_FTGREY,
    )

    slide_footer(slide, "Source: ZoomRx Lung SFEA Survey, Q3-Q4 2025.")

    out = os.path.join(OUTPUT_DIR, "demo_slide.pptx")
    result = builder.save(out)

    print(f"[OK] Created: {result['pptx_path']}")
    print(f"[OK] Registry: {result['registry_path']} ({result['shape_count']} shapes)")
    for name, rec in builder.shapes.items():
        print(f"     {name}: {rec['label']} ({rec['type']})")
    print(f"\nOpen in PowerPoint, then run: python -m slidegen.edit demo_slide.pptx")


if __name__ == "__main__":
    _demo()
