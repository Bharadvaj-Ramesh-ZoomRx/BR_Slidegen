"""
pptx_utils — SlideGen shared utility package.

Re-exports everything from submodules so existing imports like
``from slidegen.pptx_utils import textbox, C_RED`` continue working.

Submodules:
    brand          — Client brand definitions, slide constants, color palette
    lxml_helpers   — Raw OOXML XML manipulation functions
    shapes         — Layout primitives (textbox, solidrect, etc.)
    layout         — Slide chrome, decorations, LAYOUTS{} dict
    charts         — High-level chart builders, CHART_PATTERNS{} dict
    tables         — Table builders (delta columns, value columns)
    com            — COM helpers for live editing via win32com
    registry       — Shape registry CRUD operations
"""

# ── Brand constants ──────────────────────────────────────────────────────────
from .brand import (
    BRAND,
    SLIDE_W_IN, SLIDE_H_IN, SLIDE_W_EMU, SLIDE_H_EMU,
    IN, EMU_PER_IN,
    C_RYB_Q4, C_RYB_Q3, C_TAG,
    C_RED, C_GREEN, C_WHITE, C_GREY, C_FTGREY, C_LBGREY, C_HDRGREY, C_LTGREY,
    COM_RED, COM_ORANGE, COM_GREEN, COM_GREY,
    FONT_DISPLAY, FONT_TEXT,
)

# ── lxml helpers ─────────────────────────────────────────────────────────────
from .lxml_helpers import (
    _get_or_add, suppress_para_bullets, suppress_cat_axis_bullets, cell_vcenter,
    invert_cat_axis, hide_cat_labels,
    set_datalabel_pos_outside_end, set_series_no_border,
    set_val_axis_number_format, set_plot_area_gap, set_overlap,
    set_series_marker, set_series_line_style, set_series_smooth,
    set_marker_data_label_pos, set_data_label_color,
    add_val_axis_reference_line, set_stacked_label_pos,
    set_pie_slice_colors, set_donut_hole_size,
    hide_axis, set_gridlines, set_series_color,
    set_chart_plot_area, set_val_axis_scale,
)

# ── Shape primitives ─────────────────────────────────────────────────────────
from .shapes import (
    textbox, solidrect, horiz_line, insert_image,
    dashed_separator, stat_callout, callout_box,
)

# ── Layout & slide chrome ────────────────────────────────────────────────────
from .layout import (
    LAYOUTS,
    slide_header, slide_footer, manual_legend,
    section_header_bar, module_badge, section_breadcrumb,
    chart_header_row,
    divider_slide, cover_slide,
    trend_arrow_icon, scatter_quadrant_fills,
)

# ── Chart builders ───────────────────────────────────────────────────────────
from .charts import (
    CHART_PATTERNS,
    enable_data_labels, delete_data_label,
    add_single_bar_chart, add_clustered_bar_chart,
)

# ── Table builders ───────────────────────────────────────────────────────────
from .tables import (
    add_delta_col, add_delta_table, add_value_table,
)

# ── COM helpers ──────────────────────────────────────────────────────────────
from .com import (
    com_connect, com_find_shape, com_set_text,
    com_set_fill, com_move, com_resize, com_get_position,
)

# ── Registry helpers ─────────────────────────────────────────────────────────
from .registry import (
    load_registry, save_registry, registry_get,
    registry_tag_slide, registry_find_by_type, registry_diff_slide,
)

# ── Re-export pptx.util for convenience ──────────────────────────────────────
from pptx.util import Inches, Pt, Emu
