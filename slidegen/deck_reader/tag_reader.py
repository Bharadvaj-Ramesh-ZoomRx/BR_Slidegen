"""
tag_reader.py — Tier 1: Extract SlideSpecs from Connector-tagged PPTX shapes.

Reads shape tags (ReportConfigHash, DataFrameConfigHash, MappingConfig,
ColumnKeyLabelMap, AnalysisType, LastRefreshTime, RefreshErrorMsg) and
looks up Custom XML Parts by hash to recover the full ReportConfigDto /
PivotConfigDto JSON. Emits SlideSpec with fully-populated DataLineage.

Connector tag model reference:
    - Shape tags: Constants.ShapeTags in galen-powerpoint/Constants.cs
    - ReportConfigHash -> SHA256 key into Custom XML Part "ReportConfig" store
    - DataFrameConfigHash -> SHA256 key into Custom XML Part "DataFrameConfigKey" store
    - MappingConfig -> JSON string stored directly on the shape tag

CLI:
    python -m slidegen.deck_reader path/to/deck.pptx --out specs/
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Optional
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.util import Inches, Emu

from slidegen.slide_spec.schema import (
    SlideSpec, HeadlineSpec, FooterSpec, Position, DataLineage, DataLineageCandidate,
    SlideMetadata, ChartComponent, LabelTableComponent, ValueTableComponent,
    TextboxComponent, ChartData, Series, ChartChrome, DataLabelsSpec, LegendSpec, AxisSpec,
    DataFilter, DataTransform, SeriesConfig, ChartDataMapping,
    TableColumnConfig, TableDataMapping, SegmentRule,
    dump_spec, SPEC_VERSION,
)


# ── Constants matching galen-powerpoint/Constants.cs ShapeTags ────────────────

# Tag names — UPPERCASE in real PPTX files (Galen-PowerPoint Connector convention).
# Comparison is case-insensitive throughout tag_reader to handle mixed-case edge cases.
TAG_REPORT_CONFIG_HASH = "REPORTCONFIGHASH"
TAG_REPORT_CONFIG_HASH_BACKUP = "REPORTCONFIGHASH_BACKUP"
TAG_PIVOT_CONFIG_HASH = "DATAFRAMECONFIGHASH"
TAG_PIVOT_CONFIG_HASH_BACKUP = "DATAFRAMECONFIGHASH_BACKUP"
TAG_MAPPING_CONFIG = "MAPPINGCONFIG"
TAG_REFRESH_ERROR = "REFRESHERRORMSG"
TAG_LAST_REFRESH = "LASTREFRESHTIME"
TAG_COLUMN_KEY_LABEL_MAP = "COLUMNKEYLABELMAP"
TAG_ANALYSIS_TYPE = "ANALYSISTYPE"
TAG_SPLIT_GROUP_ID = "SPLITGROUPID"
TAG_SPLIT_ORDER = "SPLITORDER"

# Custom XML Part store keys (matching Constants.CustomXMLStore.Keys)
XML_STORE_REPORT_CONFIG = "ReportConfig"
XML_STORE_PIVOT_CONFIG = "DataFrameConfigKey"


@dataclass
class TagReaderSummary:
    """Summary of Tier 1 tag reading."""
    total_slides: int = 0
    total_shapes: int = 0
    tagged_shapes: int = 0
    untagged_shapes: int = 0
    custom_xml_parts_found: int = 0
    report_configs_resolved: int = 0
    pivot_configs_resolved: int = 0
    mapping_configs_resolved: int = 0
    # Fallback path: per-shape *_BACKUP tag JSON used when the deck-wide
    # customXml store has no entry for the hash. Connector's migration sometimes
    # writes the per-shape tag fully but skips the customXml store entry —
    # without this fallback those shapes look "non-connected" even though
    # their tag is intact.
    report_configs_from_backup: int = 0
    pivot_configs_from_backup: int = 0
    per_slide: dict = field(default_factory=dict)  # slide_index -> {tagged, untagged}


@dataclass
class UntaggedShape:
    """Info about an untagged shape passed to Tier 2."""
    slide_index: int
    shape_name: str
    shape_type: str  # "chart", "table", "textbox", "picture", "group", "other"
    left_emu: int = 0
    top_emu: int = 0
    width_emu: int = 0
    height_emu: int = 0
    text: str = ""
    # For charts: basic OOXML info
    chart_type: Optional[str] = None
    has_chart: bool = False
    has_table: bool = False


def _get_shape_tag(shape, tag_name: str) -> Optional[str]:
    """Read a user-defined tag from a python-pptx shape.

    python-pptx stores user tags as <p:tag> elements under the shape's
    <p:tagLst> relationship. We walk the shape's XML to find them.
    """
    try:
        # python-pptx shapes don't natively expose user tags.
        # Tags live in a separate part referenced by the shape.
        # We need to parse the shape XML for custDataLst / tag references.
        sp = shape._element

        # Look for tag elements in the shape's part relationships
        # Shape tags in PPTX are stored in /ppt/tags/tagN.xml files,
        # linked via <mc:Choice>/<p:custDataLst> in the shape XML.
        # However, the Galen Connector uses a different approach:
        # it stores tags as user-defined properties on the shape.

        # Try the python-pptx approach to reading shape tags
        # Shape._element may have <p:extLst> with custom data
        from pptx.oxml.ns import qn

        # Method 1: Check for custDataLst links to tag parts.
        # <p:custDataLst> lives inside <p:nvPr> which is nested inside
        # <p:nvSpPr> (for sp shapes) or <p:nvGraphicFramePr> (for charts/tables).
        # Search recursively to handle all shape types.
        cust_data = sp.findall('.//' + qn('p:custDataLst'))

        if not cust_data:
            return None

        # Walk through custDataLst to find tag parts
        for cdl in cust_data:
            for tag_ref in cdl:
                # Get the relationship id
                r_id = tag_ref.get(qn('r:id'))
                if r_id is None:
                    continue

                # Resolve the tag part via the slide's relationships
                try:
                    # Resolve via slide part's relationship collection
                    # (NOT shape.part.related_parts which doesn't exist)
                    slide_part = shape.part
                    rel = slide_part.rels[r_id]
                    tag_part = rel.target_part
                    if tag_part is None:
                        continue

                    # Parse the tag XML
                    tag_xml = ET.fromstring(tag_part.blob)
                    for tag_el in tag_xml:
                        name = tag_el.get('name', '')
                        val = tag_el.get('val', '')
                        if name.upper() == tag_name.upper():
                            return val if val else None
                except Exception:
                    continue

        return None
    except Exception:
        return None


def _get_shape_tags_all(shape) -> dict[str, str]:
    """Read ALL user-defined tags from a shape. Returns {name: value} dict."""
    tags = {}
    try:
        from pptx.oxml.ns import qn
        sp = shape._element

        # Search recursively — custDataLst is nested inside nvPr
        cust_data_lists = sp.findall('.//' + qn('p:custDataLst'))

        for cdl in cust_data_lists:
            for tag_ref in cdl:
                r_id = tag_ref.get(qn('r:id'))
                if r_id is None:
                    continue
                try:
                    slide_part = shape.part
                    rel = slide_part.rels[r_id]
                    tag_part = rel.target_part
                    if tag_part is None:
                        continue
                    tag_xml = ET.fromstring(tag_part.blob)
                    for tag_el in tag_xml:
                        name = tag_el.get('name', '')
                        val = tag_el.get('val', '')
                        if name:
                            tags[name.upper()] = val
                except Exception:
                    continue
    except Exception:
        pass
    return tags


def _parse_custom_xml_parts(prs: Presentation, pptx_path: Optional[Path] = None) -> dict[str, dict]:
    """Parse Custom XML Parts from the presentation.

    The Connector stores config DTOs in Custom XML Parts keyed by SHA256 hash.
    Structure: <Configs><Entry[hash]><Data>JSON</Data></Entry></Configs>

    Returns:
        { "ReportConfig": {hash: config_dict, ...},
          "DataFrameConfigKey": {hash: config_dict, ...} }
    """
    result = {XML_STORE_REPORT_CONFIG: {}, XML_STORE_PIVOT_CONFIG: {}}

    try:
        # Access Custom XML parts. python-pptx's iter_parts() may yield
        # properties XML rather than the data XML. Fall back to zipfile
        # approach if needed — read customXml/itemN.xml directly from the PPTX.
        import zipfile
        if pptx_path is None:
            pptx_path = getattr(prs.part.package, '_pkg_file', None)
            if pptx_path and hasattr(pptx_path, 'name'):
                pptx_path = pptx_path.name

        # Try zipfile approach first (more reliable for Custom XML)
        if pptx_path:
            try:
                with zipfile.ZipFile(pptx_path, 'r') as z:
                    for name in z.namelist():
                        if name.startswith('customXml/item') and name.endswith('.xml'):
                            try:
                                blob = z.read(name)
                                root = ET.fromstring(blob)
                                for store_key in [XML_STORE_REPORT_CONFIG, XML_STORE_PIVOT_CONFIG]:
                                    store_node = None
                                    for child in root.iter():
                                        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                        if tag == store_key:
                                            store_node = child
                                            break
                                    if store_node is None:
                                        continue
                                    for entry in store_node:
                                        entry_hash = entry.get('key') or entry.get('hash')
                                        if not entry_hash:
                                            continue
                                        data_node = None
                                        for child in entry:
                                            child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                            if child_tag == 'Data':
                                                data_node = child
                                                break
                                        if data_node is not None:
                                            json_text = data_node.text
                                            if not json_text or not json_text.strip():
                                                for val_child in data_node:
                                                    vt = val_child.tag.split('}')[-1] if '}' in val_child.tag else val_child.tag
                                                    if vt == 'Value' and val_child.text:
                                                        json_text = val_child.text
                                                        break
                                            if json_text and json_text.strip():
                                                try:
                                                    config = json.loads(json_text.strip())
                                                    result[store_key][entry_hash] = config
                                                except json.JSONDecodeError:
                                                    pass
                            except Exception:
                                continue
                return result
            except Exception:
                pass

        # Fallback: python-pptx iter_parts
        package = prs.part.package
        for part in package.iter_parts():
            content_type = getattr(part, 'content_type', '')
            if 'customXml' not in content_type and 'custom-xml' not in content_type:
                continue

            try:
                blob = part.blob
                if not blob:
                    continue

                # Try parsing as XML
                try:
                    root = ET.fromstring(blob)
                except ET.ParseError:
                    # Might be JSON directly
                    try:
                        data = json.loads(blob)
                        # Can't determine which store without XML structure
                        continue
                    except (json.JSONDecodeError, UnicodeDecodeError):
                        continue

                # Walk the XML looking for config entries
                # Expected structure matches CustomXmlPartService storage:
                # <Root>
                #   <ReportConfig>
                #     <Entry[hash]>
                #       <Data>{"ProjectId": 123, ...}</Data>
                #     </Entry>
                #   </ReportConfig>
                # </Root>
                for store_key in [XML_STORE_REPORT_CONFIG, XML_STORE_PIVOT_CONFIG]:
                    store_node = root.find(f".//{store_key}")
                    if store_node is None:
                        # Try without namespace
                        for child in root.iter():
                            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                            if tag == store_key:
                                store_node = child
                                break

                    if store_node is None:
                        continue

                    for entry in store_node:
                        # Entry tag may contain the hash in brackets: Entry[abc123]
                        entry_tag = entry.tag.split('}')[-1] if '}' in entry.tag else entry.tag
                        hash_match = re.search(r'\[([a-f0-9]+)\]', entry_tag)
                        if not hash_match:
                            # Try attribute-based lookup
                            entry_hash = entry.get('key') or entry.get('hash')
                            if not entry_hash:
                                continue
                        else:
                            entry_hash = hash_match.group(1)

                        # Find the Data element
                        data_node = entry.find('Data')
                        if data_node is None:
                            # Search children
                            for child in entry:
                                child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                if child_tag == 'Data':
                                    data_node = child
                                    break

                        if data_node is not None:
                            # JSON may be in data_node.text directly, OR inside
                            # a <Value> child (Galen Connector uses CDATA in <Value>)
                            json_text = data_node.text
                            if not json_text or not json_text.strip():
                                # Check for <Value> child (with or without namespace)
                                value_el = None
                                for child in data_node:
                                    child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                                    if child_tag == 'Value':
                                        value_el = child
                                        break
                                if value_el is not None and value_el.text:
                                    json_text = value_el.text

                            if json_text and json_text.strip():
                                try:
                                    config = json.loads(json_text.strip())
                                    result[store_key][entry_hash] = config
                                except json.JSONDecodeError:
                                    pass

            except Exception:
                continue

    except Exception:
        pass

    return result


def _classify_shape(shape) -> str:
    """Classify a shape into one of: chart, table, textbox, picture, group, other."""
    if shape.has_chart:
        return "chart"
    if shape.has_table:
        return "table"
    if shape.has_text_frame:
        return "textbox"
    shape_type_name = str(shape.shape_type)
    if 'PICTURE' in shape_type_name or 'PLACEHOLDER' in shape_type_name:
        return "picture"
    if 'GROUP' in shape_type_name:
        return "group"
    return "other"


def _emu_to_inches(emu: int) -> float:
    """Convert EMU to inches, rounded to 2 decimal places."""
    return round(emu / 914400, 2) if emu else 0.0


def _shape_to_position(shape) -> Position:
    """Extract Position from a python-pptx shape."""
    return Position(
        left=_emu_to_inches(shape.left),
        top=_emu_to_inches(shape.top),
        width=_emu_to_inches(shape.width),
        height=_emu_to_inches(shape.height),
    )


def _extract_headline_from_slide(slide) -> str:
    """Heuristic: find the talking headline text box on a slide.

    Real decks have two kinds of top-area text: (1) section tags ("Drivers
    and Barriers", "Non-Personal Promotion") — short, at the very top, and
    (2) talking headlines — longer, data-informed, typically 60-200 chars.

    We pick the LONGEST text in the top 1.5" zone. Talking headlines are
    almost always longer than section tags. Median PET headline: 73 chars.
    """
    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text or len(text) < 10:
            continue
        top_inches = _emu_to_inches(shape.top)
        if top_inches < 1.5:
            # Score: font size (larger = more likely headline), then text length
            font_size = 0
            try:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size:
                            font_size = r.font.size.pt
                        break
                    break
            except Exception:
                pass
            candidates.append((font_size, len(text), top_inches, text))

    if not candidates:
        return "Untitled Slide"

    # Pick by: largest font first, then longest text, then lower position
    candidates.sort(key=lambda c: (-c[0], -c[1], c[2]))
    return candidates[0][3]


def _resolve_config_from_backup(tags: dict, backup_tag_name: str) -> Optional[dict]:
    """Parse the per-shape *_BACKUP tag JSON. Used as a fallback when the
    deck-wide customXml store has no entry for the hash.

    Connector writes the full ReportConfig / PivotConfig as inline JSON in
    REPORTCONFIGHASH_BACKUP / DATAFRAMECONFIGHASH_BACKUP on every connected
    shape. The customXml store is canonical when present (one copy per hash,
    deduplicated across the deck), but cross-deck PowerPoint paste +
    incremental migration can leave hashes referenced on shapes without a
    matching customXml entry. Falling back to the per-shape backup keeps
    those shapes treated as connected — which is what the user sees in
    PowerPoint anyway.
    """
    raw = tags.get(backup_tag_name)
    if not raw:
        return None
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return None


def _report_config_to_lineage(report_config: dict, tags: dict) -> DataLineage:
    """Map a ReportConfigDto JSON to DataLineage fields.

    ReportConfigDto fields (from Connector PRD):
        ProjectId, SurveyId, ReportingPlanId, AnalysisIds[],
        SegmentIds[], StaticTimePeriodIds[], TimePeriodType,
        DynamicTimePeriod { LatestNDeliverables, IncludeLiveWave },
        RollUpDeliverables, NestSegments, IncludeOverallSegment
    """
    lineage = DataLineage()

    # Primary Synapse identifiers
    lineage.project_id = report_config.get('ProjectId')
    lineage.survey_id = report_config.get('SurveyId')
    lineage.reporting_plan_id = report_config.get('ReportingPlanId')

    analysis_ids = report_config.get('AnalysisIds', [])
    if isinstance(analysis_ids, list):
        lineage.analysis_ids = [int(x) for x in analysis_ids if x is not None]

    segment_ids = report_config.get('SegmentIds', [])
    if isinstance(segment_ids, list):
        lineage.segment_ids = [int(x) for x in segment_ids if x is not None]

    # Deliverables: static vs dynamic
    time_period_type = report_config.get('TimePeriodType', '')
    static_ids = report_config.get('StaticTimePeriodIds', [])
    if isinstance(static_ids, list) and static_ids:
        lineage.static_time_period_ids = [int(x) for x in static_ids if x is not None]
    # StaticTimePeriodNames — human-readable (always extract if available)
    static_names = report_config.get('StaticTimePeriodNames', [])
    if isinstance(static_names, list) and static_names:
        lineage.static_time_period_names = static_names

    dtp = report_config.get('DynamicTimePeriod', {})
    if isinstance(dtp, dict) and dtp:
        lineage.dynamic_latest_n = dtp.get('LatestNDeliverables')
        lineage.include_live_wave = dtp.get('IncludeLiveWave')

    # Analysis type from shape tag
    lineage.analysis_type = tags.get(TAG_ANALYSIS_TYPE)

    # ColumnKeyLabelMap from shape tag
    klm_raw = tags.get(TAG_COLUMN_KEY_LABEL_MAP)
    if klm_raw:
        try:
            lineage.column_key_label_map = json.loads(klm_raw)
        except json.JSONDecodeError:
            pass

    # Audit fields
    lineage.last_data_pull = tags.get(TAG_LAST_REFRESH)
    error_msg = tags.get(TAG_REFRESH_ERROR)
    lineage.last_refresh_error = error_msg if error_msg else None

    # Config hash (the ReportConfigHash tag value)
    lineage.config_hash = tags.get(TAG_REPORT_CONFIG_HASH)

    # SCHEMA GAP: ReportConfigDto has RollUpDeliverables, NestSegments,
    # IncludeOverallSegment fields that have no DataLineage counterpart.
    # These are Connector-specific rendering hints, not data lineage per se.
    # Documenting here per Track 2 rules: do NOT extend schema silently.

    return lineage


# ── OOXML data extraction for tagged shapes ────────────────────────────────────
# Tags carry Synapse config (what analysis to fetch); chart XML carries the
# actual rendered data. We need both for a complete spec.

from pptx.enum.chart import XL_CHART_TYPE

_CHART_TYPE_MAP = {
    XL_CHART_TYPE.BAR_CLUSTERED: "bar_clustered_horizontal",
    XL_CHART_TYPE.BAR_STACKED: "bar_stacked_100_horizontal",
    XL_CHART_TYPE.BAR_STACKED_100: "bar_stacked_100_horizontal",
    XL_CHART_TYPE.COLUMN_CLUSTERED: "column_clustered_vertical",
    XL_CHART_TYPE.COLUMN_STACKED: "column_stacked_100_vertical",
    XL_CHART_TYPE.COLUMN_STACKED_100: "column_stacked_100_vertical",
    XL_CHART_TYPE.LINE: "line_markers_trended",
    XL_CHART_TYPE.LINE_MARKERS: "line_markers_trended",
    XL_CHART_TYPE.LINE_STACKED: "line_markers_trended",
    XL_CHART_TYPE.XY_SCATTER: "xy_scatter_abacus",
    XL_CHART_TYPE.XY_SCATTER_LINES: "xy_scatter_abacus",
    XL_CHART_TYPE.DOUGHNUT: "doughnut_default",
    XL_CHART_TYPE.PIE: "doughnut_default",
}


def _classify_chart_from_ooxml(chart) -> str:
    """Classify a python-pptx chart object → chart_pattern key."""
    try:
        return _CHART_TYPE_MAP.get(chart.chart_type, "bar_clustered_horizontal")
    except Exception:
        return "bar_clustered_horizontal"


def _extract_chart_data_from_ooxml(chart) -> ChartData:
    """Extract actual chart data (categories + series) from OOXML."""
    categories = []
    series_list = []
    try:
        plot = chart.plots[0]
        try:
            cats = plot.categories
            if cats is not None:
                categories = [str(c) for c in cats]
        except Exception:
            pass

        for idx, s in enumerate(plot.series):
            name = str(s.name) if s.name else f"Series {idx}"
            values = []
            try:
                for v in s.values:
                    values.append(float(v) if v is not None else 0.0)
            except Exception:
                values = [0.0] * max(len(categories), 1)

            color = "#999999"
            try:
                fill = s.format.fill
                if fill.type is not None:
                    rgb = fill.fore_color.rgb
                    color = f"#{rgb}"
            except Exception:
                pass

            series_list.append(Series(name=name, values=values, color=color))
    except Exception:
        pass

    if not categories:
        categories = ["unknown"]
    if not series_list:
        series_list = [Series(name="unknown", values=[0.0], color="#999999")]

    return ChartData(categories=categories, series=series_list)


def _extract_chart_chrome_from_ooxml(chart) -> ChartChrome:
    """Extract chart formatting from OOXML: gapWidth, dLblPos, axis orientation, etc."""
    chrome = ChartChrome()
    try:
        cs = chart._chartSpace
        ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"

        # gapWidth
        gw = cs.find(f".//{{{ns}}}gapWidth")
        if gw is not None:
            try:
                chrome.gap_width = int(gw.get("val"))
            except (ValueError, TypeError):
                pass

        # overlap
        ov = cs.find(f".//{{{ns}}}overlap")
        if ov is not None:
            try:
                chrome.overlap = int(ov.get("val"))
            except (ValueError, TypeError):
                pass

        # dLblPos (data label position)
        dlp = cs.find(f".//{{{ns}}}dLblPos")
        if dlp is not None:
            pos_val = dlp.get("val", "")
            # Map OOXML values to our DataLabelsSpec positions
            pos_map = {"ctr": "ctr", "outEnd": "outEnd", "inEnd": "inEnd",
                       "t": "above", "b": "below", "l": "left", "r": "right",
                       "inBase": "inEnd", "bestFit": "outEnd"}
            chrome.data_labels = DataLabelsSpec(
                show=True,
                position=pos_map.get(pos_val, pos_val),
                format="0%",
            )

        # Axis orientation
        orient = cs.find(f".//{{{ns}}}catAx/{{{ns}}}scaling/{{{ns}}}orientation")
        if orient is not None and orient.get("val") == "maxMin":
            chrome.hide_category_labels = True  # typically companion table pattern

        # tickLblPos
        tlp = cs.find(f".//{{{ns}}}catAx/{{{ns}}}tickLblPos")
        if tlp is not None and tlp.get("val") == "none":
            chrome.hide_category_labels = True

        # Legend
        legend = cs.find(f".//{{{ns}}}legend")
        if legend is not None:
            chrome.legend = LegendSpec(show=True)
        else:
            chrome.legend = LegendSpec(show=False)

        # Gridlines
        mg = cs.find(f".//{{{ns}}}majorGridlines")
        chrome.gridlines = mg is not None

        # Title
        title = cs.find(f".//{{{ns}}}title")
        if title is not None:
            chrome.title = "(chart title)"

    except Exception:
        pass
    return chrome


def _extract_table_data_from_shape(shape) -> tuple[list[str], list[list[str]], dict]:
    """Extract labels, rows, and formatting from a table shape.

    Returns (labels, rows, formatting) where formatting contains:
    - font_name: most common font in the table
    - font_size_pt: most common font size
    - font_color: most common text color
    """
    labels = []
    rows = []
    formatting = {}
    try:
        table = shape.table
        n_rows = len(table.rows)
        n_cols = len(table.columns)
        if n_rows == 0 or n_cols == 0:
            return [], [], {}

        font_names = []
        font_sizes = []
        font_colors = []

        for row_idx in range(n_rows):
            row = []
            for col_idx in range(n_cols):
                cell = table.cell(row_idx, col_idx)
                row.append(cell.text.strip())
                # Extract font formatting
                try:
                    for p in cell.text_frame.paragraphs:
                        for run in p.runs:
                            if run.font.name:
                                font_names.append(run.font.name)
                            if run.font.size:
                                font_sizes.append(round(run.font.size.pt, 1))
                            if run.font.color and run.font.color.rgb:
                                font_colors.append(f"#{run.font.color.rgb}")
                            break
                        break
                except Exception:
                    pass
            rows.append(row)

        labels = [row[0] for row in rows if row]

        # Most common formatting values
        from collections import Counter
        if font_names:
            formatting["font_name"] = Counter(font_names).most_common(1)[0][0]
        if font_sizes:
            formatting["font_size_pt"] = Counter(font_sizes).most_common(1)[0][0]
        if font_colors:
            formatting["font_color"] = Counter(font_colors).most_common(1)[0][0]
    except Exception:
        pass
    return labels, rows, formatting


def read_tagged_shapes(
    pptx_path: str | Path,
) -> tuple[list[SlideSpec], TagReaderSummary, list[UntaggedShape]]:
    """Tier 1: Read Connector-tagged shapes from a PPTX.

    Returns:
        (specs, summary, untagged_shapes)
        - specs: list of SlideSpec for shapes with Connector tags
        - summary: TagReaderSummary with counts
        - untagged_shapes: list of UntaggedShape for Tier 2
    """
    pptx_path = Path(pptx_path)
    prs = Presentation(str(pptx_path))

    # Parse Custom XML Parts for config lookup
    xml_configs = _parse_custom_xml_parts(prs, pptx_path=pptx_path)
    report_configs = xml_configs.get(XML_STORE_REPORT_CONFIG, {})
    pivot_configs = xml_configs.get(XML_STORE_PIVOT_CONFIG, {})

    summary = TagReaderSummary(
        total_slides=len(prs.slides),
        custom_xml_parts_found=len(report_configs) + len(pivot_configs),
    )

    specs: list[SlideSpec] = []
    untagged: list[UntaggedShape] = []

    for slide_idx, slide in enumerate(prs.slides):
        slide_tagged = 0
        slide_untagged = 0
        slide_components = []       # ALL components on this slide
        slide_lineage = None        # best lineage found (first tagged shape with Synapse IDs)
        slide_has_synapse = False

        for shape in slide.shapes:
            summary.total_shapes += 1

            # Read all tags from this shape
            tags = _get_shape_tags_all(shape)

            # Check if this shape has any Connector tags
            has_report_hash = TAG_REPORT_CONFIG_HASH in tags
            has_pivot_hash = TAG_PIVOT_CONFIG_HASH in tags
            has_mapping = TAG_MAPPING_CONFIG in tags
            is_tagged = has_report_hash or has_pivot_hash or has_mapping

            if not is_tagged:
                slide_untagged += 1
                # Still collect as untagged for Tier 2 (backward compat)
                shape_type = _classify_shape(shape)
                us = UntaggedShape(
                    slide_index=slide_idx,
                    shape_name=shape.name,
                    shape_type=shape_type,
                    left_emu=shape.left or 0,
                    top_emu=shape.top or 0,
                    width_emu=shape.width or 0,
                    height_emu=shape.height or 0,
                    has_chart=shape.has_chart,
                    has_table=shape.has_table,
                )
                if shape.has_text_frame:
                    us.text = shape.text_frame.text[:500]
                untagged.append(us)

                # ALSO extract untagged shapes as components for full-fidelity rebuild
                position = _shape_to_position(shape)
                if shape.has_chart:
                    chart = shape.chart
                    slide_components.append(ChartComponent(
                        position=position,
                        chart_pattern=_classify_chart_from_ooxml(chart),
                        data=_extract_chart_data_from_ooxml(chart),
                        chrome=_extract_chart_chrome_from_ooxml(chart),
                    ))
                elif shape.has_table:
                    labels, rows, tbl_fmt = _extract_table_data_from_shape(shape)
                    if rows:
                        n_cols = len(rows[0]) if rows else 0
                        if n_cols <= 1:
                            slide_components.append(LabelTableComponent(
                                position=position,
                                labels=labels,
                                font_size_pt=tbl_fmt.get("font_size_pt", 9.0),
                                font_name=tbl_fmt.get("font_name"),
                                alternating_rows=False,
                            ))
                        else:
                            slide_components.append(ValueTableComponent(
                                position=position,
                                headers=[],
                                rows=rows,
                                alternating_rows=False,
                            ))
                elif shape.has_text_frame:
                    text = shape.text_frame.text.strip()
                    if text:
                        # Extract font info
                        font_size = 10.0
                        font_color = None
                        font_name = None
                        bold = False
                        try:
                            for p in shape.text_frame.paragraphs:
                                for r in p.runs:
                                    if r.font.size:
                                        font_size = round(r.font.size.pt, 1)
                                    if r.font.color and r.font.color.rgb:
                                        font_color = f"#{r.font.color.rgb}"
                                    if r.font.name:
                                        font_name = r.font.name
                                    if r.font.bold:
                                        bold = True
                                    break
                                break
                        except Exception:
                            pass
                        slide_components.append(TextboxComponent(
                            position=position,
                            text=text[:500],
                            font_size_pt=font_size,
                            font_color=font_color,
                            bold=bold,
                        ))
                continue

            # Tagged shape — extract config
            slide_tagged += 1
            summary.tagged_shapes += 1

            # Resolve ReportConfig from Custom XML Part — fall back to the
            # per-shape REPORTCONFIGHASH_BACKUP JSON if the customXml store
            # has no entry for this hash (Connector migration sometimes
            # populates the per-shape tag but not the deck-wide store).
            report_config = None
            if has_report_hash:
                r_hash = tags[TAG_REPORT_CONFIG_HASH]
                report_config = report_configs.get(r_hash)
                if report_config:
                    summary.report_configs_resolved += 1
                else:
                    report_config = _resolve_config_from_backup(
                        tags, TAG_REPORT_CONFIG_HASH_BACKUP)
                    if report_config:
                        summary.report_configs_from_backup += 1

            # Resolve PivotConfig — same fallback to DATAFRAMECONFIGHASH_BACKUP.
            pivot_config = None
            if has_pivot_hash:
                p_hash = tags[TAG_PIVOT_CONFIG_HASH]
                pivot_config = pivot_configs.get(p_hash)
                if pivot_config:
                    summary.pivot_configs_resolved += 1
                else:
                    pivot_config = _resolve_config_from_backup(
                        tags, TAG_PIVOT_CONFIG_HASH_BACKUP)
                    if pivot_config:
                        summary.pivot_configs_from_backup += 1

            # Parse MappingConfig
            mapping_config = None
            if has_mapping:
                try:
                    mapping_config = json.loads(tags[TAG_MAPPING_CONFIG])
                    summary.mapping_configs_resolved += 1
                except json.JSONDecodeError:
                    pass

            # Build DataLineage from ReportConfig + tags
            lineage = DataLineage()
            if report_config:
                lineage = _report_config_to_lineage(report_config, tags)
            else:
                lineage.analysis_type = tags.get(TAG_ANALYSIS_TYPE)
                lineage.last_data_pull = tags.get(TAG_LAST_REFRESH)
                error_msg = tags.get(TAG_REFRESH_ERROR)
                lineage.last_refresh_error = error_msg if error_msg else None
                lineage.config_hash = tags.get(TAG_REPORT_CONFIG_HASH)

            # Keep the best lineage (first one with Synapse IDs)
            has_ids = (lineage.project_id is not None or
                       lineage.reporting_plan_id is not None or
                       bool(lineage.analysis_ids))
            if has_ids and not slide_has_synapse:
                slide_lineage = lineage
                slide_has_synapse = True
            elif slide_lineage is None:
                slide_lineage = lineage

            # Extract component from this shape
            position = _shape_to_position(shape)
            shape_type = _classify_shape(shape)
            if shape_type == "chart" and shape.has_chart:
                chart = shape.chart
                chart_pattern = _classify_chart_from_ooxml(chart)
                chart_data = _extract_chart_data_from_ooxml(chart)
                chart_chrome = _extract_chart_chrome_from_ooxml(chart)
                slide_components.append(ChartComponent(
                    position=position,
                    chart_pattern=chart_pattern,
                    data=chart_data,
                    chrome=chart_chrome,
                ))
            elif shape_type == "table" and shape.has_table:
                labels, rows, tbl_fmt = _extract_table_data_from_shape(shape)
                if rows:
                    n_cols = len(rows[0]) if rows else 0
                    if n_cols <= 1:
                        # Single-column → label table with extracted font
                        slide_components.append(LabelTableComponent(
                            position=position,
                            labels=labels,
                            font_size_pt=tbl_fmt.get("font_size_pt", 9.0),
                            font_name=tbl_fmt.get("font_name"),
                            alternating_rows=False,  # don't apply SlideGen styling
                        ))
                    else:
                        # Multi-column → value table (preserves all columns)
                        # Don't split header/body — pass all rows as data
                        # (the original table may not have a header row)
                        slide_components.append(ValueTableComponent(
                            position=position,
                            headers=[],
                            rows=rows,
                            alternating_rows=False,
                        ))
            elif shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text and len(text) > 5:
                    slide_components.append(TextboxComponent(
                        position=position,
                        text=text[:200],
                    ))

        # After processing all shapes on this slide: build ONE SlideSpec
        if slide_components:
            headline_text = _extract_headline_from_slide(slide)

            # Classify layout from component composition
            n_charts = sum(1 for c in slide_components if c.type == "chart")
            n_tables = sum(1 for c in slide_components if c.type in ("label_table", "value_table"))
            layout_key = f"observed_{n_charts}_chart_{n_tables}_table"

            # Extract the original slide layout name (for template matching)
            try:
                slide_layout_name = slide.slide_layout.name
            except Exception:
                slide_layout_name = None

            completeness = "complete" if slide_has_synapse else "layout_complete_data_missing"

            spec = SlideSpec(
                slide_id=f"slide_{slide_idx:03d}",
                slide_index=slide_idx,
                layout=layout_key,
                headline=HeadlineSpec(text=headline_text),
                components=slide_components,
                data_lineage=slide_lineage or DataLineage(),
                metadata=SlideMetadata(
                    created_by="deck-reader-tier1",
                    created_at=None,
                    tier="1",
                    slide_layout_name=slide_layout_name,
                    confidence="high" if slide_has_synapse else "medium",
                ),
                spec_completeness=completeness,
                spec_version=SPEC_VERSION,
            )
            specs.append(spec)

        summary.untagged_shapes += slide_untagged
        summary.per_slide[slide_idx] = {
            "tagged": slide_tagged,
            "untagged": slide_untagged,
        }

    return specs, summary, untagged


def print_summary(summary: TagReaderSummary) -> None:
    """Print a human-readable summary of the tag reading results."""
    print(f"\n{'=' * 60}")
    print(f"Deck Reader — Tier 1 (Tag Reader) Summary")
    print(f"{'=' * 60}")
    print(f"  Total slides:              {summary.total_slides}")
    print(f"  Total shapes scanned:      {summary.total_shapes}")
    print(f"  Tagged shapes (Tier 1):    {summary.tagged_shapes}")
    print(f"  Untagged shapes (Tier 2):  {summary.untagged_shapes}")
    print(f"  Custom XML Parts found:    {summary.custom_xml_parts_found}")
    print(f"  ReportConfigs resolved:    {summary.report_configs_resolved}"
          + (f"  (+{summary.report_configs_from_backup} via per-shape backup)"
             if summary.report_configs_from_backup else ""))
    print(f"  PivotConfigs resolved:     {summary.pivot_configs_resolved}"
          + (f"  (+{summary.pivot_configs_from_backup} via per-shape backup)"
             if summary.pivot_configs_from_backup else ""))
    print(f"  MappingConfigs resolved:   {summary.mapping_configs_resolved}")
    print(f"\n  Per-slide breakdown:")
    for idx, counts in sorted(summary.per_slide.items()):
        t, u = counts['tagged'], counts['untagged']
        marker = " *" if t > 0 else ""
        print(f"    Slide {idx + 1:3d}: {t} tagged, {u} untagged{marker}")
    print(f"{'=' * 60}\n")


# ═════════════════════════════════════════════════════════════════════════════
# Config spec generation (v1.2) — spec as CONFIG, not DATA
# ═════════════════════════════════════════════════════════════════════════════
#
# Generates SlideSpec v1.2 with data_mapping on each component instead of data
# values. The spec becomes a timeless instruction file: same spec + new time
# periods = new deck.

# Connector field names → Synapse API column names.
# Kept local (mirrors synapse_chart_mapper.FIELD_MAP) to avoid circular deps.
_CONNECTOR_FIELD_MAP: dict[str, str] = {
    "value": "y_label",
    "title": "y_label",
    "question": "y_label",
    "y_label": "y_label",
    "y_code": "y_code",
    "segment_1": "segment_1",
    "segment_2": "segment_2",
    "time_period_name": "time_period_name",
    "time_period_id": "time_period_id",
    "measure": "measure",
    "n": "base",
    "n - value": "base",
    "decimal": "percentage",
    "frequency": "frequency",
    "reach": "reach",
    "SOV": "sov",
    "share": "share",
    "ME": "margin_of_error",
    "base": "base",
    "alias_label": "alias_label",
    "x_code": "x_code",
}


def _remap_field(name: str) -> str:
    """Translate a Connector field name to its Synapse API equivalent."""
    return _CONNECTOR_FIELD_MAP.get(name, name)


def _pivot_to_data_transform(
    pivot_config: dict,
    extra_filters: list[DataFilter] | None = None,
) -> DataTransform:
    """Convert Connector PivotConfig → source-agnostic DataTransform.

    PivotConfig fields (from DataFrameConfigHash Custom XML):
        RowFields, ColumnFields, ValueFields, Filters

    ColumnFields that are segment/time dimensions get promoted to component-
    level filters (passed via extra_filters) rather than becoming column_field,
    since the component already knows which segment panel it belongs to.
    """
    row_fields = pivot_config.get("RowFields", [])
    col_fields = pivot_config.get("ColumnFields", [])
    val_fields = pivot_config.get("ValueFields", [])
    raw_filters = pivot_config.get("Filters", [])

    # Map row field(s) — use first; if multiple, join with "+"
    if len(row_fields) == 1:
        row_field = _remap_field(row_fields[0])
    elif row_fields:
        row_field = "+".join(_remap_field(f) for f in row_fields)
    else:
        row_field = ""

    # Column field — skip segment/time dims (they become component filters).
    # Whatever remains is the measure/grouping dimension.
    filter_dims = {"segment_1", "segment_2", "time_period_name", "time_period_id"}
    non_filter_cols = [c for c in col_fields if _remap_field(c) not in filter_dims]
    if non_filter_cols:
        column_field = _remap_field(non_filter_cols[0])
    else:
        column_field = None

    # Value field — use first
    value_field = _remap_field(val_fields[0]) if val_fields else ""

    # Convert PivotConfig.Filters → DataFilter list
    filters: list[DataFilter] = []
    for f in raw_filters:
        col_key = _remap_field(f.get("ColumnKey", ""))
        fval = f.get("Value", "")
        criteria = f.get("filterCriteria", 1)
        op = "ne" if criteria == 2 else "eq"
        if col_key and fval:
            filters.append(DataFilter(field=col_key, value=fval, operator=op))

    if extra_filters:
        filters.extend(extra_filters)

    return DataTransform(
        row_field=row_field,
        column_field=column_field,
        value_field=value_field,
        filters=filters,
    )


def _parse_selected_columns(
    selected_columns: list[str],
    pivot_config: dict | None,
) -> dict:
    """Parse MappingConfig.selectedColumns to extract segment filter, series roles, time periods.

    selectedColumns examples:
        ['value']                           → category-only (label column)
        ['value', "Q1'26 - Overall Data"]   → category + one series
        ['value', "Q1'26 - CARD - decimal"] → category + segment-filtered series
        ['question', "Q1'26 - ... - CARD - L - Average of decimal", ...]  → stacked bar
        ['<blank:Base>']                    → computed column
        ['<blank:Easy-CARD>']               → computed metric column

    Returns dict with:
        segment_filter: str | None — detected segment value
        time_periods: list[str] — detected time period labels
        series_roles: list[str] — measure/bucket names (L, N, H, etc.)
        is_computed: bool — all columns are <blank:> (computed)
        category_field: str — the first non-blank column (row label source)
    """
    result = {
        "segment_filter": None,
        "time_periods": [],
        "series_roles": [],
        "is_computed": False,
        "category_field": "",
    }
    if not selected_columns:
        return result

    # Check if all columns are <blank:...> (computed)
    if all(c.startswith("<blank:") for c in selected_columns):
        result["is_computed"] = True
        return result

    # First non-blank column is usually the category/row field
    for c in selected_columns:
        if not c.startswith("<blank:"):
            result["category_field"] = c
            break

    # Parse compound column names from the series columns (skip first = category)
    segments_seen = set()
    times_seen = []
    roles_seen = []

    # Known segment values from ReportConfig — we'll detect them in column names
    col_fields = pivot_config.get("ColumnFields", []) if pivot_config else []
    has_segment_dim = any(_remap_field(c) in ("segment_1", "segment_2") for c in col_fields)

    for col in selected_columns[1:]:
        if col.startswith("<blank:"):
            continue
        parts = [p.strip() for p in col.split(" - ")]
        # First part is usually time period (e.g. "Q1'26")
        if parts:
            tp = parts[0]
            if tp not in times_seen:
                times_seen.append(tp)

        # Look for segment values in the compound name.
        # Common patterns: "... - CARD - ...", "... - PCP - ...", "... - Overall Data"
        if has_segment_dim and len(parts) >= 2:
            # Walk parts looking for segment values (not time, not aggregation)
            agg_markers = {"Average of decimal", "Average of frequency", "Average of reach",
                           "Average of SOV", "Average of share", "decimal", "frequency",
                           "reach", "SOV", "share", "n - value"}
            for p in parts[1:]:
                if p in agg_markers or p == tp:
                    continue
                # Single letter parts are usually measure/bucket codes (L, N, H, 0)
                if len(p) <= 2 and p not in ("", "n"):
                    if p not in roles_seen:
                        roles_seen.append(p)
                    continue
                # Multi-word parts that aren't aggregation → likely segment value
                # But "Overall" appears as both segment value and nested label
                if p not in segments_seen:
                    segments_seen.add(p)

    result["time_periods"] = times_seen
    result["series_roles"] = roles_seen

    # If we found exactly one segment value, use it as filter
    # If multiple, the slide likely has the "Overall" wrapper → don't filter
    if len(segments_seen) == 1:
        result["segment_filter"] = segments_seen.pop()
    elif len(segments_seen) > 1:
        # Multiple segments in one shape's mapping → might be cross-segment comparison
        # Keep the most specific (non-"Overall") one
        non_overall = {s for s in segments_seen if s.lower() not in ("overall", "overall data")}
        if len(non_overall) == 1:
            result["segment_filter"] = non_overall.pop()

    return result


def _extract_chart_series_colors(chart) -> list[str]:
    """Extract fill colors from each chart series in order."""
    colors = []
    try:
        for s in chart.plots[0].series:
            color = "#999999"
            try:
                fill = s.format.fill
                if fill.type is not None:
                    rgb = fill.fore_color.rgb
                    color = f"#{rgb}"
            except Exception:
                pass
            colors.append(color)
    except Exception:
        pass
    return colors


def _report_config_to_segments(report_config: dict) -> list[SegmentRule]:
    """Convert ReportConfig SegmentIds → structured SegmentRule list.

    The ReportConfig has SegmentIds (rule IDs sent to API) but not rule names
    or values — those require an API lookup. We create SegmentRule with rule_id
    populated and placeholder names; the pipeline or user fills in names later.
    """
    segment_ids = report_config.get("SegmentIds", [])
    if not segment_ids:
        return []

    rules = []
    for i, sid in enumerate(segment_ids):
        if sid is None:
            continue
        data_col = f"segment_{i + 1}"
        rules.append(SegmentRule(
            rule_id=int(sid),
            rule_name="",  # populated by API lookup or user
            values=[],     # populated by API lookup or data inspection
            data_column=data_col,
        ))
    return rules


def _build_config_lineage(report_config: dict, tags: dict) -> DataLineage:
    """Build DataLineage with structured segments for v1.2 config specs."""
    lineage = _report_config_to_lineage(report_config, tags)
    # Add structured segments (v1.2)
    lineage.segments = _report_config_to_segments(report_config)
    return lineage


def generate_config_specs(
    pptx_path: str | Path,
) -> tuple[list[SlideSpec], TagReaderSummary]:
    """Generate v1.2 config-only SlideSpecs from a Connector-tagged PPTX.

    Like read_tagged_shapes(), but produces specs with data_mapping instead of
    data values. Charts have ChartDataMapping, tables have TableDataMapping.
    No chart categories, no series values, no table cell text in the output.

    Returns:
        (specs, summary) — one SlideSpec per slide, plus tag reading summary
    """
    pptx_path = Path(pptx_path)
    prs = Presentation(str(pptx_path))

    xml_configs = _parse_custom_xml_parts(prs, pptx_path=pptx_path)
    report_configs = xml_configs.get(XML_STORE_REPORT_CONFIG, {})
    pivot_configs = xml_configs.get(XML_STORE_PIVOT_CONFIG, {})

    summary = TagReaderSummary(
        total_slides=len(prs.slides),
        custom_xml_parts_found=len(report_configs) + len(pivot_configs),
    )

    specs: list[SlideSpec] = []

    for slide_idx, slide in enumerate(prs.slides):
        slide_tagged = 0
        slide_untagged = 0
        components = []
        slide_lineage: DataLineage | None = None
        slide_has_synapse = False

        for shape in slide.shapes:
            summary.total_shapes += 1
            tags = _get_shape_tags_all(shape)

            has_report_hash = TAG_REPORT_CONFIG_HASH in tags
            has_pivot_hash = TAG_PIVOT_CONFIG_HASH in tags
            has_mapping = TAG_MAPPING_CONFIG in tags
            is_tagged = has_report_hash or has_pivot_hash or has_mapping

            position = _shape_to_position(shape)

            if not is_tagged:
                slide_untagged += 1
                # Untagged shapes: extract structural info (position, type, chrome)
                # but NO data values. Text boxes keep their text (structural).
                if shape.has_chart:
                    chart = shape.chart
                    components.append(ChartComponent(
                        position=position,
                        chart_pattern=_classify_chart_from_ooxml(chart),
                        data=None,  # config mode — no data
                        chrome=_extract_chart_chrome_from_ooxml(chart),
                    ))
                elif shape.has_table:
                    labels, rows, tbl_fmt = _extract_table_data_from_shape(shape)
                    n_rows = len(rows)
                    n_cols = len(rows[0]) if rows else 0
                    if n_cols <= 1:
                        components.append(LabelTableComponent(
                            position=position,
                            labels=[],  # config mode — no data
                            font_size_pt=tbl_fmt.get("font_size_pt", 9.0),
                            font_name=tbl_fmt.get("font_name"),
                            alternating_rows=False,
                        ))
                    else:
                        components.append(ValueTableComponent(
                            position=position,
                            headers=[],
                            rows=[],  # config mode — no data
                            alternating_rows=False,
                            row_count=n_rows,
                            col_count=n_cols,
                        ))
                elif shape.has_text_frame:
                    text = shape.text_frame.text.strip()
                    if text and len(text) > 5:
                        font_size, font_color, bold = 10.0, None, False
                        try:
                            for p in shape.text_frame.paragraphs:
                                for r in p.runs:
                                    if r.font.size:
                                        font_size = round(r.font.size.pt, 1)
                                    if r.font.color and r.font.color.rgb:
                                        font_color = f"#{r.font.color.rgb}"
                                    if r.font.bold:
                                        bold = True
                                    break
                                break
                        except Exception:
                            pass
                        components.append(TextboxComponent(
                            position=position,
                            text=text[:500],
                            font_size_pt=font_size,
                            font_color=font_color,
                            bold=bold,
                        ))
                continue

            # ── Tagged shape: extract config, build data_mapping ──
            slide_tagged += 1
            summary.tagged_shapes += 1

            # Resolve configs from Custom XML — fall back to per-shape
            # *_BACKUP tag JSON when the customXml store has no entry.
            # See `_resolve_config_from_backup` for the full reasoning.
            report_config = None
            if has_report_hash:
                r_hash = tags[TAG_REPORT_CONFIG_HASH]
                report_config = report_configs.get(r_hash)
                if report_config:
                    summary.report_configs_resolved += 1
                else:
                    report_config = _resolve_config_from_backup(
                        tags, TAG_REPORT_CONFIG_HASH_BACKUP)
                    if report_config:
                        summary.report_configs_from_backup += 1

            pivot_config = None
            if has_pivot_hash:
                p_hash = tags[TAG_PIVOT_CONFIG_HASH]
                pivot_config = pivot_configs.get(p_hash)
                if pivot_config:
                    summary.pivot_configs_resolved += 1
                else:
                    pivot_config = _resolve_config_from_backup(
                        tags, TAG_PIVOT_CONFIG_HASH_BACKUP)
                    if pivot_config:
                        summary.pivot_configs_from_backup += 1

            mapping_config = None
            if has_mapping:
                try:
                    mapping_config = json.loads(tags[TAG_MAPPING_CONFIG])
                    summary.mapping_configs_resolved += 1
                except json.JSONDecodeError:
                    pass

            # Build slide-level lineage from first tagged shape with Synapse IDs
            shape_lineage: Optional[DataLineage] = None
            if report_config:
                lineage = _build_config_lineage(report_config, tags)
                shape_lineage = lineage
                has_ids = lineage.project_id is not None or bool(lineage.analysis_ids)
                if has_ids and not slide_has_synapse:
                    slide_lineage = lineage
                    slide_has_synapse = True
                elif slide_lineage is None:
                    slide_lineage = lineage

            # Per-shape lineage dict — emitted into data_mapping so the refresh
            # pipeline can create component-level data_source keys when a
            # chart's analysis_ids diverge from the slide default.
            #
            # EXCEPTION: split visualizations (multiple charts sharing a
            # GROUPID with SPLITORDER 0..N) all share one pivot table by design
            # — don't create per-chart data_sources for those, or the refresh
            # tries to fetch the split analysis independently per shape and
            # source/refresh diverge. Empirically: CREON slide 63 has 13 charts
            # sharing GROUPID Group_25912087; their per-shape tags claim a
            # different analysis than the slide default, but source was
            # rendered from the slide default. Trust the slide default for
            # split groups.
            shape_lineage_dict = None
            is_split_visualization = (
                tags.get("GROUPID") is not None
                and tags.get(TAG_SPLIT_ORDER) is not None
            )
            if shape_lineage is not None and not is_split_visualization:
                from dataclasses import asdict as _asdict
                try:
                    shape_lineage_dict = _asdict(shape_lineage)
                except Exception:
                    shape_lineage_dict = None

            # Parse MappingConfig.selectedColumns for segment/series info
            selected_cols = mapping_config.get("selectedColumns", []) if mapping_config else []
            col_info = _parse_selected_columns(selected_cols, pivot_config)

            # Build component-level filters from detected segment
            comp_filters: list[DataFilter] = []
            if col_info["segment_filter"]:
                comp_filters.append(DataFilter(
                    field="segment_1",
                    value=col_info["segment_filter"],
                    operator="eq",
                ))

            # ── Chart shape → ChartComponent with data_mapping ──
            if shape.has_chart:
                chart = shape.chart
                chrome = _extract_chart_chrome_from_ooxml(chart)
                colors = _extract_chart_series_colors(chart)

                # Build series config from roles + colors
                roles = col_info["series_roles"]
                series_cfg: list[SeriesConfig] = []
                if roles:
                    for i, role in enumerate(roles):
                        color = colors[i] if i < len(colors) else "#999999"
                        series_cfg.append(SeriesConfig(role=role, color=color, order=i))
                elif colors:
                    # No parsed roles — use series index as role
                    for i, color in enumerate(colors):
                        series_cfg.append(SeriesConfig(role=f"series_{i}", color=color, order=i))

                # Build data transform from PivotConfig
                transform = _pivot_to_data_transform(pivot_config or {}, comp_filters)

                # Parse ColumnKeyLabelMap from shape tag
                klm = None
                klm_raw = tags.get(TAG_COLUMN_KEY_LABEL_MAP)
                if klm_raw:
                    try:
                        klm = json.loads(klm_raw)
                    except json.JSONDecodeError:
                        pass

                # Read split visualization tags
                split_order_str = tags.get(TAG_SPLIT_ORDER)
                rpo = mapping_config.get("rowsPerObject") if mapping_config else None
                top_n = mapping_config.get("topNRows") if mapping_config else None
                # Default: original template shape (no SPLITORDER tag) = order 0
                if split_order_str:
                    split_order_val = int(split_order_str)
                elif rpo and rpo > 0:
                    split_order_val = 0  # original shape = first slice
                else:
                    split_order_val = None

                components.append(ChartComponent(
                    position=position,
                    chart_pattern=_classify_chart_from_ooxml(chart),
                    data=None,
                    chrome=chrome,
                    data_mapping=ChartDataMapping(
                        transform=transform,
                        series_config=series_cfg,
                        raw_pivot_config=pivot_config,
                        raw_mapping_config=mapping_config,
                        raw_column_key_label_map=klm,
                        raw_data_lineage=shape_lineage_dict,
                        split_order=split_order_val,
                        rows_per_object=rpo,
                        top_n_rows=top_n,
                    ),
                ))

            # ── Table shape → ValueTable/LabelTable with data_mapping ──
            elif shape.has_table:
                labels, rows, tbl_fmt = _extract_table_data_from_shape(shape)
                n_rows = len(rows)
                n_cols = len(rows[0]) if rows else 0

                # Build table column config from selectedColumns
                tbl_columns: list[TableColumnConfig] = []
                for sc in selected_cols:
                    if sc.startswith("<blank:"):
                        # Computed column — header from the blank label
                        header = sc.replace("<blank:", "").rstrip(">")
                        tbl_columns.append(TableColumnConfig(
                            source_field="computed",
                            header_template=header,
                        ))
                    else:
                        # Data column — source from field name
                        tbl_columns.append(TableColumnConfig(
                            source_field=_remap_field(sc),
                            header_template=sc,
                        ))

                # Parse ColumnKeyLabelMap from shape tag
                klm = None
                klm_raw = tags.get(TAG_COLUMN_KEY_LABEL_MAP)
                if klm_raw:
                    try:
                        klm = json.loads(klm_raw)
                    except json.JSONDecodeError:
                        pass

                tbl_mapping = TableDataMapping(
                    filters=comp_filters,
                    columns=tbl_columns,
                    raw_pivot_config=pivot_config,
                    raw_mapping_config=mapping_config,
                    raw_column_key_label_map=klm,
                    raw_data_lineage=shape_lineage_dict,
                )

                if n_cols <= 1:
                    components.append(LabelTableComponent(
                        position=position,
                        labels=[],  # config mode — no data
                        font_size_pt=tbl_fmt.get("font_size_pt", 9.0),
                        font_name=tbl_fmt.get("font_name"),
                        alternating_rows=False,
                        data_mapping=tbl_mapping,
                    ))
                else:
                    components.append(ValueTableComponent(
                        position=position,
                        headers=[],
                        rows=[],  # config mode — no data
                        alternating_rows=False,
                        row_count=n_rows,
                        col_count=n_cols,
                        data_mapping=tbl_mapping,
                    ))

            # ── Text shape → TextboxComponent (structural text kept) ──
            elif shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text and len(text) > 5:
                    components.append(TextboxComponent(
                        position=position,
                        text=text[:500],
                    ))

        # ── Assemble ONE SlideSpec per slide ──
        if components:
            headline_text = _extract_headline_from_slide(slide)
            n_charts = sum(1 for c in components if c.type == "chart")
            n_tables = sum(1 for c in components
                          if c.type in ("label_table", "value_table"))
            layout_key = f"observed_{n_charts}_chart_{n_tables}_table"

            try:
                slide_layout_name = slide.slide_layout.name
            except Exception:
                slide_layout_name = None

            completeness = "complete" if slide_has_synapse else "layout_complete_data_missing"

            spec = SlideSpec(
                slide_id=f"slide_{slide_idx:03d}",
                slide_index=slide_idx,
                layout=layout_key,
                headline=HeadlineSpec(
                    text=f"{{{{headline_writer: {headline_text}}}}}",
                ),
                components=components,
                data_lineage=slide_lineage or DataLineage(),
                metadata=SlideMetadata(
                    created_by="config-spec-generator-v1.2",
                    tier="1",
                    slide_layout_name=slide_layout_name,
                    confidence="high" if slide_has_synapse else "medium",
                ),
                spec_completeness=completeness,
                spec_version=SPEC_VERSION,
            )
            specs.append(spec)

        summary.untagged_shapes += slide_untagged
        summary.per_slide[slide_idx] = {
            "tagged": slide_tagged,
            "untagged": slide_untagged,
        }

    return specs, summary


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    """CLI entry point: python -m slidegen.deck_reader path/to/deck.pptx [--out dir/]"""
    import argparse

    parser = argparse.ArgumentParser(
        description="Deck Reader — extract SlideSpecs from a PPTX file"
    )
    parser.add_argument("pptx", help="Path to the PPTX file")
    parser.add_argument("--out", default=None, help="Output directory for spec JSON files")
    parser.add_argument("--tier1-only", action="store_true", help="Only run Tier 1 (tags)")
    args = parser.parse_args()

    pptx_path = Path(args.pptx)
    if not pptx_path.exists():
        print(f"ERROR: File not found: {pptx_path}")
        sys.exit(1)

    if args.tier1_only:
        specs, summary, untagged = read_tagged_shapes(pptx_path)
        print_summary(summary)
    else:
        from slidegen.deck_reader import read_deck
        specs, summary_dict = read_deck(pptx_path)
        print_summary(summary_dict["tier1"])

    if args.out:
        out_dir = Path(args.out)
        out_dir.mkdir(parents=True, exist_ok=True)
        for spec in specs:
            out_file = out_dir / f"{spec.slide_id}.json"
            dump_spec(spec, out_file)
            print(f"  Wrote: {out_file}")

    print(f"\nTotal specs produced: {len(specs)}")


if __name__ == "__main__":
    main()
