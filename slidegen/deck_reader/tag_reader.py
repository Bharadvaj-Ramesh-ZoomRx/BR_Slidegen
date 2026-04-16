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
    SlideSpec, HeadlineSpec, FooterSpec, Position, DataLineage, SlideMetadata,
    ChartComponent, LabelTableComponent, TextboxComponent, ChartData, Series,
    ChartChrome, DataLabelsSpec, LegendSpec, AxisSpec,
    dump_spec, SPEC_VERSION,
)


# ── Constants matching galen-powerpoint/Constants.cs ShapeTags ────────────────

TAG_REPORT_CONFIG_HASH = "ReportConfigHash"
TAG_PIVOT_CONFIG_HASH = "DataFrameConfigHash"
TAG_MAPPING_CONFIG = "MappingConfig"
TAG_REFRESH_ERROR = "RefreshErrorMsg"
TAG_LAST_REFRESH = "LastRefreshTime"
TAG_COLUMN_KEY_LABEL_MAP = "ColumnKeyLabelMap"
TAG_ANALYSIS_TYPE = "AnalysisType"
TAG_SPLIT_GROUP_ID = "SplitGroupId"
TAG_SPLIT_ORDER = "SplitOrder"

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

        # Method 1: Check for custDataLst links to tag parts
        # The shape may have a <p:custDataLst> child with <p:tags r:id="rIdN"/>
        cust_data = sp.findall(qn('p:custDataLst'))
        if not cust_data:
            # Shapes sometimes have tags in an extension list
            ext_lst = sp.findall(qn('p:extLst'))
            for ext in ext_lst:
                cust_data.extend(ext.findall('.//' + qn('p:custDataLst')))

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
                    slide_part = shape.part
                    tag_part = slide_part.related_parts.get(r_id)
                    if tag_part is None:
                        continue

                    # Parse the tag XML
                    tag_xml = ET.fromstring(tag_part.blob)
                    for tag_el in tag_xml:
                        name = tag_el.get('name', '')
                        val = tag_el.get('val', '')
                        if name == tag_name:
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

        cust_data_lists = sp.findall(qn('p:custDataLst'))
        ext_lst = sp.findall(qn('p:extLst'))
        for ext in ext_lst:
            cust_data_lists.extend(ext.findall('.//' + qn('p:custDataLst')))

        for cdl in cust_data_lists:
            for tag_ref in cdl:
                r_id = tag_ref.get(qn('r:id'))
                if r_id is None:
                    continue
                try:
                    slide_part = shape.part
                    tag_part = slide_part.related_parts.get(r_id)
                    if tag_part is None:
                        continue
                    tag_xml = ET.fromstring(tag_part.blob)
                    for tag_el in tag_xml:
                        name = tag_el.get('name', '')
                        val = tag_el.get('val', '')
                        if name:
                            tags[name] = val
                except Exception:
                    continue
    except Exception:
        pass
    return tags


def _parse_custom_xml_parts(prs: Presentation) -> dict[str, dict]:
    """Parse Custom XML Parts from the presentation.

    The Connector stores config DTOs in Custom XML Parts keyed by SHA256 hash.
    Structure: <Configs><Entry[hash]><Data>JSON</Data></Entry></Configs>

    Returns:
        { "ReportConfig": {hash: config_dict, ...},
          "DataFrameConfigKey": {hash: config_dict, ...} }
    """
    result = {XML_STORE_REPORT_CONFIG: {}, XML_STORE_PIVOT_CONFIG: {}}

    try:
        # Access the package-level Custom XML parts
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

                        if data_node is not None and data_node.text:
                            try:
                                config = json.loads(data_node.text.strip())
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
    """Heuristic: find the headline text box on a slide.

    Looks for the topmost text box with large font or positioned near the top.
    """
    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text or len(text) < 5:
            continue
        top_inches = _emu_to_inches(shape.top)
        # Headline is typically in the top 1.5 inches
        if top_inches < 1.5:
            # Score by position (higher = closer to top) and text length
            candidates.append((top_inches, len(text), text, shape))

    if not candidates:
        return "Untitled Slide"

    # Sort: prefer top-most, then longest text
    candidates.sort(key=lambda c: (c[0], -c[1]))
    return candidates[0][2]


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
    if str(time_period_type).lower() == 'static' or str(time_period_type) == '0':
        static_ids = report_config.get('StaticTimePeriodIds', [])
        if isinstance(static_ids, list):
            lineage.static_time_period_ids = [int(x) for x in static_ids if x is not None]
    else:
        dtp = report_config.get('DynamicTimePeriod', {})
        if isinstance(dtp, dict):
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
    xml_configs = _parse_custom_xml_parts(prs)
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
                # Collect info for Tier 2
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
                continue

            # Tagged shape — extract config
            slide_tagged += 1
            summary.tagged_shapes += 1

            # Resolve ReportConfig from Custom XML Part
            report_config = None
            if has_report_hash:
                r_hash = tags[TAG_REPORT_CONFIG_HASH]
                report_config = report_configs.get(r_hash)
                if report_config:
                    summary.report_configs_resolved += 1

            # Resolve PivotConfig from Custom XML Part
            pivot_config = None
            if has_pivot_hash:
                p_hash = tags[TAG_PIVOT_CONFIG_HASH]
                pivot_config = pivot_configs.get(p_hash)
                if pivot_config:
                    summary.pivot_configs_resolved += 1

            # Parse MappingConfig (stored directly as JSON on shape tag)
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
                # Minimal lineage from tags alone
                lineage.analysis_type = tags.get(TAG_ANALYSIS_TYPE)
                lineage.last_data_pull = tags.get(TAG_LAST_REFRESH)
                error_msg = tags.get(TAG_REFRESH_ERROR)
                lineage.last_refresh_error = error_msg if error_msg else None
                lineage.config_hash = tags.get(TAG_REPORT_CONFIG_HASH)

            # Extract headline
            headline_text = _extract_headline_from_slide(slide)

            # Build a minimal SlideSpec
            # NOTE: We can't fully reconstruct chart data from tags alone.
            # Tags carry config (what to fetch), not data (what was fetched).
            # The spec will have the DataLineage but components are placeholders.
            position = _shape_to_position(shape)

            components = []
            shape_type = _classify_shape(shape)
            if shape_type == "chart":
                components.append(ChartComponent(
                    position=position,
                    chart_pattern="bar_clustered_horizontal",  # placeholder
                    data=ChartData(categories=["placeholder"], series=[
                        Series(name="placeholder", values=[0.0], color="#999999")
                    ]),
                    chrome=ChartChrome(),
                ))
            elif shape_type == "table":
                components.append(LabelTableComponent(
                    position=position,
                    labels=["placeholder"],
                ))
            else:
                components.append(TextboxComponent(
                    position=position,
                    text=shape.text_frame.text[:200] if shape.has_text_frame else "",
                ))

            spec = SlideSpec(
                slide_id=f"tagged_{slide_idx:03d}_{shape.name.replace(' ', '_')[:20]}",
                slide_index=slide_idx,
                layout="observed_1chart_1table",
                headline=HeadlineSpec(text=headline_text),
                components=components,
                data_lineage=lineage,
                metadata=SlideMetadata(
                    created_by="deck-reader-tier1",
                    created_at=None,
                ),
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
    print(f"  ReportConfigs resolved:    {summary.report_configs_resolved}")
    print(f"  PivotConfigs resolved:     {summary.pivot_configs_resolved}")
    print(f"  MappingConfigs resolved:   {summary.mapping_configs_resolved}")
    print(f"\n  Per-slide breakdown:")
    for idx, counts in sorted(summary.per_slide.items()):
        t, u = counts['tagged'], counts['untagged']
        marker = " *" if t > 0 else ""
        print(f"    Slide {idx + 1:3d}: {t} tagged, {u} untagged{marker}")
    print(f"{'=' * 60}\n")


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
