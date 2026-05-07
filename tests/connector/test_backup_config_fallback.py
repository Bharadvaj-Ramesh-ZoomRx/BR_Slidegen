"""Tests for the per-shape *_BACKUP fallback in tag_reader.

When the deck-wide customXml/ ReportConfig store has no entry for a hash
(e.g. cross-deck PowerPoint paste followed by incremental migration),
tag_reader should fall back to parsing the per-shape REPORTCONFIGHASH_BACKUP
(or DATAFRAMECONFIGHASH_BACKUP) tag, which carries the same JSON inline.

These tests pin the helper's contract: valid JSON parses, malformed JSON
returns None, missing tag returns None — without round-tripping through a
full PPTX.
"""
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.deck_reader.tag_reader import (
    TAG_REPORT_CONFIG_HASH_BACKUP,
    TAG_PIVOT_CONFIG_HASH_BACKUP,
    _resolve_config_from_backup,
    _report_config_to_lineage,
)


VALID_REPORT_BACKUP = (
    '{"AnalysisIds":[691709],'
    '"DynamicTimePeriod":{"IncludeLiveWave":false,"LatestNDeliverables":2},'
    '"IncludeAlias":true,"IncludeOverallSegment":true,"NestSegments":false,'
    '"ProjectId":523,"ReportingPlanId":1143,"RollUpDeliverables":false,'
    '"SegmentIds":[],"SurveyId":482371,"TimePeriodType":1}'
)

VALID_PIVOT_BACKUP = (
    '{"AggregationType":1,"ColumnFields":["time_period_name","alias_label"],'
    '"FilterCondition":0,"Filters":[],"LatestColumnsFirst":false,'
    '"RowFields":["title"],"ValueFields":["decimal"],'
    '"columnDefinitions":[{"Format":"","Formula":"","IsDefaultAlias":false,'
    '"Name":"title"}]}'
)


def test_resolve_returns_none_when_tag_missing():
    """No backup tag at all → return None (caller falls back to None lineage)."""
    assert _resolve_config_from_backup({}, TAG_REPORT_CONFIG_HASH_BACKUP) is None


def test_resolve_returns_none_when_tag_empty():
    """Empty/whitespace value → return None."""
    assert _resolve_config_from_backup(
        {TAG_REPORT_CONFIG_HASH_BACKUP: ""}, TAG_REPORT_CONFIG_HASH_BACKUP
    ) is None


def test_resolve_returns_none_when_json_malformed():
    """Malformed JSON should NOT raise; just return None."""
    assert _resolve_config_from_backup(
        {TAG_REPORT_CONFIG_HASH_BACKUP: '{"AnalysisIds": [bro'},
        TAG_REPORT_CONFIG_HASH_BACKUP,
    ) is None


def test_resolve_parses_valid_report_backup():
    """Real-world REPORTCONFIGHASH_BACKUP from a connected shape parses to the
    same dict shape that the customXml store would yield."""
    cfg = _resolve_config_from_backup(
        {TAG_REPORT_CONFIG_HASH_BACKUP: VALID_REPORT_BACKUP},
        TAG_REPORT_CONFIG_HASH_BACKUP,
    )
    assert cfg is not None
    assert cfg["ProjectId"] == 523
    assert cfg["ReportingPlanId"] == 1143
    assert cfg["AnalysisIds"] == [691709]
    assert cfg["DynamicTimePeriod"]["LatestNDeliverables"] == 2
    assert cfg["SurveyId"] == 482371


def test_resolve_parses_valid_pivot_backup():
    cfg = _resolve_config_from_backup(
        {TAG_PIVOT_CONFIG_HASH_BACKUP: VALID_PIVOT_BACKUP},
        TAG_PIVOT_CONFIG_HASH_BACKUP,
    )
    assert cfg is not None
    assert cfg["RowFields"] == ["title"]
    assert cfg["ColumnFields"] == ["time_period_name", "alias_label"]
    assert cfg["ValueFields"] == ["decimal"]


def test_backup_dict_is_lineage_compatible():
    """The dict the fallback returns must be a drop-in for what
    _report_config_to_lineage expects — same field names, same types.
    Otherwise the fallback would silently produce empty lineage."""
    cfg = _resolve_config_from_backup(
        {TAG_REPORT_CONFIG_HASH_BACKUP: VALID_REPORT_BACKUP},
        TAG_REPORT_CONFIG_HASH_BACKUP,
    )
    lineage = _report_config_to_lineage(cfg, tags={})
    assert lineage.project_id == 523
    assert lineage.reporting_plan_id == 1143
    assert lineage.analysis_ids == [691709]
    assert lineage.survey_id == 482371
    assert lineage.dynamic_latest_n == 2
    assert lineage.include_live_wave is False
    assert lineage.segment_ids == []
