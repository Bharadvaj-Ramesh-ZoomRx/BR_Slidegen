"""Hash-parity tests for slidegen.connector_tags serializers.

The doc lays out three serialization profiles (§5):
  - ReportConfig + PivotConfig: PascalCase, alphabetical Ordinal sort,
    NullValueHandling.Ignore, Formatting.None.
  - MappingConfig: camelCase, C# declaration order, nulls retained,
    Formatting.None.

These tests pin that contract using actual Connector-produced JSON strings
(extracted from a real *_BACKUP tag — the same fixtures used by
test_backup_config_fallback) as ground truth.

If a future change to the serializer breaks byte-equality with these
fixtures, the corresponding shape's hash will diverge from what the
Connector itself would produce — which is the entire failure mode this
suite exists to prevent.
"""
import json
import sys
from collections import OrderedDict
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from slidegen.connector_tags import (
    _drop_alias_when_default,
    _enforce_mapping_field_order,
    build_connector_payload,
    column_def,
    numeric_filter,
    serialize_camel,
    serialize_existing_configs,
    serialize_pascal,
    sha256_hex,
    text_filter,
)


# ─── Ground-truth fixtures from real Connector-tagged shapes ──────────────

REAL_REPORT_BACKUP = (
    '{"AnalysisIds":[691709],'
    '"DynamicTimePeriod":{"IncludeLiveWave":false,"LatestNDeliverables":2},'
    '"IncludeAlias":true,"IncludeOverallSegment":true,"NestSegments":false,'
    '"ProjectId":523,"ReportingPlanId":1143,"RollUpDeliverables":false,'
    '"SegmentIds":[],"SurveyId":482371,"TimePeriodType":1}'
)

REAL_PIVOT_BACKUP = (
    '{"AggregationType":1,"ColumnFields":["time_period_name","alias_label"],'
    '"FilterCondition":0,"Filters":[],"LatestColumnsFirst":false,'
    '"RowFields":["title"],"ValueFields":["decimal"],'
    '"columnDefinitions":[{"Format":"","Formula":"","IsDefaultAlias":false,'
    '"Name":"title"}]}'
)


# ──────────────────────────────────────────────────────────────────────────
# Profile 1: PascalCase + alphabetical + null-stripped (ReportConfig)
# ──────────────────────────────────────────────────────────────────────────


def test_report_config_round_trips_byte_identical():
    """Real Connector-produced JSON → parse → re-serialize → same bytes.

    This is the hash-parity contract: if our serializer reproduces
    Connector's output byte-for-byte for a known shape, our minted hash
    will match the hash Connector would have computed.
    """
    parsed = json.loads(REAL_REPORT_BACKUP)
    reserialized = serialize_pascal(parsed)
    assert reserialized == REAL_REPORT_BACKUP, (
        f"Serialized output diverged from Connector's:\n"
        f"  expected: {REAL_REPORT_BACKUP}\n"
        f"  actual:   {reserialized}"
    )


def test_report_config_hash_round_trips():
    """Hash of re-serialized output must equal hash of original."""
    parsed = json.loads(REAL_REPORT_BACKUP)
    reserialized = serialize_pascal(parsed)
    assert sha256_hex(reserialized) == sha256_hex(REAL_REPORT_BACKUP)


def test_pascal_strips_nulls():
    """Null-valued keys must be omitted (NullValueHandling.Ignore)."""
    d = {"A": 1, "B": None, "C": 2}
    assert serialize_pascal(d) == '{"A":1,"C":2}'


def test_pascal_strips_nested_nulls():
    """Recursive null-stripping inside nested objects."""
    d = {"Outer": {"Keep": "x", "Drop": None}, "List": [{"K": 1, "V": None}]}
    assert serialize_pascal(d) == '{"List":[{"K":1}],"Outer":{"Keep":"x"}}'


def test_pascal_alphabetical_ordinal():
    """Properties must come out alphabetically (Ordinal A→Z)."""
    d = {"Zebra": 1, "Apple": 2, "Mango": 3}
    assert serialize_pascal(d) == '{"Apple":2,"Mango":3,"Zebra":1}'


def test_pascal_no_whitespace():
    """Formatting.None — no spaces between tokens."""
    out = serialize_pascal({"A": 1, "B": 2})
    assert " " not in out and "\n" not in out


# ──────────────────────────────────────────────────────────────────────────
# Profile 2: PascalCase + sorted Filters[] + alias-when-default drop (PivotConfig)
# ──────────────────────────────────────────────────────────────────────────


def test_pivot_config_round_trips_byte_identical():
    parsed = json.loads(REAL_PIVOT_BACKUP)
    reserialized = serialize_pascal(parsed)
    assert reserialized == REAL_PIVOT_BACKUP, (
        f"Pivot serialized output diverged from Connector's:\n"
        f"  expected: {REAL_PIVOT_BACKUP}\n"
        f"  actual:   {reserialized}"
    )


def test_pivot_filters_pre_hash_sort():
    """Filters[] must be sorted by (ColumnKey, Type, filterCriteria, Value)
    before hashing (PivotConfigService.PreProcessConfig). Different input
    orders ⇒ same output JSON ⇒ same hash."""
    fa = text_filter("region", "in", ["NA", "EU"])  # ColumnKey "region"
    fb = numeric_filter("count", "gte", 30)         # ColumnKey "count"
    payload_a = serialize_existing_configs(
        report_dict={"ProjectId": 1, "ReportingPlanId": 1, "AnalysisIds": [1],
                     "TimePeriodType": 1,
                     "DynamicTimePeriod": {"LatestNDeliverables": 1,
                                           "IncludeLiveWave": False}},
        pivot_dict={"AggregationType": 0, "ColumnFields": [], "Filters": [fa, fb],
                    "FilterCondition": 0, "RowFields": [], "ValueFields": [],
                    "columnDefinitions": [], "LatestColumnsFirst": False},
        mapping_dict={"selectedColumns": ["x"], "selectedRows": [],
                      "selectAllRows": True})
    payload_b = serialize_existing_configs(
        report_dict={"ProjectId": 1, "ReportingPlanId": 1, "AnalysisIds": [1],
                     "TimePeriodType": 1,
                     "DynamicTimePeriod": {"LatestNDeliverables": 1,
                                           "IncludeLiveWave": False}},
        pivot_dict={"AggregationType": 0, "ColumnFields": [], "Filters": [fb, fa],
                    "FilterCondition": 0, "RowFields": [], "ValueFields": [],
                    "columnDefinitions": [], "LatestColumnsFirst": False},
        mapping_dict={"selectedColumns": ["x"], "selectedRows": [],
                      "selectAllRows": True})
    assert payload_a["pivot_hash"] == payload_b["pivot_hash"]
    assert payload_a["pivot_json"] == payload_b["pivot_json"]


def test_drop_alias_when_default_true():
    """IsDefaultAlias=true ⇒ Alias key omitted from JSON
    (ShapeConfigurationServiceBase.cs:324-331 ShouldSerialize predicate)."""
    cd = {"Name": "x", "IsDefaultAlias": True, "Alias": "should be dropped"}
    out = _drop_alias_when_default(cd)
    assert "Alias" not in out
    assert out == {"Name": "x", "IsDefaultAlias": True}


def test_drop_alias_when_default_false_keeps_alias():
    """IsDefaultAlias=false ⇒ Alias preserved."""
    cd = {"Name": "x", "IsDefaultAlias": False, "Alias": "Custom Label"}
    out = _drop_alias_when_default(cd)
    assert out == cd


def test_column_def_helper_default_alias():
    """column_def(name) with no alias arg ⇒ IsDefaultAlias=True, no Alias key."""
    cd = column_def("region")
    assert cd == {"Name": "region", "IsDefaultAlias": True}


def test_column_def_helper_custom_alias():
    """column_def(name, alias='X') ⇒ IsDefaultAlias=False, Alias='X'."""
    cd = column_def("region", alias="Region Label", fmt="Text")
    assert cd["Alias"] == "Region Label"
    assert cd["IsDefaultAlias"] is False
    assert cd["Format"] == "Text"


# ──────────────────────────────────────────────────────────────────────────
# Profile 3: camelCase + declaration order + nulls retained (MappingConfig)
# ──────────────────────────────────────────────────────────────────────────


def test_mapping_preserves_declaration_order():
    """ObjectMappingConfigDto serializes in C# declaration order, NOT
    alphabetical. Mixing this up changes the on-disk JSON (and
    invalidates any future hash that includes it)."""
    d = OrderedDict()
    d["selectedColumns"] = ["region"]
    d["selectedRows"] = []
    d["selectAllRows"] = True
    d["addQuestionText"] = False
    out = serialize_camel(d)
    # Must come out in insertion order — selectedColumns first, then
    # selectedRows, etc.  NOT alphabetical (which would put addQuestionText
    # first).
    assert out.startswith('{"selectedColumns":["region"]')
    assert '"addQuestionText":false}' in out


def test_mapping_retains_nulls():
    """Nulls must NOT be stripped from MappingConfig (default Newtonsoft
    null handling for this profile)."""
    d = OrderedDict([("selectedColumns", ["x"]), ("rowsPerObject", None),
                     ("topNRows", None)])
    out = serialize_camel(d)
    assert '"rowsPerObject":null' in out
    assert '"topNRows":null' in out


def test_enforce_mapping_field_order_reorders_existing_dict():
    """Free-form mapping dict gets reordered into canonical declaration order."""
    bad_order = {
        "addLegend": True,
        "selectedColumns": ["x"],
        "selectedRows": [],
        "selectAllRows": True,
        "addQuestionText": False,
        "applyTranspose": False,
        "rowsPerObject": None,
        "addSplitObjectsToSingleSlide": False,
        "topNRows": None,
        "rowIdentifierColumnName": None,
        "customChartType": None,
        "tableHeaderMode": None,
        "insertEmptyColumnsAt": None,
        "moveRowsToFirst": [],
        "moveRowsToLast": [],
    }
    out = _enforce_mapping_field_order(bad_order)
    keys = list(out.keys())
    assert keys[0] == "selectedColumns"  # NOT addLegend (alphabetical first)
    assert keys[1] == "selectedRows"
    assert keys[-1] == "moveRowsToLast"


def test_mapping_canonical_form_byte_identical():
    """Build the §15.1-style mapping config; expect canonical bytes."""
    d = OrderedDict([
        ("selectedColumns", ["region", "Count of count"]),
        ("selectedRows", []),
        ("selectAllRows", True),
        ("addQuestionText", False),
        ("applyTranspose", False),
        ("addLegend", True),
        ("rowsPerObject", None),
        ("addSplitObjectsToSingleSlide", False),
        ("topNRows", None),
        ("rowIdentifierColumnName", None),
        ("customChartType", None),
        ("tableHeaderMode", None),
        ("insertEmptyColumnsAt", None),
        ("moveRowsToFirst", []),
        ("moveRowsToLast", []),
    ])
    expected = (
        '{"selectedColumns":["region","Count of count"],'
        '"selectedRows":[],"selectAllRows":true,"addQuestionText":false,'
        '"applyTranspose":false,"addLegend":true,"rowsPerObject":null,'
        '"addSplitObjectsToSingleSlide":false,"topNRows":null,'
        '"rowIdentifierColumnName":null,"customChartType":null,'
        '"tableHeaderMode":null,"insertEmptyColumnsAt":null,'
        '"moveRowsToFirst":[],"moveRowsToLast":[]}'
    )
    assert serialize_camel(d) == expected


# ──────────────────────────────────────────────────────────────────────────
# Pre-hash sorts (ReportConfig)
# ──────────────────────────────────────────────────────────────────────────


def test_analysis_ids_sorted_ascending():
    """Different AnalysisIds input orders ⇒ same hash."""
    p1 = build_connector_payload(
        project_id=1, reporting_plan_id=1, analysis_ids=[567, 234],
        latest_n_deliverables=1, selected_columns=["x"])
    p2 = build_connector_payload(
        project_id=1, reporting_plan_id=1, analysis_ids=[234, 567],
        latest_n_deliverables=1, selected_columns=["x"])
    assert p1["report_hash"] == p2["report_hash"]
    assert '"AnalysisIds":[234,567]' in p1["report_json"]


def test_segment_ids_sorted_when_not_nested():
    """SegmentIds sorts ascending when NestSegments=False."""
    p = build_connector_payload(
        project_id=1, reporting_plan_id=1, analysis_ids=[1],
        latest_n_deliverables=1, segment_ids=[205, 101], nest_segments=False,
        selected_columns=["x"])
    assert '"SegmentIds":[101,205]' in p["report_json"]


def test_segment_ids_preserves_order_when_nested():
    """SegmentIds preserves user order when NestSegments=True (segment_1
    semantics depend on order)."""
    p = build_connector_payload(
        project_id=1, reporting_plan_id=1, analysis_ids=[1],
        latest_n_deliverables=1, segment_ids=[205, 101], nest_segments=True,
        selected_columns=["x"])
    assert '"SegmentIds":[205,101]' in p["report_json"]


# ──────────────────────────────────────────────────────────────────────────
# End-to-end: build_connector_payload returns hashes that match its JSONs
# ──────────────────────────────────────────────────────────────────────────


def test_payload_hashes_match_jsons():
    """The returned hash must be sha256(returned json), not computed from
    some intermediate state."""
    p = build_connector_payload(
        project_id=588, reporting_plan_id=7, analysis_ids=[500],
        survey_id=42, latest_n_deliverables=4,
        row_fields=["region"], column_fields=["time_period_name"],
        value_fields=["count"], aggregation_type="Count",
        column_definitions=[column_def("region", fmt="Text")],
        selected_columns=["region", "Count of count"], add_legend=True)
    assert p["report_hash"] == sha256_hex(p["report_json"])
    assert p["pivot_hash"] == sha256_hex(p["pivot_json"])
    # Hash format: 64-char lowercase hex
    assert len(p["report_hash"]) == 64
    assert all(c in "0123456789abcdef" for c in p["report_hash"])


def test_payload_dynamic_mode_minimal():
    """Smallest valid Dynamic-mode call — no segments, no filters."""
    p = build_connector_payload(
        project_id=1, reporting_plan_id=1, analysis_ids=[1],
        latest_n_deliverables=4, selected_columns=["x"])
    assert '"TimePeriodType":1' in p["report_json"]
    assert '"DynamicTimePeriod":{"IncludeLiveWave":false,"LatestNDeliverables":4}' \
        in p["report_json"]
    # Static fields must be ABSENT (NullValueHandling.Ignore)
    assert "StaticTimePeriodIds" not in p["report_json"]
    assert "StaticTimePeriodNames" not in p["report_json"]


def test_payload_static_mode_minimal():
    """Static mode: TimePeriodType=0, StaticTimePeriodIds present, no
    DynamicTimePeriod block."""
    p = build_connector_payload(
        project_id=1, reporting_plan_id=1, analysis_ids=[1],
        static_time_period_ids=[101, 102], selected_columns=["x"])
    assert '"TimePeriodType":0' in p["report_json"]
    assert '"StaticTimePeriodIds":[101,102]' in p["report_json"]
    assert "DynamicTimePeriod" not in p["report_json"]


def test_payload_rejects_no_time_period():
    """Either Static or Dynamic must be specified — neither ⇒ ValueError."""
    import pytest
    with pytest.raises(ValueError, match="Static.*Dynamic"):
        build_connector_payload(
            project_id=1, reporting_plan_id=1, analysis_ids=[1],
            selected_columns=["x"])
