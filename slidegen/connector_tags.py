"""
connector_tags.py — Build Synapse Connector connection state for a shape.

Faithful port of the rules in
    "Synapse Connector - Direct Tag Creation Reference.md"  (sections 5, 6, 20).

This module is pure data: it builds the three JSON payloads and the two
SHA256 hashes that together define a Connector "connected" shape.  It does
NOT touch a deck — the caller is free to write the resulting tags via COM
(pywin32) or via OOXML zip manipulation.

Why three serializers?
    Connector uses one Newtonsoft serialization profile for ReportConfig +
    PivotConfig (PascalCase, alphabetical Ordinal sort, NullValueHandling.Ignore)
    and a different one for MappingConfig (camelCase, C# declaration order,
    nulls retained).  Mixing them produces JSON that round-trips fine but
    hashes to a different value than what the connector itself would produce.

Hash determinism rules (§5):
    - Sort AnalysisIds ascending (Length>1).
    - Sort SegmentIds ascending if !NestSegments (Length>1).
    - Sort StaticTimePeriodIds ascending if TimePeriodType==Static (Length>1).
    - Sort Filters[] by (ColumnKey, Type, filterCriteria, Value) Ordinal.
    - Drop ColumnDefinition.Alias when IsDefaultAlias=true.

Public API:
    numeric_filter, text_filter, column_def     — array-element builders
    build_connector_payload(...)                — returns {report_json, pivot_json,
                                                  mapping_json, report_hash, pivot_hash}
    serialize_pascal, serialize_camel, sha256_hex — primitives (for tests)
"""
from __future__ import annotations

import hashlib
import json
from collections import OrderedDict


# ── Enum ordinals (match C# DTOs — see doc §5.4) ──────────────────────────
TIME_PERIOD_TYPE = {"Static": 0, "Dynamic": 1}
FILTER_CONDITION = {"And": 0, "Or": 1}
FILTER_TYPE = {"Numeric": 0, "Text": 1}
AGGREGATION_TYPE = {"Sum": 0, "Average": 1, "Count": 2}
SORT_ORDER = {"asc": 0, "Ascending": 0, "desc": 1, "Descending": 1}
CUSTOM_CHART = {"Abacus": 0, "Scatter": 1}
TABLE_HEADER = {"WithHeader": 0, "WithoutHeader": 1, "OnlyHeader": 2}
NUMERIC_OPS = {"eq": 0, "ne": 1, "gt": 2, "lt": 3, "gte": 4, "lte": 5}
TEXT_OPS = {"contains": 0, "in": 1, "not_in": 2}

ADDIN_NAMESPACE = "http://connector.synapse.com/"


# ─────────────────────────────────────────────────────────────────────────
# Builders for array-of-object fields (use these — never hand-craft dicts)
# ─────────────────────────────────────────────────────────────────────────

def numeric_filter(column_key: str, op: str, value) -> dict:
    """Build a numeric FilterDto entry for the `filters=` argument.

    op : 'eq' | 'ne' | 'gt' | 'lt' | 'gte' | 'lte'
    """
    if op not in NUMERIC_OPS:
        raise ValueError(f"numeric_filter op must be one of {list(NUMERIC_OPS)}")
    return {
        "ColumnKey": column_key,
        "Type": FILTER_TYPE["Numeric"],
        "Value": str(value),
        "filterCriteria": NUMERIC_OPS[op],
    }


def text_filter(column_key: str, op: str, values) -> dict:
    """Build a text FilterDto entry for the `filters=` argument.

    op     : 'contains' | 'in' | 'not_in'
    values : str or list[str]; multiple values are joined with ';' (the
             connector's FilterValueSeparator — see doc §12.1.2).
    """
    if op not in TEXT_OPS:
        raise ValueError(f"text_filter op must be one of {list(TEXT_OPS)}")
    if isinstance(values, (list, tuple)):
        joined = ";".join(str(v) for v in values)
    else:
        joined = str(values)
    return {
        "ColumnKey": column_key,
        "Type": FILTER_TYPE["Text"],
        "Value": joined,
        "filterCriteria": TEXT_OPS[op],
    }


def column_def(name: str, alias: str | None = None, fmt: str | None = None,
               formula: str | None = None, sort: dict | None = None) -> dict:
    """Build a ColumnDefinition entry for `column_definitions=`.

    name    : output column name (row-field key, pivot value-column name, or
              synthetic 'new_column_<i>' for formula columns)
    alias   : str or None.  None ⇒ IsDefaultAlias=True (auto-generated).
    fmt     : ColumnFormat enum string ('Text', 'NumericWithDecimal',
              'PercentageWithDec', 'PercentageWithoutDec',
              'PercentageWithoutMplWithoutDec', 'PercentageWithoutMplyWithDec',
              'CapitalText', 'NumericWithoutDecimal', 'MonthAndYear')
    formula : Excel-style formula string ('=A2/A3') for new derived columns
    sort    : dict {'level': int, 'order': 'asc'|'desc', 'custom_list': []}
    """
    d: dict = {"Name": name}
    if alias is None:
        d["IsDefaultAlias"] = True
    else:
        d["IsDefaultAlias"] = False
        d["Alias"] = alias
    if fmt:
        d["Format"] = fmt
    if formula:
        d["Formula"] = formula
    if sort:
        d["sortCriteria"] = {
            "Level": int(sort.get("level", 1)),
            "Order": SORT_ORDER[sort.get("order", "asc")],
            "CustomList": list(sort.get("custom_list", [])),
        }
    return d


# ─────────────────────────────────────────────────────────────────────────
# Serialization primitives — match Newtonsoft conventions in the connector
# ─────────────────────────────────────────────────────────────────────────

def _strip_nulls(obj):
    """Recursively drop keys whose value is None — matches NullValueHandling.Ignore."""
    if isinstance(obj, dict):
        return {k: _strip_nulls(v) for k, v in obj.items() if v is not None}
    if isinstance(obj, list):
        return [_strip_nulls(x) for x in obj]
    return obj


def serialize_pascal(d: dict) -> str:
    """ReportConfig / PivotConfig profile.

    Mirrors `OrderedContractResolver` + `NullValueHandling.Ignore` +
    `Formatting.None`: PascalCase keys (caller's responsibility),
    alphabetically sorted Ordinal, nulls omitted, no whitespace.
    """
    return json.dumps(_strip_nulls(d), separators=(",", ":"), sort_keys=True)


def serialize_camel(d: dict | OrderedDict) -> str:
    """ObjectMappingConfig profile.

    Mirrors `JsonSerializationHelper.Serialize` with
    `CamelCaseNamingStrategy`: camelCase keys (caller's responsibility),
    C# declaration order (preserved by Python's dict insertion order, so
    pass an OrderedDict or a Python 3.7+ dict), nulls retained, no
    whitespace.
    """
    return json.dumps(d, separators=(",", ":"))


def sha256_hex(s: str) -> str:
    """Lowercase hex SHA256 of UTF-8 bytes — matches doc §6 exactly."""
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def _drop_alias_when_default(col_def: dict) -> dict:
    """If IsDefaultAlias=True, drop Alias from the JSON.

    Matches the ShouldSerialize predicate at
    ShapeConfigurationServiceBase.cs:324-331.  Without this the hash differs
    from what the UI would produce for the same logical config.
    """
    out = dict(col_def)
    if out.get("IsDefaultAlias", False) and "Alias" in out:
        del out["Alias"]
    return out


# ─────────────────────────────────────────────────────────────────────────
# DTO builders (apply PreProcessConfig sort rules from §5)
# ─────────────────────────────────────────────────────────────────────────

def _build_report_config(*, project_id, survey_id, reporting_plan_id, analysis_ids,
                         segment_ids, nest_segments, static_time_period_ids,
                         static_time_period_names, latest_n_deliverables,
                         include_live_wave, rollup_deliverables, include_alias,
                         include_overall_segment) -> dict:
    """Build a ReportConfigDto-shaped dict ready for serialize_pascal()."""
    if static_time_period_ids:
        time_period_type = TIME_PERIOD_TYPE["Static"]
        dynamic_block = None
    elif latest_n_deliverables and latest_n_deliverables > 0:
        time_period_type = TIME_PERIOD_TYPE["Dynamic"]
        dynamic_block = {
            "IncludeLiveWave": bool(include_live_wave),
            "LatestNDeliverables": int(latest_n_deliverables),
        }
    else:
        raise ValueError(
            "Either static_time_period_ids must be non-empty (Static mode) "
            "or latest_n_deliverables must be > 0 (Dynamic mode).")

    # ReportConfigService.PreProcessConfig sort rules (doc §5.1)
    analyses = sorted(int(a) for a in (analysis_ids or []))
    segs = list(int(s) for s in (segment_ids or []))
    if not nest_segments and len(segs) > 1:
        segs = sorted(segs)
    static_ids = list(int(t) for t in (static_time_period_ids or []))
    if time_period_type == TIME_PERIOD_TYPE["Static"] and len(static_ids) > 1:
        static_ids = sorted(static_ids)

    return {
        "AnalysisIds": analyses,
        "DynamicTimePeriod": dynamic_block,
        "IncludeAlias": bool(include_alias),
        "IncludeOverallSegment": bool(include_overall_segment),
        "NestSegments": bool(nest_segments),
        "ProjectId": int(project_id),
        "ReportingPlanId": int(reporting_plan_id),
        "RollUpDeliverables": bool(rollup_deliverables),
        "SegmentIds": segs,
        "StaticTimePeriodIds": static_ids if static_ids else None,
        "StaticTimePeriodNames":
            list(static_time_period_names) if static_time_period_names else None,
        "SurveyId": int(survey_id) if survey_id is not None else None,
        "TimePeriodType": time_period_type,
    }


def _build_pivot_config(*, filters, filter_condition, row_fields, column_fields,
                        value_fields, aggregation_type, column_definitions,
                        latest_columns_first) -> dict:
    """Build a PivotConfigDto-shaped dict ready for serialize_pascal()."""
    # PivotConfigService.PreProcessConfig: sort filters by (ColumnKey, Type,
    # filterCriteria, Value) Ordinal — order doesn't affect semantics, but
    # determinism matters for the hash.
    filters_sorted = sorted(
        list(filters or []),
        key=lambda f: (f.get("ColumnKey", ""), f.get("Type", 0),
                       f.get("filterCriteria", 0), f.get("Value", "")))

    col_defs = [_drop_alias_when_default(c) for c in (column_definitions or [])]

    agg = (AGGREGATION_TYPE[aggregation_type]
           if isinstance(aggregation_type, str) else int(aggregation_type))
    fc = (FILTER_CONDITION[filter_condition]
          if isinstance(filter_condition, str) else int(filter_condition))

    return {
        "AggregationType": agg,
        "ColumnFields": list(column_fields or []),
        "FilterCondition": fc,
        "Filters": filters_sorted,
        "LatestColumnsFirst": bool(latest_columns_first),
        "RowFields": list(row_fields or []),
        "ValueFields": list(value_fields or []),
        "columnDefinitions": col_defs,
    }


# ObjectMappingConfigDto field declaration order, copied from the C# source
# (Dtos/WebView/ObjectMappingConfigDto.cs).  Critical: serialize_camel
# preserves this order; alphabetizing it would produce a different JSON.
_MAPPING_FIELD_ORDER = (
    "selectedColumns",
    "selectedRows",
    "selectAllRows",
    "addQuestionText",
    "applyTranspose",
    "addLegend",
    "rowsPerObject",
    "addSplitObjectsToSingleSlide",
    "topNRows",
    "rowIdentifierColumnName",
    "customChartType",
    "tableHeaderMode",
    "insertEmptyColumnsAt",
    "moveRowsToFirst",
    "moveRowsToLast",
)


def _build_mapping_config(*, selected_columns, selected_rows, select_all_rows,
                          add_question_text, apply_transpose, add_legend,
                          rows_per_object, add_split_objects_to_single_slide,
                          top_n_rows, row_identifier_column_name, custom_chart_type,
                          table_header_mode, insert_empty_columns_at,
                          move_rows_to_first, move_rows_to_last) -> OrderedDict:
    """Build an ObjectMappingConfigDto OrderedDict ready for serialize_camel()."""
    if not selected_columns:
        raise ValueError("selected_columns must be a non-empty list.")

    cct = (CUSTOM_CHART[custom_chart_type] if isinstance(custom_chart_type, str)
           else custom_chart_type)
    thm = (TABLE_HEADER[table_header_mode] if isinstance(table_header_mode, str)
           else table_header_mode)

    d: OrderedDict = OrderedDict()
    d["selectedColumns"] = list(selected_columns)
    d["selectedRows"] = list(selected_rows or [])
    d["selectAllRows"] = bool(select_all_rows)
    d["addQuestionText"] = bool(add_question_text)
    d["applyTranspose"] = bool(apply_transpose)
    d["addLegend"] = bool(add_legend)
    d["rowsPerObject"] = int(rows_per_object) if rows_per_object else None
    d["addSplitObjectsToSingleSlide"] = bool(add_split_objects_to_single_slide)
    d["topNRows"] = int(top_n_rows) if top_n_rows else None
    d["rowIdentifierColumnName"] = row_identifier_column_name
    d["customChartType"] = cct
    d["tableHeaderMode"] = thm
    d["insertEmptyColumnsAt"] = insert_empty_columns_at
    d["moveRowsToFirst"] = list(move_rows_to_first or [])
    d["moveRowsToLast"] = list(move_rows_to_last or [])
    return d


def _enforce_mapping_field_order(mapping_dict: dict) -> OrderedDict:
    """Reorder a free-form mapping dict into the canonical declaration order.

    Use when the caller already has a mapping config dict (e.g. from
    ``propose_raw_configs()`` or a spec.json) and just wants the camelCase
    declaration-order serialization without rebuilding it.
    """
    out: OrderedDict = OrderedDict()
    for key in _MAPPING_FIELD_ORDER:
        if key in mapping_dict:
            out[key] = mapping_dict[key]
    # Carry through any extra keys the caller added at the tail (defensive)
    for key, val in mapping_dict.items():
        if key not in out:
            out[key] = val
    return out


# ─────────────────────────────────────────────────────────────────────────
# Top-level convenience: build all three JSONs + the two hashes in one call
# ─────────────────────────────────────────────────────────────────────────

def build_connector_payload(
    *,
    # ── Required: identify the connection ──
    project_id: int,
    reporting_plan_id: int,
    analysis_ids,
    selected_columns,
    # ── ReportConfig: optional with defaults ──
    survey_id=None,
    static_time_period_ids=None,
    static_time_period_names=None,
    latest_n_deliverables=None,
    include_live_wave=False,
    rollup_deliverables=False,
    segment_ids=(),
    nest_segments=False,
    include_alias=True,
    include_overall_segment=True,
    # ── PivotConfig ──
    filters=(),
    filter_condition="And",
    row_fields=(),
    column_fields=(),
    value_fields=(),
    aggregation_type="Sum",
    column_definitions=(),
    latest_columns_first=False,
    # ── ObjectMappingConfig ──
    selected_rows=(),
    select_all_rows=True,
    add_question_text=False,
    apply_transpose=False,
    add_legend=False,
    rows_per_object=None,
    add_split_objects_to_single_slide=False,
    top_n_rows=None,
    row_identifier_column_name=None,
    custom_chart_type=None,
    table_header_mode=None,
    insert_empty_columns_at=None,
    move_rows_to_first=(),
    move_rows_to_last=(),
) -> dict:
    """Build the three JSONs + two hashes for one connected shape.

    Returns
    -------
    dict with keys::

        report_json   : str  (PascalCase, alphabetical, nulls omitted)
        pivot_json    : str  (PascalCase, alphabetical, nulls omitted)
        mapping_json  : str  (camelCase, declaration order, nulls retained)
        report_hash   : str  (64-char lowercase hex SHA256 of report_json)
        pivot_hash    : str  (64-char lowercase hex SHA256 of pivot_json)

    The caller is responsible for writing these into shape tags + the
    deck-level CustomXMLPart entries (see doc §2 + §4).
    """
    # Cross-field validation (§11.13 / §13.13)
    if rows_per_object and rows_per_object > 0:
        if top_n_rows and top_n_rows <= rows_per_object:
            raise ValueError("top_n_rows must be > rows_per_object")
        if (not select_all_rows) and selected_rows \
                and len(selected_rows) <= rows_per_object:
            raise ValueError("len(selected_rows) must be > rows_per_object")

    report_dict = _build_report_config(
        project_id=project_id, survey_id=survey_id,
        reporting_plan_id=reporting_plan_id, analysis_ids=analysis_ids,
        segment_ids=segment_ids, nest_segments=nest_segments,
        static_time_period_ids=static_time_period_ids,
        static_time_period_names=static_time_period_names,
        latest_n_deliverables=latest_n_deliverables,
        include_live_wave=include_live_wave,
        rollup_deliverables=rollup_deliverables,
        include_alias=include_alias,
        include_overall_segment=include_overall_segment)

    pivot_dict = _build_pivot_config(
        filters=filters, filter_condition=filter_condition,
        row_fields=row_fields, column_fields=column_fields,
        value_fields=value_fields, aggregation_type=aggregation_type,
        column_definitions=column_definitions,
        latest_columns_first=latest_columns_first)

    mapping_dict = _build_mapping_config(
        selected_columns=selected_columns,
        selected_rows=selected_rows, select_all_rows=select_all_rows,
        add_question_text=add_question_text, apply_transpose=apply_transpose,
        add_legend=add_legend, rows_per_object=rows_per_object,
        add_split_objects_to_single_slide=add_split_objects_to_single_slide,
        top_n_rows=top_n_rows,
        row_identifier_column_name=row_identifier_column_name,
        custom_chart_type=custom_chart_type, table_header_mode=table_header_mode,
        insert_empty_columns_at=insert_empty_columns_at,
        move_rows_to_first=move_rows_to_first,
        move_rows_to_last=move_rows_to_last)

    report_json = serialize_pascal(report_dict)
    pivot_json = serialize_pascal(pivot_dict)
    mapping_json = serialize_camel(mapping_dict)

    return {
        "report_json": report_json,
        "pivot_json": pivot_json,
        "mapping_json": mapping_json,
        "report_hash": sha256_hex(report_json),
        "pivot_hash": sha256_hex(pivot_json),
    }


def serialize_existing_configs(
    report_dict: dict,
    pivot_dict: dict,
    mapping_dict: dict,
) -> dict:
    """Re-serialize already-built config dicts using the correct profiles.

    Use when the caller has manually assembled the config dicts (e.g. inside
    ``intelligent_refresh.write_connector_tags``) and just wants the doc's
    serialization rules applied to them.

    The function applies pre-hash sort rules, drops Alias when
    IsDefaultAlias=true, enforces mapping field declaration order, then
    serializes with the right profile per config.
    """
    # PreProcessConfig sorts on the report side
    report_normalized = dict(report_dict)
    if isinstance(report_normalized.get("AnalysisIds"), list) \
            and len(report_normalized["AnalysisIds"]) > 1:
        report_normalized["AnalysisIds"] = sorted(
            int(a) for a in report_normalized["AnalysisIds"])
    if (isinstance(report_normalized.get("SegmentIds"), list)
            and not report_normalized.get("NestSegments", False)
            and len(report_normalized["SegmentIds"]) > 1):
        report_normalized["SegmentIds"] = sorted(
            int(s) for s in report_normalized["SegmentIds"])
    if (report_normalized.get("TimePeriodType") == TIME_PERIOD_TYPE["Static"]
            and isinstance(report_normalized.get("StaticTimePeriodIds"), list)
            and len(report_normalized["StaticTimePeriodIds"]) > 1):
        report_normalized["StaticTimePeriodIds"] = sorted(
            int(t) for t in report_normalized["StaticTimePeriodIds"])

    # PivotConfig: filter sort + alias-when-default drop
    pivot_normalized = dict(pivot_dict)
    if isinstance(pivot_normalized.get("Filters"), list):
        pivot_normalized["Filters"] = sorted(
            list(pivot_normalized["Filters"]),
            key=lambda f: (f.get("ColumnKey", ""), f.get("Type", 0),
                           f.get("filterCriteria", 0), f.get("Value", "")))
    if isinstance(pivot_normalized.get("columnDefinitions"), list):
        pivot_normalized["columnDefinitions"] = [
            _drop_alias_when_default(c)
            for c in pivot_normalized["columnDefinitions"]
        ]

    # MappingConfig: enforce declaration order
    mapping_ordered = _enforce_mapping_field_order(mapping_dict)

    report_json = serialize_pascal(report_normalized)
    pivot_json = serialize_pascal(pivot_normalized)
    mapping_json = serialize_camel(mapping_ordered)

    return {
        "report_json": report_json,
        "pivot_json": pivot_json,
        "mapping_json": mapping_json,
        "report_hash": sha256_hex(report_json),
        "pivot_hash": sha256_hex(pivot_json),
    }
