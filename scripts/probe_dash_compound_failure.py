"""Probe one P_DASH_COMPOUND_WAVE mapper-failure to find why Fix L
extension isn't fixing these charts. Picks ATU slide 29 'Clinical
remission sustained' (selectedColumns has 'Project Wave 13 - Options
- In bio-naïve patients' style entries).

Traces:
  1) The chart's selectedColumns + raw_pivot_config
  2) What Fix L's _rewrite_wave_pinned_selected_columns produces
  3) What the mapper's pivot output looks like (column names)
  4) Whether selectedColumns rewrite matches pivot column names
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import json
from slidegen.intelligent_refresh import fetch_synapse_data
from slidegen.synapse_chart_mapper import (
    _rewrite_wave_pinned_selected_columns,
    pivot_records_to_chart_data,
)

REPO = Path(__file__).resolve().parents[1]
spec = json.loads((REPO / "output/step2_test_connected/atu_q1_26_full_spec.json").read_text(encoding="utf-8"))

# Find slide 29 chart
slide = next(s for s in spec["slides"] if s["slide_index"] == 29)
chart = next(
    c for c in slide["components"]
    if c.get("type") == "chart"
    and any(" - " in s and "Project Wave" in s
            for s in c.get("raw_mapping_config", {}).get("selectedColumns", []))
)
ds_key = chart.get("data_source") or slide.get("data_source")
ds = spec["data_sources"][ds_key]
print(f"chart: {chart.get('name')!r}")
print(f"ds: {ds_key}")
print(f"ds config: {json.dumps(ds, indent=2)}")
print(f"\nselectedColumns ({len(chart['raw_mapping_config']['selectedColumns'])}):")
for sc in chart['raw_mapping_config']['selectedColumns']:
    print(f"  {sc!r}")
print(f"\nRowFields: {chart['raw_pivot_config']['RowFields']}")
print(f"ColumnFields: {chart['raw_pivot_config']['ColumnFields']}")
print(f"ValueFields: {chart['raw_pivot_config']['ValueFields']}")

# Fetch data
print("\n--- fetching data ---")
records, df = fetch_synapse_data(ds)
print(f"records: {len(records)}")
if not df.empty:
    print(f"df cols: {list(df.columns)}")
    print(f"distinct waves in df:")
    print(df.groupby('time_period_name').size())
    print(f"\nfirst row sample:")
    print(df.head(2).to_dict('records'))

# Try Fix L rewrite manually
print("\n--- Fix L rewrite ---")
sc_in = chart['raw_mapping_config']['selectedColumns']
sc_out, n_dropped = _rewrite_wave_pinned_selected_columns(
    sc_in, df, known_wave_labels=None,
)
print(f"dropped: {n_dropped}")
print(f"rewritten ({len(sc_out)}):")
for s in sc_out:
    print(f"  {s!r}")

# Run the full mapper
print("\n--- mapper output ---")
result = pivot_records_to_chart_data(
    records,
    chart['raw_pivot_config'],
    chart['raw_mapping_config'],
)
print(f"success: {result.success}")
print(f"error: {result.error}")
print(f"categories: {result.categories[:8] if result.categories else None}")
if result.series:
    for n, vals in result.series[:4]:
        print(f"  series {n!r}: {vals[:6]}")
