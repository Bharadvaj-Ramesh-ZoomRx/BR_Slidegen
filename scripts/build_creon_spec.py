"""Build spec JSON for creon/2SlideInput.pptx non-connected refresh."""
from slidegen.intelligent_refresh import fetch_synapse_data, propose_raw_configs, read_slide_context
import json, os
import pandas as pd

# Fetch both datasets
lineage_205 = {'project_id': 523, 'reporting_plan_id': 1143, 'analysis_ids': [641205], 'segment_ids': [], 'dynamic_latest_n': 3}
lineage_211 = {'project_id': 523, 'reporting_plan_id': 1143, 'analysis_ids': [641211], 'segment_ids': [], 'dynamic_latest_n': 3}
_, df_205 = fetch_synapse_data(lineage_205)
_, df_211 = fetch_synapse_data(lineage_211)

ctx0 = read_slide_context('projects/creon/2SlideInput.pptx', 0)
ctx1 = read_slide_context('projects/creon/2SlideInput.pptx', 1)

pcp_cols = ['Internal Medicine (PCP)', 'Family Medicine (PCP)', 'General Medicine / Practice (PCP)']

# --- Get raw configs ---
t3_raw = propose_raw_configs(None, df_205, value_field='count', computed_columns=[{'name': 'PCPs', 'source_columns': pcp_cols}])
t4_raw = propose_raw_configs(None, df_205, value_field='count', computed_columns=[{'name': 'PCPs', 'source_columns': pcp_cols}])
t5_raw = propose_raw_configs(None, df_205, value_field='count')
t6_raw = propose_raw_configs(None, df_205, value_field='count')

# Fix selectedColumns
t3_raw['raw_mapping_config']['selectedColumns'] = ['Gastroenterology', '<blank:PCPs>']
t4_raw['raw_mapping_config']['selectedColumns'] = ['time_period_name', 'Gastroenterology', '<blank:PCPs>']
t5_raw['raw_mapping_config']['selectedColumns'] = ['time_period_name', 'Gastroenterology']

# Charts
pms_chart = [s for s in ctx1['shapes'] if s.get('name') == 'PMS'][0]
ps_chart = [s for s in ctx1['shapes'] if s.get('name') == 'PS'][0]
pms_raw = propose_raw_configs(pms_chart, df_205)
ps_raw = propose_raw_configs(ps_chart, df_211)

# --- Compute cell_values ---
latest = ["Jan'26", "Feb'26", "Mar'26"]
df3 = df_205[df_205['time_period_name'].isin(latest)]
pivot = df3.pivot_table(index='time_period_name', columns='option', values='count', aggfunc='first')
pivot['PCPs'] = pivot['Internal Medicine (PCP)'] + pivot['Family Medicine (PCP)'] + pivot['General Medicine / Practice (PCP)']
row_order = ["Jan'26", "Feb'26", "Mar'26"]
pivot = pivot.reindex(row_order)

cv3 = [['Gastros', 'PCPs']]
for p in row_order:
    cv3.append([str(int(pivot.loc[p, 'Gastroenterology'])), str(int(pivot.loc[p, 'PCPs']))])

cv4 = [['Deliverable', 'Gastroenterology', 'PCPs']]
for p in row_order:
    cv4.append([p, str(int(pivot.loc[p, 'Gastroenterology'])), str(int(pivot.loc[p, 'PCPs']))])

cv5 = [['Deliverable', 'Gastroenterology']]
for p in row_order:
    cv5.append([p, str(int(pivot.loc[p, 'Gastroenterology']))])

cv6 = [['GASTROS', '', 'NP/PAs']]
for p in row_order:
    cv6.append([p, str(int(pivot.loc[p, 'Gastroenterology'])), 'xx'])

# --- Build spec ---
spec = {
    'source_deck': '../2SlideInput.pptx',
    'data_sources': {
        'p523_rp1143_a641205': {
            'project_id': 523, 'reporting_plan_id': 1143,
            'analysis_ids': [641205], 'segment_ids': [], 'dynamic_latest_n': 3
        },
        'p523_rp1143_a641211': {
            'project_id': 523, 'reporting_plan_id': 1143,
            'analysis_ids': [641211], 'segment_ids': [], 'dynamic_latest_n': 3
        }
    },
    'slides': [
        {
            'slide_index': 0,
            'data_source': 'p523_rp1143_a641205',
            'components': [
                {
                    'type': 'value_table',
                    'name': 'Table 3',
                    'data_source': 'p523_rp1143_a641205',
                    'table_description': 'Gastros vs PCPs counts, no Deliverable column. PCPs = IM+FM+GM computed.',
                    'source_snapshot': {'headers': ['Gastros', 'PCPs'], 'all_rows': ctx0['shapes'][0]['all_rows']},
                    'cell_values': cv3,
                    'raw_pivot_config': t3_raw['raw_pivot_config'],
                    'raw_mapping_config': t3_raw['raw_mapping_config']
                },
                {
                    'type': 'value_table',
                    'name': 'Table 6',
                    'data_source': 'p523_rp1143_a641205',
                    'table_description': 'COMPLEX: GASTROS + NP/PAs. NP/PAs not in analysis 641205 - manually updated.',
                    'source_snapshot': {'headers': ['GASTROS', '', 'NP/PAs'], 'all_rows': ctx0['shapes'][1]['all_rows']},
                    'cell_values': cv6,
                    'complex': True,
                    'raw_pivot_config': t6_raw['raw_pivot_config'],
                    'raw_mapping_config': t6_raw['raw_mapping_config']
                },
                {
                    'type': 'value_table',
                    'name': 'Table 4',
                    'data_source': 'p523_rp1143_a641205',
                    'table_description': 'Deliverable / Gastroenterology / PCPs counts. PCPs = IM+FM+GM computed.',
                    'source_snapshot': {'headers': ['Deliverable', 'Gastroenterology', 'PCPs'], 'all_rows': ctx0['shapes'][2]['all_rows']},
                    'cell_values': cv4,
                    'raw_pivot_config': t4_raw['raw_pivot_config'],
                    'raw_mapping_config': t4_raw['raw_mapping_config']
                },
                {
                    'type': 'value_table',
                    'name': 'Table 5',
                    'data_source': 'p523_rp1143_a641205',
                    'table_description': 'Deliverable / Gastroenterology counts only.',
                    'source_snapshot': {'headers': ['Deliverable', 'Gastroenterology'], 'all_rows': ctx0['shapes'][3]['all_rows']},
                    'cell_values': cv5,
                    'raw_pivot_config': t5_raw['raw_pivot_config'],
                    'raw_mapping_config': t5_raw['raw_mapping_config']
                }
            ]
        },
        {
            'slide_index': 1,
            'data_source': None,
            'components': [
                {
                    'type': 'chart',
                    'name': 'PMS',
                    'chart_pattern': 'bar_stacked_100_horizontal',
                    'data_source': 'p523_rp1143_a641205',
                    'raw_pivot_config': pms_raw['raw_pivot_config'],
                    'raw_mapping_config': pms_raw['raw_mapping_config']
                },
                {
                    'type': 'chart',
                    'name': 'PS',
                    'chart_pattern': 'column_stacked_100_vertical',
                    'data_source': 'p523_rp1143_a641211',
                    'raw_pivot_config': ps_raw['raw_pivot_config'],
                    'raw_mapping_config': ps_raw['raw_mapping_config']
                }
            ]
        }
    ]
}

# Save spec
os.makedirs('projects/creon/output', exist_ok=True)
spec_path = 'projects/creon/output/2SlideInput_spec.json'
with open(spec_path, 'w') as f:
    json.dump(spec, f, indent=2)

print(f'Spec saved to {spec_path}')
print(f'Slides: {len(spec["slides"])}')
print(f'Components: {sum(len(s["components"]) for s in spec["slides"])}')
print(f'Data sources: {list(spec["data_sources"].keys())}')
