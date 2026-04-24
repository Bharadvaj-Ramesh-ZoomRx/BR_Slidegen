"""Stamp Connector tags on creon/2SlideInput_refreshed.pptx — skip Table 6 (complex)."""
from slidegen.intelligent_refresh import write_connector_tags
import json

pptx_path = 'projects/creon/output/2SlideInput_refreshed.pptx'

with open('projects/creon/output/2SlideInput_spec.json') as f:
    spec = json.load(f)

data_lineage_205 = {
    'project_id': 523, 'reporting_plan_id': 1143,
    'segment_ids': [], 'dynamic_latest_n': 3
}
data_lineage_211 = {
    'project_id': 523, 'reporting_plan_id': 1143,
    'segment_ids': [], 'dynamic_latest_n': 3
}

# --- Slide 0: Tables 3, 4, 5 (skip Table 6 = complex) ---
slide0_configs = []
for comp in spec['slides'][0]['components']:
    if comp['name'] == 'Table 6':
        print(f"SKIPPING {comp['name']} (complex — no Connector tags)")
        continue
    slide0_configs.append({
        'shape_name': comp['name'],
        'raw_pivot_config': comp['raw_pivot_config'],
        'raw_mapping_config': comp['raw_mapping_config'],
        'analysis_id': 641205
    })

print(f"\n=== Stamping Slide 0: {len(slide0_configs)} shapes ===")
result0 = write_connector_tags(
    pptx_path, slide_index=0,
    shape_configs=slide0_configs,
    data_lineage=data_lineage_205
)
for name, info in result0.items():
    print(f"  {name}: {info['action']}")

# --- Slide 1: PMS chart (641205) ---
pms_comp = spec['slides'][1]['components'][0]
pms_configs = [{
    'shape_name': pms_comp['name'],
    'raw_pivot_config': pms_comp['raw_pivot_config'],
    'raw_mapping_config': pms_comp['raw_mapping_config'],
    'analysis_id': 641205
}]

print(f"\n=== Stamping Slide 1: PMS chart ===")
result1a = write_connector_tags(
    pptx_path, slide_index=1,
    shape_configs=pms_configs,
    data_lineage=data_lineage_205
)
for name, info in result1a.items():
    print(f"  {name}: {info['action']}")

# --- Slide 1: PS chart (641211) ---
ps_comp = spec['slides'][1]['components'][1]
ps_configs = [{
    'shape_name': ps_comp['name'],
    'raw_pivot_config': ps_comp['raw_pivot_config'],
    'raw_mapping_config': ps_comp['raw_mapping_config'],
    'analysis_id': 641211
}]

print(f"\n=== Stamping Slide 1: PS chart ===")
result1b = write_connector_tags(
    pptx_path, slide_index=1,
    shape_configs=ps_configs,
    data_lineage=data_lineage_211
)
for name, info in result1b.items():
    print(f"  {name}: {info['action']}")

print("\nDone. Tags stamped on 2SlideInput_refreshed.pptx")
