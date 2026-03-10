import pandas as pd, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE_DIR, 'Lung SFEA SB.xlsx')
tag = pd.read_excel(XLSX, sheet_name='TAG', header=None)
ryb = pd.read_excel(XLSX, sheet_name='RYB', header=None)
aa  = pd.read_excel(XLSX, sheet_name='Additonal Analysis', header=None)

print('=== TAG Q2_10Z rows 223-236 with descriptions ===')
for i,r in tag.iloc[223:237].iterrows():
    desc = str(r[1])[:80] if str(r[1]) != 'nan' else ''
    print(f'{i} | code={r[0]} | Q3={r[7]} | Q4={r[13]} | desc: {desc}')

print('\n=== TAG C1_84F follow-up rows 138-145 ===')
for i,r in tag.iloc[138:146].iterrows():
    print(f'{i} | code={r[0]} | Q3={r[7]} | Q4={r[13]} | {str(r[1])[:50]}')

print('\n=== RYB Q1_50Z rows 75-92 (topics discussed) ===')
for i,r in ryb.iloc[75:92].iterrows():
    print(f'{i} | code={r[0]} | Q3={r[7]} | Q4={r[17]} | {str(r[1])[:60]}')

print('\n=== RYB CTA rows 120-155 (vals) ===')
for i,r in ryb.iloc[120:155].iterrows():
    print(f'{i} | code={r[0]} | Q3={r[7]} | Q4={r[17]} | {str(r[1])[:50]}')

print('\n=== TAG C1_85A prescribe rows 147-176 ===')
for i,r in tag.iloc[147:176].iterrows():
    print(f'{i} | code={r[0]} | Q3={r[7]} | Q4={r[13]} | {str(r[1])[:50]}')

print('\n=== AA rep performance rows 5-19 ===')
for i,r in aa.iloc[5:20].iterrows():
    print(f'{i} | metric={str(r[1])[:60]} | RYB_Q3={r[2]} | RYB_Q4={r[3]} | TAG_Q3={r[4]} | TAG_Q4={r[5]} | RYB_d={r[7]} | TAG_d={r[8]}')

print('\n=== AA message data rows 29-41 (all cols) ===')
for i,r in aa.iloc[29:41].iterrows():
    print(f'{i} | {[str(x)[:20] for x in r[:9]]}')

print('\n=== RYB Q2_10Z message rows 235-246 (all vals) ===')
for i,r in ryb.iloc[235:246].iterrows():
    print(f'{i} | code={r[0]} | Q3_tot={r[7]} | Q3_HI={r[10]} | Q4_tot={r[17]} | {str(r[1])[:60]}')
