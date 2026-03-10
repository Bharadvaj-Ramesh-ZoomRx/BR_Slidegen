import pandas as pd, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE_DIR, 'Lung SFEA SB.xlsx')
tag = pd.read_excel(XLSX, sheet_name='TAG', header=None)
ryb = pd.read_excel(XLSX, sheet_name='RYB', header=None)
aa  = pd.read_excel(XLSX, sheet_name='Additonal Analysis', header=None)

print('=== TAG rows 222-280 cols 0,7,13 ===')
for i,r in tag.iloc[222:280].iterrows():
    print(i, str(r[0])[:50], '|Q3=',r[7],'|Q4=',r[13])

print('\n=== TAG CTA rows 80-134 ===')
for i,r in tag.iloc[80:134].iterrows():
    print(i, str(r[0])[:40], '|', str(r[1])[:40], '|Q3=',r[7],'|Q4=',r[13])

print('\n=== AA rows 57-66 ===')
for i,r in aa.iloc[57:67].iterrows():
    print(i, [str(x)[:25] for x in r[:9]])

print('\n=== AA rows 90-219 col1 ===')
for i,r in aa.iloc[90:219].iterrows():
    if str(r[1]) != 'nan':
        print(i, str(r[1])[:80])

print('\n=== RYB Q1_40_RYBZ rows 13-23 ===')
for i,r in ryb.iloc[13:23].iterrows():
    print(i, str(r[0])[:40], '|', str(r[1])[:60], '|Q3=',r[7],'|Q4=',r[17])

print('\n=== RYB C1_84FZ rows 166-175 ===')
for i,r in ryb.iloc[166:175].iterrows():
    print(i, str(r[0])[:40], '|', str(r[1])[:60], '|Q3=',r[7],'|Q4=',r[17])
