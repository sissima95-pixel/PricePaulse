import re, json
from collections import Counter

html = open(r'test_output_us/test_v1.3.1_20260916_003540.html', encoding='utf-8').read()
m = re.search(r'const rowData=(\[.*?\]);', html, re.DOTALL)
if m:
    data = json.loads(m.group(1))
    brands = [r.get('brand','') for r in data if r.get('brand')]
    c = Counter(brands)
    blocklist = {'portable','wireless','mini','neck','smart','solar','outdoor',
                 'indoor','electric','rechargeable','handheld','cordless',
                 'desk','wall','travel','home','baby','pet','set','car'}
    problems = {b: n for b, n in c.items() if b.lower() in blocklist}
    if problems:
        print('STILL PROBLEMATIC:', problems)
    else:
        print('ALL CLEAN - no descriptive-word brands')
    print('\nTop 15 brands:')
    for brand, cnt in c.most_common(15):
        print(f'  {cnt:4d}  {brand}')

    # Check tier bounds sync
    m2 = re.search(r'const TIER_BOUNDS=(\{.*?\});', html, re.DOTALL)
    if m2:
        tb = json.loads(m2.group(1))
        print('\nTier bounds (from JS):')
        for mk, tiers in tb.items():
            for t in tiers:
                print(f'  {mk}: {t["name"]} ${t["lo"]:.0f}-${t["hi"]:.0f}')
