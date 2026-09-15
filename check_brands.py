import re, json, sys
from collections import Counter

path = sys.argv[1] if len(sys.argv) > 1 else 'test_output_us/test_v1.3.5_20260916_015312.html'
html = open(path, encoding='utf-8').read()
m = re.search(r'const rowData=(\[.*?\]);', html, re.DOTALL)
data = json.loads(m.group(1))
ok = [r for r in data if r.get('status') == 'ok' and r.get('price')]
brands = [r.get('brand') or 'Unknown' for r in ok]
c = Counter(brands)
print(f'Priced ASINs: {len(ok)}   Unknown: {c.get("Unknown",0)}   Distinct brands: {len(c)-("Unknown" in c)}')
print('\nTop 12 brands:')
for b, n in c.most_common(12):
    print(f'  {n:4d}  {b}')
m2 = re.search(r'const TIER_BOUNDS=(\{.*?\});', html, re.DOTALL)
for mk, tiers in json.loads(m2.group(1)).items():
    print(f'\nTier bounds {mk}: ' + ' | '.join(f'{t["name"]} ${t["lo"]:.0f}-${t["hi"]:.0f}' for t in tiers))
