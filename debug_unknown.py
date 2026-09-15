"""Inspect why brand extraction fails for some ASINs."""
import re, sys
import openpyxl
from pricepulse.fetcher import MarketSession, _parse_price, PriceResult

DESC = {"portable","wireless","mini","smart","solar","outdoor","indoor","electric",
        "rechargeable","handheld","cordless","neck","desk","wall","car","baby","home",
        "travel","set","kit","usb","led","personal","small","large","dual"}

wb = openpyxl.load_workbook('test_output_us/asin_detail_20260811_111615.xlsx', data_only=True, read_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
hm = {str(h): i for i, h in enumerate(rows[0])}
unknown = [str(r[hm['ASIN']]) for r in rows[1:]
           if r[hm['Status']] == 'ok' and (not r[hm['Brand']] or str(r[hm['Brand']]).lower() in DESC)]
print(f'Unknown/desc-word brands: {len(unknown)}')

s = MarketSession('US')
s.prepare()
sample = unknown[:6]
for asin in sample:
    r = s.session.get(f'https://www.amazon.com/dp/{asin}/?th=1', timeout=25)
    html = r.text
    res = PriceResult(asin=asin, market='US', currency='USD')
    _parse_price(html, res)
    print(f'\n=== {asin} -> brand=[{res.brand}]  title=[{res.title[:60]}]')
    # Probe alternative patterns
    for name, pat in [
        ('bylineInfo', r'id="bylineInfo"[^>]*>([^<]{0,80})<'),
        ('po-brand', r'po-brand[\s\S]{0,400}?<span class="a-size-base po-break-word">([^<]+)<'),
        ('brand-td', r'<t[dh][^>]*>\s*Brand\s*</t[dh]>\s*<td[^>]*>\s*([^<]+?)\s*</td>'),
        ('brandname-th', r'Brand Name\s*</th>\s*<td[^>]*>\s*([^<]+?)\s*</td>'),
        ('jsonld', r'"brand"\s*:\s*\{[^}]*"name"\s*:\s*"([^"]+)"'),
        ('meta-brand', r'<meta[^>]*name="brand"[^>]*content="([^"]+)"'),
        ('stores-url', r'/stores/([A-Za-z0-9%\-]+)/page/'),
        ('brand-store-txt', r'Visit the ([^<]{1,40}?) Store'),
    ]:
        m = re.search(pat, html, re.I)
        if m:
            print(f'   {name:16} -> {m.group(1).strip()[:60]}')
