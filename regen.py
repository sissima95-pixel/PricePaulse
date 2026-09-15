import time
from pathlib import Path
from pricepulse.reporter import write_html
from pricepulse.fetcher import PriceResult
import openpyxl

wb = openpyxl.load_workbook('test_output_us/asin_detail_20260811_111615_brandfix.xlsx', data_only=True, read_only=True)
ws = wb.active
rows_raw = list(ws.iter_rows(values_only=True))
headers = [str(h) for h in rows_raw[0]]
hm = {h: i for i, h in enumerate(headers)}
results = [PriceResult(
    asin=str(row[hm['ASIN']] or ''), market=str(row[hm['Market']] or ''),
    currency=str(row[hm['Currency']] or ''),
    price=float(row[hm['Price']]) if row[hm['Price']] else None,
    display_price=str(row[hm['Display Price']] or ''),
    title=str(row[hm['Title']] or ''), brand=str(row[hm['Brand']] or ''),
    status=str(row[hm['Status']] or 'unknown'), url=str(row[hm['URL']] or ''),
    search_rank=float(row[hm['Search Rank']]) if row[hm['Search Rank']] else None,
    purchase_rank=float(row[hm['Purchase Rank']]) if row[hm['Purchase Rank']] else None,
    search_volume=float(row[hm['Search Volume']]) if row[hm['Search Volume']] else None,
) for row in rows_raw[1:]]

ts = time.strftime('%Y%m%d_%H%M%S')
out = Path(f'test_output_us/test_v1.3.5_{ts}.html')
write_html(results, out, title='US-Personal Fans ASIN Analysis', subtitle='2026.5.1 - 7.31')
print(f'Done: {out}  ({out.stat().st_size // 1024} KB)')
