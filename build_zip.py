"""Build the distributable zip. Usage: python build_zip.py"""
import zipfile, shutil
from pathlib import Path
from pricepulse import __version__

base = Path(__file__).parent
out = Path(r'C:\Users\mqia\Documents') / f'PricePulse-v{__version__}.zip'
includes = [
    'pricepulse/__init__.py', 'pricepulse/cli.py',
    'pricepulse/fetcher.py', 'pricepulse/reporter.py',
    'pyproject.toml', 'README.md', '.gitignore',
    'examples/sample_asins.csv',
    'skills/orcha/asin-pricepulse/SKILL.md',
    'skills/claude-code/asin-pricepulse/SKILL.md',
    'test_small.py', 'verify_brands.py',
]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in includes:
        fp = base / f
        assert fp.exists() and fp.stat().st_size > 0, f'missing/empty: {f}'
        zf.write(fp, 'PricePulse/' + f)
        print(f'  + {fp.stat().st_size:>6}  {f}')
print(f'\n{out}  ({out.stat().st_size // 1024} KB)')

# remove stale older zips
for old in Path(r'C:\Users\mqia\Documents').glob('PricePulse-v*.zip'):
    if old != out:
        old.unlink(); print(f'  removed old {old.name}')
