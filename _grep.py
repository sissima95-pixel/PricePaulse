import sys
pat = sys.argv[1:]
lines = open('pricepulse/reporter.py', encoding='utf-8').readlines()
for i, l in enumerate(lines):
    if any(p in l for p in pat):
        print(f'{i+1}|{l.rstrip()[:120]}')
