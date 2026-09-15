import re, sys, html as h
p = sys.argv[1]
s = open(p, encoding='utf-8').read()
m = re.search(r'C\. INDIVIDUAL PLAYER ANALYSIS.*?</div>\s*</div>', s, re.S)
txt = re.sub(r'<[^>]+>', ' ', m.group(0))
txt = h.unescape(re.sub(r'\s+', ' ', txt))
for chunk in re.split(r'(?=[A-Z]{3,}:)', txt):
    if chunk.strip():
        print(chunk.strip()[:400], '\n')
