import re
with open('completar_diag.html', encoding='utf-8', errors='ignore') as f:
    html = f.read()
print('HTML len:', len(html))
kw = re.compile(r'id="([^"]*)"', re.I)
ids = sorted(set(kw.findall(html)))
for x in ids:
    xl = x.lower()
    if any(k in xl for k in ['inalidad','ausa','uardar','iagnostic','buscar','diagn']):
        print('ID:', x)
print('---ASPxClient vars---')
for m in sorted(set(re.findall(r'var\s+(\w+)\s*=\s*new\s+ASPxClient', html))):
    print('VAR:', m)
print('---onclick con DoClick---')
for m in set(re.findall(r"window\['([^']+)'\]\.DoClick\(\)", html)):
    print('DOCLICK:', m)
