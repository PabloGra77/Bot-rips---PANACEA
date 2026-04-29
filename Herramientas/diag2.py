import re
with open('completar_diag.html', encoding='utf-8', errors='ignore') as f:
    html = f.read()
keys = re.findall(r"window\['([^']+)'\]\s*=", html)
kws = ['diag','buscar','boton','guardar','finalidad','causa','procedimiento']
print("=== window keys relevantes ===")
for k in sorted(set(keys)):
    if any(x in k.lower() for x in kws):
        print(k)
