import subprocess, json, re, unicodedata
toc = re.findall(r"\[\d, '([^']+)'\]", open('gen.js').read().split('const TOC = [')[1].split('];')[0])
n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo','out/Proyecto_Final_Piragua_Control.pdf'],capture_output=True,text=True).stdout).group(1))
res = {}
for pg in range(3, n+1):
    t = subprocess.run(['pdftotext','-f',str(pg),'-l',str(pg),'-layout','out/Proyecto_Final_Piragua_Control.pdf','-'],capture_output=True,text=True).stdout
    lines = [re.sub(r'\s+',' ',l).strip() for l in t.splitlines()]
    for h in toc:
        if h in res: continue
        if any(l.startswith(h) for l in lines): res[h] = pg
missing = [h for h in toc if h not in res]
print(json.dumps(res, ensure_ascii=False)); print('faltan', missing)
json.dump(res, open('pages.json','w'), ensure_ascii=False)
