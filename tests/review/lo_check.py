import re, subprocess, sys, os
f = sys.argv[1]; d = os.path.dirname(f) + '/lo4'
subprocess.run(['timeout','120','soffice','--headless','--convert-to','fodt',f,'--outdir',d], capture_output=True)
x = open(d + '/' + os.path.basename(f).replace('.docx','.fodt')).read()
names = {}
items = []
for a in re.findall(r'<office:annotation(?=[\s>]).*?</office:annotation>', x, re.S):
    head = re.match(r'<office:annotation[^>]*>', a).group(0)
    nm = re.search(r'office:name="([^"]+)"', head).group(1)
    par = re.search(r'parent-name="([^"]+)"', head)
    text = ' '.join(re.sub('<[^>]+>',' ',a.split('</dc:date>')[-1]).split()[1:])
    names[nm] = text
    items.append((nm, par.group(1) if par else None, 'resolved' if 'resolved="true"' in head else 'open', re.search(r'<dc:creator>(.*?)</dc:creator>', a).group(1), text))
for nm, par, st, who, text in items:
    print('   ', f'reply to "{names.get(par, "?")[:25]}"' if par else 'THREAD', '|', st, '|', who, '|', text[:45])
