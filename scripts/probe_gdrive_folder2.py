"""Extract file/folder entries from public Drive folder HTML."""
import json
import re
import urllib.request

FOLDER_ID = '18uMCSYNDgipvGzHwlhyMcxrBaAcjChSZ'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'


def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode('utf-8', errors='replace')


html = fetch(f'https://drive.google.com/drive/folders/{FOLDER_ID}')

# Pattern: ["FILE_ID","NAME",null,...
pat = re.compile(r'\["([a-zA-Z0-9_-]{20,50})","([^"\\]{1,200})"')
seen = {}
for m in pat.finditer(html):
    fid, name = m.group(1), m.group(2)
    if fid.startswith('http') or len(name) < 2:
        continue
    if fid not in seen or len(name) > len(seen[fid]):
        seen[fid] = name

items = sorted(seen.items(), key=lambda x: x[1])
print('entries', len(items))
for fid, name in items[:80]:
    print(fid, '|', name)
if len(items) > 80:
    print('...', len(items) - 80, 'more')

# MIME hints
apps = re.findall(r'application/vnd\.google-apps\.(\w+)', html)
print('google-apps types:', sorted(set(apps)))
