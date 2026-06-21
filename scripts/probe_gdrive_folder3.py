"""List public Drive folder via embedded web client API key."""
import json
import re
import urllib.parse
import urllib.request

FOLDER_ID = '18uMCSYNDgipvGzHwlhyMcxrBaAcjChSZ'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'


def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode('utf-8', errors='replace')


def list_children(parent_id, api_key):
    q = urllib.parse.quote(f"'{parent_id}' in parents and trashed=false")
    fields = urllib.parse.quote('nextPageToken,files(id,name,mimeType,size)')
    url = (
        f'https://www.googleapis.com/drive/v3/files?q={q}&fields={fields}'
        f'&pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true&key={api_key}'
    )
    all_files = []
    page_token = ''
    while True:
        u = url + (f'&pageToken={page_token}' if page_token else '')
        data = json.loads(fetch(u))
        all_files.extend(data.get('files', []))
        page_token = data.get('nextPageToken') or ''
        if not page_token:
            break
    return all_files


html = fetch(f'https://drive.google.com/drive/folders/{FOLDER_ID}')
keys = re.findall(r'AIzaSy[A-Za-z0-9_-]{20,}', html)
print('api keys found', len(set(keys)))
api_key = keys[0] if keys else ''
if not api_key:
    raise SystemExit('no api key')

root = list_children(FOLDER_ID, api_key)
print('root items', len(root))
for f in sorted(root, key=lambda x: x.get('name', '')):
    mime = f.get('mimeType', '')
    kind = 'folder' if 'folder' in mime else 'file'
    print(f"  [{kind}] {f['name']} | {f['id']}")

folders = [f for f in root if 'folder' in f.get('mimeType', '')]
for folder in folders[:5]:
    print('\n==', folder['name'], '==')
    kids = list_children(folder['id'], api_key)
    for f in sorted(kids, key=lambda x: x.get('name', ''))[:30]:
        print(' ', f['name'], '|', f['id'])
    if len(kids) > 30:
        print(' ...', len(kids) - 30, 'more')
