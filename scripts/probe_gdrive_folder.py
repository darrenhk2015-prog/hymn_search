"""Probe a public Google Drive folder structure."""
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

FOLDER_ID = '18uMCSYNDgipvGzHwlhyMcxrBaAcjChSZ'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode('utf-8', errors='replace')


def try_api_no_key():
    q = urllib.parse.quote(f"'{FOLDER_ID}' in parents and trashed=false")
    url = (
        'https://www.googleapis.com/drive/v3/files?'
        f'q={q}&fields=files(id,name,mimeType)&pageSize=100&supportsAllDrives=true'
    )
    try:
        data = json.loads(fetch(url))
        files = data.get('files', [])
        print('API no key:', len(files), 'files')
        for f in files[:10]:
            print(' ', f.get('name'), f.get('id'), f.get('mimeType', '')[:40])
        return files
    except urllib.error.HTTPError as e:
        print('API no key failed:', e.code, e.read()[:200])
    except Exception as e:
        print('API no key err:', e)
    return []


def try_embed_page():
    url = f'https://drive.google.com/drive/folders/{FOLDER_ID}'
    try:
        html = fetch(url)
        print('HTML len', len(html))
        # Drive often embeds file names in JSON blobs
        names = re.findall(r'"([^"]{2,120}\.(?:pdf|docx?|PDF|DOCX?))"', html)
        uniq = []
        for n in names:
            if n not in uniq:
                uniq.append(n)
        print('names in html', len(uniq))
        for n in uniq[:20]:
            print(' ', n)
        ids = re.findall(r'/"([a-zA-Z0-9_-]{25,44})"', html)
        print('slash ids', len(set(ids)))
    except Exception as e:
        print('HTML err:', e)


if __name__ == '__main__':
    try_api_no_key()
    print('---')
    try_embed_page()
