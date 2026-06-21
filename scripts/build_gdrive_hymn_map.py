#!/usr/bin/env python3
"""Build hymn_remote/data/gdrive_hymn_index.json from a public Google Drive folder.

Usage:
  pip install gdown
  python scripts/build_gdrive_hymn_map.py

  # 換另一個 Drive 資料夾:
  python scripts/build_gdrive_hymn_map.py --folder-id YOUR_FOLDER_ID

Output: hymn_remote/data/gdrive_hymn_index.json
Rebuild exe after updating (build_exe.ps1 bundles hymn_remote/data).
Drive folder must be shared as「知道連結的使用者均可檢視」.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

DEFAULT_FOLDER_ID = '18uMCSYNDgipvGzHwlhyMcxrBaAcjChSZ'
DEFAULT_URL = f'https://drive.google.com/drive/folders/{DEFAULT_FOLDER_ID}?usp=sharing'
OUT_PATH = ROOT / 'hymn_remote' / 'data' / 'gdrive_hymn_index.json'


def _folder_url(folder_id: str) -> str:
    return f'https://drive.google.com/drive/folders/{folder_id}?usp=sharing'


def crawl_folder(url: str):
    try:
        import gdown
    except ImportError as exc:
        raise SystemExit('pip install gdown') from exc
    files = gdown.download_folder(url, skip_download=True, quiet=True)
    entries = []
    for item in files:
        rel = item.path.replace('\\', '/')
        book, name = os.path.split(rel)
        book = book.replace('\\', '/')
        entries.append({
            'id': item.id,
            'book': book,
            'name': name,
            'rel': rel,
        })
    entries.sort(key=lambda e: (e['book'].lower(), e['name'].lower()))
    return entries


def main():
    ap = argparse.ArgumentParser(description='Build Google Drive hymn index JSON')
    ap.add_argument('--folder-id', default=DEFAULT_FOLDER_ID)
    ap.add_argument('--url', default='')
    ap.add_argument('-o', '--output', type=Path, default=OUT_PATH)
    args = ap.parse_args()

    url = args.url or _folder_url(args.folder_id)
    print(f'Listing {url} ...')
    entries = crawl_folder(url)
    payload = {
        'version': 1,
        'folder_id': args.folder_id,
        'folder_url': url,
        'built_at': time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()),
        'file_count': len(entries),
        'entries': entries,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    books = sorted({e['book'] for e in entries if e['book']})
    print(f'Wrote {len(entries)} files, {len(books)} book paths -> {args.output}')
    for b in books[:12]:
        print(f'  {b}')
    if len(books) > 12:
        print(f'  ... +{len(books) - 12} more')


if __name__ == '__main__':
    main()
