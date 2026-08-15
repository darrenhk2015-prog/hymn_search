#!/usr/bin/env python3
"""Build hymn_remote/data/gdrive_hymn_index.json from public Google Drive folders.

Usage:
  pip install gdown
  python scripts/build_gdrive_hymn_map.py

  # 只重建主資料夾（含 01–23 Word 等）:
  python scripts/build_gdrive_hymn_map.py --folder-id YOUR_FOLDER_ID

  # 略過 S1/S2 拆冊資料夾:
  python scripts/build_gdrive_hymn_map.py --no-split-folders

Output: hymn_remote/data/gdrive_hymn_index.json (+ copy beside project root)
Rebuild exe after updating (build_exe.ps1 bundles hymn_remote/data).
Drive folder must be shared as「知道連結的使用者均可檢視」.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from hymn_remote.gdrive_index_build import (  # noqa: E402
    DEFAULT_FOLDER_ID,
    SPLIT_BOOK_FOLDERS,
    build_index_payload,
)

OUT_PATH = ROOT / 'hymn_remote' / 'data' / 'gdrive_hymn_index.json'
EXTERNAL_PATH = ROOT / 'gdrive_hymn_index.json'


def main():
    ap = argparse.ArgumentParser(description='Build Google Drive hymn index JSON')
    ap.add_argument('--folder-id', default=DEFAULT_FOLDER_ID)
    ap.add_argument('--s1-folder-id', default=SPLIT_BOOK_FOLDERS[0][1])
    ap.add_argument('--s2-folder-id', default=SPLIT_BOOK_FOLDERS[1][1])
    ap.add_argument(
        '--no-split-folders',
        action='store_true',
        help='Skip S1/S2 split-PDF folders (use only --folder-id tree)',
    )
    ap.add_argument('-o', '--output', type=Path, default=OUT_PATH)
    ap.add_argument(
        '--no-external-copy',
        action='store_true',
        help=f'Do not also write {EXTERNAL_PATH.name} beside project root',
    )
    args = ap.parse_args()

    payload = build_index_payload(
        folder_id=args.folder_id,
        s1_folder_id=args.s1_folder_id,
        s2_folder_id=args.s2_folder_id,
        use_split_folders=not args.no_split_folders,
        progress_cb=print,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    args.output.write_text(text, encoding='utf-8')
    if not args.no_external_copy:
        EXTERNAL_PATH.write_text(text, encoding='utf-8')

    books = sorted({e['book'] for e in payload['entries'] if e['book']})
    print(f"Wrote {payload['file_count']} files, {len(books)} book paths -> {args.output}")
    if not args.no_external_copy:
        print(f'Also wrote {EXTERNAL_PATH}')
    for b in books[:12]:
        print(f'  {b}')
    if len(books) > 12:
        print(f'  ... +{len(books) - 12} more')


if __name__ == '__main__':
    main()
