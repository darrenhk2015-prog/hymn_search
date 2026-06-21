"""Review local hymn formats vs cog_hymn_urls.json mapping."""
import json
import os
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from hymn_remote.web_hymn_map import (  # noqa: E402
    _load_mapping,
    extract_hymn_code,
    resolve_web_hymn,
)

JSON_PATH = ROOT / 'hymn_remote' / 'data' / 'cog_hymn_urls.json'
SETTINGS = ROOT / 'settings.json'
HYMN_FOLDER = ROOT / '神家詩歌集'


def scan_local_files(root: Path):
    books = []
    files = []
    if not root.is_dir():
        return books, files
    for entry in sorted(root.iterdir()):
        if entry.is_dir():
            books.append(entry.name)
            for f in entry.iterdir():
                if f.is_file():
                    files.append((entry.name, f.name))
        elif entry.is_file():
            books.append(entry.stem)
            files.append(('', entry.name))
    return books, files


def analyze_json_keys():
    data = json.loads(JSON_PATH.read_text(encoding='utf-8'))
    by_code = data.get('by_code', {})
    h_keys = [k for k in by_code if k.startswith('h')]
    s_keys = [k for k in by_code if k.startswith('s')]
    h_vol_pad = sum(1 for k in h_keys if re.match(r'h0\d-', k))
    h_vol_nopad = sum(1 for k in h_keys if re.match(r'h\d+-', k) and not re.match(r'h0\d-', k))
    h_hymn_pad = sum(1 for k in h_keys if re.search(r'-\d{2}$', k))
    h_hymn_3 = sum(1 for k in h_keys if re.search(r'-\d{3}$', k))
    return {
        'total': len(by_code),
        'h': len(h_keys),
        's': len(s_keys),
        'h_vol_zero_pad': h_vol_pad,
        'h_vol_no_pad': h_vol_nopad,
        'h_hymn_2digit': h_hymn_pad,
        'h_hymn_3digit': h_hymn_3,
        'sample_h': sorted(h_keys)[:8],
        'sample_s': sorted(s_keys)[:8],
    }


def test_case(book, num, label=''):
    r = resolve_web_hymn(book, num, label)
    code = r['code']
    mapping = _load_mapping()
    in_json = code in mapping if code else False
    return {
        'book': book,
        'num': num,
        'label': label,
        'code': code,
        'code_source': r.get('code_source', ''),
        'url_source': r['url_source'],
        'in_json': in_json,
        'web_url': bool(r['web_url']),
    }


def main():
    print('=== JSON key patterns ===')
    stats = analyze_json_keys()
    for k, v in stats.items():
        print(f'  {k}: {v}')

    print('\n=== Settings session_history ===')
    history = []
    if SETTINGS.is_file():
        history = json.loads(SETTINGS.read_text(encoding='utf-8')).get('session_history', [])
    ok = fail = 0
    for e in history:
        t = test_case(e.get('book', ''), e.get('num', ''), e.get('label', ''))
        status = 'OK' if t['web_url'] else 'MISS'
        if t['web_url']:
            ok += 1
        else:
            fail += 1
        print(f"  [{status}] {t['code'] or '—':8} {t['code_source']:7} json={t['in_json']}  "
              f"book={e.get('book','')[:20]!r} num={e.get('num','')[:35]!r}")
    if history:
        print(f'  => {ok}/{len(history)} resolved')

    print('\n=== Sample local files (first 40) ===')
    books, files = scan_local_files(HYMN_FOLDER)
    print(f'  books: {len(books)}, files: {len(files)}')
    file_ok = file_miss = 0
    miss_samples = []
    for book, fname in files[:200]:
        t = test_case(book, fname)
        if t['web_url']:
            file_ok += 1
        else:
            file_miss += 1
            if len(miss_samples) < 15:
                miss_samples.append((book, fname, t['code']))
    print(f'  => {file_ok}/{min(200, len(files))} resolved in sample')
    if miss_samples:
        print('  misses:')
        for b, f, c in miss_samples:
            print(f'    {b!r} / {f!r} -> {c}')

    print('\n=== Book name patterns ===')
    book_patterns = Counter()
    for b in books:
        if re.match(r'^\d+\s*神家', b):
            book_patterns['NN 神家詩歌'] += 1
        elif re.match(r'^\d+神家', b):
            book_patterns['NN神家詩歌'] += 1
        elif re.match(r'^[Ss]\d', b):
            book_patterns['S series'] += 1
        else:
            book_patterns['other'] += 1
    for k, v in book_patterns.most_common():
        print(f'  {k}: {v}')

    print('\n=== Filename patterns (sample 200) ===')
    fn_patterns = Counter()
    for _, fname in files[:200]:
        if re.search(r'\bH\d+-\d+', fname, re.I):
            fn_patterns['Hxx-yy in name'] += 1
        elif re.search(r'\bP\s*\d', fname, re.I):
            fn_patterns['Pnnn'] += 1
        elif re.search(r'^\d+\t', fname):
            fn_patterns['tab number'] += 1
        elif fname.lower().endswith(('.pdf', '.doc', '.docx', '.odt')):
            fn_patterns['other doc'] += 1
        else:
            fn_patterns['other'] += 1
    for k, v in fn_patterns.most_common():
        print(f'  {k}: {v}')

    print('\n=== Edge case tests ===')
    cases = [
        ('13神家詩歌', 'P52. 真良人萬般的情愛.docx', ''),
        ('13 神家詩歌', 'P52.真良人萬般的情愛.docx', ''),
        ('01 神家詩歌', 'P001  001 父的名.doc', ''),
        ('01 神家詩歌', 'P056  053 信心前征.doc', ''),
        ('05 神家詩歌', 'P02  耶和華尼西.doc', ''),
        ('22 神家詩歌', 'H22-07 歷世歷代', ''),
        ('S1神家詩歌合訂本1大字拍子版', '2\t祢這份愛', 'S1神家詩歌合訂本1大字拍子版 · 2\t祢這份愛 · P.21'),
        ('', '79\t萬物復興的前奏', ' · 79\t萬物復興的前奏 · P.146'),
        ('12 神家詩歌', 'P44.永遠愛的傷痕.docx', ''),
        ('08 神家詩歌', 'P04  最深親情！最高父愛！.doc', ''),
    ]
    for book, num, label in cases:
        t = test_case(book, num, label)
        print(f"  {t['code'] or '—':10} json={str(t['in_json']):5} src={t['code_source']:7} "
              f"url={t['url_source']:8} | {book[:15]} / {num[:30]}")


if __name__ == '__main__':
    main()
