"""Map local hymn references to Google Drive preview URLs."""
from __future__ import annotations

import json
import re
import shutil
import sys
import time
from pathlib import Path

from hymn_remote.gdrive_index_build import DEFAULT_FOLDER_ID
from hymn_remote.web_hymn_map import _extract_song_title, _p_number

GDRIVE_PREVIEW = 'https://drive.google.com/file/d/{file_id}/preview'
INDEX_FILENAME = 'gdrive_hymn_index.json'

_P_NUM = re.compile(r'P\s*[.]?\s*0*(\d{1,3})', re.IGNORECASE)
_TAB_NUM = re.compile(r'^(\d+)[\t ]+')
_BOOK_H_VOL = re.compile(r'^(\d{1,2})\s*神家詩歌', re.IGNORECASE)
_BOOK_H_VOL_TIGHT = re.compile(r'^(\d{1,2})神家詩歌', re.IGNORECASE)
_DOC_EXT = re.compile(r'\.(docx?|pdf|odt)$', re.IGNORECASE)
_TITLE_VARIANTS = str.maketrans({'祢': '你', '妳': '你', '祂': '他'})

_BOOK_S_VOL = re.compile(r'^[Ss](\d{1,2})', re.IGNORECASE)
_BOOK_S_ALIAS = re.compile(r'^神家詩歌合訂本\s*([12])(?!\d)', re.IGNORECASE)
INTEGRATED_H_VOL_MAX = 10

_index_cache: dict | None = None
_index_mtime: float = 0.0
_book_entries: dict[str, list[dict]] | None = None
_active_index_path: Path | None = None


def _app_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def bundled_index_path() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / 'hymn_remote' / 'data' / INDEX_FILENAME
    return Path(__file__).resolve().parent / 'data' / INDEX_FILENAME


def external_index_path() -> Path:
    return _app_dir() / INDEX_FILENAME


def index_path() -> Path:
    """Prefer editable external config beside exe / project root."""
    external = external_index_path()
    if external.is_file():
        return external
    return bundled_index_path()


def ensure_external_index() -> Path:
    """Copy bundled index beside exe on first run so it can be edited."""
    external = external_index_path()
    if external.is_file():
        return external
    bundled = bundled_index_path()
    if bundled.is_file():
        try:
            shutil.copy2(bundled, external)
        except OSError:
            return bundled if bundled.is_file() else external
    return external if external.is_file() else bundled


def invalidate_index_cache() -> None:
    global _index_cache, _index_mtime, _book_entries, _active_index_path
    _index_cache = None
    _index_mtime = 0.0
    _book_entries = None
    _active_index_path = None


def _data_file() -> Path:
    return index_path()


def load_index_payload() -> dict:
    """Public accessor for the active index JSON (cached)."""
    return _load_index()


def get_index_info() -> dict:
    path = index_path()
    payload = _load_index()
    splits = payload.get('split_folders') or []
    s1_id = s2_id = ''
    for row in splits:
        book = str(row.get('book') or '')
        fid = str(row.get('folder_id') or '')
        if book.startswith('S1') and fid:
            s1_id = fid
        elif book.startswith('S2') and fid:
            s2_id = fid
    return {
        'path': str(path),
        'external': path.resolve() == external_index_path().resolve() and path.is_file(),
        'exists': path.is_file(),
        'built_at': payload.get('built_at') or '',
        'file_count': int(payload.get('file_count') or len(payload.get('entries') or [])),
        'folder_id': payload.get('folder_id') or DEFAULT_FOLDER_ID,
        's1_folder_id': s1_id,
        's2_folder_id': s2_id,
        'split_folders': splits,
    }


def save_index_payload(payload: dict, path: Path | None = None) -> Path:
    """Write index JSON and refresh in-memory cache."""
    global _index_cache, _index_mtime, _book_entries, _active_index_path
    out = path or external_index_path()
    out.parent.mkdir(parents=True, exist_ok=True)
    entries = list(payload.get('entries') or [])
    entries.sort(key=lambda e: ((e.get('book') or '').lower(), (e.get('name') or '').lower()))
    payload = dict(payload)
    payload['entries'] = entries
    payload['file_count'] = len(entries)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    _index_cache = payload
    _index_mtime = out.stat().st_mtime
    _book_entries = None
    _active_index_path = out
    return out


def _norm_spaces(text: str) -> str:
    return re.sub(r'\s+', ' ', (text or '').strip())


def _norm_key(text: str) -> str:
    return _norm_spaces(text).lower()


def _norm_title(text: str) -> str:
    cleaned = _DOC_EXT.sub('', _norm_spaces(text))
    cleaned = re.sub(r'^[Pp]\s*0*\d+\s*[.\s]+', '', cleaned)
    cleaned = re.sub(r'^\d{3}\s+', '', cleaned)
    return cleaned.translate(_TITLE_VARIANTS)


def _norm_book(book: str) -> str:
    book = _norm_spaces(book)
    m = _BOOK_H_VOL.match(book) or _BOOK_H_VOL_TIGHT.match(book)
    if m:
        return f'{int(m.group(1)):02d} 神家詩歌'
    return book


def _book_keys(book: str) -> list[str]:
    book = _norm_spaces(book)
    if not book:
        return []
    keys = [_norm_key(book)]
    norm = _norm_book(book)
    if _norm_key(norm) not in keys:
        keys.append(_norm_key(norm))
    return keys


def _load_index() -> dict:
    global _index_cache, _index_mtime, _book_entries, _active_index_path
    path = _data_file()
    if not path.is_file():
        _index_cache = {}
        _index_mtime = 0.0
        _book_entries = {}
        _active_index_path = path
        return _index_cache
    mtime = path.stat().st_mtime
    if (
        _index_cache is not None
        and mtime == _index_mtime
        and _active_index_path == path
    ):
        return _index_cache
    try:
        _index_cache = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError, TypeError):
        _index_cache = {}
    _index_mtime = mtime
    _book_entries = None
    _active_index_path = path
    return _index_cache


def _entries_by_book() -> dict[str, list[dict]]:
    global _book_entries
    if _book_entries is not None:
        return _book_entries
    payload = _load_index()
    grouped: dict[str, list[dict]] = {}
    for entry in payload.get('entries') or []:
        book = _norm_spaces(entry.get('book') or '')
        name = entry.get('name') or ''
        file_id = entry.get('id') or ''
        if not file_id or not name:
            continue
        row = {
            'id': file_id,
            'book': book,
            'name': name,
            'rel': entry.get('rel') or '',
            'name_key': _norm_key(name),
        }
        grouped.setdefault(book, []).append(row)
        if book:
            leaf = book.split('/')[-1].strip()
            if leaf and leaf != book:
                grouped.setdefault(leaf, []).append(row)
    _book_entries = grouped
    return _book_entries


def _preview_url(file_id: str) -> str:
    return GDRIVE_PREVIEW.format(file_id=file_id)


def _s_book_volume(book: str) -> int | None:
    book = _norm_spaces(book)
    if not book:
        return None
    m = _BOOK_S_VOL.match(book)
    if m:
        return int(m.group(1))
    m = _BOOK_S_ALIAS.match(book)
    if m:
        return int(m.group(1))
    return None


def _is_s_book(book: str) -> bool:
    vol = _s_book_volume(book)
    return vol is not None and vol in (1, 2)


def _integrated_entries() -> list[dict]:
    grouped = _entries_by_book()
    out: list[dict] = []
    seen: set[str] = set()
    for vol in range(1, INTEGRATED_H_VOL_MAX + 1):
        prefix = f'{vol:02d} 神家詩歌'
        for folder, rows in grouped.items():
            if not folder.startswith(prefix):
                continue
            for row in rows:
                if row['id'] in seen:
                    continue
                seen.add(row['id'])
                out.append(row)
    return out


def _pick_book_entries(book: str) -> list[dict]:
    grouped = _entries_by_book()
    out: list[dict] = []
    seen: set[str] = set()
    for key in _book_keys(book):
        for folder, rows in grouped.items():
            folder_key = _norm_key(folder)
            if folder_key == key or folder_key.startswith(key + '/') or key in folder_key:
                for row in rows:
                    if row['id'] in seen:
                        continue
                    seen.add(row['id'])
                    out.append(row)
    if out:
        return out
    vol = _norm_book(book)
    vol_key = _norm_key(vol)
    for folder, rows in grouped.items():
        if _norm_key(folder).startswith(vol_key):
            for row in rows:
                if row['id'] in seen:
                    continue
                seen.add(row['id'])
                out.append(row)
    return out


def _tab_number(*parts: str) -> int | None:
    for part in parts:
        raw = (part or '').strip()
        if not raw:
            continue
        m = _TAB_NUM.match(raw)
        if m:
            return int(m.group(1))
    return None


def _filename_from_parts(num: str, label: str, filename: str) -> str:
    for part in (filename, num, label):
        raw = (part or '').strip()
        if not raw:
            continue
        base = raw.replace('\\', '/').split('/')[-1]
        if _DOC_EXT.search(base):
            return base
    return ''


def _p_numbers(text: str) -> list[int]:
    nums = []
    for m in _P_NUM.finditer(text or ''):
        n = int(m.group(1))
        if n not in nums:
            nums.append(n)
    p = _p_number(text or '')
    if p is not None and p not in nums:
        nums.append(p)
    return nums


def _match_by_filename(entries: list[dict], filename: str) -> dict | None:
    key = _norm_key(filename)
    if not key:
        return None
    for row in entries:
        if row['name_key'] == key:
            return row
    return None


def _match_by_p_number(entries: list[dict], p_num: int, title: str = '') -> dict | None:
    if p_num is None:
        return None
    p3 = f'p{p_num:03d}'
    p2 = f'p{p_num:02d}'
    p1 = f'p{p_num}'
    candidates = []
    for row in entries:
        name = row['name_key']
        if name.startswith(p3) or name.startswith(p2) or name.startswith(p1):
            candidates.append(row)
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]
    if title:
        try:
            from hymn_features.fuzzy import fuzzy_match_score
        except ImportError:
            return candidates[0]
        best, best_score = None, 0
        for row in candidates:
            score = fuzzy_match_score(title, row['name'])
            if score > best_score:
                best_score = score
                best = row
        if best and best_score >= 300:
            return best
    return candidates[0]


def _pick_candidates(entries: list[dict], title: str = '') -> dict | None:
    if not entries:
        return None
    if len(entries) == 1:
        return entries[0]
    if title:
        try:
            from hymn_features.fuzzy import fuzzy_match_score
        except ImportError:
            return entries[0]
        best, best_score = None, 0
        for row in entries:
            score = fuzzy_match_score(title, row['name'])
            if score > best_score:
                best_score = score
                best = row
        if best and best_score >= 300:
            return best
    return entries[0]


def _match_by_tab_number(entries: list[dict], tab_num: int | None, title: str = '') -> dict | None:
    if tab_num is None:
        return None
    prefixes = (f'{tab_num} ', f'{tab_num:02d} ', f'{tab_num:03d} ')
    candidates = [row for row in entries if row['name'].startswith(prefixes)]
    return _pick_candidates(candidates, title)


def _match_by_page_marker(entries: list[dict], p_num: int | None, title: str = '') -> dict | None:
    if p_num is None:
        return None
    markers = tuple(dict.fromkeys([
        f'_p.{p_num}', f'_p{p_num}', f'_p{p_num:02d}', f'_p{p_num:03d}',
    ]))
    candidates = []
    for row in entries:
        name_key = row['name_key']
        if any(marker in name_key for marker in markers):
            candidates.append(row)
    return _pick_candidates(candidates, title)


def _match_by_title(
    entries: list[dict],
    title: str,
    *,
    min_score: int = 400,
) -> tuple[dict | None, int, str]:
    title = (title or '').strip()
    if not title or not entries:
        return None, 0, title
    try:
        from hymn_features.fuzzy import fuzzy_match_score
    except ImportError:
        return None, 0, title
    query = _norm_title(title)
    best, best_score, best_name = None, 0, ''
    for row in entries:
        candidate = _norm_title(row['name'])
        score = fuzzy_match_score(query, candidate)
        if score > best_score:
            best_score = score
            best = row
            best_name = row['name']
    if best_score < min_score:
        return None, best_score, title
    return best, best_score, best_name


def resolve_gdrive_hymn(
    book: str = '',
    num: str = '',
    label: str = '',
    filename: str = '',
) -> dict:
    book = (book or '').strip()
    num = (num or '').strip()
    label = (label or '').strip()
    filename = (filename or '').strip()
    if not label and book and num:
        label = f'{book} · {num}'
    elif not label and book:
        label = book

    query_title = _extract_song_title(book, num, label)
    entries = _pick_book_entries(book)
    match: dict | None = None
    match_source = ''
    match_score = 0
    matched_name = ''

    fname = _filename_from_parts(num, label, filename)
    if fname and entries:
        match = _match_by_filename(entries, fname)
        if match:
            match_source = 'filename'

    if not match and entries and _is_s_book(book):
        tab_num = _tab_number(num, label)
        if tab_num is not None:
            match = _match_by_tab_number(entries, tab_num, query_title)
            if match:
                match_source = 's_tab'

    if not match and entries and _is_s_book(book):
        for part in (label, num):
            for p_num in _p_numbers(part):
                match = _match_by_page_marker(entries, p_num, query_title)
                if match:
                    match_source = 's_page'
                    break
            if match:
                break

    if not match and entries and not _is_s_book(book):
        for part in (num, label, fname):
            for p_num in _p_numbers(part):
                match = _match_by_p_number(entries, p_num, query_title)
                if match:
                    match_source = 'p_number'
                    break
            if match:
                break

    if not match and query_title and entries:
        match, match_score, matched_name = _match_by_title(entries, query_title)
        if match:
            match_source = 'title'

    if not match and _is_s_book(book) and query_title:
        integrated = _integrated_entries()
        match, match_score, matched_name = _match_by_title(
            integrated, query_title, min_score=350,
        )
        if match:
            match_source = 's_title'

    file_id = (match or {}).get('id', '')
    rel = (match or {}).get('rel', '')
    drive_name = (match or {}).get('name', '')
    web_url = _preview_url(file_id) if file_id else ''
    title = query_title or drive_name or label
    if drive_name and title and _norm_key(title) not in _norm_key(drive_name):
        title = f'{title}  ·  {drive_name}'

    payload = _load_index()
    return {
        'book': book,
        'num': num,
        'label': label,
        'query_title': query_title,
        'matched_title': matched_name,
        'match_score': match_score,
        'code': '',
        'code_source': match_source,
        'map_url': web_url,
        'fallback_url': '',
        'web_url': web_url,
        'url_source': 'gdrive' if web_url else 'none',
        'map_slug': rel,
        'title': title,
        'mapped': bool(file_id),
        'drive_file_id': file_id,
        'drive_name': drive_name,
        'drive_rel': rel,
        'drive_folder_id': payload.get('folder_id') or DEFAULT_FOLDER_ID,
        'ts': time.time(),
    }
