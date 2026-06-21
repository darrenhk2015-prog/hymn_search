"""Map local hymn references to churchofgod.org.hk web pages."""
import json
import re
import sys
import time
from pathlib import Path

WEB_HYMN_BASE = 'https://churchofgod.org.hk/hymns/'
INTEGRATED_H_VOL_MAX = 10  # 冊 1–10 對應網站 S1/S2 合訂本，用歌名搵

_H_CODE = re.compile(r'\b([Hh])(\d{1,2})\s*[-_]\s*(\d{1,2})\b')
_S_CODE = re.compile(r'\b([Ss])(\d{1,2})\s*[-_]\s*(\d{1,3})\b')
_BOOK_H_VOL = re.compile(r'^(\d{1,2})\s*神家詩歌', re.IGNORECASE)
_BOOK_H_VOL_TIGHT = re.compile(r'^(\d{1,2})神家詩歌', re.IGNORECASE)
_BOOK_S_VOL = re.compile(r'^[Ss](\d{1,2})', re.IGNORECASE)
_P_NUM = re.compile(r'P\s*0*(\d{1,3})', re.IGNORECASE)
_TAB_NUM = re.compile(r'^(\d+)\t')
_DOC_EXT = re.compile(r'\.(docx?|pdf|odt)$', re.IGNORECASE)

_mapping_cache: dict | None = None
_mapping_mtime: float = 0.0
_s_title_index: list[tuple[str, str, dict]] | None = None

_TITLE_MATCH_MIN = 400


def _data_file() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / 'hymn_remote' / 'data' / 'cog_hymn_urls.json'
    return Path(__file__).resolve().parent / 'data' / 'cog_hymn_urls.json'


def normalize_hymn_code(text: str) -> str | None:
    text = (text or '').strip()
    if not text:
        return None
    m = _H_CODE.search(text)
    if m:
        return f'h{int(m.group(2))}-{int(m.group(3)):02d}'
    m = _S_CODE.search(text)
    if m:
        return f's{int(m.group(2))}-{int(m.group(3))}'
    return None


def _book_volume(book: str) -> tuple[str, int] | None:
    book = (book or '').strip()
    if not book:
        return None
    m = _BOOK_S_VOL.match(book)
    if m:
        return 's', int(m.group(1))
    m = _BOOK_H_VOL.match(book) or _BOOK_H_VOL_TIGHT.match(book)
    if m:
        return 'h', int(m.group(1))
    return None


def _h_book_volume(book: str) -> int | None:
    info = _book_volume(book)
    if info and info[0] == 'h':
        return info[1]
    return None


def _is_integrated_h_volume(book: str) -> bool:
    vol = _h_book_volume(book)
    return vol is not None and 1 <= vol <= INTEGRATED_H_VOL_MAX


def _p_number(text: str) -> int | None:
    text = (text or '').strip()
    if not text:
        return None
    m = _P_NUM.search(text)
    if m:
        return int(m.group(1))
    return None


def _clean_title(text: str) -> str:
    text = (text or '').strip()
    text = _DOC_EXT.sub('', text)
    text = re.sub(r'^[Pp]\s*0*\d+\s*[.\s]+', '', text)
    text = re.sub(r'^\d{3}\s+', '', text)
    text = re.sub(r'\s+', ' ', text).strip(' .·')
    return text


def _extract_song_title(book: str, num: str, label: str) -> str:
    for text in (num, label):
        raw = (text or '').strip()
        if not raw:
            continue
        if '\t' in raw:
            parts = raw.split('\t', 1)
            if len(parts) > 1 and parts[1].strip():
                title = _clean_title(parts[1])
                if title:
                    return title
        m = re.match(r'^[Pp]\s*0*\d+\s*[.\s]+(.+)$', raw)
        if m:
            title = _clean_title(m.group(1))
            if title:
                return title
        if raw.lower().endswith(('.doc', '.docx', '.pdf', '.odt')):
            title = _clean_title(raw)
            if title and not re.match(r'^[Pp]?\s*\d+$', title):
                return title
        if not re.match(r'^[Pp]?\s*\d+', raw):
            title = _clean_title(raw)
            if title:
                return title
    return ''


def _code_from_book_and_p(book: str, num: str, label: str) -> str | None:
    vol_info = _book_volume(book)
    if not vol_info or vol_info[0] != 'h':
        return None
    vol = vol_info[1]
    hymn_no = None
    for part in (num, label):
        hymn_no = _p_number(part)
        if hymn_no is not None:
            break
    if hymn_no is None:
        return None
    return f'h{vol}-{hymn_no:02d}'


def _code_from_s_book_tab(book: str, num: str, label: str) -> str | None:
    vol_info = _book_volume(book)
    if not vol_info or vol_info[0] != 's':
        return None
    _, vol = vol_info
    for part in (num, label):
        raw = (part or '').strip()
        m = _TAB_NUM.match(raw)
        if m:
            return f's{vol}-{int(m.group(1))}'
    return None


def _lookup_mapping(code: str) -> tuple[str, dict]:
    if not code:
        return '', {}
    mapping = _load_mapping()
    entry = mapping.get(code)
    if entry:
        return code, entry
    m = re.match(r'^h(\d+)-(\d+)$', code, re.IGNORECASE)
    if m:
        alt = f'h{int(m.group(1))}-{int(m.group(2))}'
        entry = mapping.get(alt)
        if entry:
            return alt, entry
    m = re.match(r'^s(\d+)-(\d+)$', code, re.IGNORECASE)
    if m:
        for alt in (
            f's{int(m.group(1))}-{int(m.group(2)):03d}',
            f's{int(m.group(1))}-{int(m.group(2))}',
        ):
            entry = mapping.get(alt)
            if entry:
                return alt, entry
    return code, {}


def _load_mapping() -> dict[str, dict]:
    global _mapping_cache, _mapping_mtime, _s_title_index
    path = _data_file()
    if not path.is_file():
        _mapping_cache = {}
        _mapping_mtime = 0.0
        _s_title_index = None
        return _mapping_cache
    mtime = path.stat().st_mtime
    if _mapping_cache is not None and mtime == _mapping_mtime:
        return _mapping_cache
    try:
        payload = json.loads(path.read_text(encoding='utf-8'))
        _mapping_cache = dict(payload.get('by_code') or {})
    except (OSError, json.JSONDecodeError, TypeError):
        _mapping_cache = {}
    _mapping_mtime = mtime
    _s_title_index = None
    return _mapping_cache


def _s_title_entries() -> list[tuple[str, str, dict]]:
    global _s_title_index
    if _s_title_index is not None:
        return _s_title_index
    mapping = _load_mapping()
    seen: set[str] = set()
    out: list[tuple[str, str, dict]] = []
    for code, entry in mapping.items():
        if not re.match(r'^s[12]-', code, re.IGNORECASE):
            continue
        title = (entry.get('title') or '').strip()
        if not title:
            continue
        lookup_code, canonical = _lookup_mapping(code)
        if lookup_code in seen:
            continue
        seen.add(lookup_code)
        out.append((lookup_code, title, canonical or entry))
    _s_title_index = out
    return out


def _resolve_by_title(title: str) -> tuple[str, dict, int, str]:
    title = (title or '').strip()
    if not title:
        return '', {}, 0, ''
    try:
        from hymn_features.fuzzy import fuzzy_match_score
    except ImportError:
        return '', {}, 0, title

    best_code, best_entry, best_score, best_title = '', {}, 0, ''
    for code, entry_title, entry in _s_title_entries():
        score = fuzzy_match_score(title, entry_title)
        if score > best_score:
            best_score = score
            best_code = code
            best_entry = entry
            best_title = entry_title
    if best_score < _TITLE_MATCH_MIN:
        return '', {}, best_score, title
    lookup_code, entry = _lookup_mapping(best_code)
    return lookup_code, entry or best_entry, best_score, best_title


def short_web_url(code: str) -> str:
    code = (code or '').strip().lower().rstrip('/')
    return f'{WEB_HYMN_BASE}{code}/'


def resolve_web_hymn(book: str = '', num: str = '', label: str = '') -> dict:
    book = (book or '').strip()
    num = (num or '').strip()
    label = (label or '').strip()
    if not label and book and num:
        label = f'{book} · {num}'
    elif not label and book:
        label = book

    code = ''
    code_source = ''
    mapped_entry: dict = {}
    match_score = 0
    matched_title = ''
    query_title = _extract_song_title(book, num, label)

    for part in (num, label, book):
        explicit = normalize_hymn_code(part)
        if explicit:
            code = explicit
            code_source = 'h_tag'
            break

    if not code:
        tab_code = _code_from_s_book_tab(book, num, label)
        if tab_code:
            code = tab_code
            code_source = 's_tab'

    if not code and _is_integrated_h_volume(book) and query_title:
        lc, entry, score, mt = _resolve_by_title(query_title)
        if lc:
            code = lc
            code_source = 's_title'
            mapped_entry = entry
            match_score = score
            matched_title = mt

    if not code:
        vol = _h_book_volume(book)
        if vol is not None and vol > INTEGRATED_H_VOL_MAX:
            bp = _code_from_book_and_p(book, num, label)
            if bp:
                code = bp
                code_source = 'book_p'

    if not code and not book and query_title:
        lc, entry, score, mt = _resolve_by_title(query_title)
        if lc:
            code = lc
            code_source = 's_title'
            mapped_entry = entry
            match_score = score
            matched_title = mt

    if not code:
        vol_info = _book_volume(book)
        if vol_info and vol_info[0] == 's' and query_title:
            lc, entry, score, mt = _resolve_by_title(query_title)
            if lc:
                code = lc
                code_source = 's_title'
                mapped_entry = entry
                match_score = score
                matched_title = mt

    lookup_code = code
    if code and not mapped_entry:
        lookup_code, mapped_entry = _lookup_mapping(code)

    mapped_url = (mapped_entry.get('url') or '').strip()
    fallback_url = short_web_url(lookup_code or code) if code else ''
    if mapped_url:
        url = mapped_url
        url_source = 'map'
    elif fallback_url:
        url = fallback_url
        url_source = 'fallback'
    else:
        url = ''
        url_source = 'none'

    title = (mapped_entry.get('title') or '').strip()
    if not title and label:
        title = label
    if code and title and code.upper() not in title.upper():
        title = f'{code.upper()}  {title}'

    return {
        'book': book,
        'num': num,
        'label': label,
        'query_title': query_title,
        'matched_title': matched_title,
        'match_score': match_score,
        'code': lookup_code or code or '',
        'code_source': code_source,
        'map_url': mapped_url,
        'fallback_url': fallback_url,
        'web_url': url,
        'url_source': url_source,
        'map_slug': mapped_entry.get('slug') or '',
        'title': title,
        'mapped': bool(mapped_entry),
        'ts': time.time(),
    }
