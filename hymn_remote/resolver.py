"""Resolve book / hymn references to openable hymn payloads (no Qt)."""
import re

from hymn_features.fuzzy import fuzzy_match_score, rank_titles
from hymn_features.hymn_catalog import book_matches_query, hymn_display_label, hymn_payload

HYMN_NUM_MODE_CONTAINS = 'contains'
HYMN_NUM_MODE_EXACT = 'exact'
HYMN_NUM_MODES = (HYMN_NUM_MODE_CONTAINS, HYMN_NUM_MODE_EXACT)


def _extract_primary_number(text):
    """Primary hymn number from title/filename (ignores leading zeros)."""
    t = (text or '').strip()
    if not t:
        return None
    stem = re.sub(r'\.[A-Za-z0-9]+$', '', t)
    if stem.isdigit():
        return int(stem)
    m = re.match(r'^第\s*(\d+)\s*首', stem)
    if m:
        return int(m.group(1))
    m = re.match(r'^(\d+)(?:[.。、\s_\-]|$)', stem)
    if m:
        return int(m.group(1))
    nums = re.findall(r'\d+', stem)
    if not nums:
        return None
    if re.search(r'[-_]\d', stem):
        return int(nums[-1])
    return int(nums[0])


def hymn_number_exact_match(query, text):
    """Digit query equals primary number in text (004/04/4 all match 4; 14 does not)."""
    q = str(query or '').strip()
    if not q.isdigit():
        return False
    n = _extract_primary_number(text)
    return n is not None and n == int(q)


def _find_page_in_toc(toc_list, hymn_num, raw=''):
    """Match hymn number in TOC titles (supports H22-04 style codes)."""
    n = str(hymn_num)
    raw_s = (raw if raw not in (None, '') else n).strip()
    nz2 = raw_s.zfill(2) if raw_s.isdigit() else raw_s
    n3 = n.zfill(3)

    def _num_eq(candidate):
        if not candidate:
            return False
        if candidate in (raw_s, n, nz2, n3):
            return True
        if raw_s.isdigit() and candidate.isdigit():
            return int(candidate) == int(raw_s)
        return False

    for _lvl, title, page in toc_list:
        t = title.strip()
        if t in (n, n3, raw_s, nz2):
            return page, t, _lvl
        if t.startswith((
            n + '.', n + ' ', n + '。', n + '、',
            n3 + '.', n3 + ' ',
            raw_s + '.', raw_s + ' ', raw_s + '。', raw_s + '、',
        )):
            return page, t, _lvl
        if re.match(rf'^第\s*{re.escape(n)}\s*首', t):
            return page, t, _lvl

    suffixes = [s for s in (raw_s, nz2, n, n3) if s]
    for _lvl, title, page in toc_list:
        t = title.strip()
        for suffix in suffixes:
            if re.search(rf'[-_]{re.escape(suffix)}$', t, re.I):
                return page, t, _lvl

    for _lvl, title, page in toc_list:
        t = title.strip()
        nums = re.findall(r'\d+', t)
        if not nums:
            continue
        hyphenated = bool(re.search(r'[-_]\d', t))
        candidates = [nums[-1]] if hyphenated else [nums[0]]
        if len(nums) > 1 and not hyphenated:
            candidates.append(nums[-1])
        for cand in dict.fromkeys(candidates):
            if _num_eq(cand):
                return page, t, _lvl
    return None


def _parse_ref(ref):
    if ref is None:
        return None, ''
    s = str(ref).strip()
    if not s:
        return None, ''
    if s.isdigit():
        return int(s), s
    return None, s.lower()


def _hymns(book):
    return list(book.get('hymns') or [])


def _dedupe_targets(targets):
    seen = set()
    unique = []
    for t in targets:
        key = (t.get('kind'), t.get('hymn_id') or t.get('path') or t.get('pdf_path'), t.get('page'), t.get('title'))
        if key in seen:
            continue
        seen.add(key)
        unique.append(t)
    return unique


def _split_book_refs(book_ref):
    raw = str(book_ref or '').strip()
    if not raw:
        return []
    return [part.strip() for part in re.split(r'[,;|]+', raw) if part.strip()]


def list_books(books):
    return [
        {
            'index': i + 1,
            'name': b['name'],
            'volume_id': b.get('volume_id', ''),
            'hymn_count': b.get('hymn_count', 0),
            'bookmark': bool(b.get('bookmark')),
        }
        for i, b in enumerate(books)
    ]


def resolve_books(books, book_ref):
    """Match book by 1-based index or name keyword."""
    kind, val = _parse_ref(book_ref)
    if kind is not None:
        if 1 <= kind <= len(books):
            return [books[kind - 1]]
        return []
    if not val:
        return list(books)
    exact = [b for b in books if book_matches_query(b, book_ref)]
    if exact:
        numbered = [b for b in exact if str(b.get('volume_id') or '') == str(kind or val)]
        if numbered:
            return numbered
        return exact
    return [b for b in books if val in b['name'].lower()]


def list_book_entries(book, book_ref=''):
    """Hymns for mobile autocomplete / tag lists."""
    entries = []
    book_name = book.get('name', '')
    ref = str(book_ref or '').strip()
    for hymn in _hymns(book):
        label = hymn_display_label(hymn)
        entries.append({
            'kind': 'hymn',
            'value': label,
            'label': label,
            'hymn_id': hymn.get('id', ''),
            'book': book_name,
            'book_ref': ref,
        })
    return entries


def _filter_entries(entries, query, num_mode=HYMN_NUM_MODE_CONTAINS):
    q = (query or '').strip()
    if not q:
        return list(entries)
    mode = num_mode if num_mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS
    if mode == HYMN_NUM_MODE_EXACT and q.isdigit():
        return [
            e for e in entries
            if hymn_number_exact_match(q, e.get('value', ''))
            or hymn_number_exact_match(q, e.get('label', ''))
        ]
    ql = q.lower()
    out = []
    for e in entries:
        blob = ' '.join((
            str(e.get('value') or ''),
            str(e.get('label') or ''),
            str(e.get('book') or ''),
        ))
        if ql in blob.lower():
            out.append(e)
        elif fuzzy_match_score(q, e.get('value', '')) or fuzzy_match_score(q, e.get('label', '')):
            out.append(e)
    return out


def list_catalog_entries(books):
    """All hymns across volumes (iPad tag list with no book filter)."""
    out = []
    for idx, book in enumerate(books):
        out.extend(list_book_entries(book, book_ref=str(idx + 1)))
    return out


def list_entries_for_book(books, book_ref, query='', num_mode=HYMN_NUM_MODE_CONTAINS):
    refs = _split_book_refs(book_ref)
    q = (query or '').strip()
    if not refs:
        if not q:
            return list_catalog_entries(books)
        return list_global_entries(books, q, num_mode=num_mode)

    entries = []
    seen = set()
    for ref in refs:
        matched = resolve_books(books, ref)
        for book in matched:
            try:
                idx = str(books.index(book) + 1)
            except ValueError:
                idx = str(ref)
            for e in list_book_entries(book, book_ref=idx):
                key = e.get('hymn_id') or (e.get('book'), e.get('value'))
                if key in seen:
                    continue
                seen.add(key)
                entries.append(e)
    return _filter_entries(entries, q, num_mode=num_mode)


def list_global_entries(books, query='', num_mode=HYMN_NUM_MODE_CONTAINS, limit=80):
    """Search hymn titles across all books (mobile global mode)."""
    q = (query or '').strip()
    if not q:
        return []
    mode = num_mode if num_mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS
    ql = q.lower()
    scored = []
    for idx, book in enumerate(books):
        book_ref = str(idx + 1)
        book_name = book.get('name', '')
        for e in list_book_entries(book, book_ref=book_ref):
            value = e.get('value', '')
            label = e.get('label', '')
            score = 0
            if mode == HYMN_NUM_MODE_EXACT and q.isdigit():
                if hymn_number_exact_match(q, value) or hymn_number_exact_match(q, label):
                    score = 100
                else:
                    continue
            else:
                if ql in value.lower() or ql in label.lower():
                    score = 80
                else:
                    score = max(
                        fuzzy_match_score(q, value) or 0,
                        fuzzy_match_score(q, label) or 0,
                    )
                if not score:
                    continue
            scored.append((score, {
                'kind': e.get('kind', 'hymn'),
                'value': value,
                'label': f'{book_name} · {label}',
                'book': book_name,
                'book_ref': e.get('book_ref') or book_ref,
                'hymn_id': e.get('hymn_id', ''),
            }))
    scored.sort(key=lambda x: (-x[0], x[1].get('label', '')))
    return [item for _s, item in scored[: max(1, int(limit))]]


def _hymn_matches(hymn, raw, q, num_mode):
    label = hymn_display_label(hymn)
    number = hymn.get('number') or ''
    code = hymn.get('code') or ''
    title = hymn.get('title') or ''
    if num_mode == HYMN_NUM_MODE_EXACT and str(raw).strip().isdigit():
        return (
            hymn_number_exact_match(raw, number)
            or hymn_number_exact_match(raw, code)
            or hymn_number_exact_match(raw, label)
        )
    if str(raw).strip().isdigit():
        if (
            hymn_number_exact_match(raw, number)
            or hymn_number_exact_match(raw, code)
            or hymn_number_exact_match(raw, label)
        ):
            return True
    blob = f'{label} {title} {code}'.lower()
    if q and q in blob:
        return True
    return bool(fuzzy_match_score(raw, label) or fuzzy_match_score(raw, title))


def resolve_targets(book, num_ref, num_mode=HYMN_NUM_MODE_CONTAINS):
    """Return actionable hymn payloads (same shape as file_list UserRole)."""
    hymns = _hymns(book)
    raw = str(num_ref).strip() if num_ref is not None else ''
    if not raw:
        return [hymn_payload(book, h) for h in hymns]

    mode = num_mode if num_mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS
    q = raw.lower()
    hits = [hymn_payload(book, h) for h in hymns if _hymn_matches(h, raw, q, mode)]
    return _dedupe_targets(hits)


def resolve_open_request(books, book_ref, num_ref, num_mode=HYMN_NUM_MODE_CONTAINS):
    """Resolve across matched books; dedupe by hymn id."""
    matched_books = resolve_books(books, book_ref)
    if not matched_books:
        return []

    all_targets = []
    for book in matched_books:
        all_targets.extend(resolve_targets(book, num_ref, num_mode=num_mode))

    return _dedupe_targets(all_targets)


def summarize_match(t):
    kind = t.get('kind')
    book = t.get('book', '')
    if kind == 'hymn':
        name = t.get('name') or t.get('title') or t.get('code') or ''
        return f'{book} · {name}'.strip(' ·')
    if kind == 'bookmark':
        return f"{book} · {t.get('title')} · P.{t.get('page')}"
    if kind == 'file':
        return f"{book} · {t.get('name', '')}"
    return book


# Kept for callers that still import rank_titles through this module.
__all__ = [
    'HYMN_NUM_MODE_CONTAINS',
    'HYMN_NUM_MODE_EXACT',
    'HYMN_NUM_MODES',
    '_extract_primary_number',
    '_find_page_in_toc',
    'hymn_number_exact_match',
    'list_books',
    'list_book_entries',
    'list_catalog_entries',
    'list_entries_for_book',
    'list_global_entries',
    'rank_titles',
    'resolve_books',
    'resolve_open_request',
    'resolve_targets',
    'summarize_match',
]
