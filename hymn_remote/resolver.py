"""Resolve book / hymn references to openable file payloads (no Qt)."""
import re

from hymn_features.fuzzy import fuzzy_match_score, rank_titles

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


def _file_target(fi, book):
    return {
        'kind': 'file',
        'path': fi['path'],
        'book': book['name'],
        'name': fi['name'],
    }


def _bookmark_target(book, fi, title, page, toc_level=None):
    return {
        'kind': 'bookmark',
        'title': title,
        'page': page,
        'pdf_path': fi['path'],
        'book': book['name'],
        'toc_level': toc_level,
    }


def _pdf_files(book):
    return [fi for fi in book.get('files', []) if fi['ext'] == '.pdf']


def _toc_entries(book, hymns_only=False):
    toc = book.get('toc') or []
    if not hymns_only:
        return toc
    if any(lvl == 2 for lvl, _, _ in toc):
        return [(lvl, title, page) for lvl, title, page in toc if lvl == 2]
    return toc


def _dedupe_targets(targets):
    seen = set()
    unique = []
    for t in targets:
        key = (t.get('kind'), t.get('path') or t.get('pdf_path'), t.get('page'), t.get('title'))
        if key in seen:
            continue
        seen.add(key)
        unique.append(t)
    return unique


def _prefer_bookmark_targets(targets):
    """Prefer a single level-2 hymn over PDF file / level-1 root."""
    if len(targets) <= 1:
        return targets
    bookmarks = [t for t in targets if t.get('kind') == 'bookmark']
    if not bookmarks:
        return targets
    lvl2 = [t for t in bookmarks if t.get('toc_level') == 2]
    if len(lvl2) == 1:
        return lvl2
    if lvl2:
        return lvl2
    if len(bookmarks) == 1:
        return bookmarks
    return bookmarks


def list_books(books):
    return [
        {
            'index': i + 1,
            'name': b['name'],
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
    return [b for b in books if val in b['name'].lower()]


def list_book_entries(book):
    """Files and bookmarks for mobile autocomplete."""
    entries = []
    for fi in book.get('files', []):
        ext = fi['ext']
        tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
        entries.append({
            'kind': 'file',
            'value': fi['name'],
            'label': f'[{tag}] {fi["name"]}',
        })
    if book.get('bookmark') and book.get('toc'):
        seen = set()
        for _lvl, title, page in _toc_entries(book, hymns_only=True):
            t = title.strip()
            if not t or t in seen:
                continue
            seen.add(t)
            entries.append({
                'kind': 'bookmark',
                'value': t,
                'label': f'{t} · P.{page}',
            })
    return entries


def list_entries_for_book(books, book_ref, query='', num_mode=HYMN_NUM_MODE_CONTAINS):
    book_ref_s = str(book_ref or '').strip()
    q = (query or '').strip()
    # No book selected → global search by hymn/file name (desktop「全局」)
    if not book_ref_s:
        return list_global_entries(books, q, num_mode=num_mode)

    matched = resolve_books(books, book_ref)
    if not matched:
        return []
    entries = list_book_entries(matched[0])
    if not q:
        return entries
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
        if ql in e['value'].lower() or ql in e['label'].lower():
            out.append(e)
        elif fuzzy_match_score(q, e['value']) or fuzzy_match_score(q, e['label']):
            out.append(e)
    return out


def list_global_entries(books, query='', num_mode=HYMN_NUM_MODE_CONTAINS, limit=80):
    """Search hymn/file titles across all books (mobile global mode)."""
    q = (query or '').strip()
    if not q:
        return []
    mode = num_mode if num_mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS
    ql = q.lower()
    scored = []
    for idx, book in enumerate(books):
        book_ref = str(idx + 1)
        book_name = book.get('name', '')
        for e in list_book_entries(book):
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
                'kind': e.get('kind', 'file'),
                'value': value,
                'label': f'{book_name} · {label}',
                'book': book_name,
                'book_ref': book_ref,
            }))
    scored.sort(key=lambda x: (-x[0], x[1].get('label', '')))
    return [item for _s, item in scored[: max(1, int(limit))]]


def _bookmark_hits(book, q, raw, num_mode=HYMN_NUM_MODE_CONTAINS):
    """Match bookmark titles; never include the bare PDF file."""
    hits = []
    exact_num = (
        num_mode == HYMN_NUM_MODE_EXACT and str(raw).strip().isdigit()
    )
    for fi in _pdf_files(book):
        if not book.get('bookmark') or not book.get('toc'):
            continue

        if exact_num:
            for lvl, title, page in _toc_entries(book, hymns_only=True):
                if hymn_number_exact_match(raw, title):
                    hits.append(_bookmark_target(book, fi, title.strip(), page, lvl))
            continue

        if raw.isdigit():
            found = _find_page_in_toc(book['toc'], int(raw), raw)
            if found:
                page, title, lvl = found
                hits.append(_bookmark_target(book, fi, title, page, lvl))

        if hits:
            continue

        exact = []
        partial = []
        fuzzy = []
        for lvl, title, page in book['toc']:
            t = title.strip()
            tl = t.lower()
            if tl == q:
                exact.append(_bookmark_target(book, fi, t, page, lvl))
            elif q in tl:
                partial.append(_bookmark_target(book, fi, t, page, lvl))
            else:
                score = fuzzy_match_score(raw, t)
                if score > 0:
                    fuzzy.append((score, _bookmark_target(book, fi, t, page, lvl)))

        if exact:
            hits.extend(exact)
        elif partial:
            hits.extend(partial)
        elif fuzzy:
            fuzzy.sort(key=lambda x: -x[0])
            hits.extend(t for _, t in fuzzy)

    return _prefer_bookmark_targets(_dedupe_targets(hits))


def _file_hits(book, q, raw, num_mode=HYMN_NUM_MODE_CONTAINS):
    exact_num = (
        num_mode == HYMN_NUM_MODE_EXACT and str(raw).strip().isdigit()
    )
    file_hits = []
    for fi in book.get('files', []):
        name = fi['name']
        if exact_num:
            if hymn_number_exact_match(raw, name):
                file_hits.append(_file_target(fi, book))
        elif q in name.lower():
            file_hits.append(_file_target(fi, book))
    return file_hits


def resolve_targets(book, num_ref, num_mode=HYMN_NUM_MODE_CONTAINS):
    """Return actionable payloads (same shape as file_list UserRole)."""
    raw = str(num_ref).strip() if num_ref is not None else ''
    if not raw:
        return [_file_target(fi, book) for fi in book.get('files', [])]

    mode = num_mode if num_mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS
    q = raw.lower()
    kind, _val = _parse_ref(num_ref)

    if mode == HYMN_NUM_MODE_EXACT and raw.isdigit():
        bm_hits = _bookmark_hits(book, q, raw, mode)
        if bm_hits:
            return bm_hits
        return _file_hits(book, q, raw, mode)

    if kind is not None and book.get('bookmark') and book.get('toc'):
        for fi in _pdf_files(book):
            found = _find_page_in_toc(book['toc'], kind, raw)
            if found:
                page, title, lvl = found
                return [_bookmark_target(book, fi, title, page, lvl)]

    if book.get('bookmark') and book.get('toc'):
        bm_hits = _bookmark_hits(book, q, raw, mode)
        if bm_hits:
            return bm_hits

    return _file_hits(book, q, raw, mode)


def resolve_open_request(books, book_ref, num_ref, num_mode=HYMN_NUM_MODE_CONTAINS):
    """Resolve across matched books; dedupe by kind+path+page."""
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
    if kind == 'bookmark':
        return f"{book} · {t.get('title')} · P.{t.get('page')}"
    if kind == 'file':
        return f"{book} · {t.get('name', '')}"
    return book
