"""Fuzzy title matching: subsequence + optional pinyin initials (pypinyin)."""
import re


def normalize(text):
    return re.sub(r'\s+', '', (text or '').strip().lower())


def is_subsequence(needle, haystack):
    """Return True if all chars in needle appear in order within haystack."""
    if not needle:
        return True
    if not haystack:
        return False
    i = 0
    for ch in haystack:
        if ch == needle[i]:
            i += 1
            if i >= len(needle):
                return True
    return False


def pinyin_initials(text):
    """First-letter pinyin initials; empty string if pypinyin unavailable."""
    if not text:
        return ''
    try:
        from pypinyin import lazy_pinyin, Style
        parts = lazy_pinyin(text, style=Style.FIRST_LETTER, errors='ignore')
        return ''.join(p for p in parts if p).lower()
    except ImportError:
        return ''


def fuzzy_match_score(query, title):
    """
    Score a title against query. Higher is better; 0 means no match.
    Checks normalized subsequence, raw subsequence, and pinyin-initial subsequence.
    """
    q = normalize(query)
    if not q:
        return 0

    title_norm = normalize(title)
    title_raw = (title or '').strip().lower()

    if q == title_norm or q == title_raw:
        return 1000
    if title_norm.startswith(q) or title_raw.startswith(q):
        return 900 + len(q)
    if q in title_norm or q in title_raw:
        return 800 + len(q)

    if is_subsequence(q, title_norm):
        return 500 + len(q)
    if title_raw and is_subsequence(q, title_raw):
        return 450 + len(q)

    initials = pinyin_initials(title or '')
    if initials:
        if q == initials:
            return 700
        if initials.startswith(q):
            return 650 + len(q)
        if is_subsequence(q, initials):
            return 400 + len(q)

    return 0


def fuzzy_match(query, title, min_score=1):
    return fuzzy_match_score(query, title) >= min_score


def rank_titles(query, titles):
    """Return (title, score) pairs sorted by score descending."""
    scored = []
    for title in titles:
        score = fuzzy_match_score(query, title)
        if score > 0:
            scored.append((title, score))
    scored.sort(key=lambda pair: (-pair[1], pair[0]))
    return scored
