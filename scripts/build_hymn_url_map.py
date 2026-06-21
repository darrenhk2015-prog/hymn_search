"""Build hymn code -> churchofgod.org.hk URL mapping from WordPress sitemaps.

Usage:
  python scripts/build_hymn_url_map.py

Output: hymn_remote/data/cog_hymn_urls.json
Used when viewer opens PDF (maps to churchofgod.org.hk/hymns/...).
Rebuild exe after updating (build_exe.ps1 bundles hymn_remote/data).
"""
import json
import re
import sys
import time
import urllib.error
import urllib.request
from html import unescape
from pathlib import Path

BASE = 'https://churchofgod.org.hk'
OUT = Path(__file__).resolve().parent.parent / 'hymn_remote' / 'data' / 'cog_hymn_urls.json'
USER_AGENT = 'HymnSearch/1.0 (+local mapping build)'

_H_CODE = re.compile(r'\b([Hh])(\d{1,2})\s*[-_]\s*(\d{1,2})\b')
_S_CODE = re.compile(r'\b([Ss])(\d{1,2})\s*[-_]\s*(\d{1,3})\b')


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode('utf-8', errors='replace')


def normalize_h_code(vol: int, num: int) -> str:
    return f'h{vol}-{num:02d}'


def normalize_s_code(vol: int, num: int) -> str:
    return f's{vol}-{num}'


def codes_from_text(text: str) -> set[str]:
    out: set[str] = set()
    for m in _H_CODE.finditer(text or ''):
        out.add(normalize_h_code(int(m.group(2)), int(m.group(3))))
        out.add(f'h{int(m.group(2))}-{int(m.group(3))}')
    for m in _S_CODE.finditer(text or ''):
        out.add(normalize_s_code(int(m.group(2)), int(m.group(3))))
    return out


def slug_from_url(url: str) -> str:
    path = url.split('/hymns/', 1)[-1].strip('/').split('/')[0].lower()
    m = re.match(r'([hs]\d+-\d+)', path)
    return m.group(1) if m else ''


def title_from_url(url: str) -> str:
    slug = slug_from_url(url)
    if not slug:
        return ''
    rest = url.split(f'/hymns/{slug}', 1)[-1].strip('/-')
    if not rest:
        return slug.upper()
    from urllib.parse import unquote
    return unquote(rest).replace('-', ' ')


def fetch_sitemap_urls() -> list[str]:
    index_xml = fetch(f'{BASE}/wp-sitemap.xml')
    sitemaps = re.findall(r'<loc>(https://churchofgod\.org\.hk/hymns-sitemap\d+\.xml)</loc>', index_xml)
    if not sitemaps:
        sitemaps = [f'{BASE}/wp-sitemap-posts-hymns-1.xml']
    urls: set[str] = set()
    for sm in sitemaps:
        try:
            xml = fetch(sm)
        except (urllib.error.URLError, TimeoutError) as e:
            print(f'skip {sm}: {e}', file=sys.stderr)
            continue
        for u in re.findall(r'<loc>(https://churchofgod\.org\.hk/hymns/[^<]+)</loc>', xml):
            if '/feed/' in u or '/page/' in u:
                continue
            urls.add(u.rstrip('/') + '/')
        time.sleep(0.1)
    return sorted(urls)


def build_map(urls: list[str]) -> dict[str, dict]:
    by_code: dict[str, dict] = {}
    for url in urls:
        slug = slug_from_url(url)
        title = title_from_url(url)
        codes = codes_from_text(slug)
        codes |= codes_from_text(url)
        if slug:
            codes.add(slug)
        entry = {'url': url, 'slug': slug, 'title': title}
        for code in codes:
            by_code[code] = entry
    return by_code


def main():
    print('Fetching hymn sitemaps…', file=sys.stderr)
    urls = fetch_sitemap_urls()
    print(f'Found {len(urls)} hymn URLs', file=sys.stderr)
    mapping = build_map(urls)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        'version': 1,
        'source': f'{BASE}/wp-sitemap.xml',
        'built_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'count': len(mapping),
        'by_code': mapping,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Wrote {len(mapping)} codes -> {OUT}')


if __name__ == '__main__':
    main()
