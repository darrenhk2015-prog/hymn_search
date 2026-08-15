"""Resolve viewer iframe URL — PDF → church site, Word → Google Drive.

S1/S2 合訂本（含 PDF 書籤）→ Google Drive；其餘 PDF → 教會網站。
"""
from __future__ import annotations

import os
import re

from hymn_remote.gdrive_hymn_map import _is_s_book, resolve_gdrive_hymn
from hymn_remote.web_hymn_map import resolve_web_hymn

_BOOKMARK_LABEL = re.compile(r'·\s*P\.\d+', re.IGNORECASE)
_PDF_BOOK = re.compile(r'合訂本.*\.pdf$|^\d{1,2}\s*神家詩歌\.pdf$', re.IGNORECASE)


def _path_ext(path: str) -> str:
    return os.path.splitext(path or '')[1].lower()


def _looks_like_pdf_open(
    payload: dict | None = None,
    *,
    book: str = '',
    num: str = '',
    label: str = '',
    filename: str = '',
) -> bool:
    if payload:
        kind = payload.get('kind')
        if kind == 'bookmark':
            return True
        if kind == 'content_pdf':
            return True
        if kind == 'file':
            path = payload.get('path') or ''
            name = payload.get('name') or os.path.basename(path)
            return _path_ext(path or name) == '.pdf'

    for part in (filename, num, label):
        if _path_ext(part) == '.pdf':
            return True

    if _BOOKMARK_LABEL.search(label or ''):
        return True

    book_name = (book or '').strip()
    if book_name and _PDF_BOOK.search(book_name):
        return True

    return False


def _parse_label_parts(label: str) -> tuple[str, str]:
    label = (label or '').strip()
    if not label:
        return '', ''
    parts = [p.strip() for p in label.split('·')]
    if len(parts) >= 2:
        return parts[0], parts[1]
    if ' / ' in label:
        book, num = label.split(' / ', 1)
        return book.strip(), num.strip()
    return '', label


def resolve_viewer_hymn(
    book: str = '',
    num: str = '',
    label: str = '',
    filename: str = '',
    payload: dict | None = None,
) -> dict:
    if not book and not num and label:
        parsed_book, parsed_num = _parse_label_parts(label)
        book = parsed_book or book
        num = parsed_num or num
    use_web = _looks_like_pdf_open(
        payload, book=book, num=num, label=label, filename=filename,
    )
    prefer_gdrive = _is_s_book(book)
    if use_web and not prefer_gdrive:
        info = resolve_web_hymn(book, num, label)
        info['viewer_source'] = 'web'
        info.setdefault('drive_file_id', '')
        info.setdefault('drive_name', '')
        info.setdefault('drive_rel', '')
        info.setdefault('drive_folder_id', '')
        return info

    info = resolve_gdrive_hymn(book, num, label, filename=filename)
    if prefer_gdrive and not info.get('mapped') and use_web:
        info = resolve_web_hymn(book, num, label)
        info['viewer_source'] = 'web'
        info.setdefault('drive_file_id', '')
        info.setdefault('drive_name', '')
        info.setdefault('drive_rel', '')
        info.setdefault('drive_folder_id', '')
        return info
    info['viewer_source'] = 'gdrive'
    return info
