"""One-line preview snippets for file / bookmark payloads."""
import os
import re

try:
    import fitz
except ImportError:
    fitz = None

try:
    from docx import Document
except ImportError:
    Document = None


def _collapse(text, max_len=200):
    flat = re.sub(r'\s+', ' ', (text or '').replace('\r', ' ').replace('\n', ' ')).strip()
    if len(flat) <= max_len:
        return flat
    return flat[: max_len - 1].rstrip() + '…'


def _pdf_snippet(path, max_len=200):
    if not fitz or not path or not os.path.isfile(path):
        return ''
    try:
        doc = fitz.open(path)
        parts = []
        for i in range(min(2, doc.page_count)):
            parts.append(doc[i].get_text())
        doc.close()
        return _collapse(' '.join(parts), max_len)
    except Exception:
        return ''


def _docx_snippet(path, max_len=200):
    if not Document or not path or not os.path.isfile(path):
        return ''
    try:
        doc = Document(path)
        parts = []
        for para in doc.paragraphs[:8]:
            t = para.text.strip()
            if t:
                parts.append(t)
            if sum(len(p) for p in parts) >= max_len:
                break
        return _collapse(' '.join(parts), max_len)
    except Exception:
        return ''


def preview_for_payload(payload, max_len=200):
    """
    Return a short preview string for a hymn open payload or content-search hit.
    Supports file, bookmark, content_pdf, content_word kinds.
    """
    if not payload:
        return ''

    kind = payload.get('kind')

    if kind == 'bookmark':
        title = payload.get('title', '')
        page = payload.get('page')
        if page:
            return _collapse(f"{title} · P.{page}", max_len)
        return _collapse(title, max_len)

    if kind in ('content_pdf', 'content_word'):
        snip = payload.get('snippet') or ''
        page = payload.get('page')
        if page and snip:
            return _collapse(f"P.{page} · {snip}", max_len)
        if snip:
            return _collapse(snip, max_len)
        if page:
            return f"P.{page}"

    path = payload.get('path') or payload.get('pdf_path') or ''
    if not path:
        return _collapse(payload.get('name', ''), max_len)

    name = payload.get('name') or os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()

    if ext == '.pdf':
        body = _pdf_snippet(path, max_len)
        return _collapse(f"{name} · {body}" if body else name, max_len)
    if ext in ('.docx', '.doc'):
        body = _docx_snippet(path, max_len)
        return _collapse(f"{name} · {body}" if body else name, max_len)

    return _collapse(name, max_len)


def preview_lines(text, max_lines=2, max_chars=120):
    """Split preview into at most max_lines for compact UI display."""
    text = (text or '').strip()
    if not text:
        return []
    if len(text) <= max_chars:
        return [text]
    chunk = max_chars // max_lines if max_lines > 1 else max_chars
    lines = []
    rest = text
    while rest and len(lines) < max_lines:
        if len(rest) <= chunk:
            lines.append(rest)
            break
        cut = rest.rfind(' ', 0, chunk)
        if cut < chunk // 2:
            cut = chunk
        lines.append(rest[:cut].rstrip() + '…')
        rest = rest[cut:].lstrip()
    return lines
