"""Recent hymns, pins, setlist, session history, settings export/import."""
import json

MAX_RECENT_HYMNS = 5
MAX_SESSION_HISTORY = 50

EXPORT_KEYS = (
    'theme', 'font_size', 'search_mode', 'open_overlay', 'overlay_duration',
    'overlay_mode', 'display_duplicate', 'keyword_instant', 'book_auto_focus_hymn',
    'click_to_open', 'remote_policy', 'remote_port', 'remote_token',
    'operator_mode', 'pinned_books', 'recent_hymns', 'setlist',
    'setlist_index', 'session_history', 'session_cursor', 'show_preview',
    'auto_rescan', 'startup_tray', 'minimize_to_tray', 'last_opened',
)


def _as_str_list(value):
    if not isinstance(value, list):
        return []
    out = []
    for item in value:
        s = str(item or '').strip()
        if s and s not in out:
            out.append(s)
    return out


def _touch_list(items, name, max_len):
    name = str(name or '').strip()
    if not name:
        return items
    out = [name] + [x for x in items if x != name]
    return out[:max_len]


def touch_recent_hymn(settings, book_name, hymn_ref):
    settings = dict(settings or {})
    book_name = str(book_name or '').strip()
    hymn_ref = str(hymn_ref or '').strip()
    if not book_name:
        return settings
    label = f"{book_name} / {hymn_ref}" if hymn_ref else book_name
    recent = _as_str_list(settings.get('recent_hymns'))
    settings['recent_hymns'] = _touch_list(recent, label, MAX_RECENT_HYMNS)
    return settings


def pinned_books(settings):
    return _as_str_list((settings or {}).get('pinned_books'))


def is_book_pinned(settings, book_name):
    return str(book_name or '').strip() in pinned_books(settings)


def pin_book(settings, book_name):
    settings = dict(settings or {})
    name = str(book_name or '').strip()
    if not name:
        return settings
    pins = pinned_books(settings)
    if name not in pins:
        pins.append(name)
    settings['pinned_books'] = pins
    return settings


def unpin_book(settings, book_name):
    settings = dict(settings or {})
    name = str(book_name or '').strip()
    settings['pinned_books'] = [p for p in pinned_books(settings) if p != name]
    return settings


def toggle_pin_book(settings, book_name):
    if is_book_pinned(settings, book_name):
        return unpin_book(settings, book_name)
    return pin_book(settings, book_name)


def parse_setlist(raw):
    """Parse setlist lines or stored list into [{'book': ..., 'num': ...}, ...]."""
    if raw is None:
        return []
    if isinstance(raw, list):
        if raw and isinstance(raw[0], dict):
            entries = []
            for item in raw:
                book = str(item.get('book', item.get('book_ref', ''))).strip()
                num = str(item.get('num', item.get('num_ref', ''))).strip()
                if book or num:
                    entries.append({'book': book, 'num': num})
            return entries
        lines = [str(x) for x in raw]
    else:
        lines = str(raw).splitlines()

    entries = []
    for line in lines:
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        if '|' in line:
            book, num = line.split('|', 1)
        elif '\t' in line:
            book, num = line.split('\t', 1)
        elif ' / ' in line:
            book, num = line.split(' / ', 1)
        elif '/' in line:
            book, num = line.split('/', 1)
        else:
            book, num = line, ''
        book = book.strip()
        num = num.strip()
        if book or num:
            entries.append({'book': book, 'num': num})
    return entries


def format_setlist(entries):
    """Serialize setlist entries to newline-separated text."""
    lines = []
    for entry in parse_setlist(entries):
        book = entry.get('book', '')
        num = entry.get('num', '')
        if num:
            lines.append(f"{book} / {num}")
        else:
            lines.append(book)
    return '\n'.join(lines)


def format_setlist_compact(entries, index=-1):
    entries = parse_setlist(entries)
    total = len(entries)
    if total == 0:
        return '排程：空'
    idx = max(-1, min(total - 1, int(index)))
    if idx < 0:
        return f'排程：{total} 首'
    cur = entries[idx]
    book = cur.get('book', '')
    num = cur.get('num', '')
    label = f"{book} / {num}" if num else book
    return f'排程 {idx + 1}/{total} · {label}'


def setlist_index(settings):
    try:
        return int((settings or {}).get('setlist_index', -1))
    except (TypeError, ValueError):
        return -1


def clamp_setlist_index(settings):
    settings = dict(settings or {})
    entries = parse_setlist(settings.get('setlist'))
    idx = setlist_index(settings)
    if not entries:
        settings['setlist_index'] = -1
    elif idx < 0:
        settings['setlist_index'] = 0
    elif idx >= len(entries):
        settings['setlist_index'] = len(entries) - 1
    return settings


def setlist_entry_at(settings, index=None):
    settings = clamp_setlist(settings)
    entries = parse_setlist(settings.get('setlist'))
    if not entries:
        return None
    idx = setlist_index(settings) if index is None else index
    if idx < 0 or idx >= len(entries):
        return None
    return entries[idx]


def clamp_setlist(settings):
    settings = dict(settings or {})
    settings['setlist'] = parse_setlist(settings.get('setlist'))
    return clamp_setlist_index(settings)


def advance_setlist_index(settings, delta):
    settings = clamp_setlist(settings)
    entries = parse_setlist(settings.get('setlist'))
    if not entries:
        settings['setlist_index'] = -1
        return settings
    idx = setlist_index(settings)
    if idx < 0:
        idx = 0
    else:
        idx = max(0, min(len(entries) - 1, idx + int(delta)))
    settings['setlist_index'] = idx
    return settings


def session_history_list(settings):
    raw = (settings or {}).get('session_history')
    if not isinstance(raw, list):
        return []
    out = []
    for item in raw:
        if isinstance(item, dict):
            out.append(item)
    return out[:MAX_SESSION_HISTORY]


def push_session_history(settings, entry):
    settings = dict(settings or {})
    if not isinstance(entry, dict):
        return settings
    history = session_history_list(settings)
    key = (
        entry.get('book', ''),
        entry.get('num', ''),
        entry.get('label', ''),
    )
    history = [h for h in history if (
        h.get('book', ''), h.get('num', ''), h.get('label', '')
    ) != key]
    history.insert(0, entry)
    settings['session_history'] = history[:MAX_SESSION_HISTORY]
    return settings


def navigate_session_history(settings, delta):
    """Move session cursor by delta; returns (settings, entry or None)."""
    settings = dict(settings or {})
    history = session_history_list(settings)
    if not history:
        settings['session_cursor'] = -1
        return settings, None
    try:
        cursor = int(settings.get('session_cursor', 0))
    except (TypeError, ValueError):
        cursor = 0
    cursor = max(0, min(len(history) - 1, cursor + int(delta)))
    settings['session_cursor'] = cursor
    return settings, history[cursor]


def export_settings_subset(settings):
    settings = dict(settings or {})
    return {key: settings.get(key) for key in EXPORT_KEYS if key in settings}


def import_settings_subset(settings, data):
    settings = dict(settings or {})
    if not isinstance(data, dict):
        return settings
    for key in EXPORT_KEYS:
        if key not in data:
            continue
        if key == 'setlist':
            settings[key] = parse_setlist(data[key])
        elif key in ('pinned_books', 'recent_hymns', 'session_history'):
            val = data[key]
            settings[key] = val if isinstance(val, list) else _as_str_list(val)
        elif key == 'setlist_index':
            try:
                settings[key] = int(data[key])
            except (TypeError, ValueError):
                settings[key] = -1
        elif key in ('operator_mode', 'show_preview', 'auto_rescan', 'startup_tray', 'minimize_to_tray',
                     'keyword_instant', 'book_auto_focus_hymn', 'click_to_open', 'display_duplicate',
                     'open_overlay'):
            settings[key] = bool(data[key])
        elif key == 'font_size':
            try:
                settings[key] = max(10, min(18, int(data[key])))
            except (TypeError, ValueError):
                pass
        elif key == 'remote_port':
            try:
                settings[key] = max(1024, min(65535, int(data[key])))
            except (TypeError, ValueError):
                pass
        else:
            settings[key] = data[key]
    return clamp_setlist(settings)


def export_settings_json(settings, indent=2):
    return json.dumps(export_settings_subset(settings), ensure_ascii=False, indent=indent)


def import_settings_json(settings, text):
    data = json.loads(text)
    return import_settings_subset(settings, data)


def clear_setlist(settings):
    settings = dict(settings or {})
    settings['setlist'] = []
    settings['setlist_index'] = -1
    return settings


def add_setlist_entry(settings, book, num):
    settings = dict(settings or {})
    book = str(book or '').strip()
    num = str(num or '').strip()
    if not book and not num:
        return settings
    entries = parse_setlist(settings.get('setlist'))
    entries.append({'book': book, 'num': num})
    settings['setlist'] = entries
    if setlist_index(settings) < 0 and entries:
        settings['setlist_index'] = 0
    return settings


def remove_setlist_at(settings, index):
    settings = dict(settings or {})
    entries = parse_setlist(settings.get('setlist'))
    if not (0 <= index < len(entries)):
        return settings
    entries.pop(index)
    settings['setlist'] = entries
    idx = setlist_index(settings)
    if not entries:
        settings['setlist_index'] = -1
    elif idx >= len(entries):
        settings['setlist_index'] = len(entries) - 1
    return settings


def move_setlist_entry(settings, index, delta):
    settings = dict(settings or {})
    entries = parse_setlist(settings.get('setlist'))
    j = index + int(delta)
    if not (0 <= index < len(entries) and 0 <= j < len(entries)):
        return settings
    entries[index], entries[j] = entries[j], entries[index]
    settings['setlist'] = entries
    settings['setlist_index'] = j
    return settings


def record_open(settings, book_name='', hymn_ref='', label=''):
    settings = touch_recent_hymn(settings, book_name, hymn_ref)
    if label:
        settings['last_opened'] = label
    elif book_name and hymn_ref:
        settings['last_opened'] = f"{book_name} / {hymn_ref}"
    elif book_name:
        settings['last_opened'] = book_name
    entry = {
        'book': book_name,
        'num': hymn_ref,
        'label': settings.get('last_opened', ''),
    }
    settings = push_session_history(settings, entry)
    settings['session_cursor'] = 0
    return settings
