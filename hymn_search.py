#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
詩歌冊搜索系統（Hymn Search）— Phase 1
=====================================

【程式簡介】
    本程式為 PyQt6 桌面應用，用於在本地詩歌資料夾中快速搜尋、瀏覽及開啟
    Word / PDF 詩歌檔案。支援三種搜尋模式（書本 / 全局 / 關鍵字）、PDF 書籤
    跳頁、全文索引、開啟檔案時的置頂書名提示、多螢幕延伸顯示，以及可選的
    同步畫面切換（Windows）。

【執行環境】
    - Python 3.10+
    - 依賴：PyQt6、PyMuPDF (fitz)、python-docx
    - 建議：Windows 10/11（Word 聚焦、顯示切換、文件監聽為 Windows 專用）

【資料來源】
    預設掃描 exe 旁 `神家詩歌集`；所有 UI 設定存於 settings.json，啟動時還原。
    每個子資料夾或單一檔案視為一本「詩歌冊」，可含 PDF、DOC/DOCX、ODT。

【主要模組結構】
    CONFIG / 資源路徑    — 資料夾路徑、圖示載入（支援 PyInstaller 打包）
    THEMES               — 深/淺色主題與 Qt 樣式表
    HELPERS              — 通用 UI 元件
    SCANNER              — 掃描書冊、解析 PDF 目錄（TOC）
    CONTENT SEARCH       — 即時全文搜索與持久化索引
    OPENERS              — 開啟 PDF/Word、顯示模式切換
    BookNameOverlay      — 全螢幕置頂書名浮層（多螢幕）
    DropdownList         — 書名輸入下拉建議
    MainWindow           — 主視窗與所有 UI 邏輯
    ENTRY POINT          — 應用程式進入點

【打包為 exe】
    詳見 docs/BUILD.md；執行 build_exe.ps1 產生 dist/HymnSearch.exe

【相關文檔】
    docs/SYSTEM_SPEC.md  — 系統規格
    docs/DESIGN.md       — 設計說明
    docs/USER_GUIDE.md   — 使用手冊
"""

import os, sys, re, subprocess, zipfile, json, gzip, time, threading
import fitz  # PyMuPDF

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton, QToolButton,
    QLineEdit, QLabel, QListWidget, QListWidgetItem,
    QComboBox, QSlider, QFrame, QScrollArea, QCheckBox, QSpinBox,
    QGroupBox, QFileDialog, QTabWidget, QPlainTextEdit, QMenu, QDialog,
    QSystemTrayIcon, QColorDialog,
)
from PyQt6.QtCore import (
    Qt, QTimer, QPoint, pyqtSignal, pyqtSlot, QEvent, QThread, QRectF, QRect,
    QPropertyAnimation, QEasingCurve, QFileSystemWatcher,
)
from PyQt6.QtGui import (
    QColor, QPalette, QFont, QPainter, QPainterPath, QIcon, QShortcut, QKeySequence,
    QAction, QImage, QPixmap,
)

from hymn_features.mixin import EnhancementMixin
from hymn_features.fuzzy import fuzzy_match_score
from hymn_features.preview import preview_for_payload, preview_lines
from hymn_features.session_store import (
    add_setlist_entry, advance_setlist_index, clear_setlist, clamp_setlist, move_setlist_entry,
    parse_setlist, remove_setlist_at, session_history_list, setlist_index,
)

from hymn_remote.server import RemoteServer
from hymn_remote.tunnel import (
    CloudflareTunnel,
    cloudflared_service_install_command,
    cloudflared_service_restart_command,
    cloudflared_service_uninstall_command,
    install_cloudflared_service,
    is_cloudflared_service_installed,
    normalize_tunnel_url,
    shared_tunnel_setup_hint,
    tunnel_url_path_prefix,
    tunnel_viewer_url,
)
from hymn_remote.state import (
    DEFAULT_PORT, OPEN_POLICY_AUTO, OPEN_POLICY_CONFIRM, OPEN_POLICY_UI,
)
from hymn_remote.resolver import (
    resolve_books, resolve_targets, resolve_open_request,
    list_entries_for_book, summarize_match,
)

try:
    import qrcode as _qrcode_bundle  # noqa: F401 — PyInstaller bundle
    from PIL import Image as _pil_image_bundle  # noqa: F401
except ImportError:
    pass

try:
    from version import APP_VERSION, BUILD_DATE
except ImportError:
    APP_VERSION = 0
    BUILD_DATE = ''

# ══════════════════════════════════════════════════════════════
#  CONFIG — 全域設定與資源路徑
# ══════════════════════════════════════════════════════════════
# default_hymn_folder：未設定時的預設詩歌根目錄（exe / 專案目錄下的 神家詩歌集）。
# settings.json：儲存所有 UI 設定；缺省時用 default_settings()。
# app_dir：exe 所在目錄（打包）或專案目錄（開發）；settings 與預設資料夾放這裡。
# app_resource_path：打包後讀 _MEIPASS 內嵌資源（icons），與詩歌資料夾分開。
def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def default_hymn_folder():
    return os.path.join(app_dir(), '神家詩歌集')


def settings_file_path():
    return os.path.join(app_dir(), 'settings.json')


def app_resource_path(*parts):
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


def load_app_icon():
    """載入視窗與工作列圖示（優先 app.ico，其次 app.png）。"""
    for name in ('app.ico', 'app.png'):
        path = app_resource_path('assets', name)
        if os.path.isfile(path):
            icon = QIcon(path)
            if not icon.isNull():
                return icon
    return QIcon()

# ══════════════════════════════════════════════════════════════
#  THEMES — 深/淺色主題與 Qt 樣式
# ══════════════════════════════════════════════════════════════
# THEMES 字典定義各 UI 色票；build_*_style 產生 QListWidget、輸入框等元件的 stylesheet。
THEMES = {
    'dark': {
        'bg':       '#0c0e1a', 'panel':   '#13162a', 'panel2':  '#1b1f36',
        'border':   '#252840', 'text':    '#e4e8f5', 'text2':   '#8892b8',
        'muted':    '#505870', 'accent':  '#7c6af7', 'accent_h':'#9d8fff',
        'accent_d': '#5a4ed0', 'a_text':  '#ffffff', 'a_sub':   '#1e1a40',
        'in_bg':    '#191c30', 'in_bd':   '#2e3254',
        'ok':       '#34d399', 'ok_bg':   '#012d1a',
        'warn':     '#fbbf24', 'warn_bg': '#1c1200',
        'err':      '#f87171', 'err_bg':  '#220606',
        'w_c':      '#60a5fa', 'w_bg':    '#0d1e38', 'w_bd':    '#1e3f80',
        'p_c':      '#f87171', 'p_bg':    '#260808', 'p_bd':    '#8b1818',
    },
    'light': {
        'bg':       '#f0f2ff', 'panel':   '#ffffff', 'panel2':  '#edf0fc',
        'border':   '#dde0f2', 'text':    '#1a1e40', 'text2':   '#4a5480',
        'muted':    '#8292b8', 'accent':  '#5b4ec4', 'accent_h':'#7c6af7',
        'accent_d': '#4338ca', 'a_text':  '#ffffff', 'a_sub':   '#eeeaff',
        'in_bg':    '#ffffff', 'in_bd':   '#c8ceec',
        'ok':       '#16a34a', 'ok_bg':   '#dcfce7',
        'warn':     '#d97706', 'warn_bg': '#fef3c7',
        'err':      '#dc2626', 'err_bg':  '#fee2e2',
        'w_c':      '#1d4ed8', 'w_bg':    '#eff6ff', 'w_bd':    '#bfdbfe',
        'p_c':      '#b91c1c', 'p_bg':    '#fff1f2', 'p_bd':    '#fecaca',
    },
}


def build_app_style(theme, fs):
    t = THEMES[theme]
    return f"""
* {{ font-family: 'Segoe UI', 'PingFang TC', 'Microsoft YaHei', sans-serif; }}
QMainWindow, QWidget {{ background: {t['bg']}; color: {t['text']}; }}
QLabel {{ color: {t['text']}; background: transparent; font-size: {fs}px; }}
QLineEdit {{
    background: {t['in_bg']}; border: 1.5px solid {t['in_bd']};
    border-radius: 10px; padding: 8px 14px; color: {t['text']};
    font-size: {fs}px; selection-background-color: {t['accent']};
}}
QLineEdit:focus {{ border-color: {t['accent']}; background: {t['panel2']}; }}
QLineEdit:hover:!focus {{ border-color: {t['accent_h']}; }}
QComboBox {{
    background: {t['in_bg']}; border: 1.5px solid {t['in_bd']};
    border-radius: 8px; padding: 5px 10px; color: {t['text']};
    font-size: {fs - 1}px; min-width: 78px;
}}
QComboBox:focus {{ border-color: {t['accent']}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {t['panel']}; color: {t['text']};
    border: 1.5px solid {t['border']}; outline: none;
    selection-background-color: {t['accent']};
    selection-color: {t['a_text']};
}}
QSlider::groove:horizontal {{ height: 4px; background: {t['border']}; border-radius: 2px; }}
QSlider::sub-page:horizontal {{ background: {t['accent']}; border-radius: 2px; }}
QSlider::handle:horizontal {{
    background: {t['accent']}; border: 2px solid {t['bg']};
    width: 14px; height: 14px; margin: -6px 0; border-radius: 7px;
}}
QSlider::handle:horizontal:hover {{ background: {t['accent_h']}; }}
QScrollArea {{ border: none; background: transparent; }}
QScrollBar:vertical {{ background: transparent; width: 5px; border: none; }}
QScrollBar::handle:vertical {{
    background: {t['border']}; border-radius: 2px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {t['accent']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ height: 0; }}
QSplitter::handle {{ background: {t['border']}; }}
QCheckBox {{ color: {t['text2']}; font-size: {fs - 1}px; spacing: 6px; }}
QCheckBox::indicator {{
    width: 16px; height: 16px; border-radius: 4px;
    border: 1.5px solid {t['in_bd']}; background: {t['in_bg']};
}}
QCheckBox::indicator:checked {{
    background: {t['accent']}; border-color: {t['accent']};
}}
"""


def build_book_list_style(theme, fs):
    t = THEMES[theme]
    return f"""
QListWidget {{ background: transparent; border: none; outline: none; font-size: {fs}px; }}
QListWidget::item {{
    padding: 9px 10px; color: {t['text2']};
    border-radius: 8px; margin: 1px 3px;
}}
QListWidget::item:hover {{ background: {t['panel2']}; color: {t['text']}; }}
QListWidget::item:selected {{ background: {t['accent']}; color: {t['a_text']}; }}
"""


def build_file_list_style(theme, fs):
    t = THEMES[theme]
    return f"""
QListWidget {{ background: transparent; border: none; outline: none; font-size: {fs}px; }}
QListWidget::item {{
    padding: 8px 10px; color: {t['text']};
    border-radius: 8px; margin: 1px 3px;
}}
QListWidget::item:hover {{ background: {t['panel2']}; }}
QListWidget::item:selected {{ background: {t['a_sub']}; color: {t['accent']}; }}
"""


# ══════════════════════════════════════════════════════════════
#  HELPERS — 通用 UI 小工具
# ══════════════════════════════════════════════════════════════
class HSep(QFrame):
    """水平分隔線；用於側欄、內容區視覺分區。"""
    def __init__(self, theme='dark', parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setFixedHeight(1)
        self.set_theme(theme)

    def set_theme(self, theme):
        self.setStyleSheet(
            f"background-color: {THEMES[theme]['border']}; border: none; margin: 0; padding: 0;"
        )


# ══════════════════════════════════════════════════════════════
#  SCANNER — 掃描詩歌資料夾與 PDF 書籤
# ══════════════════════════════════════════════════════════════
# scan_books：遍歷 HYMN_FOLDER，建立書冊清單（含檔案列表、PDF 頁數、TOC 書籤）。
# find_page_in_toc：依詩歌編號在 PDF 目錄中查找對應頁碼。
def scan_books(folder):
    books = []
    if not os.path.isdir(folder):
        return books

    def build_entry(entry_path, entry_name):
        pdf_path = word_path = None
        files = []

        if os.path.isfile(entry_path):
            if not os.path.basename(entry_path).startswith('~$'):
                ext = os.path.splitext(entry_path)[1].lower()
                files.append({'name': os.path.basename(entry_path), 'path': entry_path, 'ext': ext})
                if ext == '.pdf': pdf_path = entry_path
                if ext in ('.doc', '.docx', '.odt'): word_path = entry_path
        else:
            for f in sorted(os.listdir(entry_path)):
                if f.startswith('~$'): continue
                fp = os.path.join(entry_path, f)
                if os.path.isdir(fp): continue
                ext = os.path.splitext(f)[1].lower()
                files.append({'name': f, 'path': fp, 'ext': ext})
                if ext == '.pdf' and pdf_path is None: pdf_path = fp
                if ext in ('.doc', '.docx', '.odt') and word_path is None: word_path = fp

        toc_list = []
        total_pages = 0
        if pdf_path:
            try:
                doc = fitz.open(pdf_path)
                total_pages = doc.page_count
                toc = doc.get_toc()
                doc.close()
                if len(toc) > 5:
                    for lvl, title, page in toc:
                        if lvl in (1, 2):
                            toc_list.append((lvl, title.strip(), page))
            except Exception as e:
                print(f"[WARN] {pdf_path}: {e}")

        lvl2 = [e for e in toc_list if e[0] == 2]
        is_bm = len(toc_list) > 5
        books.append({
            'name': entry_name, 'pdf': pdf_path, 'docx': word_path,
            'toc': toc_list, 'bookmark': is_bm,
            'pages': total_pages,
            'hymn_count': (len(lvl2) if lvl2 else len(toc_list)) if is_bm else total_pages,
            'files': files,
        })

    for name in sorted(os.listdir(folder)):
        fp = os.path.join(folder, name)
        if os.path.isdir(fp):
            build_entry(fp, name)

    for f in sorted(os.listdir(folder)):
        fp = os.path.join(folder, f)
        if os.path.isdir(fp) or f.startswith('~$'): continue
        ext = os.path.splitext(f)[1].lower()
        if ext not in ('.pdf', '.doc', '.docx', '.odt'): continue
        build_entry(fp, os.path.splitext(f)[0])

    return books


def find_page_in_toc(toc_list, hymn_num, raw=''):
    """依詩歌編號在 PDF 目錄中查找對應頁碼（支援 H22-04 格式）。"""
    from hymn_remote.resolver import _find_page_in_toc
    found = _find_page_in_toc(toc_list, hymn_num, raw)
    return found[0] if found else None


# ══════════════════════════════════════════════════════════════
#  CONTENT SEARCH — 全文搜索（即時掃描與索引）
# ══════════════════════════════════════════════════════════════
# 支援 PDF / DOCX / DOC / ODT；search_*_content 為單檔搜索，search_books_content 為全庫掃描。
SEARCHABLE_EXTS = {'.pdf', '.docx', '.doc', '.odt'}


def _text_snippet(text, query, radius=40):
    q = query.lower()
    flat = re.sub(r'\s+', ' ', text.replace('\r', ' ').replace('\n', ' ')).strip()
    idx = flat.lower().find(q)
    if idx < 0:
        return ''
    start = max(0, idx - radius)
    end = min(len(flat), idx + len(query) + radius)
    snip = flat[start:end]
    if start > 0:
        snip = '…' + snip
    if end < len(flat):
        snip = snip + '…'
    return snip


def search_pdf_content(path, query, max_hits=15):
    hits = []
    ql = query.lower()
    try:
        doc = fitz.open(path)
        for i in range(doc.page_count):
            text = doc[i].get_text()
            if ql in text.lower():
                hits.append({'page': i + 1, 'snippet': _text_snippet(text, query)})
                if len(hits) >= max_hits:
                    break
        doc.close()
    except Exception as e:
        print(f"[WARN] pdf search {path}: {e}")
    return hits


def search_docx_content(path, query, max_hits=15):
    try:
        from docx import Document
    except ImportError:
        return None
    hits = []
    ql = query.lower()
    try:
        doc = Document(path)
        for idx, para in enumerate(doc.paragraphs):
            text = para.text
            if text and ql in text.lower():
                hits.append({'page': idx + 1, 'snippet': _text_snippet(text, query)})
                if len(hits) >= max_hits:
                    break
    except Exception as e:
        print(f"[WARN] docx search {path}: {e}")
    return hits


def search_odt_content(path, query, max_hits=15):
    hits = []
    ql = query.lower()
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read('content.xml').decode('utf-8', errors='ignore')
        blocks = re.split(r'</text:p>', xml)
        for idx, block in enumerate(blocks):
            text = re.sub(r'<[^>]+>', ' ', block)
            text = re.sub(r'\s+', ' ', text).strip()
            if text and ql in text.lower():
                hits.append({'page': idx + 1, 'snippet': _text_snippet(text, query)})
                if len(hits) >= max_hits:
                    break
    except Exception as e:
        print(f"[WARN] odt search {path}: {e}")
    return hits


def search_doc_content(path, query, max_hits=15):
    if sys.platform != 'win32':
        return []
    try:
        import win32com.client
    except ImportError:
        return []
    hits = []
    ql = query.lower()
    word = None
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        for i, para in enumerate(doc.Paragraphs):
            text = para.Range.Text
            if text and ql in text.lower():
                hits.append({'page': i + 1, 'snippet': _text_snippet(text, query)})
                if len(hits) >= max_hits:
                    break
        doc.Close(False)
    except Exception as e:
        print(f"[WARN] doc search {path}: {e}")
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
    return hits


def search_file_content(path, ext, query, max_hits=15):
    if ext == '.pdf':
        return search_pdf_content(path, query, max_hits)
    if ext == '.docx':
        return search_docx_content(path, query, max_hits)
    if ext == '.odt':
        return search_odt_content(path, query, max_hits)
    if ext == '.doc':
        return search_doc_content(path, query, max_hits)
    return []


def search_books_content(books, query, progress_cb=None, cancel_check=None, max_hits_per_file=15):
    q = query.strip()
    if not q:
        return []
    results = []
    for book in books:
        if cancel_check and cancel_check():
            return results
        for fi in book.get('files', []):
            if fi['ext'] not in SEARCHABLE_EXTS:
                continue
            if cancel_check and cancel_check():
                return results
            if progress_cb:
                progress_cb(f"{book['name']} / {fi['name']}")
            hits = search_file_content(fi['path'], fi['ext'], q, max_hits_per_file)
            if hits is None:
                hits = []
            if hits:
                results.append({
                    'book': book['name'],
                    'file': fi['name'],
                    'path': fi['path'],
                    'ext': fi['ext'],
                    'hits': hits,
                })
    return results


# ── 內容索引（持久化快取）────────────────────────────────────
# 索引儲存於 HYMN_FOLDER/content_index.json.gz，依檔案 mtime 判斷是否過期。
INDEX_VERSION = 1


def index_cache_path(folder):
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, 'content_index.json.gz')


def extract_pdf_chunks(path):
    chunks = []
    try:
        doc = fitz.open(path)
        for i in range(doc.page_count):
            text = doc[i].get_text()
            if text.strip():
                chunks.append({'page': i + 1, 'text': text})
        doc.close()
    except Exception as e:
        print(f"[WARN] index pdf {path}: {e}")
    return chunks


def extract_docx_chunks(path):
    try:
        from docx import Document
    except ImportError:
        return None
    chunks = []
    try:
        doc = Document(path)
        for idx, para in enumerate(doc.paragraphs):
            text = para.text
            if text.strip():
                chunks.append({'page': idx + 1, 'text': text})
    except Exception as e:
        print(f"[WARN] index docx {path}: {e}")
    return chunks


def extract_odt_chunks(path):
    chunks = []
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read('content.xml').decode('utf-8', errors='ignore')
        blocks = re.split(r'</text:p>', xml)
        for idx, block in enumerate(blocks):
            text = re.sub(r'<[^>]+>', ' ', block)
            text = re.sub(r'\s+', ' ', text).strip()
            if text:
                chunks.append({'page': idx + 1, 'text': text})
    except Exception as e:
        print(f"[WARN] index odt {path}: {e}")
    return chunks


def extract_doc_chunks(path):
    if sys.platform != 'win32':
        return []
    try:
        import win32com.client
    except ImportError:
        return []
    chunks = []
    word = None
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        for i, para in enumerate(doc.Paragraphs):
            text = para.Range.Text
            if text and text.strip():
                chunks.append({'page': i + 1, 'text': text})
        doc.Close(False)
    except Exception as e:
        print(f"[WARN] index doc {path}: {e}")
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
    return chunks


def extract_file_chunks(path, ext):
    if ext == '.pdf':
        return extract_pdf_chunks(path)
    if ext == '.docx':
        return extract_docx_chunks(path)
    if ext == '.odt':
        return extract_odt_chunks(path)
    if ext == '.doc':
        return extract_doc_chunks(path)
    return []


def build_content_index(books, folder, progress_cb=None, cancel_check=None):
    files_index = {}
    for book in books:
        if cancel_check and cancel_check():
            return None
        for fi in book.get('files', []):
            if fi['ext'] not in SEARCHABLE_EXTS:
                continue
            if cancel_check and cancel_check():
                return None
            path = fi['path']
            if progress_cb:
                progress_cb(f"{book['name']} / {fi['name']}")
            chunks = extract_file_chunks(path, fi['ext'])
            if chunks is None:
                chunks = []
            try:
                mtime = os.path.getmtime(path)
            except OSError:
                mtime = 0
            files_index[path] = {
                'book': book['name'],
                'file': fi['name'],
                'ext': fi['ext'],
                'mtime': mtime,
                'chunks': chunks,
            }
    return {
        'version': INDEX_VERSION,
        'folder': os.path.abspath(folder),
        'built_at': time.time(),
        'files': files_index,
    }


def save_content_index(index, folder):
    path = index_cache_path(folder)
    with gzip.open(path, 'wt', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False)
    return path


def load_content_index(folder):
    path = index_cache_path(folder)
    if not os.path.isfile(path):
        return None
    try:
        with gzip.open(path, 'rt', encoding='utf-8') as f:
            data = json.load(f)
        if data.get('folder') != os.path.abspath(folder):
            return None
        if data.get('version') != INDEX_VERSION:
            return None
        return data
    except Exception as e:
        print(f"[WARN] load index: {e}")
        return None


def index_stats(index):
    if not index:
        return 0, 0, None
    files = index.get('files', {})
    chunks = sum(len(m.get('chunks', [])) for m in files.values())
    built_at = index.get('built_at')
    return len(files), chunks, built_at


def index_is_stale(index):
    if not index:
        return True
    for path, meta in index.get('files', {}).items():
        if not os.path.isfile(path):
            return True
        try:
            if os.path.getmtime(path) != meta.get('mtime'):
                return True
        except OSError:
            return True
    return False


def search_content_index(index, query, max_hits_per_file=15):
    q = query.strip()
    if not q or not index:
        return []
    ql = q.lower()
    results = []
    for path, meta in index.get('files', {}).items():
        if not os.path.isfile(path):
            continue
        hits = []
        for chunk in meta.get('chunks', []):
            text = chunk.get('text', '')
            if text and ql in text.lower():
                hits.append({
                    'page': chunk['page'],
                    'snippet': _text_snippet(text, q),
                })
                if len(hits) >= max_hits_per_file:
                    break
        if hits:
            results.append({
                'book': meta['book'],
                'file': meta['file'],
                'path': path,
                'ext': meta['ext'],
                'hits': hits,
            })
    return results


class ContentSearchWorker(QThread):
    """背景執行緒：關鍵字全文搜索（優先使用索引，否則即時掃描）。"""
    finished = pyqtSignal(list)
    progress = pyqtSignal(str)

    def __init__(self, query, index=None, books=None):
        super().__init__()
        self.query = query
        self.index = index
        self.books = books

    def run(self):
        if self.index:
            results = search_content_index(self.index, self.query)
        else:
            results = search_books_content(
                self.books, self.query,
                progress_cb=self.progress.emit,
                cancel_check=self.isInterruptionRequested,
            )
        if not self.isInterruptionRequested():
            self.finished.emit(results)


class IndexBuildWorker(QThread):
    """背景執行緒：建立或重建全文內容索引。"""
    finished = pyqtSignal(object)
    progress = pyqtSignal(str)

    def __init__(self, books, folder):
        super().__init__()
        self.books = books
        self.folder = folder

    def run(self):
        index = build_content_index(
            self.books, self.folder,
            progress_cb=self.progress.emit,
            cancel_check=self.isInterruptionRequested,
        )
        if not self.isInterruptionRequested() and index is not None:
            self.finished.emit(index)


# ══════════════════════════════════════════════════════════════
#  OPENERS — 開啟檔案與 Windows 顯示模式
# ══════════════════════════════════════════════════════════════
# open_pdf_at_page：依序嘗試 Sumatra / Foxit / Adobe，否則系統預設。
# switch_windows_display_mode：延伸桌面 ↔ 同步（複製）畫面（DisplaySwitch.exe）。
def open_pdf_at_page(path, page):
    if sys.platform == 'win32':
        for sp in [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
        ]:
            if os.path.exists(sp):
                subprocess.Popen([sp, '-reuse-instance', '-page', str(page), path])
                return True, 'Sumatra PDF'
        for fp in [
            r"C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe",
            r"C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe",
        ]:
            if os.path.exists(fp):
                subprocess.Popen([fp, path, '/A', f'page={page}'])
                return True, 'Foxit Reader'
        for ap in [
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        ]:
            if os.path.exists(ap):
                subprocess.Popen([ap, '/A', f'page={page}=OpenActions', path])
                return True, 'Adobe Reader'
        os.startfile(path)
        return True, '系統預設'
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
        return True, 'Preview'
    else:
        subprocess.Popen(['xdg-open', path])
        return True, '系統預設'


def switch_windows_display_mode(mode='duplicate'):
    """Switch Windows multi-monitor layout: duplicate (/clone) or extend (/extend)."""
    if sys.platform != 'win32':
        return False
    exe = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'System32', 'DisplaySwitch.exe')
    if not os.path.isfile(exe):
        return False
    flag = '/clone' if mode == 'duplicate' else '/extend'
    try:
        subprocess.Popen([exe, flag], shell=False)
        return True
    except Exception as e:
        print(f"[WARN] DisplaySwitch: {e}")
        return False


# ══════════════════════════════════════════════════════════════
#  BOOK NAME OVERLAY — 開啟檔案時的置頂書名浮層
# ══════════════════════════════════════════════════════════════
# 支援多螢幕延伸顯示；可置中→右下角動畫、定時隱藏或一直顯示；
# 關閉 Word/PDF 視窗時自動隱藏（document_viewer_is_open 輪詢）。
WORD_EXTS = {'.doc', '.docx'}
DEFAULT_OVERLAY_SECS = 10
MAX_OVERLAY_SECS = 10
OVERLAY_DURATION_ALWAYS = -1
CENTER_OVERLAY_SECS = 5
OVERLAY_MODE_CORNER = 'corner'
OVERLAY_MODE_CENTER = 'center_then_corner'
DEFAULT_QR_DESKTOP_SIZE = 140
DEFAULT_QR_X = 24
DEFAULT_QR_Y = 24
DEFAULT_QR_WIDTH = 220
DEFAULT_QR_HEIGHT = 268
DEFAULT_QR_CAPTION = '掃碼觀看'
DEFAULT_QR_BG_COLOR = '#ffffff'
DEFAULT_QR_TEXT_COLOR = ''
DEFAULT_QR_TEXT_SIZE = 12
DEFAULT_QR_BG_OPACITY = 88

SETTINGS_VERSION = 2


def default_settings():
    return {
        'version': SETTINGS_VERSION,
        'hymn_folder': '',
        'theme': 'dark',
        'font_size': 13,
        'search_mode': 'book',
        'open_overlay': True,
        'overlay_duration': DEFAULT_OVERLAY_SECS,
        'overlay_mode': OVERLAY_MODE_CORNER,
        'display_duplicate': False,
        'keyword_instant': True,
        'book_auto_focus_hymn': True,
        'click_to_open': False,
        'remote_api_enabled': False,
        'remote_accept': False,
        'remote_port': DEFAULT_PORT,
        'remote_use_https': False,
        'remote_policy': OPEN_POLICY_AUTO,
        'remote_token': '',
        'enable_tunnel': False,
        'cloudflared_path': '',
        'tunnel_mode': 'service',
        'cloudflared_token': (
            'eyJhIjoiY2M5NGFhNzkzZWI5MDlmYTJiNGQwYjRkMzJiZDllYmMiLCJ0Ijoi'
            'OWJlNjdlODYtZGNiMS00ZDc4LWJhOWItZTFkNjAwYmI3NzBmIiwicyI6Ik5URXdN'
            'RGhoTmpNdE56RmxOaTAwWVRCbUxXSmhNR010TW1GbU1XSm1OVFEwTnpFdyJ9'
        ),
        'named_tunnel_url': 'https://live.churchofgodtm.com/hymn_search/',
        'remote_path_prefix': '/hymn_search',
        'viewer_show_debug': False,
        'operator_mode': False,
        'pinned_books': [],
        'setlist': [],
        'setlist_index': -1,
        'session_history': [],
        'session_cursor': 0,
        'show_preview': True,
        'auto_rescan': True,
        'startup_tray': False,
        'minimize_to_tray': False,
        'last_opened': '',
        'show_qr_code': False,
        'qr_x': DEFAULT_QR_X,
        'qr_y': DEFAULT_QR_Y,
        'qr_width': DEFAULT_QR_WIDTH,
        'qr_height': DEFAULT_QR_HEIGHT,
        'qr_caption': DEFAULT_QR_CAPTION,
        'qr_bg_color': DEFAULT_QR_BG_COLOR,
        'qr_text_color': DEFAULT_QR_TEXT_COLOR,
        'qr_text_size': DEFAULT_QR_TEXT_SIZE,
        'qr_bg_opacity': DEFAULT_QR_BG_OPACITY,
    }


def normalize_settings(raw):
    merged = default_settings()
    if isinstance(raw, dict):
        for key in merged:
            if key == 'version':
                continue
            if key in raw:
                merged[key] = raw[key]
    if merged['theme'] not in THEMES:
        merged['theme'] = 'dark'
    try:
        merged['font_size'] = max(10, min(18, int(merged['font_size'])))
    except (TypeError, ValueError):
        merged['font_size'] = 13
    if merged['search_mode'] not in ('book', 'global', 'keyword', 'schedule'):
        merged['search_mode'] = 'book'
    merged['open_overlay'] = bool(merged['open_overlay'])
    try:
        dur = int(merged['overlay_duration'])
        if dur != OVERLAY_DURATION_ALWAYS and not (1 <= dur <= MAX_OVERLAY_SECS):
            dur = DEFAULT_OVERLAY_SECS
        merged['overlay_duration'] = dur
    except (TypeError, ValueError):
        merged['overlay_duration'] = DEFAULT_OVERLAY_SECS
    if merged['overlay_mode'] not in (OVERLAY_MODE_CORNER, OVERLAY_MODE_CENTER):
        merged['overlay_mode'] = OVERLAY_MODE_CORNER
    merged['display_duplicate'] = bool(merged['display_duplicate'])
    merged['keyword_instant'] = bool(merged['keyword_instant'])
    merged['book_auto_focus_hymn'] = bool(merged['book_auto_focus_hymn'])
    merged['click_to_open'] = bool(merged['click_to_open'])
    merged['remote_api_enabled'] = bool(merged['remote_api_enabled'])
    merged['remote_accept'] = bool(merged['remote_accept'])
    try:
        merged['remote_port'] = max(1024, min(65535, int(merged['remote_port'])))
    except (TypeError, ValueError):
        merged['remote_port'] = DEFAULT_PORT
    if merged['remote_policy'] not in (OPEN_POLICY_AUTO, OPEN_POLICY_CONFIRM, OPEN_POLICY_UI):
        merged['remote_policy'] = OPEN_POLICY_AUTO
    merged['remote_token'] = str(merged['remote_token'] or '')
    merged['remote_use_https'] = False
    merged['enable_tunnel'] = bool(merged.get('enable_tunnel'))
    merged['cloudflared_path'] = str(merged.get('cloudflared_path') or '')
    if merged.get('tunnel_mode') == 'named':
        merged['tunnel_mode'] = 'service'
    if merged.get('tunnel_mode') not in ('quick', 'service'):
        merged['tunnel_mode'] = 'service'
    merged['cloudflared_token'] = str(merged.get('cloudflared_token') or '')
    merged['named_tunnel_url'] = str(merged.get('named_tunnel_url') or '').strip()
    if merged['named_tunnel_url'].startswith('http://') and '.trycloudflare.com' not in merged['named_tunnel_url']:
        merged['named_tunnel_url'] = 'https://' + merged['named_tunnel_url'][7:]
    prefix = tunnel_url_path_prefix(merged['named_tunnel_url']) or str(merged.get('remote_path_prefix') or '').strip()
    if prefix and not prefix.startswith('/'):
        prefix = '/' + prefix
    merged['remote_path_prefix'] = prefix.rstrip('/') if prefix and prefix != '/' else prefix
    if prefix and merged['named_tunnel_url']:
        merged['named_tunnel_url'] = normalize_tunnel_url(merged['named_tunnel_url']) + '/'
    merged['viewer_show_debug'] = bool(merged.get('viewer_show_debug'))
    folder = str(merged.get('hymn_folder') or '').strip()
    merged['hymn_folder'] = os.path.abspath(folder) if folder else ''
    merged['operator_mode'] = bool(merged.get('operator_mode'))
    for list_key in ('pinned_books',):
        val = merged.get(list_key)
        if isinstance(val, list):
            seen = set()
            deduped = []
            for x in val:
                s = str(x or '').strip()
                if s and s not in seen:
                    seen.add(s)
                    deduped.append(s)
            merged[list_key] = deduped
        else:
            merged[list_key] = []
    merged['setlist'] = parse_setlist(merged.get('setlist'))
    try:
        merged['setlist_index'] = int(merged.get('setlist_index', -1))
    except (TypeError, ValueError):
        merged['setlist_index'] = -1
    merged['session_history'] = session_history_list(merged)
    try:
        merged['session_cursor'] = max(0, int(merged.get('session_cursor', 0)))
    except (TypeError, ValueError):
        merged['session_cursor'] = 0
    merged['show_preview'] = bool(merged.get('show_preview', True))
    merged['auto_rescan'] = bool(merged.get('auto_rescan', True))
    merged['startup_tray'] = bool(merged.get('startup_tray'))
    merged['minimize_to_tray'] = bool(merged.get('minimize_to_tray'))
    merged['last_opened'] = str(merged.get('last_opened') or '')
    merged['show_qr_code'] = bool(merged.get('show_qr_code'))
    try:
        merged['qr_x'] = max(0, min(3840, int(merged.get('qr_x', DEFAULT_QR_X))))
    except (TypeError, ValueError):
        merged['qr_x'] = DEFAULT_QR_X
    try:
        merged['qr_y'] = max(0, min(3840, int(merged.get('qr_y', DEFAULT_QR_Y))))
    except (TypeError, ValueError):
        merged['qr_y'] = DEFAULT_QR_Y
    try:
        merged['qr_width'] = max(120, min(800, int(merged.get('qr_width', DEFAULT_QR_WIDTH))))
    except (TypeError, ValueError):
        merged['qr_width'] = DEFAULT_QR_WIDTH
    try:
        merged['qr_height'] = max(120, min(800, int(merged.get('qr_height', DEFAULT_QR_HEIGHT))))
    except (TypeError, ValueError):
        merged['qr_height'] = DEFAULT_QR_HEIGHT
    merged['qr_caption'] = str(merged.get('qr_caption') or DEFAULT_QR_CAPTION)
    merged['qr_bg_color'] = _normalize_hex_color(
        merged.get('qr_bg_color'), DEFAULT_QR_BG_COLOR,
    )
    text_color = str(merged.get('qr_text_color') or '').strip()
    merged['qr_text_color'] = (
        _normalize_hex_color(text_color, DEFAULT_QR_TEXT_COLOR)
        if text_color else DEFAULT_QR_TEXT_COLOR
    )
    try:
        merged['qr_text_size'] = max(8, min(32, int(merged.get('qr_text_size', DEFAULT_QR_TEXT_SIZE))))
    except (TypeError, ValueError):
        merged['qr_text_size'] = DEFAULT_QR_TEXT_SIZE
    try:
        if 'qr_bg_opacity' in (raw if isinstance(raw, dict) else {}):
            opacity = merged.get('qr_bg_opacity', DEFAULT_QR_BG_OPACITY)
        elif isinstance(raw, dict) and 'qr_scrim_alpha' in raw:
            opacity = min(100, int(round(int(raw.get('qr_scrim_alpha', 0)) * 100 / 120)))
        else:
            opacity = merged.get('qr_bg_opacity', DEFAULT_QR_BG_OPACITY)
        merged['qr_bg_opacity'] = max(0, min(100, int(opacity)))
    except (TypeError, ValueError):
        merged['qr_bg_opacity'] = DEFAULT_QR_BG_OPACITY
    merged = clamp_setlist(merged)
    merged['version'] = SETTINGS_VERSION
    return merged


def load_settings():
    raw = {}
    path = settings_file_path()
    if os.path.isfile(path):
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            raw = data if isinstance(data, dict) else {}
        except Exception as e:
            print(f"[WARN] load settings: {e}")
    return normalize_settings(raw)


def save_settings(settings):
    try:
        data = normalize_settings(settings)
        with open(settings_file_path(), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARN] save settings: {e}")


def resolve_hymn_folder(settings):
    custom = (settings.get('hymn_folder') or '').strip()
    if custom:
        return os.path.abspath(custom)
    return os.path.abspath(default_hymn_folder())


def _overlay_duration_index(duration):
    if duration == OVERLAY_DURATION_ALWAYS:
        return MAX_OVERLAY_SECS
    return max(0, min(MAX_OVERLAY_SECS - 1, int(duration) - 1))


def _overlay_mode_index(mode):
    return 1 if mode == OVERLAY_MODE_CENTER else 0


def _remote_policy_index(policy):
    for idx, val in enumerate((OPEN_POLICY_AUTO, OPEN_POLICY_CONFIRM, OPEN_POLICY_UI)):
        if policy == val:
            return idx
    return 0


def _overlay_subtitle(text):
    """Display name without file extension for the overlay second line."""
    text = (text or '').strip()
    if not text:
        return ''
    low = text.lower()
    for ext in ('.docx', '.doc', '.pdf', '.odt'):
        if low.endswith(ext):
            return text[:-len(ext)]
    return text


def _ps_escape(text):
    return (text or '').replace("'", "''")


def _win_ps_bool(script):
    if sys.platform != 'win32':
        return True
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command', script],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
        )
        return (result.stdout or '').strip() == '1'
    except Exception as e:
        print(f"py[WARN] document watch: {e}")
        return True


def document_viewer_is_open(path):
    """Return True if the given file still appears open in Word or a PDF viewer."""
    if sys.platform != 'win32':
        return True
    if not path:
        return False
    base = os.path.splitext(os.path.basename(path))[0].strip()
    if not base:
        return False
    ext = os.path.splitext(path)[1].lower()
    safe = _ps_escape(base)

    if ext in WORD_EXTS:
        ps = f"""
$needle = '{safe}'
$open = Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowTitle -like "*$needle*" }} |
    Select-Object -First 1
if ($open) {{ '1' }}
"""
        return _win_ps_bool(ps)

    if ext == '.pdf':
        ps = f"""
$needle = '{safe}'
$names = @('SumatraPDF','FoxitPDFReader','AcroRd32','Acrobat','msedge','chrome','firefox')
$open = $false
foreach ($n in $names) {{
    $p = Get-Process -Name $n -ErrorAction SilentlyContinue |
        Where-Object {{ $_.MainWindowTitle -like "*$needle*" }} |
        Select-Object -First 1
    if ($p) {{ $open = $true; break }}
}}
if (-not $open) {{
    $p = Get-Process -ErrorAction SilentlyContinue |
        Where-Object {{
            $_.MainWindowTitle -and
            ($_.MainWindowTitle -like "*$needle*.pdf*" -or $_.MainWindowTitle -like "*$needle*.PDF*")
        }} |
        Select-Object -First 1
    if ($p) {{ $open = $true }}
}}
if ($open) {{ '1' }}
"""
        return _win_ps_bool(ps)

    return False


class _BookNameOverlayPanel(QWidget):
    """單一螢幕上的書名浮層面板（圓角黃底、置頂、不搶焦點）。"""
    _BG_COLOR = QColor('#f59e0b')

    def __init__(self, screen):
        super().__init__(
            None,
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool,
        )
        self._screen = screen
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self._radius = 12
        self._placement = 'corner'
        self._mode = OVERLAY_MODE_CORNER
        self._corner_duration = DEFAULT_OVERLAY_SECS
        self._animating = False
        self._anim_start_rect = None
        self._anim_end_rect = None
        self._book_label = QLabel(self)
        self._book_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._item_label = QLabel(self)
        self._item_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._item_label.setWordWrap(True)
        lay = QVBoxLayout(self)
        lay.addWidget(self._book_label)
        lay.addWidget(self._item_label)
        self._apply_corner_style()

        self._phase_timer = QTimer(self)
        self._phase_timer.setSingleShot(True)
        self._phase_timer.timeout.connect(self._switch_to_corner)

        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self._request_hide)

        self._raise_timer = QTimer(self)
        self._raise_timer.setInterval(2500)
        self._raise_timer.timeout.connect(self._keep_on_top)

        self._move_anim = QPropertyAnimation(self, b"geometry", self)
        self._move_anim.setDuration(700)
        self._move_anim.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self._move_anim.valueChanged.connect(self._on_move_anim_frame)
        self._move_anim.finished.connect(self._on_move_to_corner_finished)

    def _screen_geo(self):
        if self._screen is not None:
            return self._screen.availableGeometry()
        primary = QApplication.primaryScreen()
        return primary.availableGeometry() if primary else QRect()

    @staticmethod
    def _lerp(a, b, t):
        return a + (b - a) * t

    def _apply_interpolated_style(self, t):
        """Blend center → corner visuals while geometry animates."""
        t = max(0.0, min(1.0, float(t)))
        self._radius = self._lerp(18, 12, t)
        lay = self.layout()
        m_x = int(round(self._lerp(48, 28, t)))
        m_y = int(round(self._lerp(36, 20, t)))
        lay.setContentsMargins(m_x, m_y, m_x, m_y)
        lay.setSpacing(int(round(self._lerp(14, 8, t))))

        book_fs = int(round(self._lerp(48, 24, t)))
        item_fs = int(round(self._lerp(32, 18, t)))
        self._book_label.setStyleSheet(
            f"color: #ffffff; font-size: {book_fs}px; font-weight: bold; background: transparent;"
        )
        self._item_label.setStyleSheet(
            f"color: #ffffff; font-size: {item_fs}px; background: transparent;"
        )

        self.setMinimumWidth(0)
        self.setMaximumWidth(16777215)
        self._book_label.setMinimumWidth(0)
        self._item_label.setMinimumWidth(0)
        self._item_label.setMaximumWidth(16777215)
        self.update()

    def _anim_progress(self, rect):
        start, end = self._anim_start_rect, self._anim_end_rect
        if not start or not end:
            return 1.0
        dw = end.width() - start.width()
        if dw:
            return max(0.0, min(1.0, (rect.width() - start.width()) / dw))
        dh = end.height() - start.height()
        if dh:
            return max(0.0, min(1.0, (rect.height() - start.height()) / dh))
        return 1.0

    def _on_move_anim_frame(self, value):
        if not self._animating:
            return
        rect = value
        if not isinstance(rect, QRect):
            return
        self._apply_interpolated_style(self._anim_progress(rect))

    def _apply_corner_style(self):
        self._radius = 12
        self.setMinimumWidth(320)
        self.setMaximumWidth(16777215)
        lay = self.layout()
        lay.setContentsMargins(28, 20, 28, 20)
        lay.setSpacing(8)
        self._book_label.setStyleSheet(
            "color: #ffffff; font-size: 24px; font-weight: bold; background: transparent;"
        )
        self._item_label.setStyleSheet(
            "color: #ffffff; font-size: 18px; background: transparent;"
        )
        self._item_label.setMinimumWidth(280)
        self._item_label.setMaximumWidth(520)

    def _apply_center_style(self):
        self._radius = 18
        self.setMinimumWidth(520)
        self.setMaximumWidth(900)
        lay = self.layout()
        lay.setContentsMargins(48, 36, 48, 36)
        lay.setSpacing(14)
        self._book_label.setStyleSheet(
            "color: #ffffff; font-size: 48px; font-weight: bold; background: transparent;"
        )
        self._item_label.setStyleSheet(
            "color: #ffffff; font-size: 32px; background: transparent;"
        )
        self._item_label.setMinimumWidth(440)
        self._item_label.setMaximumWidth(780)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.addRoundedRect(
            QRectF(0, 0, self.width(), self.height()),
            self._radius, self._radius,
        )
        painter.fillPath(path, self._BG_COLOR)
        super().paintEvent(event)

    def _ensure_on_screen(self):
        wh = self.windowHandle()
        if wh is None:
            self.show()
            wh = self.windowHandle()
        if wh is not None and self._screen is not None:
            wh.setScreen(self._screen)

    def _request_hide(self):
        BookNameOverlay.hide_overlay()

    @staticmethod
    def _normalize_duration(duration_sec):
        if duration_sec == OVERLAY_DURATION_ALWAYS:
            return OVERLAY_DURATION_ALWAYS
        try:
            return max(1, min(int(duration_sec), MAX_OVERLAY_SECS))
        except (TypeError, ValueError):
            return DEFAULT_OVERLAY_SECS

    def _start_corner_hide_timer(self):
        self._hide_timer.stop()
        if self._corner_duration == OVERLAY_DURATION_ALWAYS:
            return
        self._hide_timer.start(self._corner_duration * 1000)

    def _activate(self, book_name, watch_path, item_name='', duration_sec=DEFAULT_OVERLAY_SECS,
                  mode=OVERLAY_MODE_CORNER):
        self._corner_duration = self._normalize_duration(duration_sec)
        self._watch_path = watch_path or ''
        self._mode = mode
        self._book_label.setText(book_name)
        sub = _overlay_subtitle(item_name or os.path.basename(watch_path or ''))
        self._item_label.setText(sub)
        self._item_label.setVisible(bool(sub))

        self._hide_timer.stop()
        self._phase_timer.stop()
        self._move_anim.stop()
        self._animating = False
        self._anim_start_rect = None
        self._anim_end_rect = None

        if mode == OVERLAY_MODE_CENTER:
            self._placement = 'center'
            self._apply_center_style()
        else:
            self._placement = 'corner'
            self._apply_corner_style()

        self.adjustSize()
        self._position_screen()
        self.show()
        self._ensure_on_screen()
        self._position_screen()
        self.raise_()
        self._raise_timer.start()

        if mode == OVERLAY_MODE_CENTER:
            self._phase_timer.start(CENTER_OVERLAY_SECS * 1000)
        else:
            self._start_corner_hide_timer()

    def _switch_to_corner(self):
        self._phase_timer.stop()
        self._hide_timer.stop()
        start_rect = self.geometry()

        self._placement = 'corner'
        self._apply_corner_style()
        self.adjustSize()

        geo = self._screen_geo()
        if geo.isNull():
            self._finish_corner_placement()
            return

        w, h = self.width(), self.height()
        end_rect = QRect(geo.right() - w - 24, geo.bottom() - h - 24, w, h)

        self._apply_center_style()
        self.setGeometry(start_rect)

        self._anim_start_rect = start_rect
        self._anim_end_rect = end_rect
        self._move_anim.stop()
        self._animating = True
        self._apply_interpolated_style(0.0)
        self._move_anim.setStartValue(start_rect)
        self._move_anim.setEndValue(end_rect)
        self._move_anim.start()
        self.raise_()

    def _on_move_to_corner_finished(self):
        self._animating = False
        self._anim_start_rect = None
        self._anim_end_rect = None
        self._finish_corner_placement()

    def _finish_corner_placement(self):
        self._placement = 'corner'
        self._apply_corner_style()
        self._position_corner()
        self.raise_()
        self._start_corner_hide_timer()

    def _deactivate(self):
        self._phase_timer.stop()
        self._hide_timer.stop()
        self._raise_timer.stop()
        self._move_anim.stop()
        self._animating = False
        self._anim_start_rect = None
        self._anim_end_rect = None
        self.hide()

    def _position_corner(self):
        geo = self._screen_geo()
        if geo.isNull():
            return
        self.adjustSize()
        x = geo.right() - self.width() - 24
        y = geo.bottom() - self.height() - 24
        self.move(x, y)

    def _position_center(self):
        geo = self._screen_geo()
        if geo.isNull():
            return
        self.adjustSize()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)

    def _position_screen(self):
        if self._placement == 'center':
            self._position_center()
        else:
            self._position_corner()

    def _keep_on_top(self):
        if self.isVisible() and not self._animating:
            self._ensure_on_screen()
            self._position_screen()
            self.raise_()

    def _snap_to_corner(self):
        """Stop transitions and place at corner (used before display topology changes)."""
        self._phase_timer.stop()
        self._hide_timer.stop()
        self._move_anim.stop()
        self._animating = False
        self._anim_start_rect = None
        self._anim_end_rect = None
        self._placement = 'corner'
        self._mode = OVERLAY_MODE_CORNER
        self._apply_corner_style()
        self.adjustSize()
        self._position_corner()
        self.raise_()


class BookNameOverlay:
    """管理所有螢幕的書名浮層；協調顯示、動畫、文件關閉監聽與顯示模式切換後的重新定位。"""
    _panels = []
    _session = None
    _watch_path = ''
    _watch_timer = None
    _watch_busy = False

    @classmethod
    def _visible_panels(cls):
        return [p for p in cls._panels if p.isVisible()]

    @classmethod
    def _start_document_watch(cls, watch_path=''):
        path = os.path.abspath(watch_path) if watch_path else ''
        cls._watch_path = path
        if not path:
            cls._stop_document_watch()
            return
        if cls._watch_timer is None:
            cls._watch_timer = QTimer()
            cls._watch_timer.setInterval(1500)
            cls._watch_timer.timeout.connect(cls._poll_document_viewer)
        cls._watch_timer.start()

    @classmethod
    def _poll_document_viewer(cls):
        if cls._watch_busy:
            return
        if not cls._watch_path or not cls._visible_panels():
            cls._stop_document_watch()
            return
        path = cls._watch_path
        cls._watch_busy = True

        def _worker():
            try:
                if not document_viewer_is_open(path):
                    QTimer.singleShot(0, cls.hide_overlay)
            finally:
                cls._watch_busy = False

        threading.Thread(target=_worker, daemon=True).start()

    @classmethod
    def _stop_document_watch(cls):
        if cls._watch_timer is not None:
            cls._watch_timer.stop()
        cls._watch_path = ''

    @classmethod
    def prepare_for_display_change(cls):
        panels = cls._visible_panels()
        if not panels:
            return
        ref = panels[0]
        cls._session = {
            'book_name': ref._book_label.text(),
            'item_name': ref._item_label.text() if ref._item_label.isVisible() else '',
            'watch_path': getattr(ref, '_watch_path', '') or '',
            'duration_sec': ref._corner_duration,
        }
        for panel in panels:
            panel._snap_to_corner()

    @classmethod
    def refresh_after_display_change(cls):
        session = cls._session
        if not session:
            return
        app = QApplication.instance()
        if app:
            app.processEvents()
        cls._sync_panels()
        screens = QApplication.screens() or []
        primary = QApplication.primaryScreen()
        # Duplicate mode: one overlay on primary avoids doubled / misplaced hints.
        if primary and len(screens) > 1:
            target_screens = [primary]
        else:
            target_screens = screens
        target_set = set(target_screens)
        for panel in list(cls._panels):
            if panel._screen not in target_set:
                panel._deactivate()
                panel.deleteLater()
                cls._panels.remove(panel)
        by_screen = {p._screen: p for p in cls._panels}
        for screen in target_screens:
            if screen not in by_screen:
                cls._panels.append(_BookNameOverlayPanel(screen))
        for panel in cls._panels:
            if panel._screen in target_set:
                panel._activate(
                    session['book_name'],
                    session.get('watch_path', ''),
                    session.get('item_name', ''),
                    session['duration_sec'],
                    OVERLAY_MODE_CORNER,
                )
        cls._start_document_watch(session.get('watch_path', ''))

    @classmethod
    def _sync_panels(cls):
        screens = QApplication.screens() or []
        screen_set = set(screens)
        for panel in list(cls._panels):
            if panel._screen not in screen_set:
                panel._deactivate()
                panel.deleteLater()
                cls._panels.remove(panel)

        by_screen = {p._screen: p for p in cls._panels}
        for screen in screens:
            if screen not in by_screen:
                cls._panels.append(_BookNameOverlayPanel(screen))

    @classmethod
    def show_book(cls, book_name, watch_path='', item_name='', duration_sec=DEFAULT_OVERLAY_SECS,
                  mode=OVERLAY_MODE_CORNER):
        if not book_name:
            return
        cls._sync_panels()
        if not cls._panels:
            return
        for panel in cls._panels:
            panel._activate(book_name, watch_path, item_name, duration_sec, mode)
        cls._start_document_watch(watch_path)

    @classmethod
    def hide_overlay(cls):
        cls._stop_document_watch()
        for panel in cls._panels:
            panel._deactivate()


def _normalize_hex_color(value, default='#ffffff'):
    s = str(value or '').strip()
    if not s:
        return default
    if not s.startswith('#'):
        s = '#' + s
    if re.fullmatch(r'#[0-9a-fA-F]{6}', s):
        return s.lower()
    if re.fullmatch(r'#[0-9a-fA-F]{3}', s):
        return '#' + ''.join(ch * 2 for ch in s[1:]).lower()
    return default


def _text_color_for_bg(hex_color):
    c = QColor(hex_color)
    lum = 0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue()
    return '#ffffff' if lum < 140 else '#333333'


def _border_color_for_bg(hex_color):
    c = QColor(hex_color)
    return QColor(
        max(0, c.red() - 28),
        max(0, c.green() - 28),
        max(0, c.blue() - 28),
    )


def _qcolor_with_opacity(hex_color, opacity_pct):
    c = QColor(_normalize_hex_color(hex_color, DEFAULT_QR_BG_COLOR))
    alpha = max(0, min(255, int(round(255 * int(opacity_pct) / 100))))
    c.setAlpha(alpha)
    return c


def _desktop_qr_render_size(card_w, card_h, has_caption, text_size=DEFAULT_QR_TEXT_SIZE):
    """QR 像素邊長（緊湊內距、QR 佔比更大）。"""
    pad_x, pad_top, pad_bottom, spacing = 3, 3, 2, 2
    caption_h = max(18, int(text_size) + 10) if has_caption else 0
    if not has_caption:
        spacing = 0
    inner_w = max(56, int(card_w) - 2 * pad_x)
    inner_h = max(
        56,
        int(card_h) - pad_top - pad_bottom - caption_h - spacing,
    )
    return max(56, min(inner_w, inner_h, 640))


def _desktop_qr_style_from_settings(settings):
    text_color_raw = str((settings or {}).get('qr_text_color') or '').strip()
    return {
        'bg_color': _normalize_hex_color(
            (settings or {}).get('qr_bg_color'), DEFAULT_QR_BG_COLOR,
        ),
        'text_color': (
            _normalize_hex_color(text_color_raw, DEFAULT_QR_TEXT_COLOR)
            if text_color_raw else ''
        ),
        'text_size': int((settings or {}).get('qr_text_size', DEFAULT_QR_TEXT_SIZE)),
        'bg_opacity': int((settings or {}).get('qr_bg_opacity', DEFAULT_QR_BG_OPACITY)),
    }


class _DesktopQrPanel(QWidget):
    """單一螢幕 QR 浮層（PowerPoint Live 式卡片：圓角、淺陰影、置中 QR、下方說明）。"""

    _SHADOW_PAD = 2
    _RADIUS = 6
    _PAD_X = 3
    _PAD_TOP = 3
    _PAD_BOTTOM = 2
    _CAPTION_SPACING = 2

    def __init__(self, screen):
        super().__init__(
            None,
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool,
        )
        self._screen = screen
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self._bg_color = DEFAULT_QR_BG_COLOR
        self._text_color = ''
        self._text_size = DEFAULT_QR_TEXT_SIZE
        self._bg_opacity = DEFAULT_QR_BG_OPACITY
        self._card_w = DEFAULT_QR_WIDTH
        self._card_h = DEFAULT_QR_HEIGHT

        self._card = QFrame(self)
        self._card.setObjectName('qrLiveCard')
        self._card.setStyleSheet('background: transparent; border: none;')

        self._qr_label = QLabel(self._card)
        self._qr_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._qr_label.setStyleSheet('background: transparent;')

        self._caption_label = QLabel(self._card)
        self._caption_label.setAlignment(
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
        )
        self._caption_label.setWordWrap(True)
        self._apply_caption_style()

        card_lay = QVBoxLayout(self._card)
        card_lay.setContentsMargins(
            self._PAD_X, self._PAD_TOP, self._PAD_X, self._PAD_BOTTOM,
        )
        card_lay.setSpacing(self._CAPTION_SPACING)
        card_lay.addWidget(self._qr_label, 0, Qt.AlignmentFlag.AlignHCenter)
        card_lay.addWidget(self._caption_label, 0, Qt.AlignmentFlag.AlignHCenter)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(self._SHADOW_PAD, self._SHADOW_PAD, self._SHADOW_PAD, self._SHADOW_PAD)
        outer.addWidget(self._card)

        self._raise_timer = QTimer(self)
        self._raise_timer.setInterval(2500)
        self._raise_timer.timeout.connect(self._keep_on_top)

    def _screen_geo(self):
        if self._screen is not None:
            return self._screen.availableGeometry()
        primary = QApplication.primaryScreen()
        return primary.availableGeometry() if primary else QRect()

    def _ensure_on_screen(self):
        wh = self.windowHandle()
        if wh is None:
            self.show()
            wh = self.windowHandle()
        if wh is not None and self._screen is not None:
            wh.setScreen(self._screen)

    def _apply_caption_style(self):
        if self._text_color:
            text_color = self._text_color
        else:
            text_color = _text_color_for_bg(self._bg_color)
        self._caption_label.setStyleSheet(
            f"color: {text_color}; font-size: {int(self._text_size)}px; font-weight: 600; "
            f"background: transparent; padding: 0 2px;"
        )

    def _card_origin(self):
        return float(self._SHADOW_PAD)

    def _card_rect(self):
        o = self._card_origin()
        return QRectF(o, o, self._card_w, self._card_h)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        card = self._card_rect()

        shadow = QPainterPath()
        shadow.addRoundedRect(
            card.translated(1, 1), self._RADIUS, self._RADIUS,
        )
        painter.fillPath(shadow, QColor(0, 0, 0, 8))

        path = QPainterPath()
        path.addRoundedRect(card, self._RADIUS, self._RADIUS)
        painter.fillPath(path, _qcolor_with_opacity(self._bg_color, self._bg_opacity))
        super().paintEvent(event)

    def show_qr(
        self, pixmap, caption, bg_color, text_color, text_size, bg_opacity,
        x, y, width, height,
    ):
        self._bg_color = _normalize_hex_color(bg_color, DEFAULT_QR_BG_COLOR)
        raw_text = str(text_color or '').strip()
        self._text_color = _normalize_hex_color(raw_text, '') if raw_text else ''
        self._text_size = max(8, min(32, int(text_size)))
        self._bg_opacity = max(0, min(100, int(bg_opacity)))
        self._card_w = max(120, int(width))
        self._card_h = max(120, int(height))
        self._apply_caption_style()
        self._card.setFixedSize(self._card_w, self._card_h)

        caption_text = caption or ''
        self._caption_label.setText(caption_text)
        self._caption_label.setVisible(bool(caption_text))
        if caption_text:
            self._caption_label.setFixedWidth(max(80, self._card_w - 2 * self._PAD_X))

        qr_size = _desktop_qr_render_size(
            self._card_w, self._card_h, bool(caption_text), self._text_size,
        )
        scaled = pixmap.scaled(
            qr_size, qr_size,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._qr_label.setPixmap(scaled)
        self._qr_label.setFixedSize(qr_size, qr_size)

        total_w = self._card_w + 2 * self._SHADOW_PAD
        total_h = self._card_h + 2 * self._SHADOW_PAD
        self.setFixedSize(total_w, total_h)
        self.update()
        self._position(x, y)
        self.show()
        self._ensure_on_screen()
        self._position(x, y)
        self.raise_()
        self._raise_timer.start()

    def _position(self, x, y):
        geo = self._screen_geo()
        if geo.isNull():
            return
        px = geo.x() + int(x)
        py = geo.y() + int(y)
        px = max(geo.x(), min(px, geo.right() - self.width() + 1))
        py = max(geo.y(), min(py, geo.bottom() - self.height() + 1))
        self.move(px, py)

    def _keep_on_top(self):
        if self.isVisible():
            self._ensure_on_screen()
            self.raise_()

    def hide_qr(self):
        self._raise_timer.stop()
        self.hide()


class DesktopQrOverlay:
    """管理各螢幕桌面 QR Code 浮層（延伸畫面每屏各顯示；同步複製時僅主螢幕）。"""

    _panels = []
    _cached = None

    @classmethod
    def _target_screens(cls, duplicate_only=False):
        screens = QApplication.screens() or []
        primary = QApplication.primaryScreen()
        if duplicate_only and primary and len(screens) > 1:
            return [primary]
        return screens

    @classmethod
    def _sync_panels(cls, duplicate_only=False):
        screens = cls._target_screens(duplicate_only)
        screen_set = set(screens)
        for panel in list(cls._panels):
            if panel._screen not in screen_set:
                panel.hide_qr()
                panel.deleteLater()
                cls._panels.remove(panel)
        by_screen = {p._screen: p for p in cls._panels}
        for screen in screens:
            if screen not in by_screen:
                cls._panels.append(_DesktopQrPanel(screen))

    @classmethod
    def _apply_to_panel(cls, panel, cached):
        panel.show_qr(
            cached['pixmap'],
            cached.get('caption', ''),
            cached.get('bg_color', DEFAULT_QR_BG_COLOR),
            cached.get('text_color', DEFAULT_QR_TEXT_COLOR),
            cached.get('text_size', DEFAULT_QR_TEXT_SIZE),
            cached.get('bg_opacity', DEFAULT_QR_BG_OPACITY),
            cached['x'],
            cached['y'],
            cached.get('width', DEFAULT_QR_WIDTH),
            cached.get('height', DEFAULT_QR_HEIGHT),
        )

    @classmethod
    def show_qr(cls, pixmap, caption, style, x, y, width, height):
        style = style or {}
        cls._cached = {
            'pixmap': pixmap,
            'caption': caption,
            'bg_color': style.get('bg_color', DEFAULT_QR_BG_COLOR),
            'text_color': style.get('text_color', DEFAULT_QR_TEXT_COLOR),
            'text_size': style.get('text_size', DEFAULT_QR_TEXT_SIZE),
            'bg_opacity': style.get('bg_opacity', DEFAULT_QR_BG_OPACITY),
            'x': x,
            'y': y,
            'width': width,
            'height': height,
        }
        cls._sync_panels(duplicate_only=False)
        for panel in cls._panels:
            cls._apply_to_panel(panel, cls._cached)

    @classmethod
    def reposition(cls, x, y, width=None, height=None):
        if cls._cached:
            cls._cached['x'] = x
            cls._cached['y'] = y
            if width is not None:
                cls._cached['width'] = width
            if height is not None:
                cls._cached['height'] = height
        if cls._cached and width is not None and height is not None and cls._panels:
            cached = cls._cached
            cached['x'] = x
            cached['y'] = y
            cached['width'] = width
            cached['height'] = height
            for panel in cls._panels:
                if panel.isVisible():
                    cls._apply_to_panel(panel, cached)
            return
        for panel in cls._panels:
            if panel.isVisible():
                panel._position(x, y)

    @classmethod
    def hide(cls):
        cls._cached = None
        for panel in cls._panels:
            panel.hide_qr()

    @classmethod
    def refresh_after_display_change(cls):
        cached = cls._cached
        if not cached:
            return
        for panel in list(cls._panels):
            panel.hide_qr()
            panel.deleteLater()
        cls._panels = []
        cls._sync_panels(duplicate_only=True)
        for panel in cls._panels:
            cls._apply_to_panel(panel, cached)


def _schedule_word_focus_keys(doc_path='', after_callback=None):
    """Word 開啟後送出 Enter、Alt+W、Alt+O，使視圖適合投影（無需 pywin32）。"""
    def _worker():
        time.sleep(2)
        doc_name = os.path.splitext(os.path.basename(doc_path))[0].replace("'", "''")
        ps = f"""
$s = New-Object -ComObject WScript.Shell
$p = Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowTitle -like '*{doc_name}*' }} |
    Select-Object -First 1
if (-not $p) {{
    $p = Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -First 1
}}
if ($p -and $p.MainWindowTitle) {{ [void]$s.AppActivate($p.MainWindowTitle) }}
Start-Sleep -Milliseconds 250
$s.SendKeys('{{ENTER}}')
Start-Sleep -Milliseconds 150
$s.SendKeys('%w')
Start-Sleep -Milliseconds 150
$s.SendKeys('%o')
"""
        try:
            subprocess.run(
                ['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command', ps],
                check=False,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        except Exception as e:
            print(f"[WARN] Word focus keys: {e}")
        if after_callback:
            try:
                after_callback()
            except Exception as e:
                print(f"[WARN] after Word focus keys: {e}")

    threading.Thread(target=_worker, daemon=True).start()


def open_docx(path, after_focus_keys=None):
    """以系統預設方式開啟 Word 檔；可選在快捷鍵送出後執行回呼（如切換同步畫面）。"""
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 3)
        _schedule_word_focus_keys(path, after_callback=after_focus_keys)
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['libreoffice', '--writer', path])


# ══════════════════════════════════════════════════════════════
#  DROPDOWN — 書名輸入時的浮動書冊建議列表
# ══════════════════════════════════════════════════════════════
class DropdownList(QListWidget):
    """置頂浮動清單；書冊或詩歌號/檔名建議。"""
    book_selected = pyqtSignal(dict)
    text_selected = pyqtSignal(str)

    def __init__(self, theme='dark'):
        super().__init__(None)
        self._theme = theme
        self._books = []
        self._entries = []
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self._refresh_style()
        self.itemClicked.connect(self._on_click)

    def _refresh_style(self):
        t = THEMES[self._theme]
        self.setStyleSheet(f"""
            QListWidget {{
                background: {t['panel']}; border: 1.5px solid {t['accent']};
                border-radius: 10px; outline: none; font-size: 13px; padding: 4px;
            }}
            QListWidget::item {{
                padding: 8px 12px; color: {t['text']};
                border-radius: 6px; margin: 1px 2px;
            }}
            QListWidget::item:hover  {{ background: {t['panel2']}; }}
            QListWidget::item:selected {{ background: {t['accent']}; color: {t['a_text']}; }}
        """)

    def set_theme(self, theme):
        self._theme = theme
        self._refresh_style()

    def _on_click(self, item):
        idx = self.row(item)
        if self._entries:
            if 0 <= idx < len(self._entries):
                self.text_selected.emit(self._entries[idx]['value'])
            self.hide()
            return
        if 0 <= idx < len(self._books):
            self.book_selected.emit(self._books[idx])
        self.hide()

    def refresh(self, books, anchor_widget):
        self._books = books
        self._entries = []
        self.clear()
        for b in books:
            badges = []
            if b['docx']:     badges.append('WORD')
            if b['bookmark']: badges.append('PDF🔖')
            elif b['pdf']:    badges.append('PDF')
            cnt = f"  {b['hymn_count']}首" if b['hymn_count'] else ''
            self.addItem(f"  📖  {b['name']}{cnt}    {'  '.join(badges)}")
        pos = anchor_widget.mapToGlobal(QPoint(0, anchor_widget.height() + 4))
        self.move(pos)
        self.resize(anchor_widget.width(), min(len(books) * 38 + 12, 280))
        self.show()
        self.raise_()

    def refresh_entries(self, entries, anchor_widget):
        self._entries = entries
        self._books = []
        self.clear()
        for e in entries:
            self.addItem(f"  {e.get('label', e.get('value', ''))}")
        if not entries:
            self.hide()
            return
        pos = anchor_widget.mapToGlobal(QPoint(0, anchor_widget.height() + 4))
        self.move(pos)
        self.resize(anchor_widget.width(), min(len(entries) * 36 + 12, 260))
        self.show()
        self.raise_()


# ══════════════════════════════════════════════════════════════
#  MAIN WINDOW — 主視窗（UI 建構、搜尋邏輯、事件處理）
# ══════════════════════════════════════════════════════════════
class MainWindow(QMainWindow, EnhancementMixin):
    """
    應用程式主視窗。

    職責：
    - 左側固定寬度書冊列表；右側搜尋列 + 檔案/結果列表
    - 三種搜尋模式切換、設定面板、主題與字體
    - 協調開檔流程（提示 → 開啟 → Word 快捷鍵 → 可選同步畫面）
    """
    display_duplicate_requested = pyqtSignal()  # 從背景執行緒安全觸發顯示模式切換
    remote_open_payload = pyqtSignal(dict)
    remote_populate_ui = pyqtSignal(str, str, list)
    remote_pending_request = pyqtSignal(object)
    tunnel_url_signal = pyqtSignal(str)
    tunnel_ready_signal = pyqtSignal(str)
    tunnel_error_signal = pyqtSignal(str)
    tunnel_expired_signal = pyqtSignal(str)
    service_install_signal = pyqtSignal(bool, str)
    _main_invoke = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.books        = []
        self.current_book = None
        self._loading_settings = True
        self._settings    = load_settings()
        self.hymn_folder  = resolve_hymn_folder(self._settings)
        self.search_mode  = self._settings['search_mode']
        self.theme_name   = self._settings['theme']
        self.font_size    = self._settings['font_size']
        self._search_worker = None
        self._index_worker = None
        self.content_index = None
        self._app         = QApplication.instance()
        self._remote      = RemoteServer()
        self._remote_pending_id = None
        self._tunnel = None
        self.enable_tunnel = self._settings['enable_tunnel']
        self.cloudflared_path = self._settings['cloudflared_path']
        self.tunnel_mode = self._settings['tunnel_mode']
        self.cloudflared_token = self._settings['cloudflared_token']
        self.named_tunnel_url = self._settings['named_tunnel_url']
        self.remote_path_prefix = self._settings['remote_path_prefix']
        self.remote_use_https = self._settings['remote_use_https']

        ver_suffix = f" v{APP_VERSION}"
        if BUILD_DATE:
            ver_suffix += f" ({BUILD_DATE})"
        self.setWindowTitle(f"詩歌冊搜索系統{ver_suffix}")
        icon = load_app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setMinimumSize(900, 580)
        self.resize(1120, 730)

        self.dropdown = DropdownList(self.theme_name)
        self.dropdown.book_selected.connect(self._select_book)
        self.sched_entry_dropdown = DropdownList(self.theme_name)
        self.sched_entry_dropdown.text_selected.connect(self._on_sched_entry_picked)
        self.display_duplicate_requested.connect(self._switch_duplicate_with_overlay_fix)
        self.remote_open_payload.connect(self._on_remote_open_payload)
        self.remote_populate_ui.connect(self._on_remote_populate_ui)
        self.remote_pending_request.connect(self._on_remote_pending_request)
        self.tunnel_url_signal.connect(self._on_tunnel_url)
        self.tunnel_ready_signal.connect(self._on_tunnel_ready)
        self.tunnel_error_signal.connect(self._on_tunnel_error)
        self.tunnel_expired_signal.connect(self._on_tunnel_expired)
        self.service_install_signal.connect(self._on_service_install_done)
        self._main_invoke.connect(self._run_main_invoke, Qt.ConnectionType.QueuedConnection)
        self._remote.set_handlers(
            lambda: self.books,
            lambda p: self.remote_open_payload.emit(p),
            lambda b, n, m: self.remote_populate_ui.emit(b, n, m),
            lambda r: self.remote_pending_request.emit(r),
            self._remote_prepare_setlist_next,
            self._get_setlist_info,
            self._remote_prepare_setlist_open,
            self._remote_handle_setlist_add,
        )

        self._build_ui()
        self._setup_focus_shortcuts()
        self._setup_app_shortcuts()
        self._setup_tray_icon()
        self._setup_folder_watcher()
        if self._app:
            self._app.installEventFilter(self)
        self._apply_theme()
        self._restore_saved_settings()
        self._scan_all()
        self._apply_setlist_from_settings()
        self._update_footer_status()
        if self._settings.get('startup_tray'):
            self._minimize_to_tray()

    def _setup_focus_shortcuts(self):
        for seq in (
            QKeySequence(Qt.Modifier.CTRL | Qt.Key.Key_Return),
            QKeySequence(Qt.Modifier.CTRL | Qt.Key.Key_Enter),
        ):
            sc = QShortcut(seq, self)
            sc.setContext(Qt.ShortcutContext.WindowShortcut)
            sc.activated.connect(self._focus_book_input)

    def _setup_app_shortcuts(self):
        sc = QShortcut(QKeySequence(Qt.Key.Key_Question), self)
        sc.setContext(Qt.ShortcutContext.WindowShortcut)
        sc.activated.connect(self._show_shortcuts_dialog)
        for key, mode in (
            (Qt.Key.Key_F1, 'book'),
            (Qt.Key.Key_F2, 'global'),
            (Qt.Key.Key_F3, 'keyword'),
            (Qt.Key.Key_F4, 'schedule'),
        ):
            sc = QShortcut(QKeySequence(key), self)
            sc.setContext(Qt.ShortcutContext.WindowShortcut)
            sc.activated.connect(lambda m=mode: self._set_search_mode(m))
        for key, fn in (
            (Qt.Key.Key_Left, self._open_previous_session),
            (Qt.Key.Key_Right, self._open_next_session),
        ):
            sc = QShortcut(QKeySequence(Qt.Modifier.CTRL | key), self)
            sc.setContext(Qt.ShortcutContext.WindowShortcut)
            sc.activated.connect(fn)
        for key, fn in (
            (Qt.Key.Key_Left, self._setlist_prev_shortcut),
            (Qt.Key.Key_Right, self._setlist_next_shortcut),
        ):
            sc = QShortcut(QKeySequence(Qt.Modifier.ALT | key), self)
            sc.setContext(Qt.ShortcutContext.WindowShortcut)
            sc.activated.connect(fn)

    def _focus_book_input(self):
        if self.search_mode != 'book':
            self._set_search_mode('book')
        self.inp_book.setFocus(Qt.FocusReason.ShortcutFocusReason)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._sync_settings_overlay_geometry()

    def _sync_settings_overlay_geometry(self):
        if hasattr(self, 'settings_overlay') and hasattr(self, 'body_container'):
            self.settings_overlay.setGeometry(self.body_container.rect())

    def _setup_settings_overlay(self):
        self.settings_overlay = QFrame(self.body_container)
        self.settings_overlay.setObjectName('settingsOverlay')
        self.settings_overlay.setVisible(False)
        ovl = QVBoxLayout(self.settings_overlay)
        ovl.setContentsMargins(24, 18, 24, 18)
        ovl.setSpacing(12)

        hdr = QHBoxLayout()
        self.settings_overlay_title = QLabel("⚙  設定")
        hdr.addWidget(self.settings_overlay_title)
        hdr.addStretch()
        self.btn_settings_close = QPushButton("收起")
        self.btn_settings_close.setFixedHeight(30)
        self.btn_settings_close.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_settings_close.clicked.connect(lambda: self.btn_settings.setChecked(False))
        hdr.addWidget(self.btn_settings_close)
        ovl.addLayout(hdr)

        self.settings_frame.setVisible(True)
        ovl.addWidget(self.settings_scroll, 1)
        self._sync_settings_overlay_geometry()

    # ─────────────────────────────────────────────────────────────
    #  ROOT — 根版面：頂欄 + 主體（側欄+內容）+ 底部 Toast
    # ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_header())

        self.body_container = QWidget()
        container_lay = QVBoxLayout(self.body_container)
        container_lay.setContentsMargins(0, 0, 0, 0)
        container_lay.setSpacing(0)

        body_widget = QWidget()
        body_lay = QHBoxLayout(body_widget)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(0)
        body_lay.addWidget(self._build_sidebar())
        body_lay.addWidget(self._build_content(), 1)
        self.body = body_widget
        container_lay.addWidget(body_widget, 1)

        self._setup_settings_overlay()
        root.addWidget(self.body_container, 1)

        root.addWidget(self._build_footer_bar())
        root.addWidget(self._build_toast())

    # ─────────────────────────────────────────────────────────────
    #  HEADER — 頂部標題列（程式名、資料夾路徑、狀態）
    # ─────────────────────────────────────────────────────────────
    def _build_header(self):
        self.header_frame = QFrame()
        self.header_frame.setFixedHeight(48)

        h = QHBoxLayout(self.header_frame)
        h.setContentsMargins(18, 0, 18, 0)
        h.setSpacing(0)

        logo = QLabel("🎵")
        logo.setStyleSheet("font-size: 22px; padding-right: 8px; background: transparent;")
        h.addWidget(logo)

        self.hdr_title = QLabel("詩歌冊搜索")
        self.hdr_title.setStyleSheet(
            "font-size: 17px; font-weight: bold; padding-right: 5px; background: transparent;"
        )
        h.addWidget(self.hdr_title)

        bd = f" · {BUILD_DATE}" if BUILD_DATE else ""
        self.hdr_app_ver = QLabel(f"v{APP_VERSION}{bd}")
        h.addWidget(self.hdr_app_ver)

        h.addStretch()

        self.btn_reopen_last = QPushButton("↻  重開上次")
        self.btn_reopen_last.setFixedHeight(30)
        self.btn_reopen_last.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reopen_last.clicked.connect(self._reopen_last_opened)
        h.addWidget(self.btn_reopen_last)

        h.addSpacing(12)

        self.header_opts = QWidget()
        opts = QHBoxLayout(self.header_opts)
        opts.setContentsMargins(0, 0, 0, 0)
        opts.setSpacing(18)
        self.chk_remote_accept_hdr = QCheckBox("接受請求")
        self.chk_remote_accept_hdr.setChecked(self._settings['remote_accept'])
        self.chk_remote_accept_hdr.setEnabled(False)
        self.chk_remote_accept_hdr.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_remote_accept_hdr.toggled.connect(self._on_remote_accept_hdr_toggled)
        opts.addWidget(self.chk_remote_accept_hdr)

        self.combo_remote_policy_hdr = QComboBox()
        self.combo_remote_policy_hdr.addItem("唯一結果自動開", OPEN_POLICY_AUTO)
        self.combo_remote_policy_hdr.addItem("待操作員確認", OPEN_POLICY_CONFIRM)
        self.combo_remote_policy_hdr.addItem("只填 UI 不自動開", OPEN_POLICY_UI)
        self.combo_remote_policy_hdr.setCurrentIndex(
            _remote_policy_index(self._settings['remote_policy'])
        )
        self.combo_remote_policy_hdr.setFixedHeight(30)
        self.combo_remote_policy_hdr.setMinimumWidth(148)
        self.combo_remote_policy_hdr.setEnabled(False)
        self.combo_remote_policy_hdr.currentIndexChanged.connect(self._on_remote_policy_changed)
        opts.addWidget(self.combo_remote_policy_hdr)

        self.chk_click_to_open_hdr = QCheckBox("一點即開")
        self.chk_click_to_open_hdr.setChecked(self._settings['click_to_open'])
        self.chk_click_to_open_hdr.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_click_to_open_hdr.toggled.connect(self._on_click_to_open_hdr_toggled)
        opts.addWidget(self.chk_click_to_open_hdr)

        self.chk_show_qr_hdr = QCheckBox("Show QR Code")
        self.chk_show_qr_hdr.setChecked(self._settings['show_qr_code'])
        self.chk_show_qr_hdr.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_show_qr_hdr.setToolTip("在桌面顯示會眾頁 QR Code（位置見設定）")
        self.chk_show_qr_hdr.toggled.connect(self._on_show_qr_code_toggled)
        opts.addWidget(self.chk_show_qr_hdr)
        h.addWidget(self.header_opts)

        h.addSpacing(14)

        self.btn_settings = QToolButton()
        self.btn_settings.setObjectName('hdrSettingsBtn')
        self.btn_settings.setText("⚙  設定")
        self.btn_settings.setCheckable(True)
        self.btn_settings.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        self.btn_settings.setFixedHeight(30)
        self.btn_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_settings.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_settings.toggled.connect(self._on_settings_toggled)
        h.addWidget(self.btn_settings)

        self.status_lbl = QLabel("◌  掃描中...")
        self.status_lbl.hide()

        return self.header_frame

    def _build_footer_bar(self):
        self.footer_frame = QFrame()
        self.footer_frame.setFixedHeight(30)
        fl = QHBoxLayout(self.footer_frame)
        fl.setContentsMargins(14, 0, 14, 0)
        fl.setSpacing(10)

        self.footer_status_lbl = QLabel("◌  掃描中...")
        fl.addWidget(self.footer_status_lbl, 2)

        self.footer_last_lbl = QLabel("上次：—")
        fl.addWidget(self.footer_last_lbl, 1)

        self.footer_index_lbl = QLabel("索引—")
        self.footer_index_lbl.setFixedWidth(56)
        self.footer_index_lbl.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        fl.addWidget(self.footer_index_lbl)

        self.footer_pending_lbl = QLabel("")
        self.footer_pending_lbl.setFixedWidth(72)
        self.footer_pending_lbl.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        fl.addWidget(self.footer_pending_lbl)

        self.footer_remote_lbl = QLabel("遙控：關閉")
        self.footer_remote_lbl.setFixedWidth(88)
        self.footer_remote_lbl.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        fl.addWidget(self.footer_remote_lbl)

        return self.footer_frame

    # ─────────────────────────────────────────────────────────────
    #  SIDEBAR — 左側書冊列表（固定 240px，不可拖曳調整）
    # ─────────────────────────────────────────────────────────────
    def _build_sidebar(self):
        self.sidebar_frame = QFrame()
        self.sidebar_frame.setFixedWidth(240)
        lay = QVBoxLayout(self.sidebar_frame)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Title row
        self.sidebar_hdr = QWidget()
        self.sidebar_hdr.setObjectName('sidebarHeader')
        self.sidebar_hdr.setFixedHeight(46)
        hl = QHBoxLayout(self.sidebar_hdr)
        hl.setContentsMargins(14, 0, 14, 0)
        hl.setSpacing(8)
        self.sb_title = QLabel("📁  詩歌冊")
        hl.addWidget(self.sb_title)
        hl.addStretch()
        self.book_count_lbl = QLabel("0 本")
        hl.addWidget(self.book_count_lbl)
        lay.addWidget(self.sidebar_hdr)

        # Book list
        self.book_list = QListWidget()
        self.book_list.itemClicked.connect(self._sidebar_click)
        self.book_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.book_list.customContextMenuRequested.connect(self._book_list_context_menu)
        lay.addWidget(self.book_list, 1)

        # Legend
        leg_widget = QWidget()
        self.legend_widget = leg_widget
        leg_widget.setObjectName('sidebarLegend')
        ll = QVBoxLayout(leg_widget)
        ll.setContentsMargins(14, 10, 14, 12)
        ll.setSpacing(4)
        self.legend_labels = []
        for txt in ["W  =  Word 文件", "P  =  PDF 文件", "🔖 = PDF + 書籤索引"]:
            lbl = QLabel(txt)
            self.legend_labels.append(lbl)
            ll.addWidget(lbl)
        lay.addWidget(leg_widget)

        return self.sidebar_frame

    # ─────────────────────────────────────────────────────────────
    #  CONTENT — 右側主區：搜尋列 + 分隔線 + 檔案/結果面板
    # ─────────────────────────────────────────────────────────────
    def _build_content(self):
        panel = QWidget()
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        lay.addWidget(self._build_search_area())

        self.content_sep = HSep()
        lay.addWidget(self.content_sep)

        self.remote_pending_bar = QFrame()
        self.remote_pending_bar.setVisible(False)
        pbl = QHBoxLayout(self.remote_pending_bar)
        pbl.setContentsMargins(20, 8, 20, 8)
        self.remote_pending_lbl = QLabel("")
        self.btn_remote_approve = QPushButton("開啟")
        self.btn_remote_approve.setFixedWidth(72)
        self.btn_remote_approve.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_remote_approve.clicked.connect(self._approve_remote_pending)
        self.btn_remote_reject = QPushButton("忽略")
        self.btn_remote_reject.setFixedWidth(72)
        self.btn_remote_reject.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_remote_reject.clicked.connect(self._reject_remote_pending)
        pbl.addWidget(self.remote_pending_lbl, 1)
        pbl.addWidget(self.btn_remote_approve)
        pbl.addWidget(self.btn_remote_reject)
        lay.addWidget(self.remote_pending_bar)

        lay.addWidget(self._build_file_area(), 1)
        return panel

    # ─────────────────────────────────────────────────────────────
    #  SEARCH AREA — 模式切換、搜尋輸入、設定面板
    # ─────────────────────────────────────────────────────────────
    def _build_search_area(self):
        self.search_frame = QFrame()
        lay = QVBoxLayout(self.search_frame)
        lay.setContentsMargins(20, 16, 20, 14)
        lay.setSpacing(12)

        # ── 模式切換：書本 / 全局 / 關鍵字 ─────────────────────
        mode_row = QHBoxLayout()
        mode_row.setSpacing(8)
        self.btn_book     = QPushButton("📖  書本模式 (F1)")
        self.btn_global   = QPushButton("🌐  全局模式 (F2)")
        self.btn_keyword  = QPushButton("🔤  關鍵字模式 (F3)")
        self.btn_schedule = QPushButton("📋  排程模式 (F4)")
        for btn in (self.btn_book, self.btn_global, self.btn_keyword, self.btn_schedule):
            btn.setCheckable(True)
            btn.setFixedHeight(30)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_book.clicked.connect(lambda: self._set_search_mode('book'))
        self.btn_global.clicked.connect(lambda: self._set_search_mode('global'))
        self.btn_keyword.clicked.connect(lambda: self._set_search_mode('keyword'))
        self.btn_schedule.clicked.connect(lambda: self._set_search_mode('schedule'))
        mode_row.addWidget(self.btn_book)
        mode_row.addWidget(self.btn_global)
        mode_row.addWidget(self.btn_keyword)
        mode_row.addWidget(self.btn_schedule)
        mode_row.addStretch()
        lay.addLayout(mode_row)

        # ── 書本模式：書冊 + 詩歌號/名 + Filter/Open 按鈕（上下兩行）──
        self.book_row = QWidget()
        br = QVBoxLayout(self.book_row)
        br.setContentsMargins(0, 0, 0, 0)
        br.setSpacing(8)

        self.inp_book = QLineEdit()
        self.inp_book.setFixedHeight(40)
        self.inp_book.setPlaceholderText("請按 Ctrl+Enter 開始輸入書冊名稱...")
        self.inp_book.textChanged.connect(self._on_book_text_changed)
        self.inp_book.returnPressed.connect(self._on_book_filter_action)
        self.inp_book.installEventFilter(self)

        book_action_row = QWidget()
        bar = QHBoxLayout(book_action_row)
        bar.setContentsMargins(0, 0, 0, 0)
        bar.setSpacing(8)

        self.btn_sess_prev = QPushButton("◀")
        self.btn_sess_prev.setFixedSize(32, 40)
        self.btn_sess_prev.setToolTip("上一首（Ctrl+←）")
        self.btn_sess_prev.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sess_prev.clicked.connect(self._open_previous_session)

        self.btn_sess_next = QPushButton("▶")
        self.btn_sess_next.setFixedSize(32, 40)
        self.btn_sess_next.setToolTip("下一首（Ctrl+→）")
        self.btn_sess_next.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sess_next.clicked.connect(self._open_next_session)

        self.inp_num = QLineEdit()
        self.inp_num.setFixedHeight(40)
        self.inp_num.setPlaceholderText("詩歌號/名")
        self.inp_num.textChanged.connect(self._apply_sidebar_filter)
        self.inp_num.returnPressed.connect(self._on_book_filter_action)

        self.btn_filter = QPushButton("Filter")
        self.btn_filter.setFixedHeight(40)
        self.btn_filter.setFixedWidth(84)
        self.btn_filter.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_filter.clicked.connect(self._on_book_filter_action)

        bar.addWidget(self.btn_sess_prev)
        bar.addWidget(self.btn_sess_next)
        bar.addWidget(self.inp_num, 1)
        bar.addWidget(self.btn_filter)

        br.addWidget(self.inp_book)
        br.addWidget(book_action_row)
        lay.addWidget(self.book_row)

        # ── 排程模式：增減清單、上下首（預設隱藏）──────────────
        self.schedule_row = QWidget()
        sr = QVBoxLayout(self.schedule_row)
        sr.setContentsMargins(0, 0, 0, 0)
        sr.setSpacing(8)

        self.combo_sched_book = QComboBox()
        self.combo_sched_book.setFixedHeight(40)
        self.combo_sched_book.setCursor(Qt.CursorShape.PointingHandCursor)
        self.combo_sched_book.currentIndexChanged.connect(self._on_sched_book_changed)

        sched_book_row = QWidget()
        sbr = QHBoxLayout(sched_book_row)
        sbr.setContentsMargins(0, 0, 0, 0)
        sbr.setSpacing(8)
        sbr.addWidget(self.combo_sched_book, 1)
        self.lbl_sched_book_hint = QLabel("或從左側揀書")
        sbr.addWidget(self.lbl_sched_book_hint)

        sched_action_row = QWidget()
        sar = QHBoxLayout(sched_action_row)
        sar.setContentsMargins(0, 0, 0, 0)
        sar.setSpacing(8)

        self.btn_sched_prev = QPushButton("◀")
        self.btn_sched_prev.setFixedSize(32, 40)
        self.btn_sched_prev.setToolTip("上一首（Alt+←）")
        self.btn_sched_prev.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_prev.clicked.connect(self._setlist_prev)

        self.btn_sched_next = QPushButton("▶")
        self.btn_sched_next.setFixedSize(32, 40)
        self.btn_sched_next.setToolTip("下一首（Alt+→）")
        self.btn_sched_next.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_next.clicked.connect(self._setlist_next)

        self.inp_sched_num = QLineEdit()
        self.inp_sched_num.setFixedHeight(40)
        self.inp_sched_num.setPlaceholderText("詩歌號/名 — 輸入時顯示建議")
        self.inp_sched_num.textChanged.connect(self._on_sched_num_changed)
        self.inp_sched_num.returnPressed.connect(self._schedule_add)
        self.inp_sched_num.installEventFilter(self)

        self.btn_sched_add = QPushButton("加入")
        self.btn_sched_add.setFixedSize(64, 40)
        self.btn_sched_add.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_add.clicked.connect(self._schedule_add)

        self.btn_sched_open = QPushButton("開啟")
        self.btn_sched_open.setFixedSize(64, 40)
        self.btn_sched_open.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_open.clicked.connect(self._schedule_open_current)

        sar.addWidget(self.btn_sched_prev)
        sar.addWidget(self.btn_sched_next)
        sar.addWidget(self.inp_sched_num, 1)
        sar.addWidget(self.btn_sched_add)
        sar.addWidget(self.btn_sched_open)

        sched_tool_row = QWidget()
        str_lay = QHBoxLayout(sched_tool_row)
        str_lay.setContentsMargins(0, 0, 0, 0)
        str_lay.setSpacing(8)

        self.btn_sched_remove = QPushButton("移除")
        self.btn_sched_remove.setFixedHeight(32)
        self.btn_sched_remove.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_remove.clicked.connect(self._schedule_remove)

        self.btn_sched_up = QPushButton("上移")
        self.btn_sched_up.setFixedHeight(32)
        self.btn_sched_up.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_up.clicked.connect(lambda: self._schedule_move(-1))

        self.btn_sched_down = QPushButton("下移")
        self.btn_sched_down.setFixedHeight(32)
        self.btn_sched_down.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_down.clicked.connect(lambda: self._schedule_move(1))

        self.btn_sched_clear = QPushButton("清空")
        self.btn_sched_clear.setFixedHeight(32)
        self.btn_sched_clear.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sched_clear.clicked.connect(self._schedule_clear)

        self.lbl_sched_info = QLabel("排程：空")
        self.lbl_sched_info.setWordWrap(True)

        str_lay.addWidget(self.btn_sched_remove)
        str_lay.addWidget(self.btn_sched_up)
        str_lay.addWidget(self.btn_sched_down)
        str_lay.addWidget(self.btn_sched_clear)
        str_lay.addWidget(self.lbl_sched_info, 1)

        sr.addWidget(sched_book_row)
        sr.addWidget(sched_action_row)
        sr.addWidget(sched_tool_row)
        self.schedule_row.setVisible(False)
        lay.addWidget(self.schedule_row)

        # ── 全局模式：跨書冊搜索書籤與檔名（預設隱藏）──────────
        self.global_row = QWidget()
        gr = QHBoxLayout(self.global_row)
        gr.setContentsMargins(0, 0, 0, 0)
        gr.setSpacing(8)

        self.inp_global = QLineEdit()
        self.inp_global.setFixedHeight(40)
        self.inp_global.setPlaceholderText("🔍  輸入歌名  e.g. 奇異恩典  —  搜索所有書冊")
        self.inp_global.textChanged.connect(self._on_global_search)
        self.inp_global.returnPressed.connect(self._on_global_search)

        self.btn_global_search = QPushButton("搜索")
        self.btn_global_search.setFixedHeight(40)
        self.btn_global_search.setFixedWidth(84)
        self.btn_global_search.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_global_search.clicked.connect(self._on_global_search)

        gr.addWidget(self.inp_global, 1)
        gr.addWidget(self.btn_global_search)
        self.global_row.setVisible(False)
        lay.addWidget(self.global_row)

        # ── 關鍵字模式：全文搜索 + 即時搜索開關（預設隱藏）──────
        self.keyword_row = QWidget()
        kr = QHBoxLayout(self.keyword_row)
        kr.setContentsMargins(0, 0, 0, 0)
        kr.setSpacing(8)

        self.inp_keyword = QLineEdit()
        self.inp_keyword.setFixedHeight(40)
        self.inp_keyword.setPlaceholderText("🔍  輸入關鍵字  e.g. 恩典  —  搜索 PDF / Word 內容")
        self.inp_keyword.returnPressed.connect(self._on_keyword_search_action)
        self.inp_keyword.textChanged.connect(self._on_keyword_text_changed)

        self.chk_instant = QCheckBox("即時搜索")
        self.chk_instant.setChecked(self._settings['keyword_instant'])
        self.chk_instant.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_instant.toggled.connect(self._on_keyword_instant_toggled)

        self.btn_keyword_search = QPushButton("搜索")
        self.btn_keyword_search.setFixedHeight(40)
        self.btn_keyword_search.setFixedWidth(84)
        self.btn_keyword_search.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_keyword_search.clicked.connect(self._on_keyword_search_action)

        kr.addWidget(self.inp_keyword, 1)
        kr.addWidget(self.chk_instant)
        kr.addWidget(self.btn_keyword_search)
        self.keyword_row.setVisible(False)
        lay.addWidget(self.keyword_row)

        # ── 關鍵字模式：建立/刷新索引工具列（預設隱藏）──────────
        self.keyword_index_row = QWidget()
        kir = QHBoxLayout(self.keyword_index_row)
        kir.setContentsMargins(0, 0, 0, 0)
        kir.setSpacing(8)

        self.btn_build_index = QPushButton("建立索引")
        self.btn_build_index.setFixedHeight(32)
        self.btn_build_index.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_build_index.clicked.connect(lambda: self._start_index_build(force=False))

        self.btn_refresh_index = QPushButton("刷新索引")
        self.btn_refresh_index.setFixedHeight(32)
        self.btn_refresh_index.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_refresh_index.clicked.connect(lambda: self._start_index_build(force=True))

        self.index_status_lbl = QLabel("索引：未建立")
        self.index_status_lbl.setWordWrap(True)

        kir.addWidget(self.btn_build_index)
        kir.addWidget(self.btn_refresh_index)
        kir.addWidget(self.index_status_lbl, 1)
        self.keyword_index_row.setVisible(False)
        lay.addWidget(self.keyword_index_row)

        self._keyword_debounce = QTimer()
        self._keyword_debounce.setSingleShot(True)
        self._keyword_debounce.timeout.connect(self._start_keyword_search)

        # ── 目前選中書冊的摘要資訊 ─────────────────────────────
        self.sel_info = QLabel("  ▸  從左側選擇詩歌冊")
        lay.addWidget(self.sel_info)

        # ── 可摺疊設定：外觀 + 開啟 Word/PDF 時行為 ─────────────
        self.settings_frame = QFrame()
        self.settings_frame.setVisible(False)
        settings_lay = QVBoxLayout(self.settings_frame)
        settings_lay.setContentsMargins(2, 4, 2, 0)
        settings_lay.setSpacing(10)

        label_min_w = 80
        grid_margins = (18, 18, 16, 14)

        self.grp_data = QGroupBox("詩歌資料夾")
        data_grid = QGridLayout(self.grp_data)
        data_grid.setContentsMargins(*grid_margins)
        data_grid.setHorizontalSpacing(14)
        data_grid.setVerticalSpacing(12)
        data_grid.setColumnStretch(1, 1)
        data_grid.setColumnMinimumWidth(0, label_min_w)

        self.lbl_hymn_folder = QLabel("路徑")
        self.lbl_hymn_folder.setMinimumWidth(label_min_w)
        self.lbl_hymn_folder.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.inp_hymn_folder = QLineEdit(self.hymn_folder)
        self.inp_hymn_folder.setPlaceholderText(default_hymn_folder())
        self.btn_browse_hymn_folder = QPushButton("瀏覽…")
        self.btn_browse_hymn_folder.setFixedWidth(72)
        self.btn_browse_hymn_folder.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_browse_hymn_folder.clicked.connect(self._browse_hymn_folder)
        folder_row = QHBoxLayout()
        folder_row.setSpacing(8)
        folder_row.addWidget(self.inp_hymn_folder, 1)
        folder_row.addWidget(self.btn_browse_hymn_folder)

        self.btn_apply_hymn_folder = QPushButton("套用並重新掃描")
        self.btn_apply_hymn_folder.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_apply_hymn_folder.clicked.connect(self._apply_hymn_folder)

        data_grid.addWidget(self.lbl_hymn_folder, 0, 0)
        data_grid.addLayout(folder_row, 0, 1)
        data_grid.addWidget(self.btn_apply_hymn_folder, 1, 1, Qt.AlignmentFlag.AlignLeft)

        self.grp_appearance = QGroupBox("外觀")
        appearance_grid = QGridLayout(self.grp_appearance)
        appearance_grid.setContentsMargins(*grid_margins)
        appearance_grid.setHorizontalSpacing(14)
        appearance_grid.setVerticalSpacing(12)
        appearance_grid.setColumnStretch(1, 1)
        appearance_grid.setColumnMinimumWidth(0, label_min_w)

        self.lbl_theme = QLabel("主題")
        self.lbl_theme.setMinimumWidth(label_min_w)
        self.lbl_theme.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(['dark', 'light'])
        self.theme_combo.setCurrentText(self.theme_name)
        self.theme_combo.setFixedWidth(120)
        self.theme_combo.currentTextChanged.connect(self._on_theme_changed)

        self.lbl_font = QLabel("字體大小")
        self.lbl_font.setMinimumWidth(label_min_w)
        self.lbl_font.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        font_row = QHBoxLayout()
        font_row.setSpacing(8)
        self.font_slider = QSlider(Qt.Orientation.Horizontal)
        self.font_slider.setRange(10, 18)
        self.font_slider.setValue(self.font_size)
        self.font_slider.valueChanged.connect(self._on_font_changed)
        self.fs_lbl = QLabel(f"{self.font_size}px")
        self.fs_lbl.setFixedWidth(40)
        self.fs_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        font_row.addWidget(self.font_slider, 1)
        font_row.addWidget(self.fs_lbl)

        appearance_grid.addWidget(self.lbl_theme, 0, 0)
        appearance_grid.addWidget(self.theme_combo, 0, 1, Qt.AlignmentFlag.AlignLeft)
        appearance_grid.addWidget(self.lbl_font, 1, 0)
        appearance_grid.addLayout(font_row, 1, 1)

        self.chk_book_auto_focus_hymn = QCheckBox("唯一書冊結果跳至詩歌號/名")
        self.chk_book_auto_focus_hymn.setChecked(self._settings['book_auto_focus_hymn'])
        self.chk_book_auto_focus_hymn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_book_auto_focus_hymn.toggled.connect(self._on_book_auto_focus_hymn_toggled)
        appearance_grid.addWidget(self.chk_book_auto_focus_hymn, 2, 0, 1, 2)

        self.chk_click_to_open = QCheckBox("一點即開檔案（單擊列表即開啟）")
        self.chk_click_to_open.setChecked(self._settings['click_to_open'])
        self.chk_click_to_open.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_click_to_open.toggled.connect(self._on_click_to_open_settings_toggled)
        appearance_grid.addWidget(self.chk_click_to_open, 3, 0, 1, 2)

        self.chk_show_preview = QCheckBox("顯示檔案預覽")
        self.chk_show_preview.setChecked(self._settings.get('show_preview', True))
        self.chk_show_preview.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_show_preview.toggled.connect(self._on_show_preview_toggled)
        appearance_grid.addWidget(self.chk_show_preview, 4, 0, 1, 2)

        self.chk_auto_rescan = QCheckBox("資料夾變更時自動重新掃描")
        self.chk_auto_rescan.setChecked(self._settings.get('auto_rescan', True))
        self.chk_auto_rescan.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_auto_rescan.toggled.connect(self._on_auto_rescan_toggled)
        appearance_grid.addWidget(self.chk_auto_rescan, 5, 0, 1, 2)

        self.chk_startup_tray = QCheckBox("啟動時最小化到系統匣")
        self.chk_startup_tray.setChecked(self._settings.get('startup_tray', False))
        self.chk_startup_tray.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_startup_tray.toggled.connect(self._on_startup_tray_toggled)
        appearance_grid.addWidget(self.chk_startup_tray, 6, 0, 1, 2)

        self.chk_minimize_tray = QCheckBox("關閉視窗時最小化到系統匣")
        self.chk_minimize_tray.setChecked(self._settings.get('minimize_to_tray', False))
        self.chk_minimize_tray.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_minimize_tray.toggled.connect(self._on_minimize_tray_toggled)
        appearance_grid.addWidget(self.chk_minimize_tray, 7, 0, 1, 2)

        self.grp_open = QGroupBox("開啟 Word / PDF 時")
        open_grid = QGridLayout(self.grp_open)
        open_grid.setContentsMargins(*grid_margins)
        open_grid.setHorizontalSpacing(14)
        open_grid.setVerticalSpacing(12)
        open_grid.setColumnStretch(1, 1)
        open_grid.setColumnMinimumWidth(0, label_min_w)

        self.chk_open_overlay = QCheckBox("顯示置頂書名提示")
        self.chk_open_overlay.setChecked(self._settings['open_overlay'])
        self.chk_open_overlay.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_open_overlay.toggled.connect(self._on_open_overlay_toggled)

        self.lbl_overlay_secs = QLabel("提示顯示")
        self.lbl_overlay_secs.setMinimumWidth(label_min_w)
        self.lbl_overlay_secs.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_overlay_duration = QComboBox()
        for sec in range(1, MAX_OVERLAY_SECS + 1):
            self.combo_overlay_duration.addItem(f"{sec} 秒", sec)
        self.combo_overlay_duration.addItem("一直顯示", OVERLAY_DURATION_ALWAYS)
        self.combo_overlay_duration.setCurrentIndex(
            _overlay_duration_index(self._settings['overlay_duration'])
        )
        self.combo_overlay_duration.setMinimumWidth(108)
        self.combo_overlay_duration.setCursor(Qt.CursorShape.PointingHandCursor)
        self.combo_overlay_duration.currentIndexChanged.connect(self._on_overlay_duration_changed)

        self.lbl_overlay_pos = QLabel("提示位置")
        self.lbl_overlay_pos.setMinimumWidth(label_min_w)
        self.lbl_overlay_pos.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_overlay_pos = QComboBox()
        self.combo_overlay_pos.addItem("右下角", OVERLAY_MODE_CORNER)
        self.combo_overlay_pos.addItem("置中 → 右下角", OVERLAY_MODE_CENTER)
        self.combo_overlay_pos.setCurrentIndex(_overlay_mode_index(self._settings['overlay_mode']))
        self.combo_overlay_pos.setMinimumWidth(140)
        self.combo_overlay_pos.setCursor(Qt.CursorShape.PointingHandCursor)
        self.combo_overlay_pos.currentIndexChanged.connect(self._on_overlay_mode_changed)

        self.chk_display_duplicate = QCheckBox("切換同步畫面")
        self.chk_display_duplicate.setToolTip("開啟 Word / PDF 後，將延伸桌面改為同步（複製）畫面")
        self.chk_display_duplicate.setChecked(self._settings['display_duplicate'])
        self.chk_display_duplicate.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_display_duplicate.toggled.connect(self._on_display_duplicate_toggled)
        if sys.platform != 'win32':
            self.chk_display_duplicate.setEnabled(False)
            self.chk_display_duplicate.setToolTip("僅支援 Windows")

        open_grid.addWidget(self.chk_open_overlay, 0, 0, 1, 2)
        open_grid.addWidget(self.lbl_overlay_secs, 1, 0)
        open_grid.addWidget(self.combo_overlay_duration, 1, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.lbl_overlay_pos, 2, 0)
        open_grid.addWidget(self.combo_overlay_pos, 2, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.chk_display_duplicate, 3, 0, 1, 2)

        self.grp_remote = QGroupBox("Mobile 遙控")
        remote_grid = QGridLayout(self.grp_remote)
        remote_grid.setContentsMargins(*grid_margins)
        remote_grid.setHorizontalSpacing(14)
        remote_grid.setVerticalSpacing(12)
        remote_grid.setColumnStretch(1, 1)
        remote_grid.setColumnMinimumWidth(0, label_min_w)

        self.chk_remote_api = QCheckBox("啟用 API")
        self.chk_remote_api.setChecked(self._settings['remote_api_enabled'])
        self.chk_remote_api.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_remote_api.toggled.connect(self._on_remote_api_toggled)

        self.chk_remote_accept = QCheckBox("接受請求")
        self.chk_remote_accept.setChecked(self._settings['remote_accept'])
        self.chk_remote_accept.setEnabled(False)
        self.chk_remote_accept.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_remote_accept.toggled.connect(self._on_remote_accept_settings_toggled)

        self.lbl_remote_port = QLabel("Port")
        self.lbl_remote_port.setMinimumWidth(label_min_w)
        self.spin_remote_port = QSpinBox()
        self.spin_remote_port.setRange(1024, 65535)
        self.spin_remote_port.setValue(self._settings['remote_port'])
        self.spin_remote_port.setFixedWidth(100)
        self.spin_remote_port.valueChanged.connect(self._on_remote_port_changed)

        self.lbl_remote_token = QLabel("API Token")
        self.lbl_remote_token.setMinimumWidth(label_min_w)
        self.inp_remote_token = QLineEdit(self._settings['remote_token'])
        self.inp_remote_token.setPlaceholderText("留空 = 不驗證")
        self.inp_remote_token.textChanged.connect(self._on_remote_token_changed)

        self.chk_enable_tunnel = QCheckBox("外網（Cloudflare）")
        self.chk_enable_tunnel.setChecked(self.enable_tunnel)
        self.chk_enable_tunnel.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_enable_tunnel.setToolTip(
            "未勾選：僅同 WiFi 本機 IP（QR Code 指向區網網址）。"
            "勾選後才啟用公開 HTTPS 網址（須 cloudflared 服務或快速通道）。"
        )
        self.chk_enable_tunnel.toggled.connect(self._on_tunnel_toggled)

        self.tunnel_mode_combo = QComboBox()
        self.tunnel_mode_combo.addItem("共用 Windows 服務（live 固定網址）", "service")
        self.tunnel_mode_combo.addItem("快速通道（隨機網址，測試用）", "quick")
        mode_index = self.tunnel_mode_combo.findData(self.tunnel_mode)
        if mode_index < 0:
            mode_index = 0
            self.tunnel_mode = 'service'
        self.tunnel_mode_combo.setCurrentIndex(mode_index)
        self.tunnel_mode_combo.currentIndexChanged.connect(self._on_tunnel_mode_changed)

        self.lbl_public_tunnel_url = QLabel("公開網址")
        self.lbl_public_tunnel_url.setMinimumWidth(label_min_w)
        self.inp_named_tunnel_url = QLineEdit(self.named_tunnel_url)
        self.inp_named_tunnel_url.setPlaceholderText(
            "https://live.churchofgodtm.com/hymn_search/"
        )
        self.inp_named_tunnel_url.setToolTip(
            "手機／會眾用嘅 HTTPS 網址（Cloudflare 公開域名）。"
            "Path（如 /hymn_search/）會自動從網址解析；有 path 時請保留尾隨 /。"
        )
        self.inp_named_tunnel_url.textChanged.connect(self._on_named_tunnel_url_changed)

        self.inp_cloudflared_token = QLineEdit(self.cloudflared_token)
        self.inp_cloudflared_token.setPlaceholderText(
            "Cloudflare Install connector token（eyJ...）"
        )
        self.inp_cloudflared_token.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_cloudflared_token.textChanged.connect(self._on_cloudflared_token_changed)

        self.btn_install_cloudflared = QPushButton("安裝 cloudflared 服務（僅首次、無其他 app 時）")
        self.btn_install_cloudflared.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_install_cloudflared.clicked.connect(self._install_cloudflared_service)

        self.lbl_tunnel_cmd_hint = QLabel(
            "除錯指令（以管理員 CMD 執行；token 取自 settings.json）"
        )
        self.lbl_tunnel_cmd_hint.setWordWrap(True)

        (
            self._tunnel_install_row,
            self.inp_tunnel_install_cmd,
            self.btn_copy_install_cmd,
        ) = self._make_tunnel_cmd_row("安裝")
        self.btn_copy_install_cmd.clicked.connect(
            lambda: self._copy_tunnel_command(self.inp_tunnel_install_cmd.text())
        )

        (
            self._tunnel_uninstall_row,
            self.inp_tunnel_uninstall_cmd,
            self.btn_copy_uninstall_cmd,
        ) = self._make_tunnel_cmd_row("卸載")
        self.btn_copy_uninstall_cmd.clicked.connect(
            lambda: self._copy_tunnel_command(self.inp_tunnel_uninstall_cmd.text())
        )

        (
            self._tunnel_restart_row,
            self.inp_tunnel_restart_cmd,
            self.btn_copy_restart_cmd,
        ) = self._make_tunnel_cmd_row("重啟")
        self.btn_copy_restart_cmd.clicked.connect(
            lambda: self._copy_tunnel_command(self.inp_tunnel_restart_cmd.text())
        )

        self.tunnel_help_lbl = QLabel("")
        self.tunnel_help_lbl.setWordWrap(True)

        self.tunnel_status_lbl = QLabel("外網：—")
        self.tunnel_status_lbl.setWordWrap(True)

        self.lbl_local_url = QLabel("本機（WiFi）")
        self.lbl_local_url.setMinimumWidth(label_min_w)
        self.local_url_lbl = QLabel("—")
        self.local_url_lbl.setWordWrap(True)
        self.local_url_lbl.setToolTip("同 WiFi 手機用此網址；未啟用外網時 QR Code 亦指向此處")

        self.lbl_remote_url = QLabel("外網遙控")
        self.lbl_remote_url.setMinimumWidth(label_min_w)
        self.remote_url_lbl = QLabel("—")
        self.remote_url_lbl.setWordWrap(True)
        self.remote_url_lbl.setToolTip("操作員遙控頁（/admin）；須啟用外網")

        self.lbl_viewer_url = QLabel("外網會眾")
        self.lbl_viewer_url.setMinimumWidth(label_min_w)
        self.viewer_url_lbl = QLabel("—")
        self.viewer_url_lbl.setWordWrap(True)
        self.viewer_url_lbl.setToolTip("公開 HTTPS 會眾頁；須啟用外網時 QR 才指向此處")

        remote_grid.addWidget(self.chk_remote_api, 0, 0, 1, 2)
        remote_grid.addWidget(self.chk_remote_accept, 1, 0, 1, 2)
        remote_grid.addWidget(self.lbl_remote_port, 2, 0)
        remote_grid.addWidget(self.spin_remote_port, 2, 1, Qt.AlignmentFlag.AlignLeft)
        remote_grid.addWidget(self.lbl_remote_token, 3, 0)
        remote_grid.addWidget(self.inp_remote_token, 3, 1)
        remote_grid.addWidget(self.chk_enable_tunnel, 4, 0, 1, 2)
        remote_grid.addWidget(self.tunnel_mode_combo, 5, 0, 1, 2)
        remote_grid.addWidget(self.lbl_public_tunnel_url, 6, 0)
        remote_grid.addWidget(self.inp_named_tunnel_url, 6, 1)
        remote_grid.addWidget(self.inp_cloudflared_token, 7, 0, 1, 2)
        remote_grid.addWidget(self.btn_install_cloudflared, 8, 0, 1, 2)
        remote_grid.addWidget(self.lbl_tunnel_cmd_hint, 9, 0, 1, 2)
        remote_grid.addWidget(self._tunnel_install_row, 10, 0, 1, 2)
        remote_grid.addWidget(self._tunnel_uninstall_row, 11, 0, 1, 2)
        remote_grid.addWidget(self._tunnel_restart_row, 12, 0, 1, 2)
        remote_grid.addWidget(self.tunnel_help_lbl, 13, 0, 1, 2)
        remote_grid.addWidget(self.tunnel_status_lbl, 14, 0, 1, 2)
        remote_grid.addWidget(self.lbl_local_url, 15, 0)
        remote_grid.addWidget(self.local_url_lbl, 15, 1)
        remote_grid.addWidget(self.lbl_viewer_url, 16, 0)
        remote_grid.addWidget(self.viewer_url_lbl, 16, 1)
        remote_grid.addWidget(self.lbl_remote_url, 17, 0)
        remote_grid.addWidget(self.remote_url_lbl, 17, 1)

        self.chk_viewer_show_debug = QCheckBox("會眾觀看顯示除錯列")
        self.chk_viewer_show_debug.setChecked(self._settings['viewer_show_debug'])
        self.chk_viewer_show_debug.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_viewer_show_debug.setToolTip(
            "開啟後，手機 /view 頁面底部會顯示 mapping 除錯資訊"
        )
        self.chk_viewer_show_debug.toggled.connect(self._on_viewer_show_debug_toggled)
        remote_grid.addWidget(self.chk_viewer_show_debug, 18, 0, 1, 2)

        self.lbl_qr_x = QLabel("QR X")
        self.lbl_qr_x.setMinimumWidth(label_min_w)
        self.lbl_qr_x.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_qr_x = QSpinBox()
        self.spin_qr_x.setRange(0, 3840)
        self.spin_qr_x.setValue(self._settings['qr_x'])
        self.spin_qr_x.setMinimumWidth(88)
        self.spin_qr_x.setToolTip("桌面 QR Code 左上角 X 座標（相對主螢幕可用區域）")
        self.spin_qr_x.valueChanged.connect(self._on_qr_layout_changed)

        self.lbl_qr_y = QLabel("QR Y")
        self.lbl_qr_y.setMinimumWidth(label_min_w)
        self.lbl_qr_y.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_qr_y = QSpinBox()
        self.spin_qr_y.setRange(0, 3840)
        self.spin_qr_y.setValue(self._settings['qr_y'])
        self.spin_qr_y.setMinimumWidth(88)
        self.spin_qr_y.setToolTip("桌面 QR Code 左上角 Y 座標（相對主螢幕可用區域）")
        self.spin_qr_y.valueChanged.connect(self._on_qr_layout_changed)

        self.lbl_qr_w = QLabel("QR 寬")
        self.lbl_qr_w.setMinimumWidth(label_min_w)
        self.lbl_qr_w.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_qr_w = QSpinBox()
        self.spin_qr_w.setRange(120, 800)
        self.spin_qr_w.setValue(self._settings['qr_width'])
        self.spin_qr_w.setMinimumWidth(88)
        self.spin_qr_w.setToolTip("桌面 QR 卡片寬度（像素）")
        self.spin_qr_w.valueChanged.connect(self._on_qr_layout_changed)

        self.lbl_qr_h = QLabel("QR 高")
        self.lbl_qr_h.setMinimumWidth(label_min_w)
        self.lbl_qr_h.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_qr_h = QSpinBox()
        self.spin_qr_h.setRange(120, 800)
        self.spin_qr_h.setValue(self._settings['qr_height'])
        self.spin_qr_h.setMinimumWidth(88)
        self.spin_qr_h.setToolTip("桌面 QR 卡片高度（像素，含下方文字）")
        self.spin_qr_h.valueChanged.connect(self._on_qr_layout_changed)

        remote_grid.addWidget(self.lbl_qr_x, 19, 0)
        remote_grid.addWidget(self.spin_qr_x, 19, 1, Qt.AlignmentFlag.AlignLeft)
        remote_grid.addWidget(self.lbl_qr_y, 20, 0)
        remote_grid.addWidget(self.spin_qr_y, 20, 1, Qt.AlignmentFlag.AlignLeft)
        remote_grid.addWidget(self.lbl_qr_w, 21, 0)
        remote_grid.addWidget(self.spin_qr_w, 21, 1, Qt.AlignmentFlag.AlignLeft)
        remote_grid.addWidget(self.lbl_qr_h, 22, 0)
        remote_grid.addWidget(self.spin_qr_h, 22, 1, Qt.AlignmentFlag.AlignLeft)

        self.lbl_qr_caption = QLabel("QR 文字")
        self.lbl_qr_caption.setMinimumWidth(label_min_w)
        self.lbl_qr_caption.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.inp_qr_caption = QLineEdit()
        self.inp_qr_caption.setText(self._settings['qr_caption'])
        self.inp_qr_caption.setPlaceholderText(DEFAULT_QR_CAPTION)
        self.inp_qr_caption.setToolTip("桌面 QR Code 下方顯示文字")
        self.inp_qr_caption.textChanged.connect(self._on_qr_appearance_changed)

        self.lbl_qr_bg = QLabel("QR 背景")
        self.lbl_qr_bg.setMinimumWidth(label_min_w)
        self.lbl_qr_bg.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        (
            self.qr_bg_row,
            self.inp_qr_bg_color,
            self.lbl_qr_bg_preview,
            self.btn_pick_qr_bg,
        ) = self._make_qr_color_row(
            self._settings['qr_bg_color'],
            '#ffffff',
            '桌面 QR 卡片背景色',
        )
        self.inp_qr_bg_color.textChanged.connect(self._on_qr_appearance_changed)
        self.btn_pick_qr_bg.clicked.connect(
            lambda: self._pick_qr_color(
                self.inp_qr_bg_color, self.lbl_qr_bg_preview, self._on_qr_appearance_changed,
            )
        )

        self.lbl_qr_text_color = QLabel("QR 字色")
        self.lbl_qr_text_color.setMinimumWidth(label_min_w)
        self.lbl_qr_text_color.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        (
            self.qr_text_color_row,
            self.inp_qr_text_color,
            self.lbl_qr_text_color_preview,
            self.btn_pick_qr_text_color,
        ) = self._make_qr_color_row(
            self._settings.get('qr_text_color') or '',
            '自動',
            'QR 下方文字顏色；留空則依背景自動',
            allow_auto=True,
        )
        self.inp_qr_text_color.textChanged.connect(self._on_qr_appearance_changed)
        self.btn_pick_qr_text_color.clicked.connect(
            lambda: self._pick_qr_color(
                self.inp_qr_text_color,
                self.lbl_qr_text_color_preview,
                self._on_qr_appearance_changed,
            )
        )
        self.btn_qr_text_color_auto = QPushButton("自動")
        self.btn_qr_text_color_auto.setFixedWidth(52)
        self.btn_qr_text_color_auto.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_qr_text_color_auto.setToolTip("依背景自動選擇文字顏色")
        self.btn_qr_text_color_auto.clicked.connect(self._reset_qr_text_color_auto)
        self.qr_text_color_row.layout().addWidget(self.btn_qr_text_color_auto)

        self.lbl_qr_text_size = QLabel("QR 字級")
        self.lbl_qr_text_size.setMinimumWidth(label_min_w)
        self.lbl_qr_text_size.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_qr_text_size = QSpinBox()
        self.spin_qr_text_size.setRange(8, 32)
        self.spin_qr_text_size.setValue(self._settings['qr_text_size'])
        self.spin_qr_text_size.setMinimumWidth(88)
        self.spin_qr_text_size.setToolTip("QR 下方文字大小（px）")
        self.spin_qr_text_size.valueChanged.connect(self._on_qr_appearance_changed)

        self.lbl_qr_bg_opacity = QLabel("QR 背景透明度")
        self.lbl_qr_bg_opacity.setMinimumWidth(label_min_w)
        self.lbl_qr_bg_opacity.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_qr_bg_opacity = QSpinBox()
        self.spin_qr_bg_opacity.setRange(0, 100)
        self.spin_qr_bg_opacity.setSuffix(' %')
        self.spin_qr_bg_opacity.setValue(self._settings['qr_bg_opacity'])
        self.spin_qr_bg_opacity.setMinimumWidth(88)
        self.spin_qr_bg_opacity.setToolTip(
            "QR 卡片背景不透明度（0=全透明；配合背景色可做出半透明黑底）"
        )
        self.spin_qr_bg_opacity.valueChanged.connect(self._on_qr_appearance_changed)

        remote_grid.addWidget(self.lbl_qr_caption, 23, 0)
        remote_grid.addWidget(self.inp_qr_caption, 23, 1)
        remote_grid.addWidget(self.lbl_qr_bg, 24, 0)
        remote_grid.addWidget(self.qr_bg_row, 24, 1)
        remote_grid.addWidget(self.lbl_qr_text_color, 25, 0)
        remote_grid.addWidget(self.qr_text_color_row, 25, 1)
        remote_grid.addWidget(self.lbl_qr_text_size, 26, 0)
        remote_grid.addWidget(self.spin_qr_text_size, 26, 1, Qt.AlignmentFlag.AlignLeft)
        remote_grid.addWidget(self.lbl_qr_bg_opacity, 27, 0)
        remote_grid.addWidget(self.spin_qr_bg_opacity, 27, 1, Qt.AlignmentFlag.AlignLeft)
        self._update_qr_color_previews()
        self._sync_tunnel_options_visibility()

        self.remote_qr_lbl = QLabel("啟用 Mobile API 後顯示 QR")
        self.remote_qr_lbl.setObjectName('remoteQrLabel')
        self.remote_qr_lbl.setFixedSize(140, 140)
        self.remote_qr_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.remote_qr_lbl.setWordWrap(True)

        self.remote_log_list = QListWidget()
        self.remote_log_list.setMaximumHeight(80)

        tab_general = QWidget()
        tab_general_lay = QVBoxLayout(tab_general)
        tab_general_lay.setContentsMargins(0, 8, 0, 0)
        tab_general_lay.setSpacing(10)
        tab_general_lay.addWidget(self.grp_data)
        tab_general_lay.addWidget(self.grp_appearance)
        tab_general_lay.addStretch()

        tab_open = QWidget()
        tab_open_lay = QVBoxLayout(tab_open)
        tab_open_lay.setContentsMargins(0, 8, 0, 0)
        tab_open_lay.addWidget(self.grp_open)
        tab_open_lay.addStretch()

        tab_mobile = QWidget()
        tab_mobile_lay = QVBoxLayout(tab_mobile)
        tab_mobile_lay.setContentsMargins(0, 8, 0, 0)
        tab_mobile_lay.setSpacing(10)
        tab_mobile_lay.addWidget(self.grp_remote)
        tab_mobile_lay.addWidget(self.remote_qr_lbl, 0, Qt.AlignmentFlag.AlignHCenter)
        tab_mobile_lay.addWidget(self.remote_log_list)
        tab_mobile_lay.addStretch()

        tab_about = QWidget()
        tab_about_lay = QVBoxLayout(tab_about)
        tab_about_lay.setContentsMargins(0, 8, 0, 0)
        tab_about_lay.setSpacing(12)
        about_ver = QLabel(f"詩歌冊搜索系統 v{APP_VERSION}" + (f"  ·  {BUILD_DATE}" if BUILD_DATE else ""))
        about_ver.setAlignment(Qt.AlignmentFlag.AlignLeft)
        tab_about_lay.addWidget(about_ver)
        about_design = QLabel("Designed By Darren Ho")
        about_design.setAlignment(Qt.AlignmentFlag.AlignLeft)
        tab_about_lay.addWidget(about_design)
        about_btn_row = QHBoxLayout()
        about_btn_row.setSpacing(8)
        self.btn_export_settings = QPushButton("匯出設定")
        self.btn_export_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_export_settings.clicked.connect(self._export_settings)
        self.btn_import_settings = QPushButton("匯入設定")
        self.btn_import_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_import_settings.clicked.connect(self._import_settings)
        about_btn_row.addWidget(self.btn_export_settings)
        about_btn_row.addWidget(self.btn_import_settings)
        about_btn_row.addStretch()
        tab_about_lay.addLayout(about_btn_row)
        self.settings_credit = QLabel("Designed By Darren Ho")
        self.settings_credit.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        tab_about_lay.addWidget(self.settings_credit)
        tab_about_lay.addStretch()

        self.settings_tabs = QTabWidget()
        self.settings_tabs.addTab(tab_general, "一般")
        self.settings_tabs.addTab(tab_open, "開檔")
        self.settings_tabs.addTab(tab_mobile, "Mobile")
        self.settings_tabs.addTab(tab_about, "關於")
        self._settings_mobile_tab_index = 2
        self.settings_tabs.currentChanged.connect(self._on_settings_tab_changed)

        settings_lay.addWidget(self.settings_tabs)

        self.settings_scroll = QScrollArea()
        self.settings_scroll.setWidgetResizable(True)
        self.settings_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.settings_scroll.setWidget(self.settings_frame)

        return self.search_frame

    # ─────────────────────────────────────────────────────────────
    #  FILE PANEL — 右側檔案/書籤/搜索結果列表
    # ─────────────────────────────────────────────────────────────
    def _build_file_area(self):
        self.file_frame = QFrame()
        lay = QVBoxLayout(self.file_frame)
        lay.setContentsMargins(20, 14, 20, 14)
        lay.setSpacing(10)

        hdr_row = QHBoxLayout()
        self.file_list_header = QWidget()
        flh = QHBoxLayout(self.file_list_header)
        flh.setContentsMargins(0, 0, 0, 0)
        self.file_hdr = QLabel("檔案 / 書籤")
        flh.addWidget(self.file_hdr)
        flh.addStretch()
        self.file_count_lbl = QLabel("")
        flh.addWidget(self.file_count_lbl)
        hdr_row.addWidget(self.file_list_header, 1)
        lay.addLayout(hdr_row)

        self.schedule_panel = QWidget()
        sched_lay = QVBoxLayout(self.schedule_panel)
        sched_lay.setContentsMargins(0, 0, 0, 0)
        sched_lay.setSpacing(8)

        ss_hdr = QHBoxLayout()
        self.sched_search_hdr = QLabel("搜尋結果")
        ss_hdr.addWidget(self.sched_search_hdr)
        ss_hdr.addStretch()
        self.sched_search_count = QLabel("")
        ss_hdr.addWidget(self.sched_search_count)
        sched_lay.addLayout(ss_hdr)

        self.sched_search_list = QListWidget()
        self.sched_search_list.itemDoubleClicked.connect(self._schedule_add_from_search)
        self.sched_search_list.currentItemChanged.connect(self._on_sched_search_selection_changed)
        sched_lay.addWidget(self.sched_search_list, 1)

        sp_hdr = QHBoxLayout()
        self.sched_playlist_hdr = QLabel("播放清單")
        sp_hdr.addWidget(self.sched_playlist_hdr)
        sp_hdr.addStretch()
        self.sched_playlist_count = QLabel("")
        sp_hdr.addWidget(self.sched_playlist_count)
        sched_lay.addLayout(sp_hdr)

        self.sched_playlist = QListWidget()
        self.sched_playlist.itemDoubleClicked.connect(self._schedule_playlist_double_click)
        self.sched_playlist.currentItemChanged.connect(self._on_sched_playlist_selection_changed)
        sched_lay.addWidget(self.sched_playlist, 1)

        self.schedule_panel.setVisible(False)
        lay.addWidget(self.schedule_panel, 1)

        self.preview_lbl = QLabel("")
        self.preview_lbl.setWordWrap(True)
        self.preview_lbl.setMaximumHeight(40)
        lay.addWidget(self.preview_lbl)

        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self._file_clicked)
        self.file_list.itemDoubleClicked.connect(self._file_double_clicked)
        self.file_list.currentItemChanged.connect(self._on_file_selection_changed)
        lay.addWidget(self.file_list, 1)

        self.hint_lbl = QLabel("選一個左邊書冊，右邊會列出該書的所有檔案")
        self.hint_lbl.setWordWrap(True)
        lay.addWidget(self.hint_lbl)

        return self.file_frame

    # ─────────────────────────────────────────────────────────────
    #  TOAST — 底部短暫提示列（建構用，顯示邏輯見 _show_toast）
    # ─────────────────────────────────────────────────────────────
    def _build_toast(self):
        self.toast_bar = QLabel("")
        self.toast_bar.setFixedHeight(0)
        self.toast_bar.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        self._toast_timer = QTimer()
        self._toast_timer.setSingleShot(True)
        self._toast_timer.timeout.connect(lambda: self.toast_bar.setFixedHeight(0))
        return self.toast_bar

    def _book_name_for_path(self, path):
        if not path:
            return ''
        abs_path = os.path.abspath(path)
        for book in self.books:
            for fi in book.get('files', []):
                if os.path.abspath(fi['path']) == abs_path:
                    return book['name']
            for key in ('pdf', 'docx'):
                bp = book.get(key)
                if bp and os.path.abspath(bp) == abs_path:
                    return book['name']
        return os.path.splitext(os.path.basename(path))[0]

    def _show_open_book_popup(self, path, item_name=''):
        if not self.chk_open_overlay.isChecked():
            return
        book_name = self._book_name_for_path(path)
        if not book_name:
            return
        BookNameOverlay.show_book(
            book_name, path,
            _overlay_subtitle(item_name or os.path.basename(path)),
            self.combo_overlay_duration.currentData(),
            self.combo_overlay_pos.currentData(),
        )

    def _on_open_overlay_toggled(self, checked):
        for w in (self.lbl_overlay_secs, self.combo_overlay_duration,
                  self.lbl_overlay_pos, self.combo_overlay_pos):
            w.setEnabled(checked)
        if not checked:
            BookNameOverlay.hide_overlay()
        self._persist_settings(open_overlay=checked)

    def _on_overlay_duration_changed(self, _idx):
        self._persist_settings(
            overlay_duration=self.combo_overlay_duration.currentData()
        )

    def _on_overlay_mode_changed(self, _idx):
        self._persist_settings(
            overlay_mode=self.combo_overlay_pos.currentData()
        )

    def _on_display_duplicate_toggled(self, checked):
        self._persist_settings(display_duplicate=checked)

    def _on_keyword_instant_toggled(self, checked):
        self._persist_settings(keyword_instant=checked)

    def _on_book_auto_focus_hymn_toggled(self, checked):
        self._persist_settings(book_auto_focus_hymn=checked)

    def _sync_click_to_open_checkboxes(self, checked):
        for chk in (self.chk_click_to_open_hdr, self.chk_click_to_open):
            chk.blockSignals(True)
            chk.setChecked(checked)
            chk.blockSignals(False)

    def _on_click_to_open_toggled(self, checked):
        self._sync_click_to_open_checkboxes(checked)
        self._persist_settings(click_to_open=checked)

    def _on_click_to_open_hdr_toggled(self, checked):
        self._sync_click_to_open_checkboxes(checked)
        self._on_click_to_open_toggled(checked)

    def _on_click_to_open_settings_toggled(self, checked):
        self._sync_click_to_open_checkboxes(checked)
        self._on_click_to_open_toggled(checked)

    def _make_qr_color_row(self, initial_hex, placeholder, tooltip, allow_auto=False):
        row = QWidget()
        lay = QHBoxLayout(row)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        inp = QLineEdit()
        inp.setText(str(initial_hex or ''))
        inp.setPlaceholderText(placeholder)
        inp.setToolTip(tooltip)
        preview = QLabel()
        preview.setFixedSize(22, 22)
        btn = QPushButton("選色")
        btn.setFixedWidth(52)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        lay.addWidget(inp, 1)
        lay.addWidget(preview)
        lay.addWidget(btn)
        return row, inp, preview, btn

    def _update_qr_color_preview(self, preview_lbl, hex_color, fallback='#ffffff'):
        raw = str(hex_color or '').strip()
        if not raw:
            preview_lbl.setStyleSheet(
                "background: transparent; border: 1px dashed #8892b8; border-radius: 4px;"
            )
            preview_lbl.setToolTip('自動')
            return
        color = _normalize_hex_color(raw, fallback)
        border = _border_color_for_bg(color).name()
        preview_lbl.setStyleSheet(
            f'background: {color}; border: 1px solid {border}; border-radius: 4px;'
        )
        preview_lbl.setToolTip(color)

    def _update_qr_color_previews(self):
        if hasattr(self, 'lbl_qr_bg_preview'):
            self._update_qr_color_preview(
                self.lbl_qr_bg_preview,
                self.inp_qr_bg_color.text() if hasattr(self, 'inp_qr_bg_color') else '',
                DEFAULT_QR_BG_COLOR,
            )
        if hasattr(self, 'lbl_qr_text_color_preview'):
            self._update_qr_color_preview(
                self.lbl_qr_text_color_preview,
                self.inp_qr_text_color.text() if hasattr(self, 'inp_qr_text_color') else '',
                '#333333',
            )

    def _pick_qr_color(self, inp, preview_lbl, callback):
        current = str(inp.text() or '').strip()
        initial = QColor(_normalize_hex_color(current, '#ffffff'))
        color = QColorDialog.getColor(initial, self, '選擇顏色')
        if not color.isValid():
            return
        inp.blockSignals(True)
        inp.setText(color.name())
        inp.blockSignals(False)
        self._update_qr_color_previews()
        callback()

    def _reset_qr_text_color_auto(self):
        if not hasattr(self, 'inp_qr_text_color'):
            return
        self.inp_qr_text_color.blockSignals(True)
        self.inp_qr_text_color.clear()
        self.inp_qr_text_color.blockSignals(False)
        self._update_qr_color_previews()
        self._on_qr_appearance_changed()

    def _qr_appearance_values(self):
        text_color = self.inp_qr_text_color.text().strip() if hasattr(self, 'inp_qr_text_color') else ''
        return {
            'qr_caption': self.inp_qr_caption.text().strip(),
            'qr_bg_color': _normalize_hex_color(
                self.inp_qr_bg_color.text(), DEFAULT_QR_BG_COLOR,
            ),
            'qr_text_color': (
                _normalize_hex_color(text_color, DEFAULT_QR_TEXT_COLOR)
                if text_color else DEFAULT_QR_TEXT_COLOR
            ),
            'qr_text_size': self.spin_qr_text_size.value(),
            'qr_bg_opacity': self.spin_qr_bg_opacity.value(),
        }

    def _on_show_qr_code_toggled(self, checked):
        self._persist_settings(show_qr_code=checked)
        self._sync_desktop_qr_overlay()

    def _on_qr_layout_changed(self, _value):
        x = self.spin_qr_x.value()
        y = self.spin_qr_y.value()
        w = self.spin_qr_w.value()
        h = self.spin_qr_h.value()
        self._persist_settings(qr_x=x, qr_y=y, qr_width=w, qr_height=h)
        self._sync_desktop_qr_overlay()

    def _update_qr_bg_preview(self):
        self._update_qr_color_previews()

    def _on_qr_appearance_changed(self, _value=''):
        values = self._qr_appearance_values()
        self._update_qr_color_previews()
        self._persist_settings(**values)
        self._sync_desktop_qr_overlay()

    def _sync_desktop_qr_overlay(self):
        if not self._settings.get('show_qr_code'):
            DesktopQrOverlay.hide()
            return
        if not getattr(self._remote, 'running', False):
            DesktopQrOverlay.hide()
            return
        url = self._remote_share_url()
        if not url:
            DesktopQrOverlay.hide()
            return
        try:
            style = _desktop_qr_style_from_settings(self._settings)
            caption = str(self._settings.get('qr_caption', DEFAULT_QR_CAPTION))
            card_w = int(self._settings.get('qr_width', DEFAULT_QR_WIDTH))
            card_h = int(self._settings.get('qr_height', DEFAULT_QR_HEIGHT))
            qr_size = _desktop_qr_render_size(
                card_w, card_h, bool(caption.strip()), style['text_size'],
            )
            pix = self._qr_pixmap_from_url(url, size=qr_size, border=1)
            DesktopQrOverlay.show_qr(
                pix,
                caption,
                style,
                int(self._settings.get('qr_x', DEFAULT_QR_X)),
                int(self._settings.get('qr_y', DEFAULT_QR_Y)),
                card_w,
                card_h,
            )
        except Exception:
            DesktopQrOverlay.hide()

    def _setlist_prev_shortcut(self):
        if self.search_mode != 'schedule':
            self._set_search_mode('schedule')
        self._setlist_prev()

    def _setlist_next_shortcut(self):
        if self.search_mode != 'schedule':
            self._set_search_mode('schedule')
        self._setlist_next()

    @pyqtSlot(object)
    def _run_main_invoke(self, fn):
        fn()

    def _invoke_on_main(self, fn):
        """Run callable on the Qt GUI thread; block if called from API/worker thread."""
        app = QApplication.instance()
        if app is None or QThread.currentThread() == app.thread():
            return fn()
        box = {'result': None, 'error': None}
        done = threading.Event()

        def wrapped():
            try:
                box['result'] = fn()
            except Exception as exc:
                box['error'] = exc
            finally:
                done.set()

        self._main_invoke.emit(wrapped)
        if not done.wait(timeout=120):
            return False
        if box['error'] is not None:
            raise box['error']
        return box['result']

    def _remote_prepare_setlist_next(self):
        def action():
            if self.search_mode != 'schedule':
                self._set_search_mode('schedule')
            info = self._get_setlist_info()
            if not info['entries']:
                return None
            self._settings = advance_setlist_index(self._settings, 1)
            self._persist_settings(setlist_index=self._settings['setlist_index'])
            idx = self._settings['setlist_index']
            entries = parse_setlist(self._settings.get('setlist'))
            if not (0 <= idx < len(entries)):
                return None
            self._refresh_schedule_list(select_index=idx)
            return dict(entries[idx])
        return self._invoke_on_main(action)

    def _remote_prepare_setlist_open(self, index):
        def action():
            if self.search_mode != 'schedule':
                self._set_search_mode('schedule')
            info = self._get_setlist_info()
            entries = info['entries']
            if not entries:
                return None
            idx = max(0, min(len(entries) - 1, int(index)))
            self._settings['setlist_index'] = idx
            self._persist_settings(setlist_index=idx)
            self._refresh_schedule_list(select_index=idx)
            return dict(entries[idx])
        return self._invoke_on_main(action)

    def _remote_handle_setlist_add(self, book, num):
        def action():
            return self._schedule_add_entry(book, num)
        return self._invoke_on_main(action)

    def _refresh_schedule_list(self, select_index=None):
        self._settings = clamp_setlist(self._settings)
        entries = parse_setlist(self._settings.get('setlist'))
        idx = setlist_index(self._settings)
        if select_index is not None:
            idx = max(-1, min(len(entries) - 1, select_index))
            self._settings['setlist_index'] = idx
        self.sched_playlist.clear()
        for i, entry in enumerate(entries):
            book = entry.get('book', '')
            num = entry.get('num', '')
            text = f"{book} / {num}" if num else book
            prefix = '▶  ' if i == idx else '    '
            item = QListWidgetItem(f"{prefix}{i + 1}.  {text}")
            item.setData(Qt.ItemDataRole.UserRole, {
                'kind': 'schedule',
                'index': i,
                'book': book,
                'num': num,
            })
            self.sched_playlist.addItem(item)
            if i == idx:
                self.sched_playlist.setCurrentItem(item)
        self.sched_playlist_count.setText(f'{len(entries)} 首')
        info = self._get_setlist_info()
        if hasattr(self, 'lbl_sched_info'):
            self.lbl_sched_info.setText(info['label'])
        self._update_footer_status()

    def _rebuild_sched_book_combo(self, select_name=''):
        if not hasattr(self, 'combo_sched_book'):
            return
        name = select_name or self._sched_book_name()
        self.combo_sched_book.blockSignals(True)
        self.combo_sched_book.clear()
        for i, book in enumerate(self.books):
            self.combo_sched_book.addItem(f"{i + 1}.  {book['name']}", book['name'])
        if name:
            j = self.combo_sched_book.findData(name)
            if j >= 0:
                self.combo_sched_book.setCurrentIndex(j)
        elif self.combo_sched_book.count():
            self.combo_sched_book.setCurrentIndex(0)
        self.combo_sched_book.blockSignals(False)

    def _sched_book_name(self):
        if not hasattr(self, 'combo_sched_book'):
            return ''
        data = self.combo_sched_book.currentData()
        if data:
            return str(data)
        return self.combo_sched_book.currentText().strip()

    def _sched_book_ref(self):
        name = self._sched_book_name()
        if name:
            return name
        if self.combo_sched_book.count():
            return str(self.combo_sched_book.currentIndex() + 1)
        return ''

    def _sync_sched_book_combo(self, book):
        if not book or not hasattr(self, 'combo_sched_book'):
            return
        j = self.combo_sched_book.findData(book['name'])
        if j < 0:
            self._rebuild_sched_book_combo(select_name=book['name'])
            j = self.combo_sched_book.findData(book['name'])
        if j >= 0:
            self.combo_sched_book.blockSignals(True)
            self.combo_sched_book.setCurrentIndex(j)
            self.combo_sched_book.blockSignals(False)

    def _on_sched_book_changed(self, _index=-1):
        if self.search_mode != 'schedule' or self._loading_settings:
            return
        self._refresh_schedule_search()

    def _on_sched_num_changed(self, _text=''):
        if self.search_mode != 'schedule':
            return
        self._show_sched_num_suggestions()
        self._refresh_schedule_search()

    def _on_sched_entry_picked(self, value):
        self.inp_sched_num.blockSignals(True)
        self.inp_sched_num.setText(value)
        self.inp_sched_num.blockSignals(False)
        self.sched_entry_dropdown.hide()
        self._refresh_schedule_search()

    def _show_sched_num_suggestions(self):
        book_ref = self._sched_book_ref()
        if not book_ref or not self.inp_sched_num.hasFocus():
            self.sched_entry_dropdown.hide()
            return
        q = self.inp_sched_num.text().strip()
        entries = list_entries_for_book(self.books, book_ref, q)
        if not entries:
            self.sched_entry_dropdown.hide()
            return
        self.sched_entry_dropdown.refresh_entries(entries[:80], self.inp_sched_num)

    def _populate_schedule_search_list(self, matches):
        self.sched_search_list.clear()
        for payload in matches:
            if payload.get('kind') == 'file':
                path = payload['path']
                ext = os.path.splitext(path)[1].lower()
                tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
                item = QListWidgetItem(f'  [{tag}]  {os.path.basename(path)}')
            elif payload.get('kind') == 'bookmark':
                item = QListWidgetItem(
                    f"  ↳  {payload['title']}  ·  P.{payload['page']}"
                )
            else:
                item = QListWidgetItem(f"  {summarize_match(payload)}")
            item.setData(Qt.ItemDataRole.UserRole, payload)
            self.sched_search_list.addItem(item)
        self.sched_search_count.setText(f'{len(matches)} 項')

    def _refresh_schedule_search(self):
        if self.search_mode != 'schedule':
            return
        book_ref = self._sched_book_ref()
        num_ref = self.inp_sched_num.text().strip()
        if not book_ref:
            self.sched_search_list.clear()
            self.sched_search_count.setText('0 項')
            return
        if not num_ref:
            entries = list_entries_for_book(self.books, book_ref, '')
            self.sched_search_list.clear()
            matched = resolve_books(self.books, book_ref)
            book_name = matched[0]['name'] if matched else book_ref
            for e in entries[:120]:
                item = QListWidgetItem(f"  {e.get('label', e.get('value', ''))}")
                item.setData(Qt.ItemDataRole.UserRole, {
                    'kind': 'schedule_candidate',
                    'book': book_name,
                    'num': e.get('value', ''),
                })
                self.sched_search_list.addItem(item)
            self.sched_search_count.setText(f'{min(len(entries), 120)} 項')
            return
        matches = resolve_open_request(self.books, book_ref, num_ref)
        self._populate_schedule_search_list(matches)

    def _on_sched_search_selection_changed(self, current=None, previous=None):
        if self.search_mode != 'schedule':
            return
        if not self._settings.get('show_preview', True):
            self.preview_lbl.clear()
            self.preview_lbl.setToolTip('')
            return
        item = current
        if item is None and hasattr(self, 'sched_search_list'):
            item = self.sched_search_list.currentItem()
        payload = item.data(Qt.ItemDataRole.UserRole) if item else None
        if payload and payload.get('kind') == 'schedule_candidate':
            self.preview_lbl.setText(payload.get('num', '') or payload.get('book', ''))
            return
        text = preview_for_payload(payload) if payload else ''
        lines = preview_lines(text, max_lines=2, max_chars=140)
        display = '\n'.join(lines)
        self.preview_lbl.setText(display)
        self.preview_lbl.setToolTip(text if text and text != display else '')

    def _on_sched_playlist_selection_changed(self, current=None, previous=None):
        if self.search_mode != 'schedule':
            return
        if not self._settings.get('show_preview', True):
            self.preview_lbl.clear()
            self.preview_lbl.setToolTip('')
            return
        item = current or (self.sched_playlist.currentItem() if hasattr(self, 'sched_playlist') else None)
        payload = item.data(Qt.ItemDataRole.UserRole) if item else None
        if payload and payload.get('kind') == 'schedule':
            book = payload.get('book', '')
            num = payload.get('num', '')
            self.preview_lbl.setText(f"{book} / {num}" if num else book)
            return
        self._on_file_selection_changed(current, previous)

    def _schedule_add_entry(self, book, num):
        book = str(book or '').strip()
        num = str(num or '').strip()
        if not book and not num:
            self._show_toast('請選擇書冊並輸入詩歌號', 'warn')
            return False
        self._settings = add_setlist_entry(self._settings, book, num)
        self._settings = clamp_setlist(self._settings)
        idx = len(parse_setlist(self._settings.get('setlist'))) - 1
        self._settings['setlist_index'] = idx
        self._persist_settings(
            setlist=self._settings['setlist'],
            setlist_index=idx,
        )
        self._refresh_schedule_list(select_index=idx)
        self._show_toast('已加入播放清單', 'ok')
        return True

    def _schedule_add(self):
        book = self._sched_book_name() or self._sched_book_ref()
        num = self.inp_sched_num.text().strip()
        if self._schedule_add_entry(book, num):
            self.inp_sched_num.clear()
            self._refresh_schedule_search()

    def _schedule_add_from_search(self, item):
        payload = item.data(Qt.ItemDataRole.UserRole)
        if not payload:
            return
        if payload.get('kind') == 'schedule_candidate':
            book = payload.get('book', '') or self._sched_book_name()
            num = payload.get('num', '')
        else:
            meta = self._payload_open_meta(payload)
            book = meta.get('book') or self._sched_book_name()
            num = meta.get('num') or ''
        if self._schedule_add_entry(book, num):
            self.inp_sched_num.clear()
            self._refresh_schedule_search()

    def _schedule_playlist_double_click(self, item):
        payload = item.data(Qt.ItemDataRole.UserRole)
        if payload and payload.get('kind') == 'schedule':
            self._setlist_go(payload.get('index', 0))

    def _schedule_remove(self):
        row = self.sched_playlist.currentRow()
        if row < 0:
            self._show_toast('請選擇要移除的項目', 'warn')
            return
        self._settings = remove_setlist_at(self._settings, row)
        self._persist_settings(
            setlist=self._settings['setlist'],
            setlist_index=self._settings['setlist_index'],
        )
        self._refresh_schedule_list()
        self._show_toast('已移除', 'ok')

    def _schedule_move(self, delta):
        row = self.file_list.currentRow()
        if row < 0:
            self._show_toast('請選擇要移動的項目', 'warn')
            return
        self._settings = move_setlist_entry(self._settings, row, delta)
        self._persist_settings(
            setlist=self._settings['setlist'],
            setlist_index=self._settings['setlist_index'],
        )
        self._refresh_schedule_list(select_index=self._settings['setlist_index'])

    def _schedule_clear(self):
        entries = parse_setlist(self._settings.get('setlist'))
        if not entries:
            return
        self._settings = clear_setlist(self._settings)
        self._persist_settings(setlist=[], setlist_index=-1)
        self._refresh_schedule_list()
        self._show_toast('排程已清空', 'ok')

    def _schedule_open_current(self):
        row = self.sched_playlist.currentRow()
        if row >= 0:
            self._setlist_go(row)
            return
        self._setlist_go()

    def _sync_footer_index_badge(self):
        if not hasattr(self, 'footer_index_lbl'):
            return
        if not self.content_index:
            self.footer_index_lbl.setText('索引—')
        elif index_is_stale(self.content_index):
            self.footer_index_lbl.setText('索引!')
        else:
            self.footer_index_lbl.setText('索引✓')

    def _update_footer_status(self):
        EnhancementMixin._update_footer_status(self)
        self._sync_footer_index_badge()
        self._update_reopen_last_button()

    def _update_reopen_last_button(self):
        if not hasattr(self, 'btn_reopen_last'):
            return
        history = session_history_list(self._settings)
        has_last = bool(history)
        self.btn_reopen_last.setEnabled(has_last)
        if has_last:
            entry = history[0]
            label = entry.get('label') or entry.get('book') or ''
            self.btn_reopen_last.setToolTip(f'重開上次（不顯示書名提示）：{label}')
        else:
            self.btn_reopen_last.setToolTip('沒有上次開啟記錄')

    def _reopen_last_opened(self):
        history = session_history_list(self._settings)
        if not history:
            self._show_toast('沒有上次開啟記錄', 'warn')
            return
        entry = history[0]
        self._open_hymn_ref(
            entry.get('book', ''),
            entry.get('num', ''),
            record=False,
            show_overlay=False,
            show_toast=False,
        )

    def _persist_settings(self, **updates):
        if self._loading_settings:
            return
        self._settings.update(updates)
        save_settings(self._settings)

    def _restore_saved_settings(self):
        self._loading_settings = True
        s = self._settings
        self._remote.state.set_policy(s['remote_policy'])
        self._remote.state.set_token(s['remote_token'])
        self._remote.state.set_port(s['remote_port'])
        self._remote.state.set_viewer_show_debug(s['viewer_show_debug'])
        self._set_search_mode(s['search_mode'])
        self._on_open_overlay_toggled(s['open_overlay'])
        if s['remote_api_enabled']:
            self.chk_remote_api.setChecked(True)
            self._on_remote_api_toggled(True)
            if s['remote_accept']:
                self._sync_remote_accept_checkboxes(True)
                self._remote.state.set_accepting(True)
        for chk, key in (
            (self.chk_show_preview, 'show_preview'),
            (self.chk_auto_rescan, 'auto_rescan'),
            (self.chk_startup_tray, 'startup_tray'),
            (self.chk_minimize_tray, 'minimize_to_tray'),
            (self.chk_viewer_show_debug, 'viewer_show_debug'),
            (self.chk_show_qr_hdr, 'show_qr_code'),
        ):
            chk.blockSignals(True)
            chk.setChecked(bool(s.get(key)))
            chk.blockSignals(False)
        for spin, key in (
            (self.spin_qr_x, 'qr_x'),
            (self.spin_qr_y, 'qr_y'),
            (self.spin_qr_w, 'qr_width'),
            (self.spin_qr_h, 'qr_height'),
        ):
            spin.blockSignals(True)
            spin.setValue(int(s.get(key, 0)))
            spin.blockSignals(False)
        if hasattr(self, 'inp_qr_caption'):
            self.inp_qr_caption.blockSignals(True)
            self.inp_qr_caption.setText(str(s.get('qr_caption') or DEFAULT_QR_CAPTION))
            self.inp_qr_caption.blockSignals(False)
        if hasattr(self, 'inp_qr_bg_color'):
            self.inp_qr_bg_color.blockSignals(True)
            self.inp_qr_bg_color.setText(
                _normalize_hex_color(s.get('qr_bg_color'), DEFAULT_QR_BG_COLOR)
            )
            self.inp_qr_bg_color.blockSignals(False)
        if hasattr(self, 'inp_qr_text_color'):
            self.inp_qr_text_color.blockSignals(True)
            text_color = str(s.get('qr_text_color') or '').strip()
            self.inp_qr_text_color.setText(
                _normalize_hex_color(text_color, DEFAULT_QR_TEXT_COLOR) if text_color else ''
            )
            self.inp_qr_text_color.blockSignals(False)
        if hasattr(self, 'spin_qr_text_size'):
            self.spin_qr_text_size.blockSignals(True)
            self.spin_qr_text_size.setValue(int(s.get('qr_text_size', DEFAULT_QR_TEXT_SIZE)))
            self.spin_qr_text_size.blockSignals(False)
        if hasattr(self, 'spin_qr_bg_opacity'):
            self.spin_qr_bg_opacity.blockSignals(True)
            self.spin_qr_bg_opacity.setValue(int(s.get('qr_bg_opacity', DEFAULT_QR_BG_OPACITY)))
            self.spin_qr_bg_opacity.blockSignals(False)
        self._update_qr_color_previews()
        self._apply_setlist_from_settings()
        if s.get('remote_api_enabled') and s.get('last_opened'):
            self._sync_viewer_now_playing(label=s['last_opened'])
        self._loading_settings = False
        self._sync_desktop_qr_overlay()

    def _on_settings_toggled(self, checked):
        if not hasattr(self, 'settings_overlay'):
            return
        self.settings_overlay.setVisible(checked)
        self.btn_settings.setText("⚙  收起" if checked else "⚙  設定")
        if checked:
            self._sync_settings_overlay_geometry()
            self.settings_overlay.raise_()
            QTimer.singleShot(0, self._refresh_settings_layout)
            QTimer.singleShot(0, self._update_qr_code)

    def _on_settings_tab_changed(self, index):
        if index == getattr(self, '_settings_mobile_tab_index', 2):
            QTimer.singleShot(0, self._update_qr_code)
            QTimer.singleShot(0, self._sync_tunnel_options_visibility)

    def _browse_hymn_folder(self):
        start = self.inp_hymn_folder.text().strip() or self.hymn_folder
        if not os.path.isdir(start):
            start = os.path.dirname(start) if os.path.dirname(start) else default_hymn_folder()
        path = QFileDialog.getExistingDirectory(self, "選擇詩歌資料夾", start)
        if path:
            self.inp_hymn_folder.setText(path)

    def _apply_hymn_folder(self):
        path = self.inp_hymn_folder.text().strip() or default_hymn_folder()
        path = os.path.abspath(path)
        if path == os.path.abspath(self.hymn_folder):
            self._show_toast('資料夾路徑未變更', 'warn')
            return
        self.hymn_folder = path
        self._persist_settings(hymn_folder=path)
        self._stop_index_worker()
        self._stop_search_worker()
        self.content_index = None
        self.current_book = None
        self._scan_all()
        self._show_toast(f'已切換資料夾：{os.path.basename(path) or path}', 'ok')

    def _refresh_settings_layout(self):
        for w in (self.grp_data, self.grp_appearance, self.grp_open,
                  self.settings_tabs, self.settings_frame):
            w.updateGeometry()
            w.adjustSize()

    def _update_settings_scroll_height(self):
        pass

    def _maybe_duplicate_displays(self):
        if sys.platform != 'win32':
            return
        if self.chk_display_duplicate.isChecked():
            self.display_duplicate_requested.emit()

    def _switch_duplicate_with_overlay_fix(self):
        if sys.platform != 'win32' or not self.chk_display_duplicate.isChecked():
            return
        # Brief pause after Alt+O so Word settles before the topology change.
        QTimer.singleShot(450, self._apply_display_duplicate)

    def _apply_display_duplicate(self):
        if sys.platform != 'win32' or not self.chk_display_duplicate.isChecked():
            return
        BookNameOverlay.prepare_for_display_change()
        switch_windows_display_mode('duplicate')
        QTimer.singleShot(1200, BookNameOverlay.refresh_after_display_change)
        QTimer.singleShot(1200, DesktopQrOverlay.refresh_after_display_change)

    def _open_docx_with_flow(self, path, item_name='', *, show_overlay=True):
        if show_overlay:
            self._show_open_book_popup(path, item_name or os.path.basename(path))
        open_docx(path, after_focus_keys=self._maybe_duplicate_displays)

    def _open_pdf_with_flow(self, path, item_name='', page=None, *, show_overlay=True):
        if show_overlay:
            self._show_open_book_popup(path, item_name or os.path.basename(path))
        if page is not None:
            open_pdf_at_page(path, page)
        elif sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
        # PDF has no Alt+O step; switch after the viewer has had time to open.
        if sys.platform == 'win32' and self.chk_display_duplicate.isChecked():
            QTimer.singleShot(2600, self._switch_duplicate_with_overlay_fix)

    # ─────────────────────────────────────────────────────────────
    #  APPLY THEME — 統一套用深/淺色主題至所有元件
    # ─────────────────────────────────────────────────────────────
    def _apply_theme(self):
        t = THEMES[self.theme_name]
        fs = self.font_size

        # Qt palette
        pal = QPalette()
        pal.setColor(QPalette.ColorRole.Window,          QColor(t['bg']))
        pal.setColor(QPalette.ColorRole.WindowText,      QColor(t['text']))
        pal.setColor(QPalette.ColorRole.Base,            QColor(t['in_bg']))
        pal.setColor(QPalette.ColorRole.AlternateBase,   QColor(t['panel2']))
        pal.setColor(QPalette.ColorRole.Text,            QColor(t['text']))
        pal.setColor(QPalette.ColorRole.Button,          QColor(t['panel']))
        pal.setColor(QPalette.ColorRole.ButtonText,      QColor(t['text']))
        pal.setColor(QPalette.ColorRole.Highlight,       QColor(t['accent']))
        pal.setColor(QPalette.ColorRole.HighlightedText, QColor(t['a_text']))
        self._app.setPalette(pal)

        # Global stylesheet
        self.setStyleSheet(build_app_style(self.theme_name, fs))

        # ── Header ──────────────────────────────────────────────
        self.header_frame.setStyleSheet(
            f"background: {t['panel']}; border-bottom: 1px solid {t['border']};"
        )
        self.hdr_title.setStyleSheet(
            f"font-size: 17px; font-weight: bold; color: {t['text']}; "
            f"padding-right: 5px; background: transparent;"
        )
        self.hdr_app_ver.setStyleSheet(
            f"font-size: 11px; color: {t['muted']}; "
            f"padding-top: 5px; padding-right: 8px; background: transparent;"
        )
        hdr_chk_style = (
            f"QCheckBox {{ font-size: {max(10, fs - 3)}px; color: {t['text2']}; "
            f"background: transparent; spacing: 6px; padding: 0 2px; }}"
        )
        self.chk_click_to_open_hdr.setStyleSheet(hdr_chk_style)
        self.chk_remote_accept_hdr.setStyleSheet(hdr_chk_style)
        self.chk_show_qr_hdr.setStyleSheet(hdr_chk_style)
        hdr_combo_style = (
            f"QComboBox {{ background: {t['in_bg']}; border: 1.5px solid {t['in_bd']}; "
            f"border-radius: 8px; padding: 4px 8px; color: {t['text']}; "
            f"font-size: {max(10, fs - 3)}px; }}"
            f"QComboBox:disabled {{ color: {t['muted']}; }}"
            f"QComboBox::drop-down {{ border: none; width: 18px; }}"
            f"QComboBox QAbstractItemView {{ background: {t['panel']}; color: {t['text']}; "
            f"border: 1px solid {t['border']}; selection-background-color: {t['accent']}; "
            f"selection-color: {t['a_text']}; }}"
        )
        self.combo_remote_policy_hdr.setStyleSheet(hdr_combo_style)

        # ── Footer ───────────────────────────────────────────────
        footer_style = (
            f"QFrame {{ background: {t['panel']}; border-top: 1px solid {t['border']}; }}"
        )
        self.footer_frame.setStyleSheet(footer_style)
        footer_lbl = (
            f"font-size: {max(10, fs - 3)}px; color: {t['muted']}; background: transparent;"
        )
        for lbl in (self.footer_status_lbl, self.footer_last_lbl, self.footer_index_lbl,
                    self.footer_pending_lbl, self.footer_remote_lbl):
            lbl.setStyleSheet(footer_lbl)

        # ── Sidebar ──────────────────────────────────────────────
        self.sidebar_frame.setStyleSheet(
            f"QFrame {{ background: {t['panel']}; border-right: 1px solid {t['border']}; }}"
        )
        self.sidebar_hdr.setStyleSheet(
            f"QWidget#sidebarHeader {{ background: {t['panel']}; "
            f"border: none; border-bottom: 1px solid {t['border']}; }}"
        )
        self.sb_title.setStyleSheet(
            f"font-size: 13px; font-weight: bold; color: {t['text']}; "
            f"background: transparent; border: none; padding: 0;"
        )
        self.book_count_lbl.setStyleSheet(
            f"font-size: 11px; color: {t['text2']}; background: {t['panel2']}; "
            f"padding: 2px 9px; border-radius: 10px; border: none;"
        )
        self.legend_widget.setStyleSheet(
            f"QWidget#sidebarLegend {{ background: {t['panel']}; "
            f"border: none; border-top: 1px solid {t['border']}; }}"
        )
        self.book_list.setStyleSheet(build_book_list_style(self.theme_name, fs))
        for lbl in self.legend_labels:
            lbl.setStyleSheet(f"font-size: 10px; color: {t['muted']}; background: transparent;")

        # ── Search frame ─────────────────────────────────────────
        self.search_frame.setStyleSheet(f"QFrame {{ background: {t['panel']}; }}")

        # Mode toggles
        toggle_style = f"""
            QPushButton {{
                border-radius: 14px; padding: 4px 16px;
                font-size: {fs - 1}px; font-weight: 600;
                border: 1.5px solid {t['in_bd']};
                color: {t['text2']}; background: transparent;
            }}
            QPushButton:checked {{
                background: {t['accent']}; border-color: {t['accent']}; color: {t['a_text']};
            }}
            QPushButton:hover:!checked {{
                border-color: {t['accent_h']}; color: {t['text']};
            }}
        """
        self.btn_book.setStyleSheet(toggle_style)
        self.btn_global.setStyleSheet(toggle_style)
        self.btn_keyword.setStyleSheet(toggle_style)
        self.btn_schedule.setStyleSheet(toggle_style)
        hdr_btn_style = f"""
            QPushButton {{
                border-radius: 14px; padding: 4px 12px;
                font-size: {max(10, fs - 2)}px; font-weight: 600;
                border: 1.5px solid {t['in_bd']};
                color: {t['text2']}; background: transparent;
            }}
            QPushButton:checked {{
                background: {t['accent']}; border-color: {t['accent']}; color: {t['a_text']};
            }}
            QPushButton:hover:enabled:!checked {{
                border-color: {t['accent_h']}; color: {t['text']};
            }}
            QPushButton:disabled {{
                color: {t['muted']}; border-color: {t['border']};
            }}
        """
        self.btn_reopen_last.setStyleSheet(hdr_btn_style)
        self.btn_settings.setStyleSheet(f"""
            QToolButton#hdrSettingsBtn {{
                border: none; border-radius: 0; padding: 4px 8px;
                font-size: {max(10, fs - 2)}px; font-weight: 600;
                color: {t['text2']}; background: transparent;
            }}
            QToolButton#hdrSettingsBtn:hover {{
                color: {t['accent_h']}; background: transparent; border: none;
            }}
            QToolButton#hdrSettingsBtn:checked {{
                color: {t['accent']}; background: transparent; border: none;
            }}
        """)
        if hasattr(self, 'settings_overlay'):
            self.settings_overlay.setStyleSheet(
                f"QFrame#settingsOverlay {{ background: {t['bg']}; "
                f"border-top: 1px solid {t['border']}; }}"
            )
        if hasattr(self, 'settings_overlay_title'):
            self.settings_overlay_title.setStyleSheet(
                f"font-size: {fs + 1}px; font-weight: bold; color: {t['text']}; "
                f"background: transparent;"
            )
        if hasattr(self, 'btn_settings_close'):
            self.btn_settings_close.setStyleSheet(toggle_style)

        group_style = f"""
            QGroupBox {{
                font-size: {fs - 1}px;
                font-weight: 600;
                color: {t['text']};
                border: 1.5px solid {t['border']};
                border-radius: 10px;
                margin-top: 14px;
                padding: 14px 12px 12px 16px;
                background: {t['panel2']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 14px;
                padding: 0 8px;
                color: {t['text2']};
                background: {t['panel2']};
            }}
        """
        self.settings_frame.setStyleSheet("background: transparent;")
        self.settings_scroll.setStyleSheet(
            f"QScrollArea {{ background: transparent; border: none; }}"
            f"QScrollArea > QWidget > QWidget {{ background: transparent; }}"
        )
        self.grp_data.setStyleSheet(group_style)
        self.grp_appearance.setStyleSheet(group_style)
        self.grp_open.setStyleSheet(group_style)
        self.grp_remote.setStyleSheet(group_style)
        self.settings_tabs.setStyleSheet(
            f"QTabWidget::pane {{ border: 1px solid {t['border']}; border-radius: 8px; "
            f"background: {t['panel2']}; top: -1px; }}"
            f"QTabBar::tab {{ background: {t['panel']}; color: {t['text2']}; "
            f"padding: 6px 14px; margin-right: 2px; border: none; "
            f"border-top-left-radius: 6px; border-top-right-radius: 6px; "
            f"font-size: {fs - 2}px; }}"
            f"QTabBar::tab:selected {{ background: {t['panel2']}; color: {t['text']}; "
            f"font-weight: 600; border: none; }}"
            f"QTabBar::tab:hover {{ color: {t['text']}; }}"
        )
        self.settings_credit.setStyleSheet(
            f"font-size: {max(9, fs - 3)}px; color: {t['muted']}; "
            f"padding: 4px 8px 2px; background: transparent;"
        )
        self.remote_qr_lbl.setStyleSheet(
            f"QLabel#remoteQrLabel {{ background: #ffffff; color: {t['muted']}; "
            f"border: 1px solid {t['border']}; border-radius: 8px; "
            f"font-size: {max(9, fs - 3)}px; padding: 4px; }}"
        )
        self.remote_log_list.setStyleSheet(
            f"QListWidget {{ background: {t['in_bg']}; color: {t['text2']}; "
            f"border: 1px solid {t['border']}; border-radius: 6px; "
            f"font-size: {max(10, fs - 3)}px; }}"
        )

        # Action buttons
        action_style = f"""
            QPushButton {{
                background: {t['accent']}; color: {t['a_text']};
                border: none; border-radius: 10px;
                padding: 8px 18px; font-size: {fs - 1}px; font-weight: bold;
            }}
            QPushButton:hover   {{ background: {t['accent_h']}; }}
            QPushButton:pressed {{ background: {t['accent_d']}; }}
        """
        ok_hover = '#4ade80' if self.theme_name == 'dark' else '#22c55e'
        ok_press = '#10b981' if self.theme_name == 'dark' else '#15803d'
        open_style = f"""
            QPushButton {{
                background: {t['ok']}; color: #ffffff;
                border: none; border-radius: 10px;
                padding: 8px 18px; font-size: {fs - 1}px; font-weight: bold;
            }}
            QPushButton:hover   {{ background: {ok_hover}; }}
            QPushButton:pressed {{ background: {ok_press}; }}
        """
        self._action_btn_style = action_style
        self._open_btn_style = open_style
        self.btn_filter.setStyleSheet(action_style)
        self.btn_global_search.setStyleSheet(action_style)
        self.btn_keyword_search.setStyleSheet(action_style)

        index_btn_style = f"""
            QPushButton {{
                background: {t['panel2']}; color: {t['text']};
                border: 1.5px solid {t['in_bd']}; border-radius: 8px;
                padding: 4px 14px; font-size: {fs - 2}px; font-weight: 600;
            }}
            QPushButton:hover {{ border-color: {t['accent_h']}; color: {t['accent_h']}; }}
            QPushButton:disabled {{ color: {t['muted']}; border-color: {t['border']}; }}
        """
        self.btn_build_index.setStyleSheet(index_btn_style)
        self.btn_refresh_index.setStyleSheet(index_btn_style)
        self.btn_browse_hymn_folder.setStyleSheet(index_btn_style)
        self.btn_apply_hymn_folder.setStyleSheet(action_style)
        self.btn_export_settings.setStyleSheet(index_btn_style)
        self.btn_import_settings.setStyleSheet(index_btn_style)
        compact_btn_style = f"""
            QPushButton {{
                background: {t['panel2']}; color: {t['text2']};
                border: 1px solid {t['in_bd']}; border-radius: 6px;
                font-size: {max(10, fs - 3)}px; font-weight: 600; padding: 0;
            }}
            QPushButton:hover {{ border-color: {t['accent_h']}; color: {t['text']}; }}
        """
        for btn in (self.btn_sess_prev, self.btn_sess_next,
                    self.btn_sched_prev, self.btn_sched_next,
                    self.btn_sched_add, self.btn_sched_open,
                    self.btn_sched_remove, self.btn_sched_up,
                    self.btn_sched_down, self.btn_sched_clear):
            btn.setStyleSheet(compact_btn_style)
        self.lbl_sched_info.setStyleSheet(
            f"font-size: {max(10, fs - 2)}px; color: {t['text2']}; background: transparent;"
        )
        self.lbl_sched_book_hint.setStyleSheet(
            f"font-size: {max(10, fs - 3)}px; color: {t['muted']}; background: transparent;"
        )
        list_style = build_file_list_style(self.theme_name, fs)
        self.sched_search_list.setStyleSheet(list_style)
        self.sched_playlist.setStyleSheet(list_style)
        for lbl in (self.sched_search_hdr, self.sched_playlist_hdr,
                    self.sched_search_count, self.sched_playlist_count):
            lbl.setStyleSheet(
                f"font-size: {fs - 1}px; color: {t['text2']}; background: transparent;"
            )
        self.index_status_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
        )
        self.chk_instant.setStyleSheet(
            f"font-size: {fs - 1}px; color: {t['text2']}; background: transparent; spacing: 8px;"
        )
        checkbox_style = (
            f"font-size: {fs - 1}px; color: {t['text2']}; background: transparent; "
            f"spacing: 8px; padding-left: 2px;"
        )
        self.chk_book_auto_focus_hymn.setStyleSheet(checkbox_style)
        self.chk_click_to_open.setStyleSheet(checkbox_style)
        self.chk_open_overlay.setStyleSheet(checkbox_style)
        self.chk_display_duplicate.setStyleSheet(checkbox_style)
        for chk in (self.chk_remote_api, self.chk_remote_accept, self.chk_enable_tunnel,
                    self.chk_viewer_show_debug,
                    self.chk_show_preview,
                    self.chk_auto_rescan, self.chk_startup_tray, self.chk_minimize_tray):
            chk.setStyleSheet(checkbox_style)
        for lbl in (self.lbl_remote_port,
                    self.lbl_remote_token, self.lbl_remote_url, self.lbl_viewer_url,
                    self.lbl_public_tunnel_url):
            lbl.setStyleSheet(
                f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
            )
        self.remote_url_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['accent_h']}; background: transparent;"
        )
        self.viewer_url_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['ok']}; background: transparent;"
        )
        self.tunnel_status_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['text2']}; background: transparent;"
        )
        self.tunnel_help_lbl.setStyleSheet(
            f"font-size: {fs - 3}px; color: {t['muted']}; background: transparent;"
        )
        pending_style = f"""
            QFrame {{ background: {t['warn_bg']}; border-bottom: 1px solid {t['border']}; }}
            QLabel {{ color: {t['warn']}; font-size: {fs - 1}px; background: transparent; }}
            QPushButton {{
                background: {t['panel2']}; color: {t['text']};
                border: 1.5px solid {t['in_bd']}; border-radius: 8px;
                padding: 4px 10px; font-size: {fs - 2}px;
            }}
            QPushButton:hover {{ border-color: {t['accent_h']}; }}
        """
        self.remote_pending_bar.setStyleSheet(pending_style)
        self.lbl_overlay_secs.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
        )
        self.lbl_overlay_pos.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
        )

        # sel_info
        self.sel_info.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['accent_h']}; background: transparent;"
        )

        # Settings labels
        for lbl in (self.lbl_hymn_folder, self.lbl_theme, self.lbl_font, self.lbl_overlay_secs, self.lbl_overlay_pos, self.fs_lbl,
                    self.lbl_remote_port, self.lbl_remote_token, self.lbl_remote_url,
                    self.lbl_viewer_url, self.lbl_local_url, self.lbl_public_tunnel_url,
                    self.lbl_qr_x, self.lbl_qr_y, self.lbl_qr_w, self.lbl_qr_h,
                    self.lbl_qr_caption, self.lbl_qr_bg, self.lbl_qr_text_color,
                    self.lbl_qr_text_size, self.lbl_qr_bg_opacity):
            lbl.setStyleSheet(f"font-size: {fs - 1}px; color: {t['text2']}; background: transparent;")

        # ── Content sep + file area ──────────────────────────────
        self.content_sep.set_theme(self.theme_name)
        self.file_frame.setStyleSheet(f"QFrame {{ background: {t['bg']}; }}")
        self.file_hdr.setStyleSheet(
            f"font-size: {fs + 1}px; font-weight: bold; color: {t['text']}; background: transparent;"
        )
        self.file_count_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
        )
        self.preview_lbl.setStyleSheet(
            f"font-size: {max(10, fs - 2)}px; color: {t['text2']}; "
            f"background: {t['panel2']}; border-radius: 6px; padding: 2px 8px;"
        )
        self.hint_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
        )
        self.file_list.setStyleSheet(build_file_list_style(self.theme_name, fs))

        # Toast default
        self.toast_bar.setStyleSheet(
            f"background: {t['ok_bg']}; color: {t['ok']}; "
            f"font-size: {fs - 1}px; padding: 0 20px; "
            f"border-top: 1px solid {t['border']};"
        )

        # Dropdown
        self.dropdown.set_theme(self.theme_name)
        if hasattr(self, 'sched_entry_dropdown'):
            self.sched_entry_dropdown.set_theme(self.theme_name)

        self._update_qr_code()

    def _scan_all(self):
        t = THEMES[self.theme_name]
        self.books = scan_books(self.hymn_folder)

        if not self.books:
            self.book_list.clear()
            item = QListWidgetItem("  ⚠  找不到資料夾")
            item.setForeground(QColor(t['err']))
            self.book_list.addItem(item)
            self._set_status('err', f"找不到：{os.path.basename(self.hymn_folder)}")
            return

        self._rebuild_book_list()
        self._rebuild_sched_book_combo()

        if self.books:
            self.current_book = self.books[0]
            for i in range(self.book_list.count()):
                item = self.book_list.item(i)
                b = item.data(Qt.ItemDataRole.UserRole)
                if b and b['name'] == self.current_book['name']:
                    self.book_list.setCurrentRow(i)
                    break
            self._show_book_files(self.current_book)
            self._update_sel_info(self.current_book)

        self._set_status('ok', f"Ready · {len(self.books)} 本已索引")
        self._load_content_index()

    def _load_content_index(self):
        self.content_index = load_content_index(self.hymn_folder)
        self._update_index_status_label()

    def _update_index_status_label(self):
        nfiles, nchunks, built_at = index_stats(self.content_index)
        if not self.content_index:
            self.index_status_lbl.setText("索引：未建立（建議先按「建立索引」以加快搜索）")
            self.btn_refresh_index.setEnabled(False)
            self._update_footer_status()
            return
        self.btn_refresh_index.setEnabled(True)
        ts = time.strftime('%Y-%m-%d %H:%M', time.localtime(built_at)) if built_at else '?'
        stale = index_is_stale(self.content_index)
        stale_note = '  ·  ⚠ 檔案已變更，建議刷新' if stale else ''
        self.index_status_lbl.setText(
            f"索引：{nfiles} 個檔案 · {nchunks} 段  ·  建立於 {ts}{stale_note}"
        )
        self._update_footer_status()

    def _stop_index_worker(self):
        if self._index_worker and self._index_worker.isRunning():
            self._index_worker.requestInterruption()
            self._index_worker.wait(500)

    def _start_index_build(self, force=False):
        if self._index_worker and self._index_worker.isRunning():
            return
        if not force and self.content_index and not index_is_stale(self.content_index):
            self._show_toast('索引已存在且為最新', 'warn')
            return

        self._stop_search_worker()
        self._keyword_debounce.stop()
        self.btn_build_index.setEnabled(False)
        self.btn_refresh_index.setEnabled(False)
        self._set_status('warn', '正在建立內容索引…')

        self._index_worker = IndexBuildWorker(self.books, self.hymn_folder)
        self._index_worker.progress.connect(self._on_index_progress)
        self._index_worker.finished.connect(self._on_index_built)
        self._index_worker.start()

    def _on_index_progress(self, msg):
        self._set_status('warn', f'建立索引：{msg}')

    def _on_index_built(self, index):
        self.btn_build_index.setEnabled(True)
        if not index:
            self.btn_refresh_index.setEnabled(bool(self.content_index))
            self._set_status('err', '索引建立已取消')
            self._update_index_status_label()
            return
        self.content_index = index
        try:
            save_content_index(index, self.hymn_folder)
        except Exception as e:
            self._show_toast(f'索引儲存失敗：{e}', 'err')
        nfiles, nchunks, _ = index_stats(index)
        self._update_index_status_label()
        self._set_status('ok', f'索引完成 · {nfiles} 檔 · {nchunks} 段')
        self._show_toast(f'✅  索引完成：{nfiles} 個檔案 · {nchunks} 段')
        if self.search_mode == 'keyword' and self.inp_keyword.text().strip():
            self._start_keyword_search()

    def _set_status(self, kind, text):
        t = THEMES[self.theme_name]
        colours = {
            'ok':   (t['ok_bg'],   t['ok']),
            'warn': (t['warn_bg'], t['warn']),
            'err':  (t['err_bg'],  t['err']),
        }
        bg, fg = colours.get(kind, (t['panel2'], t['text2']))
        self.status_lbl.setText(f"●  {text}")
        self.status_lbl.setStyleSheet(
            f"font-size: 11px; color: {fg}; background: {bg}; "
            f"padding: 4px 14px; border-radius: 12px;"
        )
        self._update_footer_status()
        if hasattr(self, 'footer_status_lbl'):
            self.footer_status_lbl.setStyleSheet(
                f"font-size: {max(10, self.font_size - 3)}px; color: {fg}; "
                f"background: transparent;"
            )

    # ─────────────────────────────────────────────────────────────
    #  SIDEBAR SELECTION — 左側書冊點選與切換
    # ─────────────────────────────────────────────────────────────
    def _sidebar_click(self, item):
        book = item.data(Qt.ItemDataRole.UserRole)
        if book:
            self._select_book(book)

    def _select_book(self, book):
        if self.search_mode == 'schedule':
            self.current_book = book
            self._sync_sched_book_combo(book)
            self._refresh_schedule_search()
            self.sel_info.setText(f"  ▸  排程 · {book['name']}")
            return
        if self.search_mode in ('global', 'keyword'):
            self._set_search_mode('book')
        self.current_book = book
        self.inp_book.blockSignals(True)
        self.inp_book.setText(book['name'])
        self.inp_book.blockSignals(False)
        self._show_book_files(book, self.inp_num.text().strip().lower())
        self._update_sel_info(book)
        self._update_search_action_buttons()

    def _clear_file_keyword(self):
        if not self.inp_num.text():
            return
        self.inp_num.blockSignals(True)
        self.inp_num.clear()
        self.inp_num.blockSignals(False)
        if self.current_book and self.search_mode == 'book':
            self._show_book_files(self.current_book, '')
            self._update_search_action_buttons()

    # ─────────────────────────────────────────────────────────────
    #  FILE LIST — 顯示書冊檔案、書籤、全局/關鍵字搜索結果
    # ─────────────────────────────────────────────────────────────
    def _show_book_files(self, book, query=''):
        self.file_list.clear()
        q = query.strip().lower()
        t = THEMES[self.theme_name]

        def add_section_hdr(text):
            i = QListWidgetItem(f"  {text}")
            i.setFlags(Qt.ItemFlag.ItemIsEnabled)
            i.setForeground(QColor(t['muted']))
            f = i.font()
            f.setBold(True)
            i.setFont(f)
            self.file_list.addItem(i)

        if q:
            targets = resolve_targets(book, q)
            if not targets:
                ni = QListWidgetItem("  ⚠  沒有符合條件的檔案或書籤")
                ni.setForeground(QColor(t['muted']))
                self.file_list.addItem(ni)
            else:
                for payload in targets:
                    if payload.get('kind') == 'file':
                        path = payload['path']
                        ext = os.path.splitext(path)[1].lower()
                        tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
                        item = QListWidgetItem(f"  [{tag}]  {os.path.basename(path)}")
                        item.setData(Qt.ItemDataRole.UserRole, {
                            'kind': 'file', 'path': path,
                        })
                        self.file_list.addItem(item)
                    elif payload.get('kind') == 'bookmark':
                        lvl = payload.get('toc_level')
                        indent = "      " if lvl == 1 else "          "
                        bi = QListWidgetItem(
                            f"{indent}↳  {payload['title']}  ·  P.{payload['page']}"
                        )
                        bi.setData(Qt.ItemDataRole.UserRole, {
                            'kind': 'bookmark',
                            'title': payload['title'],
                            'page': payload['page'],
                            'pdf_path': payload['pdf_path'],
                        })
                        self.file_list.addItem(bi)
            self.file_count_lbl.setText(f"{self.file_list.count()} 項")
            self._update_search_action_buttons()
            return

        count = 0
        for fi in book.get('files', []):
            ext = fi['ext']
            is_pdf = ext == '.pdf'
            tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
            item = QListWidgetItem(f"  [{tag}]  {fi['name']}")
            item.setData(Qt.ItemDataRole.UserRole, {'kind': 'file', 'path': fi['path']})
            self.file_list.addItem(item)
            count += 1

            if is_pdf and book.get('bookmark') and book.get('toc'):
                toc = book['toc']
                if any(lvl == 2 for lvl, _, _ in toc):
                    toc = [(lvl, title, page) for lvl, title, page in toc if lvl == 2]
                if toc:
                    add_section_hdr("    書籤")
                    for lvl, title, page in toc:
                        indent = "      " if lvl == 1 else "          "
                        bi = QListWidgetItem(f"{indent}↳  {title}  ·  P.{page}")
                        bi.setData(Qt.ItemDataRole.UserRole, {
                            'kind': 'bookmark', 'title': title,
                            'page': page, 'pdf_path': fi['path'],
                        })
                        self.file_list.addItem(bi)
                        count += 1

        if count == 0:
            ni = QListWidgetItem("  ⚠  沒有符合條件的檔案或書籤")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)

        self.file_count_lbl.setText(f"{self.file_list.count()} 項")
        self._update_search_action_buttons()

    def _show_global_results(self, query=''):
        self.file_list.clear()
        q = query.strip().lower()
        t = THEMES[self.theme_name]

        def add_section_hdr(text):
            i = QListWidgetItem(f"  {text}")
            i.setFlags(Qt.ItemFlag.ItemIsEnabled)
            i.setForeground(QColor(t['muted']))
            f = i.font()
            f.setBold(True)
            i.setFont(f)
            self.file_list.addItem(i)

        if not q:
            ni = QListWidgetItem("  🔍  輸入歌名以搜索所有書冊")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)
            self.file_count_lbl.setText("")
            self._update_search_action_buttons()
            return

        by_book = {}
        for book in self.books:
            pdf_path = book.get('pdf')
            matches = []

            if book.get('bookmark') and book.get('toc'):
                for lvl, title, page in book['toc']:
                    score = fuzzy_match_score(q, title)
                    if score > 0:
                        matches.append({
                            'kind': 'bookmark', 'title': title, 'page': page,
                            'lvl': lvl, 'pdf_path': pdf_path, 'score': score,
                        })

            for fi in book.get('files', []):
                score = fuzzy_match_score(q, fi['name'])
                if score > 0:
                    ext = fi['ext']
                    tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
                    matches.append({
                        'kind': 'file', 'path': fi['path'], 'name': fi['name'],
                        'tag': tag, 'score': score,
                    })

            if matches:
                matches.sort(key=lambda m: m.get('score', 0), reverse=True)
                by_book[book['name']] = matches

        if not by_book:
            ni = QListWidgetItem(f"  ⚠  找不到「{query.strip()}」")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)
            self.file_count_lbl.setText("0 項")
            self._update_search_action_buttons()
            return

        count = 0
        for book_name in sorted(by_book.keys()):
            matches = by_book[book_name]
            add_section_hdr(f"📖  {book_name}  ({len(matches)} 項)")
            for m in matches:
                if m['kind'] == 'bookmark':
                    indent = "      " if m['lvl'] == 1 else "          "
                    bi = QListWidgetItem(f"{indent}↳  {m['title']}  ·  P.{m['page']}")
                    bi.setData(Qt.ItemDataRole.UserRole, {
                        'kind': 'bookmark', 'title': m['title'],
                        'page': m['page'], 'pdf_path': m['pdf_path'],
                    })
                else:
                    bi = QListWidgetItem(f"      [{m['tag']}]  {m['name']}")
                    bi.setData(Qt.ItemDataRole.UserRole, {
                        'kind': 'file', 'path': m['path'],
                    })
                self.file_list.addItem(bi)
                count += 1

        self.file_count_lbl.setText(f"{count} 項 · {len(by_book)} 本")
        self._update_search_action_buttons()

    def _stop_search_worker(self):
        if self._search_worker and self._search_worker.isRunning():
            self._search_worker.requestInterruption()
            self._search_worker.wait(500)

    def _on_keyword_text_changed(self, _text=''):
        if self.search_mode != 'keyword' or not self.chk_instant.isChecked():
            return
        query = self.inp_keyword.text().strip()
        if not query:
            self._keyword_debounce.stop()
            self._show_keyword_placeholder()
            return
        self._keyword_debounce.start(400)

    def _start_keyword_search(self):
        if self.search_mode != 'keyword':
            return
        query = self.inp_keyword.text().strip()
        if not query:
            self._show_keyword_placeholder()
            return

        self._stop_search_worker()
        t = THEMES[self.theme_name]
        using_index = bool(self.content_index)
        if using_index:
            results = search_content_index(self.content_index, query)
            self._on_keyword_results(results)
            return

        self.file_list.clear()
        ni = QListWidgetItem("  ⏳  正在搜索 PDF / Word 內容，請稍候…")
        ni.setForeground(QColor(t['muted']))
        self.file_list.addItem(ni)
        self.file_count_lbl.setText("搜索中")
        self._set_status('warn', '內容搜索中…（建議建立索引以加快速度）')

        self._search_worker = ContentSearchWorker(query, index=None, books=self.books)
        self._search_worker.progress.connect(self._on_search_progress)
        self._search_worker.finished.connect(self._on_keyword_results)
        self._search_worker.start()

    def _on_search_progress(self, msg):
        self._set_status('warn', f'搜索中：{msg}')

    def _show_keyword_placeholder(self):
        self.file_list.clear()
        t = THEMES[self.theme_name]
        if self.content_index:
            hint = "  🔍  輸入關鍵字搜索（已載入索引，可勾選即時搜索）"
        else:
            hint = "  🔍  輸入關鍵字後按搜索（建議先按「建立索引」）"
        ni = QListWidgetItem(hint)
        ni.setForeground(QColor(t['muted']))
        self.file_list.addItem(ni)
        self.file_count_lbl.setText("")
        self._update_search_action_buttons()

    def _on_keyword_results(self, results):
        if self.search_mode != 'keyword':
            return
        query = self.inp_keyword.text().strip()
        self.file_list.clear()
        t = THEMES[self.theme_name]

        def add_section_hdr(text):
            i = QListWidgetItem(f"  {text}")
            i.setFlags(Qt.ItemFlag.ItemIsEnabled)
            i.setForeground(QColor(t['muted']))
            f = i.font()
            f.setBold(True)
            i.setFont(f)
            self.file_list.addItem(i)

        if not query:
            self._show_keyword_placeholder()
            return

        if not results:
            ni = QListWidgetItem(f"  ⚠  在 PDF / Word 內容中找不到「{query}」")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)
            self.file_count_lbl.setText("0 項")
            self._set_status('ok', '搜索完成 · 0 結果')
            self._update_search_action_buttons()
            return

        by_book = {}
        for entry in results:
            by_book.setdefault(entry['book'], []).append(entry)

        count = 0
        for book_name in sorted(by_book.keys()):
            entries = by_book[book_name]
            book_hits = sum(len(e['hits']) for e in entries)
            add_section_hdr(f"📖  {book_name}  ({book_hits} 處)")
            for entry in entries:
                ext = entry['ext']
                tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
                add_section_hdr(f"    [{tag}]  {entry['file']}")
                is_pdf = ext == '.pdf'
                for hit in entry['hits']:
                    loc = f"P.{hit['page']}" if is_pdf else f"§{hit['page']}"
                    snip = hit.get('snippet', '')
                    label = f"          ↳  {loc}  ·  {snip}" if snip else f"          ↳  {loc}"
                    bi = QListWidgetItem(label)
                    payload = {
                        'kind': 'content_pdf' if is_pdf else 'content_word',
                        'path': entry['path'],
                        'page': hit['page'],
                        'snippet': snip,
                    }
                    bi.setData(Qt.ItemDataRole.UserRole, payload)
                    self.file_list.addItem(bi)
                    count += 1

        self.file_count_lbl.setText(f"{count} 處 · {len(by_book)} 本")
        self._set_status('ok', f'搜索完成 · {count} 處匹配')
        self._update_search_action_buttons()

    def _file_clicked(self, item):
        if not self._settings.get('click_to_open'):
            return
        self._file_open_item(item)

    def _file_double_clicked(self, item):
        self._file_open_item(item)

    def _file_open_item(self, item):
        payload = item.data(Qt.ItemDataRole.UserRole)
        if not payload:
            return
        if payload.get('kind') == 'schedule':
            self._setlist_go(payload.get('index', 0))
            return
        self._open_from_payload(payload)

    def _open_from_payload(self, payload, record=True, *, show_overlay=True, show_toast=True):
        kind = payload.get('kind')
        opened = False
        if kind == 'file':
            path = payload.get('path')
            if path and os.path.exists(path):
                try:
                    ext = os.path.splitext(path)[1].lower()
                    if ext in WORD_EXTS:
                        self._open_docx_with_flow(
                            path, os.path.basename(path), show_overlay=show_overlay,
                        )
                    elif ext == '.pdf':
                        self._open_pdf_with_flow(
                            path, os.path.basename(path), show_overlay=show_overlay,
                        )
                    elif sys.platform == 'win32':
                        os.startfile(path)
                    elif sys.platform == 'darwin':
                        subprocess.Popen(['open', path])
                    else:
                        subprocess.Popen(['xdg-open', path])
                    if show_toast:
                        self._show_toast(f"✅  正在開啟：{os.path.basename(path)}")
                    opened = True
                except Exception as e:
                    if show_toast:
                        self._show_toast(f"❌  無法開啟：{e}", 'err')
        elif kind == 'bookmark':
            try:
                self._open_pdf_with_flow(
                    payload['pdf_path'],
                    f"{payload['title']}  ·  P.{payload['page']}",
                    page=payload['page'],
                    show_overlay=show_overlay,
                )
                if show_toast:
                    self._show_toast(
                        f"✅  開啟 P.{payload['page']}：{payload['title'][:30]}"
                    )
                opened = True
            except Exception as e:
                if show_toast:
                    self._show_toast(f"❌  無法開啟書籤：{e}", 'err')
        elif kind == 'content_pdf':
            try:
                sub = f"P.{payload['page']}"
                snip = (payload.get('snippet') or '').strip()
                if snip:
                    sub += f"  {snip[:40]}"
                self._open_pdf_with_flow(
                    payload['path'], sub, page=payload['page'], show_overlay=show_overlay,
                )
                if show_toast:
                    self._show_toast(f"✅  開啟 P.{payload['page']}")
                opened = True
            except Exception as e:
                if show_toast:
                    self._show_toast(f"❌  無法開啟：{e}", 'err')
        elif kind == 'content_word':
            path = payload.get('path')
            if path and os.path.exists(path):
                try:
                    sub = os.path.splitext(os.path.basename(path))[0]
                    page = payload.get('page')
                    snip = (payload.get('snippet') or '').strip()
                    if page:
                        sub = f"P.{page}  {sub}"
                    if snip:
                        sub += f"  ·  {snip[:35]}"
                    self._open_docx_with_flow(path, sub, show_overlay=show_overlay)
                    if show_toast:
                        self._show_toast(f"✅  正在開啟：{os.path.basename(path)}")
                    opened = True
                except Exception as e:
                    if show_toast:
                        self._show_toast(f"❌  無法開啟：{e}", 'err')
        if opened and record:
            self._record_open(payload=payload)

    def _open_hymn_ref(self, book_ref, num_ref='', record=True, *, show_overlay=True, show_toast=True):
        matches = resolve_open_request(self.books, book_ref, num_ref)
        if not matches:
            if show_toast:
                self._show_toast(f'找不到：{book_ref} / {num_ref}', 'warn')
            return False
        if len(matches) == 1:
            self._open_from_payload(
                matches[0],
                record=record,
                show_overlay=show_overlay,
                show_toast=show_toast,
            )
            return True
        self._sync_remote_query(book_ref, num_ref)
        self._populate_matches_in_file_list(matches)
        if show_toast:
            self._show_toast(f'{len(matches)} 個結果 — 請揀選', 'warn')
        return False

    # ─────────────────────────────────────────────────────────────
    def _update_sel_info(self, book=None):
        if not book:
            self.sel_info.setText("  ▸  請從左側選擇詩歌冊")
            return
        info = f"已選：{book['name']}"
        if book['hymn_count']:
            info += f"  ·  {book['hymn_count']} 首"
        if book['bookmark']:
            info += "  ·  PDF + 書籤"
        elif book['docx'] and book['pdf']:
            info += "  ·  DOCX + PDF"
        elif book['docx']:
            info += "  ·  DOCX"
        elif book['pdf']:
            info += "  ·  PDF"
        self.sel_info.setText(f"  ▸  {info}")

    # ─────────────────────────────────────────────────────────────
    #  FILTER — 書本模式過濾、單一結果 Open、按鈕狀態
    # ─────────────────────────────────────────────────────────────
    def _actionable_file_items(self):
        items = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole):
                items.append(item)
        return items

    def _update_search_action_buttons(self):
        single = len(self._actionable_file_items()) == 1
        open_style = getattr(self, '_open_btn_style', '')
        action_style = getattr(self, '_action_btn_style', '')
        if self.search_mode == 'book':
            is_open = single
            self.btn_filter.setText('Open' if is_open else 'Filter')
            self.btn_filter.setStyleSheet(open_style if is_open else action_style)
        elif self.search_mode == 'global':
            is_open = single
            self.btn_global_search.setText('Open' if is_open else '搜索')
            self.btn_global_search.setStyleSheet(open_style if is_open else action_style)
        elif self.search_mode == 'keyword':
            is_open = single
            self.btn_keyword_search.setText('Open' if is_open else '搜索')
            self.btn_keyword_search.setStyleSheet(open_style if is_open else action_style)

    def _open_first_result(self):
        items = self._actionable_file_items()
        if items:
            self._file_open_item(items[0])

    def _try_open_single_result(self):
        if len(self._actionable_file_items()) == 1:
            self._open_first_result()
            return True
        return False

    def _on_book_filter_action(self, *_):
        if self.search_mode != 'book':
            return
        if self._try_open_single_result():
            return
        self._apply_sidebar_filter()

    def _on_keyword_search_action(self, *_):
        if self.search_mode != 'keyword':
            return
        if self._try_open_single_result():
            return
        self._start_keyword_search()

    def _on_book_text_changed(self, _):
        self._apply_sidebar_filter()

    def _apply_sidebar_filter(self, *_):
        if self.search_mode != 'book':
            return
        bq = self.inp_book.text().strip().lower()
        fq = self.inp_num.text().strip().lower()
        t  = THEMES[self.theme_name]

        visible = []
        for i in range(self.book_list.count()):
            item = self.book_list.item(i)
            book = item.data(Qt.ItemDataRole.UserRole)
            if not book:
                continue
            match = not bq or bq in book['name'].lower()
            item.setHidden(not match)
            if match:
                visible.append(book)

        preview = self.current_book if self.current_book in visible else None
        if not preview and visible:
            preview = visible[0]

        if preview:
            self._show_book_files(preview, fq)
            self._update_sel_info(preview)
        else:
            self.file_list.clear()
            ni = QListWidgetItem("  找不到符合條件的書冊")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)
            self._update_sel_info(None)
        self._update_search_action_buttons()
        self._maybe_focus_hymn_after_book_match(visible, bq)

    def _maybe_focus_hymn_after_book_match(self, visible, bq):
        if not self._settings.get('book_auto_focus_hymn'):
            return
        if not bq or len(visible) != 1:
            return
        if not self.inp_book.hasFocus():
            return
        QTimer.singleShot(0, lambda: self.inp_num.setFocus(Qt.FocusReason.OtherFocusReason))

    # ─────────────────────────────────────────────────────────────
    #  SEARCH MODE — 切換書本/全局/關鍵字模式與主題、字體
    # ─────────────────────────────────────────────────────────────
    def _apply_schedule_file_layout(self, in_schedule):
        if hasattr(self, 'schedule_panel'):
            self.schedule_panel.setVisible(in_schedule)
        if hasattr(self, 'file_list'):
            self.file_list.setVisible(not in_schedule)
        if hasattr(self, 'file_list_header'):
            self.file_list_header.setVisible(not in_schedule)
        if not in_schedule and hasattr(self, 'sched_entry_dropdown'):
            self.sched_entry_dropdown.hide()

    def _set_search_mode(self, mode):
        if mode != 'keyword':
            self._stop_search_worker()
            self._keyword_debounce.stop()

        self.search_mode = mode
        self.btn_book.setChecked(mode == 'book')
        self.btn_global.setChecked(mode == 'global')
        self.btn_keyword.setChecked(mode == 'keyword')
        self.btn_schedule.setChecked(mode == 'schedule')
        self.book_row.setVisible(mode == 'book')
        self.schedule_row.setVisible(mode == 'schedule')
        self.global_row.setVisible(mode == 'global')
        self.keyword_row.setVisible(mode == 'keyword')
        self.keyword_index_row.setVisible(mode == 'keyword')
        self.dropdown.hide()
        if hasattr(self, 'sched_entry_dropdown'):
            self.sched_entry_dropdown.hide()
        self._apply_schedule_file_layout(mode == 'schedule')

        if mode == 'book':
            self.file_hdr.setText("檔案 / 書籤")
            self.hint_lbl.setText("選一個左邊書冊，右邊會列出該書的所有檔案")
            if self.current_book:
                self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
                self._update_sel_info(self.current_book)
        elif mode == 'schedule':
            self.hint_lbl.setText("雙擊搜尋結果加入播放清單；雙擊播放清單開啟")
            self.sel_info.setText("  ▸  排程模式 — 管理今日詩歌順序")
            self._rebuild_sched_book_combo()
            self._refresh_schedule_search()
            self._refresh_schedule_list()
        elif mode == 'global':
            self.file_hdr.setText("全局搜索結果")
            self.hint_lbl.setText("輸入歌名，會在所有書冊的 PDF 書籤及檔名中搜索")
            self.sel_info.setText("  ▸  全局模式 — 跨書冊搜索")
            self._show_global_results(self.inp_global.text())
        else:
            self.file_hdr.setText("關鍵字搜索結果")
            self.hint_lbl.setText(
                "建立索引後搜索更快；勾選「即時搜索」可邊輸入邊找"
            )
            self.sel_info.setText("  ▸  關鍵字模式 — 全文搜索")
            self._update_index_status_label()
            self._show_keyword_placeholder()

        self._update_search_action_buttons()
        self._persist_settings(search_mode=mode)

    def _on_global_search(self, *_):
        if self.search_mode != 'global':
            return
        if self._try_open_single_result():
            return
        self._show_global_results(self.inp_global.text())

    def _on_theme_changed(self, name):
        self.theme_name = name
        self._apply_theme()
        if self.search_mode == 'global':
            self._show_global_results(self.inp_global.text())
        elif self.search_mode == 'keyword':
            pass  # keep current keyword results
        elif self.search_mode == 'schedule':
            self._refresh_schedule_list()
            self._refresh_schedule_search()
        elif self.current_book:
            self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
            self._update_sel_info(self.current_book)
        # Refresh error-colored items if any
        t = THEMES[self.theme_name]
        for i in range(self.book_list.count()):
            item = self.book_list.item(i)
            if not item.data(Qt.ItemDataRole.UserRole):
                item.setForeground(QColor(t['err']))
        self._persist_settings(theme=name)

    def _on_font_changed(self, val):
        self.font_size = val
        self.fs_lbl.setText(f"{val}px")
        self._apply_theme()
        if self.search_mode == 'global':
            self._show_global_results(self.inp_global.text())
        elif self.search_mode == 'keyword':
            pass  # keep current keyword results
        elif self.search_mode == 'schedule':
            self._refresh_schedule_list()
            self._refresh_schedule_search()
        elif self.current_book:
            self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
        self._persist_settings(font_size=val)

    # ─────────────────────────────────────────────────────────────
    #  TOAST — 顯示底部操作結果提示
    # ─────────────────────────────────────────────────────────────
    def _show_toast(self, msg, kind='ok'):
        t = THEMES[self.theme_name]
        colours = {
            'ok':   (t['ok_bg'],   t['ok']),
            'warn': (t['warn_bg'], t['warn']),
            'err':  (t['err_bg'],  t['err']),
        }
        bg, fg = colours.get(kind, (t['ok_bg'], t['ok']))
        self.toast_bar.setStyleSheet(
            f"background: {bg}; color: {fg}; "
            f"font-size: {self.font_size - 1}px; padding: 0 20px; "
            f"border-top: 1px solid {THEMES[self.theme_name]['border']};"
        )
        self.toast_bar.setText(f"   {msg}")
        self.toast_bar.setFixedHeight(34)
        self._toast_timer.start(3500)

    # ─────────────────────────────────────────────────────────────
    #  REMOTE — Mobile API 遙控開檔
    # ─────────────────────────────────────────────────────────────
    def _remote_local_base_url(self):
        if not self._remote.running:
            return ''
        return self._remote.url().rstrip('/')

    def _remote_wan_base_url(self):
        if not self.enable_tunnel or not self._remote.running:
            return ''
        if getattr(self, '_tunnel', None) and self._tunnel.is_ready and self._tunnel.public_url:
            return normalize_tunnel_url(self._tunnel.public_url).rstrip('/')
        if self.named_tunnel_url.strip() and self.tunnel_mode in ('named', 'service'):
            return normalize_tunnel_url(self.named_tunnel_url).rstrip('/')
        return ''

    def _remote_share_url(self):
        """QR Code：預設本機 WiFi；啟用外網後改用公開 HTTPS。"""
        if self.enable_tunnel:
            wan = self._remote_wan_base_url()
            if wan:
                return tunnel_viewer_url(wan)
        return self._remote_local_base_url()

    def _remote_public_base_url(self):
        """Backward-compatible alias for QR / mixin."""
        return self._remote_share_url()

    def _remote_viewer_url(self):
        base = self._remote_wan_base_url()
        return tunnel_viewer_url(base) if base else '—'

    def _remote_admin_url(self):
        base = self._remote_wan_base_url()
        return f"{base}/admin" if base else '—'

    def _update_remote_url_labels(self):
        if self._remote.running:
            local = self._remote_local_base_url()
            self.local_url_lbl.setText(local or '—')
            wan_on = bool(getattr(self, 'enable_tunnel', False))
            if wan_on:
                self.viewer_url_lbl.setText(self._remote_viewer_url())
                self.remote_url_lbl.setText(self._remote_admin_url())
            else:
                self.viewer_url_lbl.setText('—')
                self.remote_url_lbl.setText('—')
        else:
            self.local_url_lbl.setText('—')
            self.remote_url_lbl.setText('—')
            self.viewer_url_lbl.setText('—')

    def _start_cloudflare_tunnel(self):
        if not self.enable_tunnel or not self._remote.running:
            return
        if self.tunnel_mode == 'named' and is_cloudflared_service_installed():
            self._show_toast(
                '本機已有 cloudflared 服務，請改選「共用 Windows 服務」以免開兩個 tunnel',
                'warn',
            )
            self.tunnel_mode = 'service'
            if hasattr(self, 'tunnel_mode_combo'):
                idx = self.tunnel_mode_combo.findData('service')
                if idx >= 0:
                    self.tunnel_mode_combo.blockSignals(True)
                    self.tunnel_mode_combo.setCurrentIndex(idx)
                    self.tunnel_mode_combo.blockSignals(False)
            self._persist_settings(tunnel_mode='service')
        if self.tunnel_mode in ('named', 'service') and not self.named_tunnel_url.strip():
            self.tunnel_status_lbl.setText('外網：請設定公開網址')
            self._show_toast('啟用外網時需填寫公開網址', 'warn')
            return
        self.tunnel_status_lbl.setText('外網：正在啟動通道...')
        self._stop_cloudflare_tunnel()
        port = self._remote.state.port
        self._tunnel = CloudflareTunnel(
            local_port=port,
            cloudflared_path=self.cloudflared_path,
            tunnel_mode=self.tunnel_mode,
            cloudflared_token=self.cloudflared_token,
            named_tunnel_url=self.named_tunnel_url,
            use_https=False,
            on_url=lambda url: self.tunnel_url_signal.emit(url),
            on_ready=lambda url: self.tunnel_ready_signal.emit(url),
            on_error=lambda msg: self.tunnel_error_signal.emit(msg),
            on_expired=lambda msg: self.tunnel_expired_signal.emit(msg),
        )
        threading.Thread(target=self._tunnel.start, daemon=True).start()

    def _stop_cloudflare_tunnel(self):
        if self._tunnel:
            self._tunnel.stop()
            self._tunnel = None
        if hasattr(self, 'tunnel_status_lbl'):
            self.tunnel_status_lbl.setText('外網：—')

    def _on_tunnel_toggled(self, checked):
        self.enable_tunnel = checked
        self._persist_settings(enable_tunnel=checked)
        self._sync_tunnel_options_visibility()
        if checked and self._remote.running:
            self._start_cloudflare_tunnel()
        elif not checked:
            self._stop_cloudflare_tunnel()
            self._update_remote_url_labels()
            self._update_qr_code()

    def _make_tunnel_cmd_row(self, title):
        row = QWidget()
        lay = QHBoxLayout(row)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        lbl = QLabel(title)
        lbl.setMinimumWidth(80)
        inp = QLineEdit()
        inp.setReadOnly(True)
        btn = QPushButton("複製")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setFixedWidth(52)
        lay.addWidget(lbl)
        lay.addWidget(inp, 1)
        lay.addWidget(btn)
        return row, inp, btn

    def _copy_tunnel_command(self, text):
        cmd = str(text or '').strip()
        if not cmd:
            self._show_toast('沒有可複製的指令', 'warn')
            return
        QApplication.clipboard().setText(cmd)
        self._show_toast('已複製指令', 'ok')

    def _update_tunnel_debug_commands(self):
        if not hasattr(self, 'inp_tunnel_install_cmd'):
            return
        token = ''
        if hasattr(self, 'inp_cloudflared_token'):
            token = self.inp_cloudflared_token.text().strip()
        if not token:
            token = str(getattr(self, 'cloudflared_token', '') or '').strip()
        path = str(getattr(self, 'cloudflared_path', '') or '')
        self.inp_tunnel_install_cmd.setText(
            cloudflared_service_install_command(token, path)
        )
        self.inp_tunnel_uninstall_cmd.setText(
            cloudflared_service_uninstall_command(path)
        )
        self.inp_tunnel_restart_cmd.setText(cloudflared_service_restart_command())

    def _sync_tunnel_options_visibility(self):
        if not hasattr(self, 'chk_enable_tunnel'):
            return
        on = self.chk_enable_tunnel.isChecked()
        mode = self.tunnel_mode_combo.currentData() if on else None
        service_installed = is_cloudflared_service_installed()
        show_service = on and mode == 'service'
        show_install = show_service and not service_installed

        for w in (
            self.tunnel_mode_combo,
            self.tunnel_help_lbl,
            self.tunnel_status_lbl,
        ):
            w.setVisible(on)
        self.lbl_public_tunnel_url.setVisible(show_service)
        self.inp_named_tunnel_url.setVisible(show_service)
        self.inp_cloudflared_token.setVisible(show_install)
        self.btn_install_cloudflared.setVisible(show_install)
        for w in (
            self.lbl_viewer_url, self.viewer_url_lbl,
            self.lbl_remote_url, self.remote_url_lbl,
        ):
            w.setVisible(on)
        show_cmds = show_service
        for w in (
            self.lbl_tunnel_cmd_hint,
            self._tunnel_install_row,
            self._tunnel_uninstall_row,
            self._tunnel_restart_row,
        ):
            w.setVisible(show_cmds)
        if show_cmds:
            self._update_tunnel_debug_commands()
        if on:
            self._refresh_tunnel_service_ui()
        self._update_remote_url_labels()
        self._update_qr_code()

    def _refresh_tunnel_service_ui(self):
        if not hasattr(self, 'tunnel_help_lbl'):
            return
        port = self.spin_remote_port.value() if hasattr(self, 'spin_remote_port') else DEFAULT_PORT
        service_installed = is_cloudflared_service_installed()
        mode = getattr(self, 'tunnel_mode', 'service')
        dns_note = (
            'DNS：Cloudflare 須有 live → <tunnel-id>.cfargotunnel.com 的 CNAME，'
            '否則公開網址無法解析。'
        )
        origin_note = (
            f'Origin（Cloudflare 後台）：http://127.0.0.1:{port} '
            '（勿用 localhost；Windows 上 localhost 會走 IPv6）。'
        )
        if service_installed:
            wan_off_note = (
                ' 未勾選「外網」時 App 只顯示本機 IP；'
                '若公開網址仍可連，代表 cloudflared Windows 服務仍在背景運行。'
            )
            self.tunnel_help_lbl.setText(
                f'{shared_tunnel_setup_hint(port)} {dns_note}{wan_off_note}'
            )
            self.btn_install_cloudflared.setEnabled(False)
            self.btn_install_cloudflared.setToolTip(
                '本機已有 cloudflared 服務。'
            )
        else:
            self.btn_install_cloudflared.setEnabled(True)
            self.btn_install_cloudflared.setToolTip(
                '僅在未有 cloudflared 服務時使用。'
            )
            if mode == 'service':
                self.tunnel_help_lbl.setText(
                    f'{shared_tunnel_setup_hint(port)} {dns_note}'
                )
            else:
                self.tunnel_help_lbl.setText(
                    f'快速通道產生隨機 trycloudflare.com 網址，適合測試。{origin_note}'
                )
        self._update_tunnel_debug_commands()

    def _on_tunnel_mode_changed(self, _index):
        self.tunnel_mode = self.tunnel_mode_combo.currentData()
        self._persist_settings(tunnel_mode=self.tunnel_mode)
        self._sync_tunnel_options_visibility()

    def _on_named_tunnel_url_changed(self, text):
        self.named_tunnel_url = text.strip()
        self.remote_path_prefix = tunnel_url_path_prefix(self.named_tunnel_url)
        self._persist_settings(
            named_tunnel_url=self.named_tunnel_url,
            remote_path_prefix=self.remote_path_prefix,
        )

    def _on_cloudflared_token_changed(self, text):
        self.cloudflared_token = text
        self._persist_settings(cloudflared_token=text)
        self._update_tunnel_debug_commands()

    def _install_cloudflared_service(self):
        if is_cloudflared_service_installed():
            self._show_toast('本機已有 cloudflared 服務，請改用「共用 Windows 服務」', 'warn')
            self._refresh_tunnel_service_ui()
            return
        token = self.inp_cloudflared_token.text().strip()
        if not token:
            self._show_toast('請先輸入 Cloudflare tunnel token', 'warn')
            return
        self.btn_install_cloudflared.setEnabled(False)
        self._show_toast('正在安裝 cloudflared Windows 服務...', 'ok')

        def worker():
            ok, message = install_cloudflared_service(token, self.cloudflared_path)
            self.service_install_signal.emit(ok, message)

        threading.Thread(target=worker, daemon=True).start()

    def _on_service_install_done(self, ok, message):
        self._sync_tunnel_options_visibility()
        self._refresh_tunnel_service_ui()
        if ok:
            self._show_toast(f'cloudflared 服務：{message}', 'ok')
        else:
            self._show_toast(f'cloudflared 服務安裝失敗：{message}', 'err')

    def _on_tunnel_url(self, url):
        self.tunnel_status_lbl.setText(f'外網：{url}\n（等待 Cloudflare...）')

    def _on_tunnel_ready(self, url):
        if self._tunnel and not self._tunnel.local_dns_ok:
            self.tunnel_status_lbl.setText(
                f'外網：{url}\n'
                '✓ 通道已啟用\n'
                '⚠ 本機 DNS 無法解析 — 請用 1.1.1.1 或行動網路'
            )
        else:
            self.tunnel_status_lbl.setText(f'外網：{url}\n✓ 通道已啟用')
        self._update_remote_url_labels()
        self._update_qr_code()

    def _on_tunnel_expired(self, message):
        self.tunnel_status_lbl.setText(f'外網：⚠ {message}')
        self._update_remote_url_labels()
        self._update_qr_code()

    def _on_tunnel_error(self, message):
        self.tunnel_status_lbl.setText(f'外網：{message}\n（詳情見 exe 旁 tunnel.log）')

    def _on_remote_api_toggled(self, checked):
        self._remote.state.set_api_enabled(checked)
        if checked:
            port = self.spin_remote_port.value()
            path_prefix = self.remote_path_prefix or tunnel_url_path_prefix(self.named_tunnel_url)
            if not self._remote.start(port, path_prefix=path_prefix, use_https=False):
                self.chk_remote_api.blockSignals(True)
                self.chk_remote_api.setChecked(False)
                self.chk_remote_api.blockSignals(False)
                self._remote.state.set_api_enabled(False)
                self.spin_remote_port.setEnabled(True)
                self.remote_url_lbl.setText('—')
                self.viewer_url_lbl.setText('—')
                err = self._remote.last_error or 'port 可能已被佔用'
                self._show_toast(
                    f'API 啟動失敗：{err}（詳情見 exe 旁 remote_api.log）',
                    'err',
                )
                self._sync_remote_accept_checkboxes(checked=False)
                self._sync_remote_accept_enabled()
                self._update_remote_status_label()
                self._persist_settings(remote_api_enabled=False, remote_accept=False)
                return
            self.spin_remote_port.setEnabled(False)
            self._update_remote_url_labels()
            if self.enable_tunnel:
                self._start_cloudflare_tunnel()
            if self._settings.get('last_opened'):
                self._sync_viewer_now_playing(label=self._settings['last_opened'])
        else:
            self._stop_cloudflare_tunnel()
            self._remote.stop()
            self.spin_remote_port.setEnabled(True)
            self._update_remote_url_labels()
            self._sync_remote_accept_checkboxes(checked=False)
        self._sync_remote_accept_enabled()
        self._update_remote_status_label()
        self._update_qr_code()
        self._refresh_remote_log_view()
        self._persist_settings(remote_api_enabled=checked)
        if not checked:
            self._persist_settings(remote_accept=False)

    def _sync_remote_accept_enabled(self):
        enabled = self.chk_remote_api.isChecked() and self._remote.running
        self.chk_remote_accept.setEnabled(enabled)
        self.chk_remote_accept_hdr.setEnabled(enabled)
        self.combo_remote_policy_hdr.setEnabled(enabled)
        if not enabled:
            self._remote.state.set_accepting(False)
            self._sync_remote_accept_checkboxes(checked=False)

    def _sync_remote_accept_checkboxes(self, checked):
        for chk in (self.chk_remote_accept_hdr, self.chk_remote_accept):
            chk.blockSignals(True)
            chk.setChecked(checked)
            chk.blockSignals(False)

    def _on_remote_accept_toggled(self, checked):
        self._remote.state.set_accepting(checked)
        self._update_remote_status_label()
        self._persist_settings(remote_accept=checked)

    def _on_remote_accept_hdr_toggled(self, checked):
        self._sync_remote_accept_checkboxes(checked)
        self._on_remote_accept_toggled(checked)

    def _on_remote_accept_settings_toggled(self, checked):
        self._sync_remote_accept_checkboxes(checked)
        self._on_remote_accept_toggled(checked)

    def _on_remote_policy_changed(self, _idx):
        policy = self.combo_remote_policy_hdr.currentData()
        self._remote.state.set_policy(policy)
        self._persist_settings(remote_policy=policy)

    def _on_remote_token_changed(self, text):
        self._remote.state.set_token(text)
        self._persist_settings(remote_token=text)
        self._update_qr_code()

    def _on_viewer_show_debug_toggled(self, checked):
        self._remote.state.set_viewer_show_debug(checked)
        self._persist_settings(viewer_show_debug=checked)

    def _on_remote_port_changed(self, val):
        self._remote.state.set_port(val)
        self._persist_settings(remote_port=val)

    def _update_remote_status_label(self):
        self._update_footer_status()

    def _on_remote_open_payload(self, payload):
        self._open_from_payload(payload)
        self._show_toast('📱  Mobile 遙控已開啟', 'ok')

    def _sync_remote_query(self, book_ref, num_ref):
        matched = resolve_books(self.books, book_ref)
        if not matched:
            return
        book = matched[0]
        self._set_search_mode('book')
        self.inp_num.blockSignals(True)
        self.inp_num.setText(str(num_ref))
        self.inp_num.blockSignals(False)
        self._select_book(book)

    def _on_remote_populate_ui(self, book_ref, num_ref, matches):
        self._sync_remote_query(book_ref, num_ref)
        self._show_toast('📱  Mobile 請求 — 請揀選開啟', 'warn')

    def _on_remote_pending_request(self, req):
        self._remote_pending_id = req.id
        self.remote_pending_lbl.setText(
            f"📱  待確認：{req.book_ref} / {req.num_ref}  —  {req.summary}"
        )
        self.remote_pending_bar.setVisible(True)
        self._show_toast('📱  Mobile 請求待確認', 'warn')
        self._update_footer_status()

    def _approve_remote_pending(self):
        if not self._remote_pending_id:
            return
        result = self._remote.state.approve_pending(self._remote_pending_id)
        self._remote_pending_id = None
        self.remote_pending_bar.setVisible(False)
        self._update_footer_status()
        if result.get('status') == 'opened':
            self._show_toast('✅  已確認開啟', 'ok')
        elif result.get('status') == 'ui_populated':
            self._show_toast('📋  請在列表揀選', 'warn')

    def _reject_remote_pending(self):
        if not self._remote_pending_id:
            return
        self._remote.state.reject_pending(self._remote_pending_id)
        self._remote_pending_id = None
        self.remote_pending_bar.setVisible(False)
        self._update_footer_status()
        self._show_toast('已忽略 Mobile 請求', 'warn')

    def closeEvent(self, event):
        if self.handle_close_event(event):
            return
        DesktopQrOverlay.hide()
        self._stop_cloudflare_tunnel()
        self._remote.stop()
        super().closeEvent(event)

    # ─────────────────────────────────────────────────────────────
    #  EVENT FILTER — Enter 確認 Mobile 待辦；inp_book 聚焦行為
    # ─────────────────────────────────────────────────────────────
    def _pending_enter_modifiers_ok(self, event):
        mods = event.modifiers()
        return not (
            mods & Qt.KeyboardModifier.ControlModifier
            or mods & Qt.KeyboardModifier.AltModifier
            or mods & Qt.KeyboardModifier.MetaModifier
        )

    def eventFilter(self, obj, event):
        if (event.type() == QEvent.Type.KeyPress
                and self._remote_pending_id
                and self.remote_pending_bar.isVisible()
                and event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
                and self._pending_enter_modifiers_ok(event)):
            self._approve_remote_pending()
            return True

        if obj == self.inp_book:
            if event.type() == QEvent.Type.FocusIn and self.search_mode == 'book':
                if self.inp_book.text():
                    self.inp_book.clear()
                    self.dropdown.hide()
                self._clear_file_keyword()
            elif event.type() == QEvent.Type.FocusOut:
                QTimer.singleShot(200, self.dropdown.hide)
        elif obj == self.inp_sched_num:
            if event.type() == QEvent.Type.FocusIn and self.search_mode == 'schedule':
                self._show_sched_num_suggestions()
            elif event.type() == QEvent.Type.FocusOut:
                QTimer.singleShot(200, self.sched_entry_dropdown.hide)
        return super().eventFilter(obj, event)


# ══════════════════════════════════════════════════════════════
#  ENTRY POINT — 建立 QApplication、載入圖示、顯示主視窗
# ══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app_icon = load_app_icon()
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window,          QColor('#0c0e1a'))
    pal.setColor(QPalette.ColorRole.WindowText,      QColor('#e4e8f5'))
    pal.setColor(QPalette.ColorRole.Base,            QColor('#191c30'))
    pal.setColor(QPalette.ColorRole.AlternateBase,   QColor('#1b1f36'))
    pal.setColor(QPalette.ColorRole.Text,            QColor('#e4e8f5'))
    pal.setColor(QPalette.ColorRole.Button,          QColor('#13162a'))
    pal.setColor(QPalette.ColorRole.ButtonText,      QColor('#e4e8f5'))
    pal.setColor(QPalette.ColorRole.Highlight,       QColor('#7c6af7'))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor('#ffffff'))
    app.setPalette(pal)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())