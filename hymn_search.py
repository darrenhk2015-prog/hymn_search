#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
詩歌冊搜索系統（Hymn Search）— Phase 1
=====================================

【程式簡介】
    本程式為 PyQt6 桌面應用，用於在本地詩歌資料夾中快速搜尋、瀏覽及開啟
    Word / PDF 詩歌檔案。支援多種搜尋模式（標準 / 全局 / 關鍵字 / 排程）、
    PDF 書籤跳頁、全文索引、開啟檔案時的置頂書名提示、多螢幕延伸顯示，以及可選的
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

import os, sys, re, subprocess, zipfile, json, gzip, time, threading, shutil
import fitz  # PyMuPDF

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton, QToolButton,
    QLineEdit, QLabel, QListWidget, QListWidgetItem,
    QComboBox, QSlider, QFrame, QScrollArea, QCheckBox, QSpinBox,
    QGroupBox, QFileDialog, QTabWidget, QPlainTextEdit, QMenu, QDialog,
    QSystemTrayIcon, QColorDialog, QMessageBox, QSplitter, QFontComboBox,
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
from hymn_features.hymn_catalog import (
    default_hymns_json_path,
    find_hymn_in_books,
    hymn_payload,
    load_books,
    payload_from_legacy,
    resolve_hymns_json_path,
    search_hymn_lyrics,
)
from hymn_features.lyric_session import (
    ALIGN_JUSTIFY,
    ALIGN_LEFT,
    DEFAULT_LETTER_SPACING,
    DEFAULT_LINE_HEIGHT,
    DEFAULT_LYRIC_ALIGN,
    DEFAULT_LYRIC_BG,
    DEFAULT_LYRIC_FONT_FAMILY,
    DEFAULT_LYRIC_FONT_PT,
    DEFAULT_LYRIC_MARGINS,
    DEFAULT_LYRIC_TEXT,
    LETTER_SPACING_MAX,
    LETTER_SPACING_MIN,
    LINE_HEIGHT_MAX,
    LINE_HEIGHT_MIN,
    LYRIC_BG_IMAGE_EXTS,
    LYRIC_BG_IMAGE_STEM,
    LYRIC_FONT_PT_MAX,
    LYRIC_FONT_PT_MIN,
    LYRIC_MARGIN_MAX,
    LYRIC_MARGIN_MIN,
    MARGIN_SIDES,
    ZOOM_STEP,
    factory_lyric_best_defaults,
    lyric_best_from_session,
    lyric_best_from_values,
    normalize_letter_spacing,
    normalize_line_height,
    normalize_lyric_align,
    normalize_lyric_best_defaults,
    normalize_lyric_bg_image,
    normalize_lyric_font_family,
    normalize_lyric_font_pt,
    normalize_lyric_margins,
)
from hymn_features.lyric_windows import (
    PLACEMENT_DOCKED,
    PLACEMENT_WINDOW,
    LyricWindowPair,
)
from hymn_features.session_store import (
    add_setlist_entry, advance_setlist_index, clear_setlist, clamp_setlist, move_setlist_entry,
    parse_setlist, remove_setlist_at, session_history_list, setlist_index,
)
from hymn_features.display_flow import (
    OpenFlowController,
    close_document_viewer,
    run_pdf_projection_flow,
    run_word_projection_flow,
    send_alt_f4_close,
)
from hymn_features.win_display import (
    DISPLAY_SWITCH_CCD,
    DISPLAY_SWITCH_LEGACY,
    switch_windows_display_mode,
)
from hymn_features.phone_control_wizard import PhoneControlWizard
from hymn_features.voice_input import VoiceInputController, voice_input_available

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
    HYMN_NUM_MODE_CONTAINS, HYMN_NUM_MODE_EXACT, HYMN_NUM_MODES,
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
/* Font list uses a custom delegate that ignores dark item colors — keep popup light. */
QFontComboBox {{
    background: {t['in_bg']}; border: 1.5px solid {t['in_bd']};
    border-radius: 8px; padding: 5px 10px; color: {t['text']};
    font-size: {fs - 1}px; min-width: 160px;
}}
QFontComboBox:focus {{ border-color: {t['accent']}; }}
QFontComboBox QAbstractItemView {{
    background: #ffffff; color: #1a1e40;
    border: 1.5px solid {t['border']}; outline: none;
    selection-background-color: {t['accent']};
    selection-color: #ffffff;
}}
QFontComboBox QLineEdit {{
    background: transparent; border: none; color: {t['text']};
    padding: 0; selection-background-color: {t['accent']};
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
    pad = 14 if fs >= 20 else 9
    return f"""
QListWidget {{ background: transparent; border: none; outline: none; font-size: {fs}px; }}
QListWidget::item {{
    padding: {pad}px 12px; color: {t['text2']};
    border-radius: 8px; margin: 2px 3px;
}}
QListWidget::item:hover {{ background: {t['panel2']}; color: {t['text']}; }}
QListWidget::item:selected {{ background: {t['accent']}; color: {t['a_text']}; }}
"""


def build_file_list_style(theme, fs):
    t = THEMES[theme]
    pad = 14 if fs >= 20 else 8
    return f"""
QListWidget {{ background: transparent; border: none; outline: none; font-size: {fs}px; }}
QListWidget::item {{
    padding: {pad}px 12px; color: {t['text']};
    border-radius: 8px; margin: 2px 3px;
}}
QListWidget::item:hover {{ background: {t['panel2']}; }}
QListWidget::item:selected {{ background: {t['a_sub']}; color: {t['accent']}; }}
"""


# 簡易版 — 大字、簡化介面
SENIOR_FONT_SIZE = 22
FONT_SIZE_MIN = 10
FONT_SIZE_MAX = 28
SENIOR_SIDEBAR_W = 320
SENIOR_INPUT_H = 56
SENIOR_BTN_W = 120
NORMAL_SIDEBAR_W = 240
NORMAL_INPUT_H = 40
NORMAL_FILTER_BTN_W = 84
NORMAL_HEADER_H = 48
SENIOR_HEADER_H = 56
NORMAL_SIDEBAR_HDR_H = 46
SENIOR_SIDEBAR_HDR_H = 56


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
#  SCANNER — 由 hymns.json 組成詩歌冊
# ══════════════════════════════════════════════════════════════
def scan_books(folder=None):
    """Load volumes + hymns from hymns.json (folder may be a json path or data dir)."""
    return load_books(folder)


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
# switch_windows_display_mode：延伸桌面 ↔ 同步（複製）畫面（SetDisplayConfig API）。
def open_pdf_at_page(path, page, *, presentation=False):
    if sys.platform == 'win32':
        for sp in [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
        ]:
            if os.path.exists(sp):
                args = [sp, '-reuse-instance']
                if presentation:
                    args.append('-presentation')
                args.extend(['-page', str(page), path])
                subprocess.Popen(args)
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


def open_pdf_presentation(path):
    """Open PDF; Sumatra uses -presentation to avoid a separate F11 toggle."""
    if sys.platform == 'win32':
        for sp in [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
        ]:
            if os.path.exists(sp):
                subprocess.Popen([sp, '-reuse-instance', '-presentation', path])
                return True, 'Sumatra PDF'
        os.startfile(path)
        return True, '系統預設'
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
        return True, 'Preview'
    else:
        subprocess.Popen(['xdg-open', path])
        return True, '系統預設'


# ══════════════════════════════════════════════════════════════
#  BOOK NAME OVERLAY — 開啟檔案時的置頂書名浮層
# ══════════════════════════════════════════════════════════════
# 支援多螢幕延伸顯示；可置中→右下角動畫、定時隱藏或一直顯示；
# 關閉 Word/PDF 視窗時自動隱藏（document_viewer_is_open 輪詢）。
WORD_EXTS = {'.doc', '.docx'}
DEFAULT_OVERLAY_SECS = 10
MAX_OVERLAY_SECS = 600
OVERLAY_DURATION_ALWAYS = -1
CENTER_OVERLAY_SECS = 5
MAX_OVERLAY_CENTER_SECS = 600
OVERLAY_MODE_CORNER = 'corner'
OVERLAY_MODE_CENTER = 'center_then_corner'
DEFAULT_OVERLAY_BG_COLOR = '#f59e0b'
DEFAULT_OVERLAY_TEXT_COLOR = ''
DEFAULT_QR_DESKTOP_SIZE = 140
DEFAULT_QR_X = 24
DEFAULT_QR_Y = 24
DEFAULT_QR_WIDTH = 220
DEFAULT_QR_HEIGHT = 300
DEFAULT_QR_CAPTION = '掃碼觀看'
DEFAULT_QR_BG_COLOR = '#ffffff'
DEFAULT_QR_TEXT_COLOR = ''
DEFAULT_QR_TEXT_SIZE = 11
DEFAULT_QR_BG_OPACITY = 88
DEFAULT_QR_STACK_GAP = 12

SETTINGS_VERSION = 5

SEARCH_MODES = ('standard', 'global', 'keyword', 'schedule')
LEGACY_SEARCH_MODES = {'book': 'standard'}
BOOK_UI_MODES = ('standard',)


def default_settings():
    return {
        'version': SETTINGS_VERSION,
        'hymn_folder': '',
        'theme': 'dark',
        'font_size': 13,
        'senior_mode': False,
        'senior_snapshot': None,
        'search_mode': 'standard',
        'open_overlay': True,
        'overlay_duration': DEFAULT_OVERLAY_SECS,
        'overlay_center_secs': CENTER_OVERLAY_SECS,
        'overlay_mode': OVERLAY_MODE_CORNER,
        'overlay_bg_color': DEFAULT_OVERLAY_BG_COLOR,
        'overlay_text_color': DEFAULT_OVERLAY_TEXT_COLOR,
        'display_duplicate': False,
        'display_switch_mode': DISPLAY_SWITCH_CCD,
        'word_auto_fullscreen': True,
        'pdf_auto_fullscreen': False,
        'keyword_instant': True,
        'book_auto_focus_hymn': True,
        'hymn_num_mode': HYMN_NUM_MODE_CONTAINS,
        'click_to_open': False,
        'remote_api_enabled': False,
        'remote_accept': False,
        'remote_port': DEFAULT_PORT,
        'remote_use_https': False,
        'remote_policy': OPEN_POLICY_AUTO,
        'remote_token': '',
        'enable_tunnel': False,
        'enable_tunnel_fixed': False,
        'enable_tunnel_quick': False,
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
        'show_qr_lan': True,
        'show_qr_fixed': True,
        'show_qr_quick': True,
        'show_qr_wan': True,
        'qr_x': DEFAULT_QR_X,
        'qr_y': DEFAULT_QR_Y,
        'qr_width': DEFAULT_QR_WIDTH,
        'qr_height': DEFAULT_QR_HEIGHT,
        'qr_caption': DEFAULT_QR_CAPTION,
        'qr_bg_color': DEFAULT_QR_BG_COLOR,
        'qr_text_color': DEFAULT_QR_TEXT_COLOR,
        'qr_text_size': DEFAULT_QR_TEXT_SIZE,
        'qr_bg_opacity': DEFAULT_QR_BG_OPACITY,
        'lyric_bg_color': DEFAULT_LYRIC_BG,
        'lyric_bg_image': '',
        'lyric_text_color': DEFAULT_LYRIC_TEXT,
        'lyric_align': DEFAULT_LYRIC_ALIGN,
        'lyric_letter_spacing': DEFAULT_LETTER_SPACING,
        'lyric_line_height': DEFAULT_LINE_HEIGHT,
        'lyric_margins': dict(DEFAULT_LYRIC_MARGINS),
        'lyric_font_family': DEFAULT_LYRIC_FONT_FAMILY,
        'lyric_font_pt': DEFAULT_LYRIC_FONT_PT,
        'lyric_best_defaults': factory_lyric_best_defaults(),
        'lyric_control_placement': PLACEMENT_WINDOW,
        'lyric_start_blacked': True,
        'lyric_preview_mode': False,
        'overlay_auto_unblack': False,
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
    merged['senior_mode'] = bool(merged.get('senior_mode'))
    snap = merged.get('senior_snapshot')
    if isinstance(snap, dict):
        clean = {}
        if snap.get('theme') in THEMES:
            clean['theme'] = snap['theme']
        try:
            clean['font_size'] = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, int(snap['font_size'])))
        except (TypeError, ValueError, KeyError):
            pass
        if 'click_to_open' in snap:
            clean['click_to_open'] = bool(snap['click_to_open'])
        snap_mode = LEGACY_SEARCH_MODES.get(snap.get('search_mode'), snap.get('search_mode'))
        if snap_mode in SEARCH_MODES:
            clean['search_mode'] = snap_mode
        merged['senior_snapshot'] = clean or None
    else:
        merged['senior_snapshot'] = None
    try:
        merged['font_size'] = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, int(merged['font_size'])))
    except (TypeError, ValueError):
        merged['font_size'] = SENIOR_FONT_SIZE if merged['senior_mode'] else 13
    merged['search_mode'] = LEGACY_SEARCH_MODES.get(merged['search_mode'], merged['search_mode'])
    if merged['search_mode'] not in SEARCH_MODES:
        merged['search_mode'] = 'standard'
    merged['open_overlay'] = bool(merged['open_overlay'])
    try:
        dur = int(merged['overlay_duration'])
        if dur != OVERLAY_DURATION_ALWAYS:
            dur = max(1, min(MAX_OVERLAY_SECS, dur))
        merged['overlay_duration'] = dur
    except (TypeError, ValueError):
        merged['overlay_duration'] = DEFAULT_OVERLAY_SECS
    try:
        merged['overlay_center_secs'] = max(
            1, min(MAX_OVERLAY_CENTER_SECS, int(merged.get('overlay_center_secs', CENTER_OVERLAY_SECS))),
        )
    except (TypeError, ValueError):
        merged['overlay_center_secs'] = CENTER_OVERLAY_SECS
    if merged['overlay_mode'] not in (OVERLAY_MODE_CORNER, OVERLAY_MODE_CENTER):
        merged['overlay_mode'] = OVERLAY_MODE_CORNER
    merged['lyric_bg_color'] = _normalize_hex_color(
        merged.get('lyric_bg_color'), DEFAULT_LYRIC_BG,
    )
    merged['lyric_bg_image'] = normalize_lyric_bg_image(
        merged.get('lyric_bg_image'), app_dir(),
    )
    merged['lyric_text_color'] = _normalize_hex_color(
        merged.get('lyric_text_color'), DEFAULT_LYRIC_TEXT,
    )
    merged['lyric_align'] = normalize_lyric_align(merged.get('lyric_align'))
    merged['lyric_letter_spacing'] = normalize_letter_spacing(merged.get('lyric_letter_spacing'))
    merged['lyric_line_height'] = normalize_line_height(merged.get('lyric_line_height'))
    margins_src = merged.get('lyric_margins')
    if isinstance(raw, dict):
        if 'lyric_margins' in raw:
            margins_src = raw.get('lyric_margins')
        elif 'lyric_margin' in raw:
            margins_src = raw.get('lyric_margin')
    merged['lyric_margins'] = normalize_lyric_margins(margins_src)
    merged.pop('lyric_margin', None)
    merged['lyric_font_family'] = normalize_lyric_font_family(merged.get('lyric_font_family'))
    merged['lyric_font_pt'] = normalize_lyric_font_pt(merged.get('lyric_font_pt'))
    merged['lyric_best_defaults'] = normalize_lyric_best_defaults(merged.get('lyric_best_defaults'))
    placement = str(merged.get('lyric_control_placement') or PLACEMENT_WINDOW).strip().lower()
    merged['lyric_control_placement'] = (
        placement if placement in (PLACEMENT_WINDOW, PLACEMENT_DOCKED) else PLACEMENT_WINDOW
    )
    merged['lyric_start_blacked'] = bool(merged.get('lyric_start_blacked', True))
    merged['lyric_preview_mode'] = bool(merged.get('lyric_preview_mode', False))
    merged['overlay_auto_unblack'] = bool(merged.get('overlay_auto_unblack', False))
    merged['overlay_bg_color'] = _normalize_hex_color(
        merged.get('overlay_bg_color'), DEFAULT_OVERLAY_BG_COLOR,
    )
    overlay_text_color = str(merged.get('overlay_text_color') or '').strip()
    merged['overlay_text_color'] = (
        _normalize_hex_color(overlay_text_color, DEFAULT_OVERLAY_TEXT_COLOR)
        if overlay_text_color else DEFAULT_OVERLAY_TEXT_COLOR
    )
    merged['display_duplicate'] = bool(merged['display_duplicate'])
    mode = str(merged.get('display_switch_mode') or DISPLAY_SWITCH_CCD).strip().lower()
    merged['display_switch_mode'] = mode if mode in (DISPLAY_SWITCH_CCD, DISPLAY_SWITCH_LEGACY) else DISPLAY_SWITCH_CCD
    merged['word_auto_fullscreen'] = bool(merged.get('word_auto_fullscreen', True))
    merged['pdf_auto_fullscreen'] = bool(merged.get('pdf_auto_fullscreen', False))
    merged['keyword_instant'] = bool(merged['keyword_instant'])
    merged['book_auto_focus_hymn'] = bool(merged['book_auto_focus_hymn'])
    mode = str(merged.get('hymn_num_mode') or HYMN_NUM_MODE_CONTAINS).strip().lower()
    merged['hymn_num_mode'] = mode if mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS
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
    merged['cloudflared_path'] = str(merged.get('cloudflared_path') or '')
    if merged.get('tunnel_mode') == 'named':
        merged['tunnel_mode'] = 'service'
    if merged.get('tunnel_mode') not in ('quick', 'service'):
        merged['tunnel_mode'] = 'service'
    raw_dict = raw if isinstance(raw, dict) else {}
    if 'enable_tunnel_fixed' not in raw_dict and 'enable_tunnel_quick' not in raw_dict:
        # Migrate legacy single enable_tunnel + tunnel_mode.
        et = bool(merged.get('enable_tunnel'))
        if et and merged.get('tunnel_mode') == 'quick':
            merged['enable_tunnel_fixed'] = False
            merged['enable_tunnel_quick'] = True
        elif et:
            merged['enable_tunnel_fixed'] = True
            merged['enable_tunnel_quick'] = False
        else:
            merged['enable_tunnel_fixed'] = False
            merged['enable_tunnel_quick'] = False
    else:
        merged['enable_tunnel_fixed'] = bool(merged.get('enable_tunnel_fixed'))
        merged['enable_tunnel_quick'] = bool(merged.get('enable_tunnel_quick'))
    merged['enable_tunnel'] = bool(
        merged['enable_tunnel_fixed'] or merged['enable_tunnel_quick']
    )
    if merged['enable_tunnel_fixed']:
        merged['tunnel_mode'] = 'service'
    elif merged['enable_tunnel_quick']:
        merged['tunnel_mode'] = 'quick'
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
    if isinstance(raw, dict) and 'show_qr_lan' not in raw and 'show_qr_code' in raw:
        merged['show_qr_lan'] = bool(raw.get('show_qr_code'))
    else:
        merged['show_qr_lan'] = bool(merged.get('show_qr_lan', True))
    if isinstance(raw, dict) and 'show_qr_fixed' not in raw and 'show_qr_quick' not in raw:
        wan_flag = None
        if 'show_qr_wan' in raw:
            wan_flag = bool(raw.get('show_qr_wan'))
        elif 'show_qr_code' in raw:
            wan_flag = bool(raw.get('show_qr_code'))
        if wan_flag is None:
            wan_flag = True
        merged['show_qr_fixed'] = wan_flag
        merged['show_qr_quick'] = wan_flag
    else:
        merged['show_qr_fixed'] = bool(merged.get('show_qr_fixed', True))
        merged['show_qr_quick'] = bool(merged.get('show_qr_quick', True))
    merged['show_qr_wan'] = bool(
        merged.get('show_qr_fixed') or merged.get('show_qr_quick')
    )
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


def _normalize_overlay_duration(value):
    if value == OVERLAY_DURATION_ALWAYS:
        return OVERLAY_DURATION_ALWAYS
    try:
        n = int(value)
    except (TypeError, ValueError):
        return DEFAULT_OVERLAY_SECS
    if n == OVERLAY_DURATION_ALWAYS:
        return OVERLAY_DURATION_ALWAYS
    return max(1, min(MAX_OVERLAY_SECS, n))


def _overlay_mode_index(mode):
    return 1 if mode == OVERLAY_MODE_CENTER else 0


def _normalize_center_secs(value):
    try:
        return max(1, min(MAX_OVERLAY_CENTER_SECS, int(value)))
    except (TypeError, ValueError):
        return CENTER_OVERLAY_SECS


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


def _overlay_style_from_settings(settings):
    s = settings or {}
    text_color_raw = str(s.get('overlay_text_color') or s.get('text_color') or '').strip()
    bg_raw = s.get('overlay_bg_color', s.get('bg_color'))
    return {
        'bg_color': _normalize_hex_color(bg_raw, DEFAULT_OVERLAY_BG_COLOR),
        'text_color': (
            _normalize_hex_color(text_color_raw, DEFAULT_OVERLAY_TEXT_COLOR)
            if text_color_raw else DEFAULT_OVERLAY_TEXT_COLOR
        ),
    }


def _overlay_style_is_normalized(style):
    return isinstance(style, dict) and 'bg_color' in style


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
    """Return True if lyric windows (or a legacy Word/PDF viewer) are still open."""
    check = getattr(document_viewer_is_open, '_lyric_open_check', None)
    if callable(check):
        return bool(check())
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
        self._style = _overlay_style_from_settings({})
        self._bg_qcolor = QColor(self._style['bg_color'])
        self._radius = 12
        self._placement = 'corner'
        self._mode = OVERLAY_MODE_CORNER
        self._corner_duration = DEFAULT_OVERLAY_SECS
        self._center_secs = CENTER_OVERLAY_SECS
        self._animating = False
        self._anim_start_rect = None
        self._anim_end_rect = None
        self._book_label = QLabel(self)
        self._book_label.setObjectName('overlayBookLabel')
        self._book_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._item_label = QLabel(self)
        self._item_label.setObjectName('overlayItemLabel')
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

    def _resolve_text_color(self):
        raw = str(self._style.get('text_color') or '').strip()
        if raw:
            return raw
        return _text_color_for_bg(self._style.get('bg_color', DEFAULT_OVERLAY_BG_COLOR))

    def _apply_label_text_color(self, label, text_color, extra_style=''):
        label.setStyleSheet(
            f"QLabel#{label.objectName()} {{ color: {text_color}; background: transparent; {extra_style} }}"
        )
        pal = label.palette()
        pal.setColor(QPalette.ColorRole.WindowText, QColor(text_color))
        label.setPalette(pal)

    def _apply_style(self, style=None):
        if _overlay_style_is_normalized(style):
            self._style = dict(style)
        else:
            self._style = _overlay_style_from_settings(style or {})
        self._bg_qcolor = QColor(self._style['bg_color'])
        if self._placement == 'center':
            self._apply_center_style()
        else:
            self._apply_corner_style()
        self.update()

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
        text_color = self._resolve_text_color()
        self._apply_label_text_color(
            self._book_label, text_color,
            f"font-size: {book_fs}px; font-weight: bold;",
        )
        self._apply_label_text_color(
            self._item_label, text_color,
            f"font-size: {item_fs}px;",
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
        text_color = self._resolve_text_color()
        self._apply_label_text_color(
            self._book_label, text_color,
            "font-size: 24px; font-weight: bold;",
        )
        self._apply_label_text_color(
            self._item_label, text_color,
            "font-size: 18px;",
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
        text_color = self._resolve_text_color()
        self._apply_label_text_color(
            self._book_label, text_color,
            "font-size: 48px; font-weight: bold;",
        )
        self._apply_label_text_color(
            self._item_label, text_color,
            "font-size: 32px;",
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
        painter.fillPath(path, self._bg_qcolor)
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
                  mode=OVERLAY_MODE_CORNER, style=None, center_secs=CENTER_OVERLAY_SECS):
        self._apply_style(style)
        self._corner_duration = self._normalize_duration(duration_sec)
        self._center_secs = _normalize_center_secs(center_secs)
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
            self._phase_timer.start(self._center_secs * 1000)
        else:
            self._start_corner_hide_timer()

    def _switch_to_corner(self):
        BookNameOverlay._on_center_elapsed()
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
    _ends_at = 0.0
    _ends_at_unix = 0.0
    _on_finished = None
    _center_notified = False

    @classmethod
    def _set_ends_at(cls, duration_sec, mode, center_secs=CENTER_OVERLAY_SECS):
        cls._center_notified = False
        if mode != OVERLAY_MODE_CENTER:
            cls._ends_at = 0.0
            cls._ends_at_unix = 0.0
            return
        hold = _normalize_center_secs(center_secs)
        cls._ends_at = time.monotonic() + hold
        cls._ends_at_unix = time.time() + hold

    @classmethod
    def ends_at_unix(cls):
        return cls._ends_at_unix

    @classmethod
    def _on_center_elapsed(cls):
        if cls._center_notified:
            return
        cls._center_notified = True
        cls._ends_at = 0.0
        cls._ends_at_unix = 0.0
        if cls._on_finished:
            try:
                cls._on_finished()
            except Exception:
                pass

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
            # Already shown once; after topology change resume at corner only.
            'mode': OVERLAY_MODE_CORNER,
            'style': dict(ref._style),
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
                    session.get('style'),
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
                  mode=OVERLAY_MODE_CORNER, style=None, center_secs=CENTER_OVERLAY_SECS):
        if not book_name:
            return
        cls._sync_panels()
        if not cls._panels:
            return
        cls._set_ends_at(duration_sec, mode, center_secs)
        for panel in cls._panels:
            panel._activate(
                book_name, watch_path, item_name, duration_sec, mode, style, center_secs,
            )
        cls._start_document_watch(watch_path)

    @classmethod
    def hide_overlay(cls, finished=False):
        cls._ends_at = 0.0
        cls._ends_at_unix = 0.0
        cls._stop_document_watch()
        for panel in cls._panels:
            panel._deactivate()


class _BlackScreenPanel(QWidget):
    """One-screen black cover; Esc / double-click dismisses all."""

    def __init__(self, screen):
        super().__init__(None)
        self._screen = screen
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setStyleSheet('background: #000000;')
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            BlackScreenOverlay.hide()
            event.accept()
            return
        super().keyPressEvent(event)

    def mouseDoubleClickEvent(self, event):
        BlackScreenOverlay.hide()
        event.accept()


class BlackScreenOverlay:
    """Fullscreen black cover on all screens until next open or Esc."""

    _panels = []

    @classmethod
    def _sync_panels(cls):
        screens = QApplication.screens() or []
        screen_set = set(screens)
        for panel in list(cls._panels):
            if panel._screen not in screen_set:
                panel.hide()
                panel.deleteLater()
                cls._panels.remove(panel)
        by_screen = {p._screen: p for p in cls._panels}
        for screen in screens:
            if screen not in by_screen:
                cls._panels.append(_BlackScreenPanel(screen))

    @classmethod
    def show(cls):
        cls._sync_panels()
        primary = QApplication.primaryScreen()
        focus_panel = None
        for panel in cls._panels:
            geo = panel._screen.geometry()
            panel.setGeometry(geo)
            panel.show()
            panel.raise_()
            if panel._screen is primary:
                focus_panel = panel
        if focus_panel is None and cls._panels:
            focus_panel = cls._panels[0]
        if focus_panel:
            focus_panel.activateWindow()
            focus_panel.setFocus(Qt.FocusReason.OtherFocusReason)

    @classmethod
    def hide(cls):
        for panel in cls._panels:
            panel.hide()

    @classmethod
    def is_visible(cls):
        return any(p.isVisible() for p in cls._panels)


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


def _desktop_qr_caption_height(caption, text_size=DEFAULT_QR_TEXT_SIZE):
    text = str(caption or '').strip()
    if not text:
        return 0
    lines = max(1, text.count('\n') + 1)
    line_h = max(14, int(text_size) + 4)
    return lines * line_h + 6


def _desktop_qr_render_size(card_w, card_h, caption='', text_size=DEFAULT_QR_TEXT_SIZE):
    """QR 像素邊長（緊湊內距；caption 可含多行網址）。"""
    pad_x, pad_top, pad_bottom, spacing = 3, 3, 2, 2
    caption_h = _desktop_qr_caption_height(caption, text_size)
    if caption_h <= 0:
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
            self._card_w, self._card_h, caption_text, self._text_size,
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
    """管理各螢幕桌面 QR Code 浮層（可同時顯示內網／外網；每屏各一組）。"""

    _roles = {}

    @classmethod
    def _state(cls, role):
        role = str(role or 'default')
        if role not in cls._roles:
            cls._roles[role] = {'panels': [], 'cached': None}
        return cls._roles[role]

    @classmethod
    def _target_screens(cls, duplicate_only=False):
        screens = QApplication.screens() or []
        primary = QApplication.primaryScreen()
        if duplicate_only and primary and len(screens) > 1:
            return [primary]
        return screens

    @classmethod
    def _sync_panels(cls, role, duplicate_only=False):
        st = cls._state(role)
        screens = cls._target_screens(duplicate_only)
        screen_set = set(screens)
        for panel in list(st['panels']):
            if panel._screen not in screen_set:
                panel.hide_qr()
                panel.deleteLater()
                st['panels'].remove(panel)
        by_screen = {p._screen: p for p in st['panels']}
        for screen in screens:
            if screen not in by_screen:
                st['panels'].append(_DesktopQrPanel(screen))

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
    def show_qr(cls, pixmap, caption, style, x, y, width, height, role='default'):
        style = style or {}
        st = cls._state(role)
        st['cached'] = {
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
        cls._sync_panels(role, duplicate_only=False)
        for panel in st['panels']:
            cls._apply_to_panel(panel, st['cached'])

    @classmethod
    def reposition(cls, x, y, width=None, height=None, role='default'):
        st = cls._state(role)
        if st['cached']:
            st['cached']['x'] = x
            st['cached']['y'] = y
            if width is not None:
                st['cached']['width'] = width
            if height is not None:
                st['cached']['height'] = height
        if st['cached'] and width is not None and height is not None and st['panels']:
            cached = st['cached']
            for panel in st['panels']:
                if panel.isVisible():
                    cls._apply_to_panel(panel, cached)
            return
        for panel in st['panels']:
            if panel.isVisible():
                panel._position(x, y)

    @classmethod
    def hide(cls, role=None):
        if role is None:
            for r in list(cls._roles):
                cls.hide(r)
            return
        st = cls._state(role)
        st['cached'] = None
        for panel in st['panels']:
            panel.hide_qr()

    @classmethod
    def refresh_after_display_change(cls):
        for role, st in list(cls._roles.items()):
            cached = st.get('cached')
            if not cached:
                continue
            for panel in list(st['panels']):
                panel.hide_qr()
                panel.deleteLater()
            st['panels'] = []
            cls._sync_panels(role, duplicate_only=True)
            for panel in st['panels']:
                cls._apply_to_panel(panel, cached)


def open_docx(path):
    """以系統預設方式開啟 Word 檔。"""
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 3)
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
        self._font_size = 13
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
        fs = self._font_size
        pad = 12 if fs >= 20 else 8
        self.setStyleSheet(f"""
            QListWidget {{
                background: {t['panel']}; border: 1.5px solid {t['accent']};
                border-radius: 10px; outline: none; font-size: {fs}px; padding: 4px;
            }}
            QListWidget::item {{
                padding: {pad}px 12px; color: {t['text']};
                border-radius: 6px; margin: 1px 2px;
            }}
            QListWidget::item:hover  {{ background: {t['panel2']}; }}
            QListWidget::item:selected {{ background: {t['accent']}; color: {t['a_text']}; }}
        """)

    def set_theme(self, theme, font_size=None):
        self._theme = theme
        if font_size is not None:
            self._font_size = max(10, int(font_size))
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
    display_duplicate_requested = pyqtSignal(int)  # generation — 全螢幕後安全觸發顯示模式切換
    projection_ready = pyqtSignal(int)  # 背景執行緒 → 主執行緒：全螢幕就緒
    open_flow_esc_cancelled = pyqtSignal()  # Esc 取消開檔流程
    overlay_refresh_requested = pyqtSignal()  # 顯示模式切換後重新定位 overlay
    remote_open_payload = pyqtSignal(dict)
    remote_populate_ui = pyqtSignal(str, str, list)
    remote_pending_request = pyqtSignal(object)
    remote_display_action = pyqtSignal(str)  # close | black
    remote_lyric_control = pyqtSignal(dict)
    remote_prefs = pyqtSignal(dict)
    tunnel_url_signal = pyqtSignal(str, str)  # role, url
    tunnel_ready_signal = pyqtSignal(str, str)
    tunnel_error_signal = pyqtSignal(str, str)
    tunnel_expired_signal = pyqtSignal(str, str)
    service_install_signal = pyqtSignal(bool, str)
    _main_invoke = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.books        = []
        self.current_book = None
        self._loading_settings = True
        self._settings    = load_settings()
        self.hymn_folder  = resolve_hymn_folder(self._settings)
        self.hymns_json_path = resolve_hymns_json_path(self.hymn_folder)
        self._lyric_pair = LyricWindowPair()
        self._lyric_pair.save_defaults_cb = self._save_lyric_best_defaults
        self._lyric_pair.restore_defaults_cb = self._restore_lyric_best_defaults
        self._lyric_pair.go_live_cb = self._on_lyric_go_live
        self._lyric_pair.session.style_changed.connect(self._on_lyric_style_changed)
        self._preview_sync_timer = QTimer(self)
        self._preview_sync_timer.setSingleShot(True)
        self._preview_sync_timer.setInterval(80)
        self._preview_sync_timer.timeout.connect(self._sync_lyric_preview)
        self._lyric_layout_debounce = QTimer(self)
        self._lyric_layout_debounce.setSingleShot(True)
        self._lyric_layout_debounce.setInterval(120)
        self._lyric_layout_debounce.timeout.connect(self._apply_lyric_layout_from_settings)
        self._lyric_appearance_debounce = QTimer(self)
        self._lyric_appearance_debounce.setSingleShot(True)
        self._lyric_appearance_debounce.setInterval(120)
        self._lyric_appearance_debounce.timeout.connect(self._apply_lyric_appearance_from_settings)
        self._lyric_style_persist_debounce = QTimer(self)
        self._lyric_style_persist_debounce.setSingleShot(True)
        self._lyric_style_persist_debounce.setInterval(250)
        self._lyric_style_persist_debounce.timeout.connect(self._persist_lyric_style_from_session)
        self._lyric_ui_applying = False
        self._overlay_count_timer = QTimer(self)
        self._overlay_count_timer.setInterval(250)
        self._overlay_count_timer.timeout.connect(self._tick_overlay_countdown)
        self._lyric_pair.session.content_changed.connect(self._schedule_lyric_preview)
        self._lyric_pair.session.view_changed.connect(self._schedule_lyric_preview)
        document_viewer_is_open._lyric_open_check = self._lyric_pair.is_open
        self.search_mode  = self._settings['search_mode']
        self.theme_name   = self._settings['theme']
        self.font_size    = self._settings['font_size']
        self.senior_mode  = bool(self._settings.get('senior_mode'))
        self._search_worker = None
        self._index_worker = None
        self.content_index = None
        self._app         = QApplication.instance()
        self._remote      = RemoteServer()
        self._remote_pending_id = None
        self._tunnel_fixed = None
        self._tunnel_quick = None
        self._tunnel = None  # compat alias → last active tunnel
        self.enable_tunnel_fixed = bool(self._settings.get('enable_tunnel_fixed'))
        self.enable_tunnel_quick = bool(self._settings.get('enable_tunnel_quick'))
        self.enable_tunnel = self.enable_tunnel_fixed or self.enable_tunnel_quick
        self.cloudflared_path = self._settings['cloudflared_path']
        self.tunnel_mode = self._settings['tunnel_mode']
        self.cloudflared_token = self._settings['cloudflared_token']
        self.named_tunnel_url = self._settings['named_tunnel_url']
        self.remote_path_prefix = self._settings['remote_path_prefix']
        self.remote_use_https = self._settings['remote_use_https']
        self._tunnel_status = {'fixed': '—', 'quick': '—'}
        self._open_flow = OpenFlowController()
        self._open_flow_generation = 0
        self._open_flow.set_on_esc_cancelled(self.open_flow_esc_cancelled.emit)
        self._last_opened_path = ''

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
        self.projection_ready.connect(self._on_projection_ready)
        self.open_flow_esc_cancelled.connect(
            self._hide_overlay_on_esc, Qt.ConnectionType.QueuedConnection,
        )
        self.overlay_refresh_requested.connect(
            self._refresh_overlays_after_display, Qt.ConnectionType.QueuedConnection,
        )
        self.remote_open_payload.connect(self._on_remote_open_payload)
        self.remote_populate_ui.connect(self._on_remote_populate_ui)
        self.remote_pending_request.connect(self._on_remote_pending_request)
        self.remote_display_action.connect(self._on_remote_display_action)
        self.remote_lyric_control.connect(self._on_remote_lyric_control)
        self.remote_prefs.connect(self._on_remote_prefs)
        BookNameOverlay._on_finished = self._on_overlay_countdown_finished
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
            lambda action: self.remote_display_action.emit(action),
            lambda payload: self.remote_lyric_control.emit(payload),
        )
        if hasattr(self._remote.state, 'set_preview_needed_handler'):
            self._remote.state.set_preview_needed_handler(
                lambda: self._main_invoke.emit(self._schedule_lyric_preview)
            )
        if hasattr(self._remote.state, 'set_prefs_handler'):
            self._remote.state.set_prefs_handler(
                lambda payload: self.remote_prefs.emit(payload)
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
            (Qt.Key.Key_F1, 'standard'),
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
        sc_esc = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        sc_esc.setContext(Qt.ShortcutContext.ApplicationShortcut)
        sc_esc.activated.connect(self._dismiss_black_screen_shortcut)

    def _dismiss_black_screen_shortcut(self):
        if BlackScreenOverlay.is_visible():
            BlackScreenOverlay.hide()
            self._show_toast('已取消全黑', 'ok')
            return
        pair = getattr(self, '_lyric_pair', None)
        if pair is not None and pair.exit_display_fullscreen():
            return

    def _focus_book_input(self):
        if self.search_mode not in BOOK_UI_MODES:
            self._set_search_mode('standard')
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

        self.lyric_dock = QWidget()
        self.lyric_dock.setMinimumWidth(320)
        self.lyric_dock_lay = QVBoxLayout(self.lyric_dock)
        self.lyric_dock_lay.setContentsMargins(0, 0, 0, 0)
        self.lyric_dock_lay.setSpacing(0)
        self.lyric_dock.hide()

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.addWidget(body_widget)
        self.main_splitter.addWidget(self.lyric_dock)
        self.main_splitter.setStretchFactor(0, 3)
        self.main_splitter.setStretchFactor(1, 2)
        self.body_container_splitter = self.main_splitter
        container_lay.addWidget(self.main_splitter, 1)
        self._lyric_pair.dock_host = self.lyric_dock

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

        self.btn_senior = QPushButton("簡易版")
        self.btn_senior.setCheckable(True)
        self.btn_senior.setChecked(self.senior_mode)
        self.btn_senior.setFixedHeight(30)
        self.btn_senior.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_senior.setToolTip("一鍵切換大字簡化介面")
        self.btn_senior.toggled.connect(self._on_senior_mode_toggled)
        h.addWidget(self.btn_senior)

        h.addSpacing(8)

        self.btn_phone_control = QPushButton("手機控制")
        self.btn_phone_control.setFixedHeight(30)
        self.btn_phone_control.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_phone_control.setToolTip("內網／外網 QR，用手機遙控")
        self.btn_phone_control.clicked.connect(self._show_phone_control_wizard)
        h.addWidget(self.btn_phone_control)

        h.addSpacing(8)

        self.header_opts = QWidget()
        opts = QHBoxLayout(self.header_opts)
        opts.setContentsMargins(0, 0, 0, 0)
        opts.setSpacing(18)

        self.chk_open_overlay_hdr = QCheckBox("書名提示")
        self.chk_open_overlay_hdr.setChecked(self._settings['open_overlay'])
        self.chk_open_overlay_hdr.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_open_overlay_hdr.setToolTip("開啟 Word / PDF 時顯示置頂書名提示")
        self.chk_open_overlay_hdr.toggled.connect(self._on_open_overlay_hdr_toggled)
        opts.addWidget(self.chk_open_overlay_hdr)

        self.chk_click_to_open_hdr = QCheckBox("一點即開")
        self.chk_click_to_open_hdr.setChecked(self._settings['click_to_open'])
        self.chk_click_to_open_hdr.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_click_to_open_hdr.toggled.connect(self._on_click_to_open_hdr_toggled)
        opts.addWidget(self.chk_click_to_open_hdr)
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
        self.footer_frame.setFixedHeight(34)
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

        self.chk_lyric_preview_mode = QCheckBox('預覽模式')
        self.chk_lyric_preview_mode.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_lyric_preview_mode.setToolTip(
            '開啟後，揀詩歌先進右邊預覽（唔出街）；調好先按「轉出現場」'
        )
        self.chk_lyric_preview_mode.setChecked(bool(self._settings.get('lyric_preview_mode', False)))
        self.chk_lyric_preview_mode.toggled.connect(self._on_lyric_preview_mode_toggled)
        fl.addWidget(self.chk_lyric_preview_mode)

        return self.footer_frame

    # ─────────────────────────────────────────────────────────────
    #  SIDEBAR — 左側書冊列表（固定 240px，不可拖曳調整）
    # ─────────────────────────────────────────────────────────────
    def _build_sidebar(self):
        self.sidebar_frame = QFrame()
        self.sidebar_frame.setFixedWidth(SENIOR_SIDEBAR_W if self.senior_mode else NORMAL_SIDEBAR_W)
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

        # ── 模式切換：標準 / 全局 / 關鍵字 / 排程 ────────
        self.mode_row_widget = QWidget()
        mode_row = QHBoxLayout(self.mode_row_widget)
        mode_row.setContentsMargins(0, 0, 0, 0)
        mode_row.setSpacing(8)
        self.btn_standard = QPushButton("⭐  標準模式 (F1)")
        self.btn_book     = QPushButton("📖  書本模式")
        self.btn_book.setVisible(False)
        self.btn_global   = QPushButton("🌐  全局模式 (F2)")
        self.btn_keyword  = QPushButton("🔤  關鍵字模式 (F3)")
        self.btn_schedule = QPushButton("📋  排程模式 (F4)")
        for btn in (self.btn_standard, self.btn_global,
                    self.btn_keyword, self.btn_schedule):
            btn.setCheckable(True)
            btn.setFixedHeight(30)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_standard.clicked.connect(lambda: self._set_search_mode('standard'))
        self.btn_global.clicked.connect(lambda: self._set_search_mode('global'))
        self.btn_keyword.clicked.connect(lambda: self._set_search_mode('keyword'))
        self.btn_schedule.clicked.connect(lambda: self._set_search_mode('schedule'))
        mode_row.addWidget(self.btn_standard)
        mode_row.addWidget(self.btn_global)
        mode_row.addWidget(self.btn_keyword)
        mode_row.addWidget(self.btn_schedule)
        mode_row.addStretch()
        lay.addWidget(self.mode_row_widget)

        # ── 標準／書本：書冊 + 詩歌號/名 + Filter/Open 按鈕（上下兩行）──
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

        self.btn_voice = QPushButton("🎤")
        self.btn_voice.setFixedSize(40, 40)
        self.btn_voice.setToolTip("語音輸入（廣東話）— 再按一次可停止")
        self.btn_voice.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_voice.clicked.connect(self._toggle_voice_input)
        if not voice_input_available():
            self.btn_voice.setEnabled(False)
            self.btn_voice.setToolTip("未安裝語音辨識套件")

        self.btn_filter = QPushButton("Filter")
        self.btn_filter.setFixedHeight(40)
        self.btn_filter.setFixedWidth(84)
        self.btn_filter.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_filter.clicked.connect(self._on_book_filter_action)

        bar.addWidget(self.btn_sess_prev)
        bar.addWidget(self.btn_sess_next)
        bar.addWidget(self.inp_num, 1)
        bar.addWidget(self.btn_voice)
        bar.addWidget(self.btn_filter)

        self._voice = VoiceInputController(self, listen_seconds=5.0)
        self._voice.transcript.connect(self._on_voice_transcript)
        self._voice.error.connect(self._on_voice_error)
        self._voice.status.connect(self._on_voice_status)
        self._voice.listening_changed.connect(self._on_voice_listening_changed)

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
        self.inp_keyword.setPlaceholderText("🔍  輸入關鍵字  e.g. 恩典  —  搜索歌詞")
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

        self.grp_data = QGroupBox("詩歌資料")
        data_grid = QGridLayout(self.grp_data)
        data_grid.setContentsMargins(*grid_margins)
        data_grid.setHorizontalSpacing(14)
        data_grid.setVerticalSpacing(12)
        data_grid.setColumnStretch(1, 1)
        data_grid.setColumnMinimumWidth(0, label_min_w)

        self.lbl_hymn_folder = QLabel("路徑")
        self.lbl_hymn_folder.setMinimumWidth(label_min_w)
        self.lbl_hymn_folder.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.inp_hymn_folder = QLineEdit(self.hymns_json_path)
        self.inp_hymn_folder.setPlaceholderText(default_hymns_json_path())
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
        self.btn_edit_hymns = QPushButton("管理詩歌庫")
        self.btn_edit_hymns.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_edit_hymns.setToolTip("用介面新增、修改、刪除 hymns.json 裡的詩歌")
        self.btn_edit_hymns.clicked.connect(lambda: self._open_hymn_editor())
        data_btn_row = QHBoxLayout()
        data_btn_row.setSpacing(8)
        data_btn_row.addWidget(self.btn_apply_hymn_folder)
        data_btn_row.addWidget(self.btn_edit_hymns)
        data_btn_row.addStretch()

        data_grid.addWidget(self.lbl_hymn_folder, 0, 0)
        data_grid.addLayout(folder_row, 0, 1)
        data_grid.addLayout(data_btn_row, 1, 1)

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
        self.btn_font_minus_settings = QPushButton("−")
        self.btn_font_minus_settings.setFixedSize(36, 32)
        self.btn_font_minus_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_font_minus_settings.setToolTip("縮小字體")
        self.btn_font_minus_settings.clicked.connect(lambda: self._nudge_font(-1))
        self.fs_lbl = QLabel(f"{self.font_size}px")
        self.fs_lbl.setFixedWidth(48)
        self.fs_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.btn_font_plus_settings = QPushButton("+")
        self.btn_font_plus_settings.setFixedSize(36, 32)
        self.btn_font_plus_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_font_plus_settings.setToolTip("放大字體")
        self.btn_font_plus_settings.clicked.connect(lambda: self._nudge_font(1))
        font_row.addWidget(self.btn_font_minus_settings)
        font_row.addWidget(self.fs_lbl)
        font_row.addWidget(self.btn_font_plus_settings)
        font_row.addStretch(1)

        appearance_grid.addWidget(self.lbl_theme, 0, 0)
        appearance_grid.addWidget(self.theme_combo, 0, 1, Qt.AlignmentFlag.AlignLeft)
        appearance_grid.addWidget(self.lbl_font, 1, 0)
        appearance_grid.addLayout(font_row, 1, 1)

        self.chk_book_auto_focus_hymn = QCheckBox("唯一書冊結果跳至詩歌號/名")
        self.chk_book_auto_focus_hymn.setChecked(self._settings['book_auto_focus_hymn'])
        self.chk_book_auto_focus_hymn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_book_auto_focus_hymn.toggled.connect(self._on_book_auto_focus_hymn_toggled)
        appearance_grid.addWidget(self.chk_book_auto_focus_hymn, 2, 0, 1, 2)

        self.lbl_hymn_num_mode = QLabel("數字模式")
        self.lbl_hymn_num_mode.setMinimumWidth(label_min_w)
        self.lbl_hymn_num_mode.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_hymn_num_mode = QComboBox()
        self.combo_hymn_num_mode.addItem("模糊（含部分數字）", HYMN_NUM_MODE_CONTAINS)
        self.combo_hymn_num_mode.addItem("精確編號（4≠14）", HYMN_NUM_MODE_EXACT)
        idx = self.combo_hymn_num_mode.findData(
            self._settings.get('hymn_num_mode', HYMN_NUM_MODE_CONTAINS)
        )
        self.combo_hymn_num_mode.setCurrentIndex(idx if idx >= 0 else 0)
        self.combo_hymn_num_mode.setMinimumWidth(200)
        self.combo_hymn_num_mode.setCursor(Qt.CursorShape.PointingHandCursor)
        self.combo_hymn_num_mode.setToolTip(
            "搜尋詩歌號時：模糊會把 4 也命中 14／24；精確只命中第 4 首（含 004、04）"
        )
        self.combo_hymn_num_mode.currentIndexChanged.connect(self._on_hymn_num_mode_changed)
        appearance_grid.addWidget(self.lbl_hymn_num_mode, 3, 0)
        appearance_grid.addWidget(self.combo_hymn_num_mode, 3, 1, Qt.AlignmentFlag.AlignLeft)

        self.chk_click_to_open = QCheckBox("一點即開檔案（單擊列表即開啟）")
        self.chk_click_to_open.setChecked(self._settings['click_to_open'])
        self.chk_click_to_open.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_click_to_open.toggled.connect(self._on_click_to_open_settings_toggled)
        appearance_grid.addWidget(self.chk_click_to_open, 4, 0, 1, 2)

        self.chk_show_preview = QCheckBox("顯示檔案預覽")
        self.chk_show_preview.setChecked(self._settings.get('show_preview', True))
        self.chk_show_preview.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_show_preview.toggled.connect(self._on_show_preview_toggled)
        appearance_grid.addWidget(self.chk_show_preview, 5, 0, 1, 2)

        self.chk_auto_rescan = QCheckBox("資料夾變更時自動重新掃描")
        self.chk_auto_rescan.setChecked(self._settings.get('auto_rescan', True))
        self.chk_auto_rescan.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_auto_rescan.toggled.connect(self._on_auto_rescan_toggled)
        appearance_grid.addWidget(self.chk_auto_rescan, 6, 0, 1, 2)

        self.chk_startup_tray = QCheckBox("啟動時最小化到系統匣")
        self.chk_startup_tray.setChecked(self._settings.get('startup_tray', False))
        self.chk_startup_tray.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_startup_tray.toggled.connect(self._on_startup_tray_toggled)
        appearance_grid.addWidget(self.chk_startup_tray, 7, 0, 1, 2)

        self.chk_minimize_tray = QCheckBox("關閉視窗時最小化到系統匣")
        self.chk_minimize_tray.setChecked(self._settings.get('minimize_to_tray', False))
        self.chk_minimize_tray.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_minimize_tray.toggled.connect(self._on_minimize_tray_toggled)
        appearance_grid.addWidget(self.chk_minimize_tray, 8, 0, 1, 2)

        self.lbl_lyric_bg = QLabel("歌詞背景")
        self.lbl_lyric_bg.setMinimumWidth(label_min_w)
        self.overlay_lyric_bg_row, self.inp_lyric_bg_color, self.lbl_lyric_bg_preview, self.btn_lyric_bg = (
            self._make_qr_color_row(
                self._settings.get('lyric_bg_color', DEFAULT_LYRIC_BG),
                DEFAULT_LYRIC_BG, '展示／控制歌詞背景色',
            )
        )
        self.inp_lyric_bg_color.textChanged.connect(self._on_lyric_appearance_changed)
        self.btn_lyric_bg.clicked.connect(lambda: self._pick_qr_color(
            self.inp_lyric_bg_color, self.lbl_lyric_bg_preview, self._on_lyric_appearance_changed,
        ))
        lyric_bg_box = QWidget()
        lyric_bg_box_lay = QVBoxLayout(lyric_bg_box)
        lyric_bg_box_lay.setContentsMargins(0, 0, 0, 0)
        lyric_bg_box_lay.setSpacing(6)
        lyric_bg_box_lay.addWidget(self.overlay_lyric_bg_row)
        lyric_bg_img_row = QHBoxLayout()
        lyric_bg_img_row.setSpacing(8)
        self.btn_lyric_bg_image = QPushButton('上傳圖片')
        self.btn_lyric_bg_image.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_lyric_bg_image.setToolTip('用圖片代替純色背景')
        self.btn_lyric_bg_image.clicked.connect(self._on_lyric_bg_image_upload)
        self.btn_lyric_bg_image_clear = QPushButton('刪除圖片')
        self.btn_lyric_bg_image_clear.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_lyric_bg_image_clear.setToolTip('刪除背景圖，改回純色')
        self.btn_lyric_bg_image_clear.clicked.connect(self._on_lyric_bg_image_clear)
        self.lbl_lyric_bg_image = QLabel('純色')
        lyric_bg_img_row.addWidget(self.btn_lyric_bg_image)
        lyric_bg_img_row.addWidget(self.btn_lyric_bg_image_clear)
        lyric_bg_img_row.addWidget(self.lbl_lyric_bg_image, 1)
        lyric_bg_box_lay.addLayout(lyric_bg_img_row)
        appearance_grid.addWidget(self.lbl_lyric_bg, 9, 0)
        appearance_grid.addWidget(lyric_bg_box, 9, 1)

        self.lbl_lyric_text = QLabel("歌詞文字")
        self.lbl_lyric_text.setMinimumWidth(label_min_w)
        self.overlay_lyric_text_row, self.inp_lyric_text_color, self.lbl_lyric_text_preview, self.btn_lyric_text = (
            self._make_qr_color_row(
                self._settings.get('lyric_text_color', DEFAULT_LYRIC_TEXT),
                DEFAULT_LYRIC_TEXT, '展示／控制歌詞文字色',
            )
        )
        self.inp_lyric_text_color.textChanged.connect(self._on_lyric_appearance_changed)
        self.btn_lyric_text.clicked.connect(lambda: self._pick_qr_color(
            self.inp_lyric_text_color, self.lbl_lyric_text_preview, self._on_lyric_appearance_changed,
        ))
        appearance_grid.addWidget(self.lbl_lyric_text, 10, 0)
        appearance_grid.addWidget(self.overlay_lyric_text_row, 10, 1)

        self.lbl_lyric_align = QLabel("歌詞對齊")
        self.lbl_lyric_align.setMinimumWidth(label_min_w)
        self.combo_lyric_align = QComboBox()
        self.combo_lyric_align.addItem('靠左', ALIGN_LEFT)
        self.combo_lyric_align.addItem('左右對齊', ALIGN_JUSTIFY)
        idx = self.combo_lyric_align.findData(
            self._settings.get('lyric_align', DEFAULT_LYRIC_ALIGN)
        )
        self.combo_lyric_align.setCurrentIndex(idx if idx >= 0 else 0)
        self.combo_lyric_align.currentIndexChanged.connect(self._on_lyric_layout_settings_changed)
        appearance_grid.addWidget(self.lbl_lyric_align, 11, 0)
        appearance_grid.addWidget(self.combo_lyric_align, 11, 1, Qt.AlignmentFlag.AlignLeft)

        self.lbl_lyric_spacing = QLabel("歌詞字距")
        self.lbl_lyric_spacing.setMinimumWidth(label_min_w)
        spacing_row = QWidget()
        spacing_lay = QHBoxLayout(spacing_row)
        spacing_lay.setContentsMargins(0, 0, 0, 0)
        spacing_lay.setSpacing(8)
        self.slider_lyric_spacing = QSlider(Qt.Orientation.Horizontal)
        self.slider_lyric_spacing.setRange(LETTER_SPACING_MIN, LETTER_SPACING_MAX)
        self.slider_lyric_spacing.setSingleStep(5)
        self.slider_lyric_spacing.setPageStep(10)
        self.slider_lyric_spacing.setValue(
            normalize_letter_spacing(self._settings.get('lyric_letter_spacing'))
        )
        self.slider_lyric_spacing.valueChanged.connect(self._on_lyric_layout_settings_changed)
        self.lbl_lyric_spacing_val = QLabel(f'{self.slider_lyric_spacing.value()}%')
        self.lbl_lyric_spacing_val.setMinimumWidth(44)
        spacing_lay.addWidget(self.slider_lyric_spacing, 1)
        spacing_lay.addWidget(self.lbl_lyric_spacing_val)
        appearance_grid.addWidget(self.lbl_lyric_spacing, 12, 0)
        appearance_grid.addWidget(spacing_row, 12, 1)

        self.lbl_lyric_line_height = QLabel("歌詞行距")
        self.lbl_lyric_line_height.setMinimumWidth(label_min_w)
        line_height_row = QWidget()
        line_height_lay = QHBoxLayout(line_height_row)
        line_height_lay.setContentsMargins(0, 0, 0, 0)
        line_height_lay.setSpacing(8)
        self.slider_lyric_line_height = QSlider(Qt.Orientation.Horizontal)
        self.slider_lyric_line_height.setRange(LINE_HEIGHT_MIN, LINE_HEIGHT_MAX)
        self.slider_lyric_line_height.setSingleStep(5)
        self.slider_lyric_line_height.setPageStep(10)
        self.slider_lyric_line_height.setValue(
            normalize_line_height(self._settings.get('lyric_line_height'))
        )
        self.slider_lyric_line_height.setToolTip('行與行之間的距離（100% = 單倍行距）')
        self.slider_lyric_line_height.valueChanged.connect(self._on_lyric_layout_settings_changed)
        self.lbl_lyric_line_height_val = QLabel(f'{self.slider_lyric_line_height.value()}%')
        self.lbl_lyric_line_height_val.setMinimumWidth(52)
        line_height_lay.addWidget(self.slider_lyric_line_height, 1)
        line_height_lay.addWidget(self.lbl_lyric_line_height_val)
        appearance_grid.addWidget(self.lbl_lyric_line_height, 13, 0)
        appearance_grid.addWidget(line_height_row, 13, 1)

        self.lbl_lyric_margin = QLabel("歌詞邊距")
        self.lbl_lyric_margin.setMinimumWidth(label_min_w)
        margin_row = QWidget()
        margin_grid = QGridLayout(margin_row)
        margin_grid.setContentsMargins(0, 0, 0, 0)
        margin_grid.setHorizontalSpacing(10)
        margin_grid.setVerticalSpacing(6)
        saved_margins = normalize_lyric_margins(self._settings.get('lyric_margins'))
        self.spin_lyric_margins = {}
        side_labels = (
            ('top', '上', 0, 0),
            ('bottom', '下', 0, 2),
            ('left', '左', 1, 0),
            ('right', '右', 1, 2),
        )
        for side, label, row, col in side_labels:
            lbl = QLabel(label)
            lbl.setMinimumWidth(18)
            spin = QSpinBox()
            spin.setRange(LYRIC_MARGIN_MIN, LYRIC_MARGIN_MAX)
            spin.setSuffix(' px')
            spin.setSingleStep(4)
            spin.setValue(saved_margins[side])
            spin.setToolTip(f'歌詞與螢幕{label}邊的距離')
            spin.valueChanged.connect(self._on_lyric_layout_settings_changed)
            self.spin_lyric_margins[side] = spin
            margin_grid.addWidget(lbl, row, col)
            margin_grid.addWidget(spin, row, col + 1)
        appearance_grid.addWidget(self.lbl_lyric_margin, 14, 0)
        appearance_grid.addWidget(margin_row, 14, 1)

        self.lbl_lyric_font_family = QLabel("歌詞字體")
        self.lbl_lyric_font_family.setMinimumWidth(label_min_w)
        self.combo_lyric_font_family = QFontComboBox()
        self.combo_lyric_font_family.setMinimumWidth(220)
        self.combo_lyric_font_family.setCurrentFont(
            QFont(normalize_lyric_font_family(self._settings.get('lyric_font_family')))
        )
        self.combo_lyric_font_family.setToolTip('歌詞顯示用字體（需電腦已安裝）')
        self.combo_lyric_font_family.currentFontChanged.connect(
            self._on_lyric_layout_settings_changed
        )
        appearance_grid.addWidget(self.lbl_lyric_font_family, 15, 0)
        appearance_grid.addWidget(self.combo_lyric_font_family, 15, 1)

        self.lbl_lyric_font = QLabel("預設字級")
        self.lbl_lyric_font.setMinimumWidth(label_min_w)
        font_row = QWidget()
        font_lay = QHBoxLayout(font_row)
        font_lay.setContentsMargins(0, 0, 0, 0)
        font_lay.setSpacing(8)
        saved_font_pt = normalize_lyric_font_pt(self._settings.get('lyric_font_pt'))
        self.chk_lyric_font_auto = QCheckBox('自動填滿')
        self.chk_lyric_font_auto.setChecked(saved_font_pt <= 0)
        self.chk_lyric_font_auto.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_lyric_font_auto.setToolTip('開啟時自動調字級填滿螢幕；關閉則用右邊預設字級')
        self.chk_lyric_font_auto.toggled.connect(self._on_lyric_font_auto_toggled)
        self.spin_lyric_font = QSpinBox()
        self.spin_lyric_font.setRange(LYRIC_FONT_PT_MIN, LYRIC_FONT_PT_MAX)
        self.spin_lyric_font.setSuffix(' pt')
        self.spin_lyric_font.setValue(saved_font_pt if saved_font_pt > 0 else 48)
        self.spin_lyric_font.setEnabled(saved_font_pt > 0)
        self.spin_lyric_font.setToolTip('開啟詩歌時用呢個字級（仍可用 Ctrl+滾輪臨時放大縮小）')
        self.spin_lyric_font.valueChanged.connect(self._on_lyric_layout_settings_changed)
        font_lay.addWidget(self.chk_lyric_font_auto)
        font_lay.addWidget(self.spin_lyric_font)
        font_lay.addStretch(1)
        appearance_grid.addWidget(self.lbl_lyric_font, 16, 0)
        appearance_grid.addWidget(font_row, 16, 1)

        self.lbl_lyric_best = QLabel("最佳預設")
        self.lbl_lyric_best.setMinimumWidth(label_min_w)
        best_row = QWidget()
        best_lay = QHBoxLayout(best_row)
        best_lay.setContentsMargins(0, 0, 0, 0)
        best_lay.setSpacing(8)
        self.btn_lyric_save_best = QPushButton('保存目前為最佳')
        self.btn_lyric_save_best.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_lyric_save_best.setToolTip('將而家歌詞畫面存做一鍵還原用嘅最佳預設')
        self.btn_lyric_save_best.clicked.connect(self._save_lyric_best_defaults)
        self.btn_lyric_restore_best = QPushButton('還原最佳預設')
        self.btn_lyric_restore_best.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_lyric_restore_best.setToolTip('有問題時一鍵返去已保存嘅最佳預設')
        self.btn_lyric_restore_best.clicked.connect(self._restore_lyric_best_defaults)
        best_lay.addWidget(self.btn_lyric_save_best)
        best_lay.addWidget(self.btn_lyric_restore_best)
        best_lay.addStretch(1)
        appearance_grid.addWidget(self.lbl_lyric_best, 17, 0)
        appearance_grid.addWidget(best_row, 17, 1)

        self.lbl_lyric_placement = QLabel("控制畫面")
        self.lbl_lyric_placement.setMinimumWidth(label_min_w)
        self.combo_lyric_placement = QComboBox()
        self.combo_lyric_placement.addItem('獨立視窗', PLACEMENT_WINDOW)
        self.combo_lyric_placement.addItem('主視窗右邊', PLACEMENT_DOCKED)
        idx = self.combo_lyric_placement.findData(
            self._settings.get('lyric_control_placement', PLACEMENT_WINDOW)
        )
        self.combo_lyric_placement.setCurrentIndex(idx if idx >= 0 else 0)
        self.combo_lyric_placement.currentIndexChanged.connect(self._on_lyric_placement_changed)
        appearance_grid.addWidget(self.lbl_lyric_placement, 18, 0)
        appearance_grid.addWidget(self.combo_lyric_placement, 18, 1, Qt.AlignmentFlag.AlignLeft)

        self.chk_lyric_start_blacked = QCheckBox("開啟詩歌時先黑畫面（只外屏）")
        self.chk_lyric_start_blacked.setChecked(self._settings.get('lyric_start_blacked', True))
        self.chk_lyric_start_blacked.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_lyric_start_blacked.setToolTip(
            "主展示先全黑，預覽仍可見歌詞；調好字距／對齊後再按「取消全黑」"
        )
        self.chk_lyric_start_blacked.toggled.connect(self._on_lyric_start_blacked_toggled)
        appearance_grid.addWidget(self.chk_lyric_start_blacked, 19, 0, 1, 2)

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
        overlay_dur = _normalize_overlay_duration(self._settings['overlay_duration'])
        self.spin_overlay_duration = QSpinBox()
        self.spin_overlay_duration.setRange(1, MAX_OVERLAY_SECS)
        self.spin_overlay_duration.setSuffix(' 秒')
        self.spin_overlay_duration.setMinimumWidth(108)
        self.spin_overlay_duration.setValue(
            DEFAULT_OVERLAY_SECS if overlay_dur == OVERLAY_DURATION_ALWAYS else overlay_dur
        )
        self.spin_overlay_duration.setToolTip('提示顯示幾耐（秒），可用加減自行調校')
        self.spin_overlay_duration.valueChanged.connect(self._on_overlay_duration_changed)
        self.chk_overlay_always = QCheckBox('一直顯示')
        self.chk_overlay_always.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_overlay_always.setChecked(overlay_dur == OVERLAY_DURATION_ALWAYS)
        self.chk_overlay_always.toggled.connect(self._on_overlay_always_toggled)
        self.spin_overlay_duration.setEnabled(overlay_dur != OVERLAY_DURATION_ALWAYS)
        self.overlay_duration_row = QWidget()
        overlay_dur_lay = QHBoxLayout(self.overlay_duration_row)
        overlay_dur_lay.setContentsMargins(0, 0, 0, 0)
        overlay_dur_lay.setSpacing(8)
        overlay_dur_lay.addWidget(self.spin_overlay_duration)
        overlay_dur_lay.addWidget(self.chk_overlay_always)
        overlay_dur_lay.addStretch()

        self.lbl_overlay_center = QLabel("轉移時間")
        self.lbl_overlay_center.setMinimumWidth(label_min_w)
        self.lbl_overlay_center.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.spin_overlay_center = QSpinBox()
        self.spin_overlay_center.setRange(1, MAX_OVERLAY_CENTER_SECS)
        self.spin_overlay_center.setSuffix(' 秒')
        self.spin_overlay_center.setMinimumWidth(108)
        self.spin_overlay_center.setValue(
            _normalize_center_secs(self._settings.get('overlay_center_secs', CENTER_OVERLAY_SECS))
        )
        self.spin_overlay_center.setCursor(Qt.CursorShape.PointingHandCursor)
        self.spin_overlay_center.setToolTip("置中提示停留幾耐先移去右下角")
        self.spin_overlay_center.valueChanged.connect(self._on_overlay_center_changed)

        self.chk_overlay_auto_unblack = QCheckBox("倒數完自動取消黑畫面")
        self.chk_overlay_auto_unblack.setChecked(self._settings.get('overlay_auto_unblack', False))
        self.chk_overlay_auto_unblack.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_overlay_auto_unblack.setToolTip(
            "置中提示轉移去右下角之後自動取消外屏全黑。唔剔則維持全黑，要手動取消。"
        )
        self.chk_overlay_auto_unblack.toggled.connect(self._on_overlay_auto_unblack_toggled)

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

        self.lbl_overlay_bg = QLabel("提示背景")
        self.lbl_overlay_bg.setMinimumWidth(label_min_w)
        self.lbl_overlay_bg.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        (
            self.overlay_bg_row,
            self.inp_overlay_bg_color,
            self.lbl_overlay_bg_preview,
            self.btn_pick_overlay_bg_color,
        ) = self._make_qr_color_row(
            self._settings['overlay_bg_color'],
            DEFAULT_OVERLAY_BG_COLOR,
            "書名提示背景色",
            allow_auto=False,
        )
        self.inp_overlay_bg_color.textChanged.connect(self._on_overlay_appearance_changed)
        self.btn_pick_overlay_bg_color.clicked.connect(
            lambda: self._pick_qr_color(
                self.inp_overlay_bg_color, self.lbl_overlay_bg_preview, self._on_overlay_appearance_changed,
            )
        )

        self.lbl_overlay_text_color = QLabel("提示字色")
        self.lbl_overlay_text_color.setMinimumWidth(label_min_w)
        self.lbl_overlay_text_color.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        (
            self.overlay_text_color_row,
            self.inp_overlay_text_color,
            self.lbl_overlay_text_color_preview,
            self.btn_pick_overlay_text_color,
        ) = self._make_qr_color_row(
            self._settings.get('overlay_text_color') or '',
            '#ffffff',
            "書名提示文字色（留空 = 依背景自動）",
            allow_auto=True,
        )
        self.inp_overlay_text_color.textChanged.connect(self._on_overlay_appearance_changed)
        self.btn_pick_overlay_text_color.clicked.connect(
            lambda: self._pick_qr_color(
                self.inp_overlay_text_color,
                self.lbl_overlay_text_color_preview, self._on_overlay_appearance_changed,
            )
        )
        self.btn_overlay_text_color_auto = QPushButton("自動")
        self.btn_overlay_text_color_auto.setFixedWidth(52)
        self.btn_overlay_text_color_auto.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_overlay_text_color_auto.setToolTip("依背景自動選擇文字顏色")
        self.btn_overlay_text_color_auto.clicked.connect(self._reset_overlay_text_color_auto)
        self.overlay_text_color_row.layout().addWidget(self.btn_overlay_text_color_auto)

        self.chk_display_duplicate = QCheckBox("切換同步畫面")
        self.chk_display_duplicate.setToolTip(
            "開檔進入全螢幕後自動切換為同步（複製）畫面；"
            "只儲存設定，勾選時唔會即時切換"
        )
        self.chk_display_duplicate.setChecked(self._settings['display_duplicate'])
        self.chk_display_duplicate.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_display_duplicate.toggled.connect(self._on_display_duplicate_toggled)
        if sys.platform != 'win32':
            self.chk_display_duplicate.setEnabled(False)
            self.chk_display_duplicate.setToolTip("僅支援 Windows")

        self.lbl_display_switch_mode = QLabel("切換方式")
        self.lbl_display_switch_mode.setMinimumWidth(label_min_w)
        self.lbl_display_switch_mode.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_display_switch_mode = QComboBox()
        self.combo_display_switch_mode.addItem("快速（SetDisplayConfig）", DISPLAY_SWITCH_CCD)
        self.combo_display_switch_mode.addItem("舊模式（DisplaySwitch）", DISPLAY_SWITCH_LEGACY)
        self.combo_display_switch_mode.setToolTip(
            "快速：直接 API，通常無 Win+P 側欄；"
            "舊模式：DisplaySwitch.exe（會彈出切換提示）"
        )
        self.combo_display_switch_mode.setCursor(Qt.CursorShape.PointingHandCursor)
        self.combo_display_switch_mode.currentIndexChanged.connect(self._on_display_switch_mode_changed)
        if sys.platform != 'win32':
            self.lbl_display_switch_mode.setEnabled(False)
            self.combo_display_switch_mode.setEnabled(False)

        self.chk_word_auto_fullscreen = QCheckBox("Word 自動全螢幕")
        self.chk_word_auto_fullscreen.setToolTip(
            "開啟 Word 後自動送 Enter / Alt+W / Alt+O 進入投影全螢幕"
        )
        self.chk_word_auto_fullscreen.setChecked(self._settings['word_auto_fullscreen'])
        self.chk_word_auto_fullscreen.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_word_auto_fullscreen.toggled.connect(self._on_word_auto_fullscreen_toggled)
        if sys.platform != 'win32':
            self.chk_word_auto_fullscreen.setEnabled(False)

        self.chk_pdf_auto_fullscreen = QCheckBox("PDF 自動全螢幕")
        self.chk_pdf_auto_fullscreen.setToolTip(
            "開啟 PDF 後自動全螢幕（Sumatra 用 -presentation；其他用 F11）；"
            "Foxit 等已自動全螢幕嘅閱讀器可關閉"
        )
        self.chk_pdf_auto_fullscreen.setChecked(self._settings['pdf_auto_fullscreen'])
        self.chk_pdf_auto_fullscreen.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_pdf_auto_fullscreen.toggled.connect(self._on_pdf_auto_fullscreen_toggled)
        if sys.platform != 'win32':
            self.chk_pdf_auto_fullscreen.setEnabled(False)

        open_grid.addWidget(self.chk_open_overlay, 0, 0, 1, 2)
        open_grid.addWidget(self.lbl_overlay_secs, 1, 0)
        open_grid.addWidget(self.overlay_duration_row, 1, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.lbl_overlay_pos, 2, 0)
        open_grid.addWidget(self.combo_overlay_pos, 2, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.lbl_overlay_center, 3, 0)
        open_grid.addWidget(self.spin_overlay_center, 3, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.chk_overlay_auto_unblack, 4, 0, 1, 2)
        open_grid.addWidget(self.lbl_overlay_bg, 5, 0)
        open_grid.addWidget(self.overlay_bg_row, 5, 1)
        open_grid.addWidget(self.lbl_overlay_text_color, 6, 0)
        open_grid.addWidget(self.overlay_text_color_row, 6, 1)
        open_grid.addWidget(self.chk_display_duplicate, 7, 0, 1, 2)
        open_grid.addWidget(self.lbl_display_switch_mode, 8, 0)
        open_grid.addWidget(self.combo_display_switch_mode, 8, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.chk_word_auto_fullscreen, 9, 0, 1, 2)
        open_grid.addWidget(self.chk_pdf_auto_fullscreen, 10, 0, 1, 2)

        self.grp_remote = QGroupBox("Mobile 遙控")
        remote_outer = QVBoxLayout(self.grp_remote)
        remote_outer.setContentsMargins(*grid_margins)
        remote_outer.setSpacing(12)

        self.chk_remote_api = QCheckBox("啟用 API")
        self.chk_remote_api.setChecked(self._settings['remote_api_enabled'])
        self.chk_remote_api.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_remote_api.toggled.connect(self._on_remote_api_toggled)

        self.chk_remote_accept = QCheckBox("接受請求")
        self.chk_remote_accept.setChecked(self._settings['remote_accept'])
        self.chk_remote_accept.setEnabled(False)
        self.chk_remote_accept.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_remote_accept.toggled.connect(self._on_remote_accept_settings_toggled)

        self.lbl_remote_policy = QLabel("開檔政策")
        self.lbl_remote_policy.setMinimumWidth(label_min_w)
        self.lbl_remote_policy.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_remote_policy = QComboBox()
        self.combo_remote_policy.addItem("唯一結果自動開", OPEN_POLICY_AUTO)
        self.combo_remote_policy.addItem("待操作員確認", OPEN_POLICY_CONFIRM)
        self.combo_remote_policy.addItem("只填 UI 不自動開", OPEN_POLICY_UI)
        self.combo_remote_policy.setCurrentIndex(
            _remote_policy_index(self._settings['remote_policy'])
        )
        self.combo_remote_policy.setMinimumWidth(180)
        self.combo_remote_policy.setEnabled(False)
        self.combo_remote_policy.setCursor(Qt.CursorShape.PointingHandCursor)
        self.combo_remote_policy.setToolTip("手機遙控開檔時的處理方式（須先啟用 API）")
        self.combo_remote_policy.currentIndexChanged.connect(self._on_remote_policy_changed)

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

        self.chk_enable_tunnel_fixed = QCheckBox("固定網址（共用 Windows 服務）")
        self.chk_enable_tunnel_fixed.setChecked(self.enable_tunnel_fixed)
        self.chk_enable_tunnel_fixed.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_enable_tunnel_fixed.setToolTip(
            "使用已安裝的 cloudflared 服務與固定公開 HTTPS 網址；可與隨機通道同時啟用。"
        )
        self.chk_enable_tunnel_fixed.toggled.connect(self._on_tunnel_fixed_toggled)

        self.chk_enable_tunnel_quick = QCheckBox("隨機網址（快速通道／測試）")
        self.chk_enable_tunnel_quick.setChecked(self.enable_tunnel_quick)
        self.chk_enable_tunnel_quick.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_enable_tunnel_quick.setToolTip(
            "臨時 trycloudflare.com 網址；可與固定網址同時啟用（兩組外網 QR）。"
        )
        self.chk_enable_tunnel_quick.toggled.connect(self._on_tunnel_quick_toggled)

        # Compat aliases used by theme / restore code paths.
        self.chk_enable_tunnel = self.chk_enable_tunnel_fixed
        self.tunnel_mode_combo = None

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

        self.tunnel_status_fixed_lbl = QLabel("固定：—")
        self.tunnel_status_fixed_lbl.setWordWrap(True)
        self.tunnel_status_quick_lbl = QLabel("隨機：—")
        self.tunnel_status_quick_lbl.setWordWrap(True)
        self.tunnel_status_lbl = self.tunnel_status_fixed_lbl

        self.lbl_local_url = QLabel("內網會眾")
        self.lbl_local_url.setMinimumWidth(label_min_w)
        self.local_url_lbl = QLabel("—")
        self.local_url_lbl.setWordWrap(True)
        self.local_url_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.local_url_lbl.setToolTip("同 WiFi 會眾觀看頁")

        self.lbl_local_admin_url = QLabel("內網遙控")
        self.lbl_local_admin_url.setMinimumWidth(label_min_w)
        self.local_admin_url_lbl = QLabel("—")
        self.local_admin_url_lbl.setWordWrap(True)
        self.local_admin_url_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.local_admin_url_lbl.setToolTip("同 WiFi 操作員遙控頁（/admin）")

        self.lbl_remote_url = QLabel("固定遙控")
        self.lbl_remote_url.setMinimumWidth(label_min_w)
        self.remote_url_lbl = QLabel("—")
        self.remote_url_lbl.setWordWrap(True)
        self.remote_url_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.remote_url_lbl.setToolTip("操作員遙控頁（/admin）；固定外網")

        self.lbl_viewer_url = QLabel("固定會眾")
        self.lbl_viewer_url.setMinimumWidth(label_min_w)
        self.viewer_url_lbl = QLabel("—")
        self.viewer_url_lbl.setWordWrap(True)
        self.viewer_url_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.viewer_url_lbl.setToolTip("固定公開 HTTPS 會眾頁")

        self.lbl_quick_url = QLabel("隨機會眾")
        self.lbl_quick_url.setMinimumWidth(label_min_w)
        self.quick_url_lbl = QLabel("—")
        self.quick_url_lbl.setWordWrap(True)
        self.quick_url_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.quick_url_lbl.setToolTip("快速通道臨時 HTTPS 會眾頁")

        self.chk_viewer_show_debug = QCheckBox("會眾觀看顯示除錯列")
        self.chk_viewer_show_debug.setChecked(self._settings['viewer_show_debug'])
        self.chk_viewer_show_debug.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_viewer_show_debug.setToolTip(
            "開啟後，手機會眾頁底部會顯示歌詞同步除錯資訊"
        )
        self.chk_viewer_show_debug.toggled.connect(self._on_viewer_show_debug_toggled)

        self.chk_show_qr_lan = QCheckBox("顯示內網 QR（含連結）")
        self.chk_show_qr_lan.setChecked(self._settings.get('show_qr_lan', True))
        self.chk_show_qr_lan.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_show_qr_lan.toggled.connect(self._on_show_qr_role_toggled)

        self.chk_show_qr_fixed = QCheckBox("顯示固定外網 QR（含連結）")
        self.chk_show_qr_fixed.setChecked(self._settings.get('show_qr_fixed', True))
        self.chk_show_qr_fixed.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_show_qr_fixed.toggled.connect(self._on_show_qr_role_toggled)

        self.chk_show_qr_quick = QCheckBox("顯示隨機外網 QR（含連結）")
        self.chk_show_qr_quick.setChecked(self._settings.get('show_qr_quick', True))
        self.chk_show_qr_quick.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_show_qr_quick.toggled.connect(self._on_show_qr_role_toggled)
        self.chk_show_qr_wan = self.chk_show_qr_fixed

        (
            self.remote_qr_lan_pair,
            self.remote_qr_lan_lbl, self.remote_qr_lan_link,
            self.remote_qr_lan_admin_lbl, self.remote_qr_lan_admin_link,
        ) = self._make_settings_qr_pair(
            'remoteQrLanLabel', 'remoteQrLanAdminLabel',
            '啟用 API 後顯示', '啟用 API 後顯示',
        )
        (
            self.remote_qr_fixed_pair,
            self.remote_qr_fixed_lbl, self.remote_qr_fixed_link,
            self.remote_qr_fixed_admin_lbl, self.remote_qr_fixed_admin_link,
        ) = self._make_settings_qr_pair(
            'remoteQrFixedLabel', 'remoteQrFixedAdminLabel',
            '啟用固定外網後顯示', '啟用固定外網後顯示',
        )
        (
            self.remote_qr_quick_pair,
            self.remote_qr_quick_lbl, self.remote_qr_quick_link,
            self.remote_qr_quick_admin_lbl, self.remote_qr_quick_admin_link,
        ) = self._make_settings_qr_pair(
            'remoteQrQuickLabel', 'remoteQrQuickAdminLabel',
            '啟用隨機外網後顯示', '啟用隨機外網後顯示',
        )

        self.remote_qr_wan_lbl = self.remote_qr_fixed_lbl
        self.remote_qr_wan_link = self.remote_qr_fixed_link
        self.remote_qr_lbl = self.remote_qr_lan_lbl

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
        self.spin_qr_h.setToolTip("桌面 QR 卡片高度（含下方連結文字；過矮會自動加高）")
        self.spin_qr_h.valueChanged.connect(self._on_qr_layout_changed)

        self.lbl_qr_caption = QLabel("QR 標題")
        self.lbl_qr_caption.setMinimumWidth(label_min_w)
        self.lbl_qr_caption.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.inp_qr_caption = QLineEdit()
        self.inp_qr_caption.setText(self._settings['qr_caption'])
        self.inp_qr_caption.setPlaceholderText(DEFAULT_QR_CAPTION)
        self.inp_qr_caption.setToolTip("顯示於 QR 上方；下方會自動附「內網／外網」與完整連結")
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

        # ── 分區：本機 API ──
        grp_api = QGroupBox("本機 API")
        api_grid = QGridLayout(grp_api)
        api_grid.setHorizontalSpacing(14)
        api_grid.setVerticalSpacing(10)
        api_grid.setColumnStretch(1, 1)
        api_grid.addWidget(self.chk_remote_api, 0, 0, 1, 2)
        api_grid.addWidget(self.chk_remote_accept, 1, 0, 1, 2)
        api_grid.addWidget(self.lbl_remote_policy, 2, 0)
        api_grid.addWidget(self.combo_remote_policy, 2, 1, Qt.AlignmentFlag.AlignLeft)
        api_grid.addWidget(self.lbl_remote_port, 3, 0)
        api_grid.addWidget(self.spin_remote_port, 3, 1, Qt.AlignmentFlag.AlignLeft)
        api_grid.addWidget(self.lbl_remote_token, 4, 0)
        api_grid.addWidget(self.inp_remote_token, 4, 1)
        api_grid.addWidget(self.chk_viewer_show_debug, 5, 0, 1, 2)

        # ── 分區：內網 ──
        grp_lan = QGroupBox("內網（同 WiFi）")
        lan_grid = QGridLayout(grp_lan)
        lan_grid.setHorizontalSpacing(14)
        lan_grid.setVerticalSpacing(8)
        lan_grid.setColumnStretch(1, 1)
        lan_grid.addWidget(self.lbl_local_url, 0, 0)
        lan_grid.addWidget(self.local_url_lbl, 0, 1)
        lan_grid.addWidget(self.lbl_local_admin_url, 1, 0)
        lan_grid.addWidget(self.local_admin_url_lbl, 1, 1)
        lan_grid.addWidget(self.chk_show_qr_lan, 2, 0, 1, 2)
        lan_grid.addWidget(self.remote_qr_lan_pair, 3, 0, 1, 2)

        # ── 分區：外網固定 ──
        grp_wan_fixed = QGroupBox("外網 · 固定網址")
        wan_fixed_lay = QVBoxLayout(grp_wan_fixed)
        wan_fixed_lay.setSpacing(8)
        wan_fixed_lay.addWidget(self.chk_enable_tunnel_fixed)
        wan_service = QWidget()
        self._wan_service_box = wan_service
        wan_svc = QGridLayout(wan_service)
        wan_svc.setContentsMargins(0, 0, 0, 0)
        wan_svc.setHorizontalSpacing(14)
        wan_svc.setVerticalSpacing(8)
        wan_svc.setColumnStretch(1, 1)
        wan_svc.addWidget(self.lbl_public_tunnel_url, 0, 0)
        wan_svc.addWidget(self.inp_named_tunnel_url, 0, 1)
        wan_svc.addWidget(self.inp_cloudflared_token, 1, 0, 1, 2)
        wan_svc.addWidget(self.btn_install_cloudflared, 2, 0, 1, 2)
        wan_svc.addWidget(self.lbl_tunnel_cmd_hint, 3, 0, 1, 2)
        wan_svc.addWidget(self._tunnel_install_row, 4, 0, 1, 2)
        wan_svc.addWidget(self._tunnel_uninstall_row, 5, 0, 1, 2)
        wan_svc.addWidget(self._tunnel_restart_row, 6, 0, 1, 2)
        wan_fixed_lay.addWidget(wan_service)
        wan_fixed_lay.addWidget(self.tunnel_help_lbl)
        wan_fixed_lay.addWidget(self.tunnel_status_fixed_lbl)
        wan_urls = QGridLayout()
        wan_urls.setHorizontalSpacing(14)
        wan_urls.setVerticalSpacing(6)
        wan_urls.setColumnStretch(1, 1)
        wan_urls.addWidget(self.lbl_viewer_url, 0, 0)
        wan_urls.addWidget(self.viewer_url_lbl, 0, 1)
        wan_urls.addWidget(self.lbl_remote_url, 1, 0)
        wan_urls.addWidget(self.remote_url_lbl, 1, 1)
        wan_fixed_lay.addLayout(wan_urls)
        wan_fixed_lay.addWidget(self.chk_show_qr_fixed)
        wan_fixed_lay.addWidget(self.remote_qr_fixed_pair)

        # ── 分區：外網隨機 ──
        grp_wan_quick = QGroupBox("外網 · 隨機網址")
        wan_quick_lay = QVBoxLayout(grp_wan_quick)
        wan_quick_lay.setSpacing(8)
        wan_quick_lay.addWidget(self.chk_enable_tunnel_quick)
        wan_quick_lay.addWidget(self.tunnel_status_quick_lbl)
        quick_urls = QGridLayout()
        quick_urls.setHorizontalSpacing(14)
        quick_urls.setVerticalSpacing(6)
        quick_urls.setColumnStretch(1, 1)
        quick_urls.addWidget(self.lbl_quick_url, 0, 0)
        quick_urls.addWidget(self.quick_url_lbl, 0, 1)
        self.lbl_quick_admin_url = QLabel("隨機遙控")
        self.lbl_quick_admin_url.setMinimumWidth(label_min_w)
        self.quick_admin_url_lbl = QLabel("—")
        self.quick_admin_url_lbl.setWordWrap(True)
        self.quick_admin_url_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.quick_admin_url_lbl.setToolTip("快速通道臨時 HTTPS 操作員遙控頁（/admin）")
        quick_urls.addWidget(self.lbl_quick_admin_url, 1, 0)
        quick_urls.addWidget(self.quick_admin_url_lbl, 1, 1)
        wan_quick_lay.addLayout(quick_urls)
        wan_quick_lay.addWidget(self.chk_show_qr_quick)
        wan_quick_lay.addWidget(self.remote_qr_quick_pair)

        # ── 分區：桌面 QR 樣式（內網／外網共用）──
        grp_qr = QGroupBox("桌面 QR 樣式（內網／固定／隨機共用）")
        qr_grid = QGridLayout(grp_qr)
        qr_grid.setHorizontalSpacing(14)
        qr_grid.setVerticalSpacing(10)
        qr_grid.setColumnStretch(1, 1)
        qr_grid.addWidget(self.lbl_qr_x, 0, 0)
        qr_grid.addWidget(self.spin_qr_x, 0, 1, Qt.AlignmentFlag.AlignLeft)
        qr_grid.addWidget(self.lbl_qr_y, 1, 0)
        qr_grid.addWidget(self.spin_qr_y, 1, 1, Qt.AlignmentFlag.AlignLeft)
        qr_grid.addWidget(self.lbl_qr_w, 2, 0)
        qr_grid.addWidget(self.spin_qr_w, 2, 1, Qt.AlignmentFlag.AlignLeft)
        qr_grid.addWidget(self.lbl_qr_h, 3, 0)
        qr_grid.addWidget(self.spin_qr_h, 3, 1, Qt.AlignmentFlag.AlignLeft)
        qr_grid.addWidget(self.lbl_qr_caption, 4, 0)
        qr_grid.addWidget(self.inp_qr_caption, 4, 1)
        qr_grid.addWidget(self.lbl_qr_bg, 5, 0)
        qr_grid.addWidget(self.qr_bg_row, 5, 1)
        qr_grid.addWidget(self.lbl_qr_text_color, 6, 0)
        qr_grid.addWidget(self.qr_text_color_row, 6, 1)
        qr_grid.addWidget(self.lbl_qr_text_size, 7, 0)
        qr_grid.addWidget(self.spin_qr_text_size, 7, 1, Qt.AlignmentFlag.AlignLeft)
        qr_grid.addWidget(self.lbl_qr_bg_opacity, 8, 0)
        qr_grid.addWidget(self.spin_qr_bg_opacity, 8, 1, Qt.AlignmentFlag.AlignLeft)

        remote_outer.addWidget(grp_api)
        remote_outer.addWidget(grp_lan)
        remote_outer.addWidget(grp_wan_fixed)
        remote_outer.addWidget(grp_wan_quick)
        remote_outer.addWidget(grp_qr)
        self._update_qr_color_previews()
        self._sync_tunnel_options_visibility()

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
        tab_mobile_lay.addWidget(self.remote_log_list)
        tab_mobile_lay.addStretch()

        tab_about = QWidget()
        tab_about_lay = QVBoxLayout(tab_about)
        tab_about_lay.setContentsMargins(0, 8, 0, 0)
        tab_about_lay.setSpacing(12)
        self.about_ver_lbl = QLabel(
            f"詩歌冊搜索系統 v{APP_VERSION}" + (f"  ·  {BUILD_DATE}" if BUILD_DATE else "")
        )
        self.about_ver_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft)
        tab_about_lay.addWidget(self.about_ver_lbl)
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
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self._file_list_context_menu)
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
            self._overlay_duration_value(),
            self.combo_overlay_pos.currentData(),
            self._settings,
            center_secs=self._overlay_center_secs(),
        )
        self._schedule_lyric_preview()
        self._start_overlay_countdown_tick()

    def _show_open_hymn_popup(self, book_name, item_name=''):
        if not self.chk_open_overlay.isChecked():
            return
        if not book_name:
            return
        BookNameOverlay.show_book(
            book_name, 'lyric',
            item_name or '',
            self._overlay_duration_value(),
            self.combo_overlay_pos.currentData(),
            self._settings,
            center_secs=self._overlay_center_secs(),
        )
        self._schedule_lyric_preview()
        self._start_overlay_countdown_tick()

    def _overlay_appearance_widgets(self):
        return (
            self.lbl_overlay_bg, self.overlay_bg_row,
            self.lbl_overlay_text_color, self.overlay_text_color_row,
        )

    def _overlay_appearance_values(self):
        text_color = (
            self.inp_overlay_text_color.text().strip()
            if hasattr(self, 'inp_overlay_text_color') else ''
        )
        return {
            'overlay_bg_color': _normalize_hex_color(
                self.inp_overlay_bg_color.text(), DEFAULT_OVERLAY_BG_COLOR,
            ),
            'overlay_text_color': (
                _normalize_hex_color(text_color, DEFAULT_OVERLAY_TEXT_COLOR)
                if text_color else DEFAULT_OVERLAY_TEXT_COLOR
            ),
        }

    def _on_lyric_appearance_changed(self, _value=''):
        if getattr(self, '_loading_settings', False):
            return
        bg = _normalize_hex_color(
            self.inp_lyric_bg_color.text() if hasattr(self, 'inp_lyric_bg_color') else '',
            DEFAULT_LYRIC_BG,
        )
        text = _normalize_hex_color(
            self.inp_lyric_text_color.text() if hasattr(self, 'inp_lyric_text_color') else '',
            DEFAULT_LYRIC_TEXT,
        )
        self._update_qr_color_preview(self.lbl_lyric_bg_preview, bg, DEFAULT_LYRIC_BG)
        self._update_qr_color_preview(self.lbl_lyric_text_preview, text, DEFAULT_LYRIC_TEXT)
        timer = getattr(self, '_lyric_appearance_debounce', None)
        if timer is not None:
            timer.start()
        else:
            self._apply_lyric_appearance_from_settings()

    def _apply_lyric_appearance_from_settings(self):
        if getattr(self, '_loading_settings', False):
            return
        bg = _normalize_hex_color(
            self.inp_lyric_bg_color.text() if hasattr(self, 'inp_lyric_bg_color') else '',
            DEFAULT_LYRIC_BG,
        )
        text = _normalize_hex_color(
            self.inp_lyric_text_color.text() if hasattr(self, 'inp_lyric_text_color') else '',
            DEFAULT_LYRIC_TEXT,
        )
        self._persist_settings(lyric_bg_color=bg, lyric_text_color=text)
        pair = getattr(self, '_lyric_pair', None)
        if pair is None:
            return
        self._lyric_ui_applying = True
        try:
            pair.session.set_colors(bg, text)
        finally:
            self._lyric_ui_applying = False

    def _stored_lyric_bg_paths(self):
        folder = app_dir()
        paths = []
        try:
            names = os.listdir(folder)
        except OSError:
            return paths
        for name in names:
            stem, ext = os.path.splitext(name)
            if stem == LYRIC_BG_IMAGE_STEM and ext.lower() in LYRIC_BG_IMAGE_EXTS:
                paths.append(os.path.join(folder, name))
        return paths

    def _remove_stored_lyric_bg_images(self):
        for path in self._stored_lyric_bg_paths():
            try:
                os.remove(path)
            except OSError:
                pass

    def _sync_lyric_bg_image_label(self):
        path = normalize_lyric_bg_image(
            self._settings.get('lyric_bg_image') if hasattr(self, '_settings') else '',
            app_dir(),
        )
        if hasattr(self, 'lbl_lyric_bg_image'):
            self.lbl_lyric_bg_image.setText(os.path.basename(path) if path else '純色')
        if hasattr(self, 'btn_lyric_bg_image_clear'):
            self.btn_lyric_bg_image_clear.setEnabled(bool(path))

    def _on_lyric_bg_image_upload(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '選擇背景圖片', '',
            '圖片 (*.png *.jpg *.jpeg *.webp *.bmp *.gif)',
        )
        if not path:
            return
        if QPixmap(path).isNull():
            QMessageBox.warning(self, '背景圖片', '無法讀取這張圖片。')
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in LYRIC_BG_IMAGE_EXTS:
            ext = '.png'
        dest = os.path.join(app_dir(), f'{LYRIC_BG_IMAGE_STEM}{ext}')
        self._remove_stored_lyric_bg_images()
        try:
            shutil.copy2(path, dest)
        except OSError as exc:
            QMessageBox.warning(self, '背景圖片', f'無法儲存圖片：{exc}')
            return
        self._persist_settings(lyric_bg_image=dest)
        if getattr(self, '_lyric_pair', None) is not None:
            self._lyric_pair.session.set_bg_image(dest)
        self._sync_lyric_bg_image_label()

    def _on_lyric_bg_image_clear(self):
        self._remove_stored_lyric_bg_images()
        self._persist_settings(lyric_bg_image='')
        if getattr(self, '_lyric_pair', None) is not None:
            self._lyric_pair.session.set_bg_image('')
        self._sync_lyric_bg_image_label()

    def _on_lyric_font_auto_toggled(self, checked):
        if hasattr(self, 'spin_lyric_font'):
            self.spin_lyric_font.setEnabled(not bool(checked))
        self._on_lyric_layout_settings_changed()

    def _lyric_font_family_from_ui(self):
        if hasattr(self, 'combo_lyric_font_family'):
            return normalize_lyric_font_family(self.combo_lyric_font_family.currentFont().family())
        return normalize_lyric_font_family(self._settings.get('lyric_font_family'))

    def _lyric_font_pt_from_ui(self):
        if hasattr(self, 'chk_lyric_font_auto') and self.chk_lyric_font_auto.isChecked():
            return 0
        if hasattr(self, 'spin_lyric_font'):
            return normalize_lyric_font_pt(self.spin_lyric_font.value())
        return normalize_lyric_font_pt(self._settings.get('lyric_font_pt'))

    def _lyric_margins_from_ui(self):
        if hasattr(self, 'spin_lyric_margins') and self.spin_lyric_margins:
            return normalize_lyric_margins({
                side: self.spin_lyric_margins[side].value()
                for side in MARGIN_SIDES
                if side in self.spin_lyric_margins
            })
        return normalize_lyric_margins(self._settings.get('lyric_margins'))

    def _lyric_layout_from_ui(self):
        align = DEFAULT_LYRIC_ALIGN
        if hasattr(self, 'combo_lyric_align'):
            align = normalize_lyric_align(self.combo_lyric_align.currentData())
        spacing = DEFAULT_LETTER_SPACING
        if hasattr(self, 'slider_lyric_spacing'):
            spacing = normalize_letter_spacing(self.slider_lyric_spacing.value())
        line_height = DEFAULT_LINE_HEIGHT
        if hasattr(self, 'slider_lyric_line_height'):
            line_height = normalize_line_height(self.slider_lyric_line_height.value())
        return lyric_best_from_values(
            align=align,
            letter_spacing=spacing,
            line_height=line_height,
            margins=self._lyric_margins_from_ui(),
            font_family=self._lyric_font_family_from_ui(),
            font_pt=self._lyric_font_pt_from_ui(),
        )

    def _sync_lyric_layout_widgets(self, layout):
        layout = normalize_lyric_best_defaults(layout)
        was_loading = getattr(self, '_loading_settings', False)
        self._loading_settings = True
        try:
            if hasattr(self, 'combo_lyric_align'):
                idx = self.combo_lyric_align.findData(layout['align'])
                self.combo_lyric_align.setCurrentIndex(idx if idx >= 0 else 0)
            if hasattr(self, 'slider_lyric_spacing'):
                self.slider_lyric_spacing.setValue(layout['letter_spacing'])
                if hasattr(self, 'lbl_lyric_spacing_val'):
                    self.lbl_lyric_spacing_val.setText(f"{layout['letter_spacing']}%")
            if hasattr(self, 'slider_lyric_line_height'):
                self.slider_lyric_line_height.setValue(layout['line_height'])
                if hasattr(self, 'lbl_lyric_line_height_val'):
                    self.lbl_lyric_line_height_val.setText(f"{layout['line_height']}%")
            if hasattr(self, 'spin_lyric_margins'):
                margins = normalize_lyric_margins(layout.get('margins'))
                for side, spin in self.spin_lyric_margins.items():
                    spin.setValue(margins[side])
            if hasattr(self, 'combo_lyric_font_family'):
                self.combo_lyric_font_family.setCurrentFont(QFont(layout['font_family']))
            if hasattr(self, 'chk_lyric_font_auto') and hasattr(self, 'spin_lyric_font'):
                font_pt = layout['font_pt']
                self.chk_lyric_font_auto.setChecked(font_pt <= 0)
                self.spin_lyric_font.setValue(font_pt if font_pt > 0 else 48)
                self.spin_lyric_font.setEnabled(font_pt > 0)
        finally:
            self._loading_settings = was_loading

    def _apply_lyric_best_to_session(self, layout, *, refit=True):
        layout = normalize_lyric_best_defaults(layout)
        pair = getattr(self, '_lyric_pair', None)
        if pair is None:
            return layout
        session = pair.session
        self._lyric_ui_applying = True
        try:
            session.apply_layout(
                align=layout['align'],
                letter_spacing=layout['letter_spacing'],
                line_height=layout['line_height'],
                margins=layout['margins'],
                font_family=layout['font_family'],
                preferred_font_pt=layout['font_pt'],
            )
            session.user_zoomed = False
            if refit and (session.hymn or pair.is_open()):
                pair.fit_display()
        finally:
            self._lyric_ui_applying = False
        return layout

    def _persist_lyric_layout(self, layout, *, as_best=False):
        layout = normalize_lyric_best_defaults(layout)
        payload = {
            'lyric_align': layout['align'],
            'lyric_letter_spacing': layout['letter_spacing'],
            'lyric_line_height': layout['line_height'],
            'lyric_margins': layout['margins'],
            'lyric_font_family': layout['font_family'],
            'lyric_font_pt': layout['font_pt'],
        }
        if as_best:
            payload['lyric_best_defaults'] = layout
        self._persist_settings(**payload)
        return layout

    def _save_lyric_best_defaults(self):
        pair = getattr(self, '_lyric_pair', None)
        if pair is not None and (pair.session.hymn or pair.is_open()):
            layout = lyric_best_from_session(pair.session)
        else:
            layout = self._lyric_layout_from_ui()
        layout = self._persist_lyric_layout(layout, as_best=True)
        self._apply_lyric_best_to_session(layout, refit=True)
        self._sync_lyric_layout_widgets(layout)
        self._schedule_lyric_preview()
        if hasattr(self, 'statusBar'):
            try:
                self.statusBar().showMessage('已保存為最佳預設', 2500)
            except Exception:
                pass

    def _restore_lyric_best_defaults(self):
        layout = normalize_lyric_best_defaults(self._settings.get('lyric_best_defaults'))
        layout = self._persist_lyric_layout(layout, as_best=False)
        self._apply_lyric_best_to_session(layout, refit=True)
        self._sync_lyric_layout_widgets(layout)
        self._schedule_lyric_preview()
        if hasattr(self, 'statusBar'):
            try:
                self.statusBar().showMessage('已還原最佳預設', 2500)
            except Exception:
                pass

    def _on_lyric_layout_settings_changed(self, _value=None):
        if getattr(self, '_loading_settings', False) or getattr(self, '_lyric_ui_applying', False):
            return
        layout = self._lyric_layout_from_ui()
        if hasattr(self, 'lbl_lyric_spacing_val'):
            self.lbl_lyric_spacing_val.setText(f"{layout['letter_spacing']}%")
        if hasattr(self, 'lbl_lyric_line_height_val'):
            self.lbl_lyric_line_height_val.setText(f"{layout['line_height']}%")
        timer = getattr(self, '_lyric_layout_debounce', None)
        if timer is not None:
            timer.start()
        else:
            self._apply_lyric_layout_from_settings()

    def _apply_lyric_layout_from_settings(self):
        if getattr(self, '_loading_settings', False):
            return
        layout = self._lyric_layout_from_ui()
        self._persist_lyric_layout(layout, as_best=False)
        pair = getattr(self, '_lyric_pair', None)
        if pair is None:
            return
        self._lyric_ui_applying = True
        try:
            pair.session.apply_layout(
                align=layout['align'],
                letter_spacing=layout['letter_spacing'],
                line_height=layout['line_height'],
                margins=layout['margins'],
                font_family=layout['font_family'],
                preferred_font_pt=layout['font_pt'],
            )
            if not pair.session.user_zoomed:
                pair.fit_display()
        finally:
            self._lyric_ui_applying = False

    def _persist_lyric_style_from_session(self):
        if getattr(self, '_loading_settings', False) or getattr(self, '_lyric_ui_applying', False):
            return
        pair = getattr(self, '_lyric_pair', None)
        if pair is None:
            return
        self._persist_settings(
            lyric_align=normalize_lyric_align(pair.session.align),
            lyric_letter_spacing=normalize_letter_spacing(pair.session.letter_spacing),
            lyric_line_height=normalize_line_height(
                getattr(pair.session, 'line_height', DEFAULT_LINE_HEIGHT)
            ),
            lyric_margins=normalize_lyric_margins(
                getattr(pair.session, 'margins', DEFAULT_LYRIC_MARGINS)
            ),
            lyric_font_family=normalize_lyric_font_family(
                getattr(pair.session, 'font_family', DEFAULT_LYRIC_FONT_FAMILY)
            ),
        )

    def _on_lyric_style_changed(self):
        if getattr(self, '_loading_settings', False) or getattr(self, '_lyric_ui_applying', False):
            return
        pair = getattr(self, '_lyric_pair', None)
        if pair is None:
            return
        align = normalize_lyric_align(pair.session.align)
        spacing = normalize_letter_spacing(pair.session.letter_spacing)
        line_height = normalize_line_height(getattr(pair.session, 'line_height', DEFAULT_LINE_HEIGHT))
        margins = normalize_lyric_margins(getattr(pair.session, 'margins', DEFAULT_LYRIC_MARGINS))
        font_family = normalize_lyric_font_family(
            getattr(pair.session, 'font_family', DEFAULT_LYRIC_FONT_FAMILY)
        )
        timer = getattr(self, '_lyric_style_persist_debounce', None)
        if timer is not None:
            timer.start()
        else:
            self._persist_lyric_style_from_session()
        if hasattr(self, 'combo_lyric_align'):
            idx = self.combo_lyric_align.findData(align)
            self.combo_lyric_align.blockSignals(True)
            self.combo_lyric_align.setCurrentIndex(idx if idx >= 0 else 0)
            self.combo_lyric_align.blockSignals(False)
        if hasattr(self, 'slider_lyric_spacing'):
            self.slider_lyric_spacing.blockSignals(True)
            self.slider_lyric_spacing.setValue(spacing)
            self.slider_lyric_spacing.blockSignals(False)
        if hasattr(self, 'lbl_lyric_spacing_val'):
            self.lbl_lyric_spacing_val.setText(f'{spacing}%')
        if hasattr(self, 'slider_lyric_line_height'):
            self.slider_lyric_line_height.blockSignals(True)
            self.slider_lyric_line_height.setValue(line_height)
            self.slider_lyric_line_height.blockSignals(False)
        if hasattr(self, 'lbl_lyric_line_height_val'):
            self.lbl_lyric_line_height_val.setText(f'{line_height}%')
        if hasattr(self, 'spin_lyric_margins'):
            for side, spin in self.spin_lyric_margins.items():
                spin.blockSignals(True)
                spin.setValue(margins[side])
                spin.blockSignals(False)
        if hasattr(self, 'combo_lyric_font_family'):
            self.combo_lyric_font_family.blockSignals(True)
            self.combo_lyric_font_family.setCurrentFont(QFont(font_family))
            self.combo_lyric_font_family.blockSignals(False)

    def _on_lyric_placement_changed(self, _index=0):
        placement = self.combo_lyric_placement.currentData() or PLACEMENT_WINDOW
        self._persist_settings(lyric_control_placement=placement)
        if getattr(self, '_lyric_pair', None) is not None:
            self._lyric_pair.set_placement(placement, getattr(self, 'lyric_dock', None))

    def _on_lyric_start_blacked_toggled(self, checked):
        self._persist_settings(lyric_start_blacked=bool(checked))

    def _on_lyric_preview_mode_toggled(self, checked):
        checked = bool(checked)
        self._persist_settings(lyric_preview_mode=checked)
        pair = getattr(self, '_lyric_pair', None)
        if pair is not None:
            pair.set_preview_mode(checked)
            self._schedule_lyric_preview()

    def _on_lyric_go_live(self):
        pair = getattr(self, '_lyric_pair', None)
        if pair is None or not pair.session.hymn:
            return
        hymn = pair.session.hymn or {}
        name = hymn.get('heading') or hymn.get('name') or hymn.get('title') or ''
        self._show_open_hymn_popup(hymn.get('book', ''), name)
        self._schedule_lyric_preview()
        QTimer.singleShot(80, self._sync_lyric_preview)
        self._show_toast(f'🎥  已轉出現場：{name}')


    def _overlay_duration_value(self):
        if hasattr(self, 'chk_overlay_always') and self.chk_overlay_always.isChecked():
            return OVERLAY_DURATION_ALWAYS
        if hasattr(self, 'spin_overlay_duration'):
            return _normalize_overlay_duration(self.spin_overlay_duration.value())
        return _normalize_overlay_duration(self._settings.get('overlay_duration', DEFAULT_OVERLAY_SECS))

    def _overlay_center_secs(self):
        if hasattr(self, 'spin_overlay_center'):
            val = self.spin_overlay_center.value()
            if val is not None:
                return _normalize_center_secs(val)
        return _normalize_center_secs(self._settings.get('overlay_center_secs', CENTER_OVERLAY_SECS))

    def _sync_overlay_center_enabled(self):
        overlay_on = bool(self.chk_open_overlay.isChecked()) if hasattr(self, 'chk_open_overlay') else True
        center_mode = (
            hasattr(self, 'combo_overlay_pos')
            and self.combo_overlay_pos.currentData() == OVERLAY_MODE_CENTER
        )
        enabled = overlay_on and center_mode
        for w in (
            getattr(self, 'lbl_overlay_center', None),
            getattr(self, 'spin_overlay_center', None),
            getattr(self, 'chk_overlay_auto_unblack', None),
        ):
            if w is not None:
                w.setEnabled(enabled)

    def _on_overlay_center_changed(self, _idx):
        self._persist_settings(overlay_center_secs=self._overlay_center_secs())
        self._sync_book_name_overlay_settings()
        self._schedule_lyric_preview()

    def _on_overlay_auto_unblack_toggled(self, checked):
        checked = bool(checked)
        self._persist_settings(overlay_auto_unblack=checked)
        if hasattr(self._remote.state, 'set_overlay_auto_unblack'):
            self._remote.state.set_overlay_auto_unblack(checked)

    def _on_overlay_countdown_finished(self):
        self._schedule_lyric_preview()
        self._tick_overlay_countdown()
        if not self._settings.get('overlay_auto_unblack'):
            return
        pair = getattr(self, '_lyric_pair', None)
        if pair is not None and pair.session.blacked:
            pair.session.set_blacked(False)

    def _start_overlay_countdown_tick(self):
        timer = getattr(self, '_overlay_count_timer', None)
        if timer is None:
            return
        timer.start()
        self._tick_overlay_countdown()

    def _tick_overlay_countdown(self):
        pair = getattr(self, '_lyric_pair', None)
        ends = BookNameOverlay.ends_at_unix()
        text = ''
        if ends > 0:
            left = int(ends - time.time() + 0.999)
            if left > 0:
                text = f'轉移 {left} 秒'
                if self._settings.get('overlay_auto_unblack') and pair is not None and pair.session.blacked:
                    text += ' · 完咗取消全黑'
        if pair is not None:
            pair.set_countdown(text)
        timer = getattr(self, '_overlay_count_timer', None)
        if not text and timer is not None:
            timer.stop()

    def _schedule_lyric_preview(self):
        timer = getattr(self, '_preview_sync_timer', None)
        if timer is not None:
            timer.start()

    def _sync_lyric_preview(self):
        pair = getattr(self, '_lyric_pair', None)
        remote = getattr(self, '_remote', None)
        if pair is None or remote is None:
            return
        state = getattr(remote, 'state', None)
        # Local lyric windows update via session signals; remote push only when watched.
        if state is not None and hasattr(state, 'has_preview_subscribers'):
            try:
                if not state.has_preview_subscribers():
                    return
            except Exception:
                pass
        stage_w, stage_h = 1920, 1080
        hymn_open = bool(getattr(pair.session, 'hymn', None))
        fullscreen = False
        try:
            display = getattr(pair, 'display', None)
            canvas = getattr(display, 'canvas', None) if display is not None else None
            if canvas is not None and canvas.width() >= 40 and canvas.height() >= 40:
                stage_w = canvas.width()
                stage_h = canvas.height()
            fullscreen = bool(hymn_open and pair.is_display_fullscreen())
        except Exception:
            fullscreen = False
        snap = pair.session.preview_snapshot(
            stage_w, stage_h, display_open=hymn_open, fullscreen=fullscreen,
        )
        snap['overlay_ends_at'] = round(BookNameOverlay.ends_at_unix(), 3)
        remote.state.set_lyric_preview(snap)

    def _apply_lyric_settings(self):
        s = self._settings
        if getattr(self, '_lyric_pair', None) is None:
            return
        self._lyric_pair.session.set_colors(
            s.get('lyric_bg_color', DEFAULT_LYRIC_BG),
            s.get('lyric_text_color', DEFAULT_LYRIC_TEXT),
        )
        self._lyric_pair.session.set_bg_image(s.get('lyric_bg_image') or '')
        self._lyric_pair.session.set_align(s.get('lyric_align', DEFAULT_LYRIC_ALIGN))
        self._lyric_pair.session.set_letter_spacing(
            s.get('lyric_letter_spacing', DEFAULT_LETTER_SPACING)
        )
        self._lyric_pair.session.set_line_height(
            s.get('lyric_line_height', DEFAULT_LINE_HEIGHT)
        )
        self._lyric_pair.session.set_margins(s.get('lyric_margins', DEFAULT_LYRIC_MARGINS))
        self._lyric_pair.session.set_font_family(
            s.get('lyric_font_family', DEFAULT_LYRIC_FONT_FAMILY)
        )
        self._lyric_pair.session.set_preferred_font_pt(s.get('lyric_font_pt', DEFAULT_LYRIC_FONT_PT))
        # Keep preview staging style in sync with saved defaults
        prev = getattr(self._lyric_pair, 'preview_session', None)
        if prev is not None:
            prev.set_colors(
                s.get('lyric_bg_color', DEFAULT_LYRIC_BG),
                s.get('lyric_text_color', DEFAULT_LYRIC_TEXT),
            )
            prev.set_bg_image(s.get('lyric_bg_image') or '')
            prev.set_align(s.get('lyric_align', DEFAULT_LYRIC_ALIGN))
            prev.set_letter_spacing(s.get('lyric_letter_spacing', DEFAULT_LETTER_SPACING))
            prev.set_line_height(s.get('lyric_line_height', DEFAULT_LINE_HEIGHT))
            prev.set_margins(s.get('lyric_margins', DEFAULT_LYRIC_MARGINS))
            prev.set_font_family(s.get('lyric_font_family', DEFAULT_LYRIC_FONT_FAMILY))
            prev.set_preferred_font_pt(s.get('lyric_font_pt', DEFAULT_LYRIC_FONT_PT))
        self._lyric_pair.set_placement(
            s.get('lyric_control_placement', PLACEMENT_WINDOW),
            getattr(self, 'lyric_dock', None),
        )
        self._lyric_pair.set_preview_mode(bool(s.get('lyric_preview_mode', False)))
        if hasattr(self, 'chk_lyric_preview_mode'):
            self.chk_lyric_preview_mode.blockSignals(True)
            self.chk_lyric_preview_mode.setChecked(bool(s.get('lyric_preview_mode', False)))
            self.chk_lyric_preview_mode.blockSignals(False)

    def _update_overlay_color_previews(self):
        if hasattr(self, 'lbl_overlay_bg_preview'):
            self._update_qr_color_preview(
                self.lbl_overlay_bg_preview,
                self.inp_overlay_bg_color.text() if hasattr(self, 'inp_overlay_bg_color') else '',
                DEFAULT_OVERLAY_BG_COLOR,
            )
        if hasattr(self, 'lbl_overlay_text_color_preview'):
            self._update_qr_color_preview(
                self.lbl_overlay_text_color_preview,
                self.inp_overlay_text_color.text() if hasattr(self, 'inp_overlay_text_color') else '',
                '#ffffff',
            )

    def _reset_overlay_text_color_auto(self):
        if not hasattr(self, 'inp_overlay_text_color'):
            return
        self.inp_overlay_text_color.blockSignals(True)
        self.inp_overlay_text_color.clear()
        self.inp_overlay_text_color.blockSignals(False)
        self._update_overlay_color_previews()
        self._on_overlay_appearance_changed()

    def _sync_book_name_overlay_settings(self):
        panels = BookNameOverlay._visible_panels()
        if not panels:
            return
        ref = panels[0]
        BookNameOverlay.show_book(
            ref._book_label.text(),
            getattr(ref, '_watch_path', ''),
            ref._item_label.text() if ref._item_label.isVisible() else '',
            ref._corner_duration,
            ref._mode,
            self._settings,
            center_secs=self._overlay_center_secs(),
        )
        self._start_overlay_countdown_tick()

    def _on_overlay_appearance_changed(self, _value=''):
        values = self._overlay_appearance_values()
        self._update_overlay_color_previews()
        self._persist_settings(**values)
        self._sync_book_name_overlay_settings()

    def _sync_open_overlay_checkboxes(self, checked):
        for chk in (getattr(self, 'chk_open_overlay_hdr', None), self.chk_open_overlay):
            if chk is None:
                continue
            chk.blockSignals(True)
            chk.setChecked(checked)
            chk.blockSignals(False)

    def _sync_display_duplicate_checkboxes(self, checked):
        for chk in (getattr(self, 'chk_display_duplicate_hdr', None), self.chk_display_duplicate):
            if chk is None:
                continue
            chk.blockSignals(True)
            chk.setChecked(checked)
            chk.blockSignals(False)

    def _on_open_overlay_toggled(self, checked):
        self._sync_open_overlay_checkboxes(checked)
        for w in (self.lbl_overlay_secs, getattr(self, 'overlay_duration_row', None),
                  getattr(self, 'spin_overlay_duration', None),
                  getattr(self, 'chk_overlay_always', None),
                  self.lbl_overlay_pos, self.combo_overlay_pos):
            if w is not None:
                w.setEnabled(checked)
        if (
            checked
            and hasattr(self, 'chk_overlay_always')
            and self.chk_overlay_always.isChecked()
            and hasattr(self, 'spin_overlay_duration')
        ):
            self.spin_overlay_duration.setEnabled(False)
        for w in self._overlay_appearance_widgets():
            w.setEnabled(checked)
        self._sync_overlay_center_enabled()
        if not checked:
            BookNameOverlay.hide_overlay()
            self._schedule_lyric_preview()
        self._persist_settings(open_overlay=checked)

    def _on_open_overlay_hdr_toggled(self, checked):
        self._on_open_overlay_toggled(checked)

    def _on_overlay_always_toggled(self, checked):
        if hasattr(self, 'spin_overlay_duration'):
            self.spin_overlay_duration.setEnabled(not checked)
        self._on_overlay_duration_changed()

    def _on_overlay_duration_changed(self, _val=None):
        self._persist_settings(
            overlay_duration=self._overlay_duration_value()
        )
        self._sync_book_name_overlay_settings()
        self._schedule_lyric_preview()

    def _on_overlay_mode_changed(self, _idx):
        self._persist_settings(
            overlay_mode=self.combo_overlay_pos.currentData()
        )
        self._sync_overlay_center_enabled()
        self._sync_book_name_overlay_settings()

    def _on_display_duplicate_toggled(self, checked):
        self._sync_display_duplicate_checkboxes(checked)
        self._persist_settings(display_duplicate=checked)

    def _on_display_switch_mode_changed(self, _index=-1):
        mode = self.combo_display_switch_mode.currentData()
        if mode:
            self._persist_settings(display_switch_mode=mode)

    def _on_word_auto_fullscreen_toggled(self, checked):
        self._persist_settings(word_auto_fullscreen=checked)

    def _on_pdf_auto_fullscreen_toggled(self, checked):
        self._persist_settings(pdf_auto_fullscreen=checked)

    def _display_switch_method(self) -> str:
        if hasattr(self, 'combo_display_switch_mode'):
            mode = self.combo_display_switch_mode.currentData()
            if mode in (DISPLAY_SWITCH_CCD, DISPLAY_SWITCH_LEGACY):
                return mode
        return str(self._settings.get('display_switch_mode') or DISPLAY_SWITCH_CCD)

    def _begin_open_flow(self) -> int:
        self._open_flow_generation = self._open_flow.begin()
        return self._open_flow_generation

    def _cancel_open_flow(self):
        self._open_flow.cancel()

    def _hide_overlay_on_esc(self):
        BookNameOverlay.hide_overlay()
        self._schedule_lyric_preview()
        self._tick_overlay_countdown()

    def _on_projection_ready(self, generation: int):
        if not self._open_flow.should_switch_display(generation):
            return
        if sys.platform != 'win32' or not self.chk_display_duplicate.isChecked():
            return
        self.display_duplicate_requested.emit(generation)

    def _maybe_start_projection_flow(
        self,
        path: str,
        *,
        kind: str,
        auto_fullscreen: bool = True,
        skip_pdf_presentation: bool = False,
    ):
        gen = self._begin_open_flow()
        ready_cb = lambda g: self.projection_ready.emit(g)
        if kind == 'word':
            run_word_projection_flow(
                path, self._open_flow, gen, ready_cb,
                auto_fullscreen=auto_fullscreen,
            )
        elif kind == 'pdf':
            run_pdf_projection_flow(
                path, self._open_flow, gen, ready_cb,
                auto_fullscreen=auto_fullscreen,
                skip_presentation=skip_pdf_presentation,
            )

    def _on_keyword_instant_toggled(self, checked):
        self._persist_settings(keyword_instant=checked)

    def _on_book_auto_focus_hymn_toggled(self, checked):
        self._persist_settings(book_auto_focus_hymn=checked)

    def _hymn_num_mode(self):
        mode = self._settings.get('hymn_num_mode', HYMN_NUM_MODE_CONTAINS)
        return mode if mode in HYMN_NUM_MODES else HYMN_NUM_MODE_CONTAINS

    def _on_hymn_num_mode_changed(self, _index=None):
        if self._loading_settings:
            return
        mode = self.combo_hymn_num_mode.currentData()
        if mode not in HYMN_NUM_MODES:
            mode = HYMN_NUM_MODE_CONTAINS
        self._persist_settings(hymn_num_mode=mode)
        if getattr(self, '_remote', None):
            self._remote.state.set_hymn_num_mode(mode)
        if self.search_mode in BOOK_UI_MODES and self.current_book:
            self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
        elif self.search_mode == 'standard' and not self.inp_book.text().strip():
            self._apply_sidebar_filter()
        elif self.search_mode == 'schedule':
            self._refresh_schedule_search()

    def _sync_click_to_open_checkboxes(self, checked):
        for chk in (self.chk_click_to_open_hdr, self.chk_click_to_open):
            chk.blockSignals(True)
            chk.setChecked(checked)
            chk.blockSignals(False)

    def _on_click_to_open_toggled(self, checked):
        self._sync_click_to_open_checkboxes(checked)
        self._persist_settings(click_to_open=checked)
        try:
            self._remote.state.set_click_to_open(checked)
        except Exception:
            pass

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

    def _make_settings_qr_column(self, title, object_name, empty_text):
        col = QWidget()
        lay = QVBoxLayout(col)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        title_lbl = QLabel(title)
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        title_lbl.setObjectName('remoteQrTitle')
        qr_lbl = QLabel(empty_text)
        qr_lbl.setObjectName(object_name)
        qr_lbl.setFixedSize(120, 120)
        qr_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        qr_lbl.setWordWrap(True)
        link_lbl = QLabel("—")
        link_lbl.setWordWrap(True)
        link_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        link_lbl.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        lay.addWidget(title_lbl)
        lay.addWidget(qr_lbl, 0, Qt.AlignmentFlag.AlignHCenter)
        lay.addWidget(link_lbl)
        return col, qr_lbl, link_lbl, title_lbl

    def _make_settings_qr_pair(
        self, viewer_obj_name, admin_obj_name, viewer_empty, admin_empty,
    ):
        pair = QWidget()
        row = QHBoxLayout(pair)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(16)
        viewer_col, viewer_qr, viewer_link, _ = self._make_settings_qr_column(
            '普通（會眾）', viewer_obj_name, viewer_empty,
        )
        admin_col, admin_qr, admin_link, _ = self._make_settings_qr_column(
            'Admin（遙控）', admin_obj_name, admin_empty,
        )
        row.addWidget(viewer_col, 1)
        row.addWidget(admin_col, 1)
        return pair, viewer_qr, viewer_link, admin_qr, admin_link

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

    def _on_show_qr_role_toggled(self, _checked=False):
        self._persist_settings(
            show_qr_lan=self.chk_show_qr_lan.isChecked(),
            show_qr_fixed=self.chk_show_qr_fixed.isChecked(),
            show_qr_quick=self.chk_show_qr_quick.isChecked(),
            show_qr_wan=(
                self.chk_show_qr_fixed.isChecked() or self.chk_show_qr_quick.isChecked()
            ),
        )
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

    def _qr_caption_with_link(self, kind_label, url):
        title = str(self._settings.get('qr_caption', DEFAULT_QR_CAPTION) or '').strip()
        parts = []
        if title:
            parts.append(title)
        parts.append(str(kind_label or '').strip() or 'QR')
        if url:
            parts.append(str(url).strip())
        return '\n'.join(p for p in parts if p)

    def _show_one_desktop_qr(self, role, url, kind_label, x, y, card_w, card_h, style):
        caption = self._qr_caption_with_link(kind_label, url)
        # Keep QR readable when caption includes multi-line URL.
        need_h = (
            56
            + _desktop_qr_caption_height(caption, style['text_size'])
            + 10
        )
        use_h = max(int(card_h), need_h)
        qr_size = _desktop_qr_render_size(card_w, use_h, caption, style['text_size'])
        pix = self._qr_pixmap_from_url(url, size=qr_size, border=1)
        DesktopQrOverlay.show_qr(
            pix, caption, style, int(x), int(y), int(card_w), int(use_h), role=role,
        )
        return use_h

    def _sync_desktop_qr_overlay(self):
        if not self._settings.get('show_qr_code'):
            DesktopQrOverlay.hide()
            return
        if not getattr(self._remote, 'running', False):
            DesktopQrOverlay.hide()
            return
        try:
            style = _desktop_qr_style_from_settings(self._settings)
            card_w = int(self._settings.get('qr_width', DEFAULT_QR_WIDTH))
            card_h = int(self._settings.get('qr_height', DEFAULT_QR_HEIGHT))
            x0 = int(self._settings.get('qr_x', DEFAULT_QR_X))
            y0 = int(self._settings.get('qr_y', DEFAULT_QR_Y))
            gap = DEFAULT_QR_STACK_GAP
            y_next = y0
            any_shown = False

            def _maybe_show(role, enabled, url, kind):
                nonlocal y_next, any_shown
                if enabled and url:
                    h = self._show_one_desktop_qr(
                        role, url, kind, x0, y_next, card_w, card_h, style,
                    )
                    y_next += int(h) + gap
                    any_shown = True
                else:
                    DesktopQrOverlay.hide(role)

            _maybe_show(
                'lan',
                self._settings.get('show_qr_lan', True),
                self._remote_lan_viewer_url(),
                '內網',
            )
            _maybe_show(
                'fixed',
                self._settings.get('show_qr_fixed', True),
                self._remote_fixed_viewer_url(),
                '外網（固定）',
            )
            _maybe_show(
                'quick',
                self._settings.get('show_qr_quick', True),
                self._remote_quick_viewer_url(),
                '外網（隨機）',
            )
            # Hide legacy single-wan role if present.
            DesktopQrOverlay.hide('wan')
            if not any_shown:
                DesktopQrOverlay.hide()
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
        entries = list_entries_for_book(
            self.books, book_ref, q, num_mode=self._hymn_num_mode(),
        )
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
        matches = resolve_open_request(
            self.books, book_ref, num_ref, num_mode=self._hymn_num_mode(),
        )
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
        if hasattr(self._remote.state, 'set_hymn_num_mode'):
            self._remote.state.set_hymn_num_mode(s.get('hymn_num_mode', HYMN_NUM_MODE_CONTAINS))
        if hasattr(self._remote.state, 'set_click_to_open'):
            self._remote.state.set_click_to_open(s.get('click_to_open', False))
        if hasattr(self._remote.state, 'set_overlay_auto_unblack'):
            self._remote.state.set_overlay_auto_unblack(s.get('overlay_auto_unblack', False))
        if hasattr(self, 'chk_overlay_auto_unblack'):
            self.chk_overlay_auto_unblack.blockSignals(True)
            self.chk_overlay_auto_unblack.setChecked(s.get('overlay_auto_unblack', False))
            self.chk_overlay_auto_unblack.blockSignals(False)
        if hasattr(self, 'spin_overlay_center'):
            self.spin_overlay_center.blockSignals(True)
            self.spin_overlay_center.setValue(
                _normalize_center_secs(s.get('overlay_center_secs', CENTER_OVERLAY_SECS))
            )
            self.spin_overlay_center.blockSignals(False)
        if hasattr(self, 'spin_overlay_duration'):
            dur = _normalize_overlay_duration(s.get('overlay_duration'))
            always = dur == OVERLAY_DURATION_ALWAYS
            if hasattr(self, 'chk_overlay_always'):
                self.chk_overlay_always.blockSignals(True)
                self.chk_overlay_always.setChecked(always)
                self.chk_overlay_always.blockSignals(False)
            self.spin_overlay_duration.blockSignals(True)
            self.spin_overlay_duration.setValue(
                DEFAULT_OVERLAY_SECS if always else dur
            )
            self.spin_overlay_duration.setEnabled(not always)
            self.spin_overlay_duration.blockSignals(False)
        if hasattr(self, 'combo_hymn_num_mode'):
            idx = self.combo_hymn_num_mode.findData(
                s.get('hymn_num_mode', HYMN_NUM_MODE_CONTAINS)
            )
            self.combo_hymn_num_mode.blockSignals(True)
            self.combo_hymn_num_mode.setCurrentIndex(idx if idx >= 0 else 0)
            self.combo_hymn_num_mode.blockSignals(False)
        self._set_search_mode(s['search_mode'])
        self._on_open_overlay_toggled(s['open_overlay'])
        self._sync_display_duplicate_checkboxes(s['display_duplicate'])
        if hasattr(self, 'combo_display_switch_mode'):
            switch_mode = s.get('display_switch_mode') or DISPLAY_SWITCH_CCD
            idx = self.combo_display_switch_mode.findData(switch_mode)
            self.combo_display_switch_mode.blockSignals(True)
            self.combo_display_switch_mode.setCurrentIndex(idx if idx >= 0 else 0)
            self.combo_display_switch_mode.blockSignals(False)
        if hasattr(self, 'chk_word_auto_fullscreen'):
            self.chk_word_auto_fullscreen.blockSignals(True)
            self.chk_word_auto_fullscreen.setChecked(s.get('word_auto_fullscreen', True))
            self.chk_word_auto_fullscreen.blockSignals(False)
        if hasattr(self, 'chk_pdf_auto_fullscreen'):
            self.chk_pdf_auto_fullscreen.blockSignals(True)
            self.chk_pdf_auto_fullscreen.setChecked(s.get('pdf_auto_fullscreen', False))
            self.chk_pdf_auto_fullscreen.blockSignals(False)
        if hasattr(self, 'inp_overlay_bg_color'):
            self.inp_overlay_bg_color.blockSignals(True)
            self.inp_overlay_bg_color.setText(
                _normalize_hex_color(s.get('overlay_bg_color'), DEFAULT_OVERLAY_BG_COLOR)
            )
            self.inp_overlay_bg_color.blockSignals(False)
        if hasattr(self, 'inp_overlay_text_color'):
            self.inp_overlay_text_color.blockSignals(True)
            text_color = str(s.get('overlay_text_color') or '').strip()
            self.inp_overlay_text_color.setText(
                _normalize_hex_color(text_color, DEFAULT_OVERLAY_TEXT_COLOR) if text_color else ''
            )
            self.inp_overlay_text_color.blockSignals(False)
        if hasattr(self, 'inp_lyric_bg_color'):
            self.inp_lyric_bg_color.blockSignals(True)
            self.inp_lyric_bg_color.setText(
                _normalize_hex_color(s.get('lyric_bg_color'), DEFAULT_LYRIC_BG)
            )
            self.inp_lyric_bg_color.blockSignals(False)
        if hasattr(self, 'inp_lyric_text_color'):
            self.inp_lyric_text_color.blockSignals(True)
            self.inp_lyric_text_color.setText(
                _normalize_hex_color(s.get('lyric_text_color'), DEFAULT_LYRIC_TEXT)
            )
            self.inp_lyric_text_color.blockSignals(False)
        if hasattr(self, 'combo_lyric_placement'):
            idx = self.combo_lyric_placement.findData(
                s.get('lyric_control_placement', PLACEMENT_WINDOW)
            )
            self.combo_lyric_placement.blockSignals(True)
            self.combo_lyric_placement.setCurrentIndex(idx if idx >= 0 else 0)
            self.combo_lyric_placement.blockSignals(False)
        if hasattr(self, 'combo_lyric_align'):
            idx = self.combo_lyric_align.findData(s.get('lyric_align', DEFAULT_LYRIC_ALIGN))
            self.combo_lyric_align.blockSignals(True)
            self.combo_lyric_align.setCurrentIndex(idx if idx >= 0 else 0)
            self.combo_lyric_align.blockSignals(False)
        if hasattr(self, 'slider_lyric_spacing'):
            spacing = normalize_letter_spacing(s.get('lyric_letter_spacing'))
            self.slider_lyric_spacing.blockSignals(True)
            self.slider_lyric_spacing.setValue(spacing)
            self.slider_lyric_spacing.blockSignals(False)
            if hasattr(self, 'lbl_lyric_spacing_val'):
                self.lbl_lyric_spacing_val.setText(f'{spacing}%')
        if hasattr(self, 'slider_lyric_line_height'):
            line_height = normalize_line_height(s.get('lyric_line_height'))
            self.slider_lyric_line_height.blockSignals(True)
            self.slider_lyric_line_height.setValue(line_height)
            self.slider_lyric_line_height.blockSignals(False)
            if hasattr(self, 'lbl_lyric_line_height_val'):
                self.lbl_lyric_line_height_val.setText(f'{line_height}%')
        if hasattr(self, 'spin_lyric_margins'):
            margins = normalize_lyric_margins(s.get('lyric_margins'))
            for side, spin in self.spin_lyric_margins.items():
                spin.blockSignals(True)
                spin.setValue(margins[side])
                spin.blockSignals(False)
        if hasattr(self, 'combo_lyric_font_family'):
            family = normalize_lyric_font_family(s.get('lyric_font_family'))
            self.combo_lyric_font_family.blockSignals(True)
            self.combo_lyric_font_family.setCurrentFont(QFont(family))
            self.combo_lyric_font_family.blockSignals(False)
        if hasattr(self, 'chk_lyric_font_auto') and hasattr(self, 'spin_lyric_font'):
            font_pt = normalize_lyric_font_pt(s.get('lyric_font_pt'))
            self.chk_lyric_font_auto.blockSignals(True)
            self.spin_lyric_font.blockSignals(True)
            self.chk_lyric_font_auto.setChecked(font_pt <= 0)
            self.spin_lyric_font.setValue(font_pt if font_pt > 0 else 48)
            self.spin_lyric_font.setEnabled(font_pt > 0)
            self.chk_lyric_font_auto.blockSignals(False)
            self.spin_lyric_font.blockSignals(False)
        if hasattr(self, 'chk_lyric_start_blacked'):
            self.chk_lyric_start_blacked.blockSignals(True)
            self.chk_lyric_start_blacked.setChecked(bool(s.get('lyric_start_blacked', True)))
            self.chk_lyric_start_blacked.blockSignals(False)
        if hasattr(self, 'lbl_lyric_bg_preview'):
            self._update_qr_color_preview(
                self.lbl_lyric_bg_preview, s.get('lyric_bg_color'), DEFAULT_LYRIC_BG,
            )
            self._update_qr_color_preview(
                self.lbl_lyric_text_preview, s.get('lyric_text_color'), DEFAULT_LYRIC_TEXT,
            )
        self._apply_lyric_settings()
        self._sync_lyric_bg_image_label()
        self._update_overlay_color_previews()
        if s['remote_api_enabled']:
            self.chk_remote_api.setChecked(True)
            self._on_remote_api_toggled(True)
            if s['remote_accept']:
                self._sync_remote_accept_checkboxes(True)
                self._remote.state.set_accepting(True)
        for chk, key in (
            (self.chk_show_preview, 'show_preview'),
            (self.chk_lyric_start_blacked, 'lyric_start_blacked'),
            (self.chk_auto_rescan, 'auto_rescan'),
            (self.chk_startup_tray, 'startup_tray'),
            (self.chk_minimize_tray, 'minimize_to_tray'),
            (self.chk_viewer_show_debug, 'viewer_show_debug'),
            (self.chk_show_qr_lan, 'show_qr_lan'),
            (self.chk_show_qr_fixed, 'show_qr_fixed'),
            (self.chk_show_qr_quick, 'show_qr_quick'),
            (self.chk_enable_tunnel_fixed, 'enable_tunnel_fixed'),
            (self.chk_enable_tunnel_quick, 'enable_tunnel_quick'),
        ):
            if chk is None:
                continue
            chk.blockSignals(True)
            chk.setChecked(bool(s.get(key)))
            chk.blockSignals(False)
        self.enable_tunnel_fixed = bool(s.get('enable_tunnel_fixed'))
        self.enable_tunnel_quick = bool(s.get('enable_tunnel_quick'))
        self.enable_tunnel = self.enable_tunnel_fixed or self.enable_tunnel_quick
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
        if self.senior_mode:
            self._apply_senior_ui(True, persist=False)
            self._sync_font_size_labels()
            self._apply_theme()
        else:
            self._apply_senior_ui(False, persist=False)
            self._sync_font_size_labels()

    def _on_senior_mode_toggled(self, checked):
        if self._loading_settings:
            return
        if bool(checked) == bool(self.senior_mode):
            return
        if checked:
            self._enter_senior_mode()
        else:
            self._exit_senior_mode()

    def _enter_senior_mode(self):
        snap = {
            'theme': self.theme_name if self.theme_name in THEMES else 'dark',
            'font_size': max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, int(self.font_size))),
            'click_to_open': bool(self._settings.get('click_to_open')),
            'search_mode': self.search_mode if self.search_mode in SEARCH_MODES else 'standard',
        }
        self.senior_mode = True
        self._settings['senior_snapshot'] = snap
        if self.btn_settings.isChecked():
            self.btn_settings.setChecked(False)
        self.theme_name = 'light'
        self.theme_combo.blockSignals(True)
        self.theme_combo.setCurrentText('light')
        self.theme_combo.blockSignals(False)
        self._sync_click_to_open_checkboxes(True)
        if hasattr(self._remote.state, 'set_click_to_open'):
            self._remote.state.set_click_to_open(True)
        self._set_search_mode('standard')
        # 進入簡易版時若字偏小，升到預設大字；已夠大則保留
        enter_fs = max(SENIOR_FONT_SIZE, int(self.font_size))
        enter_fs = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, enter_fs))
        self._apply_senior_ui(True, persist=False)
        self._set_font_size(enter_fs, persist=False)
        self._apply_theme()
        self._persist_settings(
            senior_mode=True,
            senior_snapshot=snap,
            theme='light',
            font_size=enter_fs,
            click_to_open=True,
            search_mode='standard',
        )
        self.btn_senior.blockSignals(True)
        self.btn_senior.setChecked(True)
        self.btn_senior.blockSignals(False)
        self._show_toast('已切換至簡易版', 'ok')

    def _exit_senior_mode(self):
        snap = self._settings.get('senior_snapshot') or {}
        theme = snap.get('theme') if snap.get('theme') in THEMES else 'dark'
        try:
            font_size = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, int(snap.get('font_size', 13))))
        except (TypeError, ValueError):
            font_size = 13
        click_to_open = bool(snap.get('click_to_open', False))
        search_mode = snap.get('search_mode')
        if search_mode not in SEARCH_MODES:
            search_mode = 'standard'
        self.senior_mode = False
        self.theme_name = theme
        self.theme_combo.blockSignals(True)
        self.theme_combo.setCurrentText(theme)
        self.theme_combo.blockSignals(False)
        self._sync_click_to_open_checkboxes(click_to_open)
        if hasattr(self._remote.state, 'set_click_to_open'):
            self._remote.state.set_click_to_open(click_to_open)
        self._apply_senior_ui(False, persist=False)
        self._set_search_mode(search_mode)
        self._set_font_size(font_size, persist=False)
        self._apply_theme()
        self._persist_settings(
            senior_mode=False,
            senior_snapshot=None,
            theme=theme,
            font_size=font_size,
            click_to_open=click_to_open,
            search_mode=search_mode,
        )
        self.btn_senior.blockSignals(True)
        self.btn_senior.setChecked(False)
        self.btn_senior.blockSignals(False)
        self._show_toast('已返回標準版', 'ok')

    def _apply_senior_ui(self, enabled, persist=True):
        """Show/hide complex chrome and enlarge book-mode controls."""
        enabled = bool(enabled)
        self.senior_mode = enabled
        input_h = SENIOR_INPUT_H if enabled else NORMAL_INPUT_H
        filter_w = SENIOR_BTN_W if enabled else NORMAL_FILTER_BTN_W
        side_w = SENIOR_SIDEBAR_W if enabled else NORMAL_SIDEBAR_W
        hdr_h = SENIOR_HEADER_H if enabled else NORMAL_HEADER_H
        sb_hdr_h = SENIOR_SIDEBAR_HDR_H if enabled else NORMAL_SIDEBAR_HDR_H

        self.header_frame.setFixedHeight(hdr_h)
        self.sidebar_frame.setFixedWidth(side_w)
        self.sidebar_hdr.setFixedHeight(sb_hdr_h)

        self.mode_row_widget.setVisible(not enabled)
        self.header_opts.setVisible(not enabled)
        self.btn_reopen_last.setVisible(not enabled)
        self.legend_widget.setVisible(not enabled)
        self.btn_sess_prev.setVisible(not enabled)
        self.btn_sess_next.setVisible(not enabled)
        if hasattr(self, 'preview_lbl'):
            self.preview_lbl.setVisible(not enabled and bool(self._settings.get('show_preview', True)))
        # 簡易版仍可開設定（調字級等）；進入時先收起 overlay
        if enabled and hasattr(self, 'settings_overlay'):
            self.settings_overlay.setVisible(False)
            self.btn_settings.blockSignals(True)
            self.btn_settings.setChecked(False)
            self.btn_settings.setText("⚙  設定")
            self.btn_settings.blockSignals(False)

        self.inp_book.setFixedHeight(input_h)
        self.inp_num.setFixedHeight(input_h)
        if hasattr(self, 'btn_voice'):
            self.btn_voice.setFixedSize(input_h, input_h)
        self.btn_filter.setFixedHeight(input_h)
        self.btn_filter.setFixedWidth(filter_w)
        self.btn_senior.setFixedHeight(40 if enabled else 30)
        self.btn_senior.setText("標準版" if enabled else "簡易版")
        if hasattr(self, 'btn_phone_control'):
            self.btn_phone_control.setFixedHeight(40 if enabled else 30)

        if enabled:
            self.inp_book.setPlaceholderText("輸入書冊號／名稱…（可留空）")
            self.inp_num.setPlaceholderText("詩歌號或歌名（無書冊＝全局搜索）")
            self.hint_lbl.setText("可輸入書冊再搜詩歌；無書冊時歌名會全局搜索")
            self.file_hdr.setText("詩歌列表")
            self.hdr_title.setText("詩歌冊搜索 · 簡易版")
            if self.search_mode != 'standard':
                self._set_search_mode('standard')
        else:
            self.inp_book.setPlaceholderText("請按 Ctrl+Enter 開始輸入書冊名稱...")
            self.inp_num.setPlaceholderText("詩歌號/名")
            self.hdr_title.setText("詩歌冊搜索")
            if self.search_mode in BOOK_UI_MODES:
                self.file_hdr.setText("檔案 / 書籤")
                if self.search_mode == 'standard':
                    self.hint_lbl.setText("有書冊則搜該書；無書冊時歌名做全局搜索")
                else:
                    self.hint_lbl.setText("選一個左邊書冊，右邊會列出該書的所有檔案")

        self._update_search_action_buttons()
        if persist and not self._loading_settings:
            self._persist_settings(senior_mode=enabled)

    def _on_settings_toggled(self, checked):
        if not hasattr(self, 'settings_overlay'):
            return
        self.settings_overlay.setVisible(checked)
        self.btn_settings.setText("⚙  收起" if checked else "⚙  設定")
        if checked:
            self._sync_settings_overlay_geometry()
            self.settings_overlay.raise_()
            QTimer.singleShot(0, self._refresh_settings_layout)
            # Defer QR; opening settings should not block on QR/desktop sync.
            QTimer.singleShot(50, lambda: self._update_qr_code(sync_desktop=False))

    def _on_settings_tab_changed(self, index):
        if index == getattr(self, '_settings_mobile_tab_index', 2):
            QTimer.singleShot(0, self._refresh_mobile_settings_tab)

    def _refresh_mobile_settings_tab(self):
        """Light refresh when opening Mobile tab (avoid double QR + sc query lag)."""
        self._sync_tunnel_options_visibility(update_qr=False)
        self._update_remote_url_labels()
        self._refresh_remote_log_view()
        # One QR pass for settings previews only; desktop overlay already synced elsewhere.
        self._update_qr_code(sync_desktop=False)

    def _browse_hymn_folder(self):
        start = self.inp_hymn_folder.text().strip() or self.hymns_json_path
        if os.path.isfile(start):
            start_dir = os.path.dirname(start)
        elif os.path.isdir(start):
            start_dir = start
        else:
            start_dir = os.path.dirname(default_hymns_json_path())
        path, _ = QFileDialog.getOpenFileName(
            self, "選擇 hymns.json", start_dir, "JSON (*.json);;All files (*.*)",
        )
        if path:
            self.inp_hymn_folder.setText(path)

    def _open_hymn_editor(self, hymn_id='', volume_id=''):
        from hymn_features.hymn_editor import HymnCatalogEditorDialog
        path = resolve_hymns_json_path(
            (self.inp_hymn_folder.text().strip() if hasattr(self, 'inp_hymn_folder') else '')
            or getattr(self, 'hymns_json_path', '')
            or self.hymn_folder
        )
        dlg = HymnCatalogEditorDialog(
            path,
            parent=self,
            select_hymn_id=hymn_id or '',
            select_volume_id=volume_id or '',
            theme=getattr(self, 'theme_name', 'dark'),
            font_size=getattr(self, 'font_size', 14),
        )
        dlg.exec()
        if not dlg.did_save:
            return
        self.hymns_json_path = dlg.json_path
        self.hymn_folder = os.path.dirname(dlg.json_path)
        if hasattr(self, 'inp_hymn_folder'):
            self.inp_hymn_folder.setText(dlg.json_path)
        self._persist_settings(hymn_folder=dlg.json_path)
        self._scan_all()
        self._refresh_open_hymn_after_edit()
        self._show_toast(f'已更新詩歌庫 · {dlg.saved_count} 首', 'ok')

    def _refresh_open_hymn_after_edit(self):
        pair = getattr(self, '_lyric_pair', None)
        session = getattr(pair, 'session', None) if pair else None
        hymn = getattr(session, 'hymn', None) if session else None
        if not pair or not hymn:
            return
        book, found = find_hymn_in_books(self.books, hymn_id=hymn.get('id') or '')
        if not found:
            return
        try:
            pair.open_hymn(hymn_payload(book, found), keep_blacked=True)
        except Exception as exc:
            print(f'[WARN] refresh open hymn: {exc}')

    def _apply_hymn_folder(self):
        path = self.inp_hymn_folder.text().strip() or default_hymns_json_path()
        path = resolve_hymns_json_path(path)
        if path == os.path.abspath(self.hymns_json_path):
            self._show_toast('資料路徑未變更', 'warn')
            return
        self.hymns_json_path = path
        self.hymn_folder = os.path.dirname(path)
        self._persist_settings(hymn_folder=path)
        self._stop_index_worker()
        self._stop_search_worker()
        self.content_index = None
        self.current_book = None
        self._scan_all()
        self._show_toast(f'已載入：{os.path.basename(path) or path}', 'ok')

    def _refresh_settings_layout(self):
        for w in (self.grp_data, self.grp_appearance, self.grp_open,
                  self.settings_tabs, self.settings_frame):
            w.updateGeometry()
            w.adjustSize()

    def _update_settings_scroll_height(self):
        pass

    def _switch_duplicate_with_overlay_fix(self, generation: int):
        if not self._open_flow.should_switch_display(generation):
            return
        if sys.platform != 'win32' or not self.chk_display_duplicate.isChecked():
            return
        QTimer.singleShot(150, lambda g=generation: self._apply_display_mode('duplicate', g))

    def _refresh_overlays_after_display(self):
        try:
            app = QApplication.instance()
            if app:
                app.processEvents()
            BookNameOverlay.refresh_after_display_change()
            DesktopQrOverlay.refresh_after_display_change()
        except Exception as exc:
            print(f'[WARN] overlay refresh after display change: {exc}')

    def _apply_display_mode(self, mode='duplicate', generation: int | None = None):
        if generation is not None and not self._open_flow.should_switch_display(generation):
            return
        if sys.platform != 'win32':
            return
        BookNameOverlay.prepare_for_display_change()
        method = self._display_switch_method()

        def _worker():
            try:
                switch_windows_display_mode(mode, method=method)
            except Exception as exc:
                print(f'[WARN] display switch: {exc}')
            self.overlay_refresh_requested.emit()

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_display_duplicate(self, generation: int | None = None):
        if sys.platform != 'win32' or not self.chk_display_duplicate.isChecked():
            return
        self._apply_display_mode('duplicate', generation)

    def _open_docx_with_flow(self, path, item_name='', *, show_overlay=True):
        self._last_opened_path = path or ''
        if show_overlay:
            self._show_open_book_popup(path, item_name or os.path.basename(path))
        open_docx(path)
        self._maybe_start_projection_flow(
            path,
            kind='word',
            auto_fullscreen=self._settings.get('word_auto_fullscreen', True),
        )

    def _open_pdf_with_flow(self, path, item_name='', page=None, *, show_overlay=True):
        self._last_opened_path = path or ''
        if show_overlay:
            self._show_open_book_popup(path, item_name or os.path.basename(path))
        auto_fs = self._settings.get('pdf_auto_fullscreen', False)
        use_presentation = False
        if page is not None:
            _, viewer = open_pdf_at_page(path, page, presentation=auto_fs)
            use_presentation = auto_fs and viewer == 'Sumatra PDF'
        elif sys.platform == 'win32':
            if auto_fs:
                _, viewer = open_pdf_presentation(path)
                use_presentation = viewer == 'Sumatra PDF'
            else:
                os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
        self._maybe_start_projection_flow(
            path,
            kind='pdf',
            auto_fullscreen=auto_fs,
            skip_pdf_presentation=use_presentation,
        )

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

        # QFontComboBox popup uses a custom delegate that ignores dark item colors
        if hasattr(self, 'combo_lyric_font_family'):
            view = self.combo_lyric_font_family.view()
            if view is not None:
                vpal = QPalette(view.palette())
                vpal.setColor(QPalette.ColorRole.Base, QColor('#ffffff'))
                vpal.setColor(QPalette.ColorRole.AlternateBase, QColor('#f4f5fb'))
                vpal.setColor(QPalette.ColorRole.Text, QColor('#1a1e40'))
                vpal.setColor(QPalette.ColorRole.WindowText, QColor('#1a1e40'))
                vpal.setColor(QPalette.ColorRole.HighlightedText, QColor('#ffffff'))
                vpal.setColor(QPalette.ColorRole.Highlight, QColor(t['accent']))
                view.setPalette(vpal)
                view.setStyleSheet(
                    f"background: #ffffff; color: #1a1e40; outline: none; "
                    f"selection-background-color: {t['accent']}; selection-color: #ffffff;"
                )

        # ── Header ──────────────────────────────────────────────
        self.header_frame.setStyleSheet(
            f"background: {t['panel']}; border-bottom: 1px solid {t['border']};"
        )
        self.hdr_title.setStyleSheet(
            f"font-size: {22 if getattr(self, 'senior_mode', False) else 17}px; "
            f"font-weight: bold; color: {t['text']}; "
            f"padding-right: 5px; background: transparent;"
        )
        self.hdr_app_ver.setStyleSheet(
            f"font-size: {13 if getattr(self, 'senior_mode', False) else 11}px; color: {t['muted']}; "
            f"padding-top: 5px; padding-right: 8px; background: transparent;"
        )
        hdr_chk_style = (
            f"QCheckBox {{ font-size: {max(10, fs - 3)}px; color: {t['text2']}; "
            f"background: transparent; spacing: 6px; padding: 0 2px; }}"
        )
        self.chk_click_to_open_hdr.setStyleSheet(hdr_chk_style)
        self.chk_open_overlay_hdr.setStyleSheet(hdr_chk_style)

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
        if hasattr(self, 'chk_lyric_preview_mode'):
            self.chk_lyric_preview_mode.setStyleSheet(
                f"color: {t['text2']}; font-size: {max(10, fs - 2)}px; background: transparent;"
            )

        # ── Sidebar ──────────────────────────────────────────────
        self.sidebar_frame.setStyleSheet(
            f"QFrame {{ background: {t['panel']}; border-right: 1px solid {t['border']}; }}"
        )
        self.sidebar_hdr.setStyleSheet(
            f"QWidget#sidebarHeader {{ background: {t['panel']}; "
            f"border: none; border-bottom: 1px solid {t['border']}; }}"
        )
        self.sb_title.setStyleSheet(
            f"font-size: {16 if getattr(self, 'senior_mode', False) else 13}px; "
            f"font-weight: bold; color: {t['text']}; "
            f"background: transparent; border: none; padding: 0;"
        )
        self.book_count_lbl.setStyleSheet(
            f"font-size: {13 if getattr(self, 'senior_mode', False) else 11}px; "
            f"color: {t['text2']}; background: {t['panel2']}; "
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
        self.btn_standard.setStyleSheet(toggle_style)
        self.btn_book.setStyleSheet(toggle_style)
        self.btn_global.setStyleSheet(toggle_style)
        self.btn_keyword.setStyleSheet(toggle_style)
        self.btn_schedule.setStyleSheet(toggle_style)
        hdr_btn_style = f"""
            QPushButton {{
                border-radius: 14px; padding: 4px 12px;
                font-size: {max(10, fs - 2)}px; font-weight: 600;
                border: 1.5px solid {t['in_bd']};
                color: {t['text']}; background: {t['panel2']};
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
        if hasattr(self, 'btn_phone_control'):
            self.btn_phone_control.setStyleSheet(hdr_btn_style)
        senior_btn_style = f"""
            QPushButton {{
                border-radius: 14px; padding: 4px 14px;
                font-size: {fs if getattr(self, 'senior_mode', False) else max(10, fs - 1)}px;
                font-weight: 700;
                border: 1.5px solid {t['accent']};
                color: {t['a_text'] if getattr(self, 'senior_mode', False) else t['accent']};
                background: {t['accent'] if getattr(self, 'senior_mode', False) else 'transparent'};
            }}
            QPushButton:hover {{
                background: {t['accent_h']}; border-color: {t['accent_h']}; color: {t['a_text']};
            }}
            QPushButton:checked {{
                background: {t['accent']}; border-color: {t['accent']}; color: {t['a_text']};
            }}
        """
        if hasattr(self, 'btn_senior'):
            self.btn_senior.setStyleSheet(senior_btn_style)
        font_step_style = f"""
            QPushButton {{
                border-radius: 8px; padding: 0;
                font-size: {max(12, fs - 1)}px; font-weight: 700;
                border: 1.5px solid {t['in_bd']};
                color: {t['text']}; background: {t['in_bg']};
            }}
            QPushButton:hover:enabled {{
                border-color: {t['accent_h']}; color: {t['accent_h']};
            }}
            QPushButton:disabled {{
                color: {t['muted']}; border-color: {t['border']};
            }}
        """
        for btn in (
            getattr(self, 'btn_font_minus_settings', None),
            getattr(self, 'btn_font_plus_settings', None),
        ):
            if btn is not None:
                btn.setStyleSheet(font_step_style)
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
        qr_preview_ss = (
            f"background: #ffffff; color: {t['muted']}; "
            f"border: 1px solid {t['border']}; border-radius: 8px; "
            f"font-size: {max(9, fs - 3)}px; padding: 4px;"
        )
        for lbl in (
            getattr(self, 'remote_qr_lan_lbl', None),
            getattr(self, 'remote_qr_lan_admin_lbl', None),
            getattr(self, 'remote_qr_fixed_lbl', None),
            getattr(self, 'remote_qr_fixed_admin_lbl', None),
            getattr(self, 'remote_qr_quick_lbl', None),
            getattr(self, 'remote_qr_quick_admin_lbl', None),
            getattr(self, 'remote_qr_wan_lbl', None),
            getattr(self, 'remote_qr_lbl', None),
        ):
            if lbl is None:
                continue
            name = lbl.objectName() or 'remoteQrLabel'
            lbl.setStyleSheet(f"QLabel#{name} {{ {qr_preview_ss} }}")
        for link in (
            getattr(self, 'remote_qr_lan_link', None),
            getattr(self, 'remote_qr_lan_admin_link', None),
            getattr(self, 'remote_qr_fixed_link', None),
            getattr(self, 'remote_qr_fixed_admin_link', None),
            getattr(self, 'remote_qr_quick_link', None),
            getattr(self, 'remote_qr_quick_admin_link', None),
            getattr(self, 'remote_qr_wan_link', None),
        ):
            if link is not None:
                link.setStyleSheet(
                    f"font-size: {max(9, fs - 3)}px; color: {t['text2']}; background: transparent;"
                )
        for title in self.findChildren(QLabel, 'remoteQrTitle'):
            title.setStyleSheet(
                f"font-size: {max(10, fs - 2)}px; font-weight: 600; "
                f"color: {t['text']}; background: transparent;"
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
                background: #e8ebff; color: #1a1e40;
                border: 1.5px solid #c5c9ee; border-radius: 8px;
                padding: 4px 14px; font-size: {fs - 2}px; font-weight: 600;
            }}
            QPushButton:hover {{ background: #ffffff; border-color: {t['accent_h']}; color: #1a1e40; }}
            QPushButton:disabled {{ background: {t['panel2']}; color: {t['muted']}; border-color: {t['border']}; }}
        """
        self.btn_build_index.setStyleSheet(index_btn_style)
        self.btn_refresh_index.setStyleSheet(index_btn_style)
        self.btn_browse_hymn_folder.setStyleSheet(index_btn_style)
        self.btn_apply_hymn_folder.setStyleSheet(action_style)
        if hasattr(self, 'btn_edit_hymns'):
            self.btn_edit_hymns.setStyleSheet(index_btn_style)
        self.btn_export_settings.setStyleSheet(index_btn_style)
        self.btn_import_settings.setStyleSheet(index_btn_style)
        if hasattr(self, 'about_ver_lbl'):
            self.about_ver_lbl.setStyleSheet(
                f"font-size: {fs + 1}px; font-weight: 600; color: {t['text']}; background: transparent;"
            )
        compact_btn_style = f"""
            QPushButton {{
                background: #e8ebff; color: #1a1e40;
                border: 1px solid #c5c9ee; border-radius: 6px;
                font-size: {max(10, fs - 3)}px; font-weight: 600; padding: 0;
            }}
            QPushButton:hover {{ background: #ffffff; border-color: {t['accent_h']}; color: #1a1e40; }}
        """
        for btn in (self.btn_sess_prev, self.btn_sess_next,
                    self.btn_sched_prev, self.btn_sched_next,
                    self.btn_sched_add, self.btn_sched_open,
                    self.btn_sched_remove, self.btn_sched_up,
                    self.btn_sched_down, self.btn_sched_clear):
            btn.setStyleSheet(compact_btn_style)
        if hasattr(self, 'btn_voice'):
            voice_idle = compact_btn_style
            voice_listen = f"""
                QPushButton {{
                    background: {t['err_bg']}; color: {t['err']};
                    border: 1.5px solid {t['err']}; border-radius: 6px;
                    font-size: {max(12, fs - 1)}px; font-weight: 700; padding: 0;
                }}
                QPushButton:hover {{ background: {t['err_bg']}; color: {t['err']}; }}
            """
            self._voice_btn_idle_style = voice_idle
            self._voice_btn_listen_style = voice_listen
            listening = bool(getattr(self, '_voice', None) and self._voice.is_listening)
            self.btn_voice.setStyleSheet(voice_listen if listening else voice_idle)
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
        if hasattr(self, 'chk_word_auto_fullscreen'):
            self.chk_word_auto_fullscreen.setStyleSheet(checkbox_style)
        if hasattr(self, 'chk_pdf_auto_fullscreen'):
            self.chk_pdf_auto_fullscreen.setStyleSheet(checkbox_style)
        for chk in (
            self.chk_remote_api, self.chk_remote_accept,
            self.chk_enable_tunnel_fixed, self.chk_enable_tunnel_quick,
            self.chk_show_qr_lan, self.chk_show_qr_fixed, self.chk_show_qr_quick,
            self.chk_viewer_show_debug,
            self.chk_show_preview,
            getattr(self, 'chk_overlay_always', None),
            getattr(self, 'chk_lyric_start_blacked', None),
            self.chk_auto_rescan, self.chk_startup_tray, self.chk_minimize_tray,
        ):
            if chk is not None:
                chk.setStyleSheet(checkbox_style)
        for lbl in (self.lbl_remote_policy, self.lbl_remote_port,
                    self.lbl_remote_token, self.lbl_remote_url, self.lbl_viewer_url,
                    self.lbl_quick_url, getattr(self, 'lbl_quick_admin_url', None),
                    self.lbl_local_url, getattr(self, 'lbl_local_admin_url', None),
                    self.lbl_public_tunnel_url):
            if lbl is not None:
                lbl.setStyleSheet(
                    f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
                )
        self.remote_url_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['accent_h']}; background: transparent;"
        )
        self.viewer_url_lbl.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['ok']}; background: transparent;"
        )
        if hasattr(self, 'quick_url_lbl'):
            self.quick_url_lbl.setStyleSheet(
                f"font-size: {fs - 2}px; color: {t['ok']}; background: transparent;"
            )
        if hasattr(self, 'quick_admin_url_lbl'):
            self.quick_admin_url_lbl.setStyleSheet(
                f"font-size: {fs - 2}px; color: {t['accent_h']}; background: transparent;"
            )
        if hasattr(self, 'local_url_lbl'):
            self.local_url_lbl.setStyleSheet(
                f"font-size: {fs - 2}px; color: {t['ok']}; background: transparent;"
            )
        if hasattr(self, 'local_admin_url_lbl'):
            self.local_admin_url_lbl.setStyleSheet(
                f"font-size: {fs - 2}px; color: {t['accent_h']}; background: transparent;"
            )
        for status_lbl in (
            getattr(self, 'tunnel_status_fixed_lbl', None),
            getattr(self, 'tunnel_status_quick_lbl', None),
            getattr(self, 'tunnel_status_lbl', None),
        ):
            if status_lbl is not None:
                status_lbl.setStyleSheet(
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
        for lbl in (self.lbl_overlay_bg, self.lbl_overlay_text_color):
            lbl.setStyleSheet(
                f"font-size: {fs - 2}px; color: {t['muted']}; background: transparent;"
            )

        # sel_info
        self.sel_info.setStyleSheet(
            f"font-size: {fs - 2}px; color: {t['accent_h']}; background: transparent;"
        )

        # Settings labels
        for lbl in (self.lbl_hymn_folder, self.lbl_theme, self.lbl_font, self.lbl_hymn_num_mode,
                    self.lbl_overlay_secs, self.lbl_overlay_pos, self.lbl_overlay_center, self.fs_lbl,
                    self.lbl_overlay_bg, self.lbl_overlay_text_color,
                    self.lbl_lyric_bg, getattr(self, 'lbl_lyric_bg_image', None),
                    self.lbl_lyric_text, self.lbl_lyric_align,
                    self.lbl_lyric_spacing, self.lbl_lyric_line_height,
                    self.lbl_lyric_margin, getattr(self, 'lbl_lyric_font_family', None),
                    self.lbl_lyric_font, self.lbl_lyric_best,
                    self.lbl_lyric_placement,
                    self.lbl_remote_policy, self.lbl_remote_port, self.lbl_remote_token, self.lbl_remote_url,
                    self.lbl_viewer_url, self.lbl_local_url, self.lbl_quick_url,
                    getattr(self, 'lbl_quick_admin_url', None),
                    getattr(self, 'lbl_local_admin_url', None),
                    self.lbl_public_tunnel_url,
                    self.lbl_qr_x, self.lbl_qr_y, self.lbl_qr_w, self.lbl_qr_h,
                    self.lbl_qr_caption, self.lbl_qr_bg, self.lbl_qr_text_color,
                    self.lbl_qr_text_size, self.lbl_qr_bg_opacity):
            if lbl is not None:
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

        if getattr(self, '_lyric_pair', None) is not None:
            self._lyric_pair.set_theme(self.theme_name)

        # Dropdown
        self.dropdown.set_theme(self.theme_name, self.font_size)
        if hasattr(self, 'sched_entry_dropdown'):
            self.sched_entry_dropdown.set_theme(self.theme_name, self.font_size)

        self._update_qr_code()
        self._update_search_action_buttons()

    def _scan_all(self):
        t = THEMES[self.theme_name]
        self.hymns_json_path = resolve_hymns_json_path(
            getattr(self, 'hymns_json_path', '') or self.hymn_folder
        )
        self.books = scan_books(self.hymns_json_path)

        if not self.books:
            self.book_list.clear()
            item = QListWidgetItem("  ⚠  找不到詩歌資料")
            item.setForeground(QColor(t['err']))
            self.book_list.addItem(item)
            self._set_status('err', f"找不到：{os.path.basename(self.hymns_json_path)}")
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
        folder = self.hymn_folder
        if os.path.isfile(folder):
            folder = os.path.dirname(folder)
        if not os.path.isdir(folder):
            self.content_index = None
            self._update_index_status_label()
            return
        self.content_index = load_content_index(folder)
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
            self._set_search_mode('standard')
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
        if self.current_book and self.search_mode in BOOK_UI_MODES:
            self._show_book_files(self.current_book, '')
            self._update_search_action_buttons()
        elif self.search_mode == 'standard':
            self._apply_sidebar_filter()
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
            targets = resolve_targets(book, q, num_mode=self._hymn_num_mode())
            if not targets:
                ni = QListWidgetItem("  ⚠  沒有符合條件的詩歌")
                ni.setForeground(QColor(t['muted']))
                self.file_list.addItem(ni)
            else:
                for payload in targets:
                    if payload.get('kind') == 'hymn':
                        item = QListWidgetItem(
                            f"  {payload.get('name') or payload.get('title') or ''}"
                        )
                        item.setData(Qt.ItemDataRole.UserRole, payload)
                        self.file_list.addItem(item)
                    elif payload.get('kind') == 'file':
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
        hymns = book.get('hymns') or []
        if hymns:
            for hymn in hymns:
                payload = hymn_payload(book, hymn)
                item = QListWidgetItem(f"  {payload['name']}")
                item.setData(Qt.ItemDataRole.UserRole, payload)
                self.file_list.addItem(item)
                count += 1
        else:
            for fi in book.get('files', []):
                ext = fi['ext']
                tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
                item = QListWidgetItem(f"  [{tag}]  {fi['name']}")
                item.setData(Qt.ItemDataRole.UserRole, {'kind': 'file', 'path': fi['path']})
                self.file_list.addItem(item)
                count += 1

        if count == 0:
            ni = QListWidgetItem("  ⚠  沒有符合條件的詩歌")
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
            matches = []
            for hymn in book.get('hymns') or []:
                payload = hymn_payload(book, hymn)
                score = max(
                    fuzzy_match_score(q, payload.get('name', '')) or 0,
                    fuzzy_match_score(q, payload.get('title', '')) or 0,
                    fuzzy_match_score(q, payload.get('code', '')) or 0,
                )
                if q in (payload.get('name') or '').lower() or q in (payload.get('title') or '').lower():
                    score = max(score, 80)
                if q.isdigit() and (
                    q.lstrip('0') == (hymn.get('number') or '').lstrip('0')
                    or q in (hymn.get('code') or '').lower()
                ):
                    score = max(score, 100)
                if score > 0:
                    matches.append({**payload, 'score': score})

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
                bi = QListWidgetItem(f"      {m.get('name') or m.get('title') or ''}")
                bi.setData(Qt.ItemDataRole.UserRole, {k: v for k, v in m.items() if k != 'score'})
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
        results = search_hymn_lyrics(self.books, query)
        self._on_keyword_results(results)

    def _on_search_progress(self, msg):
        self._set_status('warn', f'搜索中：{msg}')

    def _show_keyword_placeholder(self):
        self.file_list.clear()
        t = THEMES[self.theme_name]
        hint = "  🔍  輸入關鍵字搜索歌詞（可勾選即時搜索）"
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
            ni = QListWidgetItem(f"  ⚠  在歌詞中找不到「{query}」")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)
            self.file_count_lbl.setText("0 項")
            self._set_status('ok', '搜索完成 · 0 結果')
            self._update_search_action_buttons()
            return

        count = 0
        for entry in results:
            hits = entry.get('hits') or []
            add_section_hdr(f"📖  {entry['book']}  ({len(hits)} 首)")
            for hit in hits:
                snip = hit.get('snippet', '')
                label = f"      {hit.get('name') or hit.get('title') or ''}"
                if snip:
                    label += f"  ·  {snip}"
                bi = QListWidgetItem(label)
                bi.setData(Qt.ItemDataRole.UserRole, hit)
                self.file_list.addItem(bi)
                count += 1

        self.file_count_lbl.setText(f"{count} 處 · {len(results)} 本")
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
        BlackScreenOverlay.hide()
        kind = payload.get('kind')
        opened = False
        hymn_payload_data = payload if kind == 'hymn' and payload.get('lyrics') else payload_from_legacy(self.books, payload)
        if hymn_payload_data:
            try:
                self._last_opened_path = 'lyric'
                preview_mode = bool(self._settings.get('lyric_preview_mode', False))
                if show_overlay and not preview_mode:
                    self._show_open_hymn_popup(
                        hymn_payload_data.get('book', ''),
                        hymn_payload_data.get('name') or hymn_payload_data.get('title', ''),
                    )
                self._lyric_pair.open_hymn(
                    hymn_payload_data,
                    start_blacked=self._settings.get('lyric_start_blacked', True),
                    preview=preview_mode,
                )
                self._schedule_lyric_preview()
                QTimer.singleShot(80, self._sync_lyric_preview)
                QTimer.singleShot(220, self._sync_lyric_preview)
                if show_toast:
                    label = hymn_payload_data.get('name') or hymn_payload_data.get('title') or ''
                    if preview_mode:
                        self._show_toast(f'👁  預覽：{label}（未出街）')
                    else:
                        self._show_toast(f"✅  開啟：{label}")
                opened = True
                payload = hymn_payload_data
            except Exception as e:
                if show_toast:
                    self._show_toast(f"❌  無法開啟詩歌：{e}", 'err')
        elif show_toast:
            self._show_toast('找不到對應歌詞，已不再開啟 Word / PDF', 'warn')
        if opened and record:
            self._record_open(payload=payload)

    def _open_hymn_ref(self, book_ref, num_ref='', record=True, *, show_overlay=True, show_toast=True):
        matches = resolve_open_request(
            self.books, book_ref, num_ref, num_mode=self._hymn_num_mode(),
        )
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
        if self.search_mode in BOOK_UI_MODES:
            is_open = single
            if getattr(self, 'senior_mode', False):
                self.btn_filter.setText('開啟' if is_open else '搜尋')
            else:
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
        if self.search_mode not in BOOK_UI_MODES:
            return
        if self._try_open_single_result():
            return
        self._apply_sidebar_filter()

    def _toggle_voice_input(self):
        if not hasattr(self, '_voice'):
            return
        if not voice_input_available():
            self._show_toast('未安裝語音辨識套件', 'err')
            return
        self._voice.toggle()

    def _on_voice_listening_changed(self, listening: bool):
        if not hasattr(self, 'btn_voice'):
            return
        if listening:
            self.btn_voice.setText('⏹')
            self.btn_voice.setToolTip('停止語音輸入')
            style = getattr(self, '_voice_btn_listen_style', '')
        else:
            self.btn_voice.setText('🎤')
            self.btn_voice.setToolTip('語音輸入（廣東話）— 再按一次可停止')
            style = getattr(self, '_voice_btn_idle_style', '')
        if style:
            self.btn_voice.setStyleSheet(style)

    def _on_voice_status(self, msg: str):
        if msg:
            self._show_toast(msg, 'warn')

    def _on_voice_error(self, msg: str):
        self._show_toast(msg or '語音辨識失敗', 'err')

    def _on_voice_transcript(self, text: str):
        text = (text or '').strip()
        if not text:
            self._show_toast('聽唔到內容', 'warn')
            return
        target = self.inp_num
        focus = QApplication.focusWidget()
        if focus is self.inp_book:
            target = self.inp_book
        target.setText(text)
        target.setFocus()
        self._show_toast(f'語音：{text}', 'ok')
        # 只過濾／更新列表，唔自動開檔（避免單一結果即刻 Open）
        if self.search_mode in BOOK_UI_MODES:
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
        if self.search_mode == 'standard':
            self._apply_standard_filter()
            return
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

    def _apply_standard_filter(self):
        """標準模式：有書冊則搜該書；無書冊時歌名做全局搜索（同簡易版／手機）。"""
        bq_raw = self.inp_book.text().strip()
        bq = bq_raw.lower()
        fq_raw = self.inp_num.text().strip()
        fq = fq_raw.lower()
        t = THEMES[self.theme_name]

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

        if not bq:
            # 顯示全部書冊；歌名 → 全局搜索
            for i in range(self.book_list.count()):
                item = self.book_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole):
                    item.setHidden(False)
            if fq:
                self.file_hdr.setText("全局搜索結果")
                self.hint_lbl.setText("無書冊：以歌名／號做全局搜索")
                self.sel_info.setText(f"  ▸  標準模式 — 全局「{fq_raw}」")
                self._show_global_results(fq_raw)
            else:
                self.file_hdr.setText("檔案 / 書籤" if not getattr(self, 'senior_mode', False) else "詩歌列表")
                self.file_list.clear()
                ni = QListWidgetItem("  輸入書冊號／名，或直接輸入歌名做全局搜索")
                ni.setForeground(QColor(t['muted']))
                self.file_list.addItem(ni)
                self.file_count_lbl.setText("")
                self._update_sel_info(None)
                if not getattr(self, 'senior_mode', False):
                    self.hint_lbl.setText("有書冊則搜該書；無書冊時歌名做全局搜索")
                self._update_search_action_buttons()
            return

        preview = self.current_book if self.current_book in visible else None
        if not preview and visible:
            preview = visible[0]

        if preview:
            self.file_hdr.setText("檔案 / 書籤" if not getattr(self, 'senior_mode', False) else "詩歌列表")
            if not getattr(self, 'senior_mode', False):
                self.hint_lbl.setText("已選書冊 — 右邊列出該書檔案／書籤")
            self._show_book_files(preview, fq)
            self._update_sel_info(preview)
        else:
            self.file_list.clear()
            ni = QListWidgetItem("  找不到符合條件的書冊")
            ni.setForeground(QColor(t['muted']))
            self.file_list.addItem(ni)
            self.file_count_lbl.setText("")
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
    #  SEARCH MODE — 切換標準/書本/全局/關鍵字模式與主題、字體
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
        mode = LEGACY_SEARCH_MODES.get(mode, mode)
        if getattr(self, 'senior_mode', False) and mode != 'standard':
            mode = 'standard'
        if mode not in SEARCH_MODES:
            mode = 'standard'
        if mode != 'keyword':
            self._stop_search_worker()
            self._keyword_debounce.stop()

        self.search_mode = mode
        self.btn_standard.setChecked(mode == 'standard')
        self.btn_book.setChecked(mode == 'book')
        self.btn_global.setChecked(mode == 'global')
        self.btn_keyword.setChecked(mode == 'keyword')
        self.btn_schedule.setChecked(mode == 'schedule')
        self.book_row.setVisible(mode in BOOK_UI_MODES)
        self.schedule_row.setVisible(mode == 'schedule')
        self.global_row.setVisible(mode == 'global')
        self.keyword_row.setVisible(mode == 'keyword')
        self.keyword_index_row.setVisible(False)
        self.dropdown.hide()
        if hasattr(self, 'sched_entry_dropdown'):
            self.sched_entry_dropdown.hide()
        self._apply_schedule_file_layout(mode == 'schedule')

        if mode == 'standard':
            if getattr(self, 'senior_mode', False):
                self.inp_book.setPlaceholderText("輸入書冊號／名稱…（可留空）")
                self.inp_num.setPlaceholderText("詩歌號或歌名（無書冊＝全局搜索）")
                self.hint_lbl.setText("可輸入書冊再搜詩歌；無書冊時歌名會全局搜索")
                self.file_hdr.setText("詩歌列表")
            else:
                self.inp_book.setPlaceholderText("書冊號／名（可留空＝全局搜歌名）")
                self.inp_num.setPlaceholderText("詩歌號／歌名")
                self.file_hdr.setText("檔案 / 書籤")
                self.hint_lbl.setText("有書冊則搜該書；無書冊時歌名做全局搜索")
            self._apply_standard_filter()
        elif mode == 'book':
            self.inp_book.setPlaceholderText("請按 Ctrl+Enter 開始輸入書冊名稱...")
            self.inp_num.setPlaceholderText("詩歌號/名")
            self.file_hdr.setText("檔案 / 書籤")
            self.hint_lbl.setText("選一個左邊書冊，右邊會列出該書的所有檔案")
            if self.current_book:
                self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
                self._update_sel_info(self.current_book)
            else:
                self._apply_sidebar_filter()
        elif mode == 'schedule':
            self.hint_lbl.setText("雙擊搜尋結果加入播放清單；雙擊播放清單開啟")
            self.sel_info.setText("  ▸  排程模式 — 管理今日詩歌順序")
            self._rebuild_sched_book_combo()
            self._refresh_schedule_search()
            self._refresh_schedule_list()
        elif mode == 'global':
            self.file_hdr.setText("全局搜索結果")
            self.hint_lbl.setText("輸入歌名或編號，會在所有書冊中搜索")
            self.sel_info.setText("  ▸  全局模式 — 跨書冊搜索")
            self._show_global_results(self.inp_global.text())
        else:
            self.file_hdr.setText("關鍵字搜索結果")
            self.hint_lbl.setText("輸入關鍵字搜索歌詞；可勾選「即時搜索」邊輸入邊找")
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
        elif self.search_mode == 'standard':
            self._apply_standard_filter()
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

    def _nudge_font(self, delta):
        try:
            cur = int(self.font_size)
        except (TypeError, ValueError):
            cur = 13
        self._set_font_size(cur + int(delta))

    def _sync_font_size_labels(self):
        text = f"{self.font_size}px"
        if hasattr(self, 'fs_lbl'):
            self.fs_lbl.setText(text)
        lo, hi = FONT_SIZE_MIN, FONT_SIZE_MAX
        at_min = self.font_size <= lo
        at_max = self.font_size >= hi
        if hasattr(self, 'btn_font_minus_settings'):
            self.btn_font_minus_settings.setEnabled(not at_min)
        if hasattr(self, 'btn_font_plus_settings'):
            self.btn_font_plus_settings.setEnabled(not at_max)

    def _set_font_size(self, val, persist=True):
        try:
            val = int(val)
        except (TypeError, ValueError):
            val = 13
        val = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, val))
        self.font_size = val
        self._sync_font_size_labels()
        self._apply_theme()
        if self.search_mode == 'global':
            self._show_global_results(self.inp_global.text())
        elif self.search_mode == 'keyword':
            pass
        elif self.search_mode == 'schedule':
            self._refresh_schedule_list()
            self._refresh_schedule_search()
        elif self.search_mode == 'standard':
            self._apply_standard_filter()
        elif self.current_book:
            self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
        if persist and not self._loading_settings:
            self._persist_settings(font_size=val)

    def _on_font_changed(self, val):
        self._set_font_size(val)

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

    def _remote_fixed_base_url(self):
        if not getattr(self, 'enable_tunnel_fixed', False) or not self._remote.running:
            return ''
        tun = getattr(self, '_tunnel_fixed', None)
        if tun and tun.is_ready and tun.public_url:
            return normalize_tunnel_url(tun.public_url).rstrip('/')
        if self.named_tunnel_url.strip():
            return normalize_tunnel_url(self.named_tunnel_url).rstrip('/')
        return ''

    def _remote_quick_base_url(self):
        if not getattr(self, 'enable_tunnel_quick', False) or not self._remote.running:
            return ''
        tun = getattr(self, '_tunnel_quick', None)
        if tun and tun.is_ready and tun.public_url:
            return normalize_tunnel_url(tun.public_url).rstrip('/')
        return ''

    def _remote_wan_base_url(self):
        """相容：優先固定，其次隨機。"""
        return self._remote_fixed_base_url() or self._remote_quick_base_url()

    def _remote_lan_viewer_url(self):
        """內網會眾入口（同 WiFi）。"""
        if not getattr(self._remote, 'running', False):
            return ''
        base = self._remote_local_base_url()
        return tunnel_viewer_url(base) if base else ''

    def _remote_lan_admin_url(self):
        """內網操作員遙控頁（/admin）。"""
        base = self._remote_local_base_url()
        return f"{base}/admin" if base else ''

    def _remote_fixed_viewer_url(self):
        base = self._remote_fixed_base_url()
        return tunnel_viewer_url(base) if base else ''

    def _remote_fixed_admin_url(self):
        base = self._remote_fixed_base_url()
        return f"{base.rstrip('/')}/admin" if base else ''

    def _remote_quick_viewer_url(self):
        base = self._remote_quick_base_url()
        return tunnel_viewer_url(base) if base else ''

    def _remote_quick_admin_url(self):
        base = self._remote_quick_base_url()
        if not base:
            return ''
        return f"{base.rstrip('/')}/admin"

    def _remote_wan_viewer_url(self):
        """相容：優先固定，其次隨機。"""
        return self._remote_fixed_viewer_url() or self._remote_quick_viewer_url()

    def _remote_share_url(self):
        """相容舊邏輯：優先外網，否則內網。"""
        wan = self._remote_wan_viewer_url()
        if wan:
            return wan
        return self._remote_lan_viewer_url()

    def _remote_public_base_url(self):
        """ Backward-compatible alias for QR / mixin."""
        return self._remote_share_url()

    def _remote_viewer_url(self):
        return self._remote_fixed_viewer_url() or '—'

    def _remote_admin_url(self):
        return self._remote_fixed_admin_url() or '—'

    def _update_remote_url_labels(self):
        if self._remote.running:
            self.local_url_lbl.setText(self._remote_lan_viewer_url() or '—')
            if hasattr(self, 'local_admin_url_lbl'):
                self.local_admin_url_lbl.setText(self._remote_lan_admin_url() or '—')
            if getattr(self, 'enable_tunnel_fixed', False):
                self.viewer_url_lbl.setText(self._remote_viewer_url())
                self.remote_url_lbl.setText(self._remote_admin_url())
            else:
                self.viewer_url_lbl.setText('—')
                self.remote_url_lbl.setText('—')
            if hasattr(self, 'quick_url_lbl'):
                if getattr(self, 'enable_tunnel_quick', False):
                    self.quick_url_lbl.setText(self._remote_quick_viewer_url() or '—')
                else:
                    self.quick_url_lbl.setText('—')
            if hasattr(self, 'quick_admin_url_lbl'):
                if getattr(self, 'enable_tunnel_quick', False):
                    self.quick_admin_url_lbl.setText(self._remote_quick_admin_url() or '—')
                else:
                    self.quick_admin_url_lbl.setText('—')
        else:
            self.local_url_lbl.setText('—')
            if hasattr(self, 'local_admin_url_lbl'):
                self.local_admin_url_lbl.setText('—')
            self.remote_url_lbl.setText('—')
            self.viewer_url_lbl.setText('—')
            if hasattr(self, 'quick_url_lbl'):
                self.quick_url_lbl.setText('—')
            if hasattr(self, 'quick_admin_url_lbl'):
                self.quick_admin_url_lbl.setText('—')
        for lbl_name, url_fn in (
            ('remote_qr_lan_link', self._remote_lan_viewer_url),
            ('remote_qr_lan_admin_link', self._remote_lan_admin_url),
            ('remote_qr_fixed_link', self._remote_fixed_viewer_url),
            ('remote_qr_fixed_admin_link', self._remote_fixed_admin_url),
            ('remote_qr_quick_link', self._remote_quick_viewer_url),
            ('remote_qr_quick_admin_link', self._remote_quick_admin_url),
            ('remote_qr_wan_link', self._remote_wan_viewer_url),
        ):
            lbl = getattr(self, lbl_name, None)
            if lbl is None:
                continue
            url = url_fn() if self._remote.running else ''
            lbl.setText(url or '—')
            lbl.setToolTip(url or '')

    def _set_tunnel_status(self, role, text):
        self._tunnel_status[role] = text
        if role == 'fixed':
            lbl = getattr(self, 'tunnel_status_fixed_lbl', None)
            prefix = '固定'
        else:
            lbl = getattr(self, 'tunnel_status_quick_lbl', None)
            prefix = '隨機'
        if lbl is not None:
            lbl.setText(f'{prefix}：{text}')

    def _start_cloudflare_tunnel(self):
        """Start whichever of fixed / quick are enabled."""
        if not self._remote.running:
            return
        if self.enable_tunnel_fixed:
            self._start_fixed_tunnel()
        else:
            self._stop_fixed_tunnel()
        if self.enable_tunnel_quick:
            self._start_quick_tunnel()
        else:
            self._stop_quick_tunnel()

    def _start_fixed_tunnel(self):
        if not self.enable_tunnel_fixed or not self._remote.running:
            return
        if not self.named_tunnel_url.strip():
            self._set_tunnel_status('fixed', '請設定公開網址')
            self._show_toast('啟用固定外網時需填寫公開網址', 'warn')
            return
        self._stop_fixed_tunnel()
        self._set_tunnel_status('fixed', '正在連線…')
        port = self._remote.state.port
        self._tunnel_fixed = CloudflareTunnel(
            local_port=port,
            cloudflared_path=self.cloudflared_path,
            tunnel_mode='service',
            cloudflared_token=self.cloudflared_token,
            named_tunnel_url=self.named_tunnel_url,
            use_https=False,
            on_url=lambda url: self.tunnel_url_signal.emit('fixed', url),
            on_ready=lambda url: self.tunnel_ready_signal.emit('fixed', url),
            on_error=lambda msg: self.tunnel_error_signal.emit('fixed', msg),
            on_expired=lambda msg: self.tunnel_expired_signal.emit('fixed', msg),
        )
        self._tunnel = self._tunnel_fixed
        threading.Thread(target=self._tunnel_fixed.start, daemon=True).start()

    def _start_quick_tunnel(self):
        if not self.enable_tunnel_quick or not self._remote.running:
            return
        self._stop_quick_tunnel()
        self._set_tunnel_status('quick', '正在啟動通道…')
        port = self._remote.state.port
        self._tunnel_quick = CloudflareTunnel(
            local_port=port,
            cloudflared_path=self.cloudflared_path,
            tunnel_mode='quick',
            cloudflared_token=self.cloudflared_token,
            named_tunnel_url='',
            use_https=False,
            on_url=lambda url: self.tunnel_url_signal.emit('quick', url),
            on_ready=lambda url: self.tunnel_ready_signal.emit('quick', url),
            on_error=lambda msg: self.tunnel_error_signal.emit('quick', msg),
            on_expired=lambda msg: self.tunnel_expired_signal.emit('quick', msg),
        )
        self._tunnel = self._tunnel_quick
        threading.Thread(target=self._tunnel_quick.start, daemon=True).start()

    def _stop_fixed_tunnel(self):
        was = self._tunnel_fixed
        if self._tunnel_fixed:
            self._tunnel_fixed.stop()
            self._tunnel_fixed = None
        if self._tunnel is was:
            self._tunnel = self._tunnel_quick
        self._set_tunnel_status('fixed', '—')

    def _stop_quick_tunnel(self):
        was = self._tunnel_quick
        if self._tunnel_quick:
            self._tunnel_quick.stop()
            self._tunnel_quick = None
        if self._tunnel is was:
            self._tunnel = self._tunnel_fixed
        self._set_tunnel_status('quick', '—')

    def _stop_cloudflare_tunnel(self):
        self._stop_fixed_tunnel()
        self._stop_quick_tunnel()
        self._tunnel = None

    def _persist_tunnel_enables(self):
        self.enable_tunnel = self.enable_tunnel_fixed or self.enable_tunnel_quick
        if self.enable_tunnel_fixed:
            self.tunnel_mode = 'service'
        elif self.enable_tunnel_quick:
            self.tunnel_mode = 'quick'
        self._persist_settings(
            enable_tunnel=self.enable_tunnel,
            enable_tunnel_fixed=self.enable_tunnel_fixed,
            enable_tunnel_quick=self.enable_tunnel_quick,
            tunnel_mode=self.tunnel_mode,
        )

    def _on_tunnel_fixed_toggled(self, checked):
        self.enable_tunnel_fixed = bool(checked)
        self._persist_tunnel_enables()
        self._sync_tunnel_options_visibility()
        if checked and self._remote.running:
            self._start_fixed_tunnel()
        elif not checked:
            self._stop_fixed_tunnel()
            self._update_remote_url_labels()
            self._update_qr_code()

    def _on_tunnel_quick_toggled(self, checked):
        self.enable_tunnel_quick = bool(checked)
        self._persist_tunnel_enables()
        self._sync_tunnel_options_visibility()
        if checked and self._remote.running:
            self._start_quick_tunnel()
        elif not checked:
            self._stop_quick_tunnel()
            self._update_remote_url_labels()
            self._update_qr_code()

    def _on_tunnel_toggled(self, checked):
        # Compat: master toggle maps to fixed when used.
        self._on_tunnel_fixed_toggled(checked)

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

    def _sync_tunnel_options_visibility(self, update_qr=True):
        if not hasattr(self, 'chk_enable_tunnel_fixed'):
            return
        fixed_on = self.chk_enable_tunnel_fixed.isChecked()
        quick_on = self.chk_enable_tunnel_quick.isChecked()
        service_installed = is_cloudflared_service_installed()
        show_install = fixed_on and not service_installed

        if hasattr(self, '_wan_service_box'):
            self._wan_service_box.setVisible(fixed_on)
        for w in (
            self.tunnel_help_lbl,
            self.tunnel_status_fixed_lbl,
            self.lbl_public_tunnel_url,
            self.inp_named_tunnel_url,
            self.lbl_viewer_url, self.viewer_url_lbl,
            self.lbl_remote_url, self.remote_url_lbl,
            self.chk_show_qr_fixed,
            self.remote_qr_fixed_pair,
        ):
            w.setVisible(fixed_on)
        self.inp_cloudflared_token.setVisible(show_install)
        self.btn_install_cloudflared.setVisible(show_install)
        for w in (
            self.lbl_tunnel_cmd_hint,
            self._tunnel_install_row,
            self._tunnel_uninstall_row,
            self._tunnel_restart_row,
        ):
            w.setVisible(fixed_on)
        if fixed_on:
            self._update_tunnel_debug_commands()
            self._refresh_tunnel_service_ui(service_installed=service_installed)

        for w in (
            self.tunnel_status_quick_lbl,
            self.lbl_quick_url, self.quick_url_lbl,
            getattr(self, 'lbl_quick_admin_url', None),
            getattr(self, 'quick_admin_url_lbl', None),
            self.chk_show_qr_quick,
            self.remote_qr_quick_pair,
        ):
            if w is not None:
                w.setVisible(quick_on)

        self._update_remote_url_labels()
        if update_qr:
            self._update_qr_code(sync_desktop=False)

    def _refresh_tunnel_service_ui(self, service_installed=None):
        if not hasattr(self, 'tunnel_help_lbl'):
            return
        port = self.spin_remote_port.value() if hasattr(self, 'spin_remote_port') else DEFAULT_PORT
        if service_installed is None:
            service_installed = is_cloudflared_service_installed()
        dns_note = (
            'DNS：Cloudflare 須有 live → <tunnel-id>.cfargotunnel.com 的 CNAME，'
            '否則公開網址無法解析。'
        )
        if service_installed:
            wan_off_note = (
                ' 未勾選「固定網址」時 App 不會主動驗證公開網址；'
                '若公開網址仍可連，代表 cloudflared Windows 服務仍在背景運行。'
            )
            self.tunnel_help_lbl.setText(
                f'{shared_tunnel_setup_hint(port)} {dns_note}{wan_off_note}'
            )
            self.btn_install_cloudflared.setEnabled(False)
            self.btn_install_cloudflared.setToolTip('本機已有 cloudflared 服務。')
        else:
            self.btn_install_cloudflared.setEnabled(True)
            self.btn_install_cloudflared.setToolTip('僅在未有 cloudflared 服務時使用。')
            self.tunnel_help_lbl.setText(
                f'{shared_tunnel_setup_hint(port)} {dns_note}'
            )
        self._update_tunnel_debug_commands()

    def _on_tunnel_mode_changed(self, _index):
        # Legacy no-op: modes are independent checkboxes now.
        return

    def _on_named_tunnel_url_changed(self, text):
        self.named_tunnel_url = text.strip()
        self.remote_path_prefix = tunnel_url_path_prefix(self.named_tunnel_url)
        self._persist_settings(
            named_tunnel_url=self.named_tunnel_url,
            remote_path_prefix=self.remote_path_prefix,
        )
        if self.enable_tunnel_fixed and self._remote.running:
            self._start_fixed_tunnel()

    def _on_cloudflared_token_changed(self, text):
        self.cloudflared_token = text
        self._persist_settings(cloudflared_token=text)
        self._update_tunnel_debug_commands()

    def _install_cloudflared_service(self):
        if is_cloudflared_service_installed():
            self._show_toast('本機已有 cloudflared 服務', 'warn')
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
            if self.enable_tunnel_fixed and self._remote.running:
                self._start_fixed_tunnel()
        else:
            self._show_toast(f'cloudflared 服務安裝失敗：{message}', 'err')

    def _tunnel_for_role(self, role):
        if role == 'fixed':
            return getattr(self, '_tunnel_fixed', None)
        if role == 'quick':
            return getattr(self, '_tunnel_quick', None)
        return getattr(self, '_tunnel', None)

    def _on_tunnel_url(self, role, url):
        self._set_tunnel_status(role, f'{url}\n（等待 Cloudflare...）')

    def _on_tunnel_ready(self, role, url):
        tun = self._tunnel_for_role(role)
        if tun and not tun.local_dns_ok:
            self._set_tunnel_status(
                role,
                f'{url}\n✓ 通道已啟用（手機／外網可用）\n'
                '⚠ 本機 DNS 解唔到此網址 — 唔影響掃碼；'
                '若要喺呢部電腦瀏覽器開，可改 DNS 1.1.1.1 或用手機流量',
            )
        else:
            self._set_tunnel_status(role, f'{url}\n✓ 通道已啟用')
        self._update_remote_url_labels()
        self._update_qr_code()

    def _on_tunnel_expired(self, role, message):
        self._set_tunnel_status(role, f'⚠ {message}')
        self._update_remote_url_labels()
        self._update_qr_code()

    def _on_tunnel_error(self, role, message):
        self._set_tunnel_status(role, f'{message}\n（詳情見 exe 旁 tunnel.log）')

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
        if hasattr(self, 'combo_remote_policy'):
            self.combo_remote_policy.setEnabled(enabled)
        if not enabled:
            self._remote.state.set_accepting(False)
            self._sync_remote_accept_checkboxes(checked=False)

    def _sync_remote_accept_checkboxes(self, checked):
        self.chk_remote_accept.blockSignals(True)
        self.chk_remote_accept.setChecked(checked)
        self.chk_remote_accept.blockSignals(False)

    def _on_remote_accept_toggled(self, checked):
        self._remote.state.set_accepting(checked)
        self._update_remote_status_label()
        self._persist_settings(remote_accept=checked)

    def _on_remote_accept_settings_toggled(self, checked):
        self._sync_remote_accept_checkboxes(checked)
        self._on_remote_accept_toggled(checked)

    def _on_remote_policy_changed(self, _idx):
        policy = self.combo_remote_policy.currentData()
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

    def _on_remote_prefs(self, payload):
        if not isinstance(payload, dict):
            return
        if 'overlay_auto_unblack' in payload:
            checked = bool(payload['overlay_auto_unblack'])
            if hasattr(self, 'chk_overlay_auto_unblack'):
                self.chk_overlay_auto_unblack.blockSignals(True)
                self.chk_overlay_auto_unblack.setChecked(checked)
                self.chk_overlay_auto_unblack.blockSignals(False)
            self._persist_settings(overlay_auto_unblack=checked)

    def _on_remote_lyric_control(self, payload):
        pair = getattr(self, '_lyric_pair', None)
        if pair is None or not isinstance(payload, dict):
            return
        action = str(payload.get('action') or '').strip().lower()
        value = payload.get('value')
        session = pair.session
        if action == 'black':
            if value is None or value == '':
                session.toggle_blacked()
            else:
                session.set_blacked(str(value).strip().lower() in ('1', 'true', 'yes', 'on'))
        elif action == 'fit':
            pair.fit_display()
        elif action == 'restore':
            self._restore_lyric_best_defaults()
        elif action == 'save_best':
            self._save_lyric_best_defaults()
        elif action == 'fullscreen':
            pair.toggle_display_fullscreen()
        elif action == 'align':
            session.set_align(value)
            if not session.user_zoomed:
                pair.fit_display()
        elif action == 'spacing':
            session.set_letter_spacing(value)
            if not session.user_zoomed:
                pair.fit_display()
        elif action == 'scroll':
            try:
                session.set_scroll_ratio(value)
            except (TypeError, ValueError):
                pass
        elif action == 'nudge':
            try:
                session.nudge_scroll(value)
            except (TypeError, ValueError):
                pass
        elif action == 'zoom':
            try:
                factor = float(value)
            except (TypeError, ValueError):
                factor = ZOOM_STEP
            if factor <= 0:
                factor = ZOOM_STEP
            session.zoom_by(factor, user=True)
        self._schedule_lyric_preview()

    def _on_remote_display_action(self, action):
        action = str(action or '').strip().lower()
        try:
            if hasattr(self, '_open_flow') and self._open_flow:
                self._open_flow.cancel()
        except Exception:
            pass
        watch = getattr(BookNameOverlay, '_watch_path', '') or ''
        path = watch or getattr(self, '_last_opened_path', '') or ''
        needle = ''
        if path:
            needle = os.path.splitext(os.path.basename(path))[0]
        if not needle:
            label = str(self._settings.get('last_opened') or '')
            if ' / ' in label:
                needle = label.rsplit(' / ', 1)[-1].strip()
            else:
                needle = label.strip()
        BookNameOverlay.hide_overlay()
        # Must close viewer BEFORE black overlay steals focus
        BlackScreenOverlay.hide()
        ok = False
        try:
            if getattr(self, '_lyric_pair', None) is not None and self._lyric_pair.is_open():
                self._lyric_pair.close()
                self._schedule_lyric_preview()
                ok = True
        except Exception as exc:
            print(f'[WARN] close lyric windows: {exc}')
            ok = False
        if action == 'black':
            BlackScreenOverlay.show()
            msg = '📱  已結束並全黑' if ok else '📱  全黑已開'
            self._show_toast(msg, 'warn' if ok else 'err')
        else:
            if ok:
                self._show_toast('📱  已結束目前詩歌', 'ok')
            else:
                self._show_toast('📱  未有開啟中的歌詞視窗', 'warn')

    def _show_phone_control_wizard(self):
        dlg = PhoneControlWizard(self, theme=self.theme_name, parent=self)
        dlg.exec()

    def _ensure_phone_control_tunnel(self):
        """One-tap: enable API + accept + quick tunnel for phone admin."""
        if hasattr(self, 'chk_remote_api') and not self.chk_remote_api.isChecked():
            self.chk_remote_api.setChecked(True)
        if not getattr(self, '_remote', None) or not self._remote.running:
            self._show_toast('API 尚未就緒，請稍候再試', 'warn')
            return
        if hasattr(self, 'chk_remote_accept') and not self.chk_remote_accept.isChecked():
            self.chk_remote_accept.setChecked(True)
        if hasattr(self, 'chk_enable_tunnel_quick'):
            if not self.chk_enable_tunnel_quick.isChecked():
                self.chk_enable_tunnel_quick.setChecked(True)
            elif not self._remote_quick_base_url():
                self._start_quick_tunnel()
        if hasattr(self, 'chk_show_qr_quick') and not self.chk_show_qr_quick.isChecked():
            self.chk_show_qr_quick.setChecked(True)
        self._show_toast('正在啟動外網隨機…', 'ok')

    def _sync_remote_query(self, book_ref, num_ref):
        matched = resolve_books(self.books, book_ref)
        if not matched:
            return
        book = matched[0]
        self._set_search_mode('standard')
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
        BlackScreenOverlay.hide()
        DesktopQrOverlay.hide()
        if getattr(self, '_lyric_pair', None) is not None:
            self._lyric_pair.close()
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
            if event.type() == QEvent.Type.FocusIn and self.search_mode in BOOK_UI_MODES:
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