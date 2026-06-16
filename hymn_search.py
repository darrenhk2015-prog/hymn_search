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
    預設掃描 HYMN_FOLDER 指向的資料夾（見下方 CONFIG）。
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
    QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton,
    QLineEdit, QLabel, QListWidget, QListWidgetItem,
    QComboBox, QSlider, QFrame, QScrollArea, QCheckBox, QSpinBox,
    QGroupBox,
)
from PyQt6.QtCore import (
    Qt, QTimer, QPoint, pyqtSignal, QEvent, QThread, QRectF, QRect,
    QPropertyAnimation, QEasingCurve,
)
from PyQt6.QtGui import QColor, QPalette, QFont, QPainter, QPainterPath, QIcon

# ══════════════════════════════════════════════════════════════
#  CONFIG — 全域設定與資源路徑
# ══════════════════════════════════════════════════════════════
# HYMN_FOLDER：詩歌資料根目錄；啟動時會遞迴掃描此路徑下的書冊與檔案。
# app_resource_path / load_app_icon：開發模式與 PyInstaller 打包後均能正確載入 assets。
HYMN_FOLDER = os.path.expanduser("~/Desktop/神家詩歌集new")


def app_resource_path(*parts):
    """解析資源檔路徑（開發環境用腳本目錄，打包後用 _MEIPASS 臨時目錄）。"""
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
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
        self.setFrameShape(QFrame.Shape.HLine)
        self.setFixedHeight(1)
        self.set_theme(theme)

    def set_theme(self, theme):
        self.setStyleSheet(f"background: {THEMES[theme]['border']}; border: none;")


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


def find_page_in_toc(toc_list, hymn_num):
    n, n3 = str(hymn_num), str(hymn_num).zfill(3)
    for _lvl, title, page in toc_list:
        t = title.strip()
        if t in (n, n3): return page
        if t.startswith((n+'.', n+' ', n+'。', n+'、', n3+'.', n3+' ')): return page
        if re.match(rf'^第\s*{n}\s*首', t): return page
    for _lvl, title, page in toc_list:
        nums = re.findall(r'\d+', title)
        if nums and int(nums[0]) == hymn_num: return page
    return None


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
        print(f"[WARN] document watch: {e}")
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
    """置頂浮動清單；依 inp_book 輸入過濾並顯示可選書冊。"""
    book_selected = pyqtSignal(dict)

    def __init__(self, theme='dark'):
        super().__init__(None)
        self._theme = theme
        self._books = []
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
        if 0 <= idx < len(self._books):
            self.book_selected.emit(self._books[idx])
        self.hide()

    def refresh(self, books, anchor_widget):
        self._books = books
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


# ══════════════════════════════════════════════════════════════
#  MAIN WINDOW — 主視窗（UI 建構、搜尋邏輯、事件處理）
# ══════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    """
    應用程式主視窗。

    職責：
    - 左側固定寬度書冊列表；右側搜尋列 + 檔案/結果列表
    - 三種搜尋模式切換、設定面板、主題與字體
    - 協調開檔流程（提示 → 開啟 → Word 快捷鍵 → 可選同步畫面）
    """
    display_duplicate_requested = pyqtSignal()  # 從背景執行緒安全觸發顯示模式切換

    def __init__(self):
        super().__init__()
        self.books        = []
        self.current_book = None
        self.search_mode  = 'book'
        self.theme_name   = 'dark'
        self.font_size    = 13
        self._search_worker = None
        self._index_worker = None
        self.content_index = None
        self._app         = QApplication.instance()

        self.setWindowTitle("詩歌冊搜索系統")
        icon = load_app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setMinimumSize(900, 580)
        self.resize(1120, 730)

        self.dropdown = DropdownList(self.theme_name)
        self.dropdown.book_selected.connect(self._select_book)
        self.display_duplicate_requested.connect(self._switch_duplicate_with_overlay_fix)

        self._build_ui()
        self._apply_theme()
        self._scan_all()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_settings_scroll_height()

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

        body_widget = QWidget()
        body_lay = QHBoxLayout(body_widget)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(0)
        body_lay.addWidget(self._build_sidebar())
        body_lay.addWidget(self._build_content(), 1)
        self.body = body_widget
        root.addWidget(body_widget, 1)

        root.addWidget(self._build_toast())

    # ─────────────────────────────────────────────────────────────
    #  HEADER — 頂部標題列（程式名、資料夾路徑、狀態）
    # ─────────────────────────────────────────────────────────────
    def _build_header(self):
        self.header_frame = QFrame()
        self.header_frame.setFixedHeight(54)

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

        self.hdr_ver = QLabel("Phase 1")
        h.addWidget(self.hdr_ver)

        self.hdr_vsep = QFrame()
        self.hdr_vsep.setFrameShape(QFrame.Shape.VLine)
        self.hdr_vsep.setFixedHeight(20)
        h.addWidget(self.hdr_vsep)

        self.hdr_path = QLabel(f"  {HYMN_FOLDER}")
        h.addWidget(self.hdr_path)
        h.addStretch()

        self.status_lbl = QLabel("◌  掃描中...")
        h.addWidget(self.status_lbl)

        return self.header_frame

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
        hdr = QWidget()
        hdr.setFixedHeight(46)
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(14, 0, 14, 0)
        self.sb_title = QLabel("📁  詩歌冊")
        self.sb_title.setStyleSheet("font-size: 13px; font-weight: bold; background: transparent;")
        hl.addWidget(self.sb_title)
        hl.addStretch()
        self.book_count_lbl = QLabel("0 本")
        hl.addWidget(self.book_count_lbl)
        lay.addWidget(hdr)

        self.sb_sep_top = HSep()
        lay.addWidget(self.sb_sep_top)

        # Book list
        self.book_list = QListWidget()
        self.book_list.itemClicked.connect(self._sidebar_click)
        lay.addWidget(self.book_list, 1)

        self.sb_sep_bot = HSep()
        lay.addWidget(self.sb_sep_bot)

        # Legend
        leg_widget = QWidget()
        self.legend_widget = leg_widget
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
        self.btn_book    = QPushButton("📖  書本模式")
        self.btn_global  = QPushButton("🌐  全局模式")
        self.btn_keyword = QPushButton("🔤  關鍵字模式")
        for btn in [self.btn_book, self.btn_global, self.btn_keyword]:
            btn.setCheckable(True)
            btn.setFixedHeight(30)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_book.setChecked(True)
        self.btn_book.clicked.connect(lambda: self._set_search_mode('book'))
        self.btn_global.clicked.connect(lambda: self._set_search_mode('global'))
        self.btn_keyword.clicked.connect(lambda: self._set_search_mode('keyword'))
        mode_row.addWidget(self.btn_book)
        mode_row.addWidget(self.btn_global)
        mode_row.addWidget(self.btn_keyword)
        mode_row.addStretch()

        self.btn_settings = QPushButton("⚙  設定")
        self.btn_settings.setCheckable(True)
        self.btn_settings.setFixedHeight(30)
        self.btn_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_settings.toggled.connect(self._on_settings_toggled)
        mode_row.addWidget(self.btn_settings)
        lay.addLayout(mode_row)

        # ── 書本模式：過濾書冊 + 檔案關鍵字 + Filter/Open 按鈕 ──
        self.book_row = QWidget()
        br = QHBoxLayout(self.book_row)
        br.setContentsMargins(0, 0, 0, 0)
        br.setSpacing(8)

        self.inp_book = QLineEdit()
        self.inp_book.setFixedHeight(40)
        self.inp_book.setPlaceholderText("🔍  過濾書冊  e.g. 讚美詩")
        self.inp_book.textChanged.connect(self._on_book_text_changed)
        self.inp_book.returnPressed.connect(self._on_book_filter_action)
        self.inp_book.installEventFilter(self)

        self.inp_num = QLineEdit()
        self.inp_num.setFixedHeight(40)
        self.inp_num.setFixedWidth(150)
        self.inp_num.setPlaceholderText("檔案關鍵字")
        self.inp_num.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.inp_num.textChanged.connect(self._apply_sidebar_filter)
        self.inp_num.returnPressed.connect(self._on_book_filter_action)

        self.btn_filter = QPushButton("Filter")
        self.btn_filter.setFixedHeight(40)
        self.btn_filter.setFixedWidth(84)
        self.btn_filter.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_filter.clicked.connect(self._on_book_filter_action)

        br.addWidget(self.inp_book, 1)
        br.addWidget(self.inp_num)
        br.addWidget(self.btn_filter)
        lay.addWidget(self.book_row)

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
        self.chk_instant.setChecked(True)
        self.chk_instant.setCursor(Qt.CursorShape.PointingHandCursor)

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

        self.grp_open = QGroupBox("開啟 Word / PDF 時")
        open_grid = QGridLayout(self.grp_open)
        open_grid.setContentsMargins(*grid_margins)
        open_grid.setHorizontalSpacing(14)
        open_grid.setVerticalSpacing(12)
        open_grid.setColumnStretch(1, 1)
        open_grid.setColumnMinimumWidth(0, label_min_w)

        self.chk_open_overlay = QCheckBox("顯示置頂書名提示")
        self.chk_open_overlay.setChecked(True)
        self.chk_open_overlay.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_open_overlay.toggled.connect(self._on_open_overlay_toggled)

        self.lbl_overlay_secs = QLabel("提示顯示")
        self.lbl_overlay_secs.setMinimumWidth(label_min_w)
        self.lbl_overlay_secs.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_overlay_duration = QComboBox()
        for sec in range(1, MAX_OVERLAY_SECS + 1):
            self.combo_overlay_duration.addItem(f"{sec} 秒", sec)
        self.combo_overlay_duration.addItem("一直顯示", OVERLAY_DURATION_ALWAYS)
        self.combo_overlay_duration.setCurrentIndex(DEFAULT_OVERLAY_SECS - 1)
        self.combo_overlay_duration.setMinimumWidth(108)
        self.combo_overlay_duration.setCursor(Qt.CursorShape.PointingHandCursor)

        self.lbl_overlay_pos = QLabel("提示位置")
        self.lbl_overlay_pos.setMinimumWidth(label_min_w)
        self.lbl_overlay_pos.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.combo_overlay_pos = QComboBox()
        self.combo_overlay_pos.addItem("右下角", OVERLAY_MODE_CORNER)
        self.combo_overlay_pos.addItem("置中 → 右下角", OVERLAY_MODE_CENTER)
        self.combo_overlay_pos.setMinimumWidth(140)
        self.combo_overlay_pos.setCursor(Qt.CursorShape.PointingHandCursor)

        self.chk_display_duplicate = QCheckBox("切換同步畫面")
        self.chk_display_duplicate.setToolTip("開啟 Word / PDF 後，將延伸桌面改為同步（複製）畫面")
        self.chk_display_duplicate.setChecked(False)
        self.chk_display_duplicate.setCursor(Qt.CursorShape.PointingHandCursor)
        if sys.platform != 'win32':
            self.chk_display_duplicate.setEnabled(False)
            self.chk_display_duplicate.setToolTip("僅支援 Windows")

        open_grid.addWidget(self.chk_open_overlay, 0, 0, 1, 2)
        open_grid.addWidget(self.lbl_overlay_secs, 1, 0)
        open_grid.addWidget(self.combo_overlay_duration, 1, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.lbl_overlay_pos, 2, 0)
        open_grid.addWidget(self.combo_overlay_pos, 2, 1, Qt.AlignmentFlag.AlignLeft)
        open_grid.addWidget(self.chk_display_duplicate, 3, 0, 1, 2)

        settings_row = QHBoxLayout()
        settings_row.setSpacing(12)
        settings_row.addWidget(self.grp_appearance, 1)
        settings_row.addWidget(self.grp_open, 1)
        settings_lay.addLayout(settings_row)

        self.settings_scroll = QScrollArea()
        self.settings_scroll.setWidgetResizable(True)
        self.settings_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.settings_scroll.setVisible(False)
        self.settings_scroll.setWidget(self.settings_frame)
        lay.addWidget(self.settings_scroll)

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
        self.file_hdr = QLabel("檔案 / 書籤")
        hdr_row.addWidget(self.file_hdr)
        hdr_row.addStretch()
        self.file_count_lbl = QLabel("")
        hdr_row.addWidget(self.file_count_lbl)
        lay.addLayout(hdr_row)

        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self._file_clicked)
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

    def _on_settings_toggled(self, checked):
        self.settings_scroll.setVisible(checked)
        self.btn_settings.setText("⚙  收起設定" if checked else "⚙  設定")
        if checked:
            QTimer.singleShot(0, self._refresh_settings_layout)
        else:
            self.settings_scroll.setMaximumHeight(0)

    def _refresh_settings_layout(self):
        for w in (self.grp_appearance, self.grp_open, self.settings_frame, self.search_frame):
            w.updateGeometry()
            w.adjustSize()
        self._update_settings_scroll_height()

    def _update_settings_scroll_height(self):
        if not self.settings_scroll.isVisible():
            return
        self.settings_frame.adjustSize()
        content_h = self.settings_frame.sizeHint().height() + 6
        reserved = 430
        avail = max(140, self.height() - reserved)
        self.settings_scroll.setMaximumHeight(min(content_h, avail))

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

    def _open_docx_with_flow(self, path, item_name=''):
        self._show_open_book_popup(path, item_name or os.path.basename(path))
        open_docx(path, after_focus_keys=self._maybe_duplicate_displays)

    def _open_pdf_with_flow(self, path, item_name='', page=None):
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
        self.hdr_ver.setStyleSheet(
            f"font-size: 11px; color: {t['muted']}; "
            f"padding-top: 5px; padding-right: 14px; background: transparent;"
        )
        self.hdr_vsep.setStyleSheet(f"background: {t['border']}; border: none;")
        self.hdr_path.setStyleSheet(
            f"font-size: 11px; color: {t['text2']}; padding-left: 10px; background: transparent;"
        )

        # ── Sidebar ──────────────────────────────────────────────
        self.sidebar_frame.setStyleSheet(
            f"QFrame {{ background: {t['panel']}; border-right: 1px solid {t['border']}; }}"
        )
        self.sb_sep_top.set_theme(self.theme_name)
        self.sb_sep_bot.set_theme(self.theme_name)
        self.sb_title.setStyleSheet(
            f"font-size: 13px; font-weight: bold; color: {t['text']}; background: transparent;"
        )
        self.book_count_lbl.setStyleSheet(
            f"font-size: 11px; color: {t['muted']}; background: transparent;"
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
        self.btn_settings.setStyleSheet(toggle_style)

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
        self.grp_appearance.setStyleSheet(group_style)
        self.grp_open.setStyleSheet(group_style)

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
        self.chk_open_overlay.setStyleSheet(checkbox_style)
        self.chk_display_duplicate.setStyleSheet(checkbox_style)
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
        for lbl in (self.lbl_theme, self.lbl_font, self.lbl_overlay_secs, self.lbl_overlay_pos, self.fs_lbl):
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

        # Status placeholder
        self.status_lbl.setStyleSheet(
            f"font-size: 11px; color: {t['muted']}; background: {t['panel2']}; "
            f"padding: 4px 14px; border-radius: 12px;"
        )

    # ─────────────────────────────────────────────────────────────
    #  SCAN — 啟動掃描書冊與內容索引管理
    # ─────────────────────────────────────────────────────────────
    def _scan_all(self):
        t = THEMES[self.theme_name]
        self.books = scan_books(HYMN_FOLDER)
        self.book_list.clear()

        if not self.books:
            item = QListWidgetItem("  ⚠  找不到資料夾")
            item.setForeground(QColor(t['err']))
            self.book_list.addItem(item)
            self._set_status('err', f"找不到：{os.path.basename(HYMN_FOLDER)}")
            return

        for b in self.books:
            badges = []
            if b['docx']:     badges.append('W')
            if b['bookmark']: badges.append('🔖')
            elif b['pdf']:    badges.append('P')
            cnt = f" ({b['hymn_count']})" if b['hymn_count'] else ''
            item = QListWidgetItem(
                f"  📖  {b['name']}{cnt}  {'  '.join(badges)}"
            )
            item.setData(Qt.ItemDataRole.UserRole, b)
            self.book_list.addItem(item)

        self.book_count_lbl.setText(f"{len(self.books)} 本")

        if self.books:
            self.current_book = self.books[0]
            self.book_list.setCurrentRow(0)
            self._show_book_files(self.current_book)
            self._update_sel_info(self.current_book)

        self._set_status('ok', f"Ready · {len(self.books)} 本已索引")
        self._load_content_index()

    def _load_content_index(self):
        self.content_index = load_content_index(HYMN_FOLDER)
        self._update_index_status_label()

    def _update_index_status_label(self):
        nfiles, nchunks, built_at = index_stats(self.content_index)
        if not self.content_index:
            self.index_status_lbl.setText("索引：未建立（建議先按「建立索引」以加快搜索）")
            self.btn_refresh_index.setEnabled(False)
            return
        self.btn_refresh_index.setEnabled(True)
        ts = time.strftime('%Y-%m-%d %H:%M', time.localtime(built_at)) if built_at else '?'
        stale = index_is_stale(self.content_index)
        stale_note = '  ·  ⚠ 檔案已變更，建議刷新' if stale else ''
        self.index_status_lbl.setText(
            f"索引：{nfiles} 個檔案 · {nchunks} 段  ·  建立於 {ts}{stale_note}"
        )

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

        self._index_worker = IndexBuildWorker(self.books, HYMN_FOLDER)
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
            save_content_index(index, HYMN_FOLDER)
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

    # ─────────────────────────────────────────────────────────────
    #  SIDEBAR SELECTION — 左側書冊點選與切換
    # ─────────────────────────────────────────────────────────────
    def _sidebar_click(self, item):
        book = item.data(Qt.ItemDataRole.UserRole)
        if book:
            self._select_book(book)

    def _select_book(self, book):
        if self.search_mode in ('global', 'keyword'):
            self._set_search_mode('book')
        self.current_book = book
        self.inp_book.blockSignals(True)
        self.inp_book.setText(book['name'])
        self.inp_book.blockSignals(False)
        self._show_book_files(book, self.inp_num.text().strip().lower())
        self._update_sel_info(book)

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

        count = 0
        for fi in book.get('files', []):
            name_lower = fi['name'].lower()
            ext        = fi['ext']
            is_pdf     = ext == '.pdf'

            file_match = bool(q and q in name_lower)
            bm_match   = bool(
                is_pdf and book.get('bookmark') and q and
                any(q in title.lower() for _lvl, title, _page in book.get('toc', []))
            )

            if q and not file_match and not bm_match:
                continue

            tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
            item = QListWidgetItem(f"  [{tag}]  {fi['name']}")
            item.setData(Qt.ItemDataRole.UserRole, {'kind': 'file', 'path': fi['path']})
            self.file_list.addItem(item)
            count += 1

            if is_pdf and book.get('bookmark') and book.get('toc'):
                matches = [
                    (lvl, title, page) for lvl, title, page in book['toc']
                    if not q or q in title.lower() or file_match
                ]
                if matches:
                    add_section_hdr("    書籤")
                    for lvl, title, page in matches:
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
                    if q in title.lower():
                        matches.append({
                            'kind': 'bookmark', 'title': title, 'page': page,
                            'lvl': lvl, 'pdf_path': pdf_path,
                        })

            for fi in book.get('files', []):
                if q in fi['name'].lower():
                    ext = fi['ext']
                    tag = ext[1:].upper() if ext.startswith('.') else ext.upper()
                    matches.append({
                        'kind': 'file', 'path': fi['path'], 'name': fi['name'], 'tag': tag,
                    })

            if matches:
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
        payload = item.data(Qt.ItemDataRole.UserRole)
        if not payload:
            return
        kind = payload.get('kind')
        if kind == 'file':
            path = payload.get('path')
            if path and os.path.exists(path):
                try:
                    ext = os.path.splitext(path)[1].lower()
                    if ext in WORD_EXTS:
                        self._open_docx_with_flow(path, os.path.basename(path))
                    elif ext == '.pdf':
                        self._open_pdf_with_flow(path, os.path.basename(path))
                    elif sys.platform == 'win32':
                        os.startfile(path)
                    elif sys.platform == 'darwin':
                        subprocess.Popen(['open', path])
                    else:
                        subprocess.Popen(['xdg-open', path])
                    self._show_toast(f"✅  正在開啟：{os.path.basename(path)}")
                except Exception as e:
                    self._show_toast(f"❌  無法開啟：{e}", 'err')
        elif kind == 'bookmark':
            try:
                self._open_pdf_with_flow(
                    payload['pdf_path'],
                    f"{payload['title']}  ·  P.{payload['page']}",
                    page=payload['page'],
                )
                self._show_toast(
                    f"✅  開啟 P.{payload['page']}：{payload['title'][:30]}"
                )
            except Exception as e:
                self._show_toast(f"❌  無法開啟書籤：{e}", 'err')
        elif kind == 'content_pdf':
            try:
                sub = f"P.{payload['page']}"
                snip = (payload.get('snippet') or '').strip()
                if snip:
                    sub += f"  {snip[:40]}"
                self._open_pdf_with_flow(payload['path'], sub, page=payload['page'])
                self._show_toast(f"✅  開啟 P.{payload['page']}")
            except Exception as e:
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
                    self._open_docx_with_flow(path, sub)
                    self._show_toast(f"✅  正在開啟：{os.path.basename(path)}")
                except Exception as e:
                    self._show_toast(f"❌  無法開啟：{e}", 'err')

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
        if self.search_mode == 'book':
            self.btn_filter.setText('Open' if single else 'Filter')
        elif self.search_mode == 'global':
            self.btn_global_search.setText('Open' if single else '搜索')
        elif self.search_mode == 'keyword':
            self.btn_keyword_search.setText('Open' if single else '搜索')

    def _open_first_result(self):
        items = self._actionable_file_items()
        if items:
            self._file_clicked(items[0])

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

    # ─────────────────────────────────────────────────────────────
    #  SEARCH MODE — 切換書本/全局/關鍵字模式與主題、字體
    # ─────────────────────────────────────────────────────────────
    def _set_search_mode(self, mode):
        if mode != 'keyword':
            self._stop_search_worker()
            self._keyword_debounce.stop()

        self.search_mode = mode
        self.btn_book.setChecked(mode == 'book')
        self.btn_global.setChecked(mode == 'global')
        self.btn_keyword.setChecked(mode == 'keyword')
        self.book_row.setVisible(mode == 'book')
        self.global_row.setVisible(mode == 'global')
        self.keyword_row.setVisible(mode == 'keyword')
        self.keyword_index_row.setVisible(mode == 'keyword')
        self.dropdown.hide()

        if mode == 'book':
            self.file_hdr.setText("檔案 / 書籤")
            self.hint_lbl.setText("選一個左邊書冊，右邊會列出該書的所有檔案")
            if self.current_book:
                self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
                self._update_sel_info(self.current_book)
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
        elif self.current_book:
            self._show_book_files(self.current_book, self.inp_num.text().strip().lower())
            self._update_sel_info(self.current_book)
        # Refresh error-colored items if any
        t = THEMES[self.theme_name]
        for i in range(self.book_list.count()):
            item = self.book_list.item(i)
            if not item.data(Qt.ItemDataRole.UserRole):
                item.setForeground(QColor(t['err']))

    def _on_font_changed(self, val):
        self.font_size = val
        self.fs_lbl.setText(f"{val}px")
        self._apply_theme()
        if self.search_mode == 'global':
            self._show_global_results(self.inp_global.text())
        elif self.search_mode == 'keyword':
            pass  # keep current keyword results
        elif self.current_book:
            self._show_book_files(self.current_book, self.inp_num.text().strip().lower())

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
    #  EVENT FILTER — inp_book 聚焦時清空、失焦時隱藏下拉
    # ─────────────────────────────────────────────────────────────
    def eventFilter(self, obj, event):
        if obj == self.inp_book:
            if (event.type() == QEvent.Type.FocusIn
                    and self.search_mode == 'book'
                    and self.inp_book.text()):
                self.inp_book.clear()
                self.dropdown.hide()
            elif event.type() == QEvent.Type.FocusOut:
                QTimer.singleShot(200, self.dropdown.hide)
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