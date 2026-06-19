"""EnhancementMixin — optional MainWindow features (session, setlist, tray, operator mode)."""
import json
import os
import sys

from PyQt6.QtCore import Qt, QTimer, QFileSystemWatcher
from PyQt6.QtGui import QFontMetrics, QImage, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QLabel,
    QListWidgetItem,
    QMenu,
    QSystemTrayIcon,
    QVBoxLayout,
)

from hymn_remote.resolver import resolve_books, resolve_open_request, summarize_match

from .preview import preview_for_payload, preview_lines
from .session_store import (
    EXPORT_KEYS,
    advance_setlist_index,
    clamp_setlist,
    export_settings_json,
    format_setlist,
    format_setlist_compact,
    import_settings_json,
    is_book_pinned,
    navigate_session_history,
    parse_setlist,
    pin_book,
    record_open,
    setlist_index,
    unpin_book,
)


class EnhancementMixin:
    """Mixin for MainWindow; expects books, _settings, _remote, and UI widgets on self."""

    # ── Footer ──────────────────────────────────────────────────────────────

    def _update_footer_status(self):
        if not hasattr(self, 'footer_status_lbl'):
            return
        if hasattr(self, 'status_lbl'):
            self.footer_status_lbl.setText(self.status_lbl.text())

        last = str(self._settings.get('last_opened') or '')
        if hasattr(self, 'footer_last_lbl'):
            if last:
                fm = QFontMetrics(self.footer_last_lbl.font())
                w = max(80, self.footer_last_lbl.width() - 8)
                self.footer_last_lbl.setText(fm.elidedText(f"上次：{last}", Qt.TextElideMode.ElideRight, w))
                self.footer_last_lbl.setToolTip(last)
            else:
                self.footer_last_lbl.setText('上次：—')
                self.footer_last_lbl.setToolTip('')

        pending = 0
        if getattr(self, '_remote', None):
            try:
                pending = len(self._remote.state.list_pending())
            except Exception:
                pending = 0
        if hasattr(self, 'footer_pending_lbl'):
            if pending:
                self.footer_pending_lbl.setText(f'待辦 {pending}')
            else:
                self.footer_pending_lbl.setText('')

        if hasattr(self, 'footer_remote_lbl'):
            remote_txt = '遙控：關閉'
            if getattr(self, '_remote', None) and getattr(self._remote, 'running', False):
                if self._remote.state.accepting_requests:
                    remote_txt = '遙控：接受'
                else:
                    remote_txt = '遙控：暫停'
            self.footer_remote_lbl.setText(remote_txt)

    # ── Open tracking ───────────────────────────────────────────────────────

    def _payload_open_meta(self, payload):
        if not payload:
            return {'book': '', 'num': '', 'label': ''}
        kind = payload.get('kind')
        book = payload.get('book', '')
        if not book and payload.get('path'):
            book = self._book_name_for_path(payload['path'])
        if kind == 'bookmark':
            num = payload.get('title', '')
            label = summarize_match(payload)
        elif kind == 'file':
            num = os.path.basename(payload.get('path', ''))
            label = summarize_match({**payload, 'book': book}) if book else num
        else:
            num = str(payload.get('page') or payload.get('name') or '')
            label = preview_for_payload(payload, max_len=80) or num
        return {'book': book, 'num': num, 'label': label}

    def _record_open(self, payload=None, book_name='', hymn_ref='', label=''):
        if payload:
            meta = self._payload_open_meta(payload)
            book_name = meta.get('book') or book_name
            hymn_ref = meta.get('num') or hymn_ref
            label = meta.get('label') or label
        self._settings = record_open(self._settings, book_name, hymn_ref, label)
        self._persist_settings(**{k: self._settings[k] for k in (
            'recent_hymns', 'last_opened', 'session_history', 'session_cursor',
        ) if k in self._settings})
        self._update_footer_status()

    def _open_hymn_ref(self, book_ref, num_ref='', record=True):
        matches = resolve_open_request(self.books, book_ref, num_ref)
        if not matches:
            self._show_toast(f'找不到：{book_ref} / {num_ref}', 'warn')
            return False
        if len(matches) == 1:
            self._open_from_payload(matches[0])
            if record:
                self._record_open(matches[0])
            return True
        if hasattr(self, '_sync_remote_query'):
            self._sync_remote_query(book_ref, num_ref)
        self._populate_matches_in_file_list(matches)
        self._show_toast(f'{len(matches)} 個結果 — 請揀選', 'warn')
        return False

    def _populate_matches_in_file_list(self, matches):
        self.file_list.clear()
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
            self.file_list.addItem(item)
        if hasattr(self, 'file_count_lbl'):
            self.file_count_lbl.setText(f'{len(matches)} 項')
        if hasattr(self, '_update_search_action_buttons'):
            self._update_search_action_buttons()

    # ── Setlist ─────────────────────────────────────────────────────────────

    def _get_setlist_info(self):
        self._settings = clamp_setlist(self._settings)
        entries = parse_setlist(self._settings.get('setlist'))
        idx = setlist_index(self._settings)
        cur = entries[idx] if entries and 0 <= idx < len(entries) else None
        return {
            'entries': entries,
            'index': idx,
            'total': len(entries),
            'current': cur,
            'label': format_setlist_compact(entries, idx),
        }

    def _apply_setlist_from_settings(self):
        self._settings = clamp_setlist(self._settings)
        if getattr(self, 'search_mode', '') == 'schedule' and hasattr(self, '_refresh_schedule_list'):
            self._refresh_schedule_list()
            if hasattr(self, '_refresh_schedule_search'):
                self._refresh_schedule_search()
        self._update_footer_status()

    def _update_setlist_compact_label(self):
        info = self._get_setlist_info()
        if hasattr(self, 'lbl_sched_info'):
            self.lbl_sched_info.setText(info['label'])

    def _setlist_go(self, index=None):
        info = self._get_setlist_info()
        if not info['entries']:
            self._show_toast('清單為空', 'warn')
            return False
        idx = info['index'] if index is None else index
        idx = max(0, min(len(info['entries']) - 1, idx))
        self._settings['setlist_index'] = idx
        self._persist_settings(setlist_index=idx)
        entry = info['entries'][idx]
        ok = self._open_hymn_ref(entry.get('book', ''), entry.get('num', ''))
        if getattr(self, 'search_mode', '') == 'schedule' and hasattr(self, '_refresh_schedule_list'):
            self._refresh_schedule_list(select_index=idx)
        else:
            self._update_setlist_compact_label()
        self._update_footer_status()
        return ok

    def _setlist_next(self):
        self._settings = advance_setlist_index(self._settings, 1)
        self._persist_settings(
            setlist_index=self._settings['setlist_index'],
        )
        return self._setlist_go(self._settings['setlist_index'])

    def _setlist_prev(self):
        self._settings = advance_setlist_index(self._settings, -1)
        self._persist_settings(
            setlist_index=self._settings['setlist_index'],
        )
        return self._setlist_go(self._settings['setlist_index'])

    # ── Session history ─────────────────────────────────────────────────────

    def _open_previous_session(self):
        self._settings, entry = navigate_session_history(self._settings, 1)
        if not entry:
            self._show_toast('沒有更早的記錄', 'warn')
            return False
        self._persist_settings(
            session_history=self._settings.get('session_history'),
            session_cursor=self._settings.get('session_cursor', 0),
        )
        return self._open_hymn_ref(entry.get('book', ''), entry.get('num', ''), record=False)

    def _open_next_session(self):
        self._settings, entry = navigate_session_history(self._settings, -1)
        if not entry:
            self._show_toast('沒有更晚的記錄', 'warn')
            return False
        self._persist_settings(
            session_history=self._settings.get('session_history'),
            session_cursor=self._settings.get('session_cursor', 0),
        )
        return self._open_hymn_ref(entry.get('book', ''), entry.get('num', ''), record=False)

    # ── File preview ────────────────────────────────────────────────────────

    def _on_file_selection_changed(self, current=None, previous=None):
        if not self._settings.get('show_preview', True):
            if hasattr(self, 'preview_lbl'):
                self.preview_lbl.clear()
                self.preview_lbl.setToolTip('')
            return
        item = current
        if item is None and hasattr(self, 'file_list'):
            item = self.file_list.currentItem()
        payload = item.data(Qt.ItemDataRole.UserRole) if item else None
        text = preview_for_payload(payload) if payload else ''
        lines = preview_lines(text, max_lines=2, max_chars=140)
        display = '\n'.join(lines)
        if hasattr(self, 'preview_lbl'):
            self.preview_lbl.setText(display)
            self.preview_lbl.setToolTip(text if text and text != display else '')

    # ── Book list pins / sections ───────────────────────────────────────────

    def _book_list_context_menu(self, pos):
        item = self.book_list.itemAt(pos)
        if not item:
            return
        book = item.data(Qt.ItemDataRole.UserRole)
        if not book:
            return
        menu = QMenu(self)
        pinned = is_book_pinned(self._settings, book['name'])
        pin_action = menu.addAction('取消釘選' if pinned else '釘選此書')
        pin_action.triggered.connect(lambda: self._pin_book_from_menu(book['name'], unpin=pinned))
        menu.exec(self.book_list.mapToGlobal(pos))

    def _pin_book_from_menu(self, book_name, unpin=False):
        if unpin:
            self._settings = unpin_book(self._settings, book_name)
        else:
            self._settings = pin_book(self._settings, book_name)
        self._persist_settings(pinned_books=self._settings['pinned_books'])
        self._rebuild_book_list()
        self._show_toast('已取消釘選' if unpin else '已釘選', 'ok')

    def _rebuild_book_list(self):
        if not self.books:
            return
        t = __import__('hymn_search', fromlist=['THEMES']).THEMES[self.theme_name]
        from PyQt6.QtGui import QColor

        selected = self.current_book
        self.book_list.clear()

        def add_header(text):
            i = QListWidgetItem(f'  {text}')
            i.setFlags(Qt.ItemFlag.ItemIsEnabled)
            i.setForeground(QColor(t['muted']))
            f = i.font()
            f.setBold(True)
            i.setFont(f)
            self.book_list.addItem(i)

        def add_book(book, prefix=''):
            badges = []
            if book.get('docx'):
                badges.append('W')
            if book.get('bookmark'):
                badges.append('🔖')
            elif book.get('pdf'):
                badges.append('P')
            cnt = f" ({book['hymn_count']})" if book.get('hymn_count') else ''
            pin = '📌 ' if is_book_pinned(self._settings, book['name']) else ''
            item = QListWidgetItem(
                f"  {prefix}{pin}📖  {book['name']}{cnt}  {'  '.join(badges)}"
            )
            item.setData(Qt.ItemDataRole.UserRole, book)
            self.book_list.addItem(item)
            return item

        by_name = {b['name']: b for b in self.books}
        pins = [by_name[n] for n in self._settings.get('pinned_books', []) if n in by_name]
        shown = set()

        if pins:
            add_header('釘選')
            for book in pins:
                add_book(book)
                shown.add(book['name'])

        add_header('全部')
        for book in self.books:
            if book['name'] in shown:
                continue
            add_book(book)

        self.book_count_lbl.setText(f'{len(self.books)} 本')

        if selected:
            for i in range(self.book_list.count()):
                item = self.book_list.item(i)
                b = item.data(Qt.ItemDataRole.UserRole)
                if b and b['name'] == selected['name']:
                    self.book_list.setCurrentRow(i)
                    break

    # ── Operator mode (removed — kept as no-op for settings import compat) ──

    def _sync_operator_mode_checkboxes(self):
        pass

    def _apply_operator_mode_ui(self):
        pass

    def _on_operator_mode_toggled(self, checked):
        pass

    def _on_show_preview_toggled(self, checked):
        self._persist_settings(show_preview=checked)
        self._on_file_selection_changed()

    def _on_auto_rescan_toggled(self, checked):
        self._persist_settings(auto_rescan=checked)

    def _on_startup_tray_toggled(self, checked):
        self._persist_settings(startup_tray=checked)

    def _on_minimize_tray_toggled(self, checked):
        self._persist_settings(minimize_to_tray=checked)

    # ── Folder watcher ──────────────────────────────────────────────────────

    def _setup_folder_watcher(self):
        self._folder_watcher = QFileSystemWatcher(self)
        self._folder_watcher.directoryChanged.connect(self._on_folder_changed)
        self._rescan_debounce = QTimer(self)
        self._rescan_debounce.setSingleShot(True)
        self._rescan_debounce.timeout.connect(self._auto_rescan_books)
        if os.path.isdir(self.hymn_folder):
            self._folder_watcher.addPath(self.hymn_folder)

    def _on_folder_changed(self, _path=''):
        if not self._settings.get('auto_rescan', True):
            return
        self._rescan_debounce.start(800)

    def _auto_rescan_books(self):
        if hasattr(self, '_scan_all'):
            self._set_status('warn', '資料夾已變更，重新掃描…')
            self._scan_all()
            if hasattr(self, '_rebuild_book_list'):
                self._rebuild_book_list()
            self._update_footer_status()

    # ── Shortcuts dialog ────────────────────────────────────────────────────

    def _show_shortcuts_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle('快捷鍵')
        lay = QVBoxLayout(dlg)
        lines = [
            'F1 / F2 / F3 / F4 — 書本 / 全局 / 關鍵字 / 排程模式',
            'Ctrl+Enter — 聚焦書冊輸入（書本模式）',
            'Enter — 確認 Mobile 待辦（有待確認時）',
            'Ctrl+← / Ctrl+→ — Session 上一首 / 下一首',
            'Alt+← / Alt+→ — 排程上一首 / 下一首',
            'Filter / Open — 書本模式過濾或開啟唯一結果',
            '雙擊搜尋結果 — 排程模式加入播放清單',
            '雙擊播放清單 — 開啟該首',
        ]
        body = QLabel('\n'.join(lines))
        body.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        lay.addWidget(body)
        dlg.resize(420, 220)
        dlg.exec()

    # ── System tray ─────────────────────────────────────────────────────────

    def _setup_tray_icon(self):
        icon = self.windowIcon()
        if icon.isNull():
            from hymn_search import load_app_icon
            icon = load_app_icon()
        self._tray = QSystemTrayIcon(icon, self)
        self._tray.setToolTip(self.windowTitle())

        menu = QMenu()
        show_action = menu.addAction('顯示主視窗')
        show_action.triggered.connect(self._show_from_tray)
        menu.addSeparator()
        quit_action = menu.addAction('結束')
        quit_action.triggered.connect(self._quit_app)
        self._tray.setContextMenu(menu)
        self._tray.activated.connect(self._tray_activated)
        if icon.isNull():
            return
        self._tray.show()

    def _show_from_tray(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def _minimize_to_tray(self):
        if hasattr(self, '_tray') and self._tray.isVisible():
            self.hide()
            self._tray.showMessage(
                self.windowTitle(),
                '程式已最小化到系統匣',
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )

    def _tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self._show_from_tray()

    def _quit_app(self):
        if getattr(self, '_remote', None):
            self._remote.stop()
        QApplication.instance().quit()

    # ── Settings export / import ────────────────────────────────────────────

    def _export_settings(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            '匯出設定',
            'hymn_settings.json',
            'JSON (*.json)',
        )
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(export_settings_json(self._settings))
            self._show_toast(f'已匯出：{os.path.basename(path)}', 'ok')
        except OSError as e:
            self._show_toast(f'匯出失敗：{e}', 'err')

    def _import_settings(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            '匯入設定',
            '',
            'JSON (*.json)',
        )
        if not path:
            return
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            self._settings = import_settings_json(self._settings, data)
            self._persist_settings(**{k: self._settings[k] for k in EXPORT_KEYS if k in self._settings})
            self._sync_operator_mode_checkboxes()
            self._apply_operator_mode_ui()
            self._apply_setlist_from_settings()
            self._rebuild_book_list()
            self._update_footer_status()
            self._show_toast('設定已匯入', 'ok')
        except (OSError, json.JSONDecodeError, TypeError) as e:
            self._show_toast(f'匯入失敗：{e}', 'err')

    # ── Remote QR / log ─────────────────────────────────────────────────────

    def _qr_pixmap_from_url(self, url, size=128):
        from io import BytesIO
        import qrcode
        buf = BytesIO()
        qrcode.make(url).save(buf, format='PNG')
        png = bytes(buf.getvalue())
        self._qr_png_cache = png
        qimg = QImage.fromData(png, 'PNG')
        if qimg.isNull():
            raise ValueError('QR PNG decode failed')
        pix = QPixmap.fromImage(qimg.copy())
        if pix.isNull():
            raise ValueError('QR pixmap failed')
        scaled = pix.scaled(
            size, size,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._qr_pixmap_cache = scaled
        return scaled

    def _update_qr_code(self):
        lbl = getattr(self, 'remote_qr_lbl', None)
        if lbl is None:
            return
        lbl.clear()
        if not getattr(self._remote, 'running', False):
            lbl.setText('啟用 Mobile API 後顯示 QR')
            lbl.setToolTip('')
            return
        url = self._remote.url()
        lbl.setToolTip(url)
        try:
            pix = self._qr_pixmap_from_url(url, size=128)
            lbl.setPixmap(pix)
        except ImportError:
            lbl.setText('缺少 qrcode\n請 pip install')
            lbl.setToolTip(url)
        except Exception as exc:
            lbl.setText('QR 無法顯示')
            lbl.setToolTip(f'{url}\n({exc})')

    def _refresh_remote_log_view(self):
        view = getattr(self, 'remote_log_list', None)
        if view is None:
            return
        log = getattr(self._remote, 'operation_log', None)
        if log is None:
            view.clear()
            return
        view.clear()
        for entry in log.list_entries(limit=50):
            line = f"{entry.get('time', '')}  {entry.get('action', '')}  {entry.get('detail', '')}"
            view.addItem(line)

    # ── Close handling ──────────────────────────────────────────────────────

    def handle_close_event(self, event):
        if self._settings.get('minimize_to_tray') and hasattr(self, '_tray'):
            if self._tray.isVisible():
                event.ignore()
                self._minimize_to_tray()
                return True
        return False
