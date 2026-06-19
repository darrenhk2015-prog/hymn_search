"""Thread-safe remote control state shared between FastAPI and MainWindow."""
import threading
import uuid
from dataclasses import dataclass
from typing import Callable, Optional

from .operation_log import RemoteOperationLog


DEFAULT_PORT = 8765
OPEN_POLICY_AUTO = 'auto_single'
OPEN_POLICY_CONFIRM = 'confirm'
OPEN_POLICY_UI = 'ui_only'


@dataclass
class PendingRequest:
    id: str
    book_ref: str
    num_ref: str
    matches: list
    summary: str


class RemoteControlState:
    def __init__(self):
        self._lock = threading.Lock()
        self.api_enabled = False
        self.accepting_requests = False
        self.open_policy = OPEN_POLICY_AUTO
        self.api_token = ''
        self.port = DEFAULT_PORT
        self._pending: dict[str, PendingRequest] = {}
        self._handle_open: Optional[Callable] = None
        self._handle_populate: Optional[Callable] = None
        self._handle_pending: Optional[Callable] = None
        self._handle_setlist_prepare_next: Optional[Callable] = None
        self._handle_setlist_prepare_open: Optional[Callable] = None
        self._handle_setlist_add: Optional[Callable] = None
        self._get_books: Optional[Callable] = None
        self._get_setlist_info: Optional[Callable] = None
        self.operation_log = RemoteOperationLog()

    def set_handlers(
        self,
        get_books,
        handle_open,
        handle_populate,
        handle_pending,
        handle_setlist_prepare_next=None,
        get_setlist_info=None,
        handle_setlist_prepare_open=None,
        handle_setlist_add=None,
    ):
        self._get_books = get_books
        self._handle_open = handle_open
        self._handle_populate = handle_populate
        self._handle_pending = handle_pending
        self._handle_setlist_prepare_next = handle_setlist_prepare_next
        self._get_setlist_info = get_setlist_info
        self._handle_setlist_prepare_open = handle_setlist_prepare_open
        self._handle_setlist_add = handle_setlist_add

    def get_books(self):
        if self._get_books:
            return self._get_books()
        return []

    def get_setlist_info(self):
        if self._get_setlist_info:
            return self._get_setlist_info()
        return {'entries': [], 'index': -1, 'total': 0, 'current': None, 'label': ''}

    def snapshot(self):
        with self._lock:
            return {
                'api_enabled': self.api_enabled,
                'accepting': self.accepting_requests,
                'policy': self.open_policy,
                'port': self.port,
                'has_token': bool(self.api_token),
                'pending_count': len(self._pending),
                'setlist': self.get_setlist_info(),
            }

    def log(self, action, detail='', client=''):
        return self.operation_log.add(action, detail, client)

    def set_api_enabled(self, enabled: bool):
        with self._lock:
            self.api_enabled = enabled
            if not enabled:
                self.accepting_requests = False

    def set_accepting(self, accepting: bool):
        with self._lock:
            if self.api_enabled:
                self.accepting_requests = accepting

    def set_policy(self, policy: str):
        with self._lock:
            self.open_policy = policy

    def set_token(self, token: str):
        with self._lock:
            self.api_token = (token or '').strip()

    def set_port(self, port: int):
        with self._lock:
            self.port = port

    def check_token(self, header_token: str) -> bool:
        with self._lock:
            if not self.api_token:
                return True
            return header_token == self.api_token

    def is_accepting(self) -> bool:
        with self._lock:
            return self.api_enabled and self.accepting_requests

    def get_policy(self) -> str:
        with self._lock:
            return self.open_policy

    def add_pending(self, book_ref, num_ref, matches) -> PendingRequest:
        from .resolver import summarize_match
        req = PendingRequest(
            id=str(uuid.uuid4())[:8],
            book_ref=str(book_ref),
            num_ref=str(num_ref),
            matches=matches,
            summary=summarize_match(matches[0]) if len(matches) == 1 else f"{len(matches)} 個結果",
        )
        with self._lock:
            self._pending[req.id] = req
        if self._handle_pending:
            self._handle_pending(req)
        return req

    def list_pending(self):
        with self._lock:
            return list(self._pending.values())

    def pop_pending(self, req_id: str) -> Optional[PendingRequest]:
        with self._lock:
            return self._pending.pop(req_id, None)

    def _dispatch_open_with_policy(self, book_ref, num_ref):
        from .resolver import resolve_open_request, summarize_match

        books = self.get_books()
        matches = resolve_open_request(books, book_ref, num_ref)
        if not matches:
            return {
                'status': 'not_found',
                'book_ref': str(book_ref),
                'num_ref': str(num_ref),
                'match_count': 0,
                'matches': [],
            }

        enriched = [{**m, 'summary': summarize_match(m)} for m in matches]
        policy = self.get_policy()

        if policy == OPEN_POLICY_CONFIRM:
            req = self.add_pending(book_ref, num_ref, matches)
            return {
                'status': 'queued',
                'request_id': req.id,
                'book_ref': str(book_ref),
                'num_ref': str(num_ref),
                'match_count': len(matches),
                'matches': enriched,
            }

        if policy == OPEN_POLICY_UI or len(matches) != 1:
            if self._handle_populate:
                self._handle_populate(book_ref, num_ref, matches)
            return {
                'status': 'ui_populated',
                'book_ref': str(book_ref),
                'num_ref': str(num_ref),
                'match_count': len(matches),
                'matches': enriched,
            }

        if self._handle_populate:
            self._handle_populate(book_ref, num_ref, matches)
        if self._handle_open:
            self._handle_open(matches[0])
        return {
            'status': 'opened',
            'book_ref': str(book_ref),
            'num_ref': str(num_ref),
            'match_count': 1,
            'matches': enriched,
        }

    def dispatch_open_request(self, book_ref, num_ref, client=''):
        self.log('open', f'{book_ref} / {num_ref}', client)

        if not self.is_accepting():
            return {'status': 'not_accepting', 'message': '桌面端暫停接受請求，請聯絡操作員'}

        return self._dispatch_open_with_policy(book_ref, num_ref)

    def dispatch_next(self, client=''):
        self.log('next', '', client)
        if not self.is_accepting():
            return {'status': 'not_accepting', 'message': '桌面端暫停接受請求，請聯絡操作員'}
        if not self._handle_setlist_prepare_next:
            return {'status': 'unsupported', 'message': '桌面端未支援排程'}
        entry = self._handle_setlist_prepare_next()
        if not entry:
            return {'status': 'failed', 'setlist': self.get_setlist_info()}
        result = self._dispatch_open_with_policy(entry.get('book', ''), entry.get('num', ''))
        result['setlist'] = self.get_setlist_info()
        return result

    def dispatch_setlist_open(self, index, client=''):
        self.log('setlist_open', str(index), client)
        if not self.is_accepting():
            return {'status': 'not_accepting', 'message': '桌面端暫停接受請求，請聯絡操作員'}
        if not self._handle_setlist_prepare_open:
            return {'status': 'unsupported', 'message': '桌面端未支援排程開啟'}
        try:
            idx = int(index)
        except (TypeError, ValueError):
            return {'status': 'invalid', 'message': '索引無效'}
        entry = self._handle_setlist_prepare_open(idx)
        if not entry:
            return {'status': 'failed', 'setlist': self.get_setlist_info()}
        result = self._dispatch_open_with_policy(entry.get('book', ''), entry.get('num', ''))
        result['setlist'] = self.get_setlist_info()
        return result

    def dispatch_setlist_add(self, book_ref, num_ref, client=''):
        from .resolver import resolve_books

        self.log('setlist_add', f'{book_ref} / {num_ref}', client)
        with self._lock:
            if not self.api_enabled:
                return {'status': 'disabled', 'message': '桌面端 API 未啟用'}
        if not self._handle_setlist_add:
            return {'status': 'unsupported', 'message': '桌面端未支援加入排程'}
        book_ref = str(book_ref or '').strip()
        num_ref = str(num_ref or '').strip()
        if not book_ref and not num_ref:
            return {'status': 'invalid', 'message': '請輸入書冊或詩歌號'}
        books = resolve_books(self.get_books(), book_ref)
        if not books:
            return {'status': 'not_found', 'message': '找不到書冊'}
        if len(books) > 1:
            return {
                'status': 'ambiguous',
                'message': f'找到 {len(books)} 本書冊，請指定更精確的書名或序號',
            }
        ok = self._handle_setlist_add(books[0]['name'], num_ref)
        if ok:
            return {'status': 'added', 'setlist': self.get_setlist_info()}
        return {'status': 'failed', 'setlist': self.get_setlist_info()}

    def approve_pending(self, req_id: str, client=''):
        req = self.pop_pending(req_id)
        if not req:
            return {'status': 'not_found', 'message': '請求不存在或已處理'}
        self.log('approve', f'{req.book_ref} / {req.num_ref}', client)
        if len(req.matches) == 1:
            if self._handle_populate:
                self._handle_populate(req.book_ref, req.num_ref, req.matches)
            if self._handle_open:
                self._handle_open(req.matches[0])
            return {
                'status': 'opened',
                'book_ref': req.book_ref,
                'num_ref': req.num_ref,
                'match_count': 1,
            }
        if self._handle_populate:
            self._handle_populate(req.book_ref, req.num_ref, req.matches)
        return {'status': 'ui_populated', 'match_count': len(req.matches)}

    def reject_pending(self, req_id: str, client=''):
        req = self.pop_pending(req_id)
        if not req:
            return {'status': 'not_found', 'message': '請求不存在或已處理'}
        self.log('reject', f'{req.book_ref} / {req.num_ref}', client)
        return {'status': 'rejected'}
