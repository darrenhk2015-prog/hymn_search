"""In-memory remote operation log (thread-safe)."""
import threading
import time
from collections import deque

MAX_ENTRIES = 200


class RemoteOperationLog:
    def __init__(self, max_entries=MAX_ENTRIES):
        self._lock = threading.Lock()
        self._entries = deque(maxlen=max_entries)

    def add(self, action, detail='', client=''):
        entry = {
            'time': time.strftime('%Y-%m-%d %H:%M:%S'),
            'ts': time.time(),
            'action': action,
            'detail': detail,
            'client': client,
        }
        with self._lock:
            self._entries.appendleft(entry)
        return entry

    def list_entries(self, limit=50):
        with self._lock:
            return list(self._entries)[:limit]
