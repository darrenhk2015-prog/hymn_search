"""Background uvicorn server lifecycle."""
import os
import socket
import sys
import threading
import time
import traceback

from .api import create_app
from .local_tls import cloudflared_origin_ca_hint, deploy_cloudflared_origin_ca, ensure_localhost_cert
from .state import RemoteControlState, DEFAULT_PORT


def _app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _log_remote_error(exc):
    try:
        path = os.path.join(_app_dir(), 'remote_api.log')
        with open(path, 'a', encoding='utf-8') as f:
            f.write(f'\n--- {time.strftime("%Y-%m-%d %H:%M:%S")} ---\n')
            f.write(''.join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except OSError:
        pass


def _prepare_stdio_for_uvicorn():
    """PyInstaller --windowed sets stdout/stderr to None; uvicorn logging needs them."""
    import io
    if sys.stdout is None:
        sys.stdout = io.StringIO()
    if sys.stderr is None:
        sys.stderr = io.StringIO()


def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return '127.0.0.1'


class RemoteServer:
    def __init__(self):
        self.state = RemoteControlState()
        self._thread = None
        self._ipv6_thread = None
        self._server = None
        self._ipv6_server = None
        self._app = None
        self._start_error = None
        self._use_https = False

    def set_handlers(self, get_books, handle_open, handle_populate, handle_pending,
                     handle_setlist_prepare_next=None, get_setlist_info=None,
                     handle_setlist_prepare_open=None, handle_setlist_add=None):
        self.state.set_handlers(
            get_books, handle_open, handle_populate, handle_pending,
            handle_setlist_prepare_next, get_setlist_info,
            handle_setlist_prepare_open, handle_setlist_add,
        )

    @property
    def running(self):
        return self._is_port_open('127.0.0.1', self.state.port)

    @property
    def last_error(self):
        return self._start_error

    def _is_port_open(self, host, port):
        try:
            family = socket.AF_INET6 if ':' in host else socket.AF_INET
            s = socket.socket(family, socket.SOCK_STREAM)
            s.settimeout(0.3)
            s.connect((host, port))
            s.close()
            return True
        except OSError:
            return False

    def _wait_until_listening(self, port, timeout=4.0, use_https=False):
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self._start_error:
                return False
            if self._thread and not self._thread.is_alive():
                return False
            if self._is_port_open('127.0.0.1', port):
                if not use_https or not sys.platform.startswith('win'):
                    return True
                if self._is_port_open('::1', port):
                    return True
            time.sleep(0.1)
        if not self._is_port_open('127.0.0.1', port):
            return False
        if use_https and sys.platform.startswith('win'):
            return self._is_port_open('::1', port)
        return True

    def start(self, port=None, path_prefix='', use_https=False):
        if (
            self._is_port_open('127.0.0.1', self.state.port)
            and self._thread
            and self._thread.is_alive()
            and self._use_https == bool(use_https)
        ):
            return True
        port = port or self.state.port
        self.state.set_port(port)
        self._start_error = None
        self._use_https = bool(use_https)
        self._app = create_app(self.state, path_prefix=path_prefix)
        ssl_certfile = ssl_keyfile = None
        if self._use_https:
            try:
                cert_path, key_path = ensure_localhost_cert(_app_dir())
                ssl_certfile = str(cert_path)
                ssl_keyfile = str(key_path)
                deploy_cloudflared_origin_ca(cert_path)
            except RuntimeError as exc:
                self._start_error = str(exc)
                self._use_https = False
                _log_remote_error(exc)
                return False

        def _run(host):
            try:
                _prepare_stdio_for_uvicorn()
                import uvicorn
                config = uvicorn.Config(
                    self._app,
                    host=host,
                    port=port,
                    log_level='warning',
                    access_log=False,
                    log_config=None,
                    ssl_certfile=ssl_certfile,
                    ssl_keyfile=ssl_keyfile,
                )
                server = uvicorn.Server(config)
                if host == '::':
                    self._ipv6_server = server
                else:
                    self._server = server
                server.run()
            except Exception as e:
                self._start_error = str(e)
                _log_remote_error(e)

        self._thread = threading.Thread(target=_run, args=('0.0.0.0',), daemon=True, name='HymnRemoteAPI')
        self._thread.start()
        if self._use_https and sys.platform.startswith('win'):
            self._ipv6_thread = threading.Thread(
                target=_run, args=('::',), daemon=True, name='HymnRemoteAPI-IPv6'
            )
            self._ipv6_thread.start()
        return self._wait_until_listening(port, use_https=self._use_https)

    def stop(self):
        if self._server:
            self._server.should_exit = True
        if self._ipv6_server:
            self._ipv6_server.should_exit = True
        self._thread = None
        self._ipv6_thread = None
        self._server = None
        self._ipv6_server = None

    def url(self, path: str = ''):
        scheme = 'https' if self._use_https else 'http'
        base = f"{scheme}://{get_lan_ip()}:{self.state.port}"
        if not path:
            return base
        if not path.startswith('/'):
            path = '/' + path
        return base + path
