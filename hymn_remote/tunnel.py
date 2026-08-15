"""Cloudflare tunnel helpers for Hymn Search remote API."""
from __future__ import annotations

import http.client
import logging
import os
import re
import shutil
import socket
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Callable, TextIO
from urllib.parse import urlparse

from .local_tls import cloudflared_origin_ca_hint
from .state import DEFAULT_PORT

logger = logging.getLogger(__name__)

_TUNNEL_EVENT_LOG = "tunnel.log"
_CLOUDFLARED_STDERR_LOG = "cloudflared-tunnel.log"
_TUNNEL_LOG_LOCK = threading.Lock()
_SHUTDOWN_JOIN_TIMEOUT_S = 3.0
_TUNNEL_URL_RE = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com", re.I)
_IPV4_RE = re.compile(r"\b(\d{1,3}(?:\.\d{1,3}){3})\b")
_PUBLIC_DNS = ("1.1.1.1", "8.8.8.8")
_DNS_SERVERS = frozenset(_PUBLIC_DNS)
_HEALTH_INTERVAL_S = 10.0
_PROBE_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
_HTTPS_ORIGIN_RE = re.compile(
    r'https://(?:localhost|127\.0\.0\.1)(?::\d+)?',
    re.I,
)
_UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
    re.I,
)


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(os.path.dirname(os.path.abspath(sys.executable)))
    return Path(__file__).resolve().parent.parent


def tunnel_log_path() -> Path:
    return app_dir() / _TUNNEL_EVENT_LOG


def cloudflared_stderr_log_path() -> Path:
    return app_dir() / _CLOUDFLARED_STDERR_LOG


def append_tunnel_log(level: str, message: str, **fields: str) -> None:
    """Append one line to tunnel.log beside the exe / project root."""
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    extras = " ".join(f"{key}={value}" for key, value in fields.items() if value)
    line = f"[{ts}] [{level}] {message}"
    if extras:
        line += f" | {extras}"
    with _TUNNEL_LOG_LOCK:
        try:
            with open(tunnel_log_path(), "a", encoding="utf-8") as handle:
                handle.write(line + "\n")
        except OSError as exc:
            logger.debug("tunnel log write failed: %s", exc)


def _append_cloudflared_session_header(mode: str, local_port: int, url: str = "") -> None:
    append_tunnel_log(
        "INFO",
        "cloudflared subprocess session",
        mode=mode,
        port=str(local_port),
        url=url,
        stderr_log=str(cloudflared_stderr_log_path()),
    )
    with _TUNNEL_LOG_LOCK:
        try:
            with open(cloudflared_stderr_log_path(), "a", encoding="utf-8") as handle:
                handle.write(
                    f"\n--- cloudflared session {time.strftime('%Y-%m-%d %H:%M:%S')} "
                    f"mode={mode} port={local_port} url={url or '-'} ---\n"
                )
        except OSError:
            pass


def normalize_tunnel_url(url: str) -> str:
    url = url.strip().rstrip("/")
    if not url:
        return ""
    if not url.startswith(("http://", "https://")):
        return f"https://{url}"
    return url


def tunnel_url_path_prefix(url: str) -> str:
    """Return URL path prefix (e.g. /hymn_search) from a named tunnel URL."""
    path = urlparse(normalize_tunnel_url(url)).path.rstrip("/")
    return path if path else ""


def tunnel_viewer_url(url: str) -> str:
    """Public viewer page URL; path-prefix tunnels need a trailing slash."""
    base = normalize_tunnel_url(url)
    if not base:
        return ""
    if tunnel_url_path_prefix(base):
        return f"{base}/"
    return base


def validate_run_token(token: str) -> tuple[bool, str]:
    token = token.strip()
    if not token:
        return False, "Cloudflare tunnel token is required"
    if token.startswith("eyJ") and len(token) > 80:
        return True, ""
    if _UUID_RE.match(token):
        return (
            False,
            "This is a Tunnel ID, not a run token. In Cloudflare Zero Trust go to "
            "Networks → Tunnels → your tunnel → Install connector, and copy the token "
            "(starts with eyJ...).",
        )
    return (
        False,
        "Invalid tunnel token. Use the JWT from Cloudflare Install connector, not the tunnel ID.",
    )


def _windows_candidates() -> list[Path]:
    candidates: list[Path] = []
    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_name)
        if base:
            candidates.append(Path(base) / "cloudflared" / "cloudflared.exe")
    localappdata = os.environ.get("LOCALAPPDATA")
    if localappdata:
        candidates.append(Path(localappdata) / "cloudflared" / "cloudflared.exe")
    candidates.append(Path.home() / "cloudflared" / "cloudflared.exe")
    candidates.append(app_dir() / "cloudflared.exe")
    return candidates


def find_cloudflared(custom_path: str = "") -> str | None:
    if custom_path:
        path = Path(custom_path)
        if path.is_file():
            return str(path.resolve())

    found = shutil.which("cloudflared")
    if found:
        return found

    for candidate in _windows_candidates():
        if candidate.is_file():
            logger.info("Found cloudflared at: %s", candidate)
            return str(candidate.resolve())

    return None


def wait_for_local_server(
    port: int = DEFAULT_PORT,
    timeout: float = 8.0,
    use_https: bool = False,
) -> bool:
    hosts = ("127.0.0.1", "localhost") if use_https else ("127.0.0.1",)
    scheme = "https" if use_https else "http"
    ctx = None
    if use_https:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    deadline = time.time() + timeout
    while time.time() < deadline:
        for host in hosts:
            url = f"{scheme}://{host}:{port}/api/health"
            try:
                with urllib.request.urlopen(url, timeout=1.5, context=ctx) as resp:
                    if resp.status == 200:
                        return True
            except (urllib.error.URLError, TimeoutError, OSError):
                pass
        time.sleep(0.25)
    return False


def _hostname_from_url(url: str) -> str | None:
    return urlparse(url).hostname


def _health_path(url: str) -> str:
    base = urlparse(url).path or ""
    path = (base.rstrip("/") + "/api/health").replace("//", "/")
    if not path.startswith("/"):
        path = "/" + path
    return path


def _extract_ipv4_addresses(text: str) -> list[str]:
    ips: list[str] = []
    for match in _IPV4_RE.finditer(text):
        ip = match.group(1)
        if ip in _DNS_SERVERS:
            continue
        if ip not in ips:
            ips.append(ip)
    return ips


def _resolve_via_dns(hostname: str, dns_server: str | None = None) -> list[str]:
    flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    cmd = ["nslookup", hostname]
    if dns_server:
        cmd.append(dns_server)
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=6,
            creationflags=flags,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []

    output = f"{proc.stdout}\n{proc.stderr}"
    if re.search(r"Non-existent domain|NXDOMAIN", output, re.I):
        return []
    return _extract_ipv4_addresses(output)


def local_dns_resolves(hostname: str) -> bool:
    try:
        socket.getaddrinfo(hostname, 443, type=socket.SOCK_STREAM)
        return True
    except OSError:
        return False


def public_dns_resolves(hostname: str) -> bool:
    for dns in _PUBLIC_DNS:
        if _resolve_via_dns(hostname, dns):
            return True
    return False


def _localhost_ipv6_unreachable(port: int, use_https: bool = False) -> bool:
    """True when ::1 is down but 127.0.0.1 works (common Windows + cloudflared issue)."""
    if not use_https:
        return False
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=1.0):
            v4_ok = True
    except OSError:
        v4_ok = False
    try:
        with socket.create_connection(("::1", port), timeout=1.0):
            v6_ok = True
    except OSError:
        v6_ok = False
    return v4_ok and not v6_ok


def _https_origin_cloudflare_hint(local_port: int = DEFAULT_PORT) -> str:
    ca = cloudflared_origin_ca_hint()
    return (
        f"Zero Trust → Public Hostname → /hymn_search："
        f"Service=https://127.0.0.1:{local_port}；"
        f"No TLS Verify=開；Origin Server Name=127.0.0.1；"
        f"CA Pool={ca}；"
        "然後以管理員執行：net stop cloudflared && net start cloudflared。"
        f"若仍 502，改 Service=http://127.0.0.1:{local_port} 並關閉「本機 API 用 HTTPS」。"
    )


def _origin_setup_hint(local_port: int = DEFAULT_PORT, path: str = "/hymn_search") -> str:
    return (
        f"Cloudflare Zero Trust → Networks → Tunnels → Public Hostname："
        f"{path} → http://127.0.0.1:{local_port}（必須用 http，勿用 https 或 localhost）。"
    )


def _probe_request(host: str, path: str, timeout: float) -> tuple[int | None, str]:
    try:
        req = urllib.request.Request(f"https://{host}{path}", method="GET")
        req.add_header("User-Agent", _PROBE_UA)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read(256).decode("utf-8", errors="replace")
            return resp.status, body
    except urllib.error.HTTPError as exc:
        body = exc.read(256).decode("utf-8", errors="replace")
        return exc.code, body
    except Exception:
        return None, ""


def _probe_https_host(host: str, path: str, timeout: float) -> bool:
    status, _ = _probe_request(host, path, timeout)
    return status in (200, 401, 302)


def _probe_https_ip(ip: str, host: str, path: str, timeout: float) -> bool:
    conn: http.client.HTTPSConnection | None = None
    sock: socket.socket | None = None
    try:
        sock = socket.create_connection((ip, 443), timeout=timeout)
        ctx = ssl.create_default_context()
        wrapped = ctx.wrap_socket(sock, server_hostname=host)
        conn = http.client.HTTPSConnection(host, timeout=timeout, context=ctx)
        conn.sock = wrapped
        conn.request("GET", path, headers={"User-Agent": _PROBE_UA})
        return conn.getresponse().status in (200, 401, 302)
    except Exception:
        return False
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
        elif sock is not None:
            try:
                sock.close()
            except Exception:
                pass


def _latest_cloudflared_config_line() -> str:
    path = cloudflared_stderr_log_path()
    if not path.is_file():
        return ""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""
    for line in reversed(text.splitlines()):
        if "Updated to new configuration" in line:
            return line.replace("\\", "")
    return ""


def cloudflared_logged_hymn_origin(local_port: int = DEFAULT_PORT) -> str | None:
    """Best-effort parse of hymn_search origin from cloudflared stderr log."""
    text = _latest_cloudflared_config_line()
    if not text:
        return None
    match = re.search(
        rf'hymn_search.*?"service":"(https?://[^"]+:{local_port}[^"]*)"',
        text,
        re.I,
    )
    return match.group(1) if match else None


def diagnose_cloudflared_origin_mismatch(
    local_port: int = DEFAULT_PORT,
    use_https: bool = False,
) -> str | None:
    """Warn when cloudflared's logged origin does not match the app's HTTP/HTTPS mode."""
    origin = cloudflared_logged_hymn_origin(local_port)
    if not origin:
        return None
    if "localhost" in origin.lower():
        return (
            f"cloudflared 仍指向 {origin}（localhost 會走 IPv6 ::1）。"
            f"請改為 http://127.0.0.1:{local_port} 並重啟 cloudflared 服務。"
        )
    if not use_https and origin.lower().startswith("https://"):
        return (
            f"cloudflared 仍指向 {origin}，但本程式用 HTTP。"
            f"請在 Cloudflare 改 Service 為 http://127.0.0.1:{local_port}，"
            "關閉 No TLS Verify，儲存後以管理員執行 net stop cloudflared && net start cloudflared。"
        )
    if use_https and origin.lower().startswith("http://"):
        return (
            f"cloudflared 指向 {origin}，但本程式用 HTTPS。"
            f"請改 Cloudflare 為 https://127.0.0.1:{local_port} 或關閉「本機 API 用 HTTPS」。"
        )
    return None


def detect_https_origin_misconfiguration(
    local_port: int = DEFAULT_PORT,
    log_path: Path | None = None,
    use_https: bool = False,
) -> str | None:
    """Return a hint if the latest cloudflared log shows HTTPS origin while app uses HTTP."""
    if use_https and wait_for_local_server(local_port, timeout=2.0, use_https=True):
        return None
    origin = cloudflared_logged_hymn_origin(local_port)
    if not origin:
        return None
    if not use_https:
        if origin.lower().startswith("http://127.0.0.1"):
            return None
        return _origin_setup_hint(local_port)
    if origin.lower().startswith("https://127.0.0.1") or origin.lower().startswith("https://localhost"):
        return None
    return None


def diagnose_tunnel_probe(
    url: str,
    local_port: int = DEFAULT_PORT,
    timeout: float = 8.0,
    use_https: bool = False,
    tunnel_mode: str = "service",
) -> str | None:
    """Return a specific failure hint, or None if the tunnel health endpoint looks OK."""
    host = _hostname_from_url(url)
    if not host:
        return "Invalid tunnel URL"

    path = _health_path(url)
    status, body = _probe_request(host, path, timeout)
    if status in (200, 401, 302):
        return None

    body_l = body.lower()
    cf_code = "1033" if "1033" in body or "error code: 1033" in body_l else str(status or "")
    is_quick = tunnel_mode == "quick" or "trycloudflare.com" in host.lower()
    if status in (530, 502, 504) or "1033" in body or "error code: 1033" in body_l:
        if not use_https:
            if not wait_for_local_server(local_port, timeout=2.0, use_https=False):
                return (
                    "本機 HTTP API 未啟動。請勾選「啟用 API」後再試。"
                )
            if is_quick:
                return (
                    f"本機 HTTP 正常，但快速通道回 error {cf_code}。"
                    "請關閉後重新啟用「隨機網址」取得新連結；"
                    "詳情見 exe 旁 cloudflared-tunnel.log。"
                    "（快速通道無需設定 Zero Trust Public Hostname。）"
                )
            mismatch = diagnose_cloudflared_origin_mismatch(local_port, use_https=False)
            if mismatch:
                return mismatch
            return (
                f"本機 HTTP 正常，但固定外網回 error {cf_code}。"
                f"{_origin_setup_hint(local_port, tunnel_url_path_prefix(url) or '/hymn_search')}"
                "儲存後以管理員執行：net stop cloudflared && net start cloudflared。"
            )
        if use_https and wait_for_local_server(local_port, timeout=2.0, use_https=True):
            if status == 502 and _localhost_ipv6_unreachable(local_port, use_https=True):
                return (
                    f"Cloudflare error 502：本機 HTTPS 正常，但 cloudflared 連 localhost 時會走 IPv6 (::1)，"
                    f"而 API 未聽 IPv6。{_https_origin_cloudflare_hint(local_port)}"
                )
            if status == 502:
                return (
                    f"Cloudflare error 502：tunnel 已通，但 cloudflared 無法讀取 HTTPS 自簽 origin。"
                    f"{_https_origin_cloudflare_hint(local_port)}"
                )
            return (
                f"Cloudflare error {cf_code}：本機 HTTPS 已就緒。"
                f"{_https_origin_cloudflare_hint(local_port)}"
            )
        hint = detect_https_origin_misconfiguration(local_port, use_https=use_https)
        if hint:
            return f"Cloudflare 無法連接本機 API（error 1033）。{hint}"
        if is_quick:
            return (
                f"快速通道無法連接本機 API（error {cf_code or '1033'}）。"
                "請確認已啟用 API，然後重新啟用隨機網址。"
            )
        return (
            f"Cloudflare 無法連接本機 API（error 1033）。"
            f"{_origin_setup_hint(local_port, tunnel_url_path_prefix(url) or '/hymn_search')}"
        )
    if status == 403 and ("1010" in body or "error code: 1010" in body_l):
        return (
            "Cloudflare 封鎖自動健康檢查（error 1010）。"
            "若手機瀏覽器可開啟，可忽略；否則請檢查 Tunnel Public Hostname 設定。"
        )
    if status is None:
        host = _hostname_from_url(url)
        if host and not public_dns_resolves(host):
            if is_quick:
                return (
                    f"DNS 無法解析 {host}。"
                    "請換網路／用手機流量再試，或重新啟用「隨機網址」。"
                )
            return (
                f"DNS 無法解析 {host}（域名不存在）。"
                "請在 Cloudflare → churchofgodtm.com → DNS 新增 CNAME："
                "live → <tunnel-id>.cfargotunnel.com（隧道 ID 見 Zero Trust → Tunnels），"
                "或重新在 Tunnel Public Hostname 儲存以自動建立 DNS。"
            )
        return "無法連線至 Cloudflare — 請檢查網路或 DNS"
    return None


def probe_tunnel_url(url: str, timeout: float = 8.0) -> bool:
    host = _hostname_from_url(url)
    if not host:
        return False

    path = _health_path(url)
    if _probe_https_host(host, path, timeout):
        return True

    for dns in _PUBLIC_DNS:
        for ip in _resolve_via_dns(host, dns):
            if _probe_https_ip(ip, host, path, timeout):
                return True
    return False


_SERVICE_QUERY_CACHE: tuple[float, int, str] | None = None
_SERVICE_QUERY_TTL_S = 20.0


def _query_cloudflared_service(force: bool = False) -> tuple[int, str]:
    global _SERVICE_QUERY_CACHE
    if not force and _SERVICE_QUERY_CACHE is not None:
        ts, code, output = _SERVICE_QUERY_CACHE
        if time.time() - ts < _SERVICE_QUERY_TTL_S:
            return code, output
    if not sys.platform.startswith("win"):
        return 1, ""
    flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    try:
        proc = subprocess.run(
            ["sc", "query", "cloudflared"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=flags,
        )
    except (OSError, subprocess.TimeoutExpired):
        result = (1, "")
        _SERVICE_QUERY_CACHE = (time.time(), *result)
        return result
    result = (proc.returncode, f"{proc.stdout}\n{proc.stderr}")
    _SERVICE_QUERY_CACHE = (time.time(), *result)
    return result


def invalidate_cloudflared_service_cache() -> None:
    global _SERVICE_QUERY_CACHE
    _SERVICE_QUERY_CACHE = None


def is_cloudflared_service_installed(force: bool = False) -> bool:
    """Return True if the cloudflared Windows service is registered."""
    code, output = _query_cloudflared_service(force=force)
    return code == 0 and "STATE" in output


def is_cloudflared_service_running(force: bool = False) -> bool:
    """Return True if the cloudflared Windows service is in RUNNING state."""
    code, output = _query_cloudflared_service(force=force)
    if code != 0:
        return False
    return re.search(r"STATE\s+:\s+\d+\s+RUNNING", output) is not None


def _cloudflared_exe_for_command(cloudflared_path: str = "") -> str:
    exe = find_cloudflared(cloudflared_path) or str(cloudflared_path or "").strip() or "cloudflared"
    if " " in exe:
        return f'"{exe}"'
    return exe


def cloudflared_service_install_command(token: str, cloudflared_path: str = "") -> str:
    """Admin CMD: install Windows service (token from settings.json)."""
    exe = _cloudflared_exe_for_command(cloudflared_path)
    tok = str(token or "").strip()
    if not tok:
        return f"{exe} service install <cloudflared_token in settings.json>"
    return f"{exe} service install {tok}"


def cloudflared_service_uninstall_command(cloudflared_path: str = "") -> str:
    """Admin CMD: remove cloudflared Windows service."""
    exe = _cloudflared_exe_for_command(cloudflared_path)
    return f"{exe} service uninstall"


def cloudflared_service_restart_command() -> str:
    """Admin CMD: reload tunnel config after Cloudflare dashboard changes."""
    return "net stop cloudflared && net start cloudflared"


def shared_tunnel_setup_hint(local_port: int = DEFAULT_PORT) -> str:
    return (
        "一部電腦只能有一個 cloudflared Windows 服務。"
        "Cloudflare Zero Trust → Tunnels → Public Hostname："
        f"live.churchofgodtm.com /hymn_search → http://127.0.0.1:{local_port}"
        "（勿用 localhost）；DNS 須有 live 子域名 CNAME。"
        "pptlive 用 /pptlive → http://127.0.0.1:<port>。"
        "本程式選「共用 Windows 服務」。"
        "改 Cloudflare 後以管理員執行 net stop cloudflared && net start cloudflared。"
    )


def install_cloudflared_service(token: str, cloudflared_path: str = "") -> tuple[bool, str]:
    ok, err = validate_run_token(token)
    if not ok:
        return False, err

    if is_cloudflared_service_installed():
        return False, shared_tunnel_setup_hint()

    cloudflared = find_cloudflared(cloudflared_path)
    if not cloudflared:
        return False, "cloudflared not found"

    flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    try:
        proc = subprocess.run(
            [cloudflared, "service", "install", token],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=flags,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, f"Service install failed: {exc}"

    output = f"{proc.stdout}\n{proc.stderr}".strip()
    if proc.returncode == 0:
        invalidate_cloudflared_service_cache()
        append_tunnel_log("INFO", "cloudflared Windows service installed", detail=output[:500])
        return True, output or "cloudflared service installed"
    if "access is denied" in output.lower() or "administrator" in output.lower():
        append_tunnel_log("ERROR", "cloudflared service install denied", detail=output[:500])
        return False, "Run Hymn Search as Administrator to install cloudflared service"
    append_tunnel_log("ERROR", "cloudflared service install failed", detail=output[:500])
    return False, output or f"Service install failed (exit {proc.returncode})"


class CloudflareTunnel:
    """Cloudflare tunnel: quick (trycloudflare), named token, or Windows service."""

    def __init__(
        self,
        local_port: int = DEFAULT_PORT,
        cloudflared_path: str = "",
        tunnel_mode: str = "quick",
        cloudflared_token: str = "",
        named_tunnel_url: str = "",
        use_https: bool = False,
        on_url: Callable[[str], None] | None = None,
        on_error: Callable[[str], None] | None = None,
        on_ready: Callable[[str], None] | None = None,
        on_expired: Callable[[str], None] | None = None,
    ) -> None:
        self.local_port = local_port
        self.cloudflared_path = cloudflared_path
        self.tunnel_mode = tunnel_mode if tunnel_mode in ("quick", "named", "service") else "quick"
        self.cloudflared_token = cloudflared_token.strip()
        self.named_tunnel_url = normalize_tunnel_url(named_tunnel_url)
        self.use_https = bool(use_https)
        self._on_url = on_url
        self._on_error = on_error
        self._on_ready = on_ready
        self._on_expired = on_expired
        self._proc: subprocess.Popen[str] | None = None
        self._log_handle: TextIO | None = None
        self._log_path: Path | None = None
        self._watcher: threading.Thread | None = None
        self._health: threading.Thread | None = None
        self._stop = threading.Event()
        self._active = False
        self.public_url: str | None = None
        self.is_ready = False
        self.local_dns_ok = True

    def _log_event(self, level: str, message: str, **fields: str) -> None:
        log_fields = {
            "mode": self.tunnel_mode,
            "port": str(self.local_port),
            "url": self.public_url or self.named_tunnel_url,
        }
        log_fields.update(fields)
        append_tunnel_log(level, message, **log_fields)

    def is_running(self) -> bool:
        if self.tunnel_mode == "service":
            return self._active and not self._stop.is_set()
        return self._proc is not None and self._proc.poll() is None

    def start(self) -> None:
        if self._active:
            return

        self._stop.clear()
        self._active = True
        self.public_url = None
        self.is_ready = False
        self.local_dns_ok = True

        if not wait_for_local_server(self.local_port, use_https=self.use_https):
            scheme = "https" if self.use_https else "http"
            self._emit_error(f"Local server not ready on {scheme}://127.0.0.1:{self.local_port}")
            return

        self._log_event(
            "INFO",
            "tunnel start",
            named_url=self.named_tunnel_url,
            service_installed=str(is_cloudflared_service_installed()),
            service_running=str(is_cloudflared_service_running()),
            use_https=str(self.use_https),
        )

        if self.tunnel_mode == "service":
            self._start_service_mode()
            return
        if self.tunnel_mode == "named":
            self._start_named_mode()
            return
        self._start_quick_mode()

    def _start_service_mode(self) -> None:
        if not self.named_tunnel_url:
            self._emit_error("Named tunnel URL is required for service mode")
            return
        if not is_cloudflared_service_installed():
            self._emit_error(
                "cloudflared Windows 服務未安裝。請按「安裝 cloudflared 服務」或以管理員執行。"
            )
            return
        if not is_cloudflared_service_running():
            self._emit_error(
                "cloudflared Windows 服務未運行。請在 services.msc 啟動 cloudflared，或重新安裝服務。"
            )
            return
        if not wait_for_local_server(self.local_port, use_https=self.use_https):
            scheme = "https" if self.use_https else "http"
            self._emit_error(f"本機 API 未就緒：{scheme}://127.0.0.1:{self.local_port}")
            return
        self.public_url = self.named_tunnel_url
        self._emit_url(self.public_url)
        threading.Thread(
            target=self._verify_tunnel_ready,
            args=(self.public_url,),
            daemon=True,
        ).start()

    def _read_log_tail(self) -> str:
        if not self._log_path or not self._log_path.exists():
            return ""
        try:
            return self._log_path.read_text(encoding="utf-8", errors="replace").strip()
        except OSError:
            return ""

    def _startup_failure_message(self) -> str:
        log = self._read_log_tail().lower()
        if "token is not valid" in log:
            return (
                "Cloudflare tunnel token is invalid. Copy the JWT from "
                "Zero Trust → Networks → Tunnels → Install connector (starts with eyJ...)."
            )
        if "already connected" in log or "already running" in log:
            return (
                "cloudflared tunnel already running (Windows service?). "
                "Try tunnel mode: Windows service."
            )
        if _HTTPS_ORIGIN_RE.search(self._read_log_tail().replace("\\", "")):
            return detect_https_origin_misconfiguration(self.local_port, use_https=self.use_https) or (
                "Cloudflare origin 設為 https — 請改為 "
                + _origin_setup_hint(self.local_port)
            )
        tail = self._read_log_tail()
        if tail:
            last = tail.splitlines()[-1].strip()
            if last:
                return f"Cloudflare tunnel failed: {last}"
        return (
            "Cloudflare tunnel stopped before becoming reachable — "
            "check tunnel.log and cloudflared-tunnel.log"
        )

    def _watch_named_exit(self) -> None:
        if not self._proc:
            return
        self._proc.wait()
        if self._stop.is_set() or self.is_ready:
            return
        self._emit_error(self._startup_failure_message())

    def _start_named_mode(self) -> None:
        ok, err = validate_run_token(self.cloudflared_token)
        if not ok:
            self._emit_error(err)
            return
        if not self.named_tunnel_url:
            self._emit_error("Named tunnel URL is required (e.g. https://hymn.example.com)")
            return

        cloudflared = find_cloudflared(self.cloudflared_path)
        if not cloudflared:
            self._emit_error(
                "cloudflared not found. Set cloudflared_path in settings.json or reinstall."
            )
            return

        self.public_url = self.named_tunnel_url
        self._emit_url(self.public_url)

        self._log_path = cloudflared_stderr_log_path()
        _append_cloudflared_session_header("named", self.local_port, self.public_url or "")
        try:
            self._log_handle = open(self._log_path, "a", encoding="utf-8")
        except OSError as exc:
            self._emit_error(f"Cannot write tunnel log: {exc}")
            return

        flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        try:
            self._proc = subprocess.Popen(
                [
                    cloudflared,
                    "tunnel",
                    "--no-autoupdate",
                    "run",
                    "--token",
                    self.cloudflared_token,
                ],
                stdout=subprocess.DEVNULL,
                stderr=self._log_handle,
                creationflags=flags,
            )
        except OSError as exc:
            self._close_log()
            self._emit_error(f"Failed to start named cloudflared tunnel: {exc}")
            return

        logger.info("Named Cloudflare tunnel starting with %s", cloudflared)
        threading.Thread(target=self._watch_named_exit, daemon=True).start()
        threading.Thread(
            target=self._verify_tunnel_ready,
            args=(self.public_url,),
            daemon=True,
        ).start()

    def _start_quick_mode(self) -> None:
        cloudflared = find_cloudflared(self.cloudflared_path)
        if not cloudflared:
            self._emit_error(
                "cloudflared not found. Set cloudflared_path in settings.json or reinstall."
            )
            return

        self._log_path = cloudflared_stderr_log_path()
        _append_cloudflared_session_header("quick", self.local_port)
        try:
            self._log_handle = open(self._log_path, "a", encoding="utf-8")
        except OSError as exc:
            self._emit_error(f"Cannot write tunnel log: {exc}")
            return

        flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        try:
            self._proc = subprocess.Popen(
                [
                    cloudflared,
                    "tunnel",
                    "--no-autoupdate",
                    "--url",
                    f"http://127.0.0.1:{self.local_port}",
                ],
                stdout=subprocess.DEVNULL,
                stderr=self._log_handle,
                creationflags=flags,
            )
        except OSError as exc:
            self._close_log()
            self._emit_error(f"Failed to start cloudflared: {exc}")
            return

        self._watcher = threading.Thread(target=self._watch_log, daemon=True)
        self._watcher.start()
        logger.info("Cloudflare tunnel starting with %s on port %s", cloudflared, self.local_port)

    def _emit_error(self, msg: str) -> None:
        logger.error(msg)
        if not self.is_ready:
            self._active = False
        self._log_event("ERROR", msg)
        if self._on_error:
            self._on_error(msg)

    def _emit_expired(self, msg: str) -> None:
        logger.warning(msg)
        self.is_ready = False
        self._log_event("WARN", msg)
        if self._on_expired:
            self._on_expired(msg)

    def _emit_url(self, url: str) -> None:
        self._log_event("INFO", "tunnel url assigned", url=url)
        if self._on_url:
            self._on_url(url)

    def _emit_ready(self, url: str) -> None:
        self._log_event(
            "INFO",
            "tunnel ready",
            url=url,
            local_dns_ok=str(self.local_dns_ok),
        )
        if self._on_ready:
            self._on_ready(url)

    def _close_log(self) -> None:
        if self._log_handle:
            try:
                self._log_handle.close()
            except Exception:
                pass
            self._log_handle = None

    def _watch_log(self) -> None:
        # Only scan new lines — previous sessions leave trycloudflare.com URLs in the log.
        last_pos = 0
        if self._log_path and self._log_path.exists():
            try:
                last_pos = self._log_path.stat().st_size
            except OSError:
                last_pos = 0
        deadline = time.time() + 60
        while not self._stop.is_set() and self.is_running() and time.time() < deadline:
            if self._log_path and self._log_path.exists():
                try:
                    with open(self._log_path, encoding="utf-8", errors="replace") as log_file:
                        log_file.seek(last_pos)
                        chunk = log_file.read()
                        last_pos = log_file.tell()
                    for line in chunk.splitlines():
                        match = _TUNNEL_URL_RE.search(line)
                        if match and not self.public_url:
                            self.public_url = match.group(0)
                            logger.info("Tunnel URL: %s", self.public_url)
                            self._emit_url(self.public_url)
                            threading.Thread(
                                target=self._verify_tunnel_ready,
                                args=(self.public_url,),
                                daemon=True,
                            ).start()
                            return
                except OSError as exc:
                    logger.debug("Tunnel log read: %s", exc)
            time.sleep(0.4)

        if self._stop.is_set():
            return
        if not self.public_url:
            self._emit_error("Tunnel URL not received — check tunnel.log and cloudflared-tunnel.log")
        elif self._proc and self._proc.poll() not in (None, 0):
            self._emit_error("Cloudflare tunnel exited before URL was ready")

    def _verify_tunnel_ready(self, url: str) -> None:
        url = normalize_tunnel_url(url)
        deadline = time.time() + 60
        while time.time() < deadline and not self._stop.is_set():
            if self.tunnel_mode != "service" and not self.is_running():
                if not self._stop.is_set():
                    self._emit_error(self._startup_failure_message())
                return
            host = _hostname_from_url(url)
            if host and probe_tunnel_url(url):
                self.local_dns_ok = local_dns_resolves(host)
                self.is_ready = True
                if not self.local_dns_ok:
                    logger.warning(
                        "Tunnel reachable via public DNS but local DNS cannot resolve %s",
                        host,
                    )
                else:
                    logger.info("Tunnel reachable: %s", url)
                self._emit_ready(url)
                self._health = threading.Thread(target=self._health_loop, daemon=True)
                self._health.start()
                return
            time.sleep(2)

        if self._stop.is_set():
            return
        # Quick tunnels have no Zero Trust Public Hostname — keep the hint mode-aware.
        if self.tunnel_mode == "quick":
            hint = diagnose_tunnel_probe(
                url,
                self.local_port,
                use_https=self.use_https,
                tunnel_mode="quick",
            )
            self._emit_error(
                hint
                or (
                    "快速通道網址未能就緒（請稍候再試，或檢查 exe 旁 cloudflared-tunnel.log）。"
                    "若本機已裝 cloudflared Windows 服務，可先停服務再試快速通道。"
                )
            )
            return
        hint = diagnose_tunnel_probe(
            url,
            self.local_port,
            use_https=self.use_https,
            tunnel_mode=self.tunnel_mode,
        )
        origin_hint = detect_https_origin_misconfiguration(self.local_port, use_https=self.use_https)
        if origin_hint:
            self._log_event("ERROR", origin_hint, probe=hint or "")
        if hint:
            self._emit_error(hint)
        else:
            self._emit_error("Tunnel URL not reachable — check tunnel.log, firewall, or restart session")

    def _health_loop(self) -> None:
        while not self._stop.wait(_HEALTH_INTERVAL_S):
            if self._stop.is_set():
                return
            if self.tunnel_mode != "service" and not self.is_running():
                self._emit_expired("Tunnel stopped — URL expired. Stop and Start Session again.")
                return
            if self.public_url and not probe_tunnel_url(self.public_url):
                if self.tunnel_mode == "service":
                    self._emit_expired("Named tunnel URL unreachable — check cloudflared service.")
                else:
                    self._emit_expired("Tunnel URL unreachable — restart session to get a new link.")

    def stop(self) -> None:
        self._stop.set()
        self._active = False
        if self._proc:
            try:
                self._proc.terminate()
                self._proc.wait(timeout=_SHUTDOWN_JOIN_TIMEOUT_S)
            except Exception:
                try:
                    self._proc.kill()
                except Exception:
                    pass
        self._proc = None
        self._close_log()
        self.public_url = None
        self.is_ready = False
        self.local_dns_ok = True
        self._log_event("INFO", "tunnel stop")
        logger.info("Cloudflare tunnel stopped")
