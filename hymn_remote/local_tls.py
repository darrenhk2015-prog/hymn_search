"""Self-signed localhost TLS material for the remote API (Cloudflare https origin)."""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

CERT_NAME = "remote_local.crt"
KEY_NAME = "remote_local.key"
CLOUDFLARED_CA_NAME = "origin-ca.pem"


def cert_paths(base_dir: str | Path) -> tuple[Path, Path]:
    base = Path(base_dir)
    return base / CERT_NAME, base / KEY_NAME


def cloudflared_origin_ca_path() -> Path:
    base = Path(os.environ.get("ProgramData", r"C:\ProgramData"))
    return base / "cloudflared" / CLOUDFLARED_CA_NAME


def cloudflared_origin_ca_hint() -> str:
    return str(cloudflared_origin_ca_path())


def _find_openssl() -> str | None:
    found = shutil.which("openssl")
    if found:
        return found
    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_name)
        if not base:
            continue
        candidate = Path(base) / "Git" / "usr" / "bin" / "openssl.exe"
        if candidate.is_file():
            return str(candidate)
    return None


def _cert_is_current(cert_path: Path) -> bool:
    try:
        text = cert_path.read_text(encoding="ascii", errors="ignore")
    except OSError:
        return False
    return (
        "CN=127.0.0.1" in text
        and "127.0.0.1" in text
        and "0:0:0:0:0:0:0:1" in text
    )


def _generate_cert(cert_path: Path, key_path: Path) -> None:
    openssl = _find_openssl()
    if not openssl:
        raise RuntimeError(
            "找不到 openssl，無法產生本機 HTTPS 憑證。"
            "請安裝 Git for Windows，或改 Cloudflare origin 為 http://127.0.0.1:<port>。"
        )

    cert_path.parent.mkdir(parents=True, exist_ok=True)
    flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    cmd = [
        openssl,
        "req",
        "-x509",
        "-newkey",
        "rsa:2048",
        "-nodes",
        "-days",
        "825",
        "-keyout",
        str(key_path),
        "-out",
        str(cert_path),
        "-subj",
        "/CN=127.0.0.1",
        "-addext",
        "subjectAltName=DNS:localhost,IP:127.0.0.1,IP:0:0:0:0:0:0:0:1",
    ]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=flags,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"openssl 產生憑證失敗：{exc}") from exc

    output = f"{proc.stdout}\n{proc.stderr}".strip()
    if proc.returncode != 0 or not cert_path.is_file() or not key_path.is_file():
        raise RuntimeError(output or f"openssl 產生憑證失敗（exit {proc.returncode}）")


def ensure_localhost_cert(base_dir: str | Path) -> tuple[Path, Path]:
    """Return (cert, key) paths, generating or refreshing a self-signed cert if needed."""
    cert_path, key_path = cert_paths(base_dir)
    if cert_path.is_file() and key_path.is_file() and _cert_is_current(cert_path):
        return cert_path, key_path

    for path in (cert_path, key_path):
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
    _generate_cert(cert_path, key_path)
    return cert_path, key_path


def deploy_cloudflared_origin_ca(cert_path: str | Path) -> Path | None:
    """Copy the local origin cert where cloudflared Windows service can load it."""
    if not sys.platform.startswith("win"):
        return None
    src = Path(cert_path)
    if not src.is_file():
        return None
    dest = cloudflared_origin_ca_path()
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
    except OSError:
        return None
    return dest
