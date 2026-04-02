"""Shared daemon utilities."""

from __future__ import annotations

import os
import shutil
import socket
from functools import cache
from pathlib import Path
from sys import platform

from headless_word.daemon.config import PID_FILE, PORT_FILE, Config, get_daemon_port
from headless_word.errors import DaemonError, LibreOfficeNotFoundError

_INSTALL_INSTRUCTIONS = {
    "darwin": "brew install --cask libreoffice",
    "linux": "sudo apt install libreoffice libreoffice-writer",
    "win32": "Download from https://www.libreoffice.org/download/",
}

_WINDOWS_SOFFICE_PATHS = [
    Path(os.environ.get("PROGRAMFILES", "C:\\Program Files"))
    / "LibreOffice"
    / "program"
    / "soffice.com",
    Path(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"))
    / "LibreOffice"
    / "program"
    / "soffice.com",
    Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local")))
    / "Programs"
    / "LibreOffice"
    / "program"
    / "soffice.com",
]


def cleanup_daemon_files() -> None:
    PID_FILE.unlink(missing_ok=True)
    PORT_FILE.unlink(missing_ok=True)


@cache
def get_soffice_path() -> str | None:
    if platform == "win32":
        for path in _WINDOWS_SOFFICE_PATHS:
            if path.exists():
                return str(path)
        found = shutil.which("soffice.com") or shutil.which("soffice")
        return found
    if platform == "darwin":
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(mac_path):
            return mac_path
    return shutil.which("soffice")


def ensure_libreoffice_installed() -> None:
    if get_soffice_path():
        return
    key = (
        "win32"
        if platform == "win32"
        else ("darwin" if platform == "darwin" else "linux")
    )
    instructions = _INSTALL_INSTRUCTIONS.get(key, "Install LibreOffice")
    raise LibreOfficeNotFoundError(
        f"LibreOffice not found. Install it:\n  {instructions}"
    )


def send_daemon_command(cmd: str, timeout: float | None = None) -> str:
    port = get_daemon_port()
    if port is None:
        raise DaemonError("Daemon port file not found")

    config = Config.from_env()
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout or config.socket_timeout)
    try:
        sock.connect((config.daemon_host, port))
        sock.sendall((cmd + "\n").encode("utf-8"))
        chunks = []
        while True:
            try:
                chunk = sock.recv(1048576)
                if not chunk:
                    break
                chunks.append(chunk)
            except TimeoutError:
                break
        return b"".join(chunks).decode("utf-8").strip()
    except ConnectionRefusedError as err:
        raise DaemonError("Daemon is not running (connection refused)") from err
    finally:
        sock.close()


def is_daemon_running(cleanup_if_stale: bool = True) -> bool:
    port = get_daemon_port()
    if port is None:
        if cleanup_if_stale:
            cleanup_daemon_files()
        return False

    try:
        resp = send_daemon_command("PING", timeout=3)
        return resp == "PONG"
    except (DaemonError, OSError):
        if cleanup_if_stale:
            cleanup_daemon_files()
        return False
