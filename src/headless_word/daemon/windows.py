"""Windows daemon: Python macro inside LibreOffice."""

import os
import subprocess
import time
from pathlib import Path

from headless_word.daemon.base import (
    PID_FILE,
    cleanup_daemon_files,
    get_soffice_path,
    is_daemon_running,
    send_daemon_command,
)
from headless_word.daemon.config import Config, find_free_port, write_daemon_port
from headless_word.daemon.macro import get_word_macro
from headless_word.errors import DaemonError


def _get_python_macro_dir() -> Path:
    appdata = os.environ.get("APPDATA", str(Path.home() / "AppData/Roaming"))
    return Path(appdata) / "LibreOffice/4/user/Scripts/python"


def _install_daemon_macro(config: Config) -> None:
    macro_dir = _get_python_macro_dir()
    macro_dir.mkdir(parents=True, exist_ok=True)
    macro_file = macro_dir / "wordbridge.py"
    macro_file.write_text(get_word_macro(config))


def start_daemon_windows(wait: bool, timeout: float) -> int:
    soffice = get_soffice_path()
    if not soffice:
        raise DaemonError("LibreOffice soffice not found")

    config = Config.from_env()
    port = find_free_port(config.daemon_port)
    config = Config(daemon_port=port, idle_timeout=config.idle_timeout)

    _install_daemon_macro(config)

    cmd = [
        soffice,
        "--headless",
        "--invisible",
        "--nologo",
        "--norestore",
        "vnd.sun.star.script:wordbridge.py$start_server?language=Python&location=user",
    ]

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )

    PID_FILE.parent.mkdir(parents=True, exist_ok=True)
    PID_FILE.write_text(str(proc.pid))
    write_daemon_port(port)

    if wait:
        start_time = time.time()
        while time.time() - start_time < timeout:
            if is_daemon_running(cleanup_if_stale=False):
                return proc.pid
            time.sleep(0.2)
        stop_daemon_windows()
        raise DaemonError(f"Daemon failed to start within {timeout}s")

    return proc.pid


def stop_daemon_windows() -> bool:
    stopped = False

    if is_daemon_running():
        try:
            send_daemon_command("QUIT", timeout=5)
            stopped = True
            time.sleep(0.5)
        except DaemonError:
            pass

    if PID_FILE.exists():
        try:
            pid = int(PID_FILE.read_text().strip())
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                capture_output=True,
                timeout=5,
            )
            stopped = True
        except Exception:
            pass

    cleanup_daemon_files()
    return stopped
