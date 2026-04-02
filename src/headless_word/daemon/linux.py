"""Linux daemon: system Python + UNO socket connection."""

import contextlib
import os
import signal
import subprocess
import time
from pathlib import Path

from headless_word.daemon.base import (
    PID_FILE,
    cleanup_daemon_files,
    is_daemon_running,
    send_daemon_command,
)
from headless_word.daemon.config import Config, find_free_port, write_daemon_port
from headless_word.errors import DaemonError


def _get_helper_script_path() -> Path:
    return Path(__file__).parent / "linux_helper.py"


def _check_uno_available() -> bool:
    try:
        result = subprocess.run(
            ["/usr/bin/python3", "-c", "import uno"],
            capture_output=True,
            timeout=5,
        )
        return result.returncode == 0
    except Exception:
        return False


def start_daemon_linux(wait: bool, timeout: float) -> int:
    if not _check_uno_available():
        raise DaemonError(
            "python3-uno is required on Linux.\n"
            "Install it with: sudo apt install python3-uno"
        )

    helper_path = _get_helper_script_path()
    if not helper_path.exists():
        raise DaemonError(f"Helper script not found: {helper_path}")

    config = Config.from_env()
    port = find_free_port(config.daemon_port)

    env = os.environ.copy()
    env["HEADLESS_WORD_PORT"] = str(port)
    env["HEADLESS_WORD_IDLE_TIMEOUT"] = str(config.idle_timeout)

    proc = subprocess.Popen(
        ["/usr/bin/python3", str(helper_path)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
        env=env,
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
        stop_daemon_linux()
        raise DaemonError(f"Daemon failed to start within {timeout}s")

    return proc.pid


def stop_daemon_linux() -> bool:
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
            os.killpg(pid, signal.SIGTERM)
            stopped = True
        except (ProcessLookupError, ValueError, PermissionError):
            pass

    cleanup_daemon_files()

    with contextlib.suppress(Exception):
        subprocess.run(
            ["pkill", "-f", "accept=socket.*2002"],
            capture_output=True,
            timeout=5,
        )

    return stopped
