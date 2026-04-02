"""Cross-platform daemon API."""

from sys import platform

from headless_word.daemon.base import (
    PID_FILE,
    ensure_libreoffice_installed,
    is_daemon_running,
)


def start_daemon(wait: bool = True, timeout: float = 15) -> int:
    ensure_libreoffice_installed()

    if is_daemon_running():
        if PID_FILE.exists():
            return int(PID_FILE.read_text().strip())
        return -1

    if platform.startswith("linux"):
        from headless_word.daemon.linux import start_daemon_linux

        return start_daemon_linux(wait, timeout)
    elif platform == "win32":
        from headless_word.daemon.windows import start_daemon_windows

        return start_daemon_windows(wait, timeout)
    else:
        from headless_word.daemon.macos import start_daemon_macos

        return start_daemon_macos(wait, timeout)


def stop_daemon() -> bool:
    if platform.startswith("linux"):
        from headless_word.daemon.linux import stop_daemon_linux

        return stop_daemon_linux()
    elif platform == "win32":
        from headless_word.daemon.windows import stop_daemon_windows

        return stop_daemon_windows()
    else:
        from headless_word.daemon.macos import stop_daemon_macos

        return stop_daemon_macos()
