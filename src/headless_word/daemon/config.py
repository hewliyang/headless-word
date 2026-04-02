from __future__ import annotations

import os
import random
import socket
from dataclasses import dataclass

from headless_word.constants import HEADLESS_WORD_DIR

PRIVATE_PORT_START = 49152
PRIVATE_PORT_END = 65535

PID_FILE = HEADLESS_WORD_DIR / "daemon.pid"
PORT_FILE = HEADLESS_WORD_DIR / "daemon.port"


@dataclass(frozen=True)
class Config:
    daemon_host: str = "127.0.0.1"
    daemon_port: int = PRIVATE_PORT_START
    uno_port: int = 2002
    socket_timeout: int = 30
    idle_timeout: int = 300

    @classmethod
    def from_env(cls) -> Config:
        return cls(
            daemon_port=int(
                os.environ.get("HEADLESS_WORD_PORT", str(PRIVATE_PORT_START))
            ),
            idle_timeout=int(os.environ.get("HEADLESS_WORD_IDLE_TIMEOUT", "300")),
        )


def find_free_port(preferred: int | None = None, max_attempts: int = 100) -> int:
    def try_bind(port: int) -> bool:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port))
                return True
        except OSError:
            return False

    if (
        preferred is not None
        and PRIVATE_PORT_START <= preferred <= PRIVATE_PORT_END
        and try_bind(preferred)
    ):
        return preferred

    for _ in range(max_attempts):
        port = random.randint(PRIVATE_PORT_START, PRIVATE_PORT_END)
        if try_bind(port):
            return port

    raise RuntimeError(f"No free port found after {max_attempts} attempts")


def get_daemon_port() -> int | None:
    if not PORT_FILE.exists():
        return None
    try:
        return int(PORT_FILE.read_text().strip())
    except (ValueError, OSError):
        return None


def write_daemon_port(port: int) -> None:
    PORT_FILE.parent.mkdir(parents=True, exist_ok=True)
    PORT_FILE.write_text(str(port))
