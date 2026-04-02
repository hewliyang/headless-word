from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]


class CliResult(subprocess.CompletedProcess[str]):
    pass


@pytest.fixture(scope="session")
def cli_env(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    home = tmp_path_factory.mktemp("home")
    env = os.environ.copy()
    env["HOME"] = str(home)
    env.setdefault("HEADLESS_WORD_IDLE_TIMEOUT", "120")
    pythonpath = str(PROJECT_ROOT / "src")
    if env.get("PYTHONPATH"):
        env["PYTHONPATH"] = pythonpath + os.pathsep + env["PYTHONPATH"]
    else:
        env["PYTHONPATH"] = pythonpath
    return env


@pytest.fixture(scope="session", autouse=True)
def libreoffice_daemon(cli_env: dict[str, str]) -> None:
    result = run_cli("start", "--timeout", "30", env=cli_env, check=False)
    if result.returncode != 0:
        raise RuntimeError(
            f"Failed to start headless-word daemon\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )

    yield

    run_cli("stop", env=cli_env, check=False)


@pytest.fixture
def docx_path(tmp_path: Path) -> Path:
    return tmp_path / "integration.docx"


@pytest.fixture
def populated_session(docx_path: Path, cli_env: dict[str, str]) -> dict[str, str]:
    new_result = run_cli("new", str(docx_path), env=cli_env)
    session_id = json.loads(new_result.stdout)["session_id"]

    code = """
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

cursor.gotoEnd(False)
cursor.setPropertyValue("ParaStyleName", "Heading 1")
text.insertString(cursor, "Quarterly Report", False)
text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
cursor.setPropertyValue("ParaStyleName", "Standard")
text.insertString(cursor, "Revenue increased 10% year over year.", False)
result = "ok"
""".strip()

    run_cli("execute", session_id, env=cli_env, input=code)

    yield {"session_id": session_id, "path": str(docx_path)}

    run_cli("close", session_id, env=cli_env, check=False)


def run_cli(
    *args: str,
    env: dict[str, str] | None = None,
    input: str | None = None,
    timeout: int = 60,
    check: bool = True,
) -> CliResult:
    result = subprocess.run(
        [sys.executable, "-m", "headless_word.cli", *args],
        cwd=PROJECT_ROOT,
        env=env,
        input=input,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if check and result.returncode != 0:
        raise AssertionError(
            f"Command failed: {' '.join(args)}\n"
            f"exit={result.returncode}\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )
    return result  # type: ignore[return-value]
