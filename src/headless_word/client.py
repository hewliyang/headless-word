"""Typed client for the headless-word daemon."""

from __future__ import annotations

import base64
import json
from typing import Any

from headless_word.daemon.base import send_daemon_command
from headless_word.errors import DaemonError, SessionError, ToolError
from headless_word.models import (
    ExecuteResult,
    GetDocumentStructureResult,
    GetDocumentTextResult,
    GetOoxmlResult,
    InsertOoxmlResult,
    ScreenshotDocumentResult,
)


class WordClient:
    def _send(self, cmd: str, timeout: float | None = None) -> str:
        return send_daemon_command(cmd, timeout=timeout)

    def _tool(
        self,
        sid: str,
        tool_name: str,
        params: dict | None = None,
        timeout: float | None = None,
    ) -> dict:
        if params:
            params_b64 = base64.b64encode(json.dumps(params).encode()).decode()
            cmd = f"TOOL:{sid}:{tool_name}:{params_b64}"
        else:
            cmd = f"TOOL:{sid}:{tool_name}"

        resp = self._send(cmd, timeout=timeout)

        if resp.startswith("RESULT:"):
            return json.loads(resp[7:])
        if resp.startswith("TOOL_ERROR:"):
            raise ToolError(resp[11:])
        if resp.startswith("ERROR:"):
            raise SessionError(resp[6:])
        raise DaemonError(f"Unexpected response: {resp[:200]}")

    # Session management

    def open(self, path: str) -> str:
        resp = self._send(f"OPEN:{path}")
        if resp.startswith("SESSION:"):
            return resp[8:]
        raise SessionError(resp)

    def new(self, path: str) -> str:
        resp = self._send(f"NEW:{path}")
        if resp.startswith("SESSION:"):
            return resp[8:]
        raise SessionError(resp)

    def save(self, sid: str, path: str | None = None) -> str:
        cmd = f"SAVE:{sid}:{path}" if path else f"SAVE:{sid}"
        resp = self._send(cmd)
        if resp.startswith("OK:"):
            return resp[3:]
        raise SessionError(resp)

    def close(self, sid: str) -> None:
        resp = self._send(f"CLOSE:{sid}")
        if resp != "OK":
            raise SessionError(resp)

    def export_pdf(self, sid: str, path: str, **opts: Any) -> str:
        opts_json = json.dumps(opts) if opts else ""
        cmd = f"EXPORT_PDF:{sid}:{path}"
        if opts_json:
            cmd += f":{opts_json}"
        resp = self._send(cmd)
        if resp.startswith("OK:"):
            return resp[3:]
        raise SessionError(resp)

    def list_sessions(self) -> dict:
        resp = self._send("LIST")
        if resp.startswith("RESULT:"):
            return json.loads(resp[7:])
        raise DaemonError(resp)

    # Tools

    def get_document_text(
        self,
        sid: str,
        start_paragraph: int = 0,
        end_paragraph: int | None = None,
        include_formatting: bool = True,
    ) -> GetDocumentTextResult:
        params: dict[str, Any] = {
            "start_paragraph": start_paragraph,
            "include_formatting": include_formatting,
        }
        if end_paragraph is not None:
            params["end_paragraph"] = end_paragraph
        return GetDocumentTextResult(**self._tool(sid, "get_document_text", params))

    def get_document_structure(self, sid: str) -> GetDocumentStructureResult:
        return GetDocumentStructureResult(**self._tool(sid, "get_document_structure"))

    def get_ooxml(
        self,
        sid: str,
        start_child: int | None = None,
        end_child: int | None = None,
    ) -> GetOoxmlResult:
        params: dict[str, Any] = {}
        if start_child is not None:
            params["start_child"] = start_child
        if end_child is not None:
            params["end_child"] = end_child
        return GetOoxmlResult(**self._tool(sid, "get_ooxml", params))

    def screenshot(
        self,
        sid: str,
        page: int = 1,
        dpi: int = 200,
        show_comments: bool = True,
        show_changes: bool = True,
    ) -> ScreenshotDocumentResult:
        params = {
            "page": page,
            "dpi": dpi,
            "show_comments": show_comments,
            "show_changes": show_changes,
        }
        return ScreenshotDocumentResult(
            **self._tool(sid, "screenshot", params, timeout=30)
        )

    def insert_ooxml(
        self,
        sid: str,
        ooxml: str,
        position: str = "end",
    ) -> InsertOoxmlResult:
        params = {"ooxml": ooxml, "position": position}
        return InsertOoxmlResult(**self._tool(sid, "insert_ooxml", params))

    def execute(self, sid: str, code: str) -> ExecuteResult:
        params = {"code": code}
        return ExecuteResult(**self._tool(sid, "execute", params))

    def execute_raw(self, sid: str, code: str) -> Any:
        code_b64 = base64.b64encode(code.encode()).decode()
        resp = self._send(f"EXEC:{sid}:{code_b64}")
        if resp.startswith("RESULT:"):
            return json.loads(resp[7:])
        if resp.startswith("EXEC_ERROR:"):
            raise ToolError(resp[11:])
        if resp.startswith("ERROR:"):
            raise SessionError(resp[6:])
        raise DaemonError(f"Unexpected response: {resp[:200]}")
