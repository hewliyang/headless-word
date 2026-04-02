"""CLI for headless-word."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

from headless_word.client import WordClient
from headless_word.daemon import is_daemon_running, start_daemon, stop_daemon
from headless_word.daemon.base import PID_FILE, get_soffice_path
from headless_word.daemon.config import get_daemon_port
from headless_word.errors import HeadlessWordError
from headless_word.postprocess import fix_comment_threading_with_state

GREEN = "\033[32m"
RED = "\033[31m"
DIM = "\033[2m"
RESET = "\033[0m"


def _ok(msg: str) -> None:
    print(f"{GREEN}✓{RESET} {msg}")


def _fail(msg: str) -> None:
    print(f"{RED}✗{RESET} {msg}", file=sys.stderr)


def _info(msg: str) -> None:
    print(f"  {DIM}{msg}{RESET}")


def _json_out(data: Any) -> None:
    if hasattr(data, "model_dump"):
        print(json.dumps(data.model_dump(), indent=2, default=str))
    else:
        print(json.dumps(data, indent=2, default=str))


def _get_lo_version() -> str | None:
    soffice = get_soffice_path()
    if not soffice:
        return None
    try:
        result = subprocess.run(
            [soffice, "--version"], capture_output=True, text=True, timeout=5
        )
        return result.stdout.strip()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------


def cmd_start(args: argparse.Namespace) -> int:
    if is_daemon_running():
        _ok("Daemon is already running")
        if PID_FILE.exists():
            _info(f"PID: {PID_FILE.read_text().strip()}")
        port = get_daemon_port()
        if port:
            _info(f"Port: {port}")
        return 0

    print("Starting LibreOffice daemon...")
    try:
        pid = start_daemon(wait=True, timeout=args.timeout)
        _ok(f"Daemon started (PID: {pid})")
        port = get_daemon_port()
        if port:
            _info(f"Port: {port}")
        _info("Stop with: headless-word stop")
        return 0
    except Exception as e:
        _fail(f"Failed to start: {e}")
        return 1


def cmd_stop(args: argparse.Namespace) -> int:
    if not is_daemon_running():
        print("Daemon is not running")
        stop_daemon()
        return 0

    print("Stopping daemon...")
    if stop_daemon():
        _ok("Daemon stopped")
        return 0
    else:
        _fail("Failed to stop daemon")
        return 1


def cmd_status(args: argparse.Namespace) -> int:
    if is_daemon_running():
        _ok("Daemon is running")
        if PID_FILE.exists():
            _info(f"PID: {PID_FILE.read_text().strip()}")
        port = get_daemon_port()
        if port:
            _info(f"Port: {port}")
        return 0
    else:
        print("Daemon is not running")
        return 1


def cmd_check(args: argparse.Namespace) -> int:
    print("headless-word environment check\n")

    soffice = get_soffice_path()
    if soffice:
        _ok(f"LibreOffice found: {soffice}")
        version = _get_lo_version()
        if version:
            _info(version)
    else:
        _fail("LibreOffice not found")
        if sys.platform == "darwin":
            _info("Install: brew install --cask libreoffice")
        elif sys.platform == "win32":
            _info("Download from https://www.libreoffice.org/download/")
        else:
            _info("Install: sudo apt install libreoffice libreoffice-writer")
        return 1

    print()
    if is_daemon_running():
        _ok("Daemon is running")
    else:
        _info("Daemon is not running (start with: headless-word start)")

    print()
    print("All checks passed!" if soffice else "Some checks failed.")
    return 0


def cmd_open(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        sid = client.open(str(Path(args.path).absolute()))
        _json_out({"session_id": sid, "path": str(Path(args.path).absolute())})
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_new(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        sid = client.new(str(Path(args.path).absolute()))
        _json_out({"session_id": sid, "path": str(Path(args.path).absolute())})
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def _get_comment_threading(client: WordClient, sid: str) -> list[dict]:
    """Query comment threading state from the daemon before save."""
    try:
        result = client.execute(
            sid,
            code="""
fields = doc.getTextFields().createEnumeration()
anns = []
while fields.hasMoreElements():
    f = fields.nextElement()
    if f.supportsService("com.sun.star.text.textfield.Annotation"):
        anns.append({
            "ParaId": f.getPropertyValue("ParaId"),
            "ParaIdParent": f.getPropertyValue("ParaIdParent"),
            "Resolved": f.getPropertyValue("Resolved"),
            "Author": f.getPropertyValue("Author"),
            "Content": f.getPropertyValue("Content")[:20],
        })
result = anns
""",
        )
        if not result.success or not result.result:
            return []
        anns = result.result
        has_threading = any(a["ParaIdParent"] != "0" or a["Resolved"] for a in anns)
        if not has_threading:
            return []
        return anns
    except Exception:
        return []


def cmd_save(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        # Query threading state before save (IDs may shift on save)
        threading_state = _get_comment_threading(client, args.session)

        path = client.save(args.session, args.path)

        # Post-process if there are threaded/resolved comments
        if threading_state:
            _apply_comment_threading(path, threading_state)

        _json_out({"saved": path})
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def _apply_comment_threading(docx_path: str, anns: list[dict]) -> None:
    """Map UNO annotation state to comment IDs and fix threading in docx."""
    import xml.dom.minidom
    import zipfile

    # Read comments.xml to get comment IDs and match by content/author
    try:
        with zipfile.ZipFile(docx_path) as z:
            if "word/comments.xml" not in z.namelist():
                return
            dom = xml.dom.minidom.parseString(z.read("word/comments.xml"))
    except Exception:
        return

    comments = dom.getElementsByTagName("w:comment")

    # Build lookup: (author, content_prefix) -> comment w:id
    xml_comments = []
    for c in comments:
        cid = int(c.getAttribute("w:id"))
        author = c.getAttribute("w:author")
        texts = c.getElementsByTagName("w:t")
        content = "".join(
            t.firstChild.nodeValue
            for t in texts
            if t.firstChild and t.firstChild.nodeValue
        )[:20]
        xml_comments.append({"id": cid, "author": author, "content": content})

    # Match UNO annotations to XML comments by author + content prefix
    # Build a mapping: UNO ParaId -> comment w:id
    uno_to_xml = {}  # UNO ParaId -> xml comment w:id
    used_xml = set()
    for ann in anns:
        for xc in xml_comments:
            if xc["id"] in used_xml:
                continue
            if (
                ann["Author"] == xc["author"]
                and ann["Content"][:15] == xc["content"][:15]
            ):
                uno_to_xml[ann["ParaId"]] = xc["id"]
                used_xml.add(xc["id"])
                break

    # Build threading list for post-processor
    threading = []
    for ann in anns:
        para_id = ann["ParaId"]
        if para_id not in uno_to_xml:
            continue
        cid = uno_to_xml[para_id]
        parent_cid = None
        if ann["ParaIdParent"] != "0" and ann["ParaIdParent"] in uno_to_xml:
            parent_cid = uno_to_xml[ann["ParaIdParent"]]
        threading.append(
            {
                "comment_id": cid,
                "parent_comment_id": parent_cid,
                "resolved": ann.get("Resolved", False),
            }
        )

    if threading:
        fix_comment_threading_with_state(docx_path, threading)


def cmd_close(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        client.close(args.session)
        _ok(f"Session {args.session} closed")
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_list(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        sessions = client.list_sessions()
        _json_out(sessions)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_export_pdf(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        opts = {}
        if args.no_comments:
            opts["show_comments"] = False
        if args.no_changes:
            opts["show_changes"] = False
        if args.pages:
            opts["page_range"] = args.pages
        path = client.export_pdf(args.session, str(Path(args.path).absolute()), **opts)
        _json_out({"pdf": path})
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


# Tool commands


def cmd_get_document_text(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        result = client.get_document_text(
            args.session,
            start_paragraph=args.start or 0,
            end_paragraph=args.end,
            include_formatting=not args.no_formatting,
        )
        _json_out(result)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_get_document_structure(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        result = client.get_document_structure(args.session)
        _json_out(result)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_get_ooxml(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        result = client.get_ooxml(
            args.session,
            start_child=args.start_child,
            end_child=args.end_child,
        )
        _json_out(result)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_screenshot(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        result = client.screenshot(
            args.session,
            page=args.page or 1,
            dpi=args.dpi or 200,
            show_comments=not args.no_comments,
            show_changes=not args.no_changes,
        )
        # If --out specified, copy the file
        if args.out:
            import shutil

            shutil.copy2(result.png_path, args.out)
            result = result.model_copy(
                update={"png_path": str(Path(args.out).absolute())}
            )
        _json_out(result)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_insert_ooxml(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        if args.file:
            ooxml = Path(args.file).read_text(encoding="utf-8")
        elif args.xml:
            ooxml = args.xml
        else:
            if sys.stdin.isatty():
                _fail("Provide OOXML via --xml, --file, or stdin")
                return 1
            ooxml = sys.stdin.read()

        result = client.insert_ooxml(
            args.session, ooxml, position=args.position or "end"
        )
        _json_out(result)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


def cmd_watch(args: argparse.Namespace) -> int:
    import asyncio

    try:
        from headless_word.watch import watch
    except ImportError as e:
        _fail(f"Missing dependency: {e}")
        _info("Install with: pip install headless-word[watch]")
        return 1

    try:
        asyncio.run(
            watch(
                args.path,
                session_id=args.session,
                http_port=args.port,
                ws_port=args.ws_port,
                dpi=args.dpi,
                open_browser=args.open,
            )
        )
    except KeyboardInterrupt:
        print("\nStopped")
    except Exception as e:
        _fail(str(e))
        return 1
    return 0


def cmd_execute(args: argparse.Namespace) -> int:
    client = WordClient()
    try:
        if args.file:
            code = Path(args.file).read_text(encoding="utf-8")
        elif args.code:
            code = args.code
        else:
            if sys.stdin.isatty():
                _fail("Provide code via --code, --file, or stdin")
                return 1
            code = sys.stdin.read()

        if args.raw:
            result = client.execute_raw(args.session, code)
            _json_out({"result": result})
        else:
            result = client.execute(args.session, code)
            _json_out(result)
        return 0
    except HeadlessWordError as e:
        _fail(str(e))
        return 1


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        prog="headless-word",
        description="Word document automation via LibreOffice",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # Daemon commands
    p_start = sub.add_parser("start", help="Start the LibreOffice daemon")
    p_start.add_argument(
        "--timeout", type=float, default=15, help="Startup timeout (seconds)"
    )

    sub.add_parser("stop", help="Stop the daemon")
    sub.add_parser("status", help="Check daemon status")
    sub.add_parser("check", help="Check environment setup")
    sub.add_parser("list", help="List open sessions")

    # Session commands
    p_open = sub.add_parser("open", help="Open a document")
    p_open.add_argument("path", help="Path to .docx file")

    p_new = sub.add_parser("new", help="Create a new document")
    p_new.add_argument("path", help="Path for the new .docx file")

    p_save = sub.add_parser("save", help="Save a document")
    p_save.add_argument("session", help="Session ID")
    p_save.add_argument("--path", help="Save-as path (optional)")

    p_close = sub.add_parser("close", help="Close a session")
    p_close.add_argument("session", help="Session ID")

    p_pdf = sub.add_parser("export-pdf", help="Export as PDF")
    p_pdf.add_argument("session", help="Session ID")
    p_pdf.add_argument("path", help="Output PDF path")
    p_pdf.add_argument("--no-comments", action="store_true")
    p_pdf.add_argument("--no-changes", action="store_true")
    p_pdf.add_argument("--pages", help="Page range (e.g. '1-3')")

    # Tool commands
    p_text = sub.add_parser("get-document-text", help="Get document text")
    p_text.add_argument("session", help="Session ID")
    p_text.add_argument("--start", type=int, help="Start paragraph index")
    p_text.add_argument("--end", type=int, help="End paragraph index (exclusive)")
    p_text.add_argument("--no-formatting", action="store_true")

    p_struct = sub.add_parser("get-document-structure", help="Get document structure")
    p_struct.add_argument("session", help="Session ID")

    p_ooxml = sub.add_parser("get-ooxml", help="Get document OOXML")
    p_ooxml.add_argument("session", help="Session ID")
    p_ooxml.add_argument("--start-child", type=int, help="Start body-child index")
    p_ooxml.add_argument(
        "--end-child", type=int, help="End body-child index (inclusive)"
    )

    p_shot = sub.add_parser("screenshot", help="Screenshot a page")
    p_shot.add_argument("session", help="Session ID")
    p_shot.add_argument("--page", type=int, help="Page number (1-based)")
    p_shot.add_argument("--dpi", type=int, help="DPI (default: 200)")
    p_shot.add_argument("--no-comments", action="store_true")
    p_shot.add_argument("--no-changes", action="store_true")
    p_shot.add_argument("--out", help="Copy PNG to this path")

    p_insert = sub.add_parser("insert-ooxml", help="Insert OOXML into document")
    p_insert.add_argument("session", help="Session ID")
    p_insert.add_argument("--xml", help="OOXML string")
    p_insert.add_argument("--file", help="OOXML file path")
    p_insert.add_argument("--position", help="'end', 'start', or 'after:<index>'")

    p_watch = sub.add_parser("watch", help="Live document viewer with auto-reload")
    p_watch.add_argument("path", help="Path to .docx file")
    p_watch.add_argument("--session", help="Existing session ID (optional)")
    p_watch.add_argument(
        "--port", type=int, default=8080, help="HTTP port (default: 8080)"
    )
    p_watch.add_argument(
        "--ws-port", type=int, default=8765, help="WebSocket port (default: 8765)"
    )
    p_watch.add_argument(
        "--dpi", type=int, default=150, help="Render DPI (default: 150)"
    )
    p_watch.add_argument(
        "--open", action="store_true", help="Open browser automatically"
    )

    p_exec = sub.add_parser("execute", help="Execute Python/UNO code")
    p_exec.add_argument("session", help="Session ID")
    p_exec.add_argument("--code", help="Python code string")
    p_exec.add_argument("--file", help="Python file path")
    p_exec.add_argument(
        "--raw", action="store_true", help="Use raw EXEC protocol (no tool wrapper)"
    )

    args = parser.parse_args()

    handlers = {
        "start": cmd_start,
        "stop": cmd_stop,
        "status": cmd_status,
        "check": cmd_check,
        "list": cmd_list,
        "open": cmd_open,
        "new": cmd_new,
        "save": cmd_save,
        "close": cmd_close,
        "export-pdf": cmd_export_pdf,
        "get-document-text": cmd_get_document_text,
        "get-document-structure": cmd_get_document_structure,
        "get-ooxml": cmd_get_ooxml,
        "screenshot": cmd_screenshot,
        "insert-ooxml": cmd_insert_ooxml,
        "execute": cmd_execute,
        "watch": cmd_watch,
    }

    handler = handlers.get(args.command)
    if handler:
        sys.exit(handler(args))
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
