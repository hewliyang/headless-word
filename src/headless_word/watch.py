"""Live Word document viewer with auto-reload."""

import asyncio
import base64
import contextlib
import json
import os
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

from headless_word.client import WordClient

VIEWER_HTML = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{filename} - headless-word</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            background: #1e1e1e;
            font-family: system-ui, -apple-system, sans-serif;
            color: #ccc;
        }}
        #toolbar {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 100;
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 8px 16px;
            background: #252526;
            border-bottom: 1px solid #3c3c3c;
            font-size: 13px;
        }}
        #toolbar .file {{ color: #dcdcaa; font-family: 'SF Mono', Consolas, monospace; font-size: 12px; }}
        #toolbar .status {{ color: #4ec9b0; font-size: 11px; }}
        #toolbar .status.loading {{ color: #e8ab53; }}
        #toolbar .page-info {{ color: #808080; font-size: 12px; margin-left: auto; }}
        .nav-btn {{
            background: #3c3c3c;
            color: #ccc;
            border: 1px solid #555;
            padding: 4px 12px;
            border-radius: 3px;
            cursor: pointer;
            font-size: 12px;
        }}
        .nav-btn:hover {{ background: #4c4c4c; }}
        .nav-btn:disabled {{ opacity: 0.3; cursor: default; }}
        #pages {{
            padding: 52px 0 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 16px;
        }}
        .page-container {{
            background: white;
            box-shadow: 0 2px 12px rgba(0,0,0,0.5);
            position: relative;
        }}
        .page-container img {{
            display: block;
            max-width: 900px;
            width: 100%;
            height: auto;
        }}
        .page-label {{
            position: absolute;
            top: -22px;
            left: 0;
            color: #666;
            font-size: 11px;
        }}
        #loading {{
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            color: #666;
            font-size: 14px;
        }}
    </style>
</head>
<body>
    <div id="toolbar">
        <span class="file">{filename}</span>
        <span id="status" class="status">connected</span>
        <span class="page-info"><span id="page-count">-</span> pages</span>
    </div>
    <div id="pages">
        <div id="loading">Loading document...</div>
    </div>

    <script>
        let ws;
        let reconnectTimer;

        function connect() {{
            ws = new WebSocket('ws://localhost:{ws_port}');
            ws.onopen = () => {{
                document.getElementById('status').textContent = 'connected';
                document.getElementById('status').className = 'status';
                clearTimeout(reconnectTimer);
            }};
            ws.onmessage = (e) => {{
                const data = JSON.parse(e.data);
                if (data.type === 'update') {{
                    renderPages(data.pages);
                }} else if (data.type === 'refreshing') {{
                    document.getElementById('status').textContent = 'refreshing...';
                    document.getElementById('status').className = 'status loading';
                }}
            }};
            ws.onclose = () => {{
                document.getElementById('status').textContent = 'disconnected';
                document.getElementById('status').className = 'status loading';
                reconnectTimer = setTimeout(connect, 2000);
            }};
        }}

        function renderPages(pages) {{
            const container = document.getElementById('pages');
            container.innerHTML = '';
            document.getElementById('page-count').textContent = pages.length;
            document.getElementById('status').textContent = 'connected';
            document.getElementById('status').className = 'status';

            pages.forEach((page, i) => {{
                const div = document.createElement('div');
                div.className = 'page-container';

                const label = document.createElement('div');
                label.className = 'page-label';
                label.textContent = `Page ${{i + 1}}`;

                const img = document.createElement('img');
                img.src = 'data:image/png;base64,' + page.data;
                img.alt = `Page ${{i + 1}}`;

                div.appendChild(label);
                div.appendChild(img);
                container.appendChild(div);
            }});
        }}

        connect();
    </script>
</body>
</html>"""


def _render_all_pages(client: WordClient, sid: str, dpi: int = 150) -> list[dict]:
    """Render all pages to base64 PNG."""
    # Get page count
    struct = client.get_document_structure(sid)
    page_count = struct.page_count or 1

    pages = []
    for p in range(1, page_count + 1):
        result = client.screenshot(sid, page=p, dpi=dpi)
        with open(result.png_path, "rb") as f:
            data = base64.b64encode(f.read()).decode()
        pages.append({"page": p, "data": data, "size": result.size_bytes})
        # Clean up the temp file
        with contextlib.suppress(Exception):
            os.unlink(result.png_path)

    return pages


def create_handler(html_content: str):
    class Handler(SimpleHTTPRequestHandler):
        def do_GET(self):
            if self.path == "/" or self.path == "/index.html":
                content = html_content.encode()
                self.send_response(200)
                self.send_header("Content-Type", "text/html")
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_error(404)

        def log_message(self, format, *args):
            pass  # Suppress HTTP logs

    return Handler


async def watch(
    file_path: str,
    session_id: str | None = None,
    http_port: int = 8080,
    ws_port: int = 8765,
    dpi: int = 150,
    open_browser: bool = True,
):
    """Start live document viewer."""
    try:
        import websockets
    except ImportError as err:
        raise ImportError(
            "Missing dependency: websockets. Install with: pip install websockets"
        ) from err

    file_path = str(Path(file_path).absolute())
    filename = Path(file_path).name
    client = WordClient()

    html = VIEWER_HTML.format(filename=filename, ws_port=ws_port)
    connected_clients: set = set()
    last_mtime: float = 0
    watch_sid: str | None = None  # watcher's own read-only session

    def _open_watch_session() -> str:
        """Open a read-only session for rendering. Separate from any editing session."""
        nonlocal watch_sid
        if watch_sid:
            with contextlib.suppress(Exception):
                client.close(watch_sid)
        watch_sid = client.open(file_path)
        return watch_sid

    def _close_watch_session():
        nonlocal watch_sid
        if watch_sid:
            with contextlib.suppress(Exception):
                client.close(watch_sid)
            watch_sid = None

    # Initial render
    print(f"Rendering {filename}...")
    sid = _open_watch_session()
    pages = _render_all_pages(client, sid, dpi)
    _close_watch_session()
    print(f"  {len(pages)} pages rendered")

    async def ws_handler(websocket):
        connected_clients.add(websocket)
        try:
            # Send current state
            await websocket.send(json.dumps({"type": "update", "pages": pages}))
            async for _ in websocket:
                pass  # Keep alive
        finally:
            connected_clients.discard(websocket)

    async def broadcast(data: dict):
        if connected_clients:
            msg = json.dumps(data)
            await asyncio.gather(
                *[c.send(msg) for c in connected_clients],
                return_exceptions=True,
            )

    async def file_watcher():
        nonlocal pages, last_mtime
        last_mtime = os.path.getmtime(file_path)

        while True:
            await asyncio.sleep(1)
            try:
                mtime = os.path.getmtime(file_path)
                if mtime > last_mtime:
                    last_mtime = mtime
                    print("  File changed, re-rendering...")
                    await broadcast({"type": "refreshing"})

                    # Open, render, close — don't hold the file
                    sid = _open_watch_session()
                    pages = _render_all_pages(client, sid, dpi)
                    _close_watch_session()
                    print(f"  {len(pages)} pages rendered")
                    await broadcast({"type": "update", "pages": pages})
            except Exception as e:
                print(f"  Watch error: {e}")

    # Start HTTP server in a thread
    handler_class = create_handler(html)
    httpd = HTTPServer(("127.0.0.1", http_port), handler_class)
    http_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    http_thread.start()

    url = f"http://localhost:{http_port}"
    print(f"  Viewer: {url}")
    print(f"  Watching: {file_path}")
    print("  Press Ctrl+C to stop\n")

    if open_browser:
        webbrowser.open(url)

    # Start WebSocket server and file watcher
    async with websockets.serve(ws_handler, "127.0.0.1", ws_port):
        await file_watcher()
