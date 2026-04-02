"""Generate the Python macro that runs inside LibreOffice.

The macro starts a TCP server, manages document sessions,
and implements all Word tools as named handlers.
"""

from __future__ import annotations

from headless_word.daemon.config import Config


def get_word_macro(config: Config | None = None) -> str:
    if config is None:
        config = Config.from_env()

    return f'''\
"""TCP bridge for headless-word: Writer document tools daemon."""
import base64
import json
import os
import shutil
import socket
import subprocess
import tempfile
import time
import traceback
import uuid
import zipfile
from xml.dom.minidom import parseString as parse_xml_string

import uno
from com.sun.star.beans import PropertyValue

DAEMON_PORT = {config.daemon_port}
IDLE_TIMEOUT = {config.idle_timeout}

sessions = {{}}  # session_id -> (doc, filepath, work_dir)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

NOISE_ATTRS = [
    "rsidR", "rsidRDefault", "rsidRPr", "rsidP",
    "rsidDel", "rsidSect", "rsidTr",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def recv_all(conn, initial_data=""):
    data = initial_data
    while True:
        if "\\n" in data:
            return data.split("\\n")[0]
        try:
            chunk = conn.recv(65536).decode("utf-8")
            if not chunk:
                break
            data += chunk
        except socket.timeout:
            break
    return data.strip()


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def file_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))


def get_desktop():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def get_paragraphs(doc):
    text = doc.getText()
    enum = text.createEnumeration()
    paras = []
    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            paras.append(el)
    return paras


def get_page_count(doc):
    try:
        return doc.getCurrentController().getPropertyValue("PageCount")
    except Exception:
        return None


def save_tmp_docx(doc, work_dir):
    tmp = os.path.join(work_dir, f"_tmp_{{int(time.time() * 1000)}}.docx")
    doc.storeToURL(
        file_url(tmp),
        (make_prop("FilterName", "MS Word 2007 XML"), make_prop("Overwrite", True)),
    )
    return tmp


def clean_element(el):
    for attr in NOISE_ATTRS:
        try:
            el.removeAttributeNS(W_NS, attr)
        except Exception:
            pass
        try:
            if el.hasAttribute(f"w:{{attr}}"):
                el.removeAttribute(f"w:{{attr}}")
        except Exception:
            pass
    try:
        el.removeAttributeNS(W14_NS, "paraId")
    except Exception:
        pass
    try:
        el.removeAttributeNS(W14_NS, "textId")
    except Exception:
        pass
    for child in list(el.childNodes):
        if child.nodeType == 1:
            clean_element(child)


def get_text_content(el):
    texts = []
    for t in el.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild:
            texts.append(t.firstChild.nodeValue or "")
    return "".join(texts)


def pretty_xml(xml_str):
    formatted = xml_str.replace("><", ">\\n<")
    lines = formatted.split("\\n")
    depth = 0
    result = []
    for line in lines:
        trimmed = line.strip()
        if not trimmed:
            continue
        is_closing = trimmed.startswith("</")
        is_self_closing = trimmed.endswith("/>")
        is_opening = trimmed.startswith("<") and not is_closing and not is_self_closing
        if is_closing:
            depth = max(0, depth - 1)
        result.append("  " * depth + trimmed)
        if is_opening:
            depth += 1
    return "\\n".join(result)


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

def tool_get_document_text(doc, work_dir, params):
    start = params.get("start_paragraph", 0)
    end_p = params.get("end_paragraph")
    include_fmt = params.get("include_formatting", True)

    paras = get_paragraphs(doc)
    total = len(paras)
    end = min(end_p, total) if end_p is not None else total

    results = []
    for i in range(start, end):
        p = paras[i]
        info = {{"index": i, "text": p.getString()}}
        if include_fmt:
            info["style"] = p.getPropertyValue("ParaStyleName")
            try:
                align_val = p.getPropertyValue("ParaAdjust")
                align_map = {{0: "Left", 1: "Right", 2: "Justified", 3: "Center"}}
                info["alignment"] = align_map.get(align_val, str(align_val))
            except Exception:
                pass
            try:
                rules = p.getPropertyValue("NumberingRules")
                level = p.getPropertyValue("NumberingLevel")
                if rules and level >= 0:
                    info["list_level"] = level
                    try:
                        info["list_string"] = p.getPropertyValue("ListLabelString")
                    except Exception:
                        pass
            except Exception:
                pass
        results.append(info)

    return {{
        "total_paragraphs": total,
        "showing": {{"start": start, "end": end}},
        "paragraphs": results,
    }}


def tool_get_document_structure(doc, work_dir, params):
    text = doc.getText()
    enum = text.createEnumeration()
    headings = []
    tables = []
    para_count = 0
    table_idx = 0

    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            outline_level = el.getPropertyValue("OutlineLevel")
            if 1 <= outline_level <= 9:
                headings.append({{
                    "text": el.getString()[:120],
                    "level": outline_level,
                    "paragraph_index": para_count,
                }})
            para_count += 1
        elif el.supportsService("com.sun.star.text.TextTable"):
            rows = el.getRows().getCount()
            cols = el.getColumns().getCount()
            style = None
            try:
                style = el.getPropertyValue("TableTemplateName") or None
            except Exception:
                pass
            tables.append({{
                "index": table_idx,
                "rows": rows,
                "columns": cols,
                "style": style,
            }})
            table_idx += 1

    section_count = 1
    try:
        section_count = max(1, doc.getTextSections().getCount())
    except Exception:
        pass

    cc_info = []
    try:
        fenum = doc.getTextFields().createEnumeration()
        cc_idx = 0
        while fenum.hasMoreElements():
            f = fenum.nextElement()
            if f.supportsService("com.sun.star.text.textfield.InputField"):
                cc_info.append({{
                    "id": cc_idx,
                    "title": f.getPropertyValue("Hint") or "",
                    "tag": "",
                    "type": "InputField",
                }})
                cc_idx += 1
    except Exception:
        pass

    return {{
        "paragraph_count": para_count,
        "section_count": section_count,
        "table_count": table_idx,
        "content_control_count": len(cc_info),
        "page_count": get_page_count(doc),
        "headings": headings,
        "tables": tables,
        "content_controls": cc_info,
    }}


def _extract_body_content(docx_path, start_child=None, end_child=None):
    with zipfile.ZipFile(docx_path) as z:
        raw = z.read("word/document.xml")
        try:
            raw_styles = z.read("word/styles.xml")
        except KeyError:
            raw_styles = None
        try:
            raw_numbering = z.read("word/numbering.xml")
        except KeyError:
            raw_numbering = None

    dom = parse_xml_string(raw)
    body_els = dom.getElementsByTagNameNS(W_NS, "body")
    if not body_els:
        return "", [], None, None
    body = body_els[0]

    all_elements = [c for c in body.childNodes if c.nodeType == 1]
    s = start_child if start_child is not None else 0
    e = end_child if end_child is not None else len(all_elements) - 1

    # Collect referenced style IDs
    style_ids = set()
    for tag in ("pStyle", "rStyle", "tblStyle"):
        for el in body.getElementsByTagNameNS(W_NS, tag):
            val = el.getAttributeNS(W_NS, "val") or el.getAttribute("w:val")
            if val:
                style_ids.add(val)

    style_xml = None
    if raw_styles and style_ids:
        styles_dom = parse_xml_string(raw_styles)
        style_map = {{}}
        for el in styles_dom.getElementsByTagNameNS(W_NS, "style"):
            sid = el.getAttributeNS(W_NS, "styleId") or el.getAttribute("w:styleId")
            if sid:
                style_map[sid] = el
        to_include = set(style_ids)
        for sid in style_ids:
            el = style_map.get(sid)
            if el:
                for based_on in el.getElementsByTagNameNS(W_NS, "basedOn"):
                    base = based_on.getAttributeNS(W_NS, "val") or based_on.getAttribute("w:val")
                    if base:
                        to_include.add(base)
        parts = []
        for dd in styles_dom.getElementsByTagNameNS(W_NS, "docDefaults"):
            parts.append(dd.toxml())
        for sid in to_include:
            el = style_map.get(sid)
            if el:
                clean_element(el)
                parts.append(el.toxml())
        if parts:
            style_xml = "\\n".join(parts)

    numbering_xml = None
    num_ids = set()
    for el in body.getElementsByTagNameNS(W_NS, "numId"):
        val = el.getAttributeNS(W_NS, "val") or el.getAttribute("w:val")
        if val and val != "0":
            num_ids.add(val)

    if raw_numbering and num_ids:
        num_dom = parse_xml_string(raw_numbering)
        parts = []
        abstract_ids = set()
        for num_el in num_dom.getElementsByTagNameNS(W_NS, "num"):
            nid = num_el.getAttributeNS(W_NS, "numId") or num_el.getAttribute("w:numId")
            if nid and nid in num_ids:
                parts.append(num_el.toxml())
                for abs_ref in num_el.getElementsByTagNameNS(W_NS, "abstractNumId"):
                    aid = abs_ref.getAttributeNS(W_NS, "val") or abs_ref.getAttribute("w:val")
                    if aid:
                        abstract_ids.add(aid)
        for abs_el in num_dom.getElementsByTagNameNS(W_NS, "abstractNum"):
            aid = abs_el.getAttributeNS(W_NS, "abstractNumId") or abs_el.getAttribute("w:abstractNumId")
            if aid and aid in abstract_ids:
                parts.insert(0, abs_el.toxml())
        if parts:
            numbering_xml = "\\n".join(parts)

    output_parts = []
    children = []
    line_offset = 1
    para_offset = 0
    table_idx = 0

    for i, el in enumerate(all_elements):
        tag = el.localName
        clean_element(el)
        p_count = len(el.getElementsByTagNameNS(W_NS, "p")) if tag != "p" else 1

        if i < s or i > e:
            if tag == "tbl":
                table_idx += 1
            para_offset += p_count
            continue

        summary = {{"index": i, "type": tag, "line": line_offset}}

        if tag == "tbl":
            rows = el.getElementsByTagNameNS(W_NS, "tr")
            first_row = rows[0] if rows else None
            cols = len(first_row.getElementsByTagNameNS(W_NS, "tc")) if first_row else 0
            label = f"table ({{len(rows)}} rows x {{cols}} cols)"
            summary["table_index"] = table_idx
            summary["rows"] = len(rows)
            summary["cols"] = cols
            summary["paragraph_range"] = [para_offset, para_offset + p_count - 1]
            table_idx += 1
        elif tag == "p":
            txt = get_text_content(el)
            summary["paragraph_index"] = para_offset
            if txt:
                summary["text"] = txt[:80]
                label = f'paragraph: "{{txt[:80]}}"'
            else:
                label = "paragraph (empty)"
        elif tag == "sdt":
            alias_els = el.getElementsByTagNameNS(W_NS, "alias")
            title = ""
            if alias_els:
                title = alias_els[0].getAttributeNS(W_NS, "val") or alias_els[0].getAttribute("w:val") or ""
            label = f"sdt: {{title}}" if title else "sdt"
            summary["paragraph_range"] = [para_offset, para_offset + p_count - 1]
        elif tag == "sectPr":
            label = "sectPr"
        else:
            label = tag

        para_offset += p_count
        children.append(summary)

        raw_xml = el.toxml()
        pretty = pretty_xml(raw_xml)
        comment_line = f"<!-- Body child {{i}}: {{label}} -->"
        block = f"{{comment_line}}\\n{{pretty}}"
        block_lines = len(block.split("\\n"))
        output_parts.append(block)
        line_offset += block_lines + 1

    xml_content = "\\n\\n".join(output_parts)
    return xml_content, children, style_xml, numbering_xml


def tool_get_ooxml(doc, work_dir, params):
    start_child = params.get("start_child")
    end_child = params.get("end_child")

    tmp = save_tmp_docx(doc, work_dir)
    try:
        xml_content, children, style_xml, numbering_xml = _extract_body_content(
            tmp, start_child, end_child
        )
    finally:
        os.unlink(tmp)

    file_parts = []
    if style_xml:
        file_parts.append(f"<!-- Referenced styles -->\\n{{pretty_xml(style_xml)}}")
    if numbering_xml:
        file_parts.append(f"<!-- Numbering definitions -->\\n{{pretty_xml(numbering_xml)}}")
    file_parts.append(xml_content)
    file_content = "\\n\\n".join(file_parts)

    if style_xml or numbering_xml:
        header_lines = len(file_content.split("\\n")) - len(xml_content.split("\\n"))
        for child in children:
            child["line"] += header_lines

    range_label = ""
    if start_child is not None or end_child is not None:
        range_label = f"-{{start_child or 0}}-{{end_child or 'end'}}"
    out_path = os.path.join(work_dir, f"body{{range_label}}.xml")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(file_content)

    lines = len(file_content.split("\\n"))
    size_kb = round(len(file_content) / 1024)

    return {{
        "file": out_path,
        "size": f"{{size_kb}}KB",
        "lines": lines,
        "children": children,
    }}


def tool_screenshot(doc, work_dir, params):
    page = params.get("page", 1)
    dpi = params.get("dpi", 200)
    show_comments = params.get("show_comments", True)
    show_changes = params.get("show_changes", True)

    page_count = get_page_count(doc)
    if page_count and page > page_count:
        raise ValueError(f"Page {{page}} out of range (1-{{page_count}})")

    pdf_path = os.path.join(work_dir, f"_page_{{page}}.pdf")
    fd_props = [make_prop("PageRange", str(page))]
    if show_comments:
        fd_props.append(make_prop("ExportNotesInMargin", True))
    if show_changes:
        fd_props.append(make_prop("Changes", 1))
    filter_data = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        tuple(fd_props),
    )
    doc.storeToURL(
        file_url(pdf_path),
        (
            make_prop("FilterName", "writer_pdf_Export"),
            make_prop("Overwrite", True),
            make_prop("FilterData", filter_data),
        ),
    )

    png_path = os.path.join(work_dir, f"page_{{page}}.png")
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("pdftoppm not found. Install poppler-utils.")

    subprocess.run(
        [pdftoppm, "-png", "-r", str(dpi), "-singlefile", pdf_path, png_path.replace(".png", "")],
        capture_output=True,
        check=True,
    )
    os.unlink(pdf_path)

    size = os.path.getsize(png_path)
    return {{
        "page": page,
        "page_count": page_count or 0,
        "png_path": png_path,
        "size_bytes": size,
    }}


def tool_insert_ooxml(doc, work_dir, params):
    desktop = get_desktop()
    ooxml = params["ooxml"]
    position = params.get("position", "end")

    tmp_path = save_tmp_docx(doc, work_dir)
    with zipfile.ZipFile(tmp_path, "r") as z:
        all_files = {{name: z.read(name) for name in z.namelist()}}

    dom = parse_xml_string(all_files["word/document.xml"])
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    elements = [c for c in body.childNodes if c.nodeType == 1]

    wrapper = f'<w:body xmlns:w="{{W_NS}}">{{ooxml}}</w:body>'
    snippet_dom = parse_xml_string(wrapper)
    snippet_body = snippet_dom.getElementsByTagNameNS(W_NS, "body")[0]
    new_nodes = [c for c in snippet_body.childNodes if c.nodeType == 1]

    if position == "start":
        ref_node = elements[0] if elements else None
    elif position.startswith("after:"):
        idx = int(position.split(":")[1])
        if idx >= len(elements):
            raise ValueError(f"Body child index {{idx}} out of range (0-{{len(elements) - 1}})")
        ref_node = elements[idx].nextSibling
        while ref_node and ref_node.nodeType != 1:
            ref_node = ref_node.nextSibling
    else:
        sect_prs = body.getElementsByTagNameNS(W_NS, "sectPr")
        ref_node = sect_prs[0] if sect_prs else None

    for node in new_nodes:
        imported = dom.importNode(node, True)
        if ref_node:
            body.insertBefore(imported, ref_node)
        else:
            body.appendChild(imported)

    all_files["word/document.xml"] = dom.toxml().encode("utf-8")
    out_path = tmp_path + ".new.docx"
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in all_files.items():
            z.writestr(name, data)

    doc.close(True)
    new_doc = desktop.loadComponentFromURL(file_url(out_path), "_blank", 0, ())
    os.unlink(tmp_path)

    return {{"success": True, "new_doc": id(new_doc), "_doc": new_doc, "_path": out_path}}


def tool_execute(doc, work_dir, params):
    code = params["code"]
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    text = doc.getText()
    cursor = text.createTextCursor()
    cursor.gotoEnd(False)

    exec_ns = {{
        "doc": doc,
        "text": text,
        "cursor": cursor,
        "desktop": desktop,
        "uno": uno,
        "smgr": smgr,
        "ctx": ctx,
        "prop": make_prop,
    }}
    exec(code, exec_ns)
    result = exec_ns.get("result", None)
    return {{"success": True, "result": result}}


TOOLS = {{
    "get_document_text": tool_get_document_text,
    "get_document_structure": tool_get_document_structure,
    "get_ooxml": tool_get_ooxml,
    "screenshot": tool_screenshot,
    "insert_ooxml": tool_insert_ooxml,
    "execute": tool_execute,
}}


# ---------------------------------------------------------------------------
# Command handler
# ---------------------------------------------------------------------------

def handle_command(cmd, desktop):
    try:
        if cmd == "PING":
            return "PONG"

        if cmd == "QUIT":
            for sid, (doc, fp, wd) in list(sessions.items()):
                try:
                    doc.close(False)
                except Exception:
                    pass
                try:
                    shutil.rmtree(wd, ignore_errors=True)
                except Exception:
                    pass
            sessions.clear()
            return "QUIT_OK"

        if cmd == "LIST":
            info = {{}}
            for sid, (doc, fp, wd) in sessions.items():
                info[sid] = {{"path": fp}}
            return f"RESULT:{{json.dumps(info)}}"

        if cmd.startswith("OPEN:"):
            filepath = cmd[5:]
            if not os.path.exists(filepath):
                return f"ERROR:File not found: {{filepath}}"
            doc = desktop.loadComponentFromURL(file_url(filepath), "_blank", 0, ())
            if not doc:
                return f"ERROR:Failed to open: {{filepath}}"
            sid = str(uuid.uuid4())[:8]
            work_dir = tempfile.mkdtemp(prefix=f"hw_{{sid}}_")
            sessions[sid] = (doc, os.path.abspath(filepath), work_dir)
            return f"SESSION:{{sid}}"

        if cmd.startswith("NEW:"):
            filepath = cmd[4:]
            doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
            if not doc:
                return "ERROR:Failed to create new document"
            sid = str(uuid.uuid4())[:8]
            work_dir = tempfile.mkdtemp(prefix=f"hw_{{sid}}_")
            sessions[sid] = (doc, os.path.abspath(filepath), work_dir)
            return f"SESSION:{{sid}}"

        if cmd.startswith("SAVE:"):
            parts = cmd[5:].split(":", 1)
            sid = parts[0]
            if sid not in sessions:
                return "ERROR:Invalid session"
            doc, filepath, work_dir = sessions[sid]
            save_path = parts[1] if len(parts) > 1 else filepath
            doc.storeToURL(
                file_url(save_path),
                (make_prop("FilterName", "MS Word 2007 XML"), make_prop("Overwrite", True)),
            )
            sessions[sid] = (doc, os.path.abspath(save_path), work_dir)
            return f"OK:{{os.path.abspath(save_path)}}"

        if cmd.startswith("EXPORT_PDF:"):
            parts = cmd[11:].split(":", 2)
            if len(parts) < 2:
                return "ERROR:Expected EXPORT_PDF:session_id:path[:options_json]"
            sid, pdf_path = parts[0], parts[1]
            opts = json.loads(parts[2]) if len(parts) > 2 else {{}}
            if sid not in sessions:
                return "ERROR:Invalid session"
            doc, filepath, work_dir = sessions[sid]

            fd_props = []
            if opts.get("show_comments", True):
                fd_props.append(make_prop("ExportNotesInMargin", True))
            if opts.get("show_changes", True):
                fd_props.append(make_prop("Changes", 1))
            if "page_range" in opts:
                fd_props.append(make_prop("PageRange", opts["page_range"]))

            filter_data = uno.Any(
                "[]com.sun.star.beans.PropertyValue",
                tuple(fd_props),
            ) if fd_props else ()

            export_props = [
                make_prop("FilterName", "writer_pdf_Export"),
                make_prop("Overwrite", True),
            ]
            if fd_props:
                export_props.append(make_prop("FilterData", filter_data))

            doc.storeToURL(file_url(pdf_path), tuple(export_props))
            return f"OK:{{os.path.abspath(pdf_path)}}"

        if cmd.startswith("CLOSE:"):
            sid = cmd[6:]
            if sid not in sessions:
                return "ERROR:Invalid session"
            doc, filepath, work_dir = sessions.pop(sid)
            try:
                doc.close(True)
            except Exception:
                pass
            try:
                shutil.rmtree(work_dir, ignore_errors=True)
            except Exception:
                pass
            return "OK"

        if cmd.startswith("TOOL:"):
            parts = cmd[5:].split(":", 2)
            if len(parts) < 2:
                return "ERROR:Expected TOOL:session_id:tool_name[:b64_params]"
            sid = parts[0]
            tool_name = parts[1]
            params_b64 = parts[2] if len(parts) > 2 else ""

            if sid not in sessions:
                return "ERROR:Invalid session"
            doc, filepath, work_dir = sessions[sid]

            if tool_name not in TOOLS:
                return f"ERROR:Unknown tool: {{tool_name}}"

            params = json.loads(base64.b64decode(params_b64).decode()) if params_b64 else {{}}

            try:
                result = TOOLS[tool_name](doc, work_dir, params)
                # Handle insert_ooxml special case (replaces doc)
                if tool_name == "insert_ooxml" and result.get("_doc"):
                    new_doc = result.pop("_doc")
                    new_path = result.pop("_path")
                    sessions[sid] = (new_doc, new_path, work_dir)
                    result.pop("new_doc", None)
                return f"RESULT:{{json.dumps(result)}}"
            except Exception:
                return f"TOOL_ERROR:{{traceback.format_exc()}}"

        if cmd.startswith("EXEC:"):
            parts = cmd[5:].split(":", 1)
            if len(parts) != 2:
                return "ERROR:Expected EXEC:session_id:b64_code"
            sid, code_b64 = parts
            if sid not in sessions:
                return "ERROR:Invalid session"
            doc, filepath, work_dir = sessions[sid]

            try:
                code = base64.b64decode(code_b64).decode()
            except Exception as e:
                return f"ERROR:Failed to decode code: {{e}}"

            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager

            exec_ns = {{
                "doc": doc,
                "text": doc.getText(),
                "cursor": doc.getText().createTextCursor(),
                "desktop": desktop,
                "uno": uno,
                "smgr": smgr,
                "ctx": ctx,
                "prop": make_prop,
            }}
            exec_ns["cursor"].gotoEnd(False)

            try:
                exec(code, exec_ns)
                result = exec_ns.get("result", None)
                return f"RESULT:{{json.dumps(result)}}"
            except Exception:
                return f"EXEC_ERROR:{{traceback.format_exc()}}"

        return "ERROR:Unknown command"

    except Exception:
        return f"ERROR:{{traceback.format_exc()}}"


# ---------------------------------------------------------------------------
# TCP server entry point (called by LibreOffice as a macro)
# ---------------------------------------------------------------------------

def start_server(*args):
    desktop = get_desktop()

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(("127.0.0.1", DAEMON_PORT))
    server.listen(5)
    server.settimeout(60)

    last_activity = time.time()

    while True:
        try:
            conn, addr = server.accept()
            last_activity = time.time()
            conn.settimeout(30)

            initial = conn.recv(65536).decode("utf-8")
            data = recv_all(conn, initial).strip()

            response = handle_command(data, desktop)

            if response == "QUIT_OK":
                conn.send(b"OK")
                conn.close()
                server.close()
                desktop.terminate()
                break

            conn.sendall(response.encode("utf-8"))
            conn.close()
        except socket.timeout:
            if time.time() - last_activity > IDLE_TIMEOUT:
                for sid, (doc, _, wd) in list(sessions.items()):
                    try:
                        doc.close(False)
                    except Exception:
                        pass
                    try:
                        shutil.rmtree(wd, ignore_errors=True)
                    except Exception:
                        pass
                sessions.clear()
                server.close()
                desktop.terminate()
                break
            continue
        except Exception:
            break


g_exportedScripts = (start_server,)
'''
