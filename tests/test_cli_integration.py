from __future__ import annotations

import json
from pathlib import Path

from .conftest import run_cli


def test_check_runs(cli_env: dict[str, str]) -> None:
    result = run_cli("check", env=cli_env, check=False)

    assert result.returncode == 0
    assert "headless-word environment check" in result.stdout
    assert "LibreOffice" in result.stdout


def test_new_execute_save_reopen_and_read(
    populated_session: dict[str, str], cli_env: dict[str, str]
) -> None:
    session_id = populated_session["session_id"]
    path = Path(populated_session["path"])

    text_result = run_cli("get-document-text", session_id, env=cli_env)
    text_data = json.loads(text_result.stdout)
    paragraph_texts = [p["text"] for p in text_data["paragraphs"]]

    assert any("Quarterly Report" in text for text in paragraph_texts)
    assert any("Revenue increased 10%" in text for text in paragraph_texts)

    structure_result = run_cli("get-document-structure", session_id, env=cli_env)
    structure = json.loads(structure_result.stdout)

    assert structure["paragraph_count"] >= 2
    assert structure["page_count"] >= 1
    assert any(h["text"] == "Quarterly Report" for h in structure["headings"])

    save_result = run_cli("save", session_id, env=cli_env)
    saved = Path(json.loads(save_result.stdout)["saved"])

    assert saved == path
    assert path.exists()
    assert path.stat().st_size > 0

    run_cli("close", session_id, env=cli_env, check=False)

    reopen_result = run_cli("open", str(path), env=cli_env)
    reopened_session_id = json.loads(reopen_result.stdout)["session_id"]

    try:
        reopened_text_result = run_cli(
            "get-document-text", reopened_session_id, env=cli_env
        )
        reopened = json.loads(reopened_text_result.stdout)
        reopened_texts = [p["text"] for p in reopened["paragraphs"]]
        assert any("Quarterly Report" in text for text in reopened_texts)
        assert any("Revenue increased 10%" in text for text in reopened_texts)
    finally:
        run_cli("close", reopened_session_id, env=cli_env, check=False)


def test_export_pdf_and_screenshot(
    populated_session: dict[str, str], tmp_path: Path, cli_env: dict[str, str]
) -> None:
    session_id = populated_session["session_id"]
    pdf_path = tmp_path / "document.pdf"
    png_path = tmp_path / "page1.png"

    pdf_result = run_cli("export-pdf", session_id, str(pdf_path), env=cli_env)
    pdf_data = json.loads(pdf_result.stdout)

    assert Path(pdf_data["pdf"]) == pdf_path.resolve()
    assert pdf_path.exists()
    assert pdf_path.stat().st_size > 0

    screenshot_result = run_cli(
        "screenshot", session_id, "--page", "1", "--out", str(png_path), env=cli_env
    )
    screenshot = json.loads(screenshot_result.stdout)

    assert screenshot["page"] == 1
    assert screenshot["page_count"] >= 1
    assert Path(screenshot["png_path"]) == png_path.resolve()
    assert png_path.exists()
    assert png_path.stat().st_size > 0


def test_insert_ooxml_round_trip(docx_path: Path, cli_env: dict[str, str]) -> None:
    new_result = run_cli("new", str(docx_path), env=cli_env)
    session_id = json.loads(new_result.stdout)["session_id"]

    ooxml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:r><w:t>Inserted via OOXML</w:t></w:r>"
        "</w:p>"
    )

    try:
        insert_result = run_cli(
            "insert-ooxml",
            session_id,
            "--xml",
            ooxml,
            env=cli_env,
            timeout=90,
        )
        insert_data = json.loads(insert_result.stdout)
        assert insert_data["success"] is True

        save_result = run_cli("save", session_id, env=cli_env, timeout=90)
        saved_path = Path(json.loads(save_result.stdout)["saved"])
        assert saved_path.exists()
        assert saved_path.stat().st_size > 0

        text_result = run_cli("get-document-text", session_id, env=cli_env)
        text_data = json.loads(text_result.stdout)
        paragraph_texts = [p["text"] for p in text_data["paragraphs"]]
        assert any("Inserted via OOXML" in text for text in paragraph_texts)
    finally:
        run_cli("close", session_id, env=cli_env, check=False)


def test_threaded_comments_are_written_to_docx(
    docx_path: Path, cli_env: dict[str, str]
) -> None:
    new_result = run_cli("new", str(docx_path), env=cli_env)
    session_id = json.loads(new_result.stdout)["session_id"]

    code = """
import datetime


def now_dt():
    n = datetime.datetime.now()
    dt = uno.createUnoStruct("com.sun.star.util.DateTime")
    dt.Year, dt.Month, dt.Day = n.year, n.month, n.day
    dt.Hours, dt.Minutes, dt.Seconds = n.hour, n.minute, n.second
    return dt

parent = doc.createInstance("com.sun.star.text.textfield.Annotation")
parent.setPropertyValue("Content", "Please review")
parent.setPropertyValue("Author", "Reviewer")
parent.setPropertyValue("DateTimeValue", now_dt())
text.insertTextContent(cursor, parent, False)

reply = doc.createInstance("com.sun.star.text.textfield.Annotation")
reply.setPropertyValue("Content", "Addressed in v2")
reply.setPropertyValue("Author", "Author")
reply.setPropertyValue("DateTimeValue", now_dt())
reply.setPropertyValue("ParaIdParent", parent.getPropertyValue("ParaId"))
reply.setPropertyValue("Resolved", True)
text.insertTextContent(cursor, reply, False)

result = "ok"
""".strip()

    try:
        run_cli("execute", session_id, env=cli_env, input=code, timeout=90)
        save_result = run_cli("save", session_id, env=cli_env, timeout=90)
        saved_path = Path(json.loads(save_result.stdout)["saved"])

        assert saved_path.exists()

        import zipfile

        with zipfile.ZipFile(saved_path) as zf:
            names = set(zf.namelist())
            assert "word/comments.xml" in names
            assert "word/commentsExtended.xml" in names

            comments_ext = zf.read("word/commentsExtended.xml").decode("utf-8")
            assert "w15:commentEx" in comments_ext
            assert "w15:paraIdParent" in comments_ext
            assert 'w15:done="1"' in comments_ext

            content_types = zf.read("[Content_Types].xml").decode("utf-8")
            assert "commentsExtended" in content_types

            rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
            assert "commentsExtended" in rels
    finally:
        run_cli("close", session_id, env=cli_env, check=False)
