from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Input models
# ---------------------------------------------------------------------------


class GetDocumentTextParams(BaseModel):
    start_paragraph: int = Field(0, description="0-based start paragraph index")
    end_paragraph: int | None = Field(
        None, description="0-based end paragraph index (exclusive). None = all."
    )
    include_formatting: bool = Field(
        True, description="Include style names and list info"
    )


class GetDocumentStructureParams(BaseModel):
    pass


class GetOoxmlParams(BaseModel):
    start_child: int | None = Field(
        None, description="0-based body-child index to start from"
    )
    end_child: int | None = Field(
        None, description="0-based body-child index to end at (inclusive)"
    )


class ExecuteParams(BaseModel):
    code: str = Field(
        ...,
        description="Python code to execute. Receives doc, text, cursor, desktop, uno, smgr, ctx, prop.",
    )


class ScreenshotDocumentParams(BaseModel):
    page: int = Field(1, ge=1, description="1-based page number to render")
    dpi: int = Field(200, description="DPI for the rendered image")
    show_comments: bool = Field(True, description="Render comments in the margin")
    show_changes: bool = Field(True, description="Show tracked changes")


class InsertOoxmlParams(BaseModel):
    ooxml: str = Field(
        ..., description="Raw OOXML string (<w:p>, <w:tbl>, etc.) to insert"
    )
    position: str = Field(
        "end", description="'end', 'start', or 'after:<paragraph_index>'"
    )


# ---------------------------------------------------------------------------
# Output models
# ---------------------------------------------------------------------------


class ParagraphInfo(BaseModel):
    index: int
    text: str
    style: str | None = None
    alignment: str | None = None
    list_level: int | None = None
    list_string: str | None = None


class GetDocumentTextResult(BaseModel):
    total_paragraphs: int
    showing: dict[str, int]
    paragraphs: list[ParagraphInfo]


class HeadingInfo(BaseModel):
    text: str
    level: int
    paragraph_index: int


class TableInfo(BaseModel):
    index: int
    rows: int
    columns: int
    style: str | None = None


class ContentControlInfo(BaseModel):
    id: int
    title: str
    tag: str
    type: str


class GetDocumentStructureResult(BaseModel):
    paragraph_count: int
    section_count: int
    table_count: int
    content_control_count: int
    page_count: int | None = None
    headings: list[HeadingInfo]
    tables: list[TableInfo]
    content_controls: list[ContentControlInfo]


class BodyChildSummary(BaseModel):
    index: int
    type: str
    line: int
    paragraph_index: int | None = None
    table_index: int | None = None
    paragraph_range: tuple[int, int] | None = None
    rows: int | None = None
    cols: int | None = None
    text: str | None = None


class GetOoxmlResult(BaseModel):
    file: str
    size: str
    lines: int
    children: list[BodyChildSummary]


class ExecuteResult(BaseModel):
    success: bool
    result: Any = None
    error: str | None = None


class ScreenshotDocumentResult(BaseModel):
    page: int
    page_count: int
    png_path: str
    size_bytes: int


class InsertOoxmlResult(BaseModel):
    success: bool
    error: str | None = None
