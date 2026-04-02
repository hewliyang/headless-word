# headless-word

Word document automation via LibreOffice for headless environments.

## Install

```bash
uv tool install headless-word
```

## Agent Skill

```bash
npx skills add hewliyang/headless-word
```

## Requirements

- Python 3.10+
- LibreOffice 25.2+ (with Writer)
- `poppler-utils` for screenshots (`pdftoppm`)

## Quick Start

```bash
headless-word check              # verify LibreOffice is installed
headless-word start              # start the daemon
headless-word new doc.docx       # create a new document → session ID
headless-word get-document-text <sid>
headless-word screenshot <sid> --page 1 --out page1.png
headless-word stop               # stop the daemon
```

## Commands

```
headless-word <command> [options]

Daemon:
  start                Start the LibreOffice daemon
  stop                 Stop the daemon
  status               Check daemon status
  check                Check environment setup

Sessions:
  open                 Open a document
  new                  Create a new document
  save                 Save a document
  close                Close a session
  list                 List open sessions
  export-pdf           Export as PDF

Reading:
  get-document-text    Get document text (paragraphs, styles, lists)
  get-document-structure  Get document structure (headings, tables, sections)
  get-ooxml            Extract OOXML for inspection
  screenshot           Render a page to PNG

Writing:
  execute              Execute Python/UNO code inside LibreOffice
  insert-ooxml         Insert raw OOXML elements into document

Live Preview:
  watch                Live document viewer with auto-reload
```
