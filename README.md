# headless-word

Word document automation via LibreOffice for headless environments.

## Quick Start

```bash
uv tool install headless-word

headless-word check              # verify LibreOffice is installed
headless-word start              # start the daemon
headless-word open doc.docx      # open a document → session ID
headless-word get-document-text <sid>
headless-word screenshot <sid> --page 1 --out page1.png
headless-word stop               # stop the daemon
```

## Requirements

- Python 3.10+
- LibreOffice (with Writer)
- `poppler-utils` for screenshots (`pdftoppm`)
