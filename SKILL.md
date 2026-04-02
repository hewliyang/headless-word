---
name: headless-word
description: "Creating & editing Word documents (.docx) via CLI. Use when the user needs to create, read, edit, or manipulate Word documents."
---

# headless-word CLI

## Setup

```bash
which headless-word 2>/dev/null || pip install headless-word
```

Requires LibreOffice installed:

- macOS: `brew install --cask libreoffice`
- Linux: `sudo apt install libreoffice libreoffice-writer python3-uno`
- Windows: download from libreoffice.org

For screenshots, also need `poppler-utils` (`brew install poppler` / `apt install poppler-utils`).

```bash
headless-word check    # verify environment
```

## Daemon Lifecycle

```bash
headless-word start              # start daemon (waits until ready)
headless-word stop               # stop daemon
headless-word status             # check if running
```

The daemon auto-exits after 5 minutes of inactivity.

## Session Management

Every document operation requires a session. Open returns a session ID used by all subsequent commands:

```bash
headless-word open /path/to/doc.docx        # → {"session_id": "abc123", ...}
headless-word new /path/to/new.docx         # create new doc → session ID
headless-word save <sid> [--path out.docx]  # save (optionally save-as)
headless-word close <sid>                   # close session
headless-word list                          # list open sessions
headless-word export-pdf <sid> out.pdf [--no-comments] [--no-changes] [--pages 1-3]
```

## Reading Tools

All output is JSON on stdout.

### get-document-text

Returns paragraphs with text, style, alignment, list info:

```bash
headless-word get-document-text <sid>
headless-word get-document-text <sid> --start 5 --end 10    # paragraph range
headless-word get-document-text <sid> --no-formatting       # text only
```

### get-document-structure

Returns paragraph count, headings, tables, sections, page count, content controls:

```bash
headless-word get-document-structure <sid>
# → {"paragraph_count": 42, "page_count": 3, "headings": [...], "tables": [...], ...}
```

### get-ooxml

Extracts OOXML and writes it to a temp file. Returns the file path plus a summary with body-child indices, types, line numbers, and text previews. Includes referenced styles and numbering definitions. **Never load the full XML into context** — use `grep` or `read` with offset/limit.

```bash
headless-word get-ooxml <sid>                                # entire body → temp file
headless-word get-ooxml <sid> --start-child 0 --end-child 5  # range of body children

# Inspect formatting:
grep -n 'rFonts\|w:sz \|w:color\|w:b/' /path/to/body-0-5.xml | head -20
grep -n 'w:numPr\|w:pStyle' /path/to/body-0-5.xml | head -20

# Read specific lines (use line numbers from the summary):
read /path/to/body-0-5.xml --offset 710 --limit 40
```

Body children are direct elements under `<w:body>`: paragraphs (`<w:p>`), tables (`<w:tbl>`), content controls (`<w:sdt>`), section properties (`<w:sectPr>`).

### screenshot

Renders a page to PNG via LO PDF export + pdftoppm:

```bash
headless-word screenshot <sid> --page 1 --out page1.png
headless-word screenshot <sid> --page 2 --dpi 300 --no-comments --no-changes
```

## Writing Tools

### execute (primary editing tool)

Execute Python code inside LibreOffice with full UNO API access.

Available in scope: `doc`, `text`, `cursor`, `desktop`, `uno`, `smgr`, `ctx`, `prop`. Set `result` to return a value.

```bash
# Heredoc via stdin (recommended — single-quoted EOF prevents shell expansion)
cat <<'EOF' | headless-word execute <sid>
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
text.insertString(cursor, "Hello World", False)
text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
result = "done"
EOF

# Short inline code
headless-word execute <sid> --code 'text.insertString(cursor, "Hello", False)'

# From file
headless-word execute <sid> --file script.py

# Raw exec mode (returns result directly, no tool wrapper)
headless-word execute <sid> --raw --code 'result = doc.getText().getString()'
```

### insert-ooxml

Insert raw OOXML elements (`<w:p>`, `<w:tbl>`, etc.) into the document body via direct ZIP/XML manipulation:

```bash
headless-word insert-ooxml <sid> --xml '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold text</w:t></w:r></w:p>'
headless-word insert-ooxml <sid> --file snippet.xml --position start
headless-word insert-ooxml <sid> --file snippet.xml --position "after:5"
cat snippet.xml | headless-word insert-ooxml <sid> --position end
```

Positions: `end` (default, before sectPr), `start`, `after:<body_child_index>`

## Live Preview

```bash
headless-word watch /path/to/doc.docx --open          # open browser automatically
headless-word watch /path/to/doc.docx --port 8090 --dpi 150
```

Auto-refreshes when the .docx changes on disk. The watcher opens/closes sessions for each render, so it won't block saves from other sessions.

## ⚠️ Preserving Formatting When Editing Existing Content

**`setString()` and `cell.setString()` DESTROY direct run-level formatting** (`<w:rPr>` — fonts, sizes, colors, bold/italic), reverting to bare style defaults. Most real-world documents rely on direct formatting.

### Editing decision tree:

1. **Simple text replacement** → Use search/replace via `execute` (auto-preserves formatting):
   ```python
   search = doc.createSearchDescriptor()
   search.SearchString = "old text"
   search.ReplaceString = "new text"
   doc.replaceAll(search)
   ```

2. **New content in new/empty document** → Safe to use `execute` with `insertString()`

3. **Editing formatted content** (templates, styled docs) → Inspect first, then use OOXML:
   ```bash
   headless-word get-ooxml <sid> --start-child 2 --end-child 2
   grep -n 'rFonts\|w:sz \|w:color\|w:b/' /path/to/body.xml | head -20
   # If <w:rPr> exists: extract it, construct replacement OOXML with same rPr, use insert-ooxml
   # If only <w:pStyle> and no <w:rPr>: safe to use setString()
   ```

4. **List paragraphs** with `<w:numPr>` — bullets/numbers render automatically. Do NOT prefix text with `- ` or `1. `.

### Adding new content to existing documents

Match the document's formatting — don't guess. Grep a nearby body child for font/size/color and apply the same properties via `execute` or matching `<w:rPr>` blocks in `insert-ooxml`.

## OOXML Reference

**Run properties (`<w:rPr>`)** — direct character formatting:

```xml
<w:r>
  <w:rPr>
    <w:rFonts w:ascii="Open Sans" w:hAnsi="Open Sans" w:cs="Open Sans"/>
    <w:sz w:val="18"/>          <!-- font size in half-points: 18 = 9pt -->
    <w:szCs w:val="18"/>
    <w:color w:val="002060"/>   <!-- hex color without # -->
    <w:b/>                      <!-- bold -->
    <w:i/>                      <!-- italic -->
  </w:rPr>
  <w:t>Formatted text</w:t>
</w:r>
```

**Paragraph properties (`<w:pPr>`)** — element order matters:

```xml
<w:pPr>
  <w:pStyle w:val="Normal"/>     <!-- 1. style -->
  <w:numPr>                      <!-- 2. list numbering -->
    <w:ilvl w:val="0"/>
    <w:numId w:val="1"/>
  </w:numPr>
  <w:spacing w:after="200"/>     <!-- 3. spacing -->
  <w:jc w:val="left"/>           <!-- 4. justification -->
  <w:rPr>...</w:rPr>             <!-- 5. default run properties (LAST) -->
</w:pPr>
```

**Whitespace preservation** — required when text has leading/trailing spaces:

```xml
<w:t xml:space="preserve"> text with spaces </w:t>
```

## Execute Cookbook

### Text & Paragraphs

```python
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

text.insertString(cursor, "Hello World", False)
text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
cursor.setPropertyValue("ParaStyleName", "Heading 1")
text.insertString(cursor, "Section Title", False)
```

### Formatting

```python
from com.sun.star.awt.FontWeight import BOLD

cursor.setPropertyValue("CharHeight", 12)          # font size in pt
cursor.setPropertyValue("CharWeight", BOLD)
cursor.setPropertyValue("CharColor", 0x2E75B6)      # text color (RGB int)
cursor.setPropertyValue("ParaAdjust", 3)             # 0=Left, 1=Right, 2=Justified, 3=Center
cursor.setPropertyValue("ParaTopMargin", 200)        # spacing in 1/100mm
cursor.setPropertyValue("ParaBottomMargin", 100)
cursor.setPropertyValue("ParaLeftMargin", 500)       # indent
```

### Tables

```python
table = doc.createInstance("com.sun.star.text.TextTable")
table.initialize(3, 4)  # rows, cols
text.insertTextContent(cursor, table, False)
table.getCellByName("A1").setString("Header")
table.getCellByName("B1").setValue(42)
```

### Comments

```python
import datetime

def now_dt():
    n = datetime.datetime.now()
    dt = uno.createUnoStruct("com.sun.star.util.DateTime")
    dt.Year, dt.Month, dt.Day = n.year, n.month, n.day
    dt.Hours, dt.Minutes, dt.Seconds = n.hour, n.minute, n.second
    return dt

comment = doc.createInstance("com.sun.star.text.textfield.Annotation")
comment.setPropertyValue("Content", "Review this section")
comment.setPropertyValue("Author", "Reviewer")
comment.setPropertyValue("DateTimeValue", now_dt())  # must set explicitly, defaults to 1900-01-01
text.insertTextContent(cursor, comment, False)

comment.setPropertyValue("Resolved", True)  # resolve
```

**Threaded replies:** Set `ParaIdParent` on the reply to the parent's `ParaId`. `headless-word save` auto-injects proper `commentsExtended.xml` threading (LO's OOXML export doesn't serialize this correctly).

```python
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
text.insertTextContent(cursor, reply, False)
```

### Tracked Changes

```python
doc.setPropertyValue("RecordChanges", True)   # enable
doc.setPropertyValue("RecordChanges", False)  # disable
doc.setPropertyValue("RedlineMode", 0)        # accept all (hide changes)
```

### Images

```python
img = doc.createInstance("com.sun.star.text.TextGraphicObject")
img.setPropertyValue("GraphicURL", uno.systemPathToFileUrl("/path/to/image.png"))
size = uno.createUnoStruct("com.sun.star.awt.Size")
size.Width = 5000   # 50mm
size.Height = 3000  # 30mm
img.setPropertyValue("Size", size)
text.insertTextContent(cursor, img, False)
```

### Headers & Footers

```python
styles = doc.getStyleFamilies().getByName("PageStyles")
ps = styles.getByName("Standard")
ps.setPropertyValue("HeaderIsOn", True)
header = ps.getPropertyValue("HeaderText")
hcursor = header.createTextCursor()
header.insertString(hcursor, "Company Name", False)
```

### Page Setup

```python
styles = doc.getStyleFamilies().getByName("PageStyles")
ps = styles.getByName("Standard")
ps.setPropertyValue("TopMargin", 1000)      # 10mm
ps.setPropertyValue("BottomMargin", 1000)
ps.setPropertyValue("LeftMargin", 1200)
ps.setPropertyValue("RightMargin", 1200)
```

### Read Document Content

```python
result = doc.getText().getString()  # all text

# Iterate paragraphs
enum = text.createEnumeration()
paragraphs = []
while enum.hasMoreElements():
    p = enum.nextElement()
    if p.supportsService("com.sun.star.text.Paragraph"):
        paragraphs.append(p.getString())
result = paragraphs

# Read table cell
tables = doc.getTextTables()
result = tables.getByIndex(0).getCellByName("A1").getString()
```

## Best Practices

1. **Read before writing** — use `get-document-structure` for layout, `get-document-text` for content, `get-ooxml` + `grep` for formatting. Never assume fonts/sizes/styles.
2. **Build incrementally** — add one section at a time, `save` → `screenshot` → verify. Use `watch --open` for live preview.
3. **Use heredocs** — `cat <<'EOF' | headless-word execute <sid>` avoids shell expansion issues.

### Unit reference

- Spacing: 1/100mm in UNO (1000 = 10mm)
- Font size: points in UNO (`CharHeight`), **half-points** in OOXML (`<w:sz w:val="18"/>` = 9pt)
- Colors: RGB int in UNO (`0xFF0000`), hex string without `#` in OOXML (`w:val="002060"`)
- Borders: use `uno.createUnoStruct("com.sun.star.table.BorderLine2")` (don't import)

## UNO API Reference

IDL type definitions in `docs/uno-idl/` (relative to this skill directory). Grep to discover properties, methods, or type names:

```bash
grep -r "Annotation\|ParentName" docs/uno-idl/com/sun/star/text/
grep -r "CellProperties\|BackColor" docs/uno-idl/com/sun/star/table/
grep -r "CharHeight\|ParaAdjust" docs/uno-idl/com/sun/star/awt/ docs/uno-idl/com/sun/star/style/
```

Key namespaces: `text/` (paragraphs, tables, fields, annotations), `sheet/` (cells, ranges, formulas), `table/` (borders, cell properties), `drawing/` + `presentation/` (shapes, slides), `awt/` (font weights, colors, sizes), `style/` (paragraph/character/page styles), `document/` (doc-level properties), `frame/` (desktop, dispatch), `util/` (date, URL, search).

**Note**: Some runtime properties (e.g. `ParentName`, `Resolved`, `ParaId` on Annotation) exist at runtime but aren't declared in the IDL — they're undocumented C++ extensions.
