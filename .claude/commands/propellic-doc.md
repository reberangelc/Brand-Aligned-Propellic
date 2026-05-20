Create a fully branded Propellic document about: $ARGUMENTS

Write a Python script using python-docx and run it to generate an actual .docx file, then immediately upload it to Google Drive as a Google Doc. Use the shebang `#!/usr/bin/env python3` at the top. Save the .docx to the current working directory with a descriptive kebab-case filename.

The logo file is at: propellic-logo-dark.png (in the current working directory — pink flame + Midnight wordmark, transparent background, works on white).

---

## Propellic Brand Spec — apply exactly

### Colors (RGB tuples for python-docx RGBColor)
- Midnight: RGBColor(0x15, 0x25, 0x34)
- Pink:     RGBColor(0xE2, 0x1A, 0x6B)
- White:    RGBColor(0xFF, 0xFF, 0xFF)
- Blush:    RGBColor(0xF7, 0xBF, 0xD5)
- Berry:    RGBColor(0x8F, 0x1F, 0x55)

### Fonts
- All text: Montserrat (available in Google Docs — approved substitute for Proxima Nova)
- Fallback: Calibri

### Google Docs compatibility rules (CRITICAL)
1. Use Word's built-in named styles — "Heading 1", "Heading 2", "Heading 3", "Normal" — then override their font/color/size. These map to Google Docs outline levels correctly.
2. Never use paragraph shading for callout boxes — use single-cell tables instead. Paragraph shading is unreliable in Google Docs.
3. Use table-based cover block (not paragraph shading) so Midnight background renders in Docs.
4. Keep borders simple — table borders, not paragraph XML borders.
5. Line spacing via paragraph_format.line_spacing_rule and line_spacing — avoid raw Pt() spacing only.
6. Images embed cleanly — always use add_picture() with a width constraint.

### Document structure

**Page header (repeats on every page) — full-width 1-row table, no borders:**
- Cell background: Midnight (152534)
- Logo: propellic-logo.png, left-aligned, height ~0.3in, inside the header table cell
- The header table fills the full page width including margins (use section.page_width)
- Set header distance to 0 so it sits flush at the top: section.header_distance = Pt(0)

**Cover block — full-width 1-row table, no borders (NO logo here — logo is in header):**
- Cell background: Midnight (152534)
- Title text in white bold 28pt, subtitle in white 13pt
- Pink accent: add a second 1-row table below cover, full width, Pink fill, height ~6pt — acts as accent divider

**Eyebrow labels:**
- Paragraph above each H1: Pink, Montserrat Bold, ALL CAPS, 9pt, space_after=2pt

**Heading styles (override built-ins for Google Docs compatibility):**
- Heading 1: Montserrat Bold, 20pt, Midnight, sentence case
- Heading 2: Montserrat Bold, 14pt, Pink, sentence case
- Heading 3: Montserrat Bold, 12pt, Midnight, sentence case

**Body text:**
- Normal style, Montserrat Regular, 11pt, Midnight
- line_spacing = Pt(17), space_after = Pt(6)

**Callout box (insight) — single-cell table:**
- Table width: full page width
- Cell background: Blush (F7BFD5)
- Left border: Pink (E21A6B), 4pt thick
- Other borders: none
- Text: Montserrat Regular, 11pt, Midnight, 8pt padding

**Callout box (warning) — single-cell table:**
- Cell background: Berry (8F1F55)
- All borders: none
- Text: Montserrat Regular, 11pt, White, 8pt padding

**Tables (data tables):**
- Header row: Midnight fill, white bold ALL CAPS Montserrat, 10pt
- Alternate rows: White / Blush
- All borders: light gray, 0.5pt

**Footer:**
- Right-aligned: "propellic.com" in Pink, Montserrat, 9pt
- Auto page number

---

### Python implementation pattern

```python
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

doc = Document()

MIDNIGHT = RGBColor(0x15, 0x25, 0x34)
PINK     = RGBColor(0xE2, 0x1A, 0x6B)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLUSH    = RGBColor(0xF7, 0xBF, 0xD5)
BERRY    = RGBColor(0x8F, 0x1F, 0x55)

# Page setup
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1)
    section.right_margin  = Inches(1)

PAGE_WIDTH = Inches(6.5)  # 8.5 - 1 - 1

def cell_shading(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, left_color=None, left_sz=32, others_color=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBdr = OxmlElement('w:tcBdr')
    for side in ['top', 'left', 'bottom', 'right']:
        el = OxmlElement(f'w:{side}')
        if side == 'left' and left_color:
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(left_sz))
            el.set(qn('w:color'), left_color)
        elif others_color:
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), '4')
            el.set(qn('w:color'), others_color)
        else:
            el.set(qn('w:val'), 'none')
        tcBdr.append(el)
    tcPr.append(tcBdr)

def set_table_width(table, width):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(int(width.inches * 1440)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

def add_run(p, text, size=11, bold=False, italic=False, color=None):
    r = p.add_run(text)
    r.font.name = 'Montserrat'
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    if color:
        r.font.color.rgb = color
    return r

def add_cover(title, subtitle='', logo_path='propellic-logo.png'):
    # Cover table — Midnight background
    cover_tbl = doc.add_table(rows=1, cols=1)
    cover_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_width(cover_tbl, PAGE_WIDTH)
    cell = cover_tbl.rows[0].cells[0]
    cell_shading(cell, '152534')
    set_cell_border(cell)  # no borders
    cell.width = PAGE_WIDTH

    # Logo
    if os.path.exists(logo_path):
        logo_p = cell.paragraphs[0]
        logo_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        logo_p.paragraph_format.space_before = Pt(16)
        logo_p.paragraph_format.space_after  = Pt(8)
        r = logo_p.add_run()
        r.add_picture(logo_path, height=Inches(0.35))

    # Title
    title_p = cell.add_paragraph()
    title_p.paragraph_format.space_before = Pt(8)
    title_p.paragraph_format.space_after  = Pt(4)
    add_run(title_p, title, size=28, bold=True, color=WHITE)

    # Subtitle
    if subtitle:
        sub_p = cell.add_paragraph()
        sub_p.paragraph_format.space_before = Pt(0)
        sub_p.paragraph_format.space_after  = Pt(16)
        add_run(sub_p, subtitle, size=13, color=WHITE)

    # Pink accent divider table
    divider = doc.add_table(rows=1, cols=1)
    set_table_width(divider, PAGE_WIDTH)
    d_cell = divider.rows[0].cells[0]
    cell_shading(d_cell, 'E21A6B')
    set_cell_border(d_cell)
    d_cell.height = Pt(5)
    d_p = d_cell.paragraphs[0]
    d_p.paragraph_format.line_spacing = Pt(5)
    d_p.paragraph_format.space_before = Pt(0)
    d_p.paragraph_format.space_after  = Pt(0)
    add_run(d_p, ' ', size=4, color=PINK)

    doc.add_paragraph()  # spacer after cover

def add_eyebrow(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(1)
    add_run(p, text.upper(), size=9, bold=True, color=PINK)

def add_h1(text):
    p = doc.add_heading(level=1)
    p.clear()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(4)
    add_run(p, text, size=20, bold=True, color=MIDNIGHT)

def add_h2(text):
    p = doc.add_heading(level=2)
    p.clear()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(3)
    add_run(p, text, size=14, bold=True, color=PINK)

def add_h3(text):
    p = doc.add_heading(level=3)
    p.clear()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(2)
    add_run(p, text, size=12, bold=True, color=MIDNIGHT)

def add_body(text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = Pt(17)
    p.paragraph_format.space_after  = Pt(6)
    add_run(p, text, size=11, color=MIDNIGHT)

def add_bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    add_run(p, text, size=11, color=MIDNIGHT)

def add_callout(text, warning=False):
    """Single-cell table callout — reliable in Google Docs."""
    t = doc.add_table(rows=1, cols=1)
    set_table_width(t, PAGE_WIDTH)
    cell = t.rows[0].cells[0]
    cell_shading(cell, '8F1F55' if warning else 'F7BFD5')
    if warning:
        set_cell_border(cell)
    else:
        set_cell_border(cell, left_color='E21A6B', left_sz=48)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    txt_color = WHITE if warning else MIDNIGHT
    add_run(p, text, size=11, color=txt_color)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_data_table(headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    set_table_width(t, PAGE_WIDTH)
    # Header row
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.paragraphs[0].clear()
        p = cell.paragraphs[0]
        add_run(p, h.upper(), size=10, bold=True, color=WHITE)
        cell_shading(cell, '152534')
        set_cell_border(cell, others_color='CCCCCC')
    # Data rows
    for r_i, row_data in enumerate(rows):
        bg = 'F7BFD5' if r_i % 2 == 1 else 'FFFFFF'
        for c_i, val in enumerate(row_data):
            cell = t.rows[r_i + 1].cells[c_i]
            cell.paragraphs[0].clear()
            add_run(cell.paragraphs[0], str(val), size=10, color=MIDNIGHT)
            cell_shading(cell, bg)
            set_cell_border(cell, others_color='CCCCCC')
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_header(logo_path='propellic-logo.png'):
    """Midnight header bar with logo — repeats on every page."""
    section = doc.sections[0]
    section.header_distance = Pt(0)
    header = section.header
    # Clear default paragraph
    for p in header.paragraphs:
        p._element.getparent().remove(p._element)
    # Full-width Midnight table
    full_width = section.page_width
    t = header.add_table(rows=1, cols=1, width=full_width)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = t.rows[0].cells[0]
    cell_shading(cell, '152534')
    set_cell_border(cell)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    if os.path.exists(logo_path):
        p.add_run().add_picture(logo_path, height=Inches(0.3))

def add_footer():
    section = doc.sections[0]
    footer = section.footer
    footer.paragraphs[0].clear()
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(p, 'propellic.com  |  ', size=9, color=PINK)
    r = p.add_run()
    r.font.name = 'Montserrat'
    r.font.size = Pt(9)
    r.font.color.rgb = PINK
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    r._r.append(fldChar1)
    r._r.append(instrText)
    r._r.append(fldChar2)
```

Using this pattern, build a complete document for: $ARGUMENTS

Call add_footer() once at the end before saving. After saving, immediately upload to Google Drive using this pattern — do NOT use the Google Drive MCP tool for this, it cannot handle binary uploads:

```python
import subprocess, sys

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
GDOC_MIME = 'application/vnd.google-apps.document'
TOKEN_PATH  = os.path.expanduser('~/.propellic/token.json')
VENV_PYTHON = os.path.expanduser('~/.propellic/venv/bin/python3')

# If running outside the venv, re-exec with it
if sys.executable != VENV_PYTHON and os.path.exists(VENV_PYTHON):
    os.execv(VENV_PYTHON, [VENV_PYTHON] + sys.argv)

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

creds = Credentials.from_authorized_user_file(TOKEN_PATH)
if creds.expired and creds.refresh_token:
    creds.refresh(Request())

service = build('drive', 'v3', credentials=creds)
media = MediaFileUpload(out_path, mimetype=DOCX_MIME, resumable=False)
file_meta = {'name': doc_title, 'mimeType': GDOC_MIME}
uploaded = service.files().create(body=file_meta, media_body=media, fields='id').execute()
file_id = uploaded['id']
print(f"File saved: {out_path}")
print(f"Google Doc: https://docs.google.com/document/d/{file_id}/edit")
```

Set `doc_title` to a human-readable title string (e.g. `"Propellic — Guide to Cats"`) before the upload block. Run the script with plain `python3` — the re-exec handles switching to the venv automatically.
