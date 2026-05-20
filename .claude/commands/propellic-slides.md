Create a fully branded Propellic presentation deck about: $ARGUMENTS

Write a Python script using python-pptx and run it to generate an actual .pptx file, then immediately upload it to Google Drive as a Google Slides presentation. Use the shebang `#!/usr/bin/env python3` at the top. Save the .pptx to the current working directory with a descriptive kebab-case filename.

The logo files are in the current working directory:
- `propellic-logo.png` — white flame + white wordmark, use on Midnight backgrounds
- `propellic-logo-dark.png` — pink flame + Midnight wordmark, use on white backgrounds

---

## Propellic Brand Spec

### Colors
- Midnight: RGBColor(0x15, 0x25, 0x34) — dark slide backgrounds, sidebars
- Pink:     RGBColor(0xE2, 0x1A, 0x6B) — eyebrows, accents, CTAs, divider bars
- White:    RGBColor(0xFF, 0xFF, 0xFF) — text on dark slides
- Blush:    RGBColor(0xF7, 0xBF, 0xD5) — callout backgrounds
- Berry:    RGBColor(0x8F, 0x1F, 0x55) — warning callouts, secondary accents

### Typography
- All text: Montserrat (set via run.font.name = 'Montserrat')
- Headlines: Bold, sentence case
- Eyebrows: Bold, ALL CAPS, Pink, 9–11pt
- Body/bullets: Regular weight
- CTAs: Bold, Title Case

### Slide types to use

**Title slide** — Midnight background, white logo top-left, Pink accent bar at bottom, title in white 40pt bold, subtitle in Pink 20pt

**Agenda slide** — White background, Midnight left sidebar (0.18in wide), eyebrow in Pink, headline in Midnight, numbered list items in Midnight 16pt

**Section break** — Midnight background, white logo top-left, Pink accent bar at bottom, section title centered in white 36pt bold

**Content slide** — White background, Midnight left sidebar (0.18in wide), eyebrow ALL CAPS Pink 9pt, headline Midnight 28pt bold, Pink divider line under headline, max 5 bullet points at 16pt Midnight, speaker notes in notes pane

**Closing slide** — Midnight background, white logo top-left, Pink accent bar at bottom, CTA centered in white 32pt bold, URL in Pink 16pt below

---

### Python implementation pattern

```python
#!/usr/bin/env python3
import os, sys

VENV_PYTHON = os.path.expanduser('~/.propellic/venv/bin/python3')
if sys.executable != VENV_PYTHON and os.path.exists(VENV_PYTHON):
    os.execv(VENV_PYTHON, [VENV_PYTHON] + sys.argv)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

MIDNIGHT = RGBColor(0x15, 0x25, 0x34)
PINK     = RGBColor(0xE2, 0x1A, 0x6B)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLUSH    = RGBColor(0xF7, 0xBF, 0xD5)
BERRY    = RGBColor(0x8F, 0x1F, 0x55)

BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
LOGO_WHITE     = os.path.join(BASE_DIR, 'propellic-logo.png')       # for dark slides
LOGO_DARK      = os.path.join(BASE_DIR, 'propellic-logo-dark.png')  # for light slides

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]
W, H  = prs.slide_width, prs.slide_height


def set_bg(slide, color):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = color


def add_rect(slide, left, top, width, height, color, line=False):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if not line:
        shape.line.fill.background()
    return shape


def add_text(slide, text, left, top, width, height,
             size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = 'Montserrat'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txb


def add_logo(slide, dark_bg=True, left=Inches(0.5), top=Inches(0.3), height=Inches(0.42)):
    path = LOGO_WHITE if dark_bg else LOGO_DARK
    if os.path.exists(path):
        slide.shapes.add_picture(path, left, top, height=height)


def pink_bar(slide, height_pt=7):
    h = Pt(height_pt)
    add_rect(slide, 0, H - int(h), W, int(h), PINK)


def midnight_sidebar(slide, width=Inches(0.18)):
    add_rect(slide, 0, 0, width, H, MIDNIGHT)


def set_notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


def title_slide(title, subtitle=''):
    slide = prs.slides.add_slide(BLANK)
    set_bg(slide, MIDNIGHT)
    add_logo(slide, dark_bg=True)
    pink_bar(slide)
    add_text(slide, title,
             Inches(0.5), Inches(2.1), Inches(12.3), Inches(1.8),
             size=40, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, subtitle,
                 Inches(0.5), Inches(3.9), Inches(10.0), Inches(0.8),
                 size=20, bold=False, color=PINK)


def agenda_slide(items):
    slide = prs.slides.add_slide(BLANK)
    set_bg(slide, WHITE)
    midnight_sidebar(slide)
    add_text(slide, 'AGENDA', Inches(0.5), Inches(0.45), Inches(12.0), Inches(0.4),
             size=9, bold=True, color=PINK)
    add_text(slide, 'Today\'s agenda',
             Inches(0.5), Inches(0.85), Inches(12.0), Inches(0.8),
             size=28, bold=True, color=MIDNIGHT)
    add_rect(slide, Inches(0.5), Inches(1.7), Inches(12.0), Pt(2), PINK)
    for i, item in enumerate(items):
        add_text(slide, f'{i + 1}.  {item}',
                 Inches(0.6), Inches(2.0) + Inches(i * 0.72), Inches(11.8), Inches(0.65),
                 size=16, bold=False, color=MIDNIGHT)


def section_break(title):
    slide = prs.slides.add_slide(BLANK)
    set_bg(slide, MIDNIGHT)
    add_logo(slide, dark_bg=True)
    pink_bar(slide)
    add_text(slide, title,
             Inches(0.5), Inches(2.9), Inches(12.3), Inches(1.4),
             size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def content_slide(eyebrow, headline, bullets=None, notes=''):
    slide = prs.slides.add_slide(BLANK)
    set_bg(slide, WHITE)
    midnight_sidebar(slide)
    add_text(slide, eyebrow.upper(),
             Inches(0.5), Inches(0.45), Inches(12.0), Inches(0.4),
             size=9, bold=True, color=PINK)
    add_text(slide, headline,
             Inches(0.5), Inches(0.85), Inches(12.3), Inches(0.9),
             size=28, bold=True, color=MIDNIGHT)
    add_rect(slide, Inches(0.5), Inches(1.75), Inches(12.3), Pt(2), PINK)
    if bullets:
        for i, bullet in enumerate(bullets[:5]):
            add_text(slide, f'•  {bullet}',
                     Inches(0.6), Inches(2.05) + Inches(i * 0.88), Inches(11.8), Inches(0.8),
                     size=16, bold=False, color=MIDNIGHT)
    set_notes(slide, notes)


def closing_slide(cta, url='propellic.com'):
    slide = prs.slides.add_slide(BLANK)
    set_bg(slide, MIDNIGHT)
    add_logo(slide, dark_bg=True)
    pink_bar(slide)
    add_text(slide, cta,
             Inches(0.5), Inches(2.8), Inches(12.3), Inches(1.1),
             size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, url,
             Inches(0.5), Inches(3.95), Inches(12.3), Inches(0.55),
             size=16, bold=False, color=PINK, align=PP_ALIGN.CENTER)
```

Using this pattern, build a complete deck for: $ARGUMENTS

Design the right number and mix of slides for the topic. After building all slides, save and upload:

```python
PPTX_MIME    = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
GSLIDES_MIME = 'application/vnd.google-apps.presentation'
TOKEN_PATH   = os.path.expanduser('~/.propellic/token.json')

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

creds = Credentials.from_authorized_user_file(TOKEN_PATH)
if creds.expired and creds.refresh_token:
    creds.refresh(Request())

service  = build('drive', 'v3', credentials=creds)
media    = MediaFileUpload(out_path, mimetype=PPTX_MIME, resumable=False)
uploaded = service.files().create(
    body={'name': deck_title, 'mimeType': GSLIDES_MIME},
    media_body=media,
    fields='id'
).execute()

print(f"File saved: {out_path}")
print(f"Google Slides: https://docs.google.com/presentation/d/{uploaded['id']}/edit")
```

Set `deck_title` to a human-readable string (e.g. `"Propellic — Kartrite SEO Kickoff"`) and `out_path` to the kebab-case .pptx filename before the upload block. Run with plain `python3` — the re-exec handles switching to the venv automatically.
