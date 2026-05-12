Generate a branded Propellic document about: $ARGUMENTS

You have access to the `propellic_doc` MCP tool. Use it to generate and upload the document.

## How to use the tool

1. Based on the user's request, plan the document structure
2. Call `propellic_doc` with these parameters:
   - **title**: Clear, descriptive document title
   - **subtitle**: Supporting context line
   - **sections**: JSON array of section objects (see below)
   - **style**: "report", "proposal", or "brief"

## Section format

Each section in the JSON array should have:
- `heading` (str): Section heading
- `level` (int): 1, 2, or 3
- `type` (str): One of:
  - `"text"` — body paragraph (include `body` field)
  - `"bullets"` — bullet list (include `items` array)
  - `"callout"` — pink-bordered insight box (include `body` field)
  - `"warning"` — berry-colored warning box (include `body` field)
  - `"table"` — data table (include `headers` and `rows` fields)
- `eyebrow` (str, optional): Pink ALL CAPS label above H1 headings

## Brand voice guidelines

- Professional yet approachable — Propellic is a premium SEO and digital marketing agency
- Lead with insights and data, not fluff
- Use sentence case for headings (not Title Case)
- Keep paragraphs concise — aim for 2-4 sentences each
- Use callout boxes for key takeaways and actionable insights
- Use warning boxes sparingly — only for critical blockers or risks

The tool returns a Google Drive link. Share it with the user.
