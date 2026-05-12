Generate a branded Propellic presentation about: $ARGUMENTS

You have access to the `propellic_slides` MCP tool. Use it to generate and upload the deck.

## How to use the tool

1. Based on the user's request, plan the slide structure
2. Call `propellic_slides` with these parameters:
   - **title**: Deck title for the title slide
   - **subtitle**: Supporting subtitle
   - **slides**: JSON array of slide objects (see below)
   - **template**: "standard", "pitch", or "case_study"

## Slide types

Each slide in the JSON array should have a `type` key:

- **"agenda"** — Numbered agenda list
  - `items`: list of strings
- **"section"** — Section break (Midnight background)
  - `title`: section title
- **"content"** — Standard content slide
  - `eyebrow`: Pink ALL CAPS category label
  - `headline`: Main slide heading
  - `bullets`: list of up to 5 bullet points
  - `notes`: optional speaker notes
- **"closing"** — Final CTA slide
  - `cta`: Call-to-action text
  - `url`: URL (default: "propellic.com")

## Deck design guidelines

- Start with a title slide (auto-generated) and an agenda slide
- Use section breaks to separate major topics
- Max 5 bullets per content slide — if you have more, split into multiple slides
- Keep bullet text concise (1 line each)
- End with a closing slide
- Add speaker notes for complex slides
- Aim for 8-15 slides total unless the topic demands more

The tool returns a Google Drive link. Share it with the user.
