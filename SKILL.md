---
name: adi-presentation
description: Build PowerPoint decks on the official Analog Devices (ADI) 2026 brand template. Use whenever the user asks for an ADI presentation, slides, or deck, or wants content put on the ADI template, or asks to convert an old deck to ADI branding. Short informal or conversation decks use the clean click-through style (text-led, one point per click, Morph on every slide, notes with click plan). Also covers customer pitches, design-in decks, product intros, technical deep dives, demos, account reviews and roadmaps (story arc, minimal on-slide text, speaker notes with conditions, appendix, optional Morph motion) and linting or restructuring an existing deck.
---

# ADI presentations (brand template + deck builder)

Two ways to build, one brand. Pick the route first.

| Route | Use when | How |
|---|---|---|
| **C. Clean click-through (default for conversation decks)** | Informal or introductory meetings, university and customer conversations, short updates, anything where slides back a discussion (about 5 to 8 slides, 10 to 20 minutes). The user wants minimal text, content revealed point by point, Morph transitions, speaker notes with sources | `scripts/clean_deck.py` (`CleanDeck`), starting from `scripts/example_clean_deck.py`. Style and rules: `references/clean-click-through.md`. Reusable prompt: `references/prompt-template.md` |
| **A. Template-native** | Short everyday/internal decks (about 10 slides), conversions of old decks, org charts, anything needing the template's own layouts, library patterns, icons or photos | python-pptx directly on the template layouts. Steps in "Route A" below |
| **B. Deck builder** | Customer pitches, design-in meetings, product intros, deep dives, demos, account reviews: where story, notes, claims-checking and lint matter | Brief JSON -> `scripts/build_deck.py` -> lint. Steps in "Route B" below |

All routes draw on `assets/ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx`, so the brand rules below apply to all. Route B reads `private/theme.json`, which points at that template (Barlow fonts, ADI palette, template masters, logo, classification footer). If the user's request is ambiguous: C for conversation or meeting decks, B for sales pitches and design-ins, A for static internal decks and conversions.

### Route C in brief (the house style the user has approved)
Read `references/clean-click-through.md` first, then `references/morph-patterns.md` for the hand-crafted Morph techniques (make-room pairs, focus-row accordion, parking off-canvas, arrows and paragraph builds, slide links to a hidden appendix; `scripts/example_morph_patterns.py` demonstrates each). Non-negotiables: **text-led with thin rules; no coloured cards, bars, strips, chevrons or highlight rings; every point appears on its own click (about 7 clicks per slide, about 30 per 10-minute deck); Morph transition on every slide; notes start with a `CLICKS (n)` cue line, then script, sources and `[TO CONFIRM]` cautions; never over-promise (ideas for discussion, no implied funding, equipment, vacancies or approvals); UK English; Public footer unless told otherwise.** Workflow: (1) search the user's source folders and read what is relevant, (2) outline slides with the click plan, (3) copy `scripts/example_clean_deck.py`, replace the content, run it, (4) fix any `CHECK` lines, (5) deliver under a new version name if the file is open, and report clicks, `[TO CONFIRM]` items, sources used and that no visual render was done if so. The user can paste `references/prompt-template.md` to get this every time.

## Brand rules (from the template's own guidance slides)
- Fonts: **Barlow Medium** (titles) and **Barlow** (body). They come from the theme; never name other fonts on slides.
- Body text sizes: level 1 = 18 pt, no bullet; level 2 = 16 pt; level 3 = 14 pt. Tables 14 pt. 14 pt is the floor for any text except the footer.
- Titles use Title Case (articles, conjunctions and short prepositions lowercase unless first word or over four letters). A cover title longer than 3 lines may shrink, never below 36 pt.
- **Do not move, resize or restyle** the template's text containers. Fit content inside them.
- Prefer light backgrounds. Use the blue / dark blue layouts only for key messages, call-outs and big data visuals.
- Palette only (theme colors): dk2/primary blue `0067B9`, accent1 dark blue `00325C`, accent2 `1B9CD0`, accent3 `8637BA`, accent4 `179963`, accent5 `FED141`, accent6 `C81A28`, gray `9EA1AE`. Route A references theme colors, not hex. Keep contrast high.
- Max ~10 slides / 20 minutes for everyday decks. One message per slide, visuals over text. No clip art, sound effects or gratuitous animation. This template is NOT for auditorium or keynote stages.
- Large, accessible charts. Use native charts, theme accent colors, 14 pt labels.
- Footer text must match the classification (set on every content slide, keep the slide number):
  - Public: `©2026 Analog Devices, Inc. All Rights Reserved.`
  - Confidential (NDA): `Analog Devices Confidential Information. ©2026 Analog Devices, Inc. All Rights Reserved.`
  - Internal only: `Analog Devices Confidential Information—Not for External Distribution. ©2026 Analog Devices, Inc. All Rights Reserved.`
  If the user does not say, ask which classification applies; default to Public only for obviously public content. In Route B set `"meta": {"classification": "public|confidential|internal"}` in the brief.
- Images: only from the Brand Image Library or brand-approved photography. Do not invent or download images.

### Where the deck builder is deliberately more restrained than its original
The builder was written for keynote-style pitching. On the ADI template:
- **Motion defaults to `minimal`** (quiet fade, no builds). Use `standard` for a customer meeting only if the user wants Morph moments and click builds, and `showcase` only if they explicitly ask. The brand rule against gratuitous animation wins when in doubt.
- Keep Morph moments to the few that carry meaning (3 at most unless asked).
- Main-flow slides carry few words, but never drop below 14 pt to achieve it.
- Builder layouts draw on the template's "Title Only – Blank" layout and place their own text boxes; footer and slide number are cloned from the layout. If a user wants strict placeholder-driven titles (accessibility, outline view), use Route A.

## Layouts (slide size 13.333 x 7.5 in) — placeholder idx in brackets
| Layout | Use | Placeholders |
|---|---|---|
| Cover with Pattern / with Image | Opening slide | title [10], name [14], date [15] |
| Title and Content | Default content slide | title, body [10] |
| Two Line Title and Content | Long titles | title, body [11] |
| Section Slide with Pattern / Image | Dividers, or "Key Message" | text [10] |
| Two Column Content | Comparisons | title, left [10], right [11] |
| Title and Content with Sub Header | Body with sub header | sub header [16], body [11] |
| Content with Sub Header Two Column | Two columns with headers | sub [16],[22], body [11],[23] |
| Content and Image | Text plus photo | title, picture [16], body [11] |
| Title Only – Blank | Charts, tables, custom visuals (Route B builds on this) | title |
| Two Column Alternate | Two panels with headers | [10],[19],[20],[21] |
| Title and Content – Primary Blue / Dark Blue (+ Two Line, + Sub Header) | Emphasis slides | title, body [11] |
| Org Chart 5x2 | People grids | 10 picture + 2 text placeholders each |
| Big Statement Two Column (+ Blue) | Statement plus supporting content | statement [27], content [23] / [11] |
| Big Statement – Primary Blue | One bold sentence | statement [28] |
| Video | Full-bleed video | media [12] |
| End Card | Final slide | none |

Reusable slide patterns (agenda, dashboard, roadmap, timeline, progress roadmap, next steps, before/after, pull-out quote, team success, etc.) live in `assets/Brand-PPT-Slide-Library.pptx`. Duplicate the slide with the pptx skill's `add_slide.py` and replace the text rather than rebuilding it.

## Other brand assets (all in `assets/`)
- **Icons** — `Brand-Icon-Library_2025-v2.pptx` (10 slides; slide 2 = usage rules). Sections: Capabilities (3), Industry Solutions (4), Car/Charging (5), Healthcare (6), Energy/Environment (7), Technology (8), Technology—Consumer (9), Workplace (10). Vector graphics: copy the shape XML (or duplicate the slide and delete the rest). Recolor with brand colors only; preferred: white, black, ADI blue, ADI light blue.
- **Images** — `ADI-Brand-PowerPoint-Image-Library.pptx` (62 slides, one licensed photo each; `image-index.md` describes them). Licensed for PowerPoint cover and section slides only. Slide 1 has instructions, slide 2 shows placement. Pre-cropped to 33.87 x 19.05 cm with the focal point on the right. Swap via the Slide Master (cover / section layout): delete the old picture, insert the new, send to back. Always show the user which image you picked.
- **Org chart** — `ADI-Brand-PowerPoint-Template-Org-Chart-2026-v1.pptx` (cover + 2 org chart slides). Same grid as layout "Org Chart 5x2". Internal Only unless told otherwise.
- **Previews** — `assets/previews/sheet_<name>_<n>.png` are 12-slide contact sheets of the template (`std`), slide library (`lib`), icons (`icons`), org chart (`org`) and images (`imgs`). View them to choose a layout, pattern, icon or photo without re-rendering.

## Route A: template-native build
1. Clarify only what you cannot infer: topic, audience, length, footer classification. Outline one message per slide, titles that state the message.
2. `python scripts/prep_template.py <template.potx> <work/adi_base.pptx>` converts the template to an openable .pptx and prints every layout and placeholder idx.
3. Build with python-pptx (see `scripts/example_build.py`): delete the guidance slides, add slides by **layout name**, fill placeholders by **idx**, copy the layout's footer [3] and slide number [4] placeholders onto each non-cover/divider slide and set the footer text. Deps: `python -m pip install --user python-pptx lxml defusedxml Pillow`. Do not run with `python -I`.
4. Native charts and tables, theme colors only. For visual slides (agenda, status, next steps) build cards on "Title Only" using the `box` / `text` / `badge` helpers in `scripts/example_build.py`. Avoid plain title-plus-bullets slides.
5. QA: `python <pptx skill>/scripts/office/validate.py deck.pptx --original <base.pptx>`, then `slide_text.py` to proof the content and grep for leftover template text. Optionally `python scripts/lint_deck.py deck.pptx` for claims/units/competitor checks.
6. Render and look at every slide (see "Rendering").

## Route B: deck builder (story, notes, lint)
The principle: **the slides are a backdrop for the presenter, not a document.** One idea per slide in a few large words; the script and every number's conditions live in the speaker notes; datasheet detail lives in an appendix.

| Path | What it is | When to read |
|---|---|---|
| `references/story.md` | Pitch arc, slide sequence, ADI adaptations | Every customer pitch, before outlining |
| `references/brief-format.md` | Brief JSON format and every layout's fields | Before writing a brief |
| `references/design.md` | Type, colour, photos, layout, what not to do | Choosing layouts or hand-building slides |
| `references/motion.md` | Transition grammar, Morph, click builds, levels | Before choosing a motion level above minimal |
| `references/deck-types.md` | Demos, deep dives, account reviews, internal updates, training, conference talks | When not a straight customer pitch |
| `references/adi-template.md` | Template/theme setup and open questions (see "Setup state" below) | If branding looks wrong |
| `scripts/build_deck.py` | Brief JSON -> .pptx, runs lint, writes `<deck>.lint.md` | Always |
| `scripts/lint_deck.py` | Story, slide, claims and build checks; works on any .pptx | After every build; to review a user's deck |
| `scripts/deckkit.py`, `layouts.py`, `motion.py` | Drawing engine, 22 layouts, transitions and animation | When a slide needs something new |
| `examples/brief_example.json` | Complete worked brief (fictional customer) | Model for a new brief |

Run: `python <skill_dir>/scripts/build_deck.py brief.json -o deck.pptx` (add `--static` for a no-motion PDF/leave-behind copy). Deps: `python-pptx`, `Pillow`, `lxml`; `shapely` optional (mosaic title only). On Windows set `PYTHONIOENCODING=utf-8` if output shows garbled characters.

Workflow:
1. **Deck type.** Customer pitch / design-in uses the full arc in `references/story.md`; anything else starts from `references/deck-types.md`. If the user only wants a deck checked, skip to step 6.
2. **Inputs.** Customer and application, who is in the room, their goal and pain (with numbers), their 2–3 criteria, the ADI part or signal chain, 2–3 headline specs **with conditions**, today's alternative, proof, the ask with a date, and the footer classification. Ask for what is missing in one grouped message. If a draft is wanted now, mark gaps `[TBC: ...]`. **Never invent specs, prices, customer data or reference wins.** Competitor names and comparative claims stay flagged for review.
3. **Story before slides.** State the one idea of the deck, then a slide-by-slide outline (role, words on screen, what the presenter says). Check against the arc: problem before feature, criteria before reveal, proof after claims, a dated ask at the end. Show a substantial outline to the user before building.
4. **Brief JSON.** Copy the shape of `examples/brief_example.json`. About six words per main-flow slide; `notes.conditions` on every slide with a number; `notes.source` on comparisons; `[CLICK]` markers in `notes.say` only if motion is above minimal; set `meta.classification`. Asset paths are relative to the brief; missing assets render as labelled dashed boxes.
5. **Build.** Theme resolution: `--theme`, else the brief's `theme`, else `private/theme.json` (the real ADI theme, present here), else the placeholder palette (not brand-ready; say so).
6. **Lint and fix.** Fix every FAIL in the brief and rebuild. REVIEW items (competitors, `[TBC]` sources, approximate dates) go to the user as explicit actions; the deck is not presentable until resolved. For a hand-made deck: `python scripts/lint_deck.py their_deck.pptx --competitors "A,B" --require-footer "Analog Devices"` (story-order checks are partial without a `role: ...` first notes line).
7. **Look at it** (see "Rendering"). Lint does not check geometry; overflow and collisions need eyes. Shorten text rather than shrinking it.
8. **Deliver.** The .pptx (and `--static` copy if it will be a PDF), then in a few lines: the one idea, slide count, open items (`[TBC]`, placeholder assets, REVIEW flags), and a note to rehearse in presenter view.

Judgement calls: keep it about the customer; candour about weaknesses buys credibility; superlatives only after evidence with conditions; builds and motion must read as stills because decks get forwarded and exported to PDF.

Going beyond the layouts: compose in Python with `deckkit.Deck` (text, rect, line, picture, `d.tl(slide).add(...)`, `d.morph_as(...)`) and follow the naming rules: `!!name` for Morph-paired objects, `aux_` for labels lint should not word-count, `photo_` for photographs, `ftr_` for chrome. Tag the role with `set_slide_tag(slide, role="proof")`. To edit an existing .pptx rather than regenerate it, use the general `anthropic-skills:pptx` skill, then lint the result.

## Rendering (both routes)
LibreOffice is not installed on this machine; use PowerPoint: `powershell scripts/render_slides.ps1 -In deck.pptx -OutDir <dir> -Width 1280` writes `Slide1.PNG`, … It only works with the shell sandbox disabled (inside it PowerPoint fails with COM error 0x80080005). If rendering is not possible, say plainly that visual QA was not done and ask the user to open the file. Delete old images before re-rendering. Morph and animation do not show in a render: check each state-A slide reads as a still and tell the user to run Slide Show once.

## Setup state (this install)
`private/theme.json` was generated by `scripts/extract_template.py` from the 2026 Standard template and hand-tuned (Barlow fonts, ADI palette, classification footers, light template background, motion `minimal`). Regenerate or edit it if the template changes; `references/adi-template.md` describes the process (its "placeholder palette" caveats are superseded by this file). `SKILL.deckbuilder.orig.md` is the original standalone deck-builder SKILL.md, kept for reference.

## Gotchas
- python-pptx refuses `.potx`; `prep_template.py` and `deckkit.open_template` rewrite the content type in a copy.
- Cover and section layouts get their image or pattern from the layout/master; do not add pictures to them.
- Empty picture placeholders show prompt text in edit view; fill them or use a layout without one.
- Dropping slides with python-pptx: call `prs.part.drop_rel(rId)` and remove the `sldId`, otherwise orphaned slides are saved.
- Scripts must read/write JSON and markdown as UTF-8 (the `©` footer corrupts to "Â©" under the Windows default codepage). When editing files with non-ASCII characters use the Edit tool, not a PowerShell `Get-Content`/`Set-Content` round trip.
