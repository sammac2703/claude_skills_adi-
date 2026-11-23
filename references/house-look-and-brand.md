# House look, brand rules and motion (moved from SKILL.md)

## House look: keynote style (dominant, all routes)
Read `references/keynote-style.md` before building anything. Every deck uses it unless the user asks for the light look:
- **Canvas:** black to ADI dark blue (`00325C`) diagonal gradient on every content slide, on the template's *Title and Content - Dark Blue* layout so the white ADI logo, footer and slide number stay fixed. Cover, section and end card keep their template layouts.
- **One accent:** ADI light blue `1B9CD0`, only for hero numbers, the key phrase, the subtitle line and small labels. White type; ADI grey `9EA1AE` for secondary lines.
- **Elements:** short white titles with one accent subtitle line; dark charcoal cards with a slim accent header tab and Brand Icon Library line icons in accent circles (three or four per slide at most); faint white hairlines at 12% opacity; generous empty space.
- **Words:** one message per slide; at most four points of six words; no full stops or commas on slides (sentences go in the notes); lead with verbs; a punchy takeaway line on key slides; frame positively ("to be defined", not "missing").
- **Motion:** Morph "by object" on every slide (1.0 to 1.25 s), builds Fade or Float In upwards at about 0.5 s staggered, chapter dividers as a two or three word statement or hero number. Travelling objects are `!!named`, unique per slide, built as slide pairs; canvas and logo never move.
- **ADI rules still win** on fonts (Barlow), palette, logo, footers and classification, the 14 pt floor, Title Case, UK English, approved imagery and icons, template container geometry and the content guardrails. The conflicts are settled in `references/keynote-style.md` section 6.
- The previous light look remains available on request: `CleanDeck(style="light")` in Route C, `--theme private/theme-light.json` in Route B, light layouts in Route A. Use it for appendix tables, print and PDF-first leave-behinds.

| Route | Use when | How |
|---|---|---|
| **C. Clean click-through (default for conversation decks)** | Informal or introductory meetings, university and customer conversations, short updates, anything where slides back a discussion (about 5 to 8 slides, 10 to 20 minutes). The user wants minimal text, content revealed point by point, Morph transitions, speaker notes with sources | `scripts/clean_deck.py` (`CleanDeck`), starting from `scripts/example_clean_deck.py`. Style and rules: `references/clean-click-through.md`. Reusable prompt: `references/prompt-template.md` |
| **A. Template-native** | Short everyday/internal decks (about 10 slides), conversions of old decks, org charts, anything needing the template's own layouts, library patterns, icons or photos | python-pptx directly on the template layouts. Steps in "Route A" below |
| **B. Deck builder** | Customer pitches, design-in meetings, product intros, deep dives, demos, account reviews: where story, notes, claims-checking and lint matter | Brief JSON -> `scripts/build_deck.py` -> lint. Steps in "Route B" below |

All routes draw on `assets/ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx`, so the brand rules below apply to all. Route B reads `private/theme.json`, the keynote theme on that template (Barlow fonts, ADI palette, dark canvas, template masters, logo, classification footer). If the user's request is ambiguous: C for conversation or meeting decks, B for sales pitches and design-ins, A for static internal decks and conversions.

### Route C in brief (the approved click-through build, in keynote style)
Read `references/keynote-style.md` and `references/clean-click-through.md` first, then `references/morph-patterns.md` for the hand-crafted Morph techniques (make-room pairs, focus-row accordion, parking off-canvas, arrows and paragraph builds, slide links to a hidden appendix; `scripts/example_morph_patterns.py` demonstrates each). Non-negotiables: **keynote look (`CleanDeck` defaults to `style="keynote"`: dark canvas, one accent, hairlines; `d.card`, `d.icon`, `d.hero` for at most three or four elements per slide; no bars, strips, chevrons or highlight rings); every point appears on its own click (about 7 clicks per slide, about 30 per 10-minute deck); Morph transition on every slide; notes start with a `CLICKS (n)` cue line, then script, sources and `[TO CONFIRM]` cautions; never over-promise (ideas for discussion, no implied funding, equipment, vacancies or approvals); UK English; Public footer unless told otherwise.** Workflow: (1) search the user's source folders and read what is relevant, (2) outline slides with the click plan, (3) copy `scripts/example_clean_deck.py`, replace the content, run it, (4) fix any `CHECK` lines, (5) deliver under a new version name if the file is open, and report clicks, `[TO CONFIRM]` items, sources used and that no visual render was done if so. The user can paste `references/prompt-template.md` to get this every time.

## Brand rules (from the template's own guidance slides)
- Fonts: **Barlow Medium** (titles) and **Barlow** (body). They come from the theme; never name other fonts on slides.
- Body text sizes (Standard template): level 1 = 18 pt, no bullet; level 2 = 16 pt; level 3 = 14 pt. Tables 14 pt (row height 1 cm minimum). 14 pt is the floor for any text except the footer. **Large Format Venues template** (auditoriums, stages, keynote halls only): 22 / 20 / 18 pt, fewer words, mix of dark and light by room lighting; never use it for everyday decks or trade shows, and never use the Standard template on a stage.
- Typography: flush left, ragged right; weight before colour for emphasis; standard letter spacing; no effects.
- Slide numbers, copyright and classification: see Footer below. Trademark symbols once per slide.
- Titles use Title Case (articles, conjunctions and short prepositions lowercase unless first word or over four letters). A cover title longer than 3 lines may shrink, never below 36 pt.
- **Do not move, resize or restyle** the template's text containers. Fit content inside them.
- Backgrounds: the keynote dark canvas on the Dark Blue layout is the house default (see "House look"). Light layouts are for appendix tables, printed or PDF-first copies, or when the user asks.
- Palette only (theme colors): dk2/primary blue `0067B9`, accent1 dark blue `00325C` (keynote canvas tint), accent2 light blue `1B9CD0` (keynote accent), accent3 `8637BA`, accent4 `179963`, accent5 `FED141`, accent6 `C81A28`, gray `9EA1AE`, dark grey `3C4157` (keynote cards). Route A references theme colors, not hex. Keep contrast high; on the dark canvas use white and `1B9CD0` for text, never `0067B9` or `00325C`.
- Max ~10 slides / 20 minutes for everyday decks. One message per slide, visuals over text. No clip art, sound effects or gratuitous animation (keynote Morph and 0.5 s builds are purposeful motion, not gratuitous). The look is keynote quality, but the Standard template itself is not meant for auditorium stages: for those, check with ADI brand first.
- Large, accessible charts. Use native charts, theme accent colors, 14 pt labels.
- Language: follow the user's spelling (this skill defaults to UK English; ADI's own Editorial Standards say U.S. English, so flag that on external decks). Units without a space (`1V`), no spaces around slashes, "Analog Devices" first then "ADI". Wording limits (no "partnership" for a specific relationship, no "sustainable" or "eco-friendly", dated and sourced claims) are in `references/compliance.md`.
- Footer text must match the classification (set on every content slide, keep the slide number; never append extra text to it):
  - Public: `©2026 Analog Devices, Inc. All Rights Reserved.`
  - Confidential (NDA): `Analog Devices Confidential Information. ©2026 Analog Devices, Inc. All Rights Reserved.`
  - Internal only: `Analog Devices Confidential Information—Not for External Distribution. ©2026 Analog Devices, Inc. All Rights Reserved.`
  If the user does not say, ask which classification applies; default to Public only for obviously public content. In Route B set `"meta": {"classification": "public|confidential|internal"}` in the brief.
- Images: only from the Brand Image Library or brand-approved photography. Do not invent or download images.

### Motion in the deck builder (keynote style)
The builder was written for keynote-style pitching, and the house look now uses it that way:
- **Motion defaults to `standard` with Morph as the base transition** (1200 ms, by object; fade through black on the opener; cut on the appendix; click builds at 500 ms). Use `minimal` for board-level rooms or laggy video calls, `showcase` only if the user asks. No loud transitions.
- Morph is on every slide, but design the pairs: `!!name` the objects that should travel (hero number, hub, cards). Lint's `build.morph_unpaired` WARN is expected on slides where everything is new; act on it only where you meant objects to travel.
- Main-flow slides carry few words, but never drop below 14 pt to achieve it.
- Builder layouts draw on the template's "Title and Content - Dark Blue" layout (dark slides, canvas painted over it; theme key `blank_layout_dark`) or "Title Only – Blank" (light slides) and place their own text boxes; footer and slide number are cloned from the layout. If a user wants strict placeholder-driven titles (accessibility, outline view), use Route A.
