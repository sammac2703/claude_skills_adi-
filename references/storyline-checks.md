# Storyline checks and the QA gate

Ideas adapted from consulting-deck skills (action titles, ghost decks, partner review) and from the Mck QA gate, fitted to ADI rules (Title Case, no full stops on slides, one message per slide, 14 pt floor). Use with `story.md` for pitches and `use-cases.md` for the recipe.

## 1. Action titles (titles that say the point)
A title is a claim, not a topic: "Customer Visits" says what the slide is about; "Seven Visits Put Interns in Front of Customers" says what it proves.
- One message per slide. If the title needs "and", it is two slides or one belongs in the appendix.
- Short, with a verb, Title Case, no full stop, no more than about 8 words. The proof (numbers, conditions) goes in the body or notes.
- Keep claims dated and sourced (`compliance.md`). A title must never promise more than the body proves.
- Conversation decks may keep short topic titles when the presenter talks the point: choose per use case, and say so in the delivery report.

**Title-only read.** Write the titles in order and read only them. They must tell the story and answer the audience's question without the bodies. Where the read jumps, a slide is missing or in the wrong place. Do this before building and again before delivering.

## 2. Ghost deck (optional first step for decks over about 8 slides or when scope is unclear)
Before any design, list every slide with: number, final title, one-line sketch of the exhibit ("three-column comparison", "timeline of five phases"), the library layout or motif planned, the data needed and whether it exists, and a confidence tag (solid, likely, speculative). Show it to the user and build only after they agree. Speculative slides are candidates to cut or move to the appendix. Vary the motif column so no two neighbours repeat (`variation.md`).

## 3. Review order before delivery (story, then titles, then craft)
1. **Story:** one clear answer or ask, up front; the audience's real question is answered; the deck ends with the next step or decision.
2. **Title chain:** the title-only read works; each title matches its slide.
3. **Logic:** groupings do not overlap and are complete; no orphan slide that supports nothing.
4. **Exhibit fit:** each chart or visual matches its message; one insight highlighted.
5. **Craft:** alignment, one type scale, ADI palette, one accent, icons from one family and one style per slide, no stacked boxes.
6. **Density:** one message per slide; detail goes to the notes or appendix.
7. **Consistency:** same colour and label for the same thing everywhere; same number formats.
8. **Evidence:** every number traceable, claims dated, assumptions stated, nothing invented, `[TO CONFIRM]` items listed.
9. **Compliance:** `compliance.md` (spelling, units, banned words, footer classification, logo, imagery).
Rank findings: blocker (story or logic broken), major (title or exhibit wrong), minor (alignment, colour).

## 4. Template-fill rules (any route that fills layouts)
- Count the content pieces first, then choose the layout with that many slots: two items to two columns, three to three, four or more to a list or split across slides.
- Image layouts only when there is an approved image; quote layouts only for real quotes with attribution.
- Never leave an empty placeholder or invent filler to fill a slot; remove unused ones (`b.tidy`).

## 5. QA gate: the script decides, not the model
Run `python scripts/qa_layout.py deck.pptx` before delivering. It prints PASSED or FAILED and exits 1 on failure; it writes no file. Never write "QA passed" unless the script printed PASSED; fix the ERRORS, rebuild, run it again.
- **ERROR:** shape off the slide; text under 14 pt (footer, legal line and slide number exempt); icon or text too close to the colour behind it (invisible icons on dark layouts); leftover template text; empty placeholder.
- **WARN:** text boxes overlapping; text likely to overflow; no ADI logo or slide number on a content slide; no transition. Read each: overlaps between a layout placeholder and your own text are real, box-only overlaps on tall placeholders can be fine.
- **INFO:** under a third of the content area used: enlarge text or icons.
The gate checks geometry and colour from the file, so it does not replace looking at a render. Rendering: `scripts/render_slides.ps1` (PowerPoint COM; it fails with `CO_E_SERVER_EXEC_FAILURE` when PowerPoint is already open or locked down, and LibreOffice is not installed on this machine). When no render is possible, say so and ask the user to run Slide Show once; record each slide as pass or fail from their feedback.

## 6. Added checks and preview (this session)
- `qa_layout.py` also warns on: text edges that nearly but not quite align (0.02 to 0.15 in), titles over 10 words, the same library layout twice in a row (or three custom slides in a row), and INFO for numbers on a slide that the notes never mention. `-q` hides INFO.
- `python scripts/preview_png.py deck.pptx outdir` draws every slide from the file geometry in Barlow (icons and the logo as outlined boxes in their colour, charts as labelled boxes, no Morph or builds). Open the PNGs with the Read tool: it shows empty halves, crowding, clipping and wrong colours that numbers alone miss. Look at every slide at least once before delivering; it is a layout check, not a render.
- `python tests/regression.py` rebuilds `tests/regression_spec.json` (every spec type), runs the gate and writes previews. Run it after changing any script.

## 7. Output rule: the .pptx alone
A delivery leaves exactly one new file in the user's folder: the .pptx (plus the generator script only if the user asked for it). QA, lint and preview output go to the terminal or a temp folder, never next to the deck: no `.qa.json`, no `.lint.md` (only with `--lint-report`), no `png` folders (`preview_png.py` outdir must be a temp path), no stale copies. If a target is open in PowerPoint, save under a new version name and say so.
