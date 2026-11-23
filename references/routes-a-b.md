# Route A, Route B, rendering, setup and gotchas (moved from SKILL.md)

## Route A: template-native build
1. Clarify only what you cannot infer: topic, audience, length, footer classification. Outline one message per slide, titles that state the message.
2. `python scripts/prep_template.py <template.potx> <work/adi_base.pptx>` converts the template to an openable .pptx and prints every layout and placeholder idx.
3. Build with python-pptx (see `scripts/example_build.py`): delete the guidance slides, add slides by **layout name**, fill placeholders by **idx**, copy the layout's footer [3] and slide number [4] placeholders onto each non-cover/divider slide and set the footer text. Deps: `python -m pip install --user python-pptx lxml defusedxml Pillow`. Do not run with `python -I`.
4. Keynote look: put content on the Dark Blue layouts ("Title and Content - Dark Blue", "Two Line Title and Content - Dark Blue", "Title and Content with Sub Header - Dark Blue") and paint the black to `00325C` diagonal canvas on each slide background (see `CleanDeck.canvas` in `scripts/clean_deck.py`); white text, `1B9CD0` accent, Morph on every slide. Native charts and tables, theme colors only. For visual slides (agenda, status, next steps) build charcoal cards with an accent header tab (as `CleanDeck.card`) or the `box` / `text` / `badge` helpers in `scripts/example_build.py`, three or four per slide at most. Avoid plain title-plus-bullets slides.
5. QA: `python <pptx skill>/scripts/office/validate.py deck.pptx --original <base.pptx>`, then `slide_text.py` to proof the content and grep for leftover template text. Optionally `python scripts/lint_deck.py deck.pptx` for claims/units/competitor checks.
6. Render and look at every slide (see "Rendering").

## Route B: deck builder (story, notes, lint)
The principle: **the slides are a backdrop for the presenter, not a document.** One idea per slide in a few large words; the script and every number's conditions live in the speaker notes; datasheet detail lives in an appendix.

| Path | What it is | When to read |
|---|---|---|
| `references/story.md` | Pitch arc, slide sequence, ADI adaptations | Every customer pitch, before outlining |
| `references/brief-format.md` | Brief JSON format and every layout's fields | Before writing a brief |
| `references/keynote-style.md` | The dominant house look: canvas, accent, cards, icons, words, Morph | Every deck, before outlining |
| `references/design.md` | Type, colour, photos, layout, what not to do | Choosing layouts or hand-building slides |
| `references/motion.md` | Transition grammar, Morph, click builds, levels | Before choosing a motion level above minimal |
| `references/deck-types.md` | Demos, deep dives, account reviews, internal updates, training, conference talks | When not a straight customer pitch |
| `references/use-cases.md` | Intake questions, recipe per use case, delivery report | Every request, first |
| `references/variation.md` | What stays fixed (Morph, builds) and what varies (motifs, look rhythm, imagery) | Every deck, when outlining |
| `references/storyline-checks.md` | Action titles, title-only read, ghost deck, review order, template-fill rules, the QA gate | Before outlining a deck over about 8 slides, and before every delivery |
| `references/compliance.md` | Editorial rules, claims, banned words, logo, imagery, approvals, silent periods | Before delivering; always for external decks |
| `references/new-slides-index.md` | What the 2026-11-23 downloads contain and what to learn from each | When converting or modelling on those decks |
| `references/library-layouts.md` | Slide Library layouts as placeholder maps, `adi_blocks.py` helpers, ADI icon workflow | Whenever slides should use ADI's own building blocks and icons (the default for variety) |
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
4. **Brief JSON.** Copy the shape of `examples/brief_example.json`. At most four points of six words on a main-flow slide, no full stops or commas on slides; `notes.conditions` on every slide with a number; `notes.source` on comparisons; `[CLICK]` markers in `notes.say` only if motion is above minimal; set `meta.classification`. Asset paths are relative to the brief; missing assets render as labelled dashed boxes.
5. **Build.** Theme resolution: `--theme`, else the brief's `theme`, else `private/theme.json` (the real ADI theme, present here), else the placeholder palette (not brand-ready; say so).
6. **Lint and fix.** Fix every FAIL in the brief and rebuild. REVIEW items (competitors, `[TBC]` sources, approximate dates) go to the user as explicit actions; the deck is not presentable until resolved. For a hand-made deck: `python scripts/lint_deck.py their_deck.pptx --competitors "A,B" --require-footer "Analog Devices"` (story-order checks are partial without a `role: ...` first notes line).
7. **Look at it** (see "Rendering"). Lint does not check geometry; overflow and collisions need eyes. Shorten text rather than shrinking it.
8. **Deliver.** The .pptx (and `--static` copy if it will be a PDF), then in a few lines: the one idea, slide count, open items (`[TBC]`, placeholder assets, REVIEW flags), and a note to rehearse in presenter view.

Judgement calls: keep it about the customer; candour about weaknesses buys credibility; superlatives only after evidence with conditions; builds and motion must read as stills because decks get forwarded and exported to PDF.

Going beyond the layouts: compose in Python with `deckkit.Deck` (text, rect, line, picture, `d.tl(slide).add(...)`, `d.morph_as(...)`) and follow the naming rules: `!!name` for Morph-paired objects, `aux_` for labels lint should not word-count, `photo_` for photographs, `ftr_` for chrome. Tag the role with `set_slide_tag(slide, role="proof")`. To edit an existing .pptx rather than regenerate it, use the general `anthropic-skills:pptx` skill, then lint the result.

## Rendering (both routes)
LibreOffice is not installed on this machine; use PowerPoint: `powershell scripts/render_slides.ps1 -In deck.pptx -OutDir <dir> -Width 1280` writes `Slide1.PNG`, … It only works with the shell sandbox disabled (inside it PowerPoint fails with COM error 0x80080005). ADI machines enforce the `RemoteSigned` execution policy, so if the skill came from a downloaded zip the script fails with "is not digitally signed"; the user must run `Unblock-File <skill_dir>\scripts\render_slides.ps1` once themselves (do not bypass the policy). If rendering is not possible, say plainly that visual QA was not done and ask the user to open the file. Delete old images before re-rendering. Morph and animation do not show in a render: check each state-A slide reads as a still and tell the user to run Slide Show once.

## Setup state (this install)
`private/theme.json` was generated by `scripts/extract_template.py` from the 2026 Standard template and hand-tuned into the keynote theme (Barlow fonts, ADI palette, classification footers, dark canvas `000000` to `00325C` at 315 degrees over the Dark Blue layout, accent `1B9CD0`, motion `standard` with Morph base). `private/theme-light.json` is the earlier light theme (light template background, motion `minimal`). Regenerate or edit it if the template changes; `references/adi-template.md` describes the process (its "placeholder palette" caveats are superseded by this file). `SKILL.deckbuilder.orig.md` is the original standalone deck-builder SKILL.md, kept for reference.

## Gotchas
- python-pptx refuses `.potx`; `prep_template.py` and `deckkit.open_template` rewrite the content type in a copy.
- Cover and section layouts get their image or pattern from the layout/master; do not add pictures to them.
- Empty picture placeholders show prompt text in edit view; fill them or use a layout without one.
- Dropping slides with python-pptx: call `prs.part.drop_rel(rId)` and remove the `sldId`, otherwise orphaned slides are saved.
- Scripts must read/write JSON and markdown as UTF-8 (the `©` footer corrupts to "Â©" under the Windows default codepage). When editing files with non-ASCII characters use the Edit tool, not a PowerShell `Get-Content`/`Set-Content` round trip.
