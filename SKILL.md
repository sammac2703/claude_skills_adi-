---
name: adi-presentation
description: Build PowerPoint decks on the official Analog Devices (ADI) 2026 brand template in the keynote house style (dark black-to-ADI-dark-blue canvas, ADI light blue as the one accent, minimal words, cards and icon circles, Morph on every slide). Use whenever the user asks for an ADI presentation, slides, or deck, or wants content put on the ADI template, or asks to convert an old deck to ADI branding. Short informal or conversation decks use the clean click-through build (one point per click, Morph on every slide, notes with click plan). Also covers customer pitches, design-in decks, product intros, technical deep dives, demos, account reviews and roadmaps (story arc, minimal on-slide text, speaker notes with conditions, appendix, Morph motion) and linting or restructuring an existing deck. Also covers executive and partner briefings, training and workshop decks, corporate overviews, large-venue keynotes, leave-behinds and old-deck conversions, each with its own recipe, while keeping Morph and the build timings identical and varying composition, look rhythm and imagery from deck to deck.
---

# ADI presentations (router: read only what the task needs)

Brand template, Slide Library layouts and icons, Morph or fade motion, notes, QA gate. Classify first, build with scripts, gate before delivery.

## 1. Start (always)
1. **Classify** with `references/use-cases.md` (intake plus one recipe: conversation, pitch, executive briefing, deep dive, training, launch, demo, account review, internal update, corporate overview, keynote, leave-behind, conversion, org chart, ADEF, media). Say the recipe in one line. Save intake answers once in `brief.json` next to the deck and reuse them for later versions.
2. **Constants vs variation** (`references/variation.md`): always the same = 0.5 s float-up builds, `!!named` travelling objects, fixed canvas and logo, Barlow, ADI palette, one accent, 14 pt floor, footer classification. Varied per deck = motifs, look rhythm (dark / light / mixed), sequence. Transition: Morph for slide pairs that share objects; a 0.7 s fade when neighbouring layouts differ a lot; cut for the appendix.
3. **Titles are claims**: outline titles first and check the title-only read (`references/storyline-checks.md`); ghost deck for decks over about 8 slides.

## 2. Build (pick one path)
| Path | Use for | Do |
|---|---|---|
| **Spec build (default)** | Most decks | Write a short JSON spec, run `python scripts/build_from_spec.py spec.json out.pptx`. Slide types and fields: `references/spec-format.md` |
| **Blocks in Python** | Custom compositions | `scripts/adi_blocks.py` (library layouts, icons, logo, numbers), `scripts/adi_data.py` (tables, charts), `scripts/exhibits.py` (waterfall, 2x2, Gantt, harvey table, funnel). Detail: `references/blocks-detail.md`, `references/library-layouts.md` |
| **Route C click-through** | Conversation decks of 5 to 8 slides, one point per click | `scripts/clean_deck.py`, `references/clean-click-through.md`, `references/prompt-template.md` |
| **Route A / B** | Conversions, org charts, customer pitches with brief + lint | `references/routes-a-b.md` |

Find a layout or icon without reading catalogues: `python scripts/find_layout.py "four numbers"`, `python scripts/find_icon.py "team"`. Layout facts (dark or light, placeholders, logo colour) live in `assets/layout_facts.json`.

## 3. Rules that always apply (details in `references/house-look-and-brand.md`)
Barlow; ADI palette only (`1B9CD0` is the one accent); Title Case titles, no full stops or commas on slides; one message per slide; at most four points of six words; 14 pt floor (footer, legal line, slide number exempt); Standard template unless a stage keynote (Large Format); US vs UK spelling as the audience needs (flag it); units without space; claims dated and sourced; no banned words, "partnership" only for real partnerships (`references/compliance.md`); images only from the Brand Image Library; icons always white on blue or dark circles and backgrounds (never black or dark grey), ADI blue on light; every library-layout slide gets `b.brand(slide, dark)` (logo top right, slide number); fewer boxes (`b.open_up`, hairlines); exhibits follow `references/exhibit-rules.md`.

## 4. Before delivery (script decides, not you)
`python scripts/qa_layout.py deck.pptx` (ERRORS fail; also title length, repeated layouts, alignment and numbers-in-notes checks), `python scripts/lint_deck.py deck.pptx` for pitches. Deliver only when it prints PASSED (it writes no file). **Output rule: leave the .pptx alone in the user's folder** (QA and previews go to the terminal or a temp dir, never beside the deck). Preview without PowerPoint: `python scripts/preview_png.py deck.pptx outdir` and look at the PNGs. Regression: `python tests/regression.py` after changing any script.

## 5. Report
Use case, recipe, motifs, look rhythm, clicks, `[TO CONFIRM]` items, sources, QA result, and whether a real render was seen.

## Reference map
`keynote-style.md` look · `house-look-and-brand.md` brand and motion rules · `layouts-and-assets.md` template layouts and assets · `library-layouts.md` library idx map · `exhibit-rules.md` charts and exhibits · `spec-format.md` spec build · `storyline-checks.md` titles, review, QA gate · `compliance.md` editorial rules · `variation.md` · `use-cases.md` · `story.md`, `brief-format.md`, `deck-types.md`, `design.md`, `motion.md`, `morph-patterns.md`, `clean-click-through.md` (Routes B and C) · `new-slides-index.md` · `adi-template.md`.
