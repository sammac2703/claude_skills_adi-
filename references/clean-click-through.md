# Route C: clean click-through conversation decks

The house style for short, informal, presenter-led ADI decks (introductory meetings, customer or university conversations, internal updates, anything where the slides back a discussion). Proven on the Bristol careers deck. Build with `scripts/clean_deck.py`; start from `scripts/example_clean_deck.py`.

## The look (keynote style, the default: see `keynote-style.md`)
- **Dark canvas.** `CleanDeck()` defaults to `style="keynote"`: every content slide uses `Title and Content - Dark Blue` with the black to ADI dark blue diagonal canvas; the template's white logo, footer and slide number stay fixed. `Cover with Pattern` for the opener.
- **One accent.** White type; ADI light blue (`TC.ACCENT_2`) for the subtitle line (`d.sub`), numerals, hero numbers, key phrases and the closing takeaway. Light-style colour roles are remapped automatically (`TEXT_1`, `ACCENT_1` to white; `TEXT_2`, `ACCENT_3` to the accent), so older slide code still reads.
- **Few elements.** Bold labels with 12% white hairlines (`d.rule`), pale numerals, and where they help: up to three or four charcoal cards with an accent header tab (`d.card`), Brand Icon Library icons in accent circles (`d.icon`), or one hero number (`d.hero`). No accent bars, header bands, banner strips, chevrons, chips or highlight rings.
- **Words:** at most four points of six words, no full stops or commas on slides, lead with verbs; sentences go in the notes. Template fonts only; 14 pt minimum, 18 pt and up for body text.
- **Light look on request:** `CleanDeck(style="light")` restores the original text-only light style (`Title Only` layout, `ACCENT_1` labels, grey rules, `ACCENT_3` status line, Dark Blue layout only for `d.slide(..., dark=True)`).
- **Titles** in Title Case, one line, stating the message. UK English. Footer set by classification (default Public).
- **Photos:** none unless asked; the image library is licensed for cover and section slides only.

## The build (this is the point of the style)
- **Nothing appears at once.** Title, footer and the one-line intro (`d.sub`) show first. Every point then arrives on a click.
- **One idea per click.** Lists with more than about 4 short items either land in sequence on one click (items 220 ms apart) or get a click each if each is a talking point (ideas, questions, chain stages). The last click is the closing line or guardrail.
- **Effects:** text uses `float_up` (500 ms, 220 ms stagger); rules `wipe` in from the left (500 ms). No emphasis rings, pulses, spins or exits.
- **Budget:** at most about 7 clicks on a slide and about 30 for a 10-minute deck. If a slide needs more, group items.
- **Morph on every slide** (`d.finish`, 1.2 s, by object). The intro line is named `!!sub` on every slide so it glides; the rest builds on click. Morph is a smooth hand-over, not a substitute for the builds.
- **Make at least one hand-crafted Morph moment per deck** where it helps the story, using the state-pair patterns in `references/morph-patterns.md`: a make-room pair (a list steps aside for new content), a focus row (accordion / spotlight), arrows that land with their step, or a link to a hidden appendix slide. Use them only where they carry meaning, not on every slide; give state pairs 1.6 to 2.0 s. Two or three is plenty in a 10-minute deck.
- **Appendix:** detail the presenter may need but should not show goes on a hidden slide (`show="0"`) reached by a link and left by a Back link, with a `cut` transition.

## Notes (every slide)
1. First line: `CLICKS (n): [1] ... [2] ...` (written automatically by `d.build`), so the presenter knows what each click shows.
2. Then the script: timing, what to say, when to stop talking and ask.
3. Then sources and cautions, with `[TO CONFIRM]` on every name, date, figure or detail not in the source material.

## Content guardrails
- Never over-promise. Frame everything as ideas for discussion unless the user says otherwise. No equipment, funding, placements, vacancies, recruitment routes, budgets or sponsorship implied or offered. State guardrails on the slide when it matters ("We would only promote roles that are confirmed and approved").
- Use only facts from the source material in the working folders; omit dated statistics (say so in notes); never invent. Older or internal documents are format references only.
- Flag privacy (student data), brand/legal review for external speakers and rights for any imagery in the notes.
- Overseas programme documents are reference models only.
- Competitor names, comparative claims, customer names: flag for review.

## QA before delivering
1. `d.finish` prints click counts and a rough text-fit check; fix any `CHECK` lines (shorten text first).
2. `python scripts/lint_deck.py deck.pptx`: for this style ignore the sales-arc story checks, the claims.conditions numerals false positives and the dense-slide warnings that come from build text; act on everything else.
3. Rendering needs PowerPoint outside the sandbox. If it cannot be done, say plainly that no visual check was done and ask the user to open the file and run Slide Show once (it also confirms click counts and Morph).
4. If the target .pptx is open in PowerPoint the save fails with a permission error: save under a new version name instead of overwriting.
5. Report: file path, slide count, click count and plan, every `[TO CONFIRM]`, what sources were and were not used.
