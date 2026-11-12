# Other deck types

The customer pitch arc is the default. These patterns reuse the same layouts, notes and lint, with a different spine. Set `role` on slides where the default role does not fit, so lint checks the right things; expect some story FAILs (no criteria, no reveal) on decks that legitimately skip those beats, and say so to the user rather than padding the deck.

## Technical deep dive (design review, architecture session)

Audience: engineers who already want the part. Goal: confidence it will work in their design.
1. `statement` (hook): their design goal and the one risk they worry about most
2. `roadmap` (agenda): the signal chain as stages (sensor -> AFE -> ADC -> processor), cover optional
3. Per block: `feature` or `hardware` with the key spec translated, then `side_by_side` or `bars` with measured data
4. `objections`: known limitations and mitigations (candour buys credibility)
5. `ask`: design review date, schematic review, samples
6. Appendix: full tables, test conditions, layout guidelines, errata

Denser is acceptable here, but still one point per slide; put full plots in the appendix.

## Product demo

Goal: they see it working and believe it is ready.
1. `statement`: what they will see and why it matters to them
2. `criteria`: what "working" means (latency, accuracy, power)
3. `hardware`: the demo setup, annotated
4. Live demo (a `section` slide titled "Live" as the cue), with `notes.fallback` pointing at a recorded backup slide
5. `big_number` / `bars`: the result in their units
6. `ask`
Rules: pre-load the demo state; rehearse on the actual laptop; have a recording one click away; never debug live.

## Account review / QBR (customer or internal)

Goal: a decision or alignment, usually from managers. Here an assertion-headline style is right: each slide's words state the conclusion ("Design-in at risk: samples late by 3 weeks"), because it may be read without you.
1. `statement`: the one-line summary and the decision needed
2. `bars` / `big_number`: the headline metric vs target
3. `status_quo`: what is going well / at risk (rows)
4. `roadmap`: the plan
5. `ask`: the decision, owner, date
Appendix: pipeline detail, backlog, forecasts.

## Internal update or proposal (to a manager or team)

Pyramid structure: answer first, then support. Slide 1 says what you want and why; then three supporting points (`criteria` as the three reasons, or one `feature` each); then risks (`objections`); then the ask. Keep to 5 to 7 slides.

## Training / enablement (e.g. a product family for colleagues or distributors)

1. `section` per topic, `roadmap` as the agenda
2. For each concept: `statement` (the idea), `hardware` or `feature` (the example), `big_number` (the number to remember)
3. A recap slide (bullets are fine here, as a recap)
4. Quick-reference `table` in the appendix
More text on slides is acceptable when the deck will be read afterwards; build a `--static` copy for distribution.

## Conference / external talk

Closest to the Jobs model: one idea, strong hook, story, one demo or proof moment, a memorable close. Remove ADI-confidential content; check with marketing before anything external.

## Leave-behind / pre-read

A deck sent without a presenter must carry its own narrative: assertion headlines, a support line on each slide, conditions in visible footnotes rather than notes. Build from the same brief with `--static`, then raise word limits deliberately (put the notes' key line into `support` fields). Tell the user this is a different artefact from the presented deck.
