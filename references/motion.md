# Motion: transitions, Morph and animation like a pro

Contents: 1 The principle · 2 The motion system (levels) · 3 Transitions · 4 Morph techniques · 5 In-slide animation · 6 Timing numbers · 7 What the engine does for you · 8 Brief fields · 9 Avoid · 10 Checking it

## 1. The principle

Professional decks use a lot of motion, but almost none of it is decorative. Motion does three jobs:
1. **Pacing:** the presenter controls when each idea appears (click builds), so the room listens instead of reading ahead.
2. **Continuity:** the eye follows one object as it changes (Morph), so a comparison, a zoom or a before/after reads as *one* thing changing, not two slides.
3. **Attention:** one emphasis on the one number that matters.

The test used by presentation agencies: if the audience remembers the transition, it was the wrong choice. So the deck uses **every** PowerPoint capability the engine has, but through a small grammar, consistently. A deck that mixes Vortex, Origami and Curtains looks amateur; a deck with one quiet base transition, Morph for the moments that move, and click builds with fades and wipes looks expensive.

## 2. The motion system

Set with `"motion": "minimal" | "standard" | "showcase"` in the brief (default `standard`), or override any field: `"motion": {"level": "standard", "base": "push", "base_ms": 500}`.

| | minimal | standard (default) | showcase |
|---|---|---|---|
| Base transition | Fade 300 ms | Fade 400 ms | Fade 450 ms |
| First slide | Fade 500 ms | Fade through black 1000 ms | Fade through black 1400 ms |
| Section dividers | Fade 400 ms | Fade through black 800 ms | Smooth Reveal (through black) 1400 ms |
| Appendix | Cut | Cut | Cut |
| Morph pairs | Yes | Yes | Yes |
| Click builds and staggered entrances | No | Yes | Yes |
| Ken Burns push-in on photos | No | Hook photo only | Hook and credibility photos |

Use `minimal` for board-level or very senior rooms and anything presented over a laggy video call; `standard` for customer meetings; `showcase` for conferences, launches and recorded videos. `--static` (for PDF and email) removes Morph state slides and all animation.

## 3. Transitions

All 58 PowerPoint transitions are available (`motion.TRANSITIONS`): the 2007 set (fade, fade_black, push, wipe, split, cover, uncover, zoom, shapes, blinds, checker, comb, dissolve, random_bars, strips, newsflash, random), the 2010 set (reveal, pan, vortex, ripple, honeycomb, glitter, shred, flash, switch, flip, gallery, conveyor, ferris, doors, window, warp, flythrough, cube, box, rotate, orbit), the 2013 set (curtains, wind, prestige, fracture, crush, peel_off, page_curl, page_curl_single, airplane, origami, fall_over, drape) and Morph (morph, morph_words, morph_chars). They are written exactly as PowerPoint writes them, inside `mc:AlternateContent` with a fade fallback, so older PowerPoint versions degrade to a fade instead of failing.

What to actually use:
- **Fade** as the base. It is the industry standard for business decks.
- **Fade through black** for the opening and for topic changes (it reads as a "chapter").
- **Push** (up or left) when slides are a sequence the audience should feel moving forward, e.g. a multi-slide walkthrough. Pick a direction and keep it.
- **Wipe** only when the content itself is directional (left-to-right process).
- **Smooth Reveal** (`reveal`) for section breaks in showcase decks.
- **Morph** whenever two consecutive slides share objects (section 4).
- **Cut** (no transition) for appendix and reference slides, and when pace matters more than polish.
- The rest (vortex, origami, curtains, page curl, fracture, ...) exist so you can match an existing deck or make a deliberately playful internal one. Lint warns on them in customer decks.

Per slide: `"transition": "push"` or `"transition": {"kind": "push", "dir": "l", "ms": 600}`.

## 4. Morph techniques

Morph animates every object that appears on both slides: position, size, rotation, colour, transparency, font size, crop and 3-D rotation. Objects only on the second slide fade in; objects only on the first fade out. `!!name` pairs objects explicitly; the engine names everything that should pair.

| Technique | How the engine builds it | Layout / field |
|---|---|---|
| **Camera zoom into a diagram** | The whole diagram is redrawn at a new scale and offset on each stop; text, outlines and arrows scale with it. Morph flies the audience in. | `zoom_tour` |
| **Infinite zoom** | Blocks contain sub-blocks drawn tiny in the overview (5 pt text); zooming in makes them legible. | `zoom_tour` with `inside` |
| **Zoom back out** | The overview repeats at the end so the detail lands in context. | `zoom_out` (default true) |
| **Pan and zoom across a photo** | The board photo is scaled 2.4x and re-centred on each callout; a focus ring morphs from point to point; a side panel carries the label. | `hardware` with `zoom: true` |
| **Build a table row by row** | Rows parked off-canvas at staggered distances slide in one per click. | `status_quo` |
| **Fill the empty corner** | Object parked off the slide moves into the 2x2. | `quadrant` |
| **Grow bars** | Zero-height bars on slide A, full height on B. | `bars` with `grow` |
| **Three become one** | Tease circles shrink into dots under the product name. | `tease` then `reveal` |
| **Word swap** | Morph *by Words*: words that stay glide to their new place; words that change cross-fade. "Three chips. 41 mW." becomes "One chip. 8 mW." | `statement` with `then` |
| **Number change** | Morph *by Characters*: digits slide and swap. "3 days" becomes "14 days". | `big_number` with `from` |
| **3-D turn** | Product is flat on the reveal and tilted in 3-D on the next slide; Morph rotates it. | `product` (default `tilt`) |
| **Staggered fly-in** | New objects parked off-canvas at increasing distances, so further ones travel faster (multi-speed entrance). | `enter` on any slide |
| **Expand from a cover** | Roadmap columns start as stripes off the left edge and open out. | `roadmap` with `cover` |

Rules: Morph only between slides that are the same scene in two states. Keep 1000 to 2000 ms (1500 default; camera moves 1600 to 1800). Never Morph between unrelated slides; it produces chaotic motion.

## 5. In-slide animation

Built per layout with `Timeline` (`motion.py`). Every effect PowerPoint offers through its animation pane can be written: entrance (appear, fade, wipe, fly, float_up, zoom, split), emphasis (pulse, grow, ken_burns, spin, dim), exit (fade_out, wipe_out) and motion paths (`move`), with On Click / With Previous / After Previous, delays, smooth start and end, and text by paragraph, word or letter.

What the layouts do by default (standard level):

| Layout | Opens with | On click |
|---|---|---|
| title, mosaic, section | Kicker fades, title floats up 0.25 s later, subtitle follows (one cascade, With Previous plus delays). Mosaic photo slowly pushes in. | |
| statement | | `build: "lines"`: one line per click. Support line on click |
| credibility | Number floats up, unit and line follow | |
| objections | Concerns visible | Each answer wipes in from the left: voice the doubt, then answer it |
| criteria | Kicker only | One yardstick per click (get a nod on each) |
| quadrant | Axes draw themselves (wipe), labels fade | Each alternative pops in; then the empty target corner. Next slide: ours moves in (Morph) and pulses once |
| bars | | With `grow`: Morph grows bars, ours pulses. Without: each bar wipes up from the baseline |
| reveal | Name arrives with the Morph; definition fades 0.4 s later | |
| feature | Glyph zooms in | Support line |
| product | Product turns in 3-D (Morph) | One point per click, floating up |
| hardware | | Each callout: dot pops, line draws towards the label (wipe), label fades in |
| roadmap | Morph expands the stages | Each stage's detail wipes in from the left |
| big_number | Old number Morphs into the new one by character, then pulses | |
| side_by_side | Left side | Right side, then caption |
| ask | Kicker, ask, owner, date cascade in | |

The engine writes the click count and a one-line cue per click into the speaker notes (`CLICKS (3): 1) ... 2) ...`), so the presenter knows how many clicks a slide holds.

## 6. Timing numbers

- Entrances: 300 to 600 ms (the engine uses 400 to 600). Exits: about 300 ms.
- Stagger inside a cascade: 150 to 300 ms between items.
- Ordinary transitions: 300 to 600 ms. Section and opening fades: up to about 1.4 s.
- Morph: 1000 to 2000 ms.
- Ken Burns: 9 to 12 s, scale 1.06, so slow nobody consciously notices it.
- Smooth end (deceleration) on anything that moves: fly, float, zoom, motion paths.
- Maximum: about 6 clicks per slide (split the slide beyond that), and two or three pulses in a whole deck.
- Clicks, not timers, for anything the presenter talks over. Auto-play (With/After Previous) only for cascades that complete in under about 1.5 s.

## 7. What the engine does for you

- Names Morph pairs, parks off-canvas objects at staggered distances, adds Morph where pairs exist, and picks Morph by words or characters where the layout calls for it.
- Applies the level's transition grammar by slide role.
- Writes p:timing exactly as PowerPoint structures it (tmRoot, mainSeq, click groups, after-previous chains, bldLst), so the Animation Pane shows normal, editable effects.
- Strips everything for `--static`.

## 8. Brief fields

Deck level: `"motion"` (level string or object). Slide level: `"transition"`, `"enter"`, `"enter_only"`, `"stagger_in"`, `"morph_ms"`, plus layout fields: `statement.build` / `then`, `big_number.from` / `from_label`, `hardware.zoom` / `zoom_scale` / `zoom_out`, `product.tilt` / `click_points`, `side_by_side.build`, and the whole `zoom_tour` layout. See `brief-format.md`.

## 9. Avoid

- Mixing transition styles; more than one "loud" transition in a customer deck.
- Animating titles, logos and footers on every slide.
- Bounce, spin, swivel, random bars, boomerang.
- Timed builds the presenter has to race.
- Morph between unrelated slides.
- Long durations: a 2 s fade on every build adds minutes of dead time to a deck.
- Exit animations in a deck that will also be exported to PDF (the PDF shows every object).

## 10. Checking it

LibreOffice renders the stills but not the motion. Before presenting, open the deck in PowerPoint (Microsoft 365 or 2019+ for Morph), run Slide Show from the start once and check: every click advances something, nothing appears early, Morph pairs travel rather than cross-fade, and the click counts match the notes. Run it on the presenting laptop and projector; motion that is smooth on a new Mac can stutter on an old projector link.

Sources: Microsoft, "Morph transition: Tips and tricks" (support.microsoft.com); Deckary, "PowerPoint Transitions: The Professional's Guide to Slide Flow" and "PowerPoint Animation: A Practical Guide for Professionals"; BrightCarbon animation guidance; Slidor, "PowerPoint animation tutorial: professional motion effects"; SlideModel, "PowerPoint Morph Transition Explained"; Indezine, "Morph transition with words / characters"; LibreOffice OOXML transition import and export source (oox, sd/filter/eppt) and Microsoft Open XML SDK docs for `p15:prstTrans`.
