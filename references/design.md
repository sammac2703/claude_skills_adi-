# Design layer: the visual system

Contents: 1 Content rules · 2 Type · 3 Colour · 4 Photos, icons, product shots · 5 Layout and finish · 6 Charts and data · 7 Motion · 8 Accessibility · 9 Avoid · 10 Generatable vs manual

## 1. Content rules (from the keynotes)

- One idea per slide: a year, a word, a product, a number.
- Few words, set large: typically 1 to 6 on main-flow slides.
- Hero images and simple icons do the work a bullet list would.
- Dark, clean backgrounds so products and text stand out (light theme for appendix tables and printed copies).
- Numbers as anchors, not tables, in the main flow. Tables live in the appendix.
- Comparison in one picture: a 2x2, a lineup, a two-bar chart.
- Builds: elements appear, merge or move on cue so the screen moves with the words. Comparison tables build row by row.
- Lists only as proof, late in the flow, backing a claim already made. Bullets recap a section; they never script it.
- Annotation callouts for hardware tours: thin white lines and short labels pointing at parts. Maps directly onto eval boards and pinouts.

## 2. Type: hierarchy by weight

Theme font roles: `heavy` (huge titles, stage letters, big numbers), `head` (bold subheads, kickers), `body`, `light` (body that must not compete).

| Element | Size (13.33 x 7.5 in slide) | Role |
|---|---|---|
| Big number / stage letter | 110 to 180 pt | heavy, accent colour |
| Hero title / statement | 48 to 66 pt | heavy |
| Benefit line | 40 to 48 pt | heavy |
| Subhead | 22 to 28 pt | head, often accent |
| Body / support | 17 to 24 pt | light, muted colour |
| Kicker | 13 to 15 pt, caps, letter-spaced | head, accent |
| Footer, captions | 9 to 12 pt | light, muted |

Fonts come from the ADI template. If the template font has a Light weight, put it in `fonts.light`; the light/heavy contrast is what makes the slides look designed. Never hard-code the videos' fonts (Avenir Next, Bebas Neue).

## 3. Colour

- One dominant colour (the background), one accent, one muted text tone, plus the ours/rival pair for comparisons. Accent is for the thing you want seen first, not decoration.
- `ours` vs `rival`: ADI's option in accent, alternatives in a neutral grey. Do not colour competitors red; it reads as an attack.
- Colour sequence encodes order: dark to light means first to last (roadmap stages, computed from `ramp_start` to `ramp_end`).
- Text colour on a filled column is chosen by luminance automatically (white on dark, dark on light).
- Theme variants come from swapping tokens (dark, light, template); no slide should need hand-recolouring.

## 4. Photos, icons, product shots

- Application photos in greyscale, so they never clash with brand colour; add a ~45% transparent colour overlay where text sits on them.
- Product renders and eval boards as **pre-cut transparent PNGs** (PowerPoint's Remove Background cannot be scripted). Leave them in colour; give them a soft shadow (about 55% transparent, 16 pt blur).
- Icons recoloured to the palette; decorative watermark icons at about 75% transparency.
- Same scale for side-by-side comparisons (the Retina rule): if one image is zoomed, the comparison lies.
- Never use photos of real people or competitor product images without clearance.

## 5. Layout and finish

- 0.8 in outer margin; equal-column grids (N columns of width / N, one idea per column); rule of thirds for single focal points.
- Left-align text; centre only short labels under centred visuals.
- Consistent title position across slides of the same layout.
- Gradient backgrounds give depth without decoration: on ADI the house canvas is black to ADI dark blue `00325C`, diagonal (see `keynote-style.md`).
- Inner shadow (inside top-left) for the layered paper look on the mosaic cut-out.
- Footer, slide number, logo and legal line come from the template master (mode `template`), or from the theme's footer when painting our own backgrounds.

## 6. Charts and data

- One message per chart; put the message in the slide words, not a chart title like "Results".
- Two bars beat a ten-bar chart in the main flow (`bars`). Full data goes in the appendix.
- Label values directly; drop gridlines, legends and axis clutter where a direct label works.
- Same axis scale when comparing; start bars at zero.
- Units on the slide, conditions in the notes, raw data in the appendix.
- Measured plots from the lab (PNG) go in `side_by_side` or `feature` with the conditions in notes; crop to the region that makes the point and say what to look at.

## 7. Motion

The full motion system (transition grammar, Morph techniques, click builds, timings, what to avoid) is in `motion.md`. In short: one quiet base transition (fade), fade through black for openings and chapters, Morph for every moment where one thing changes into another, click builds with fades and wipes so the presenter sets the pace, and at most two or three emphasis pulses in a deck.

Morph mechanics the engine relies on:
- **State-based, not animated.** Build the finished slide (B); slide A has the same objects collapsed, moved or parked off-canvas. Morph on B does the motion.
- **Stable names.** `!!name` forces Morph to pair objects across slides; lint fails any slide that shares `!!` names with its predecessor but has no Morph.
- **Off-canvas staggering.** Parked objects sit at increasing distances; further away travels faster, giving a multi-speed entrance.
- **PDF and forwarding.** Neither Morph nor animation survives PDF export; every slide must read as a still, and `--static` drops build steps.
- **XML.** Transitions are written as PowerPoint writes them: `mc:AlternateContent` with the 2010/2013/Morph element in the choice and a fade fallback, after `p:clrMapOvr`; animations as `p:timing` with a `bldLst`.

## 8. Accessibility and robustness

- Contrast: body text at least 4.5:1 against its background; muted text is for secondary lines only.
- Minimum on-screen size 14 pt for anything the audience must read (footers excepted).
- Do not encode meaning by colour alone: ours vs rival is also labelled.
- Fonts: if the template font is not installed on the presenting machine, PowerPoint substitutes; check on the actual laptop and projector.
- Keep the file small: compress photos to roughly 1600 px wide before inserting.

## 9. Avoid

- Bullet lists as the main content; title + bullets slides.
- Decorative accent bars under titles or down slide edges.
- Clip art, stock handshakes, 3D charts, shadows on everything.
- More than one dense slide in the main flow.
- Superlatives before evidence; unsourced competitor claims.
- Morph on unrelated slides; loud transitions (vortex, origami, curtains) in customer decks.
- Reading the slide aloud: if the notes repeat the slide words, the slide has too many.

## 10. What the engine generates vs what needs a person

Generated: every PowerPoint transition, Morph pairs (by object, word, character), camera zooms, 3-D turns, click builds and staggered entrances with their notes cues, off-canvas parking and stagger, overlays and transparency, shadows, gradients, greyscale, the pill-mosaic union (computed with shapely, written as a freeform with a picture fill and `rotWithShape="0"`), chevron columns and colour ramps, row builds, notes, appendix tables, lint.

Needs a person: background removal for product cut-outs, real ADI brand values, real measured data, clearance for reference wins, quotes and competitor comparisons, rehearsal.
