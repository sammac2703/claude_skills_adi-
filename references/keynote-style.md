# Keynote deck style (the dominant ADI house look)

**In one line:** dark, calm canvas, one accent colour, very few words, big visuals, and smooth motion, so every slide feels like a product keynote rather than a report, while staying on the ADI 2026 template and inside the ADI brand rules.

This style is the default for every route (A, B and C) unless the user asks for the light look. Where this guide and an ADI brand rule disagree, the conflict is settled in section 6; all other ADI rules in `SKILL.md` still apply.

## 1. Brand settings (fixed for ADI)

| Setting | ADI value | Guidance |
|---|---|---|
| Accent colour | ADI light blue `1B9CD0` (theme `accent2`) | One bright colour only. It is the highlighter, not the paint. Primary blue `0067B9` is too dark on the canvas to read as the accent; use it only for charts that need a second ADI blue |
| Dark tint | ADI dark blue `00325C` (theme `accent1`) | The far end of the background gradient |
| Font | Barlow Medium (titles), Barlow (body) | From the theme. Barlow Medium is the "bold" of this style; never name another face |
| Logo | The white ADI logo the template's dark layouts already carry, bottom right | Never add, move, resize or recolour a logo. Using the Dark Blue layouts is what puts the white logo on every slide |
| Card fill | ADI dark grey `3C4157` at 60% opacity | Reads as charcoal over the canvas |
| Muted text | ADI grey `9EA1AE` | Secondary lines only |
| Footer | Classification text plus slide number, from the layout | Public, Confidential or Internal, as in `SKILL.md` |

## 2. The look

1. **Canvas:** 16:9, black (`000000`) to ADI dark blue (`00325C`) diagonal gradient on every content slide, laid over the template's **Title and Content - Dark Blue** layout (so the white logo, footer and slide number stay in place). Cover, section and end-card slides keep the template's own layouts and artwork.
2. **Titles:** short, white, Title Case, stating the message. A single accent-coloured subtitle line sits beneath (`d.sub` in Route C). On template layouts the title stays in the template's title placeholder, unmoved (ADI rule); on builder-drawn slides (Route B) centre the title.
3. **Cards:** dark charcoal panels with a slim accent header tab. Three or four per slide at most (`d.card`).
4. **Icons:** line icons from the Brand Icon Library, recoloured white, inside accent-coloured circles, all the same style and size on a slide (`d.icon`). No other icon sources, no clip art.
5. **Accent use:** reserved for hero numbers, the key phrase, the subtitle line and small labels. Never for paragraphs. Do not introduce a second bright colour (purple, green, yellow and red stay for charts and status only).
6. **Lines:** faint white hairlines at about 12% opacity, never bright rules (Route C draws every `d.rule` this way in keynote style).
7. **Space:** generous margins (at least the template's 0.46 in left edge; 0.8 in on builder slides) and empty space. If a slide feels full, split it.
8. **Type sizes:** ADI floors still hold: nothing the audience reads below 14 pt; body 18 pt and up; card headings about 20 pt, card body 16 pt; hero numbers 96 to 160 pt.

## 3. The words

1. One message per slide, and the title states that message.
2. Cut every line to the bone: no more than four points, no more than six words each.
3. No full stops or commas on slides. Full sentences belong in the speaker notes. (Titles, numbers with decimals and the legal footer are exempt.)
4. Lead with verbs, for example "Start Earlier · Think Bigger · Finish Stronger".
5. Replace words with a picture, icon or number wherever possible (approved ADI imagery and icons only).
6. Close key slides with a single punchy takeaway line (`d.statement` in Route C, which renders in the accent).
7. Frame strategically and positively for leadership. Show what is next ("to be defined") rather than what is missing. This never overrides the ADI content guardrails: no over-promising, no invented figures, `[TO CONFIRM]` / `[TBC]` on anything unverified.
8. UK English, Title Case titles (ADI rules).

## 4. The motion

**Principle:** the deck should feel like one continuous canvas. Objects move, they do not jump. Motion guides the eye and is never decoration.

### 4.1 Three layers of motion

| Layer | What it does | Setting |
|---|---|---|
| **Slide transition** | Moves the audience from one slide to the next | Morph on every slide, "By object", about 1.0 to 1.25 s (Route C `d.finish` default 1.2 s; Route B theme `base: morph`, 1200 ms). Appendix slides: cut |
| **Builds** | Reveal points on a slide, one idea at a time | Fade or Float In (upwards), about 0.5 s, staggered (Route C default `float_up` 500 ms, 220 ms stagger) |
| **Dividers** | Mark a new chapter | Big two or three word statement or hero number, entered with Morph or a slow Fade (`d.hero`) |

This replaces the old "minimal motion" default. It is still not gratuitous animation in the ADI sense: no spins, bounces, loud transitions or sound, and every slide must read as a still for PDF.

### 4.2 How Morph is made to glide

Morph compares the slide you are leaving with the slide you are arriving at, and animates every object it can pair. The pairing is what creates the smoothness, so it is designed deliberately.

1. **Build slide pairs, not single slides.** Duplicate the previous slide, then move, resize or recolour objects into their new state. Never rebuild from scratch, or nothing will pair. In code: create the next slide with the same keys (`d.m_text`, `d.m_rule`, `d.card(..., key=)`, `d.icon(..., key=)`, `d.hero(key, ...)`).
2. **Name the objects that should travel.** Each travelling object gets a `!!` prefix, for example `!!heroNumber`, `!!hub`, `!!card1`. The same name on both slides forces PowerPoint to pair them, even if their text, size or colour has changed. (Hand edits: Selection Pane, Alt+F10.)
3. **Keep each `!!` name unique on a slide.** Two objects with the same name on one slide break the pairing.
4. **Let unpaired objects fade.** Anything that exists on only one slide fades in or out automatically. Use this for supporting text.
5. **Keep the background and logo fixed.** The canvas gradient and the template logo are identical on every slide, so they stay still while the content moves. Never animate the logo, footer or title chrome.

### 4.3 Typical Morph moves

- **Hero number to detail:** a large central figure (`d.hero(s, "heroNumber", ...)`) shrinks into a corner card as the breakdown appears on the next slide (same key, smaller size and new position).
- **Hub diagram build:** the centre circle (`!!hub`) stays put while partner nodes (`!!node1`...) slide out around it one slide at a time; nodes not yet shown are parked off-canvas (`d.park`).
- **Zoom in:** one card from a row of four (`key="card2"`) grows to fill the slide for its deep dive, then shrinks back on the following slide.

The patterns in `morph-patterns.md` (make room, focus row, park off-canvas, interleaved builds, slide links) all work in this style.

## 5. Route C code (clean_deck.py)

`CleanDeck(classification="public")` now defaults to `style="keynote"`:
- `d.slide(title)` uses the Dark Blue layout and paints the canvas.
- Colour roles are remapped so existing slide code reads correctly on the dark canvas: `TEXT_1` and `ACCENT_1` become white, `TEXT_2` and `ACCENT_3` become the accent. Write new code with `TC.BACKGROUND_1` (white) and `TC.ACCENT_2` (accent) directly.
- Every `d.rule` is a 12% white hairline.
- New helpers: `d.card(s, x, y, w, h, head, body, n=, key=)`, `d.icon(s, x, y, d, picture=, glyph=, key=)`, `d.hero(s, key, x, y, w, h, value)`. Each returns shapes ready to hand to `d.build`.
- `CleanDeck(style="light")` gives the original light, text-only look.

## 6. Where this style and the ADI brand rules meet

| Topic | Keynote guide says | On ADI, do this |
|---|---|---|
| Background | Dark gradient everywhere | Dark canvas is the default; the light look stays available for appendix tables, print and PDF-first leave-behinds, or when the user asks |
| Font | Any clean sans serif, bold titles | Barlow Medium / Barlow only, from the theme |
| Colours | One accent and a dark tint | ADI light blue and ADI dark blue; palette colours only |
| Logo | White logo bottom right | The template's own logo; never add or move one |
| Title position | Centred | Template placeholders stay where the template puts them. ADI typography is flush left, ragged right, so on builder-drawn slides prefer flush-left titles; centre only hero numbers and big statements |
| Background mix | Dark everywhere | The Standard template asks for mostly light with dark for key messages; the Large Format template asks for a mix by room light. Dark stays the default for conversation and keynote decks; dense, printed and training decks use the rhythms in `variation.md` |
| Spelling | n/a | ADI Editorial Standards say U.S. English; this skill defaults to UK at the user's request. See `compliance.md` |
| Cards and icons | Allowed, three or four per slide | Allowed (replaces the old "no coloured cards" Route C rule); icons from the Brand Icon Library only |
| Motion | Morph on every slide | Yes, with the ADI limits: no loud transitions, every slide reads as a still |
| Words | No punctuation, six words a line | Yes; ADI 14 pt floor, Title Case and classification footer still apply |
| Images | Strong visuals | Brand Image Library or approved photography only; never downloaded or invented |
