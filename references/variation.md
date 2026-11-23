# Variation: what stays fixed, what changes every deck

Goal: every deck feels like the same brand and the same hand, but no two decks (and no two neighbouring slides) look cloned. Fixed things build trust; varied things keep it alive.

## 1. Fixed (never vary)

| Fixed | Value |
|---|---|
| **Transition** | Morph on every slide, "by object", 1.0 to 1.25 s (appendix: cut, opener may fade through black) |
| **Builds** | Fade or Float In upwards, about 0.5 s, staggered; one point per click on conversation decks |
| **Morph craft** | `!!named` travelling objects, unique per slide, slide pairs made by duplicating, canvas and logo never move |
| **Brand** | Template layouts, Barlow, ADI palette, logo, classification footer, 14 pt floor, Title Case titles |
| **Words** | One message per slide, at most four points of six words, no full stops or commas on slides, verbs first |
| **Accent** | ADI light blue `1B9CD0`, one accent only |

Motion is the signature. Do not change transition type to create variety.

## 2. Varied (change on purpose)

### 2.1 Slide composition: the motif menu

Pick motifs from this menu for each slide. Each is built from Route C helpers or the matching Route B layout.

| Motif | Good for | Morph move that goes with it |
|---|---|---|
| **Hero number** (`d.hero`) | One proof figure, a result | Number shrinks into a card on the next slide |
| **Big statement** (`d.statement`, Big Statement layouts) | Chapter turn, key message | Words regroup into the next slide's title |
| **Card trio or quartet** (`d.card`) | Parallel points, options, pillars | Make-room: one card grows, others step aside |
| **Icon row** (`d.icon`) | Capabilities, steps, domains | Icons slide into a diagram |
| **Focus row / accordion** (`d.focus_row`) | Choosing between options | Selected row expands, others compress |
| **Hub and nodes** | System, ecosystem, signal chain | Nodes arrive one per slide, unborn ones parked off-canvas |
| **Before and after** | Change, problem to outcome | Left state morphs into right state |
| **Timeline or roadmap** | Plans, phases | Marker travels along the line |
| **Split with photo** | Applications, people, emotion | Photo crop widens to full bleed on the section slide |
| **Chart or table slide** | Evidence | Series highlight by recolouring, not by new objects |
| **Quote or pull-out** | Customer voice (with approval) | Quote marks anchor, text fades |
| **Quiz and answer pair** | Training | Same layout, answers colour in |
| **Block diagram build** | Technical decks | Blocks appear in signal-chain order |
| **Dashboard** (Slide Library) | Portfolio, status | Tile expands to its detail slide |

### 2.2 Look rhythm (canvas and layout mix)

The keynote look is the house default, but the ADI template guidance itself asks for a mix: the Standard template says use the light background predominantly and keep the blue and dark blue for key messages, call-outs and large visualisations; the Large Format template encourages a mix of dark and light by room lighting. Use the rhythm that fits the use case:

| Rhythm | Use for | How |
|---|---|---|
| **Dark throughout** (default for conversation and keynote decks) | Short decks, dim rooms, calls | Dark Blue layouts with the diagonal canvas (Route C default) |
| **Dark with a light evidence slide** | Pitches and briefings with a chart or table | Route A or B light layout for the evidence slide; keep the same title style |
| **Light-dominant, dark for key messages** | Training, deep dives, corporate overviews, anything dense or printed | Standard light layouts; Dark Blue or Primary Blue for chapter openers and takeaways |
| **Mixed by room light** | Large Format Venues | Dark in a darkened hall, light-grey where content is heavy |

Rules: never more than four slides in a row with the same look and composition; open and close strong (cover and end card on their template layouts); do not substitute any other colour for brand blue or dark blue; keep backgrounds identical between Morph partners so the canvas stays still. Route C currently builds dark only: for a mixed deck, build the dark slides in Route C and the light ones in Route A, or use Route B.

### 2.3 Cover, section and imagery

Alternate between "Cover with Pattern", "Cover with Image", "Section Slide with Pattern" and "Section Slide with Image" across decks. Pick the photograph from `assets/image-index.md` to suit the audience and say which one you picked. Do not reuse the same cover photograph two decks in a row.

### 2.4 Accent and number treatment

The accent stays `1B9CD0`, but vary where it lands: sometimes the hero figure, sometimes the key verb in the title, sometimes only the subtitle line. One accent per slide area, never everywhere.

### 2.5 Pace

Vary the beat: after two dense slides, give a one-word statement or a single hero figure. After a build-heavy slide, give a still slide. A deck of identical rhythm reads as a template.

## 3. How to make each deck different (the procedure)

1. **Pick a signature**: choose two or three motifs from 2.1 that suit this story and use them as the deck's recurring devices (for example hero number, hub build and quiz pair for a training deck).
2. **Plan the sequence**: list each slide's motif in the outline; no motif twice in a row; the signature motifs return at least twice so the deck feels designed.
3. **Choose the rhythm** from 2.2 and the cover and section artwork from 2.3.
4. **Check against the previous deck** if the user built one in this session or names one: change at least the signature motifs and the cover image.
5. **Record it**: put `Motifs: ...` and `Look: ...` in the delivery report and the first line of the speaker notes on slide 1, so the next request can choose differently.

## 4. Guardrails on variety

- Variety never beats clarity or the brand: no new colours, fonts, clip art or loud effects.
- If a motif needs more than four elements to work, simplify it.
- A hand-crafted Morph moment must carry meaning (a number becoming a card, a hub gaining nodes). Two or three per deck; the rest are plain Morph.
- Everything must still read as a still, because decks are exported to PDF and forwarded.
