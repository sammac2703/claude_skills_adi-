# Morph patterns (learned from a hand-crafted reference deck)

Studied from slides 13 to 65 of an internal "Gold" deck. What made it feel hand-crafted was **not** fancy click effects (53 of its effects were plain Appear). It was:
- **Morph chains** (21 of them, "by object", slow, about 2.0 s), each slide a duplicate of the previous with shapes moved, resized or restyled while keeping their names;
- **off-canvas parking** so shapes slide in and out instead of fading;
- **interleaved click builds**, and a few Slide Zoom / Section Zoom links.

The rule behind every pattern: **Morph pairs shapes that share a name**. In `clean_deck.py` use `d.m_text(slide, key, ...)` and `d.m_rule(slide, key, ...)`; the same `key` on consecutive slides becomes one shape that glides between its two states (position, size, font size, colour). Use 1.0 s for normal slides and 1.6 to 2.0 s for state pairs (`d.transition(slide, "morph", 1600)`).

Keep the house style: text, thin rules and photos only, no coloured blocks. Several reference techniques used coloured shapes; the text-only equivalents are given.

## Built into clean_deck.py

| Pattern | Reference slides | How | Helper |
|---|---|---|---|
| **Make room** | 18 to 20, 43 to 45 | Slide A shows a list at full size. Slide B uses the same keys smaller and tucked to one side (for example 24 pt to 18 pt, two columns to one) and new content enters in the freed space. Park the new content off-canvas on A so it slides in. | `m_text`, `m_rule`, `park`, `transition` |
| **Focus row (accordion / spotlight)** | 24 to 29, 33 to 36 | One slide per item. The active item widens and shows its description; the others narrow and shrink; the previous description collapses to 1 pt (a reference trick so Morph shrinks it instead of cutting). Labels must be short (about 8 characters at 18 pt). | `focus_row(title, [(label, desc), ...])` |
| **Park off-canvas** | 22 to 23, 40 to 42 | Hidden content sits beyond the slide edge (x above 13.4 in, y above 7.5 in, or negative). Morph slides it across. Use for entries and exits. | `park(slide, key, kind, edge)` |
| **Interleaved builds** | 31, 39 | One click can reveal a paragraph of a text box, then a picture, then another paragraph. Arrows land with the step they point to (put the arrow item first). | `build` accepts `(shape, paragraph_index)`; `arrow()` |
| **Slide link (zoom stand-in)** | 45, 57, 65 | A line on the main slide opens an appendix slide; a Back link returns. Mark the appendix hidden so the normal run skips it. | `link(shape, target)`; `slide._element.set("show", "0")` |

## Documented, not yet automated

| Pattern | Reference slides | Recipe |
|---|---|---|
| **Scatter then pick a winner** | 22 to 23 | State A: 8 to 12 items scattered with slight rotations (about 10 to 15 degrees) around the slide. State B: all but one parked off-canvas; the chosen item grows (for example 1.2 in to 3.9 in), straightens to 0 degrees and goes from 16 pt to 60 pt text. Use for narrowing to one idea. For a conversation deck do not pick the winner for the audience. |
| **Active-state highlight** | 49 to 55 | A row of phase labels. On each slide the active label grows (about 1.3 times) and its colour changes while the others shrink back. Text-only version: active label larger and `ACCENT_1`, others smaller and muted. |
| **Timeline zoom in / out** | 49 to 53 | A rule across the slide shortens to a third and moves left; detail content opens on the right (picture, annotated layout with dimension lines, a numbered checklist). The next slide expands the rule again. |
| **Pipeline zoom-out** | 43 to 45 | A process diagram scales to about 85% (14 pt to 11 pt) and shifts left on every state so the next stage fits; the title shrinks to a small corner label. |
| **Active-element swell with inner shadow** | 24 to 29 | Pictures or icons: the hero swells (1 to 2.7 in), the others shrink to 0.67 in with an inner shadow (looks recessed). Needs approved imagery. |
| **Motion path pan** | 64 | A picture pans left by 77% of the slide width over 2 s on one click, then drops 25% of the height on the next. `motion.Timeline` has `move` (dx, dy in inches). |
| **True Slide Zoom / Section Zoom** | 45, 57, 65 | PowerPoint's own Insert > Zoom objects. They need a thumbnail image of the target slide, so they cannot be generated safely without rendering. Create them by hand in PowerPoint (Insert > Zoom) if wanted; otherwise use slide links. |

## Do not copy (leftovers seen in the reference)
Working labels such as "Done" boxes, placeholder junk text, to-do notes on slides, and shape names left as "Rectangle 11". Name shapes with `!!key` and clear working notes before delivering.

## Checks
- Every `!!key` used on slide B has a partner on slide A (or is parked there); `d.finish` does not check this, so compare visually in Slide Show.
- A shape with an entrance animation is hidden until its click. Morph still pairs it, but for the smoothest hand-over keep state-pair shapes (the ones that move) unanimated and animate only the new content.
- Morph needs PowerPoint 365 or 2019 and later; older versions fall back to a fade.
