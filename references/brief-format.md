# Brief format and layout fields

The brief is one JSON file. Asset paths are relative to it. See `examples/brief_example.json` for a complete one.

```json
{
  "meta": {
    "customer": "...", "application": "...", "audience": ["..."],
    "presenter": "...", "meeting_date": "2026-11-06",
    "competitors": ["Name used on slides"]
  },
  "theme": null,
  "theme_overrides": {"mode": "light"},
  "motion": "standard",
  "slides":   [ {"layout": "...", "...": "...", "notes": {...}} ],
  "appendix": [ {"layout": "table", "...": "..."} ]
}
```

- `theme`: path to a theme JSON, or null to use `private/theme.json` then the placeholder.
- `theme_overrides`: partial theme merged on top (e.g. switch to `"mode": "light"` for one deck).
- `motion`: `minimal`, `standard` (default) or `showcase`, or an object overriding fields of the motion system (`base`, `base_ms`, `hook`, `section`, ...). See `motion.md`.
- `meta.competitors`: every name here is flagged by lint wherever it appears, and needs `notes.source`.
- `appendix`: built after an automatic "Appendix" divider; lint relaxes word and numbering rules there.
- `meta.presenter` and `meta.audience` are not drawn on slides; they are there for you when writing the notes. The cover slide shows `subtitle` if you want a name on it. Do not invent a presenter name: use what the user gave, or leave it out.

## Fields every slide accepts

| Field | Meaning |
|---|---|
| `layout` | One of the layouts below (required) |
| `role` | Override the lint role (hook, credibility, objections, problem, criteria, positioning, tease, reveal, feature, proof, partners, agenda, ask, section) |
| `id` | Your own label, stored on the slide |
| `answers` | For feature slides: the `id` of the problem slide or the criterion word this feature answers, e.g. `"answers": "cost"` or `"answers": "Stable"`. Lint only checks it is set; it is for you and the presenter, and it is stored on the slide |
| `notes` | String, or object with `say`, `ask`, `pause_before`, `pause_after`, `conditions`, `source`, `fallback`, `transition` |
| `build_notes` | List of short cues for the extra build slides of a sequence (status_quo, quadrant, bars) |
| `enter` | `below`, `left`, `right` or `above`: this slide's Morph objects start parked off-canvas on the previous slide, staggered, and fly in |
| `enter_only` | List of name prefixes to park (default: all of the slide's `!!` objects not already on the previous slide) |
| `stagger_in` | Extra distance per parked object, inches (default 1.2). Further away travels faster |
| `transition` | Override this slide's entry transition: `"push"` or `{"kind": "push", "dir": "l", "ms": 600}`. Any key of `motion.TRANSITIONS`; normally leave it to the motion system |
| `morph_ms` | Morph duration for this slide's Morph steps (default 1500; camera moves 1600 to 1800) |

Any slide showing a number needs `notes.conditions`; lint fails it otherwise. Any comparative claim needs `notes.source`.

## Layouts

| Layout | Default role | Fields | Morph behaviour |
|---|---|---|---|
| `title` | hook | `title`, `kicker`, `subtitle`, `watermark` (PNG, shown at 75% transparency), `size` | `!!title` |
| `mosaic` | hook | `title`, `kicker`, `subtitle`, `photo` | Pill mosaic filled with the greyscale photo; gradient if no photo |
| `section` | section | `number`, `title`, `subtitle` | `!!section_title` |
| `statement` | problem | `text` (use `\n` for line breaks), `kicker`, `support`, `size`, `build: "lines"` (one line per click), `then`: [follow-on texts] | `then` adds one slide per text with Morph by Words: unchanged words glide, changed words swap |
| `credibility` | credibility | `number`, `unit`, `line`, `photo`, `kicker` | Photo greyscaled with colour overlay |
| `objections` | objections | `items`: [{`concern`, `answer`}] up to 3, `kicker` | Concerns shown; each answer wipes in on click |
| `status_quo` | problem | `rows`: [{`item`, `value`}], `total`: {`label`, `value`}, `ours`: {`label`, `value`, `note`}, `build_rows` (default true), `title`, `kicker` | One slide per row, then total, then ours; rows slide in from the left |
| `criteria` | criteria | `items`: ["word"] or [{`word`, `sub`}], 2 to 4, `kicker` | One yardstick per click |
| `quadrant` | positioning | `x_axis`, `y_axis`, `x_low`, `y_low`, `rivals`: [{`name`, `x`, `y`}], `ours`: {`name`, `x`, `y`}, `target_corner` (tr/tl/br/bl), `target_label`, `reveal` (default true) | Slide A: target empty; slide B: `!!ours` moves in. Coordinates 0..1, (1,1) top-right |
| `bars` | proof | `metric`, `unit`, `bars`: [{`label`, `value`, `display`, `ours`}], `grow`, `kicker` | With `grow`, bars start at zero and Morph grows them |
| `tease` | tease | `items`: [{`title`, `body`, `glyph` or `icon`}] up to 3, `kicker` | Circles `!!b1..b3`; pair with `enter: "below"` for the staggered entrance |
| `reveal` | reveal | `name`, `definition`, `image` (transparent PNG) or `glyph` (big accent text when there is no render, e.g. "1" or "AFE"), `kicker`, `merge_dots` (default true), `dots` | Tease circles shrink into dots under the name; `!!product` |
| `feature` | feature | `feature` (small caps line), `benefit` (the big line), `support`, `image` or `glyph`, `is_product`, `answers` | `!!product` if `is_product` |
| `product` | feature | `image`, `title`, `points` (up to 4), `kicker`, `tilt` (default true), `click_points` (default true) | `!!product` scales from the reveal and turns in 3-D; points float in one per click |
| `hardware` | feature/proof | `image`, `callouts`: [{`label`, `detail`, `x`, `y`}] (fractions of the image box), `kicker`, `zoom` (camera tour), `zoom_scale` (default 2.4), `zoom_out` | Callouts draw on click (dot, line, label). With `zoom`: one slide per callout, photo panned and zoomed around it, then back out |
| `roadmap` | agenda | `stages`: [{`letter`, `header`, `body`}] 3 to 6, `cover`: {`title`, `kicker`, `subtitle`} | With a cover, columns start as thin stripes off the left edge and Morph expands them |
| `big_number` | proof | `value`, `unit`, `label`, `kicker`, `size`, `from`, `from_label`, `from_kicker` | With `from`: old value first, then Morph by Characters into the new one, then a pulse |
| `zoom_tour` | feature | `chain`: [{`id`, `label`, `sub`, `ours`, `inside`: ["label" or {`label`, `sub`}]}] 3 to 6 blocks, `stops`: [{`focus` (block id, or `id.N` for the Nth inside item), `kicker`, `caption`}], `title`, `kicker`, `zoom_out` (default true), `area` | Camera move: overview, then each stop zooms the same shapes so that block fills the screen; caption band morphs between stops |
| `side_by_side` | proof | `left`/`right`: {`label`, `image` or `value`, `ours`}, `caption`, `kicker` | `!!sbs1`, `!!sbs2` |
| `partners` | partners | `items`: [{`name`, `logo`, `role`}], `quote`, `kicker` | None |
| `ask` | ask | `ask`, `owner`, `date` (a real date; if the user only gave a rough one, write it as given plus a marker, e.g. "early December [TBC: date]", which lint flags as REVIEW instead of FAIL), `kicker` | `!!ask` |
| `table` | appendix | `title`, `columns`, `rows`, `footnote` | Light background, native table |

## Writing good slide text

- Main flow: about six words on screen. Put the rest in `notes.say`.
- Headlines are statements or nouns the presenter talks over, not labels: "Three chips. 41 mW." beats "Current architecture overview".
- Numbers carry units on the slide ("14 days", "±0.5 %"); the conditions go in the notes.
- No full stops in kickers; kickers are short context ("Today", "Proof", "What has to be true"). Kickers are set in caps, but words with digits or units (0.9 mW, 5 nV/°C, kHz) keep their case.
- Write `[TBC: what is missing]` where a real value is not yet known. In big display slots (status_quo values, the ours box, `big_number`, bar labels) keep the marker short, `[TBC]`, and put what is missing in the notes; long markers shrink or wrap the slot.

## Rough length limits

The layouts are fixed boxes, and lint does not measure them. Stay near these and check the render:

| Field | Comfortable | Hard limit before it wraps or shrinks badly |
|---|---|---|
| `statement.text`, `mosaic.title` | 2 lines of about 25 characters | 3 lines |
| `feature.benefit` | 40 characters | 70 |
| `kicker`, `feature.feature` | 30 characters | 45 |
| `status_quo` row `item` / `value` | 30 / 8 characters | 45 / 14 |
| `status_quo.ours.value` | 6 characters ("$4.10") | 10 |
| `big_number.value` | 4 characters ("0.6", "±5") | 8 |
| `criteria` word | 1 word, about 10 characters | 14 |
| `tease` item `title` | 20 characters | 28 |
| `ask.ask` | 2 lines of about 35 characters | 3 lines |
| `roadmap` stage `header` / `body` | 14 / 60 characters | 20 / 90 |

For one-line captions in `zoom_tour` stops, keep under about 55 characters. Inside items are drawn at 5.5 pt in the overview on purpose (texture until you zoom); keep their labels to two or three words.
