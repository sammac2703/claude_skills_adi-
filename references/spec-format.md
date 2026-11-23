# Spec format for `scripts/build_from_spec.py`

```json
{"classification": "internal|public|confidential", "look": "dark|light", "transition": "fade|morph",
 "slides": [ {"type": "...", ...} ]}
```
Common fields on every slide: `notes` (speaker script, sources, `[TO CONFIRM]`), `pillar` (messaging pillar line added to notes), `look` (dark or light for custom slides), `appendix: true` (hidden, cut transition). Every item gets one click; `CLICKS (n)` is written to the notes. Titles are claims (`storyline-checks.md`). No full stops on slides. Icons are matched by label from the Brand Icon Library (`python scripts/find_icon.py team`); an unmatched icon is skipped.

| type | Fields | Layout |
|---|---|---|
| `cover` | title, subtitle, date, detail ("a\|b" lines) | Cover Slide with Pattern |
| `statement` | text, attribution?, icon? | Big Statement/Quote - Primary Blue |
| `exec_summary` | title, answer, points[{head,text}] up to 3 | custom: answer left, three supports right |
| `priorities` | title, rows[{label,lines[],icon}] x3, statement? | Strategic Priorities (dark) |
| `kpi` | title, items[{value,label,caption}] x4 | KPI Slide (x4) (dark) |
| `agenda` | title, items[{head,text}] up to 6, hero?, hero_label? | ToC / Agenda |
| `timeline` | title, phases[{head,text,when,icon}] up to 6 | Event Roadmap/Timeline |
| `icons` | title, sub?, items[{head,text,icon}] 3 to 4, statement? | custom columns |
| `rows` | title, sub?, items[{text,icon}] up to 6, statement? | custom numbered rows |
| `ask` | title, steps[], ask | custom rows plus the ask line |
| `table` | title, header[], rows[[]], widths[], sub? | paginated native table |
| `chart` | title, kind (column, stacked, bar, line, pie, doughnut, area), categories[], series{}, number_format?, takeaway?, source? | native chart |
| `waterfall` | title, labels[], values[] (first total, deltas, last null = computed), highlight?, takeaway?, source? | bridge |
| `matrix` | title, x_label, y_label, quads[4], points[[label,x0to1,y0to1]], highlight? | 2x2 |
| `gantt` | title, periods, tasks[[name,start,end]], highlight? | Gantt |
| `ratings` | title, header[], rows[[label,n,...]], widths[], highlight? | dot ratings |
| `funnel` | title, stages[[label,value]], highlight? | funnel |
| `end` | none | End Card |

Example: `tests/regression_spec.json` (every type once). Workflow: write the spec, `python scripts/build_from_spec.py spec.json deck.pptx`, `python scripts/qa_layout.py deck.pptx`, `python scripts/preview_png.py deck.pptx out/` and look at the PNGs, fix, deliver when the gate passes. Anything the spec cannot express: build the slide with `adi_blocks.py` and add it to the same deck (see `library-layouts.md`).
