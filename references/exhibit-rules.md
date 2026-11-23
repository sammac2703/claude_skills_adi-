# Exhibit rules (charts, tables and consulting-style visuals)

Short rules for any data slide. Builders: `scripts/adi_data.py` (`add_chart`, `paginate_table`) and `scripts/exhibits.py` (`waterfall`, `matrix2x2`, `gantt`, `rating_table`, `funnel`, `takeaway`, `source_line`). In a spec: types `chart`, `waterfall`, `matrix`, `gantt`, `ratings`, `funnel`, `table`.

## Pick the exhibit by the message
| Message | Exhibit |
|---|---|
| Compare a few values | column or bar (sorted, one highlighted) |
| Change over time | line, or column when under about 8 periods |
| Part of a whole | stacked 100% column or doughnut with at most 4 parts; never a pie of many |
| How we got from A to B | waterfall |
| Where things sit on two criteria | 2x2 matrix |
| What happens when | Gantt (or the Event Roadmap layout for phases) |
| Options against criteria | rating table (filled dots) |
| Narrowing numbers | funnel |
| Exact figures | table, 14 pt, at most about 8 rows per slide |

## Rules
1. **Title states the finding** ("Volume Is the Swing"), not the topic ("Revenue Bridge"). One message per exhibit.
2. **One accent marks the point.** Everything is ADI blue or grey; the bar, point or row the title is about is ADI light blue. Never a rainbow; theme colours only, blues first.
3. **Label the data, drop the clutter:** values on the bars, no gridline clutter, no 3D, no legend when there is one series, axes only when they add something, 14 pt minimum everywhere.
4. **Add a takeaway line** (24 pt, accent) under the exhibit and a **source line** (14 pt grey) above the footer: "Source: planner v0.1, November 2026". Illustrative data says "Illustrative". Claims follow `compliance.md`.
5. **Anchor comparisons:** show the baseline, target or prior period in the same exhibit.
6. **Consistent encoding across the deck:** the same category keeps the same colour and label.
7. **Sort and order deliberately:** biggest first for ranking, chronological for time, logical for process.
8. **Numbers in the notes:** every figure on the slide also appears in the notes with its source (QA INFO check).
9. **Build order:** exhibits are static; reveal the takeaway on one click. Do not animate bars one by one.
10. **Do not use a chart where a number will do:** one big number (KPI layout) beats a one-bar chart.

## Typographic scale and grid (all slides)
Title 28 to 32 pt (layout). Key number 60 to 110 pt. Heading 20 to 24 pt bold. Body 18 pt (16 pt secondary). Labels and sources 14 pt. Left margin 0.46 in; content area from y 1.45 to 7.0; footer, number and logo are fixed. Columns share left edges to within 0.02 in (the QA gate warns on near-misses between 0.02 and 0.15 in); text boxes made with `Blocks.text` have no inset, placeholders may, so align by text edge, not box edge.
