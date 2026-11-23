# Slide Library layouts and icons (build blocks)

`assets/Brand-PPT-Slide-Library.pptx` carries the ADI building blocks as **layouts with placeholders**, so a slide is made by adding a slide on the layout and filling placeholders by idx. Open the library as the deck base (`CleanDeck(template=...Slide-Library.pptx)`; it drops the 16 example slides and keeps every layout), then use `scripts/adi_blocks.py`:

```python
from clean_deck import CleanDeck
from adi_blocks import Blocks, NAVY, BLUE, LIGHT, WHITE
d = CleanDeck(classification="internal", template="assets/Brand-PPT-Slide-Library.pptx")
b = Blocks(d)
s = b.slide("KPI Slide (x4)", "Title")             # title filled and named !!title (Morph glides it between layouts)
n = b.fill(s, 26, "8", size=60, bold=True)         # fill by idx; inherits layout font and colour
i = b.icon(s, "PRESENTING", x, y, 0.7, "circle", WHITE, LIGHT)   # SVG icon, recoloured; "bare" = line icon only
b.tidy(s); b.footer(s)                             # drop unused placeholders; classification footer from the layout
d.build(s, [[[n, *i]]], ["Click label"], "notes")  # one click per group; Morph via d.finish()
```
Icons: run `python scripts/extract_icons.py` once (writes `assets/icons/` and `index.json`, 193 icons with their library labels). Find one by label (`b.find_icon("TEAM MEETING")`). Icons are SVG with no PNG fallback (as in the library itself): PowerPoint 2016 or later. Recolour to white, ADI blue, ADI light blue, black or navy only. One style and size per slide.

## Layout map (idx of each placeholder; positions in inches)

| Layout (library name) | Look | Fill these idx | Good for |
|---|---|---|---|
| Cover Slide with Pattern | Pattern | 10 title, 14 subtitle, 15 date | Opening |
| Big Statement/Quote - Primary Blue | Blue | 27 statement, 28 attribution, 29 picture (circle at 5.92,0.85; replace with an icon) | Purpose, quote, one key message |
| Big Statement Two Column | Blue left, light right | 27 statement (left), 23 content (right) | Message plus three supports (messaging pillars) |
| Strategic Priorities | Dark | 0 title; rows 10,11,12 (labels) and 13,14,15 (text); numbers 1 to 3 are on the layout | Three wins, three priorities |
| ToC / Agenda Slide | Blue left, dark right | 27 title (left); 34 to 39 list items (y 1.24 to 6.02); draw your own number dots at x 4.62 | Agenda, formats, steps |
| KPI Slide (x4) | Dark | 0 title; per tile big=26/35/38/41, head=19/36/39/42, foot=29/37/40/43 | Four numbers |
| KPIs plus Side Content | Dark (accent1 background) | 0 title, 16 sub, 36 side head, 35 side text; tiles big=37/40/43/46, head=38/41/44/47, foot=39/42/45/48 | Dashboard with commentary |
| Event Roadmap/Timeline | Light, pins | 0 title; 6 columns text=17,33,34,35,36,37, label=38,23,42,41,40,39, image=16,28,29,30,31,32 (replace images with icons) | Phases, milestones |
| Progress Status/Timeline | Light grid | 0 title, 16 sub, rows 17 to 26 (label and text pairs) | RAG status over time |
| Project Roadmap 2 | Light | 0 title, 12 and 13 headings, table or bars drawn by you | Gantt-style plan |
| Employee Recognition / Team Success | Blue left, light right | 0 title, 10 left body; rows text=11,13,15, picture=18,16,17 (replace with icons) | People, buddies, stories |
| Product Features / Stories | Light | 0 title, 16 sub; top band heads 17,18,19 and bodies 20,21,22; bottom band heads 23,24 and bodies 25,26 | Five features in 3 plus 2 |
| Two Columns, Dark | Light with two panels and an arrow | 0 title; left panel head 22, bars 23,26,25,24; right panel head 11, bars 13,21,20,19; 18 statement | Before and after, start and end |
| Four Columns | Light | 0 title; column frames 27,35,36,37 (heading at top) and body 34,31,32,33; add number dots | Four items |
| Next Steps / Action Items (Light and Dark) | Both | 0, 16 sub, 29/39 and 20/30 left items, 18 body, 32/34/35 bottom band, 36 to 38 chart placeholders | Actions with a chart |
| Product or Strategic Dashboard | Light, 3 rows | rows with 5 fields each (ph 11 to 35) | Portfolio rows |
| Content with Subtitle - Dark Blue, 1_Title and Content - Dark Blue, 2_Title Only – Blank | Dark, light | 0 title, 16 sub (first) | Custom compositions (icons, rules, hero text) |
| End Card | Pattern | none | Closing |

Use layouts for the structural slides and custom compositions (icon rows, rising ladder, timeline line, hero text) for the rest, so no two neighbouring slides look alike. Do not draw cards over layouts that already have tiles. Unused placeholders must be removed (`b.tidy`). Layout containers are not moved or resized; add text boxes beside them instead.

## Gotchas learned building the internship deck
- Library layouts carry the footer placeholder but not always a slide-number placeholder; `b.footer` clones the footer only.
- Picture and image placeholders left empty show prompt text: remove them and place an icon.
- Text in a layout placeholder inherits its size and colour. Override only to raise to the 14 pt floor or for hero numbers.
- Morph pairs shapes by name, so the title is named `!!title` on every layout; Morph then glides the title between the different title positions.
- When slides alternate between very different library layouts, Morph makes the `!!title` fly across the slide and looks abrupt. The user asked for something quieter on the internship deck: a 0.7 s `fade` on every slide (`d.transition(s, "fade", 700)`), cut on the appendix, fade through black on cover and end. Offer this whenever neighbouring layouts differ a lot.
- The two "Two Columns, Dark" panels carry an arrow between them: use it only for start to end or before to after content.
- Content layouts carry no logo and no slide number. Call `b.brand(slide, dark)` (logo top right, 1.45 in wide above the title, plus a slide-number field bottom right). Dark backgrounds: Big Statement Primary Blue, Strategic Priorities, KPI Slide (x4), KPIs plus Side Content, ToC right panel, Dark Blue layouts; everything else is light (F0F0F2 or F2F2F2). Check a layout's background before choosing icon colours: BLUE icons vanish on the dark ones.
- `b.open_up(placeholder)` removes a placeholder's box so tiles and columns become open type with a hairline (`b.line`); use it when a layout stacks many boxes (Event Roadmap columns, KPI tiles).
- A layout's own shapes (arrows, number boxes) cannot be animated. To build an arrow after the first panel, cover it with a rectangle in the layout's background colour and add your own arrow to the click sequence.
- Text stacks above a rising bar must be placed above the bar top, never below it, or it crosses the baseline.

## Tables and charts (scripts/adi_data.py)
```python
from adi_data import paginate_table, add_chart
slides = paginate_table(b, "2_Title Only – Blank", "Appendix: Plan", header, rows, [0.8, 2.4, 3.3, 3.0, 2.9])   # one slide per page
add_chart(s, "column", ["Q1", "Q2"], {"2025": [3, 4], "2026": [5, 6]}, 0.46, 1.9, 8, 4.8, dark=False, number_format="0")
```
Row heights grow for wrapped cells and rows never split across slides. Tested on a 25-row table (4 pages, bottoms at 6.72 in) and five chart types; not visually rendered. Pie and doughnut take one series; a combo chart needs two charts.

## Icon colour rule (hard)
Glyphs on a blue or dark circle, panel or background are **always white**; never black or dark grey. `Blocks.icon` forces white for `style="circle"` and recolours every hex colour in the SVG (a few library icons, such as Mobile Robot, use 2B2B2B or 231F20 rather than 000000, which stayed dark until this was fixed). Bare icons: white on dark, ADI blue on light. `qa_layout.py` fails any icon that is not white on blue or dark, has contrast under 2, or still mixes colours.
