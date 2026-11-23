# Template layouts and brand assets (moved from SKILL.md)

## Layouts (slide size 13.333 x 7.5 in) — placeholder idx in brackets
| Layout | Use | Placeholders |
|---|---|---|
| Cover with Pattern / with Image | Opening slide | title [10], name [14], date [15] |
| Title and Content | Default content slide | title, body [10] |
| Two Line Title and Content | Long titles | title, body [11] |
| Section Slide with Pattern / Image | Dividers, or "Key Message" | text [10] |
| Two Column Content | Comparisons | title, left [10], right [11] |
| Title and Content with Sub Header | Body with sub header | sub header [16], body [11] |
| Content with Sub Header Two Column | Two columns with headers | sub [16],[22], body [11],[23] |
| Content and Image | Text plus photo | title, picture [16], body [11] |
| Title Only – Blank | Charts, tables, custom visuals (Route B builds on this) | title |
| Two Column Alternate | Two panels with headers | [10],[19],[20],[21] |
| Title and Content – Primary Blue / Dark Blue (+ Two Line, + Sub Header) | Emphasis slides | title, body [11] |
| Org Chart 5x2 | People grids | 10 picture + 2 text placeholders each |
| Big Statement Two Column (+ Blue) | Statement plus supporting content | statement [27], content [23] / [11] |
| Big Statement – Primary Blue | One bold sentence | statement [28] |
| Video | Full-bleed video | media [12] |
| End Card | Final slide | none |

Reusable slide patterns (agenda, dashboard, roadmap, timeline, progress roadmap, next steps, before/after, pull-out quote, team success, etc.) live in `assets/Brand-PPT-Slide-Library.pptx`. Duplicate the slide with the pptx skill's `add_slide.py` and replace the text rather than rebuilding it.

## Other brand assets (all in `assets/`)
- **Icons** — `Brand-Icon-Library_2025-v2.pptx` (10 slides; slide 2 = usage rules). Sections: Capabilities (3), Industry Solutions (4), Car/Charging (5), Healthcare (6), Energy/Environment (7), Technology (8), Technology—Consumer (9), Workplace (10). Vector graphics: copy the shape XML (or duplicate the slide and delete the rest). Recolor with brand colors only; preferred: white, black, ADI blue, ADI light blue.
- **Images** — `ADI-Brand-PowerPoint-Image-Library.pptx` (62 slides, one licensed photo each; `image-index.md` describes them). Licensed for PowerPoint cover and section slides only. Slide 1 has instructions, slide 2 shows placement. Pre-cropped to 33.87 x 19.05 cm with the focal point on the right. Swap via the Slide Master (cover / section layout): delete the old picture, insert the new, send to back. Always show the user which image you picked.
- **Org chart** — `ADI-Brand-PowerPoint-Template-Org-Chart-2026-v1.pptx` (cover + 2 org chart slides). Same grid as layout "Org Chart 5x2". Internal Only unless told otherwise.
- **Gold deck (worked example)** — `Gold_Deck_HearScore_V1.pptx` (65 slides, 2025 intern project pitch for HearScore, a hearing-health wearable on MAX78000). A reference for structure and Morph technique only: problem-first arc (hearing crisis, causes, market gap) before the product reveal, then market sizing, approach options, hardware, algorithm, roadmap, demo and a backup appendix (slides 58 to 65). Study its slide pairs that build one idea across several slides with Morph: 15 to 17 (crisis stats), 18 to 20 (causes chart), 24 to 29 (feature reveal), 33 to 36 (market approach focus), 49 to 55 (timeline phases). Do not copy its content or numbers, and do not copy its deviations: 2025 footer (use the 2026 text above), off-palette colours (pink, bright green, purple, bright blue blocks; use the ADI palette and one accent), its own HearScore logo in the footer, off-template layouts ("Title Slide", "Section Header", "Title Only"), long body sentences, and draft notes left on slides 64 and 65. Internal material.
- **New slides (2026-11-23)** — `assets/new-slides/` holds the Corporate Overview and Messaging decks, the Alif and ADI briefing and workshop decks (`alif/`), and the approved logo kits (`ADI-Logo-Kit/`, `ADI-A2B_Logo-Kit/`); `references/new-slides/` holds the brand, logo, template, language, ADEF and media PDFs. Index and lessons: `references/new-slides-index.md`. Logos not already on a template layout come only from the logo kits. These folders are git-ignored.
- **Slide Library patterns** — Agenda, Performance, Dashboard, Roadmap with Narrative, Timeline, Progress Roadmap, Next Steps, Before and After with Statement, Product Features / Stories (see `assets/Brand-PPT-Slide-Library.pptx` and the PDF in `references/new-slides/`). Prefer these over inventing a layout for executive and status slides.
- **Previews** — `assets/previews/sheet_<name>_<n>.png` are 12-slide contact sheets of the template (`std`), slide library (`lib`), icons (`icons`), org chart (`org`), images (`imgs`) and gold deck (`gold`). View them to choose a layout, pattern, icon or photo without re-rendering. Generate or refresh them with `python scripts/make_previews.py [keys]` (needs PowerPoint rendering, see "Rendering"); if the folder is missing, say so and work from `image-index.md` and the layout table instead.
