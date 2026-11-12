# Plugging in the real ADI template

The repo ships with a PLACEHOLDER palette (dark navy, generic blue accent, Arial/Calibri). Every deck built with it says so in the build output. The real ADI look comes from the corporate template, which is company material and must not be committed to a personal or public repository.

## One-time setup on the work machine

1. Get the current ADI PowerPoint template (.potx or a blank .pptx made from it) from the usual internal source.
2. Copy it to `<skill_dir>/private/` (the folder is git-ignored).
3. Run:
   ```bash
   python <skill_dir>/scripts/extract_template.py private/ADI_template.potx -o private/theme.json
   ```
   This writes `private/theme.json` and `private/theme.report.md` with the slide size, theme colours, fonts, layout names and any fixed footer or legal text on the master.
4. Open `theme.json` and check by eye:
   - `colors`: the role mapping (dark_bg, accent, ours, ramp_start, ramp_end) is a guess from the theme scheme. Set them from the brand guide.
   - `fonts`: set `light` to the Light weight of the brand font if one exists; set `heavy` to the Black/Heavy weight if one exists.
   - `mode`: `template` keeps the master's backgrounds, logo and footers (most brand-safe). Set `template_background` to `light` or `dark` to match the master so text colours are right. `dark` paints the dark gradient over the template's masters while keeping their layouts.
   - `blank_layout`: the template's blank or title-only layout name, as listed in the report.
   - `footer.required_text`: the exact legal / confidentiality string, so lint fails any slide without it.
5. Build the example brief with it and look at every slide:
   ```bash
   python <skill_dir>/scripts/build_deck.py <skill_dir>/examples/brief_example.json -o test.pptx
   ```
6. Record answers to the open questions below in `private/notes.md` (also git-ignored).

## Using the skill in a work Claude account

The skill folder (with `private/` filled in) is what you upload or install there. Two routes, depending on how the account is set up:
- **Claude app (web/desktop):** zip the `adi-deck-builder` folder, including `private/`, and upload it as a skill in the app's skills settings. Organisation admins may need to allow custom skills and code execution. Do not upload the template anywhere ADI policy does not allow; if in doubt, ask IT/InfoSec first.
- **Claude Code on the work machine:** clone the repo into `~/.claude/skills/adi-deck-builder` (or a project's `.claude/skills/`), then add `private/` locally.

If the template cannot leave ADI systems at all, use the skill without it: build with the placeholder theme, then apply the ADI template in PowerPoint (Design > Browse themes / apply the .potx, or paste slides into a template deck with "Use destination theme") and re-check fonts.

## Open questions to answer at ADI

- [ ] Slide masters and layouts: is a .potx available, and which layout is the blank one?
- [ ] Real theme colours and fonts, including Light and Heavy weights
- [ ] Logo rules: position, minimum size, clear space, on dark vs light
- [ ] Required legal / confidentiality footers, and whether they differ for NDA vs public material
- [ ] Rules on naming competitors and comparative claims in customer decks
- [ ] Rules on citing customer wins and quotes (clearance process)
- [ ] Whether a one-page leave-behind is expected alongside decks
- [ ] Two or three past decks to use as test cases (lint them first: `lint_deck.py`)
- [ ] PowerPoint version on ADI laptops renders Morph (Microsoft 365 / PowerPoint 2019+ do; older versions show the fade fallback)
- [ ] Approved product imagery source (renders, eval board photos)
