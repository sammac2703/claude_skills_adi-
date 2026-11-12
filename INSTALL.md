# Installing the adi-presentation skill

A Claude skill for building PowerPoint decks on the official Analog Devices 2026 brand template. Routes: **C** clean click-through conversation decks (default for meetings), **A** template-native decks, **B** deck builder for pitches and design-ins. Start with `SKILL.md`.

## 1. Get the skill
```powershell
git clone <this repo URL> $env:USERPROFILE\.claude\skills\adi-presentation
```
(or unzip the shared zip so that `...\skills\adi-presentation\SKILL.md` exists). For the Claude app (web or desktop), zip the folder and upload it in the skills settings; the admin may need to allow custom skills.

## 2. Install the Python packages
```powershell
pip install -r $env:USERPROFILE\.claude\skills\adi-presentation\requirements.txt
```
Needs python-pptx, lxml and Pillow. Do not run the scripts with `python -I`.

## 3. Add the ADI brand files (not stored in Git)
Copy these from the internal brand source into the skill's `assets\` folder:
- `ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx` (required)
- `ADI-Brand-PowerPoint-Image-Library.pptx`, `Brand-Icon-Library_2025-v2.pptx`, `Brand-PPT-Slide-Library.pptx`, `ADI-Brand-PowerPoint-Template-Org-Chart-2026-v1.pptx` (optional; used by Route A)

Install the **Barlow** and **Barlow Medium** fonts so decks display correctly.

## 4. Check it works
```powershell
python $env:USERPROFILE\.claude\skills\adi-presentation\scripts\example_clean_deck.py check.pptx
```
It writes a 7-slide example deck and prints click counts. Open it in PowerPoint and run Slide Show once.

## 5. Use it
Paste the prompt from `references/prompt-template.md` into Claude Code, filling in the brackets. Visual rendering (`scripts/render_slides.ps1`) needs PowerPoint on Windows; without it, open the file and check by eye.

## Rules for sharing
- Keep this repo **private** and internal to ADI. Never commit the brand files (`.gitignore` excludes them) or customer or university material.
- Generated decks stay local; do not commit `.pptx` or `.pdf` outputs.
