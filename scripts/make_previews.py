"""Render the brand files in assets/ and write 12-slide contact sheets to
assets/previews/sheet_<name>_<n>.png (4 x 3 grid, slide numbers labelled).

    python scripts/make_previews.py            # all brand files
    python scripts/make_previews.py lib icons  # only some

Needs PowerPoint (via scripts/render_slides.ps1, so the shell sandbox must be
off) and Pillow. Re-run whenever a brand file in assets/ is replaced.
"""
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SKILL = Path(__file__).resolve().parent.parent
ASSETS = SKILL / "assets"
OUT = ASSETS / "previews"
SOURCES = {
    "std": "ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx",
    "lib": "Brand-PPT-Slide-Library.pptx",
    "icons": "Brand-Icon-Library_2025-v2.pptx",
    "org": "ADI-Brand-PowerPoint-Template-Org-Chart-2026-v1.pptx",
    "imgs": "ADI-Brand-PowerPoint-Image-Library.pptx",
    "gold": "Gold_Deck_HearScore_V1.pptx",
}
COLS, ROWS, THUMB_W, GAP, LABEL_H = 4, 3, 480, 12, 28


def render(src: Path, out_dir: Path) -> list[Path]:
    subprocess.run(
        ["powershell", "-NoProfile", "-File",
         str(SKILL / "scripts" / "render_slides.ps1"),
         "-In", str(src), "-OutDir", str(out_dir), "-Width", str(THUMB_W)],
        check=True,
    )
    pngs = list(out_dir.glob("*.PNG")) + list(out_dir.glob("*.png"))
    num = lambda p: int(re.search(r"(\d+)", p.stem).group(1))
    return sorted(set(pngs), key=num)


def sheets(key: str, pngs: list[Path]) -> int:
    th = int(THUMB_W * 9 / 16)
    try:
        font = ImageFont.truetype("arial.ttf", 18)
    except OSError:
        font = ImageFont.load_default()
    per = COLS * ROWS
    count = 0
    for start in range(0, len(pngs), per):
        batch = pngs[start:start + per]
        sheet = Image.new("RGB", (COLS * (THUMB_W + GAP) + GAP,
                                  ROWS * (th + LABEL_H + GAP) + GAP), "white")
        draw = ImageDraw.Draw(sheet)
        for i, p in enumerate(batch):
            x = GAP + (i % COLS) * (THUMB_W + GAP)
            y = GAP + (i // COLS) * (th + LABEL_H + GAP)
            draw.text((x, y + 4), f"Slide {start + i + 1}", fill="black", font=font)
            with Image.open(p) as im:
                sheet.paste(im.convert("RGB").resize((THUMB_W, th)), (x, y + LABEL_H))
            draw.rectangle([x - 1, y + LABEL_H - 1, x + THUMB_W, y + LABEL_H + th], outline="#9EA1AE")
        count += 1
        sheet.save(OUT / f"sheet_{key}_{count}.png")
    return count


def main(keys: list[str]) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for key in keys or SOURCES:
        src = ASSETS / SOURCES[key]
        if not src.exists():
            print(f"skip {key}: {src.name} not found")
            continue
        for old in OUT.glob(f"sheet_{key}_*.png"):
            old.unlink()
        tmp = Path(tempfile.mkdtemp())
        try:
            n = sheets(key, render(src, tmp))
            print(f"{key}: {n} sheet(s)")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main(sys.argv[1:])
