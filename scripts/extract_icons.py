"""Build the icon catalogue: pull every SVG icon out of the Brand Icon Library with its label.

    python scripts/extract_icons.py

Writes assets/icons/<slide>_<label>.svg (black line art, recolour at use) and assets/icons/index.json (file -> label).
The library stores icons as SVG-only pictures; each label is the text box nearest below the icon.
"""
import json
import re
import sys
from pathlib import Path

from pptx import Presentation

SKILL = Path(__file__).resolve().parent.parent
SRC = SKILL / "assets" / "Brand-Icon-Library_2025-v2.pptx"
OUT = SKILL / "assets" / "icons"
SVG_BLIP = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"
EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(SRC))
    cat = {}
    for si in range(3, len(prs.slides) + 1):          # slide 1 cover, slide 2 usage rules
        s = prs.slides[si - 1]
        pics = [sh for sh in s.shapes if sh.shape_type == 13]
        txts = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and sh.shape_type != 14
                and sh.top > 914400 * 0.9]
        for k, p in enumerate(pics):
            cx, bottom, best = p.left + p.width / 2, p.top + p.height, None
            for t in txts:
                tcx = t.left + t.width / 2
                if abs(tcx - cx) < p.width * 0.7 and -0.15 * 914400 < t.top - bottom < 0.45 * 914400:
                    dist = abs(t.top - bottom) + abs(tcx - cx) * 0.2
                    if best is None or dist < best[0]:
                        best = (dist, t)
            label = " ".join(best[1].text_frame.text.split()) if best else f"s{si}_g{k}"
            el = p._element.find(".//" + SVG_BLIP)
            if el is None:
                continue
            name = f"{si}_" + re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")[:36]
            (OUT / f"{name}.svg").write_bytes(s.part.related_part(el.get(EMBED)).blob)
            cat[name] = label
    (OUT / "index.json").write_text(json.dumps(cat, indent=0), encoding="utf-8")
    print(f"{len(cat)} icons written to {OUT}")


if __name__ == "__main__":
    sys.exit(main())
