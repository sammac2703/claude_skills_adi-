"""Code-checked layout QA gate for any .pptx (idea from Mck-ppt-design-skill: pass/fail comes from a script, never
from the model saying "looks fine").

    python scripts/qa_layout.py deck.pptx [-q] [--json report.json]   (no file is written unless --json)

Checks per slide (ERROR fails the gate, WARN and INFO are listed):
  ERROR  shape off the slide · text under 14 pt (footer, legal line and slide number exempt) · icon or text whose
         colour is too close to the background it sits on · leftover template prompt text · empty text placeholder
  WARN   text boxes overlapping each other · text likely to overflow its box · missing ADI logo or slide number on a
         content slide · slide with a transition missing
  INFO   large empty areas (less than a third of the content area used)
Hidden slides are checked too. It prints the result and writes nothing next to the deck. Exit code 0 = passed, 1 = failed. Not a substitute for looking at a render, but it
catches what this skill's decks have got wrong before (invisible icons, clipped text, empty slots, missing branding).
"""
from __future__ import annotations

import json
import math
import re
import sys
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

E = 914400
SCHEME = {"accent1": "00325C", "tx2": "0067B9", "bg1": "FFFFFF", "tx1": "000000", "accent2": "1B9CD0", "bg2": "F0F0F2",
          "dk2": "0067B9", "lt1": "FFFFFF", "dk1": "000000"}
EXEMPT = ("Slide Number", "Legal", "Footer")
PROMPT = re.compile(r"click to (edit|add)|lorem ipsum|\[?insert .{0,20} here|your text here|subtitle goes here", re.I)
NO_BRAND = ("Cover", "End Card", "Section", "Video")


def lum(h):
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4   # noqa: E731
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a, b):
    la, lb = sorted((lum(a), lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def solid_fill(el):
    """Hex of a solid fill in a spPr/bgPr element, or None. Scheme colours with lumMod/alpha are treated as unknown."""
    f = el.find("{*}solidFill")
    if f is None or len(f) == 0:
        return None
    c = f[0]
    if c.tag.endswith("srgbClr") and not len(c):
        return c.get("val").upper()
    if c.tag.endswith("schemeClr") and not len(c):
        return SCHEME.get(c.get("val"))
    if c.tag.endswith("schemeClr") and any(k.tag.endswith("lumMod") for k in c) and c.get("val") == "bg1":
        m = [k for k in c if k.tag.endswith("lumMod")][0]
        v = int(255 * int(m.get("val")) / 100000)
        return f"{v:02X}{v:02X}{v:02X}"
    return None


def slide_bg(slide):
    for src in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        bg = src._element.find(".//{*}bg")
        if bg is not None:
            pr = bg.find("{*}bgPr")
            if pr is not None:
                c = solid_fill(pr)
                if c:
                    return c
                if pr.find("{*}gradFill") is not None:
                    return "00325C"            # ADI keynote gradient: black to 00325C, treat as dark
                if pr.find("{*}blipFill") is not None:
                    return None                # picture background (cover, end card): contrast not checked
    return "FFFFFF"


def _overridden(slide, lsh):
    """True when a layout placeholder's box is dropped or restyled box-free on the slide."""
    if not lsh.is_placeholder:
        return False
    mine = [p for p in slide.placeholders if p.placeholder_format.idx == lsh.placeholder_format.idx]
    if not mine:
        return True
    sp = mine[0]._element.find("{*}spPr")
    return sp is not None and sp.find("{*}noFill") is not None


def colour_at(slide, x, y, below=None):
    """Background colour at a point: topmost solid-filled shape under it (slide shapes, then layout shapes), else bg."""
    for owner in (slide, slide.slide_layout):
        shapes = list(owner.shapes)
        if below is not None and owner is slide:
            shapes = shapes[:shapes.index(below)] if below in shapes else shapes
        for sh in reversed(shapes):
            if sh.is_placeholder and owner is slide and not sh.has_text_frame:
                continue
            if owner is slide.slide_layout and _overridden(slide, sh):
                continue
            try:
                if sh.width is None or not (sh.left <= x <= sh.left + sh.width and sh.top <= y <= sh.top + sh.height):
                    continue
                sp = sh._element.find("{*}spPr")
            except Exception:
                continue
            if sp is None or sh.width > 13.2 * E and sh.height > 7.4 * E:
                continue
            c = solid_fill(sp)
            if c and sh.shape_type != MSO_SHAPE_TYPE.PICTURE:
                return c
    return slide_bg(slide)


def svg_text(shape):
    try:
        blip = shape._element.find(".//{*}svgBlip")
        rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        return shape.part.related_part(rid).blob.decode("utf-8", "ignore")
    except Exception:
        return ""


def svg_colour(shape):
    try:
        blip = shape._element.find(".//{*}svgBlip")
        if blip is None:
            return None
        rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        svg = shape.part.related_part(rid).blob.decode("utf-8", "ignore")
    except Exception:
        return None
    cols = re.findall(r'(?:fill|stroke)[=:]\s*"?#([0-9A-Fa-f]{6})', svg) or \
        [c * 2 if len(c) == 3 else c for c in re.findall(r"(?:fill|stroke)[=:]\s*\"?#([0-9A-Fa-f]{3,6})", svg)]
    cols = [c.upper() for c in cols if len(c) == 6]
    return max(set(cols), key=cols.count) if cols else None


def run_colours(shape):
    out = []
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip():
                try:
                    if r.font.color is not None and r.font.color.type is not None and r.font.color.rgb is not None:
                        out.append(str(r.font.color.rgb))
                except Exception:
                    pass
    return out


def est_height(shape, width_in):
    need = 0.0
    for p in shape.text_frame.paragraphs:
        sz = max([r.font.size.pt for r in p.runs if r.font.size] or [18])
        per = max(1, int(width_in * 72 / (sz * 0.5)))
        need += math.ceil(max(1, len(p.text)) / per) * sz * 1.2 / 72
    return need


def text_inset(sh, slide):
    """Left text inset in EMU: slide bodyPr, else layout placeholder, else master placeholder of the same type, else
    the PowerPoint default 91440. Text boxes made by python-pptx carry their own bodyPr."""
    cands = [sh]
    if sh.is_placeholder:
        idx, typ = sh.placeholder_format.idx, sh.placeholder_format.type
        cands += [p for p in slide.slide_layout.placeholders if p.placeholder_format.idx == idx]
        cands += [p for p in slide.slide_layout.slide_master.placeholders if p.placeholder_format.type == typ
                  or (typ in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) and p.placeholder_format.type == PP_PLACEHOLDER.BODY)]
    for src in cands:
        bp = src._element.find(".//{*}bodyPr")
        if bp is not None and bp.get("lIns") is not None:
            return int(bp.get("lIns"))
    return 91440


def check(path):
    prs = Presentation(path)
    issues = []

    def add(level, n, msg):
        issues.append({"level": level, "slide": n, "msg": msg})

    W, H = prs.slide_width, prs.slide_height
    for n, s in enumerate(prs.slides, 1):
        lay = s.slide_layout.name
        has_logo = has_num = False
        texts = []
        edges = []
        for sh in s.shapes:
            name = sh.name or ""
            if name.startswith("ADI Logo"):
                has_logo = True
            if name == "Slide Number" or (sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER):
                has_num = True
            if sh.width is None:
                continue
            l_, t_, w_, h_ = sh.left, sh.top, sh.width, sh.height
            if getattr(sh, "rotation", 0) in (90, 270, -90):          # rotated 90 degrees: swap about the centre
                cx_, cy_ = l_ + w_ / 2, t_ + h_ / 2
                l_, t_, w_, h_ = cx_ - h_ / 2, cy_ - w_ / 2, h_, w_
            if l_ < -E * 0.02 or t_ < -E * 0.02 or l_ + w_ > W + E * 0.02 or t_ + h_ > H + E * 0.02:
                add("ERROR", n, f"'{name}' extends off the slide")
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or sh._element.tag.endswith("pic"):
                col = svg_colour(sh)
                if col and not name.startswith("ADI Logo"):
                    bg = colour_at(s, sh.left + sh.width // 2, sh.top + sh.height // 2, below=sh)
                    if bg and contrast(col, bg) < 2.0:
                        add("ERROR", n, f"icon '{name}' is {col} on {bg} (contrast {contrast(col, bg):.1f}): hard to see")
                    elif bg and lum(bg) < 0.4 and col not in ("FFFFFF",):
                        add("ERROR", n, f"icon '{name}' is {col} on blue or dark {bg}: icons on blue or dark must be white")
                    if "MsftOfcThm" in svg_text(sh):
                        add("ERROR", n, f"icon '{name}' still has Office theme classes: PowerPoint will draw it black")
                    allc = set(re.findall(r"#([0-9A-Fa-f]{6})(?![0-9A-Fa-f])", svg_text(sh)))
                    if len({c.upper() for c in allc}) > 1:
                        add("ERROR", n, f"icon '{name}' mixes colours {sorted({c.upper() for c in allc})}: not recoloured")
                continue
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text.strip()
            if sh.is_placeholder and not txt:
                if sh.placeholder_format.type not in (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE):
                    add("ERROR", n, f"empty placeholder '{name}' (shows prompt text in edit view)")
                continue
            if not txt:
                continue
            if PROMPT.search(txt):
                add("ERROR", n, f"leftover template text in '{name}': {txt[:40]!r}")
            exempt = any(e in name for e in EXEMPT) or (sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.FOOTER)
            if not exempt:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size and r.font.size.pt < 14 and r.text.strip():
                            add("ERROR", n, f"'{name}' has {r.font.size.pt:g} pt text (14 pt floor)")
                            break
                cols = run_colours(sh)
                if cols:
                    own = solid_fill(sh._element.find("{*}spPr")) if sh._element.find("{*}spPr") is not None else None
                    bg = own or colour_at(s, sh.left + min(sh.width // 2, int(0.4 * E)), sh.top + min(sh.height // 2, int(0.2 * E)), below=sh)
                    bad = [c for c in cols if bg and contrast(c, bg) < 2.5]
                    if bad:
                        add("ERROR", n, f"text '{txt[:24]}' is {bad[0]} on {bg} (contrast {contrast(bad[0], bg):.1f})")
                need = est_height(sh, max(0.2, sh.width / E - 0.2))
                if need > sh.height / E * 1.25 and sh.top / E + need > 7.0:
                    add("WARN", n, f"'{name}' may overflow: needs about {need:.2f} in, box is {sh.height / E:.2f} in")
                if sh.width > 0.6 * E and not name.startswith(("Rule", "Step bar", "Divider")):
                    edges.append((name, sh.left + text_inset(sh, s), sh.top))
                texts.append((name, sh.is_placeholder, sh.left, sh.top, sh.left + sh.width, sh.top + max(sh.height, int(min(need, 3) * E))))
        for i in range(len(texts)):
            for j in range(i + 1, len(texts)):
                a, b = texts[i], texts[j]
                if a[1] and b[1]:
                    continue                      # two placeholders: overlap is the layout's own design
                a, b = (a[0],) + a[2:], (b[0],) + b[2:]
                ox, oy = min(a[3], b[3]) - max(a[1], b[1]), min(a[4], b[4]) - max(a[2], b[2])
                if ox > 0.15 * E and oy > 0.12 * E:
                    area = ox * oy
                    small = min((a[3] - a[1]) * (a[4] - a[2]), (b[3] - b[1]) * (b[4] - b[2]))
                    if small and area / small > 0.3:
                        add("WARN", n, f"'{a[0]}' overlaps '{b[0]}'")
        near = [(a[0], b2[0], abs(a[1] - b2[1]) / E) for i, a in enumerate(edges) for b2 in edges[i + 1:]
                if 0.02 * E < abs(a[1] - b2[1]) < 0.15 * E and abs(a[2] - b2[2]) < 4 * E]
        for a, b2, dx in near[:3]:
            add("WARN", n, f"left edges of '{a}' and '{b2}' differ by {dx:.2f} in: align them")
        ttl = s.shapes.title.text_frame.text if s.shapes.title is not None else ""
        if len(ttl.split()) > 10:
            add("WARN", n, f"title has {len(ttl.split())} words (aim for 8 or fewer): {ttl[:50]!r}")
        if not any(k in lay for k in NO_BRAND):
            if not has_logo and not any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in s.slide_layout.shapes):
                add("WARN", n, "no ADI logo on this content slide")
            if not has_num:
                add("WARN", n, "no slide number")
            # empty-space check on the content area
            cells, hit = 24, set()
            for sh in list(s.shapes) + list(s.slide_layout.shapes):
                if sh.width is None or sh.top is None or (sh.is_placeholder and sh.shape_type is None):
                    continue
                if sh.top / E < 1.4 and sh.height / E < 1:
                    continue
                if sh.width > 13 * E and sh.height > 7 * E:
                    continue
                x0, x1 = int(sh.left / W * cells), int(min(sh.left + sh.width, W - 1) / W * cells)
                y0, y1 = int(sh.top / H * cells), int(min(sh.top + sh.height, H - 1) / H * cells)
                for xx in range(max(0, x0), min(cells, x1 + 1)):
                    for yy in range(max(0, y0), min(cells, y1 + 1)):
                        hit.add((xx, yy))
            area = [(x, y) for x in range(cells) for y in range(int(1.4 / 7.5 * cells), int(7.0 / 7.5 * cells))]
            used = sum(1 for c in area if c in hit) / max(1, len(area))
            if used < 0.33:
                add("INFO", n, f"only {used:.0%} of the content area is used: consider larger text or icons")
        notes = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
        body = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.name not in ("Slide Number",) and not sh.name.startswith("Legal"))
        missing = [x for x in dict.fromkeys(re.findall(r"\d[\d.,]*%?", body)) if len(x) > 1 and x not in notes and x not in ("10",)][:4]
        if notes and missing and not any(k in lay for k in NO_BRAND) and len(notes) > 40:
            add("INFO", n, f"numbers on the slide not mentioned in the notes: {', '.join(missing)}")
        tr = zipfile.ZipFile(path).read(f"ppt/slides/slide{n}.xml").decode("utf-8", "ignore")
        if "<p:transition" not in tr and "p159:morph" not in tr:
            add("WARN", n, "no slide transition")
    layouts = [s.slide_layout.name for s in prs.slides if s._element.get("show") != "0"]
    same_ok = {"1_Title and Content - Dark Blue", "2_Title Only – Blank"}
    for i in range(1, len(layouts)):
        if layouts[i] == layouts[i - 1] and layouts[i] not in same_ok:
            add("WARN", i + 1, f"same layout '{layouts[i]}' as the slide before: vary the motif")
        if i > 1 and layouts[i] == layouts[i - 1] == layouts[i - 2]:
            add("WARN", i + 1, f"three slides in a row on '{layouts[i]}': vary the look")
    return issues


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    issues = check(path)
    errs = [i for i in issues if i["level"] == "ERROR"]
    quiet = "-q" in sys.argv
    for lv in ("ERROR", "WARN", "INFO"):
        if quiet and lv == "INFO":
            continue
        for i in issues:
            if i["level"] == lv:
                print(f"{lv:5} slide {i['slide']:>2}: {i['msg']}")
    res = {"file": path, "passed": not errs, "errors": len(errs),
           "warnings": sum(1 for i in issues if i["level"] == "WARN"), "issues": issues}
    note = ""
    if "--json" in sys.argv:                       # report file only when asked: the deck folder gets the .pptx alone
        out = Path(sys.argv[sys.argv.index("--json") + 1])
        out.write_text(json.dumps(res, indent=1), encoding="utf-8")
        note = f" -> {out}"
    print(f"\n{'PASSED' if res['passed'] else 'FAILED'}: {res['errors']} errors, {res['warnings']} warnings{note}")
    return 0 if res["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
