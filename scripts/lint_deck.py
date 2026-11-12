#!/usr/bin/env python3
"""
Lint a customer deck against the story, slide, claims and build rules.

    python lint_deck.py deck.pptx [--brief brief.json] [--theme theme.json]
                                   [--competitors "Acme,Foo"] [--require-footer "Confidential"]

Works on decks made by build_deck.py (which tags each slide's role in
p:cSld/@name) and on hand-made decks. For hand-made decks, tag a slide by
putting "role: problem" (etc.) as the first line of its speaker notes.

Exit code 1 if any FAIL, else 0. WARNs are judgement calls for the presenter.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from motion import LOUD, count_clicks, effect_summary, get_transition, transition_ms  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))

MAIN_WORD_LIMIT = 8          # "about six words"; a little slack for units and names
LAYOUT_WORD_LIMIT = {"tease": 30, "criteria": 24, "ask": 30, "credibility": 18, "reveal": 20,
                     "side_by_side": 20, "partners": 30, "bars": 16}
DENSE_WORDS = 30
MAX_BEATS = 14
WORD_EXEMPT_ROLES = {"problem_table", "objections", "positioning", "agenda", "appendix",
                     "partners", "section"}
NUMBER_EXEMPT_ROLES = {"section", "agenda", "ask", "appendix"}
SUPERLATIVES = ["best", "fastest", "lowest", "highest", "smallest", "most accurate", "world's",
                "industry-leading", "industry leading", "unmatched", "unbeatable", "revolutionary",
                "first-ever", "leading-edge", "best-in-class", "best in class", "unrivalled", "unrivaled"]
JARGON_HINT = re.compile(r"\b(SNR|ENOB|THD|SFDR|INL|DNL|LSB|PSRR|CMRR|nV/√Hz|ppm/°C|SPI|I2C|JESD|LVDS|AFE|PGA|SAR|ΣΔ|sigma-delta)\b", re.I)
UNIT_AFTER = re.compile(
    r"^\s*(%|x\b|×|-?[a-zA-Zµμ°Ω$£€]{1,}|/)", re.U)
NUM = re.compile(r"(?<![\w.\-])([$£€]?\d[\d,]*(?:\.\d+)?)")
DATE_HINT = re.compile(r"\b(\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b|\d{4}-\d{2}-\d{2}|"
                       r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2}|Q[1-4]\s*\d{2,4}|"
                       r"week\s+\d+|by\s+(Mon|Tue|Wed|Thu|Fri)[a-z]*|end of (the )?(week|month))", re.I)


def slide_tags(slide):
    name = slide._element.find(qn("p:cSld")).get("name") or ""
    tags = dict(kv.split("=", 1) for kv in name.split(";") if "=" in kv)
    if "role" not in tags and slide.has_notes_slide:
        first = slide.notes_slide.notes_text_frame.text.strip().split("\n")[0]
        m = re.match(r"role\s*:\s*(\w+)", first, re.I)
        if m:
            tags["role"] = m.group(1).lower()
    return tags


def iter_text_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            yield from iter_text_shapes(sh.shapes)
        else:
            yield sh


def visible(sh, W, H):
    try:
        x, y, w, h = sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0
    except Exception:
        return True
    return x < W and y < H and x + w > 0 and y + h > 0


def slide_text(slide, W, H, include_aux=False):
    out = []
    for top in slide.shapes:
        top_vis = visible(top, W, H)
        if not top_vis:
            continue
        group = [top] if top.shape_type != 6 else list(iter_text_shapes(top.shapes))
        for sh in group:
            nm = sh.name or ""
            if nm.startswith("ftr_") or (not include_aux and nm.startswith("aux_")):
                continue
            if sh.has_text_frame and sh.text_frame.text.strip():
                out.append(sh.text_frame.text.strip())
            if getattr(sh, "has_table", False) and sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        if c.text.strip():
                            out.append(c.text.strip())
    return out


def names_top(slide):
    return {sh.name for sh in slide.shapes if sh.name and sh.name.startswith("!!")}


def has_morph(slide):
    return slide._element.find(".//{http://schemas.microsoft.com/office/powerpoint/2015/09/main}morph") is not None


def fonts_used(slide):
    faces = set()
    for el in slide._element.iter(qn("a:latin")):
        tf = el.get("typeface")
        if tf and not tf.startswith("+"):
            faces.add(tf)
    return faces


def lint(path, theme=None, brief=None, competitors=None, require_footer=None, static=False):
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    meta = (brief or {}).get("meta", {})
    competitors = [c.strip() for c in (competitors or meta.get("competitors") or []) if c and c.strip()]
    require_footer = require_footer or ((theme or {}).get("footer") or {}).get("required_text")
    allowed_fonts = None
    if theme:
        allowed_fonts = {f["face"] for f in theme["fonts"].values()}
    res = []  # (level, rule, slide_no, message)

    def add(level, rule, n, msg):
        res.append((level, rule, n, msg))

    slides = list(prs.slides)
    info = []
    for i, sl in enumerate(slides, start=1):
        tags = slide_tags(sl)
        role = tags.get("role", "unknown")
        txt = slide_text(sl, W, H)
        notes = sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else ""
        info.append(dict(n=i, slide=sl, tags=tags, role=role, text=txt, notes=notes,
                         alltext=slide_text(sl, W, H, include_aux=True)))

    main = [s for s in info if s["role"] != "appendix"]
    roles = [s["role"] for s in main]

    def first(role):
        return next((s["n"] for s in main if s["role"] == role), None)

    # ---------------- story
    if "unknown" in roles:
        add("WARN", "story", None, f"{roles.count('unknown')} slide(s) have no role tag; story-order checks are partial")
    p, c, r = first("problem"), first("criteria"), first("reveal")
    feats = [s["n"] for s in main if s["role"] == "feature"]
    if p is None:
        add("FAIL", "story.problem", None, "No problem / cost-of-status-quo slide")
    elif feats and min(feats) < p:
        add("FAIL", "story.problem", min(feats), "A feature slide comes before the problem slide")
    if c is None:
        add("FAIL", "story.criteria", None, "No criteria slide before the solution")
    elif r and c > r:
        add("FAIL", "story.criteria", c, "Criteria slide comes after the reveal")
    if r is None:
        add("WARN", "story.reveal", None, "No reveal slide (product + one-sentence definition)")
    else:
        rs = next(s for s in main if s["n"] == r)
        if not any(len(t.split()) >= 5 for t in rs["text"]):
            add("FAIL", "story.definition", r, "Reveal has no one-sentence definition")
    for s in main:
        if s["role"] == "feature" and not s["tags"].get("answers") and "ANSWERS:" not in s["notes"].upper() \
                and "PROBLEM:" not in s["notes"].upper():
            add("WARN", "story.feature_maps", s["n"], "Feature slide does not say which problem it answers (answers= in brief, or 'Problem:' in notes)")
    asks = [s for s in main if s["role"] == "ask"]
    if not asks:
        add("FAIL", "story.ask", None, "Deck does not end with a specific ask")
    else:
        a = asks[-1]
        if main and a["n"] != main[-1]["n"]:
            add("WARN", "story.ask", a["n"], "Ask is not the last main-flow slide")
        ask_txt = " ".join(a["text"]) + " " + a["notes"]
        if not DATE_HINT.search(ask_txt):
            if re.search(r"\[TBC[^\]]*\]", ask_txt) and re.search(r"(early|mid|late|end)[- ]\w+|week|month|\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b", ask_txt, re.I):
                add("REVIEW", "story.ask_date", a["n"], "Ask date is approximate and marked [TBC]; confirm a real date before the meeting")
            else:
                add("FAIL", "story.ask_date", a["n"], "Ask has no date (use a real date, or e.g. 'early December [TBC: date]' if only approximate)")
    proof_n = first("proof")

    # ---------------- length
    def _is_final_state(s):
        b = s["tags"].get("build", "")
        if "/" in b:
            k, n = b.split("/", 1)
            return k == n
        return True
    beats = [s for s in main if s["role"] != "section" and _is_final_state(s)]
    if len(beats) > MAX_BEATS:
        add("WARN", "story.length", None, f"{len(beats)} main-flow slides (build steps and sections not counted); "
            f"aim for 8 to 14 for a 30 to 45 minute meeting. Cut repeats or move them to the appendix")

    # ---------------- slides
    dense = 0
    for s in main:
        words = sum(len(t.split()) for t in s["text"])
        if s["role"] not in WORD_EXEMPT_ROLES and s["tags"].get("layout") not in ("status_quo", "hardware", "objections", "roadmap", "quadrant", "zoom_tour"):
            if words > DENSE_WORDS:
                dense += 1
                add("WARN", "slide.dense", s["n"], f"{words} words on a main-flow slide; move detail to notes or the appendix")
            elif words > LAYOUT_WORD_LIMIT.get(s["tags"].get("layout"), MAIN_WORD_LIMIT + 6):
                add("WARN", "slide.words", s["n"], f"{words} words on slide (target about six); is there one idea here?")
    if dense > 1:
        add("FAIL", "slide.dense_count", None, f"{dense} dense slides in the main flow (max one, and only as proof)")
    for s in main:
        joined = " ".join(s["text"])
        if proof_n is None or s["n"] < (proof_n or 10**6):
            if s["role"] not in ("proof", "appendix") and JARGON_HINT.search(joined) and s["role"] in ("hook", "credibility", "criteria", "tease"):
                add("WARN", "slide.jargon", s["n"], f"Jargon before its benefit is shown: {JARGON_HINT.search(joined).group(0)}")

    # ---------------- numbers and claims
    for s in info:
        if s["role"] in NUMBER_EXEMPT_ROLES:
            continue
        has_num = False
        for t in s["text"]:
            for m in NUM.finditer(t):
                tok = m.group(1)
                if re.fullmatch(r"0?\d", tok) and len(t.strip()) <= 2:
                    continue  # list numbering
                has_num = True
                after = t[m.end():]
                if not tok[0] in "$£€" and not UNIT_AFTER.match(after):
                    add("WARN", "claims.units", s["n"], f"Number '{tok}' in \"{t[:40]}\" has no unit next to it")
        if has_num and s["role"] != "appendix" and "CONDITIONS:" not in s["notes"].upper():
            add("FAIL", "claims.conditions", s["n"], "Slide shows numbers but the notes carry no CONDITIONS: line")
    for s in info:
        joined = " ".join(s["alltext"]).lower()
        for comp in competitors:
            if comp.lower() in joined:
                add("REVIEW", "claims.competitor", s["n"], f"Competitor named on slide: {comp}. Check ADI's comparison rules and cite a source")
                if "SOURCE:" not in s["notes"].upper():
                    add("FAIL", "claims.source", s["n"], f"Comparative slide naming {comp} has no SOURCE: in notes")
                elif re.search(r"SOURCE:\s*\[?TBC", s["notes"], re.I):
                    add("REVIEW", "claims.source", s["n"], f"Source for the {comp} comparison is still [TBC]; do not present until filled")
        if s["role"] not in ("appendix", "criteria", "problem") and (proof_n is None or s["n"] < proof_n):
            for w in SUPERLATIVES:
                if re.search(r"\b" + re.escape(w) + r"\b", joined):
                    add("WARN", "claims.superlative", s["n"], f"Superlative '{w}' before any proof slide")

    # ---------------- build quality
    for prev, cur in zip(info, info[1:]):
        shared = names_top(prev["slide"]) & names_top(cur["slide"])
        if shared and not has_morph(cur["slide"]) and not static:
            add("FAIL", "build.morph", cur["n"], f"Shares Morph names with the previous slide ({', '.join(sorted(shared)[:3])}) but has no Morph transition")
    for s in info:
        sl = s["slide"]
        if has_morph(sl) and s["n"] > 1 and not (names_top(sl) & names_top(info[s['n'] - 2]["slide"])):
            add("WARN", "build.morph_unpaired", s["n"], "Morph transition but no !!-named objects shared with the previous slide (Morph will guess)")
        # lockstep: off-canvas parked objects at identical offsets
        offs = []
        for sh in sl.shapes:
            if sh.name and sh.name.startswith("!!") and not visible(sh, W, H):
                offs.append((sh.left, sh.top))
        dup = len(offs) - len(set(offs))
        if dup:
            add("WARN", "build.stagger", s["n"], f"{dup} parked object(s) share an identical off-canvas position (lockstep motion)")
        if allowed_fonts:
            extra = fonts_used(sl) - allowed_fonts
            if extra:
                add("FAIL", "build.fonts", s["n"], f"Fonts outside the theme: {', '.join(sorted(extra))}")
        for top in sl.shapes:
            for sh in ([top] if top.shape_type != 6 else list(iter_text_shapes(top.shapes))):
                if sh.shape_type == 13 and (sh.name or "").startswith("photo_"):
                    if sh._element.find(".//" + qn("a:grayscl")) is None:
                        add("WARN", "build.photo_bw", s["n"], f"Photo '{sh.name}' is in colour; greyscale or overlay it on brand colour")
        if require_footer:
            found = any(require_footer.lower() in (t or "").lower()
                        for t in slide_text(sl, W, H, include_aux=True) + [x.text_frame.text for x in sl.shapes if x.has_text_frame])
            if not found:
                for part in (sl.slide_layout, sl.slide_layout.slide_master):
                    if any(sh.has_text_frame and require_footer.lower() in sh.text_frame.text.lower() for sh in part.shapes):
                        found = True
            if not found:
                add("FAIL", "build.footer", s["n"], f"Required footer text '{require_footer}' not found")
        if not s["notes"].strip() and s["role"] != "section" and s["tags"].get("layout") != "section":
            add("WARN", "notes.missing", s["n"], "No speaker notes")

    # ---------------- motion (references/motion.md)
    base_kinds = {}
    emph = 0
    for s in info:
        sl = s["slide"]
        kind = get_transition(sl)
        ms = transition_ms(sl)
        if kind in LOUD and s["role"] != "appendix":
            add("WARN", "motion.gimmick", s["n"], f"'{kind}' transition: the audience will remember the effect, not the point. Use fade, push or Morph")
        if kind and kind.startswith("morph"):
            if ms and ms > 2500:
                add("WARN", "motion.slow", s["n"], f"Morph of {ms} ms; 1000 to 2000 ms keeps a customer meeting moving")
        elif kind and ms and ms > 1000 and s["n"] > 1 and s["role"] != "section" and s["tags"].get("layout") != "section":
            add("WARN", "motion.slow", s["n"], f"{kind} transition of {ms} ms; keep ordinary transitions at 300 to 600 ms")
        if kind and not kind.startswith("morph") and s["n"] > 1 and s["role"] not in ("section", "appendix"):
            base_kinds.setdefault(kind, []).append(s["n"])
        clicks = count_clicks(sl)
        if static and clicks:
            add("WARN", "motion.static", s["n"], "Static copy still has click animations")
        if clicks > 6:
            add("WARN", "motion.clicks", s["n"], f"{clicks} clicks on one slide; split it, or the audience waits on you")
        if clicks and "CLICK" not in s["notes"].upper():
            add("WARN", "notes.clicks", s["n"], f"{clicks} click build(s) but no CLICK cues in the notes")
        for cls, pid, dur in effect_summary(sl):
            if cls == "emph" and pid in (26, 6) and dur < 5000:
                emph += 1
            if cls == "entr" and dur > 1000:
                add("WARN", "motion.slow", s["n"], f"Entrance effect of {dur} ms; 300 to 600 ms reads as confident")
            if cls == "path":
                add("REVIEW", "motion.path", s["n"], "Motion path: check it in Slide Show on the presenting laptop")
        if s["role"] == "appendix" and clicks:
            add("WARN", "motion.appendix", s["n"], "Appendix slides should be static reference pages")
    if len(base_kinds) > 2:
        detail = "; ".join(f"{k}: slides {', '.join(map(str, v[:4]))}" for k, v in base_kinds.items())
        add("WARN", "motion.consistency", None, f"{len(base_kinds)} different ordinary transitions ({detail}). Pick one base transition and use Morph for the moments that move")
    if emph > 4:
        add("WARN", "motion.emphasis", None, f"{emph} emphasis effects; keep pulses to the two or three numbers that matter")
    return res


def render(res, path):
    order = {"FAIL": 0, "REVIEW": 1, "WARN": 2}
    res = sorted(res, key=lambda r: (order[r[0]], r[2] or 0))
    fails = sum(1 for r in res if r[0] == "FAIL")
    warns = sum(1 for r in res if r[0] == "WARN")
    revs = sum(1 for r in res if r[0] == "REVIEW")
    lines = [f"# Lint: {Path(path).name}", "",
             f"**{fails} fail, {revs} review, {warns} warn**", ""]
    if not res:
        lines.append("Clean.")
    for lvl, rule, n, msg in res:
        where = f"slide {n}" if n else "deck"
        lines.append(f"- {lvl} `{rule}` ({where}): {msg}")
    return "\n".join(lines) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--brief")
    ap.add_argument("--theme")
    ap.add_argument("--competitors", default="")
    ap.add_argument("--require-footer", default=None)
    ap.add_argument("--static", action="store_true", help="deck is a deliberate no-Morph copy; skip Morph pairing checks")
    a = ap.parse_args()
    brief = json.loads(Path(a.brief).read_text(encoding="utf-8")) if a.brief else None
    theme = None
    if a.theme:
        from deckkit import load_theme
        theme = load_theme(a.theme)
    comps = [c for c in a.competitors.split(",") if c.strip()] or None
    res = lint(a.deck, theme, brief, comps, a.require_footer, static=a.static)
    print(render(res, a.deck))
    sys.exit(1 if any(r[0] == "FAIL" for r in res) else 0)


if __name__ == "__main__":
    main()
