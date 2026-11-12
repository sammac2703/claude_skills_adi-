#!/usr/bin/env python3
"""
Build a customer deck from a brief JSON.

    python build_deck.py brief.json -o deck.pptx [--theme theme.json] [--static]

--static   leave out Morph build steps (extra row-by-row slides, empty-state
           2x2, zero-height bars) for a PDF or leave-behind version.

The brief format is documented in references/brief-format.md and
examples/brief_example.json. Lint runs automatically after saving and its
report is written next to the deck as <deck>.lint.md.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from deckkit import Deck, auto_morph, deep_merge, has_morph, load_theme, park_from, set_slide_tag, write_notes  # noqa: E402
from layouts import LAYOUTS  # noqa: E402
from motion import TRANSITIONS, set_transition  # noqa: E402

# --static: one still per idea (for PDF, print, email). No Morph state slides, no animation.
STATIC_OFF = {"build_rows": False, "reveal": False, "grow": False, "zoom": False, "zoom_out": False,
              "from": None, "build": None}

MOTION_LEVELS = {
    # minimal: quiet fades and Morph only, nothing animates inside a slide
    "minimal": {"level": "minimal", "base": "fade", "base_ms": 300, "hook": "fade", "hook_ms": 500,
                "section": "fade", "section_ms": 400},
    "standard": {"level": "standard"},
    # showcase: same grammar, more cinema (Ken Burns on photos, smooth reveal between sections)
    "showcase": {"level": "showcase", "hook_ms": 1400, "section": "reveal", "section_ms": 1400,
                 "base_ms": 450},
}


def _transition_spec(v):
    """'push' | {'kind': 'push', 'dir': 'l', 'ms': 600} -> (kind, ms, opts)"""
    if isinstance(v, str):
        return v, None, {}
    v = dict(v)
    kind = v.pop("kind")
    ms = v.pop("ms", None)
    return kind, ms, v


def apply_motion(deck, static=False):
    """Transitions on every slide without Morph, chosen by role; then write the
    animation timelines and add click cues to the notes."""
    m = deck.motion
    slides = list(deck.prs.slides)
    for i, sl in enumerate(slides):
        sid = sl.slide_id
        role = deck.roles.get(sid, "")
        if sid in deck.transition_for:
            kind, ms, opts = deck.transition_for[sid]
            if kind.startswith("morph") and not has_morph(sl):
                deck.log.append(f"slide {i+1}: transition '{kind}' asked for but nothing on it pairs with the slide before")
            set_transition(sl, kind, duration_ms=ms, **opts)
            continue
        if has_morph(sl):
            continue
        if static:
            set_transition(sl, "fade", duration_ms=300) if i else None
            continue
        if i == 0:
            set_transition(sl, m.get("hook", "fade_black"), duration_ms=m.get("hook_ms", 1000))
        elif role == "appendix" and deck.layouts.get(sid) != "section":
            set_transition(sl, m.get("appendix", "cut"))
        elif role in ("section",) or (role == "appendix" and deck.layouts.get(sid) == "section"):
            set_transition(sl, m.get("section", "fade_black"), duration_ms=m.get("section_ms", 800))
        else:
            set_transition(sl, m.get("base", "fade"), duration_ms=m.get("base_ms", 400))
    clicks_total = 0
    for sl in slides:
        tl = deck._tls.get(sl.slide_id)
        if not tl or not tl.items:
            continue
        n = tl.commit()
        clicks_total += n
        if n:
            labels = tl.click_labels()
            cue = "CLICKS (" + str(n) + "): " + "; ".join(f"{j+1}) {lab}" for j, lab in enumerate(labels))
            tf = sl.notes_slide.notes_text_frame
            tf.text = (tf.text + "\n\n" if tf.text.strip() else "") + cue
    return clicks_total


def build(brief: dict, out: Path, theme_ref=None, static=False, base_dir: Path | None = None):
    meta = dict(brief.get("meta", {}))
    meta["_base"] = str(base_dir or Path.cwd())
    ref = theme_ref if theme_ref is not None else brief.get("theme")
    if ref is None:
        # private/ is git-ignored: the real ADI template and theme live there on the work machine
        private = Path(__file__).resolve().parent.parent / "private" / "theme.json"
        if private.exists():
            ref = str(private)
    theme = load_theme(ref, base_dir)
    if ref is None:
        print("  note: no theme given and no private/theme.json; using the PLACEHOLDER palette")
    if brief.get("theme_overrides"):
        theme = deep_merge(theme, brief["theme_overrides"])
    mot = brief.get("motion")
    if mot:
        lvl = mot if isinstance(mot, str) else mot.get("level", theme["motion"].get("level", "standard"))
        if lvl not in MOTION_LEVELS:
            raise SystemExit(f"motion level '{lvl}' unknown; use minimal, standard or showcase")
        theme["motion"] = deep_merge(theme["motion"], MOTION_LEVELS[lvl])
        if isinstance(mot, dict):
            theme["motion"] = deep_merge(theme["motion"], {k: v for k, v in mot.items() if k != "level"})
    deck = Deck(theme, meta)
    deck.roles, deck.layouts = {}, {}
    if static:
        deck.animate = False

    def run(specs, appendix=False):
        prev_last = None
        for idx, spec in enumerate(specs):
            spec = dict(spec)
            if static:
                if spec.get("then"):
                    spec["text"] = spec["then"][-1]
                    spec["then"] = None
                if spec.get("from") and spec.get("layout") == "big_number":
                    spec["from"] = None
                spec.update({k: v for k, v in STATIC_OFF.items() if k != "from"})
            lay = spec.get("layout")
            if lay not in LAYOUTS:
                raise SystemExit(f"Unknown layout '{lay}' in slide {idx+1}. Valid: {', '.join(LAYOUTS)}")
            fn, default_role = LAYOUTS[lay]
            slides = fn(deck, spec)
            role = "appendix" if appendix else spec.get("role", default_role)
            for k, sl in enumerate(slides):
                tags = {"role": role, "layout": lay}
                if len(slides) > 1 and k < len(slides) - 1:
                    tags["build"] = f"{k+1}/{len(slides)}"
                if spec.get("answers"):
                    tags["answers"] = spec["answers"]
                if spec.get("id"):
                    tags["id"] = spec["id"]
                set_slide_tag(sl, **tags)
                deck.roles[sl.slide_id] = role
                deck.layouts[sl.slide_id] = lay
            if spec.get("transition") and not static:
                kind, ms, opts = _transition_spec(spec["transition"])
                if kind not in TRANSITIONS and kind != "none":
                    raise SystemExit(f"Unknown transition '{kind}'. Valid: {', '.join(TRANSITIONS)}")
                deck.transition_for[slides[0].slide_id] = (kind, ms, opts)
            # notes: full notes on the first slide of a build, short cues on the rest
            write_notes(slides[0], spec.get("notes"))
            cues = spec.get("build_notes") or []
            for k, sl in enumerate(slides[1:]):
                cue = cues[k] if k < len(cues) else "[CLICK] Build continues. Let it land."
                n = spec.get("notes")
                if isinstance(n, dict):
                    cue = {"say": cue, "conditions": n.get("conditions"), "source": n.get("source")}
                write_notes(sl, cue)
            # staggered entrance: park this slide's moving objects on the previous slide
            enter = spec.get("enter")
            if enter and prev_last is not None and not static:
                first = slides[0]
                n = park_from(deck, prev_last, first, direction=enter,
                              stagger_in=spec.get("stagger_in", 1.2), only=spec.get("enter_only"))
                if not n:
                    deck.log.append(f"slide '{lay}': enter='{enter}' found nothing new to park")
            prev_last = slides[-1]

    run(brief["slides"])
    if brief.get("appendix"):
        run([{"layout": "section", "title": "Appendix", "subtitle": brief.get("appendix_subtitle",
              "Detail, conditions and full comparisons"), "role": "appendix"}] + brief["appendix"],
            appendix=True)
    morphs = 0 if static else auto_morph(deck)
    deck.clicks = apply_motion(deck, static)
    out.parent.mkdir(parents=True, exist_ok=True)
    deck.prs.save(str(out))
    return deck, morphs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("brief")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--theme", default=None, help="theme JSON (overrides the brief's theme)")
    ap.add_argument("--static", action="store_true")
    ap.add_argument("--no-lint", action="store_true")
    a = ap.parse_args()
    bp = Path(a.brief).resolve()
    brief = json.loads(bp.read_text(encoding="utf-8"))
    out = Path(a.out or bp.with_suffix(".pptx")).resolve()
    theme_arg = str(Path(a.theme).resolve()) if a.theme else None
    deck, morphs = build(brief, out, theme_arg, a.static, base_dir=bp.parent)
    print(f"Saved {out}  ({len(deck.prs.slides)} slides, {morphs} Morph transitions, "
          f"{deck.clicks} click builds, motion: {'static' if a.static else deck.level})")
    for m in deck.log:
        print("  note:", m)
    if not a.no_lint:
        from lint_deck import lint, render
        theme = deck.t
        res = lint(str(out), theme=theme, brief=brief, static=a.static)
        report = render(res, str(out))
        Path(str(out) + ".lint.md").write_text(report, encoding="utf-8")
        print(report)


if __name__ == "__main__":
    main()
