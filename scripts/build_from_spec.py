"""Build an ADI deck from a short JSON spec: the model writes content, this script does layout, branding and motion.

    python scripts/build_from_spec.py spec.json out.pptx          # then: python scripts/qa_layout.py out.pptx
Spec fields and slide types: references/spec-format.md. Uses the Slide Library layouts, Brand Icon Library icons, the
logo and slide number on every content slide, 0.5 s float-up builds (one per item), 0.7 s fade (or Morph when the spec
says "transition": "morph"), cut on hidden appendix slides, notes with a CLICKS cue.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from adi_blocks import BLUE, CHAR, LIGHT, NAVY, WHITE, Blocks      # noqa: E402
from adi_data import add_chart, paginate_table                      # noqa: E402
from clean_deck import CleanDeck, MSO_ANCHOR, PP_ALIGN, Pt, Inches  # noqa: E402
from exhibits import (funnel, gantt, matrix2x2, rating_table, source_line, takeaway, waterfall)  # noqa: E402
from motion import get_transition                                    # noqa: E402
from pptx.dml.color import RGBColor                                  # noqa: E402

SK = Path(__file__).resolve().parent.parent
DARK_LAYOUTS = {"Big Statement/Quote - Primary Blue", "Strategic Priorities", "KPI Slide (x4)", "KPIs plus Side Content",
                "1_Title and Content - Dark Blue", "ToC / Agenda Slide"}


class Spec:
    def __init__(self, spec):
        self.sp = spec
        self.d = CleanDeck(classification=spec.get("classification", "internal"),
                           template=SK / "assets" / "Brand-PPT-Slide-Library.pptx")
        self.b = Blocks(self.d)
        self.links = []
        self.appendix = []
        self.look = spec.get("look", "dark")
        self.names = {}

    # ---- helpers -------------------------------------------------------------------------------
    def icon_for(self, item, key="head"):
        lab = item.get("icon") or item.get(key, "")
        try:
            self.b.find_icon(lab)
            return lab
        except KeyError:
            return None

    def custom(self, title, sub, dark):
        b, d = self.b, self.d
        s = b.slide("1_Title and Content - Dark Blue" if dark else "2_Title Only – Blank", title)
        if dark:
            p11 = b.ph(s, 11)
            if p11 is not None:
                p11._element.getparent().remove(p11._element)
            d.canvas(s)
        if sub:
            b.text(s, 0.46, 1.45, 12.4, 0.4, sub, 18, LIGHT if dark else BLUE, name="!!sub")
        return s

    def finish(self, s, groups, labels, notes, pillar=None):
        b = self.b
        b.tidy(s)
        f = b.footer(s)
        dark = s.slide_layout.name in DARK_LAYOUTS
        if f is not None and not dark:
            for r in f.text_frame.paragraphs[0].runs:
                r.font.color.rgb = RGBColor.from_string(CHAR)
        b.brand(s, dark)
        head = f"MESSAGE PILLAR: {pillar}.\n\n" if pillar else ""
        groups = [[g] for g in groups if g]          # one click per group; the group is one build item
        if groups:
            self.d.build(s, groups, labels[:len(groups)], head + (notes or ""))
        else:
            s.notes_slide.notes_text_frame.text = head + (notes or "")

    def item(self, *parts):
        out = []
        for p in parts:
            if isinstance(p, (list, tuple)) and not (isinstance(p, tuple) and len(p) == 2 and isinstance(p[1], int)):
                out += self.item(*p)
            elif p is not None:
                out.append(p)
        return out

    def is_dark(self, sl):
        return {"dark": True, "light": False}.get(sl.get("look", self.look), True)

    # ---- slide types ---------------------------------------------------------------------------
    def cover(self, sl):
        b = self.b
        s = b.slide("Cover Slide with Pattern", None)
        b.fill(s, 10, sl["title"])
        b.fill(s, 14, sl.get("subtitle", ""))
        b.fill(s, 15, sl.get("date", ""))
        if sl.get("detail"):
            b.text(s, 0.47, 6.1, 8.0, 0.7, sl["detail"].split("|"), 14, WHITE, name="Meeting details")
        b.tidy(s)
        s.notes_slide.notes_text_frame.text = sl.get("notes", "")
        return s

    def statement(self, sl):
        b = self.b
        s = b.slide("Big Statement/Quote - Primary Blue", None)
        st = b.fill(s, 27, sl["text"], name="!!title")
        at = b.fill(s, 28, sl.get("attribution", "")) if sl.get("attribution") else None
        p29 = b.ph(s, 29)
        p29._element.getparent().remove(p29._element)
        ico = self.icon_for(sl, "text") if sl.get("icon") else None
        ic = b.icon(s, ico, 5.92, 0.85, 1.49, "bare", WHITE, key="ico") if ico else []
        self.finish(s, [self.item(ic), self.item(st), self.item(at)], ["Icon", "Statement", "Attribution"],
                    sl.get("notes"), sl.get("pillar"))
        return s

    def kpi(self, sl):
        b = self.b
        s = b.slide("KPI Slide (x4)", sl["title"])
        pos = [(26, 19, 29, (2.63, 1.72)), (35, 36, 37, (6.76, 1.72)), (38, 39, 40, (2.63, 4.09)), (41, 42, 43, (6.76, 4.09))]
        groups = []
        for (big, head, foot, (tx, ty)), it in zip(pos, sl["items"]):
            n = b.open_up(b.fill(s, big, it["value"], size=60, bold=True, anchor=MSO_ANCHOR.MIDDLE))
            h = b.open_up(b.fill(s, head, it["label"]))
            f = b.fill(s, foot, it.get("caption", ""))
            rule = b.line(s, tx, ty, tx + 3.94, ty, LIGHT, 1.5, name=f"Rule {it['label'][:12]}")
            groups.append(self.item(rule, n, h, f))
        self.finish(s, groups, [i["label"] for i in sl["items"]], sl.get("notes"), sl.get("pillar"))
        return s

    def priorities(self, sl):
        b = self.b
        s = b.slide("Strategic Priorities", sl["title"])
        tops = [1.96, 3.44, 4.96]
        groups = []
        for i, r in enumerate(sl["rows"][:3]):
            lab = b.fill(s, 10 + i, r["label"])
            txt = b.fill(s, 13 + i, r["lines"], size=18)
            txt.left, txt.top, txt.width, txt.height = Inches(4.13), Inches(tops[i]), Inches(6.9), Inches(1.14)
            ico = self.icon_for(r, "label")
            ic = b.icon(s, ico, 11.7, tops[i] + 0.04, 0.85, "circle", WHITE, LIGHT) if ico else []
            groups.append(self.item(lab, txt, ic))
        if sl.get("statement"):
            st = b.text(s, 1.19, 6.4, 11.4, 0.6, sl["statement"], 24, LIGHT, True, name="Statement")
            groups.append(self.item(st))
        self.finish(s, groups, [r["label"] for r in sl["rows"]] + ["Takeaway"], sl.get("notes"), sl.get("pillar"))
        return s

    def agenda(self, sl):
        b = self.b
        s = b.slide("ToC / Agenda Slide", None)
        b.fill(s, 27, sl["title"], name="!!title")
        ys = [1.24, 2.20, 3.15, 4.11, 5.06, 6.02]
        groups = []
        for i, it in enumerate(sl["items"][:6]):
            t = b.fill(s, 34 + i, [{"t": it["head"], "bold": True}] + ([{"t": it["text"], "size": 16}] if it.get("text") else []))
            dot = b.dot(s, 4.62, ys[i] - 0.01, 0.36, LIGHT, name=f"Number {i+1}")
            tf = dot.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = str(i + 1)
            r.font.size, r.font.bold = Pt(14), True
            r.font.color.rgb = RGBColor.from_string(WHITE)
            groups.append(self.item(dot, t))
        if sl.get("hero"):
            h1 = b.text(s, 0.38, 3.75, 3.4, 1.9, sl["hero"], 110, WHITE, True, name="!!hero")
            h2 = b.text(s, 0.46, 5.55, 3.4, 0.5, sl.get("hero_label", ""), 28, WHITE, name="Hero label")
            groups.insert(0, self.item(h1, h2))
        self.finish(s, groups, (["Hero"] if sl.get("hero") else []) + [i["head"] for i in sl["items"]], sl.get("notes"),
                    sl.get("pillar"))
        return s

    def timeline(self, sl):
        b = self.b
        s = b.slide("Event Roadmap/Timeline", sl["title"])
        TXT, LAB, CX = [17, 33, 34, 35, 36, 37], [38, 23, 42, 41, 40, 39], [0.46, 2.55, 4.63, 6.72, 8.81, 10.89]
        groups = []
        for i, ph in enumerate(sl["phases"][:6]):
            t = b.open_up(b.fill(s, TXT[i], [{"t": ph["head"], "bold": True, "size": 20}, {"t": ph.get("text", ""), "size": 16}],
                                 top_inset=1.0))
            hl = b.line(s, CX[i], 1.62, CX[i] + 1.99, 1.62, BLUE, 1.5, name=f"Rule phase {i+1}")
            ico = self.icon_for(ph)
            ic = b.icon(s, ico, CX[i] + 0.62, 1.74, 0.75, "bare", BLUE) if ico else []
            lb = b.fill(s, LAB[i], ph.get("when", ""))
            groups.append(self.item(t, hl, ic, lb))
        self.finish(s, groups, [p["head"] for p in sl["phases"]], sl.get("notes"), sl.get("pillar"))
        return s

    def icons(self, sl):
        """Three or four columns: big icon, head, text. Custom slide, dark or light."""
        b = self.b
        dark = self.is_dark(sl)
        s = self.custom(sl["title"], sl.get("sub"), dark)
        ink, acc = (WHITE, LIGHT) if dark else (NAVY, BLUE)
        items = sl["items"][:4]
        n = len(items)
        cw = 12.4 / n
        groups = []
        for i, it in enumerate(items):
            cx = 0.46 + i * cw
            ico = self.icon_for(it)
            ic = b.icon(s, ico, cx, 2.35, 1.1, "circle", WHITE, acc) if ico else []
            hd = b.text(s, cx, 3.6, cw - 0.3, 0.85, it["head"], 24, ink, True, anchor=MSO_ANCHOR.BOTTOM, name=f"Head {i+1}")
            rl = b.line(s, cx, 4.55, cx + cw - 0.4, 4.55, acc, 1.5, name=f"Rule {i+1}")
            tx = b.text(s, cx, 4.7, cw - 0.3, 1.5, it.get("text", ""), 18, ink if dark else CHAR, name=f"Text {i+1}")
            groups.append(self.item(ic, hd, rl, tx))
        if sl.get("statement"):
            groups.append(self.item(b.text(s, 0.46, 6.4, 12.4, 0.5, sl["statement"], 24, acc, True, name="Statement")))
        self.finish(s, groups, [i["head"] for i in items] + ["Takeaway"], sl.get("notes"), sl.get("pillar"))
        return s

    def rows(self, sl):
        """Up to six numbered rows with icon, text and hairline (two columns of three)."""
        b = self.b
        dark = self.is_dark(sl)
        s = self.custom(sl["title"], sl.get("sub"), dark)
        ink, acc = (WHITE, LIGHT) if dark else (NAVY, BLUE)
        groups = []
        for i, it in enumerate(sl["items"][:6]):
            col, row = i % 2, i // 2
            x, y = 0.46 + col * 6.5, 2.2 + row * 1.2
            ico = self.icon_for(it, "text")
            ic = b.icon(s, ico, x, y, 0.8, "circle", WHITE, acc) if ico else []
            tx = b.text(s, x + 1.05, y + 0.05, 5.0, 0.7, f"{i+1:02d}  {it['text']}", 20, ink, True, anchor=MSO_ANCHOR.MIDDLE,
                        name=f"Item {i+1}")
            rl = b.line(s, x, y + 1.0, x + 5.9, y + 1.0, ink if dark else BLUE, 1.0, 0.12 if dark else 0.3, name=f"Rule {i+1}")
            groups.append(self.item(ic, tx, rl))
        paired = [groups[k] + (groups[k + 1] if k + 1 < len(groups) else []) for k in range(0, len(groups), 2)]
        if sl.get("statement"):
            paired.append(self.item(b.text(s, 0.46, 6.35, 12.4, 0.6, sl["statement"], sl.get("statement_size", 28), acc, True, name="Statement")))
        self.finish(s, paired, [f"Items {2*k+1}-{2*k+2}" for k in range(len(paired))], sl.get("notes"), sl.get("pillar"))
        return s

    def exec_summary(self, sl):
        """The answer first: one hero sentence left, up to three supports right."""
        b = self.b
        dark = self.is_dark(sl)
        s = self.custom(sl["title"], None, dark)
        ink, acc = (WHITE, LIGHT) if dark else (NAVY, BLUE)
        hero = b.text(s, 0.46, 2.0, 5.2, 3.6, sl["answer"], 36, acc, True, name="!!hero")
        groups = [self.item(hero)]
        for i, p in enumerate(sl["points"][:3]):
            y = 1.9 + i * 1.5
            num = b.text(s, 6.4, y, 0.8, 0.8, f"{i+1}", 44, acc, True, name=f"Num {i+1}")
            hd = b.text(s, 7.3, y, 5.5, 0.5, p["head"], 24, ink, True, name=f"Head {i+1}")
            tx = b.text(s, 7.3, y + 0.55, 5.5, 0.8, p.get("text", ""), 18, ink if dark else CHAR, name=f"Text {i+1}")
            rl = b.line(s, 6.4, y + 1.35, 12.85, y + 1.35, ink if dark else BLUE, 1.0, 0.12 if dark else 0.3, name=f"Rule {i+1}")
            groups.append(self.item(num, hd, tx, rl))
        self.finish(s, groups, ["Answer"] + [p["head"] for p in sl["points"]], sl.get("notes"), sl.get("pillar"))
        return s

    def ask(self, sl):
        """The decision or ask: hero sentence plus numbered next steps."""
        sl = {**sl, "items": [{"text": t} for t in sl["steps"]]}
        s = self.rows({**sl, "statement": sl.get("ask")})
        return s

    def table(self, sl):
        slides = paginate_table(self.b, "2_Title Only – Blank", sl["title"], sl["header"], sl["rows"], sl["widths"],
                                top=sl.get("top", 2.0), max_y=6.3, dark=False)
        for s in slides:
            if sl.get("sub"):
                self.b.text(s, 0.46, 1.45, 12.4, 0.4, sl["sub"], 18, BLUE, name="!!sub")
            s.notes_slide.notes_text_frame.text = sl.get("notes", "")
            self.d_light_footer(s)
        return slides

    def d_light_footer(self, s):
        pass

    def exhibit(self, sl, draw):
        b = self.b
        dark = self.is_dark(sl)
        s = self.custom(sl["title"], sl.get("sub"), dark)
        draw(s, dark)
        shapes = []
        if sl.get("takeaway"):
            shapes.append(takeaway(b, s, sl["takeaway"], dark=dark))
        if sl.get("source"):
            source_line(b, s, sl["source"], dark=dark)
        self.finish(s, [self.item(shapes)] if shapes else [], ["Takeaway"], sl.get("notes"), sl.get("pillar"))
        return s

    def chart(self, sl):
        return self.exhibit(sl, lambda s, dark: add_chart(s, sl.get("kind", "column"), sl["categories"], sl["series"], 0.46, 1.95,
                                                          12.4, 4.2, dark=dark, number_format=sl.get("number_format", "General")))

    def waterfall(self, sl):
        return self.exhibit(sl, lambda s, dark: waterfall(s, sl["labels"], sl["values"], 0.46, 1.95, 12.4, 4.2, dark=dark,
                                                          highlight=sl.get("highlight")))

    def matrix(self, sl):
        return self.exhibit(sl, lambda s, dark: matrix2x2(self.b, s, 1.2, 1.95, 7.6, 3.9, sl["x_label"], sl["y_label"], sl["quads"],
                                                          [tuple(p) for p in sl.get("points", [])], dark, sl.get("highlight")))

    def gantt(self, sl):
        return self.exhibit(sl, lambda s, dark: gantt(self.b, s, [tuple(t) for t in sl["tasks"]], sl["periods"], 0.46, 1.95, 12.4,
                                                      row_h=sl.get("row_h", min(0.9, 4.2 / (len(sl["tasks"]) + 1))), dark=dark, highlight=sl.get("highlight")))

    def ratings(self, sl):
        return self.exhibit(sl, lambda s, dark: rating_table(self.b, s, sl["header"], [tuple(r) for r in sl["rows"]], 0.46, 1.95,
                                                             sl["widths"], dark=dark, highlight_row=sl.get("highlight")))

    def funnel(self, sl):
        return self.exhibit(sl, lambda s, dark: funnel(self.b, s, [tuple(t) for t in sl["stages"]], 0.46, 1.95, 12.4, 4.2, dark=dark,
                                                       highlight=sl.get("highlight")))

    def end(self, sl):
        s = self.d.prs.slides.add_slide(self.b.layout("End Card"))
        s.notes_slide.notes_text_frame.text = sl.get("notes", "End card.")
        return s

    # ---- run -----------------------------------------------------------------------------------
    def build(self, out):
        made = []
        for sl in self.sp["slides"]:
            fn = getattr(self, sl["type"].replace("-", "_"))
            res = fn(sl)
            res = res if isinstance(res, list) else [res]
            for s in res:
                made.append((sl, s))
                if sl.get("appendix"):
                    s._element.set("show", "0")
                    self.d.transition(s, "cut")
        kind = self.sp.get("transition", "fade")
        for sl, s in made:
            if get_transition(s) is None:
                if sl["type"] in ("cover", "end"):
                    self.d.transition(s, "fade_black", 1500)
                else:
                    self.d.transition(s, kind, 700 if kind == "fade" else 1200)
        self.d.finish(out)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    spec = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    Spec(spec).build(sys.argv[2])
    return 0


if __name__ == "__main__":
    sys.exit(main())
