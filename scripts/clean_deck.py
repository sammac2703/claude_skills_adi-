"""Route C helper: clean, text-led ADI decks with click-through builds and Morph transitions.

    from clean_deck import CleanDeck, TC
    d = CleanDeck(classification="public")
    d.cover("Title", "Subtitle", details=["Informal meeting", "[TO CONFIRM: date]"])
    s = d.slide("Slide Title")
    d.sub(s, "Intro line")                                   # visible from the start, named !!sub for Morph
    items = [d.numbered(s, 0.46, 2.25, 5.9, 1, "First point")]
    d.build(s, [items], ["First point"], "Speaker notes ...")  # one click per group
    d.finish("deck.pptx")

Rules baked in (see references/keynote-style.md and references/clean-click-through.md): style="keynote" (default) puts
every content slide on the template's Dark Blue layout with a black-to-ADI-dark-blue diagonal canvas, white type, ADI light
blue as the single accent and 12% white hairlines; cards (d.card), icon circles (d.icon) and hero numbers (d.hero) are
available, three or four per slide at most. style="light" keeps the original text-only light look. Either way: each point
appears on its own click; Morph on every slide; CLICKS cue line at the top of the notes.
"""
from __future__ import annotations

import copy
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx.enum.dml import MSO_THEME_COLOR as TC            # noqa: E402,F401
from pptx.enum.shapes import MSO_CONNECTOR                # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR           # noqa: E402,F401
from pptx.util import Pt, Inches                          # noqa: E402

from pptx.dml.color import RGBColor                       # noqa: E402
from pptx.enum.shapes import MSO_SHAPE                    # noqa: E402
from pptx.oxml.ns import qn                               # noqa: E402
from lxml import etree                                    # noqa: E402

from deckkit import open_template, _set_alpha             # noqa: E402
from motion import Timeline, set_transition, get_transition, count_clicks  # noqa: E402

SKILL = Path(__file__).resolve().parent.parent
TEMPLATE = SKILL / "assets" / "ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx"
FOOTERS = {
    "public": "©2026 Analog Devices, Inc. All Rights Reserved.",
    "confidential": "Analog Devices Confidential Information. ©2026 Analog Devices, Inc. All Rights Reserved.",
    "internal": "Analog Devices Confidential Information—Not for External Distribution. ©2026 Analog Devices, Inc. All Rights Reserved.",
}
W, H = 13.333, 7.5

# Keynote canvas (references/keynote-style.md): ADI palette only.
CANVAS = ("000000", "00325C")   # black to ADI dark blue, diagonal
CANVAS_ANGLE = 315              # top-left (black) to bottom-right (dark blue)
CARD_FILL, CARD_OPACITY = "3C4157", 0.6   # ADI dark grey over the canvas reads as charcoal
HAIRLINE_OPACITY = 0.12
# On the dark canvas the light-style colour roles are remapped so existing slide code still reads:
# labels and body go white, closing lines and status lines go to the one accent (ADI light blue).
KEYNOTE_COLORS = {TC.TEXT_1: TC.BACKGROUND_1, TC.ACCENT_1: TC.BACKGROUND_1,
                  TC.TEXT_2: TC.ACCENT_2, TC.ACCENT_3: TC.ACCENT_2}


class CleanDeck:
    def __init__(self, classification: str = "public", template: str | None = None, style: str = "keynote"):
        self.prs = open_template(str(template or TEMPLATE))
        self.layouts = {l.name: l for l in self.prs.slide_layouts}
        self.footer_text = FOOTERS[classification]
        self.keynote = style == "keynote"

    def col(self, color):
        """Colour role for the current style (keynote remaps light-style roles onto the dark canvas)."""
        return KEYNOTE_COLORS.get(color, color) if self.keynote else color

    @staticmethod
    def canvas(slide):
        """Black to ADI dark blue diagonal gradient on the slide background. The layout's white logo, footer
        and slide number still draw on top, unmoved."""
        fill = slide.background.fill
        fill.gradient()
        fill.gradient_angle = CANVAS_ANGLE
        st = fill.gradient_stops
        st[0].color.rgb, st[1].color.rgb = RGBColor.from_string(CANVAS[0]), RGBColor.from_string(CANVAS[1])

    # ---- layout / placeholder helpers ------------------------------------------------
    def L(self, prefix):
        return next(l for n, l in self.layouts.items() if n.startswith(prefix))

    @staticmethod
    def ph(slide, idx):
        return next(s for s in slide.placeholders if s.placeholder_format.idx == idx)

    def drop_ph(self, slide, idx):
        el = self.ph(slide, idx)._element
        el.getparent().remove(el)

    def add_footer(self, slide):
        for p in slide.slide_layout.placeholders:
            if p.placeholder_format.idx in (3, 4):
                slide.shapes._spTree.append(copy.deepcopy(p._element))
        self.ph(slide, 3).text_frame.text = self.footer_text

    # ---- slides ------------------------------------------------------------------------
    def cover(self, title, subtitle, details=None, notes=""):
        s = self.prs.slides.add_slide(self.L("Cover with Pattern"))
        self.ph(s, 10).text_frame.text = title
        self.ph(s, 14).text_frame.text = subtitle
        self.drop_ph(s, 15)
        if details:
            self.text(s, 0.47, 6.2, 8.0, 0.7, details, 14, TC.BACKGROUND_1, name="Meeting details")
        s.notes_slide.notes_text_frame.text = notes
        return s

    def slide(self, title, dark=False):
        """Content slide. Keynote style: always the 'Title and Content - Dark Blue' layout on the diagonal canvas.
        Light style: 'Title Only', or the Dark Blue layout with dark=True for a key-message / discussion slide."""
        dark = dark or self.keynote
        s = self.prs.slides.add_slide(self.L("Title and Content - Dark Blue" if dark else "Title Only"))
        s.shapes.title.text = title
        if dark:
            self.drop_ph(s, 11)
        if self.keynote:
            self.canvas(s)
        return s

    def sub(self, slide, s, color=None, bold=False):
        """Intro / status line under the title (the accent subtitle in keynote style). Visible from the start;
        shared name lets Morph glide it."""
        color = color or (TC.ACCENT_2 if self.keynote else TC.TEXT_1)
        return self.text(slide, 0.46, 1.45, 12.4, 0.4, s, 18, color, bold, name="!!sub")

    # ---- primitives -------------------------------------------------------------------
    def text(self, slide, x, y, w, h, s, size, color=TC.TEXT_1, bold=False, anchor=MSO_ANCHOR.TOP, name=None,
             brightness=0.0, align=PP_ALIGN.LEFT):
        color = self.col(color)
        t = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = t.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, ln in enumerate(s if isinstance(s, list) else [s]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = ln
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.theme_color = color
            if brightness:
                r.font.color.brightness = brightness
        if name:
            t.name = name
        return t

    def rule(self, slide, x, y, w, color=TC.BACKGROUND_2, name=None):
        """Thin rule. Keynote style draws every rule as a faint white hairline (12% opacity)."""
        hairline = self.keynote          # keynote: every rule is a faint hairline, never a bright rule
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
        c.line.color.theme_color = TC.BACKGROUND_1 if hairline else self.col(color)
        c.line.width = Pt(1.0 if hairline else 1.25)
        if hairline:
            clr = c.line._get_or_add_ln().find(qn("a:solidFill"))[0]
            etree.SubElement(clr, qn("a:alpha")).set("val", str(int(HAIRLINE_OPACITY * 100000)))
        c.name = name or "Rule"
        return c

    # ---- keynote elements (references/keynote-style.md): three or four per slide at most -------------------
    def card(self, slide, x, y, w, h, head, body=None, n=1, head_size=20, body_size=16, key=None):
        """Dark charcoal panel with a slim accent header tab, a white bold heading and short body lines.
        Returns the shapes as one build item. key='card1' names them !!card1, !!card1_tab ... so Morph can grow
        the card into a deep dive on the next slide (the zoom-in move)."""
        nm = (lambda part: f"!!{key}{part}") if key else (lambda part: f"Card {n}{part}")
        panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        panel.fill.solid()
        panel.fill.fore_color.rgb = RGBColor.from_string(CARD_FILL)
        _set_alpha(panel, CARD_OPACITY)
        panel.line.fill.background()
        panel.shadow.inherit = False
        panel.name = nm("")
        tab = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.08))
        tab.fill.solid()
        tab.fill.fore_color.theme_color = TC.ACCENT_2
        tab.line.fill.background()
        tab.shadow.inherit = False
        tab.name = nm("_tab")
        shapes = [panel, tab, self.text(slide, x + 0.3, y + 0.35, w - 0.6, 0.6, head, head_size, TC.BACKGROUND_1, True,
                                        name=nm("_head"))]
        if body:
            shapes.append(self.text(slide, x + 0.3, y + 1.0, w - 0.6, h - 1.2, body, body_size, TC.BACKGROUND_1,
                                    name=nm("_body")))
        return shapes

    def icon(self, slide, x, y, d=1.0, picture=None, glyph=None, n=1, key=None):
        """Line icon inside an accent circle. picture: an icon exported from the Brand Icon Library, recoloured
        white (PNG/SVG); glyph: a short stand-in (a digit or symbol) when no approved icon is to hand.
        Keep every icon on a slide the same diameter."""
        name = f"!!{key}" if key else f"Icon {n}"
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        c.fill.solid()
        c.fill.fore_color.theme_color = TC.ACCENT_2
        c.line.fill.background()
        c.shadow.inherit = False
        c.name = name
        out = [c]
        if picture:
            p = slide.shapes.add_picture(str(picture), Inches(x + d * 0.22), Inches(y + d * 0.22), Inches(d * 0.56),
                                         Inches(d * 0.56))
            p.name = name + "_art"
            out.append(p)
        elif glyph:
            out.append(self.text(slide, x, y, d, d, glyph, max(14, int(d * 28)), TC.BACKGROUND_1, True,
                                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, name=name + "_glyph"))
        return out

    def hero(self, slide, key, x, y, w, h, value, size=120, align=PP_ALIGN.CENTER):
        """Hero number or two-to-three word statement in the accent, named !!key so it can shrink into a corner
        card on the next slide (hero-number-to-detail move)."""
        return self.text(slide, x, y, w, h, value, size, TC.ACCENT_2, True, name=f"!!{key}", align=align,
                         anchor=MSO_ANCHOR.MIDDLE)

    def numbered(self, slide, x, y, w, n, label, size=22, color=TC.ACCENT_1, num_color=TC.ACCENT_2,
                 num_bright=0.0, line_color=TC.BACKGROUND_2):
        """One build item: pale numeral, bold label, thin rule. Returns the list of shapes to reveal together."""
        return [self.text(slide, x, y, 0.8, 0.6, f"{n:02d}", 28, num_color, True, brightness=num_bright, name=f"Number {n}"),
                self.text(slide, x + 0.95, y, w - 0.95, 0.9, label, size, color, True, name=f"Item {n}"),
                self.rule(slide, x, y + 0.95, w, line_color, name=f"Rule {n}")]

    def column(self, slide, x, y, w, head, body, head_size=30, body_size=20, n=1, rule_color=TC.ACCENT_2):
        """Heading, thin rule and body text: one build item for column-style slides."""
        return [self.text(slide, x, y, w, 0.6, head, head_size, TC.ACCENT_1, True, name=f"Heading {n}"),
                self.rule(slide, x, y + 0.65, w - 0.2, rule_color, name=f"Rule {n}"),
                self.text(slide, x, y + 0.9, w - 0.1, 1.8, body, body_size, TC.TEXT_1, name=f"Body {n}")]

    def statement(self, slide, y, s, size=24, color=TC.TEXT_2, h=0.7, name="Statement"):
        """Closing or key line shown on the final click of a slide."""
        return [self.text(slide, 0.46, y, 12.4, h, s, size, color, True, name=name)]

    # ---- Morph state-pair helpers (technique: duplicate a slide, move/resize/restyle shapes, keep the names) ----
    # Morph pairs shapes that share a name. These helpers give every shape a stable "!!key" name, so the
    # SAME key on consecutive slides glides between its two states (position, size, font size, colour).
    def m_text(self, slide, key, x, y, w, h, s, size, color=TC.TEXT_1, bold=False, **kw):
        return self.text(slide, x, y, w, h, s, size, color, bold, name=f"!!{key}", **kw)

    def m_rule(self, slide, key, x, y, w, color=TC.BACKGROUND_2):
        return self.rule(slide, x, y, w, color, name=f"!!{key}")

    def park(self, slide, key, kind="text", edge="right", x=None, y=None, w=3.0):
        """Add an off-canvas copy of a shape that exists on a neighbouring slide, so Morph slides it in or out
        instead of fading it. edge: right | left | bottom | top. Invisible in the show; no text, so nothing leaks."""
        px = {"right": W + 1.0, "left": -w - 1.0}.get(edge, x if x is not None else 0.46)
        py = {"bottom": H + 1.0, "top": -1.5}.get(edge, y if y is not None else 2.0)
        if kind == "rule":
            return self.m_rule(slide, key, px, py, w)
        return self.m_text(slide, key, px, py, w, 0.5, "", 18)

    def focus_row(self, title, items, y=2.3, intro=None, hero_w=4.6, size_hero=30, size_other=18):
        """Text-only accordion / spotlight: one slide per item. The active item widens and shows its description;
        the others narrow and shrink (their description collapses to 1 pt, a trick from the Gold deck so Morph
        shrinks it instead of cutting). items: list of (short label, description). Returns the slides."""
        n = len(items)
        others = (12.42 - hero_w - 0.25 * (n - 1)) / (n - 1)
        slides = []
        for s in range(n):
            sl = self.slide(title)
            if intro:
                self.sub(sl, intro)
            x = 0.46
            for i, (label, desc) in enumerate(items):
                hero = i == s
                w = hero_w if hero else others
                self.m_text(sl, f"fl{i}", x, y, w, 0.7, label, size_hero if hero else size_other,
                            (TC.ACCENT_2 if self.keynote else TC.ACCENT_1) if hero else (TC.BACKGROUND_1 if self.keynote else TC.ACCENT_2), True)
                self.m_rule(sl, f"flr{i}", x, y + 0.8, w - 0.1, TC.ACCENT_2 if hero else TC.BACKGROUND_2)
                self.m_text(sl, f"fld{i}", x, y + 1.05, (w - 0.1) if hero else 0.3, 2.0 if hero else 0.2,
                            desc, 20 if hero else 1, TC.TEXT_1)
                x += w + 0.25
            self.add_footer(sl)
            slides.append(sl)
        return slides

    # ---- arrows, links ------------------------------------------------------------------
    def arrow(self, slide, x1, y1, x2, y2, color=TC.ACCENT_2, name="Arrow"):
        """Thin connector with an arrowhead at the end (used with builds so the arrow lands with its label)."""
        from pptx.oxml.ns import qn
        from lxml import etree
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.theme_color = color
        c.line.width = Pt(2)
        ln = c.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        c.name = name
        return c

    @staticmethod
    def link(shape, target_slide):
        """Click this shape in the show to jump to another slide (safe alternative to Slide Zoom, which needs
        PowerPoint's own Insert > Zoom to generate its thumbnail). Pair with a 'Back' link on the target."""
        shape.click_action.target_slide = target_slide
        return shape

    def transition(self, slide, kind="morph", ms=1000, **opts):
        """Set a slide's transition now; finish() leaves it alone. Morph state pairs read well at 1500-2000 ms."""
        set_transition(slide, kind, duration_ms=ms, **opts)

    # ---- animation + notes ------------------------------------------------------------
    def build(self, slide, groups, labels, notes_text, effect="float_up", stagger_ms=220):
        """groups: one entry per click; each entry is a list of items; each item is a list of shapes shown together.
        Items inside a click land one after another (stagger_ms apart), so put an arrow item before the step it
        points to. A shape may be given as (shape, paragraph_index) to reveal a single paragraph of a text box,
        which lets text lines, pictures and diagram parts interleave in any order. Rules wipe in from the left."""
        tl = Timeline(slide, W, H)
        for g, label in zip(groups, labels):
            for k, item in enumerate(g):
                for j, sh in enumerate(item):
                    first = (k == 0 and j == 0)
                    para = None
                    if isinstance(sh, tuple):
                        sh, para = sh
                    wipe = para is None and (sh.shape_type == 9 or sh.name.startswith("Rule") or sh.name.startswith("!!rule"))
                    start = "click" if first else "with"
                    delay = 0 if first else k * stagger_ms
                    if para is None:
                        tl.add(sh, "wipe" if wipe else effect, start=start, delay=delay, duration=500,
                               label=label if first else None, **(dict(direction="left") if wipe else {}))
                    else:
                        tl.items.append(dict(shape=sh, effect=effect, start=start, delay=delay, duration=500,
                                             direction=None, para=para, by=None, iterate_ms=60, smooth=True,
                                             scale=1.1, opacity=0.35, dx=0.0, dy=0.0, repeat=None,
                                             label=label if first else None))
        clicks = tl.commit()
        plan = f"CLICKS ({clicks}): " + "  ".join(f"[{i+1}] {l}." for i, l in enumerate(labels))
        slide.notes_slide.notes_text_frame.text = plan + "\n\n" + notes_text
        return clicks

    # ---- finish + QA -------------------------------------------------------------------
    def finish(self, out, morph_ms=1200):
        """Morph on every slide, save, print click counts and a rough text-fit check."""
        for sl in self.prs.slides:
            if get_transition(sl) is None:          # slides given a transition via d.transition() keep theirs
                set_transition(sl, "morph", duration_ms=morph_ms)
        self.prs.save(out)
        total = 0
        for i, sl in enumerate(self.prs.slides, 1):
            c = count_clicks(sl)
            total += c
            print(f"slide {i}: {c} clicks")
            for sh in sl.shapes:
                if not sh.has_text_frame or sh.is_placeholder or sh.name == "Slide Number" or not sh.text_frame.text.strip():
                    continue
                w = sh.width / 914400
                need = 0.0
                for para in sh.text_frame.paragraphs:
                    sz = max([r.font.size.pt for r in para.runs if r.font.size] or [18])
                    need += math.ceil(max(1, len(para.text)) / max(1, int(w * 72 / (sz * 0.5)))) * sz * 1.2 / 72
                if need > sh.height / 914400 * 1.02 and (sh.top / 914400 + need) > 7.0:
                    print(f"   CHECK overflow risk: {sh.name} needs {need:.2f}in")
        print(f"saved {out}: {len(self.prs.slides)} slides, {total} clicks. Not visually rendered: open in PowerPoint and run Slide Show once.")
