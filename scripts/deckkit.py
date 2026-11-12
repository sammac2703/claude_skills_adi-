"""
deckkit: the drawing engine behind the ADI deck builder skill.

Everything that turns a slide spec (a dict) into shapes lives here:
theme handling, template loading (.pptx or .potx), the layout library,
Morph transitions, off-canvas parking, picture effects and speaker notes.

build_deck.py is the command-line front end; import this module directly
if you want to compose slides by hand.

Units: layout maths is done in inches (floats) and converted with IN().
"""
from __future__ import annotations

import copy
import json
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

EMU_PER_IN = 914400
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"


def IN(x: float) -> Emu:
    return Emu(int(round(x * EMU_PER_IN)))


def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str.strip("#").upper())


def lerp_hex(a: str, b: str, t: float) -> str:
    a, b = a.strip("#"), b.strip("#")
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def luminance(hex_str: str) -> float:
    h = hex_str.strip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    lin = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------

DEFAULT_THEME = {
    "name": "placeholder-dark",
    "_note": "PLACEHOLDER palette. Replace with real ADI template values (run extract_template.py).",
    "template": None,                 # path to the ADI .potx/.pptx, or null
    "blank_layout": ["Blank", "Title Only"],  # layout names to try, in order
    "mode": "dark",                   # dark | light | template (template = keep master background)
    "fonts": {
        "heavy": {"face": "Arial", "bold": True},     # huge titles, stage letters, big numbers
        "head": {"face": "Arial", "bold": True},      # subheads
        "body": {"face": "Calibri", "bold": False},   # normal text
        "light": {"face": "Calibri Light", "bold": False},  # body that must not compete
    },
    "colors": {
        "dark_bg": "0D1117", "dark_bg2": "000000",
        "light_bg": "FFFFFF", "light_bg2": "EEF1F5",
        "on_dark": "FFFFFF", "on_dark_muted": "9AA4B2",
        "on_light": "16202C", "on_light_muted": "5D6875",
        "accent": "2F8CFF", "accent2": "7CC4FF",
        "ramp_start": "0A2540", "ramp_end": "A9D2F7",
        "rival": "6B7685", "ours": "2F8CFF", "positive": "34C38F", "warn": "F5A524",
    },
    "footer": {"text": "PLACEHOLDER | Add ADI confidentiality line from template",
               "slide_numbers": True},
    "logo": {"path": None, "width_in": 1.1, "position": "br"},
    "photo_overlay_opacity": 0.55,    # 0.55 opacity = 45% transparency, per the videos
    "morph_ms": 1500,
    # Motion system (see references/motion.md). level: minimal | standard | showcase
    "motion": {
        "level": "standard",
        "base": "fade", "base_ms": 400,            # every slide without a Morph
        "hook": "fade_black", "hook_ms": 1000,     # first slide: up from black
        "section": "fade_black", "section_ms": 800,  # topic change
        "appendix": "cut",                         # reference material: no transition
        "anim_ms": 400,                            # default in-slide entrance length
    },
}


def deep_merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_theme(theme_ref, base_dir: Path | None = None) -> dict:
    """theme_ref: None, a path to a theme JSON, or an inline dict of overrides."""
    if theme_ref is None:
        return copy.deepcopy(DEFAULT_THEME)
    if isinstance(theme_ref, str):
        p = Path(theme_ref)
        if not p.is_absolute() and base_dir:
            p = (base_dir / p).resolve()
        data = json.loads(p.read_text(encoding="utf-8"))
        if data.get("template") and not Path(data["template"]).is_absolute():
            data["template"] = str((p.parent / data["template"]).resolve())
        if (data.get("logo") or {}).get("path") and not Path(data["logo"]["path"]).is_absolute():
            data["logo"]["path"] = str((p.parent / data["logo"]["path"]).resolve())
        return deep_merge(DEFAULT_THEME, data)
    return deep_merge(DEFAULT_THEME, theme_ref)


def open_template(path: str | None) -> Presentation:
    """Open a .pptx or .potx. python-pptx refuses .potx because of its content
    type, so a patched temporary copy is opened instead. Existing slides are removed."""
    if not path:
        prs = Presentation()
        prs.slide_width, prs.slide_height = IN(13.333), IN(7.5)
        return prs
    src = Path(path)
    if src.suffix.lower() in (".potx", ".potm"):
        tmp = Path(tempfile.mkdtemp()) / (src.stem + ".pptx")
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = re.sub(rb"presentationml\.template\.main\+xml",
                                  b"presentationml.presentation.main+xml", data)
                    data = data.replace(b"application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
                                        b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
                zout.writestr(item, data)
        prs = Presentation(str(tmp))
    else:
        prs = Presentation(str(src))
    # drop any example slides in the template
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    return prs


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def _set_alpha(shape, opacity: float):
    """Opacity 0..1 on a solid fill (0.55 opacity = 45% transparency)."""
    clr = shape._element.spPr.find(qn("a:solidFill"))
    if clr is None or len(clr) == 0:
        return
    c = clr[0]
    for old in c.findall(qn("a:alpha")):
        c.remove(old)
    a = etree.SubElement(c, qn("a:alpha"))
    a.set("val", str(int(opacity * 100000)))


def _spPr_insert_effect(spPr, effect_xml: str):
    eff = spPr.find(qn("a:effectLst"))
    if eff is None:
        eff = etree.SubElement(spPr, qn("a:effectLst"))
        # effectLst must come after fill and ln
        ln = spPr.find(qn("a:ln"))
        if ln is not None:
            ln.addnext(eff)
    eff.append(etree.fromstring(effect_xml))


def add_outer_shadow(shape, blur_pt=16, dist_pt=4, opacity=0.45, color="000000"):
    """Subtle drop shadow. Videos: images ~55% transparency / 16 pt blur."""
    _spPr_insert_effect(shape._element.spPr, (
        f'<a:outerShdw xmlns:a="{NS_A}" blurRad="{int(blur_pt*12700)}" dist="{int(dist_pt*12700)}" '
        f'dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="{color}">'
        f'<a:alpha val="{int(opacity*100000)}"/></a:srgbClr></a:outerShdw>'))


def add_inner_shadow(shape, blur_pt=10, dist_pt=6, opacity=0.6):
    """Inside top-left inner shadow: the layered paper look on cut-outs."""
    _spPr_insert_effect(shape._element.spPr, (
        f'<a:innerShdw xmlns:a="{NS_A}" blurRad="{int(blur_pt*12700)}" dist="{int(dist_pt*12700)}" '
        f'dir="2700000"><a:srgbClr val="000000"><a:alpha val="{int(opacity*100000)}"/></a:srgbClr></a:innerShdw>'))


def set_grayscale(picture):
    blip = picture._element.find(".//" + qn("a:blip"))
    if blip is not None and blip.find(qn("a:grayscl")) is None:
        blip.insert(0, etree.Element(qn("a:grayscl")))


def set_picture_opacity(picture, opacity: float):
    blip = picture._element.find(".//" + qn("a:blip"))
    if blip is not None:
        el = etree.SubElement(blip, qn("a:alphaModFix"))
        el.set("amt", str(int(opacity * 100000)))


class _NullTimeline:
    clicks = 0
    items = []

    def add(self, *a, **k):
        return self

    def commit(self):
        return 0


def add_morph(slide, duration_ms=1500, kind="morph"):
    """Morph transition on this slide (slide B of a state pair). kind: morph (by object),
    morph_words or morph_chars. Written as PowerPoint writes it, with a fade fallback."""
    from motion import set_transition
    set_transition(slide, kind, duration_ms=duration_ms)


def has_morph(slide) -> bool:
    return slide._element.find(".//{%s}morph" % NS_P159) is not None


def set_slide_tag(slide, **tags):
    """Store role metadata in p:cSld/@name (valid OOXML, kept by PowerPoint,
    invisible to the audience). Read back by lint_deck.py."""
    slide._element.find(qn("p:cSld")).set(
        "name", ";".join(f"{k}={v}" for k, v in tags.items() if v not in (None, "")))


def element_name(el) -> str | None:
    nv = el.find(".//" + qn("p:cNvPr"))
    return nv.get("name") if nv is not None else None


def top_level_named(slide, prefix="!!"):
    out = {}
    for el in slide.shapes._spTree.iterchildren():
        if el.tag in (qn("p:sp"), qn("p:pic"), qn("p:grpSp"), qn("p:cxnSp"), qn("p:graphicFrame")):
            nm = element_name(el)
            if nm and nm.startswith(prefix):
                out[nm] = el
    return out


def _xfrm_off_ext(el):
    if el.tag == qn("p:grpSp"):
        x = el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    elif el.tag == qn("p:graphicFrame"):
        x = el.find(qn("p:xfrm"))
    else:
        x = el.find(qn("p:spPr")).find(qn("a:xfrm"))
    return x.find(qn("a:off")), x.find(qn("a:ext"))


# ---------------------------------------------------------------------------
# Deck builder
# ---------------------------------------------------------------------------

_UNIT_WORD = re.compile(r"^[(\[]?([pnµμmkMG]?(W|V|A|Hz|F|Ω|bit|bits|B|dB|dBm|SPS)(/[\w°√]+)?|[pnµμm]s|[\w.]*\d[\w./°%-]*|.*[µμ°Ω/√].*)[)\],.;:]?$")


def caps_safe(text):
    """Uppercase a line for kickers without corrupting units or values:
    words that are numbers, carry a digit or look like units (mW, nV/°C, kHz) keep their case."""
    return " ".join(w if _UNIT_WORD.match(w) else w.upper() for w in str(text).split(" "))


def fit_pt(text, width_in, height_in, max_pt, min_pt=14, line_spacing=1.0, char_w=0.55):
    """Largest point size (stepping down by 2) at which text roughly fits the box.
    Estimate only: average glyph width ~0.55 em for the heavy fonts used here."""
    text = str(text)
    pt = max_pt
    while pt > min_pt:
        cpl = max(1, int(width_in * 72 / (pt * char_w)))
        lines = 0
        for para in text.split("\n"):
            words, cur = para.split(), 0
            n = 1
            for w in words:
                if cur and cur + 1 + len(w) > cpl:
                    n += 1; cur = len(w)
                else:
                    cur += (1 if cur else 0) + len(w)
            lines += n
        if lines * pt * 1.2 * line_spacing / 72 <= height_in:
            return pt
        pt -= 2
    return min_pt


class Deck:
    def __init__(self, theme: dict, meta: dict | None = None):
        self.t = theme
        self.meta = meta or {}
        self.prs = open_template(theme.get("template"))
        self.W = self.prs.slide_width / EMU_PER_IN
        self.H = self.prs.slide_height / EMU_PER_IN
        self.c = theme["colors"]
        self.slide_no = 0
        self.log: list[str] = []
        m = theme.get("motion") or {}
        self.motion = m
        self.level = m.get("level", "standard")
        self.animate = self.level != "minimal"     # build_deck sets False for --static
        self._tls = {}                             # slide_id -> Timeline
        self.morph_kind = {}                       # slide_id -> morph | morph_words | morph_chars
        self.morph_ms_for = {}                     # slide_id -> duration override
        self.transition_for = {}                   # slide_id -> (kind, ms, opts) override

    # ----- motion -------------------------------------------------------------
    def tl(self, slide):
        """Animation timeline for a slide. Returns a no-op when animation is off
        (static build or level 'minimal'), so layouts can call it unconditionally."""
        if not self.animate:
            return _NullTimeline()
        from motion import Timeline
        sid = slide.slide_id
        if sid not in self._tls:
            self._tls[sid] = Timeline(slide, self.W, self.H)
        return self._tls[sid]

    def showcase(self):
        return self.animate and self.level == "showcase"

    def morph_as(self, slide, kind="morph", ms=None):
        """Ask for a particular Morph flavour on this slide (words, chars) or duration."""
        self.morph_kind[slide.slide_id] = kind
        if ms:
            self.morph_ms_for[slide.slide_id] = ms

    # ----- colour shortcuts -------------------------------------------------
    @property
    def dark(self):
        mode = self.t["mode"]
        if mode == "template":  # master paints the background; say which kind it is
            return self.t.get("template_background", "light") == "dark"
        return mode != "light"

    def fg(self, dark=None):
        d = self.dark if dark is None else dark
        return self.c["on_dark"] if d else self.c["on_light"]

    def muted(self, dark=None):
        d = self.dark if dark is None else dark
        return self.c["on_dark_muted"] if d else self.c["on_light_muted"]

    # ----- slide creation ---------------------------------------------------
    def _blank_layout(self):
        layouts = list(self.prs.slide_layouts)
        for want in self.t.get("blank_layout", []):
            for lay in layouts:
                if lay.name.strip().lower() == want.lower():
                    return lay
        return min(layouts, key=lambda l: len(l.placeholders))

    def new_slide(self, dark=None):
        s = self.prs.slides.add_slide(self._blank_layout())
        # delete any placeholders the layout brought along; we place everything ourselves
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        mode = self.t["mode"]
        if mode != "template":
            d = self.dark if dark is None else dark
            fill = s.background.fill
            fill.gradient()
            fill.gradient_angle = 90
            st = fill.gradient_stops
            if d:
                st[0].color.rgb, st[1].color.rgb = rgb(self.c["dark_bg"]), rgb(self.c["dark_bg2"])
            else:
                st[0].color.rgb, st[1].color.rgb = rgb(self.c["light_bg"]), rgb(self.c["light_bg2"])
        self.slide_no += 1
        return s

    def chrome(self, slide, dark=None, appendix=False):
        """Footer, slide number and logo. Skipped when the template master already
        provides them (mode == template); named ftr_* so lint ignores them."""
        f = self.t.get("footer") or {}
        if self.t["mode"] == "template":
            self._template_chrome(slide, f)
            return
        col = self.muted(dark)
        if f.get("text") and self.t["mode"] != "template":
            self.text(slide.shapes, 0.5, self.H - 0.42, self.W * 0.6, 0.25, f["text"],
                      size=9, color=col, font="light", name="ftr_conf")
        if f.get("slide_numbers") and self.t["mode"] != "template":
            self.text(slide.shapes, self.W - 1.0, self.H - 0.42, 0.5, 0.25, str(self.slide_no),
                      size=9, color=col, font="light", align="r", name="ftr_num")
        logo = self.t.get("logo") or {}
        if logo.get("path") and Path(logo["path"]).exists() and self.t["mode"] != "template":
            w = logo.get("width_in", 1.1)
            pic = slide.shapes.add_picture(logo["path"], IN(self.W - w - 0.5), IN(0.35), width=IN(w))
            pic.name = "ftr_logo"

    def _template_chrome(self, slide, f):
        """Template mode: python-pptx does not copy the layout's footer / slide-number
        placeholders onto a new slide, so clone them and set the classification text.
        Layouts without them (cover, section, end card) are left alone."""
        text = f.get("text")
        cls = (self.meta.get("classification") or "").strip().lower()
        text = (f.get("classifications") or {}).get(cls, text)
        for ph in slide.slide_layout.placeholders:
            kind = str(ph.placeholder_format.type)
            if "FOOTER" not in kind and "SLIDE_NUMBER" not in kind:
                continue
            slide.shapes._spTree.append(copy.deepcopy(ph._element))
            new = slide.shapes[-1]
            new.name = "ftr_" + ("conf" if "FOOTER" in kind else "num")
            if "FOOTER" in kind and text:
                new.text_frame.text = text

    # ----- primitives -------------------------------------------------------
    def text(self, shapes, x, y, w, h, txt, size=18, color=None, font="body", bold=None,
             align="l", anchor="t", name=None, caps=False, spacing=None, line_spacing=None,
             italic=False):
        tb = shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        f = self.t["fonts"][font]
        lines = str(txt).split("\n")
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            if line_spacing:
                p.line_spacing = line_spacing
            r = p.add_run()
            r.text = caps_safe(ln) if caps else ln
            r.font.size = Pt(size)
            r.font.name = f["face"]
            r.font.bold = f.get("bold", False) if bold is None else bold
            r.font.italic = italic
            r.font.color.rgb = rgb(color or self.fg())
            if spacing is not None:
                r.font._element.set("spc", str(int(spacing * 100)))
        return tb

    def rect(self, shapes, x, y, w, h, color, name=None, opacity=None, kind=MSO_SHAPE.RECTANGLE,
             line=None, line_w=1.0, dash=False):
        s = shapes.add_shape(kind, IN(x), IN(y), IN(w), IN(h))
        if name:
            s.name = name
        if color:
            s.fill.solid()
            s.fill.fore_color.rgb = rgb(color)
            if opacity is not None:
                _set_alpha(s, opacity)
        else:
            s.fill.background()
        if line:
            s.line.color.rgb = rgb(line)
            s.line.width = Pt(line_w)
            if dash:
                s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        else:
            s.line.fill.background()
        s.shadow.inherit = False
        if s.has_text_frame:
            s.text_frame.text = ""
        return s

    def line(self, shapes, x1, y1, x2, y2, color, width=1.0, name=None):
        c = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, IN(x1), IN(y1), IN(x2), IN(y2))
        c.line.color.rgb = rgb(color)
        c.line.width = Pt(width)
        if name:
            c.name = name
        return c

    def triangle_right(self, shapes, x, y, w, h, color, name=None):
        """Right-pointing triangle as a freeform: no rotation maths needed."""
        fb = shapes.build_freeform(IN(x), IN(y))
        fb.add_line_segments([(IN(x + w), IN(y + h / 2)), (IN(x), IN(y + h))], close=True)
        s = fb.convert_to_shape()
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(color)
        s.line.fill.background()
        s.shadow.inherit = False
        if name:
            s.name = name
        return s

    def picture(self, shapes, path, x, y, w, h, name=None, mode="cover", grayscale=False):
        """mode='cover' crops to fill the box; mode='contain' fits inside it."""
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        box_r, img_r = w / h, iw / ih
        if mode == "contain":
            if img_r > box_r:
                nw, nh = w, w / img_r
            else:
                nw, nh = h * img_r, h
            pic = shapes.add_picture(path, IN(x + (w - nw) / 2), IN(y + (h - nh) / 2), IN(nw), IN(nh))
        else:
            pic = shapes.add_picture(path, IN(x), IN(y), IN(w), IN(h))
            if img_r > box_r:      # image wider: crop left/right
                c = (1 - box_r / img_r) / 2
                pic.crop_left = pic.crop_right = c
            else:
                c = (1 - img_r / box_r) / 2
                pic.crop_top = pic.crop_bottom = c
        if name:
            pic.name = name
        if grayscale:
            set_grayscale(pic)
        return pic

    def placeholder_art(self, shapes, x, y, w, h, label, name=None):
        """Dashed box shown when an asset is missing, so gaps are obvious in review."""
        r = self.rect(shapes, x, y, w, h, None, line=self.muted(), dash=True, name=name or "aux_missing")
        self.text(shapes, x, y + h / 2 - 0.25, w, 0.5, f"[ {label} ]", size=12,
                  color=self.muted(), align="c", font="light", name="aux_missing_lbl")
        self.log.append(f"missing asset: {label}")
        return r

    def asset(self, path):
        if not path:
            return None
        p = Path(path)
        if not p.is_absolute() and self.meta.get("_base"):
            p = Path(self.meta["_base"]) / p
        return str(p) if p.exists() else None


# ---------------------------------------------------------------------------
# Off-canvas parking (the videos' staggered entrance) and Morph pairing
# ---------------------------------------------------------------------------

def park_from(deck: Deck, prev_slide, next_slide, direction="below", stagger_in=1.2, only=None):
    """Copy next_slide's !!-named shapes onto prev_slide, parked off-canvas at
    increasing distances. With Morph on next_slide they fly in, further ones faster.
    Shapes already on prev_slide (same name) are left alone so they morph in place."""
    have = top_level_named(prev_slide)
    targets = [(n, el) for n, el in top_level_named(next_slide).items()
               if n not in have and (only is None or any(n.startswith(o) for o in only))]
    if not targets:
        return 0
    tree = prev_slide.shapes._spTree
    max_id = max([int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))] + [1])
    W, H = deck.prs.slide_width, deck.prs.slide_height
    # sort by current position so the stagger follows reading order
    def key(item):
        off, _ = _xfrm_off_ext(item[1])
        return (int(off.get("x")), int(off.get("y"))) if direction in ("below", "above") else (int(off.get("y")), int(off.get("x")))
    targets.sort(key=key)
    for i, (nm, el) in enumerate(targets):
        clone = copy.deepcopy(el)
        # re-point image relationships
        for node in clone.iter():
            for attr in (qn("r:embed"), qn("r:link")):
                rid = node.get(attr)
                if rid:
                    part = next_slide.part.related_part(rid)
                    node.set(attr, prev_slide.part.relate_to(part, RT.IMAGE))
        for nv in clone.iter(qn("p:cNvPr")):
            max_id += 1
            nv.set("id", str(max_id))
        off, ext = _xfrm_off_ext(clone)
        gap = int(IN(0.4) + i * IN(stagger_in))
        if direction == "below":
            off.set("y", str(H + gap))
        elif direction == "above":
            off.set("y", str(-int(ext.get("cy")) - gap))
        elif direction == "left":
            off.set("x", str(-int(ext.get("cx")) - gap))
        elif direction == "right":
            off.set("x", str(W + gap))
        tree.append(clone)
    return len(targets)


def auto_morph(deck: Deck):
    """Add Morph to every slide that shares a !!-named object with the slide before."""
    slides = list(deck.prs.slides)
    count = 0
    for prev, cur in zip(slides, slides[1:]):
        if set(top_level_named(prev)) & set(top_level_named(cur)):
            add_morph(cur, deck.morph_ms_for.get(cur.slide_id, deck.t.get("morph_ms", 1500)),
                      deck.morph_kind.get(cur.slide_id, "morph"))
            count += 1
    return count


# ---------------------------------------------------------------------------
# Speaker notes
# ---------------------------------------------------------------------------

def write_notes(slide, notes):
    if not notes:
        return
    if isinstance(notes, str):
        body = notes
    else:
        parts = []
        if notes.get("pause_before"):
            parts.append("[PAUSE] Let the slide land before speaking.")
        if notes.get("say"):
            parts.append("SAY: " + notes["say"])
        if notes.get("ask"):
            parts.append("ASK THE ROOM: " + notes["ask"])
        if notes.get("pause_after"):
            parts.append("[PAUSE] Wait for reactions to finish.")
        if notes.get("conditions"):
            parts.append("CONDITIONS: " + notes["conditions"])
        if notes.get("source"):
            parts.append("SOURCE: " + notes["source"])
        if notes.get("fallback"):
            parts.append("IF IT FAILS: " + notes["fallback"])
        if notes.get("transition"):
            parts.append("NEXT: " + notes["transition"])
        body = "\n\n".join(parts)
    slide.notes_slide.notes_text_frame.text = body


# ---------------------------------------------------------------------------
# Mosaic cut-out geometry (Merge Shapes > Union, then Intersect with a photo)
# ---------------------------------------------------------------------------

def _mosaic_polys(W, H, region=(0.45, 1.0), seed_pattern=None):
    """Return a list of polygons (lists of (x,y) inches) for a pill mosaic
    rotated 45 degrees, clipped to the slide. Needs shapely; returns None if absent."""
    try:
        from shapely.geometry import LineString, box
        from shapely.ops import unary_union
        from shapely import affinity
    except ImportError:
        return None
    x0 = W * region[0]
    # (centre x fraction of region, centre y fraction of H, length, width)
    pattern = seed_pattern or [
        (0.30, 0.10, 5.0, 1.05), (0.55, 0.20, 6.5, 1.25), (0.82, 0.12, 4.2, 0.9),
        (0.40, 0.45, 5.8, 1.15), (0.70, 0.52, 7.0, 1.35), (0.95, 0.40, 5.0, 1.0),
        (0.25, 0.78, 4.0, 0.85), (0.58, 0.85, 5.5, 1.1), (0.90, 0.88, 4.5, 0.95),
        (0.12, 0.32, 2.2, 0.55), (0.98, 0.70, 2.4, 0.6), (0.45, 0.66, 0.6, 0.6),
        (0.72, 0.30, 0.55, 0.55), (0.18, 0.60, 0.5, 0.5),
    ]
    rw = W - x0
    shapes = []
    for fx, fy, L, Wd in pattern:
        cx, cy = x0 + fx * rw, fy * H
        half = max(L / 2 - Wd / 2, 0.001)
        pill = LineString([(cx - half, cy), (cx + half, cy)]).buffer(Wd / 2, quad_segs=12)
        shapes.append(affinity.rotate(pill, -45, origin=(cx, cy)))
    merged = unary_union(shapes).intersection(box(x0 - 1.0, 0, W, H))
    geoms = list(getattr(merged, "geoms", [merged]))
    return [list(g.exterior.coords) for g in geoms if g.area > 0.05]


def mosaic_shape(deck: Deck, slide, photo_path, name="!!mosaic", grayscale=True):
    polys = _mosaic_polys(deck.W, deck.H)
    if not polys:
        return None
    first = polys[0]
    fb = slide.shapes.build_freeform(IN(first[0][0]), IN(first[0][1]))
    for k, poly in enumerate(polys):
        if k > 0:
            fb.move_to(IN(poly[0][0]), IN(poly[0][1]))
        fb.add_line_segments([(IN(x), IN(y)) for x, y in poly[1:]], close=True)
    shp = fb.convert_to_shape()
    shp.name = name
    shp.line.fill.background()
    spPr = shp._element.spPr
    for tag in ("a:solidFill", "a:gradFill", "a:noFill", "a:blipFill", "a:pattFill"):
        for old in spPr.findall(qn(tag)):
            spPr.remove(old)
    geom = spPr.find(qn("a:custGeom"))
    if photo_path:
        from PIL import Image
        _, rId = slide.part.get_or_add_image_part(photo_path)
        with Image.open(photo_path) as im:
            iw, ih = im.size
        bw, bh = shp.width, shp.height
        box_r, img_r = bw / bh, iw / ih
        if img_r > box_r:
            c = int((1 - box_r / img_r) / 2 * 100000); src = f'<a:srcRect l="{c}" r="{c}"/>'
        else:
            c = int((1 - img_r / box_r) / 2 * 100000); src = f'<a:srcRect t="{c}" b="{c}"/>'
        gray = "<a:grayscl/>" if grayscale else ""
        fill = etree.fromstring(
            f'<a:blipFill xmlns:a="{NS_A}" xmlns:r="{NS_R}" rotWithShape="0">'
            f'<a:blip r:embed="{rId}">{gray}</a:blip>{src}<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
    else:
        a, b = deck.c["ramp_start"], deck.c["accent"]
        fill = etree.fromstring(
            f'<a:gradFill xmlns:a="{NS_A}" rotWithShape="0"><a:gsLst>'
            f'<a:gs pos="0"><a:srgbClr val="{a}"/></a:gs><a:gs pos="100000"><a:srgbClr val="{b}"/></a:gs>'
            f'</a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>')
    geom.addnext(fill)
    shp.shadow.inherit = False
    add_inner_shadow(shp, opacity=0.5)
    # freeforms carry a theme style; drop its text/line references visually
    return shp
