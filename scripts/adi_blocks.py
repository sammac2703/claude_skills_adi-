"""Building blocks for decks made from the ADI Brand Slide Library layouts and the Brand Icon Library.

    from adi_blocks import Blocks, ICON_DIR
    b = Blocks(deck)                               # deck: a clean_deck.CleanDeck opened on the Slide Library .pptx
    s = b.slide("KPI Slide (x4)", "Title")         # slide on a library layout, title filled
    b.fill(s, 26, "8", size=60)                    # fill a placeholder by idx (inherits the layout's formatting)
    b.icon(s, "PRESENTING", x, y, 0.8, "circle")   # ADI icon (SVG) recoloured, optional accent circle
    b.footer(s); b.tidy(s)                         # footer text from the layout, drop unfilled placeholders

Why: the Slide Library's layouts already carry the brand building blocks (KPI tiles, agenda, timeline, team success,
dashboard, product features, big statement). Using them, plus Brand Icon Library icons, gives variety without
hand-drawn boxes. Icons are embedded as SVG (vector, as the library's own instructions say) and recoloured by
editing the SVG stroke colour. SVG needs PowerPoint 2016 or later to display.

Icon catalogue: `assets/icons/index.json` (label -> file) is built once from `assets/Brand-Icon-Library_2025-v2.pptx`
by `scripts/extract_icons.py`.
"""
from __future__ import annotations

import copy
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lxml import etree                                           # noqa: E402
from pptx.dml.color import RGBColor                              # noqa: E402
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER           # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR                  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT           # noqa: E402
from pptx.opc.package import Part                                # noqa: E402
from pptx.opc.packuri import PackURI                             # noqa: E402
from pptx.oxml import parse_xml                                  # noqa: E402
from pptx.oxml.ns import qn                                      # noqa: E402
from pptx.util import Inches, Pt                                 # noqa: E402

SKILL = Path(__file__).resolve().parent.parent
ICON_DIR = SKILL / "assets" / "icons"
LIBRARY = SKILL / "assets" / "Brand-PPT-Slide-Library.pptx"

# ADI palette only
NAVY, BLUE, LIGHT, WHITE, GREY, CHAR = "00325C", "0067B9", "1B9CD0", "FFFFFF", "9EA1AE", "3C4157"

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"


class Blocks:
    def __init__(self, deck, footer_text: str | None = None):
        self.d = deck
        self.footer_text = footer_text or deck.footer_text
        self._n = 0
        idx = ICON_DIR / "index.json"
        self.catalogue = json.loads(idx.read_text(encoding="utf-8")) if idx.exists() else {}

    # ---- slides ------------------------------------------------------------------------------
    def layout(self, name: str):
        return next(l for l in self.d.prs.slide_layouts if l.name == name)

    def slide(self, layout_name: str, title: str | None = None, key_title: bool = True):
        """Add a slide on a Slide Library layout. The title placeholder (idx 0) is filled and, when key_title,
        named !!title so Morph glides the title between layouts."""
        s = self.d.prs.slides.add_slide(self.layout(layout_name))
        if title is not None and s.shapes.title is not None:
            s.shapes.title.text_frame.text = title
        if key_title and s.shapes.title is not None:
            s.shapes.title.name = "!!title"
        return s

    @staticmethod
    def ph(slide, idx):
        return next((p for p in slide.placeholders if p.placeholder_format.idx == idx), None)

    def fill(self, slide, idx, lines, size=None, bold=None, color=None, anchor=None, align=None, top_inset=None,
             name=None):
        """Fill a placeholder by idx. lines: str or list (one paragraph each). Inherits the layout's font, size and
        colour unless overridden. Returns the shape (use as a build item)."""
        sh = self.ph(slide, idx)
        if sh is None:
            raise KeyError(f"no placeholder idx {idx} on layout {slide.slide_layout.name}")
        tf = sh.text_frame
        lines = lines if isinstance(lines, list) else [lines]
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            if i == 0:
                for r in list(p.runs):
                    r._r.getparent().remove(r._r)
            spec = ln if isinstance(ln, dict) else {"t": ln}
            r = p.add_run()
            r.text = spec["t"]
            sz = spec.get("size", size)
            if sz:
                r.font.size = Pt(sz)
            b = spec.get("bold", bold)
            if b is not None:
                r.font.bold = b
            c = spec.get("color", color)
            if c:
                r.font.color.rgb = RGBColor.from_string(c)
            if align is not None:
                p.alignment = align
        if anchor is not None:
            tf.vertical_anchor = anchor
        if top_inset is not None:
            tf.margin_top = Inches(top_inset)
        if name:
            sh.name = name
        return sh

    def tidy(self, slide):
        """Remove placeholders left empty (they would show prompt text in edit view)."""
        for p in list(slide.placeholders):
            if p.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                continue
            if p.has_text_frame and p.text_frame.text.strip():
                continue
            p._element.getparent().remove(p._element)

    def footer(self, slide):
        """Clone the layout's footer placeholder onto the slide and set the classification text."""
        for p in slide.slide_layout.placeholders:
            if p.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                el = copy.deepcopy(p._element)
                slide.shapes._spTree.append(el)
                ids = [int(i) for i in slide.shapes._spTree.xpath(".//p:cNvPr/@id")]
                el.xpath(".//p:cNvPr")[0].set("id", str(max(ids) + 1))
                shp = slide.shapes[-1]
                shp.text_frame.text = self.footer_text
                return shp

    # ---- text primitives (explicit RGB, ADI palette) -----------------------------------------------
    def text(self, slide, x, y, w, h, s, size, color=NAVY, bold=False, anchor=MSO_ANCHOR.TOP, name=None,
             align=PP_ALIGN.LEFT, underline=False):
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
            r.font.underline = underline
            r.font.color.rgb = RGBColor.from_string(color)
        if name:
            t.name = name
        return t

    def line(self, slide, x1, y1, x2, y2, color=WHITE, width=1.0, alpha=None, name="Line"):
        from pptx.enum.shapes import MSO_CONNECTOR
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.rgb = RGBColor.from_string(color)
        c.line.width = Pt(width)
        if alpha is not None:
            clr = c.line._get_or_add_ln().find(qn("a:solidFill"))[0]
            etree.SubElement(clr, qn("a:alpha")).set("val", str(int(alpha * 100000)))
        c.name = name
        return c

    def dot(self, slide, x, y, d, color=BLUE, name="Dot"):
        o = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        o.fill.solid()
        o.fill.fore_color.rgb = RGBColor.from_string(color)
        o.line.fill.background()
        o.shadow.inherit = False
        o.name = name
        return o

    # ---- icons ---------------------------------------------------------------------------------
    def find_icon(self, label: str):
        """Match an icon by its Brand Icon Library label (case-insensitive, exact first, then substring)."""
        want = re.sub(r"\s+", " ", label.upper()).strip()
        hits = [(k, v) for k, v in self.catalogue.items() if v.upper() == want]
        hits = hits or [(k, v) for k, v in self.catalogue.items() if want in v.upper()]
        if not hits:
            raise KeyError(f"icon '{label}' not in catalogue ({len(self.catalogue)} icons)")
        return ICON_DIR / f"{hits[0][0]}.svg", hits[0][1]

    def icon(self, slide, label, x, y, d=0.9, style="circle", color=WHITE, circle=LIGHT, key=None):
        """ADI icon as SVG recoloured to `color`. style: 'circle' = icon on a filled circle (default circle colour is
        ADI light blue, as the keynote look), 'bare' = just the line icon. Returns [shapes] for a build item."""
        path, real = self.find_icon(label)
        if style == "circle":
            color = WHITE                                  # rule: glyphs on a blue or dark circle are always white
        raw = path.read_bytes()
        # Library icons are single-colour line art but not all use #000000 (some use 2B2B2B or 231F20): recolour every hex.
        svg = re.sub(rb"#[0-9A-Fa-f]{6}(?![0-9A-Fa-f])", ("#" + color).encode(), raw)
        svg = self._plain_svg(svg, color)
        shapes = []
        base = f"!!{key}" if key else f"Icon {label.title()}"
        if style == "circle":
            c = self.dot(slide, x, y, d, circle, name=base)
            shapes.append(c)
            ix, iy, isz = x + d * 0.2, y + d * 0.2, d * 0.6
        else:
            ix, iy, isz = x, y, d
        shapes.append(self._svg_pic(slide, svg, ix, iy, isz, base + "_art" if style == "circle" else base, real.title()))
        return shapes

    @staticmethod
    def _plain_svg(svg: bytes, color: str) -> bytes:
        """Strip Office theme classes (MsftOfcThm_*) and the style block. PowerPoint maps those classes to the theme's
        Text 1 colour (black) and ignores the colour written in the file, so an icon recoloured only by hex stays
        black on screen. Stroke and fill are then set as plain attributes."""
        s = re.sub(rb"<style[^>]*>.*?</style>", b"", svg, flags=re.S)
        def tag(m):
            t_ = m.group(0)
            cls = re.search(rb'class="(MsftOfcThm_[^"]*)"', t_)
            if not cls:
                return t_
            t_ = t_.replace(cls.group(0), b"")
            if b"Fill" in cls.group(1) and not re.search(rb'\sfill="(?!none)', t_):
                t_ = re.sub(rb'\sfill="[^"]*"', b"", t_)
                t_ = t_.replace(b"<", b"<", 1)
                t_ = re.sub(rb"^(<\w+)", rb' fill="#' + color.encode() + b'"', t_)
            elif b"Stroke" in cls.group(1) and not re.search(rb'\sstroke="', t_):
                t_ = re.sub(rb"^(<\w+)", rb' stroke="#' + color.encode() + b'"', t_)
            return t_
        return re.sub(rb"<[a-zA-Z][^>]*>", tag, s)

    def _svg_pic(self, slide, svg: bytes, x, y, size, name, alt, h=None):
        self._n += 1
        pkg = slide.part.package
        part = Part(PackURI(f"/ppt/media/adi_icon_{id(slide) % 10**6}_{self._n}.svg"), "image/svg+xml", pkg, svg)
        rid = slide.part.relate_to(part, RT.IMAGE)
        ids = [int(i) for i in slide.shapes._spTree.xpath(".//p:cNvPr/@id")] or [1]
        sid = max(ids) + 1
        emu = lambda v: int(v * 914400)  # noqa: E731
        xml = (f'<p:pic xmlns:p="{NS_P}" xmlns:a="{NS_A}" xmlns:r="{NS_R}"><p:nvPicPr>'
               f'<p:cNvPr id="{sid}" name="{name}" descr="{alt}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>'
               f'</p:nvPicPr><p:blipFill><a:blip><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">'
               f'<asvg:svgBlip xmlns:asvg="{NS_SVG}" r:embed="{rid}"/></a:ext></a:extLst></a:blip>'
               f'<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/>'
               f'<a:ext cx="{emu(size)}" cy="{emu(h if h else size)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
               f'</p:spPr></p:pic>')
        el = parse_xml(xml)
        slide.shapes._spTree.append(el)
        return slide.shapes[-1]

    # ---- branding: logo, slide number, open (box-free) placeholders ----------------------------------
    def logo(self, slide, dark, x=11.43, y=0.2, w=1.45):
        """ADI logo, top right above the title (layouts above content slides carry none). The library's own white SVG
        is recoloured to ADI blue 1668B3 on light slides. Aspect 432 x 122.63."""
        import zipfile
        if not hasattr(self, "_logo"):
            with zipfile.ZipFile(LIBRARY) as z:
                self._logo = z.read("ppt/media/image3.svg")
        svg = self._logo if dark else self._logo.replace(b"#FFFFFF", b"#1668B3")
        return self._svg_pic(slide, svg, x, y, w, "ADI Logo", "Analog Devices logo", h=w * 122.63 / 432)

    def slide_number(self, slide, dark, x=12.0, y=7.17, w=0.88):
        """Slide number field (library layouts carry no number placeholder), 10 pt, grey on light, white on dark."""
        t = self.text(slide, x, y, w, 0.2, "", 10, WHITE if dark else GREY, name="Slide Number", align=PP_ALIGN.RIGHT)
        p = t.text_frame.paragraphs[0]._p
        for r in p.findall(qn("a:r")):
            p.remove(r)
        colour = WHITE if dark else GREY
        fld = parse_xml(f'<a:fld xmlns:a="{NS_A}" id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
                        f'<a:rPr lang="en-GB" sz="1000"><a:solidFill><a:srgbClr val="{colour}"/></a:solidFill></a:rPr>'
                        f'<a:t>1</a:t></a:fld>')
        end = p.find(qn("a:endParaRPr"))
        (end.addprevious(fld) if end is not None else p.append(fld))
        return t

    @staticmethod
    def is_dark(slide):
        """Dark or light from assets/layout_facts.json (built by extract_layout_facts.py); False when unknown."""
        f = SKILL / "assets" / "layout_facts.json"
        try:
            return bool(json.loads(f.read_text(encoding="utf-8"))[slide.slide_layout.name]["dark"])
        except (OSError, KeyError, ValueError):
            return False

    def brand(self, slide, dark=None):
        """Logo plus slide number on a content slide (dark=None reads the layout's background from layout_facts.json)."""
        dark = self.is_dark(slide) if dark is None else dark
        self.logo(slide, dark)
        self.slide_number(slide, dark)

    @staticmethod
    def open_up(shape):
        """Drop a library placeholder's box (fill and outline) so only its text remains."""
        shape.fill.background()
        shape.line.fill.background()
        return shape
