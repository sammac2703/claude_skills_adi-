"""Approximate PNG preview of a .pptx without PowerPoint (python and Pillow only), so layout can be looked at.

    python scripts/preview_png.py deck.pptx outdir [--scale 90] [--slides 3,8]

Draws backgrounds, layout panels, filled shapes, lines, text in Barlow (wrapped to its box, real sizes), tables and
charts as labelled boxes, SVG icons and the logo as outlined squares in their colour (no SVG renderer here).
It uses the same geometry as the file, so overlaps, clipping, off-slide shapes and empty areas are visible. It does not
reproduce Morph, builds, gradients or exact font metrics: it is a layout check, not a render.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from PIL import Image, ImageDraw, ImageFont          # noqa: E402
from pptx import Presentation                        # noqa: E402
from pptx.enum.shapes import MSO_SHAPE_TYPE          # noqa: E402
from qa_layout import slide_bg, solid_fill, svg_colour, text_inset, E   # noqa: E402

FONTS = Path("C:/Windows/Fonts")
_cache = {}


def font(size_pt, bold, scale):
    key = (round(size_pt * scale / 72 * 10), bold)
    if key not in _cache:
        px = max(6, int(size_pt * scale / 72 * 1.0))
        names = ["Barlow-Bold.ttf" if bold else "Barlow-Regular.ttf", "arialbd.ttf" if bold else "arial.ttf"]
        for n in names:
            try:
                _cache[key] = ImageFont.truetype(str(FONTS / n), px)
                break
            except OSError:
                continue
        else:
            _cache[key] = ImageFont.load_default()
    return _cache[key]


def hex_rgb(h, default=(255, 255, 255)):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) if h and len(h) == 6 else default


def lum(h):
    r, g, b = hex_rgb(h)
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255


def inherited_size(sh, slide):
    """Font size (pt) a placeholder inherits from its layout or master when the run sets none."""
    if not sh.is_placeholder:
        return 18
    idx = sh.placeholder_format.idx
    typ = sh.placeholder_format.type
    for src in (slide.slide_layout, slide.slide_layout.slide_master):
        for p in src.placeholders:
            if p.placeholder_format.idx == idx or (src is not slide.slide_layout and p.placeholder_format.type == typ):
                d = p._element.find(".//{*}lstStyle/{*}lvl1pPr/{*}defRPr")
                if d is not None and d.get("sz"):
                    return int(d.get("sz")) / 100
    return 18


def draw_text(d, sh, scale, default_ink, anchor_x, anchor_y, slide=None):
    tf = sh.text_frame
    li = text_inset(sh, slide) if slide is not None else 91440
    x0 = sh.left / E * scale + li / E * scale
    w = (sh.width / E - 2 * li / E) * scale
    y = sh.top / E * scale
    lines = []
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            lines.append(("", 18, False, default_ink, 0))
            continue
        sz = max([r.font.size.pt for r in runs if r.font.size] or [inherited_size(sh, slide) if slide else 18])
        bold = any(r.font.bold for r in runs)
        col = default_ink
        for r in runs:
            try:
                if r.font.color is not None and r.font.color.type is not None:
                    col = str(r.font.color.rgb)
            except Exception:
                pass
        al = p.alignment
        words, cur = p.text.split(" "), ""
        f = font(sz, bold, scale)
        for wd in words:
            t = (cur + " " + wd).strip()
            if d.textlength(t, font=f) <= w or not cur:
                cur = t
            else:
                lines.append((cur, sz, bold, col, al))
                cur = wd
        lines.append((cur, sz, bold, col, al))
    total = sum(sz * 1.2 / 72 * scale for _, sz, *_ in lines)
    va = tf.vertical_anchor
    top = sh.top / E * scale + (tf.margin_top if tf.margin_top is not None else 45720) / E * scale
    if va is not None and "MIDDLE" in str(va):
        top = sh.top / E * scale + (sh.height / E * scale - total) / 2
    elif va is not None and "BOTTOM" in str(va):
        top = (sh.top + sh.height) / E * scale - total
    y = top
    for t, sz, bold, col, al in lines:
        f = font(sz, bold, scale)
        tw = d.textlength(t, font=f)
        x = x0
        if al is not None and "CENTER" in str(al):
            x = x0 + (w - tw) / 2
        elif al is not None and "RIGHT" in str(al):
            x = x0 + w - tw
        d.text((x, y), t, font=f, fill=hex_rgb(col))
        y += sz * 1.2 / 72 * scale


def render(path, outdir, scale=90, only=None):
    prs = Presentation(path)
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    W, H = int(13.333 * scale), int(7.5 * scale)
    for n, s in enumerate(prs.slides, 1):
        if only and n not in only:
            continue
        bgc = slide_bg(s) or "7A8CA0"
        im = Image.new("RGB", (W, H), hex_rgb(bgc))
        d = ImageDraw.Draw(im)
        dark = lum(bgc) < 0.45
        ink = "FFFFFF" if dark else "00325C"
        owners = [s.slide_layout] + [s]
        for own in owners:
            for sh in own.shapes:
                if sh.width is None or sh.left is None:
                    continue
                if own is s.slide_layout and sh.is_placeholder:
                    sp = sh._element.find("{*}spPr")
                    mine = [p for p in s.placeholders if p.placeholder_format.idx == sh.placeholder_format.idx]
                    if not mine or (mine[0]._element.find("{*}spPr") is not None and mine[0]._element.find("{*}spPr").find("{*}noFill") is not None):
                        continue
                x0, y0 = sh.left / E * scale, sh.top / E * scale
                x1, y1 = x0 + sh.width / E * scale, y0 + sh.height / E * scale
                sp = sh._element.find("{*}spPr")
                fill = solid_fill(sp) if sp is not None else None
                if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or sh._element.tag.endswith("pic"):
                    if own is s.slide_layout:
                        continue
                    col = svg_colour(sh) or "888888"
                    d.rounded_rectangle([x0, y0, x1, y1], radius=6, outline=hex_rgb(col), width=2)
                    d.line([x0, y0, x1, y1], fill=hex_rgb(col), width=1)
                    continue
                if sh.shape_type == MSO_SHAPE_TYPE.LINE or sh._element.tag.endswith("cxnSp"):
                    ln = sh._element.find(".//{*}ln")
                    lc = solid_fill(ln) if ln is not None else None
                    d.line([x0, y0, x1, y1], fill=hex_rgb(lc, (150, 150, 150)), width=max(1, int(scale / 45)))
                    continue
                if getattr(sh, "has_chart", False) and sh.has_chart:
                    d.rectangle([x0, y0, x1, y1], outline=(120, 140, 160), width=2)
                    d.text((x0 + 8, y0 + 6), "chart", font=font(14, False, scale), fill=(120, 140, 160))
                    continue
                if getattr(sh, "has_table", False) and sh.has_table:
                    yy = y0
                    for r in sh.table.rows:
                        rh = r.height / E * scale
                        xx = x0
                        for ci, c in enumerate(r.cells):
                            cw = sh.table.columns[ci].width / E * scale
                            cf = c._tc.find(".//{*}solidFill")
                            col = cf[0].get("val") if cf is not None and len(cf) else "FFFFFF"
                            d.rectangle([xx, yy, xx + cw, yy + rh], fill=hex_rgb(col), outline=(200, 200, 205))
                            d.text((xx + 4, yy + 3), c.text[:int(cw / (scale * 0.1))], font=font(14, False, scale),
                                   fill=hex_rgb("FFFFFF" if col in ("0067B9",) else "00325C"))
                            xx += cw
                        yy += rh
                    continue
                prst = sh._element.find(".//{*}prstGeom")
                shape = prst.get("prst") if prst is not None else "rect"
                if fill and (own is s or not sh.is_placeholder or True):
                    if shape == "ellipse":
                        d.ellipse([x0, y0, x1, y1], fill=hex_rgb(fill))
                    elif shape == "roundRect":
                        d.rounded_rectangle([x0, y0, x1, y1], radius=8, fill=hex_rgb(fill))
                    elif shape == "rightArrow":
                        d.polygon([(x0, y0 + (y1 - y0) * .25), (x0 + (x1 - x0) * .6, y0 + (y1 - y0) * .25), (x0 + (x1 - x0) * .6, y0),
                                   (x1, (y0 + y1) / 2), (x0 + (x1 - x0) * .6, y1), (x0 + (x1 - x0) * .6, y0 + (y1 - y0) * .75),
                                   (x0, y0 + (y1 - y0) * .75)], fill=hex_rgb(fill))
                    else:
                        d.rectangle([x0, y0, x1, y1], fill=hex_rgb(fill))
                if sh.has_text_frame and sh.text_frame.text.strip() and (own is s or not sh.is_placeholder):
                    cover = fill or colour_under(s, sh, bgc)
                    draw_text(d, sh, scale, "FFFFFF" if lum(cover) < 0.45 else "00325C", 0, 0, s)
        im.save(out / f"slide{n:02d}.png")
    return len(prs.slides)


def colour_under(slide, sh, bgc):
    from qa_layout import colour_at
    try:
        return colour_at(slide, sh.left + min(sh.width // 2, 400000), sh.top + min(sh.height // 2, 200000), below=sh) or bgc
    except Exception:
        return bgc


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Approximate PNG preview of a .pptx")
    ap.add_argument("pptx")
    ap.add_argument("outdir")
    ap.add_argument("--scale", type=int, default=90)
    ap.add_argument("--slides", default="")
    a = ap.parse_args()
    only = {int(v) for v in a.slides.split(",") if v} or None
    print(f"{render(a.pptx, a.outdir, a.scale, only)} slides -> {a.outdir}")
