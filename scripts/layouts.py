"""
Layout library. Each function takes (deck, spec) and returns the list of slides
it created (usually one; build sequences return several).

Naming convention, which Morph and lint both depend on:
  !!name   object that Morph should pair across slides
  aux_*    supporting text (axis labels, numbering) that lint does not word-count
  photo_*  photographs (lint expects them greyscale or overlaid)
  ftr_*    footer, slide number, logo

Every layout reads well as a static slide too, because Morph does not survive
PDF export.
"""
from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE

from deckkit import (Deck, IN, fit_pt, add_outer_shadow, lerp_hex, luminance, mosaic_shape,
                     set_picture_opacity)
from motion import set_3d_rotation

M = 0.8  # outer margin, inches


def _kicker(d: Deck, sh, txt, y=0.75, color=None, x=M, w=None):
    if txt:
        return d.text(sh, x, y, w or d.W - 2 * M, 0.4, txt, size=14, font="head",
                      color=color or d.c["accent"], caps=True, spacing=2, name="aux_kicker")
    return None


def _intro(d: Deck, slide, steps):
    """Staggered entrance as the slide opens: [(shape, effect, delay_ms), ...].
    All 'With Previous' with rising delays, the pro way to cascade a title block."""
    t = d.tl(slide)
    for shp, eff, delay in steps:
        if shp is not None:
            t.add(shp, eff, start="with", delay=delay)
    return t


# ---------------------------------------------------------------- hook / titles

def title(d: Deck, s):
    """Heavy-type title: huge heavy title, accent bold subhead, light body."""
    sl = d.new_slide(); sh = sl.shapes
    wm = d.asset(s.get("watermark"))
    if wm:
        pic = d.picture(sh, wm, d.W - 4.6, 0.6, 4.0, 4.0, mode="contain", name="aux_watermark")
        set_picture_opacity(pic, 0.25)  # ~75% transparent watermark
    k = _kicker(d, sh, s.get("kicker"), y=2.35)
    t = d.text(sh, M, 2.85, d.W - 2 * M, 2.2, s["title"], size=s.get("size", 60), font="heavy",
               line_spacing=0.92, name="!!title")
    sub = None
    if s.get("subtitle"):
        sub = d.text(sh, M, 5.15, d.W * 0.62, 1.0, s["subtitle"], size=22, font="light", color=d.muted(),
                     name="sub_title")
    _intro(d, sl, [(k, "fade", 200), (t, "float_up", 450), (sub, "fade", 1000)])
    d.chrome(sl)
    return [sl]


def mosaic(d: Deck, s):
    """Mosaic cut-out title (video 2): union of 45-degree pills filled with an
    upright, greyscale application photo, inner shadow, dark gradient."""
    sl = d.new_slide(dark=True); sh = sl.shapes
    photo = d.asset(s.get("photo"))
    shp = mosaic_shape(d, sl, photo)
    if shp is None:  # no shapely: fall back to plain pills
        for i, (x, y) in enumerate([(8.2, 0.3), (9.6, 1.9), (7.6, 3.6), (10.4, 4.5)]):
            r = d.rect(sh, x, y, 3.6, 1.0, d.c["accent"], opacity=0.85 - i * 0.15,
                       kind=MSO_SHAPE.ROUNDED_RECTANGLE, name=f"aux_pill{i}")
            r.adjustments[0] = 0.5
            r.rotation = -45
    if not photo:
        d.log.append("mosaic: no photo given, used gradient fill")
    k = _kicker(d, sh, s.get("kicker"), y=2.4, w=6.2)
    t = d.text(sh, M, 2.9, 6.4, 2.4, s["title"], size=s.get("size", 54), font="heavy",
               color=d.c["on_dark"], line_spacing=0.92, name="!!title")
    sub = None
    if s.get("subtitle"):
        sub = d.text(sh, M, 5.2, 6.0, 1.0, s["subtitle"], size=20, font="light", color=d.c["on_dark_muted"],
                     name="sub_title")
    tl = d.tl(sl)
    if shp is not None:  # slow push-in on the photo while the title lands (Ken Burns)
        tl.add(shp, "ken_burns", start="with", duration=12000, scale=1.06)
    for shp_, eff, delay in [(k, "fade", 300), (t, "float_up", 600), (sub, "fade", 1200)]:
        if shp_ is not None:
            tl.add(shp_, eff, start="with", delay=delay)
    d.chrome(sl, dark=True)
    return [sl]


def section(d: Deck, s):
    sl = d.new_slide(); sh = sl.shapes
    num = sub = None
    if s.get("number"):
        num = d.text(sh, M, 1.6, 4, 2.2, s["number"], size=140, font="heavy", color=d.c["accent"],
                     name="aux_section_no")
    t = d.text(sh, M, 4.1, d.W - 2 * M, 1.4, s["title"], size=48, font="heavy", name="!!section_title")
    if s.get("subtitle"):
        sub = d.text(sh, M, 5.35, d.W * 0.6, 0.8, s["subtitle"], size=20, font="light", color=d.muted())
    if s.get("role") != "appendix":
        _intro(d, sl, [(num, "fade", 0), (t, "float_up", 250), (sub, "fade", 700)])
    d.chrome(sl)
    return [sl]


def statement(d: Deck, s):
    """One big sentence: the problem, the hook line, a candour moment.
    build="lines": each line arrives on its own click (problem stated beat by beat).
    then=[...]: follow-on versions of the sentence; Morph by Words moves the words that
    stay and swaps the ones that change ("Three chips. 41 mW." -> "One chip. 8 mW.")."""
    texts = [s["text"]] + list(s.get("then") or [])
    slides = []
    for i, txt in enumerate(texts):
        sl = d.new_slide(); sh = sl.shapes
        _kicker(d, sh, s.get("kicker"), y=2.2)
        st = d.text(sh, M, 2.7, d.W - 2.4 * M, 3.0, txt, size=s.get("size", 48), font="heavy",
                    line_spacing=0.95, name="!!statement",
                    color=d.c["accent"] if (i and s.get("then_accent", True)) else None)
        t = d.tl(sl)
        if i == 0 and s.get("build") == "lines" and "\n" in txt:
            t.add(st, "fade", start="with", by="paragraph", label="next line of the problem")
        if i:
            d.morph_as(sl, "morph_words", s.get("morph_ms"))
        if s.get("support") and i == len(texts) - 1:
            sup = d.text(sh, M, 5.6, d.W * 0.6, 0.8, s["support"], size=20, font="light", color=d.muted(),
                         name="support")
            t.add(sup, "fade", start="click", label="support line")
        d.chrome(sl)
        slides.append(sl)
    return slides


# ---------------------------------------------------------------- credibility

def credibility(d: Deck, s):
    """One reference win: greyscale photo + overlay on the left, number + line right."""
    sl = d.new_slide(); sh = sl.shapes
    pw = d.W * 0.46
    photo = d.asset(s.get("photo"))
    pic = None
    if photo:
        pic = d.picture(sh, photo, 0, 0, pw, d.H, grayscale=True, name="photo_credibility")
        d.rect(sh, 0, 0, pw, d.H, d.c["ramp_start"], opacity=d.t["photo_overlay_opacity"], name="aux_overlay")
    else:
        d.placeholder_art(sh, 0.4, 0.4, pw - 0.8, d.H - 0.8, "application photo")
    x = pw + 0.7; w = d.W - x - M
    k = _kicker(d, sh, s.get("kicker", "Done before"), y=1.6, x=x, w=w)
    num = d.text(sh, x, 2.1, w, 1.9, s["number"], size=110, font="heavy", color=d.c["accent"], name="!!cred_number")
    unit = None
    if s.get("unit"):
        unit = d.text(sh, x, 3.95, w, 0.6, s["unit"], size=24, font="head", name="cred_unit")
    line = d.text(sh, x, 4.7, w, 1.4, s["line"], size=24, font="light", color=d.muted(), name="cred_line")
    tl = d.tl(sl)
    if pic is not None and d.showcase():
        tl.add(pic, "ken_burns", start="with", duration=10000, scale=1.06)
    for shp_, eff, delay in [(k, "fade", 200), (num, "float_up", 400), (unit, "fade", 900), (line, "fade", 1100)]:
        if shp_ is not None:
            tl.add(shp_, eff, start="with", delay=delay)
    d.chrome(sl)
    return [sl]


def objections(d: Deck, s):
    """Name the concern before they do: concern (muted) -> one-line answer."""
    sl = d.new_slide(); sh = sl.shapes
    _kicker(d, sh, s.get("kicker", "What you might be thinking"), y=0.8)
    items = s["items"][:3]
    rowh = (d.H - 2.4) / len(items)
    for i, it in enumerate(items):
        y = 1.55 + i * rowh
        g = sh.add_group_shape(); g.name = f"!!obj{i+1}"
        d.text(g.shapes, M, y, 5.4, rowh - 0.2, f"“{it['concern']}”", size=22, font="light",
               italic=True, color=d.muted(), anchor="m")
        a = sh.add_group_shape(); a.name = f"obj{i+1}_answer"
        d.rect(a.shapes, 6.5, y + rowh / 2 - 0.02, 0.6, 0.04, d.c["accent"])
        d.text(a.shapes, 7.4, y, d.W - 7.4 - M, rowh - 0.2, it["answer"], size=24, font="head", anchor="m")
        # concern is on screen; the answer wipes in on click, so you voice the doubt first
        d.tl(sl).add(a, "wipe", start="click", direction="left", duration=500,
                     label=f"answer {i+1}: {it['answer'][:30]}")
    d.chrome(sl)
    return [sl]


# ---------------------------------------------------------------- problem

def status_quo(d: Deck, s):
    """Cost of the status quo, building one row per slide (iPod cost-per-song table).
    Rows not yet shown are parked off the left edge at staggered distances."""
    rows = s["rows"]
    total = s.get("total")
    ours = s.get("ours")
    build = s.get("build_rows", True)
    if build:
        steps = list(range(1, len(rows) + 1))
        if total:
            steps.append("total")
        if ours:
            steps.append("ours")
    else:  # no build: one finished slide
        steps = ["ours" if ours else ("total" if total else len(rows))]
    slides = []
    rowh = min(0.85, (d.H - 3.2) / (len(rows) + 1))
    colx = M; colw = 6.8
    for step in steps:
        sl = d.new_slide(); sh = sl.shapes
        _kicker(d, sh, s.get("kicker", "What it costs today"), y=0.8)
        if s.get("title"):
            d.text(sh, M, 1.2, d.W - 2 * M, 0.8, s["title"], size=32, font="heavy", name="!!sq_title")
        shown = len(rows) if step in ("total", "ours") else step
        for i, r in enumerate(rows):
            y = 2.25 + i * rowh
            x = colx if i < shown else -colw - 0.5 - (i - shown) * 1.5
            g = sh.add_group_shape(); g.name = f"!!row{i+1}"
            d.text(g.shapes, x, y, colw - 2.2, rowh, r["item"], size=22, font="light", anchor="m")
            d.text(g.shapes, x + colw - 2.2, y, 2.2, rowh, r["value"], size=fit_pt(r["value"], 2.2, rowh, 24, 14), font="head", align="r", anchor="m")
            d.rect(g.shapes, x, y + rowh - 0.02, colw, 0.015, d.muted())
        if total:
            y = 2.35 + len(rows) * rowh
            vis = step in ("total", "ours")
            x = colx if vis else -colw - 3
            g = sh.add_group_shape(); g.name = "!!total"
            d.text(g.shapes, x, y, colw - 2.6, rowh + 0.2, total["label"], size=26, font="heavy", anchor="m")
            d.text(g.shapes, x + colw - 2.6, y, 2.6, rowh + 0.2, total["value"], size=fit_pt(total["value"], 2.6, rowh + 0.2, 34, 16), font="heavy",
                   color=d.c["warn"], align="r", anchor="m")
        if ours:
            vis = step == "ours"
            x = 8.4 if vis else d.W + 0.6
            g = sh.add_group_shape(); g.name = "!!ours_price"
            d.rect(g.shapes, x, 2.25, 4.1, 3.6, d.c["accent"], opacity=0.14,
                   kind=MSO_SHAPE.ROUNDED_RECTANGLE)
            d.text(g.shapes, x + 0.35, 2.55, 3.4, 0.5, ours["label"], size=fit_pt(ours["label"], 3.4, 0.55, 18, 12), font="head", color=d.c["accent"], caps=True)
            d.text(g.shapes, x + 0.35, 3.15, 3.4, 1.4, ours["value"], size=fit_pt(ours["value"], 3.4, 1.4, 64, 24), font="heavy", anchor="m")
            if ours.get("note"):
                d.text(g.shapes, x + 0.35, 4.65, 3.4, 1.0, ours["note"], size=fit_pt(ours["note"], 3.4, 1.0, 16, 11), font="light", color=d.muted())
        d.chrome(sl)
        slides.append(sl)
    return slides


# ---------------------------------------------------------------- criteria

def criteria(d: Deck, s):
    """Two or three yardsticks the customer agrees matter, one word each,
    set before any product appears. Equal columns."""
    sl = d.new_slide(); sh = sl.shapes
    _kicker(d, sh, s.get("kicker", "What has to be true"), y=1.2)
    items = s["items"][:4]
    n = len(items); colw = (d.W - 2 * M) / n
    for i, it in enumerate(items):
        word, sub = (it, None) if isinstance(it, str) else (it["word"], it.get("sub"))
        x = M + i * colw
        g = sh.add_group_shape(); g.name = f"!!crit{i+1}"
        d.text(g.shapes, x, 2.4, colw - 0.3, 0.6, f"0{i+1}", size=18, font="head", color=d.c["accent"])
        d.text(g.shapes, x, 3.0, colw - 0.3, 1.4, word, size=54 if n < 4 else 44, font="heavy")
        if sub:
            d.text(g.shapes, x, 4.4, colw - 0.4, 1.2, sub, size=18, font="light", color=d.muted())
        # one yardstick per click: get a nod on each before the next
        d.tl(sl).add(g, "float_up", start="click", label=f"criterion: {word}")
    d.chrome(sl)
    return [sl]


# ---------------------------------------------------------------- positioning

def quadrant(d: Deck, s):
    """2x2: rivals plotted, target corner empty; then ours moves in (Morph).
    Coordinates are 0..1 on each axis; (1,1) is top-right."""
    slides = []
    states = ["empty", "filled"] if s.get("reveal", True) and s.get("ours") else ["filled"]
    px, py, pw, ph = 3.4, 1.15, 6.6, 5.4  # plot box
    for st in states:
        sl = d.new_slide(); sh = sl.shapes
        tc = s.get("target_corner", "tr")
        tx = px + (pw / 2 if "r" in tc else 0); ty = py + (0 if "t" in tc else ph / 2)
        d.rect(sh, tx + 0.08, ty + 0.08, pw / 2 - 0.16, ph / 2 - 0.16, d.c["accent"], opacity=0.08,
               line=d.c["accent"], dash=True, name="!!target")
        if s.get("target_label"):
            d.text(sh, tx + 0.25, ty + 0.2, pw / 2 - 0.5, 0.4, s["target_label"], size=12, font="head",
                   color=d.c["accent"], caps=True, name="aux_target_lbl")
        d.line(sh, px, py + ph / 2, px + pw, py + ph / 2, d.muted(), 1.0, name="aux_axis_x")
        d.line(sh, px + pw / 2, py, px + pw / 2, py + ph, d.muted(), 1.0, name="aux_axis_y")
        d.text(sh, px + pw + 0.15, py + ph / 2 - 0.2, 2.6, 0.4, s["x_axis"], size=16, font="head", name="aux_xlab")
        d.text(sh, px + pw / 2 - 1.5, py - 0.55, 3.0, 0.4, s["y_axis"], size=16, font="head", align="c", name="aux_ylab")
        if s.get("x_low"):
            d.text(sh, px - 2.7, py + ph / 2 - 0.2, 2.55, 0.4, s["x_low"], size=13, font="light", color=d.muted(), align="r", name="aux_xlow")
        if s.get("y_low"):
            d.text(sh, px + pw / 2 - 1.5, py + ph + 0.08, 3.0, 0.4, s["y_low"], size=13, font="light", color=d.muted(), align="c", name="aux_ylow")
        def dot(name, x, y, label, col, big=False):
            r = 0.34 if big else 0.24
            cx, cy = px + x * pw, py + (1 - y) * ph
            g = sh.add_group_shape(); g.name = name
            d.rect(g.shapes, cx - r / 2, cy - r / 2, r, r, col, kind=MSO_SHAPE.OVAL)
            left = x < 0.5 and cx + r / 2 + 0.1 + 2.6 > px + pw / 2
            lx = cx - r / 2 - 0.1 - 2.6 if left else cx + r / 2 + 0.1
            d.text(g.shapes, lx, cy - 0.2, 2.6, 0.4, label, size=18 if big else 14,
                   font="head" if big else "light", color=d.fg() if big else d.muted(),
                   align="r" if left else "l", name="aux_dotlbl")
            return g
        rivals = []
        for i, rv in enumerate(s.get("rivals", [])):
            rivals.append((dot(f"!!rival{i+1}", rv["x"], rv["y"], rv["name"], d.c["rival"]), rv["name"]))
        o = s.get("ours")
        g = None
        if o:
            if st == "filled":
                g = dot("!!ours", o["x"], o["y"], o["name"], d.c["ours"], big=True)
            else:
                g = dot("!!ours", 2.2, o["y"], o["name"], d.c["ours"], big=True)  # parked off the right edge
        tl = d.tl(sl)
        if st == states[0]:
            # draw the axes as the slide opens, then place each alternative on click
            names = {x.name: x for x in sl.shapes}
            tl.add(names["aux_axis_x"], "wipe", start="with", direction="left", duration=600)
            tl.add(names["aux_axis_y"], "wipe", start="with", direction="bottom", duration=600)
            for nm in ("aux_xlab", "aux_ylab", "aux_xlow", "aux_ylow"):
                if nm in names:
                    tl.add(names[nm], "fade", start="with", delay=400)
            for grp, nm in rivals:
                tl.add(grp, "zoom", start="click", label=f"plot {nm}")
            tl.add(names["!!target"], "fade", start="click", label="the empty target corner")
            if "aux_target_lbl" in names:
                tl.add(names["aux_target_lbl"], "fade", start="with")
        elif g is not None:
            tl.add(g, "pulse", start="after", delay=100)
        d.chrome(sl)
        slides.append(sl)
    return slides


def bars(d: Deck, s):
    """Two-bar metric (performance per watt: 15 vs 70). With grow=true the first
    slide shows bars at zero height and Morph grows them."""
    vals = [b["value"] for b in s["bars"]]
    vmax = max(vals) or 1
    base_y, top_y = 6.2, 1.9
    n = len(s["bars"]); bw = 1.9; gap = 1.4
    x0 = (d.W - (n * bw + (n - 1) * gap)) / 2 + 1.6
    slides = []
    for st in (["zero", "full"] if s.get("grow") else ["full"]):
        sl = d.new_slide(); sh = sl.shapes
        _kicker(d, sh, s.get("kicker"), y=0.8)
        d.text(sh, M, 1.2, 4.0, 1.6, s["metric"], size=36, font="heavy", name="!!metric")
        if s.get("unit"):
            d.text(sh, M, 2.75, 4.4, 0.9, s["unit"], size=18, font="light", color=d.muted(), name="aux_unit")
        d.line(sh, x0 - 0.4, base_y, x0 + n * bw + (n - 1) * gap + 0.4, base_y, d.muted(), 1.0, name="aux_base")
        tl = d.tl(sl)
        for i, b in enumerate(s["bars"]):
            hgt = (base_y - top_y) * (b["value"] / vmax) if st == "full" else 0.04
            x = x0 + i * (bw + gap)
            col = d.c["ours"] if b.get("ours") else d.c["rival"]
            bar = d.rect(sh, x, base_y - hgt, bw, hgt, col, name=f"!!bar{i+1}")
            val = d.text(sh, x - 0.7, base_y - hgt - 1.0, bw + 1.4, 0.95, b.get("display", f"{b['value']:g}"),
                         size=44, font="heavy", align="c", anchor="b", name=f"!!val{i+1}")
            d.text(sh, x - 0.5, base_y + 0.12, bw + 1.0, 0.5, b["label"], size=16, font="head",
                   align="c", color=d.fg() if b.get("ours") else d.muted(), name=f"aux_barlbl{i+1}")
            if not s.get("grow"):
                # no Morph pair: build each bar from the baseline on click
                tl.add(bar, "wipe", start="click", direction="bottom", duration=700, label=f"bar: {b['label']}")
                tl.add(val, "fade", start="after", duration=300)
            elif st == "full" and b.get("ours"):
                tl.add(val, "pulse", start="after", delay=150)
        d.chrome(sl)
        slides.append(sl)
    return slides


# ---------------------------------------------------------------- tease / reveal

def tease(d: Deck, s):
    """Three key benefits (video 2): circles with glyph or icon, accent subhead,
    light body. Circles are !!b1..!!b3 so the reveal slide can merge them."""
    sl = d.new_slide(); sh = sl.shapes
    _kicker(d, sh, s.get("kicker"), y=0.8)
    items = s["items"][:3]
    n = len(items); colw = (d.W - 2 * M) / n
    for i, it in enumerate(items):
        cx = M + i * colw + colw / 2
        dia = 1.7
        top = 1.9
        c = d.rect(sh, cx - dia / 2, top, dia, dia, d.c["on_dark"] if d.dark else d.c["on_light"],
                   kind=MSO_SHAPE.OVAL, name=f"!!b{i+1}")
        icon = d.asset(it.get("icon"))
        g = sh.add_group_shape(); g.name = f"!!b{i+1}_txt"
        if icon:
            d.picture(g.shapes, icon, cx - 0.5, top + dia / 2 - 0.5, 1.0, 1.0, mode="contain", name="aux_icon")
        else:
            d.text(g.shapes, cx - dia / 2, top, dia, dia, it.get("glyph", str(i + 1)), size=48,
                   font="heavy", align="c", anchor="m", color=d.c["dark_bg"] if d.dark else d.c["on_dark"])
        d.text(g.shapes, cx - colw / 2 + 0.2, 4.0, colw - 0.4, 0.9, it["title"], size=26, font="head",
               align="c", color=d.c["accent"])
        if it.get("body"):
            d.text(g.shapes, cx - colw / 2 + 0.3, 4.85, colw - 0.6, 1.4, it["body"], size=17, font="light",
                   align="c", color=d.muted())
    d.chrome(sl)
    return [sl]


def reveal(d: Deck, s):
    """Reveal + one-sentence definition. If the previous slide was a tease, its
    circles (!!b1..!!b3) land here as small dots under the name: three become one."""
    sl = d.new_slide(); sh = sl.shapes
    img = d.asset(s.get("image"))
    if img:
        pic = d.picture(sh, img, 0.9, 1.0, 5.2, 5.2, mode="contain", name="!!product")
        add_outer_shadow(pic, blur_pt=16, dist_pt=6, opacity=0.45)
    elif s.get("glyph"):
        gsize = s.get("glyph_size", max(60, min(170, int(520 / max(len(s["glyph"]), 1)))))
        d.text(sh, 0.9, 1.0, 5.2, 5.2, s["glyph"], size=gsize, font="heavy", color=d.c["accent"],
               align="c", anchor="m", name="!!product")
    else:
        d.placeholder_art(sh, 0.9, 1.0, 5.2, 5.2, "product render (transparent PNG)", name="!!product")
    x = 6.7; w = d.W - x - M
    _kicker(d, sh, s.get("kicker"), y=1.7, x=x, w=w)
    d.text(sh, x, 2.15, w, 1.4, s["name"], size=s.get("size", 64), font="heavy", name="!!prod_name")
    if s.get("merge_dots", True):
        for i in range(s.get("dots", 3)):
            d.rect(sh, x + i * 0.45, 3.65, 0.28, 0.28, d.c["accent"], kind=MSO_SHAPE.OVAL, name=f"!!b{i+1}")
    df = d.text(sh, x, 4.2, w, 1.8, s["definition"], size=26, font="light", color=d.muted(), name="definition")
    # name lands with the Morph; the one-line definition follows a beat later
    d.tl(sl).add(df, "fade", start="after", delay=400, duration=600)
    d.chrome(sl)
    return [sl]


# ---------------------------------------------------------------- how it works

def feature(d: Deck, s):
    """One feature, one benefit line, one image. The problem it answers lives in the notes."""
    sl = d.new_slide(); sh = sl.shapes
    img = d.asset(s.get("image"))
    x_img = d.W * 0.52
    glyph = None
    if img:
        d.picture(sh, img, x_img, 0.9, d.W - x_img - 0.6, d.H - 1.8, mode="contain", name="!!product" if s.get("is_product") else "feature_img")
    elif s.get("glyph"):
        gsize = s.get("glyph_size", max(60, min(170, int(560 / max(len(s["glyph"]), 1)))))
        glyph = d.text(sh, x_img - 0.3, 1.2, d.W - x_img - M + 0.6, 4.5, s["glyph"], size=gsize, font="heavy",
                       color=d.c["accent"], align="c", anchor="m", name="aux_glyph")
    else:
        d.placeholder_art(sh, x_img, 0.9, d.W - x_img - 0.6, d.H - 1.8, "feature image or plot")
    _kicker(d, sh, s.get("feature"), y=2.0, w=x_img - M - 0.4)
    d.text(sh, M, 2.5, x_img - M - 0.5, 2.6, s["benefit"], size=s.get("size", 44), font="heavy",
           line_spacing=0.95, name="benefit")
    tl = d.tl(sl)
    if glyph is not None:
        tl.add(glyph, "zoom", start="with", delay=250, duration=600)
    if s.get("support"):
        sup = d.text(sh, M, 5.2, x_img - M - 0.6, 1.0, s["support"], size=18, font="light", color=d.muted(), name="support")
        tl.add(sup, "fade", start="click", label="support line")
    d.chrome(sl)
    return [sl]


def product(d: Deck, s):
    """Scale + slide (video 1): product PNG scaled up and pushed partly off the
    right edge; numbered points on the left (park them with enter='left')."""
    sl = d.new_slide(); sh = sl.shapes
    img = d.asset(s.get("image"))
    if img:
        pic = d.picture(sh, img, d.W * 0.5, 0.2, d.W * 0.62, d.H - 0.4, mode="contain", name="!!product")
        if s.get("tilt", True) and d.animate:
            # a 3-D turn: flat on the reveal slide, tilted here; Morph rotates it in 3-D
            set_3d_rotation(pic, lat=8, lon=338, camera="perspectiveRelaxed")
    else:
        d.placeholder_art(sh, d.W * 0.55, 0.6, d.W * 0.5, d.H - 1.2, "product render", name="!!product")
    _kicker(d, sh, s.get("kicker"), y=0.9, w=6)
    if s.get("title"):
        d.text(sh, M, 1.3, 6.2, 1.0, s["title"], size=36, font="heavy", name="!!prod_name")
    for i, p in enumerate(s["points"][:4]):
        y = 2.6 + i * 1.0
        g = sh.add_group_shape(); g.name = f"pt{i+1}" if s.get("click_points", True) else f"!!pt{i+1}"
        d.text(g.shapes, M, y, 0.7, 0.8, f"{i+1}", size=30, font="heavy", color=d.c["accent"], name="aux_num")
        d.text(g.shapes, M + 0.7, y + 0.05, 5.3, 0.8, p, size=22, font="light")
        if s.get("click_points", True):
            d.tl(sl).add(g, "float_up", start="click", label=f"point {i+1}")
    d.chrome(sl)
    return [sl]


def hardware(d: Deck, s):
    """Annotated hardware tour (iPhone 2007): eval board photo with thin callout
    lines to short labels. Callout x,y are fractions of the image box.
    Each callout draws on click: point, then the line wipes out to its label.
    zoom=true adds a camera tour: one slide per callout with the photo scaled up
    around that point (Morph pans and zooms between them), then back out."""
    bx, by, bw, bh = 0.6, 0.8, 7.6, 5.9
    img = d.asset(s.get("image"))
    calls = sorted(s.get("callouts", [])[:6], key=lambda c: (c["y"], c["x"]))
    slides = []

    sl = d.new_slide(); sh = sl.shapes
    if img:
        pic = d.picture(sh, img, bx, by, bw, bh, mode="contain", name="!!board")
        # actual picture box (contain may letterbox)
        pbx, pby, pbw, pbh = pic.left / IN(1), pic.top / IN(1), pic.width / IN(1), pic.height / IN(1)
    else:
        d.placeholder_art(sh, bx, by, bw, bh, "eval board photo", name="!!board")
        pbx, pby, pbw, pbh = bx, by, bw, bh
    _kicker(d, sh, s.get("kicker"), y=0.5, x=8.8, w=4)
    lx = 9.0
    prev = 0.6
    tl = d.tl(sl)
    for i, c in enumerate(calls):
        px, py = bx + c["x"] * bw, by + c["y"] * bh
        ly = min(max(py, prev + 0.8), d.H - 1.0 - (len(calls) - 1 - i) * 0.8)
        prev = ly
        dot = d.rect(sh, px - 0.06, py - 0.06, 0.12, 0.12, d.fg(), kind=MSO_SHAPE.OVAL, name=f"aux_pt{i}")
        ln = d.line(sh, px, py, lx - 0.15, ly, d.fg(), 0.75, name=f"aux_ln{i}")
        lab = d.text(sh, lx, ly - 0.32, d.W - lx - 0.5, 0.4, c["label"], size=17, font="head", name=f"aux_call{i}")
        det = None
        if c.get("detail"):
            det = d.text(sh, lx, ly + 0.05, d.W - lx - 0.5, 0.4, c["detail"], size=13, font="light",
                         color=d.muted(), name=f"aux_calld{i}")
        # line drawing: dot appears, the line wipes toward the label, the label fades in
        tl.add(dot, "zoom", start="click", duration=250, label=f"callout: {c['label']}")
        tl.add(ln, "wipe", start="after", direction="left", duration=400)
        tl.add(lab, "fade", start="after", duration=300)
        if det is not None:
            tl.add(det, "fade", start="with", duration=300)
    d.chrome(sl)
    slides.append(sl)

    if s.get("zoom") and img:
        k = float(s.get("zoom_scale", 2.4))
        cxs, cys = 4.6, 3.75          # where the focus point lands on screen
        for i, c in enumerate(calls):
            z = d.new_slide(); zs = z.shapes
            # point in the picture's own fraction space
            fx = (bx + c["x"] * bw - pbx) / pbw
            fy = (by + c["y"] * bh - pby) / pbh
            nw, nh = pbw * k, pbh * k
            p2 = d.picture(zs, img, cxs - fx * nw, cys - fy * nh, nw, nh, mode="contain", name="!!board")
            d.rect(zs, 8.6, 0, d.W - 8.6, d.H, d.c["dark_bg"], opacity=0.92, name="aux_panel")
            ring = d.rect(zs, cxs - 0.55, cys - 0.55, 1.1, 1.1, None, line=d.c["accent"], line_w=2.5,
                          kind=MSO_SHAPE.OVAL, name="!!focus_ring")
            _kicker(d, zs, f"{i + 1} / {len(calls)}", y=2.2, x=9.0, w=3.8)
            d.text(zs, 9.0, 2.65, d.W - 9.5, 1.2, c["label"], size=fit_pt(c["label"], d.W - 9.5, 1.2, 32, 18),
                   font="heavy", name="!!zoom_label")
            if c.get("detail"):
                d.text(zs, 9.0, 3.9, d.W - 9.5, 1.6, c["detail"], size=18, font="light", color=d.muted(),
                       name="!!zoom_detail")
            d.morph_as(z, "morph", s.get("morph_ms", 1600))
            d.chrome(z)
            slides.append(z)
        if s.get("zoom_out", True):
            o = d.new_slide(); os_ = o.shapes
            d.picture(os_, img, bx, by, bw, bh, mode="contain", name="!!board")
            _kicker(d, os_, s.get("kicker"), y=0.5, x=8.8, w=4)
            prev = 0.6
            for i, c in enumerate(calls):
                px, py = bx + c["x"] * bw, by + c["y"] * bh
                ly = min(max(py, prev + 0.8), d.H - 1.0 - (len(calls) - 1 - i) * 0.8)
                prev = ly
                d.rect(os_, px - 0.06, py - 0.06, 0.12, 0.12, d.fg(), kind=MSO_SHAPE.OVAL, name=f"aux_pt{i}")
                d.line(os_, px, py, lx - 0.15, ly, d.fg(), 0.75, name=f"aux_ln{i}")
                d.text(os_, lx, ly - 0.32, d.W - lx - 0.5, 0.4, c["label"], size=17, font="head", name=f"aux_call{i}")
            d.morph_as(o, "morph", s.get("morph_ms", 1600))
            d.chrome(o)
            slides.append(o)
    return slides


def roadmap(d: Deck, s):
    """Chevron roadmap (video 3). N equal columns, dark -> light = first -> last,
    giant stage letter, header, light body; a right-pointing triangle at each
    boundary in the left column's colour. With a cover, the first slide stacks
    the columns off the left edge as thin stripes and Morph expands them."""
    stages = s["stages"][:6]
    n = len(stages)
    colw = d.W / n
    letter_pt = int(166 * min(1.0, 4 / max(n, 4)))
    head_w = colw - 0.9 - 0.25  # tri_w + 0.3 padding
    longest = max(len(st["header"]) for st in stages)
    head_pt = int(min(26 if n <= 4 else 21, head_w * 72 / (0.56 * max(longest, 1))))
    body_pt = 16 if n <= 4 else 14
    cols = [lerp_hex(d.c["ramp_start"], d.c["ramp_end"], i / max(n - 1, 1)) for i in range(n)]
    tri_w, tri_h = 0.6, 1.6
    stripe = 0.13
    pad = tri_w + 0.3

    def draw(sl, cover):
        sh = sl.shapes
        # draw last column first so each column's triangle sits on top of its neighbour
        for i in reversed(range(n)):
            st = stages[i]
            x = (-colw + stripe * (i + 1)) if cover else i * colw
            col = cols[i]
            txt = "FFFFFF" if luminance(col) < 0.4 else d.c["on_light"]
            g = sh.add_group_shape(); g.name = f"!!stage{i+1}"
            d.rect(g.shapes, x, 0, colw + 0.01, d.H, col)
            d.text(g.shapes, x + pad, 0.45, colw - pad, 2.5, st.get("letter", chr(65 + i)), size=letter_pt,
                   font="heavy", color=txt, name="aux_stage_letter")
            d.text(g.shapes, x + pad, 4.75, colw - pad - 0.25, 0.9, st["header"], size=head_pt, font="head", color=txt)
            if i < n - 1:
                d.triangle_right(g.shapes, x + colw, 2.7, tri_w, tri_h, col)
            if st.get("body") and not cover:
                bodies[i] = (d.text(sh, x + pad, 5.5, colw - pad - 0.3, 1.5, st["body"], size=body_pt,
                                    font="light", color=txt, name=f"stage_body{i+1}"), st["header"])
    slides = []
    bodies = {}
    if s.get("cover"):
        cv = d.new_slide(); sh = cv.shapes
        draw(cv, cover=True)
        cx = 1.6
        _kicker(d, sh, s["cover"].get("kicker"), y=2.5, x=cx)
        d.text(sh, cx, 3.0, d.W - cx - M, 1.8, s["cover"]["title"], size=72, font="heavy",
               color=d.c["accent"] if not d.dark else d.fg(), name="cover_title")
        if s["cover"].get("subtitle"):
            d.text(sh, cx, 4.8, 8, 0.8, s["cover"]["subtitle"], size=22, font="light", color=d.muted())
        d.chrome(cv)
        slides.append(cv)
    sl = d.new_slide()
    bodies.clear()
    draw(sl, cover=False)
    # walk the plan stage by stage: each stage's detail wipes in from the left on click
    for i in sorted(bodies):
        shp, head = bodies[i]
        d.tl(sl).add(shp, "wipe", start="click", direction="left", duration=500, label=f"stage: {head}")
    slides.append(sl)
    return slides


# ---------------------------------------------------------------- proof

def big_number(d: Deck, s):
    """One number, huge. from="3 days" adds a first slide with the old value; Morph by
    Characters turns the old number into the new one, digits sliding and swapping."""
    vals = ([s["from"]] if s.get("from") else []) + [s["value"]]
    size = s.get("size") or min(fit_pt(v, d.W - 2 * M, 2.9, 180, 60, line_spacing=0.85, char_w=0.6) for v in vals)
    slides = []
    for i, v in enumerate(vals):
        sl = d.new_slide(); sh = sl.shapes
        last = i == len(vals) - 1
        _kicker(d, sh, s.get("kicker") if last else s.get("from_kicker", s.get("kicker")), y=1.3)
        big = d.text(sh, M, 1.8, d.W - 2 * M, 2.9, v, size=size, font="heavy",
                     color=d.c["accent"] if last else d.c["rival"], line_spacing=0.85, name="!!big_value")
        lbl = s.get("label") if last else s.get("from_label", s.get("label"))
        if s.get("unit"):
            d.text(sh, M, 4.55, d.W - 2 * M, 0.8, s["unit"], size=34, font="head", name="big_unit")
        if lbl:
            d.text(sh, M, 5.35, d.W * 0.65, 1.0, lbl, size=22, font="light", color=d.muted(), name="!!big_label")
        if i:
            d.morph_as(sl, "morph_chars", s.get("morph_ms", 1400))
            d.tl(sl).add(big, "pulse", start="after", delay=150)
        d.chrome(sl)
        slides.append(sl)
    return slides


def side_by_side(d: Deck, s):
    """Before/after or competitor/ours at the SAME scale (Retina old vs new)."""
    sl = d.new_slide(); sh = sl.shapes
    _kicker(d, sh, s.get("kicker"), y=0.7)
    pw = (d.W - 2 * M - 0.6) / 2
    for i, side in enumerate([s["left"], s["right"]]):
        x = M + i * (pw + 0.6)
        img = d.asset(side.get("image"))
        if img:
            d.picture(sh, img, x, 1.3, pw, 3.7, mode="contain", name=f"!!sbs{i+1}")
        elif side.get("value"):
            d.text(sh, x, 1.3, pw, 3.7, side["value"], size=110, font="heavy", align="c", anchor="m",
                   color=d.c["ours"] if side.get("ours") else d.c["rival"], name=f"!!sbs{i+1}")
        else:
            d.placeholder_art(sh, x, 1.3, pw, 3.7, side["label"], name=f"!!sbs{i+1}")
        lbl = d.text(sh, x, 5.15, pw, 0.5, side["label"], size=22, font="head", align="c",
                     color=d.fg() if side.get("ours") else d.muted(), name=f"aux_sbslbl{i+1}")
        if i == 1 and s.get("build", True):
            names = {x_.name: x_ for x_ in sh}
            d.tl(sl).add(names[f"!!sbs{i+1}"], "fade", start="click", duration=500, label=f"show {side['label']}")
            d.tl(sl).add(lbl, "fade", start="with", duration=500)
    if s.get("caption"):
        cap = d.text(sh, M, 5.85, d.W - 2 * M, 0.6, s["caption"], size=20, font="light", align="c",
                     color=d.muted(), name="caption")
        d.tl(sl).add(cap, "fade", start="after", delay=300)
    d.chrome(sl)
    return [sl]


def partners(d: Deck, s):
    sl = d.new_slide(); sh = sl.shapes
    _kicker(d, sh, s.get("kicker", "Not just us"), y=1.0)
    items = s["items"][:4]
    n = len(items); colw = (d.W - 2 * M) / n
    for i, it in enumerate(items):
        x = M + i * colw
        logo = d.asset(it.get("logo"))
        if logo:
            pl = d.picture(sh, logo, x + 0.3, 2.2, colw - 0.6, 1.4, mode="contain", name=f"partner_logo{i}")
        else:
            pl = d.text(sh, x, 2.4, colw - 0.3, 1.0, it["name"], size=30, font="heavy", name=f"partner{i}")
        d.tl(sl).add(pl, "fade", start="with", delay=150 * i + 200)
        if it.get("role"):
            d.text(sh, x, 3.8, colw - 0.4, 1.4, it["role"], size=17, font="light", color=d.muted())
    if s.get("quote"):
        d.text(sh, M, 5.3, d.W - 2 * M, 1.0, f"“{s['quote']}”", size=22, font="light", italic=True)
    d.chrome(sl)
    return [sl]


# ---------------------------------------------------------------- close

def ask(d: Deck, s):
    """The specific next step, owner and date."""
    sl = d.new_slide(); sh = sl.shapes
    k = _kicker(d, sh, s.get("kicker", "Next step"), y=1.6)
    ask_pt = s.get("size") or fit_pt(s["ask"], d.W - 2 * M, 2.2, 52, 30, line_spacing=0.95)
    a = d.text(sh, M, 2.1, d.W - 2 * M, 2.2, s["ask"], size=ask_pt, font="heavy",
               line_spacing=0.95, name="!!ask")
    steps = [(k, "fade", 200), (a, "float_up", 400)]
    y = 4.65
    for j, lbl in enumerate(("owner", "date")):
        if s.get(lbl):
            g = sh.add_group_shape(); g.name = f"ask_{lbl}"
            d.text(g.shapes, M, y, 1.6, 0.5, lbl.upper(), size=13, font="head", color=d.c["accent"], spacing=2, name=f"aux_{lbl}_lbl")
            d.text(g.shapes, M + 1.6, y - 0.05, 9, 0.6, s[lbl], size=22, font="light", name=f"ask_{lbl}_txt")
            steps.append((g, "fade", 1000 + 250 * j))
            y += 0.65
    _intro(d, sl, steps)
    d.chrome(sl)
    return [sl]


def table(d: Deck, s):
    """Appendix table: the detail engineers will ask for. Light, dense, readable."""
    sl = d.new_slide(dark=False); sh = sl.shapes
    d.text(sh, M, 0.5, d.W - 2 * M, 0.7, s["title"], size=28, font="heavy", color=d.c["on_light"], name="tbl_title")
    cols = s["columns"]; rows = s["rows"]
    tw = d.W - 2 * M
    th = min(0.42 * (len(rows) + 1), d.H - 2.3)
    gf = sh.add_table(len(rows) + 1, len(cols), IN(M), IN(1.4), IN(tw), IN(th))
    gf.name = "appendix_table"
    tbl = gf.table
    from pptx.util import Pt
    from deckkit import rgb
    for j, c in enumerate(cols):
        cell = tbl.cell(0, j); cell.text = c
        cell.fill.solid(); cell.fill.fore_color.rgb = rgb(d.c["ramp_start"])
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = rgb("FFFFFF")
                r.font.name = d.t["fonts"]["head"]["face"]
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j); cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb("FFFFFF" if i % 2 else d.c["light_bg2"])
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(12); r.font.color.rgb = rgb(d.c["on_light"])
                    r.font.name = d.t["fonts"]["body"]["face"]
    if s.get("footnote"):
        d.text(sh, M, d.H - 1.0, d.W - 2 * M, 0.45, s["footnote"], size=11, font="light",
               color=d.c["on_light_muted"], name="aux_footnote")
    d.chrome(sl, dark=False)
    return [sl]


# ---------------------------------------------------------------- camera zoom (Morph)

def _chain_canvas(d: Deck, s):
    """Lay out a signal chain as canvas items in overview (slide) inches.
    Returns {id: item} with x, y, w, h, label, kind, ours, parent, pt."""
    chain = s["chain"]
    n = len(chain)
    gap = 0.55
    bw = (d.W - 2 * M - (n - 1) * gap) / n
    bh = 2.3
    y0 = 2.55
    items = {}
    for i, b in enumerate(chain):
        x = M + i * (bw + gap)
        bid = b.get("id", f"b{i+1}")
        items[bid] = dict(x=x, y=y0, w=bw, h=bh, label=b["label"], kind="block", ours=b.get("ours"),
                          sub=b.get("sub"), pt=20)
        inside = b.get("inside") or []
        if inside:
            k = len(inside)
            pad = 0.08
            iw = (bw - pad * (k + 1)) / k
            for j, lab in enumerate(inside):
                iid = f"{bid}.{j+1}"
                lab_txt, lab_sub = (lab, None) if isinstance(lab, str) else (lab["label"], lab.get("sub"))
                items[iid] = dict(x=x + pad + j * (iw + pad), y=y0 + bh * 0.52, w=iw, h=bh * 0.38,
                                  label=lab_txt, sub=lab_sub, kind="inner", ours=b.get("ours"), pt=5.5)
        if i < n - 1:
            items[f"arrow{i+1}"] = dict(x=x + bw + 0.08, y=y0 + bh / 2 - 0.16, w=gap - 0.16, h=0.32,
                                        label="", kind="arrow", pt=1)
    return items


def _camera(d: Deck, items, focus, margin=1.35, area=None):
    """Return (scale, dx, dy) mapping canvas -> screen so `focus` fills `area`."""
    ax, ay, aw, ah = area or (0, 0.6, d.W, d.H - 2.2)
    if focus is None:
        return 1.0, 0.0, 0.0
    it = items[focus]
    fw, fh = it["w"] * margin, it["h"] * margin
    sc = min(aw / fw, ah / fh)
    cx, cy = it["x"] + it["w"] / 2, it["y"] + it["h"] / 2
    return sc, ax + aw / 2 - cx * sc, ay + ah / 2 - cy * sc


def zoom_tour(d: Deck, s):
    """Camera move through a block diagram, the pro 'zoom into the detail' Morph.
    Slide 1 shows the whole chain; each stop scales and pans the SAME shapes so one
    block (or a block inside a block) fills the screen. Text, outlines and arrows all
    scale with the camera, so Morph flies the audience in rather than cutting.
    Blocks can hold `inside` items that are unreadably small in the overview and
    become legible when you zoom in: the infinite-zoom reveal."""
    items = _chain_canvas(d, s)
    stops = [None] + [st.get("focus") for st in s.get("stops", [])]
    if s.get("zoom_out", True) and len(stops) > 1:
        stops.append(None)
    slides = []
    for si, focus in enumerate(stops):
        sl = d.new_slide(); sh = sl.shapes
        sc, dx, dy = _camera(d, items, focus, area=s.get("area"))
        for iid, it in items.items():
            x, y = it["x"] * sc + dx, it["y"] * sc + dy
            w, h = it["w"] * sc, it["h"] * sc
            nm = "!!cv_" + iid.replace(".", "_")
            if it["kind"] == "arrow":
                d.rect(sh, x, y, w, h, d.muted(), kind=MSO_SHAPE.RIGHT_ARROW, name=nm)
                continue
            ours = it.get("ours")
            if it["kind"] == "block":
                fill, op, line = (d.c["accent"], 0.16, d.c["accent"]) if ours else (d.c["rival"], 0.18, d.muted())
            else:
                fill, op, line = (d.c["accent"], 0.45, None) if ours else (d.c["rival"], 0.4, None)
            box = d.rect(sh, x, y, w, h, fill, opacity=op, line=line, line_w=max(0.5, 1.5 * sc),
                         kind=MSO_SHAPE.ROUNDED_RECTANGLE, name=nm)
            box.adjustments[0] = 0.08
            pt = it["pt"] * sc
            if it["kind"] == "block":
                lab = d.text(sh, x + 0.12 * sc, y + 0.15 * sc, w - 0.24 * sc, 0.6 * sc, it["label"],
                             size=max(1, round(pt, 1)), font="heavy", align="c", name=nm + "_lbl")
                if it.get("sub"):
                    d.text(sh, x + 0.12 * sc, y + 0.68 * sc, w - 0.24 * sc, 0.45 * sc, it["sub"],
                           size=max(1, round(pt * 0.6, 1)), font="light", color=d.muted(), align="c",
                           name=nm + "_sub")
            else:
                d.text(sh, x + 0.04 * sc, y + 0.08 * sc, w - 0.08 * sc, h * 0.55, it["label"],
                       size=max(1, round(pt, 1)), font="head", align="c", anchor="m", name=nm + "_lbl")
                if it.get("sub"):
                    d.text(sh, x + 0.04 * sc, y + h * 0.58, w - 0.08 * sc, h * 0.38, it["sub"],
                           size=max(1, round(pt * 0.7, 1)), font="light", color=d.muted(), align="c",
                           name=nm + "_sub")
        # caption band (not part of the canvas): morphs its text between stops
        stop = s["stops"][si - 1] if (focus is not None and 0 < si <= len(s.get("stops", []))) else None
        if stop:
            band = d.rect(sh, 0, d.H - 1.55, d.W, 1.55, d.c["dark_bg"] if d.dark else d.c["light_bg"],
                          opacity=0.9, name="!!cv_band")
            _kicker(d, sh, stop.get("kicker") or items[focus]["label"], y=d.H - 1.4)
            cap = d.text(sh, M, d.H - 1.02, d.W - 2 * M, 0.6, stop.get("caption", ""), size=26,
                         font="heavy", name="!!cv_caption")
        else:
            _kicker(d, sh, s.get("kicker", "The signal chain"), y=0.6)
            if s.get("title"):
                d.text(sh, M, 1.0, d.W - 2 * M, 0.9, s["title"], size=36, font="heavy", name="!!cv_title")
        if si:
            d.morph_as(sl, "morph", s.get("morph_ms", 1800))
        d.chrome(sl)
        slides.append(sl)
    return slides


LAYOUTS = {
    "title": (title, "hook"), "mosaic": (mosaic, "hook"), "section": (section, "section"),
    "statement": (statement, "problem"), "credibility": (credibility, "credibility"),
    "objections": (objections, "objections"), "status_quo": (status_quo, "problem"),
    "criteria": (criteria, "criteria"), "quadrant": (quadrant, "positioning"),
    "bars": (bars, "proof"), "tease": (tease, "tease"), "reveal": (reveal, "reveal"),
    "feature": (feature, "feature"), "product": (product, "feature"),
    "hardware": (hardware, "feature"), "roadmap": (roadmap, "agenda"),
    "big_number": (big_number, "proof"), "side_by_side": (side_by_side, "proof"),
    "partners": (partners, "partners"), "ask": (ask, "ask"), "table": (table, "appendix"),
    "zoom_tour": (zoom_tour, "feature"),
}
