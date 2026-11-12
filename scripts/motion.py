"""
motion.py - slide transitions and in-slide animation, written as PowerPoint writes them.

Two layers, matching how professional decks are built:

1. Transitions (between slides): `set_transition(slide, kind, ...)`.
   Every PowerPoint transition is available: the 2007 set (fade, push, wipe, split,
   reveal-style cover/pull, zoom, shapes, ...), the 2010 set (p14: vortex, ripple,
   pan, conveyor, ferris, gallery, doors, window, shred, flash, glitter, honeycomb,
   flythrough, warp, switch, flip, prism/cube/box/rotate/orbit, reveal), the 2013 set
   (p15 presets: curtains, wind, prestige, fracture, crush, peelOff, pageCurl*,
   airplane, origami, fallOver, drape) and Morph (by object, word or character).
   Newer ones are wrapped in mc:AlternateContent with a fade fallback, exactly like
   PowerPoint's own files, so older viewers degrade gracefully.

2. Animations (inside a slide): `Timeline(slide)`.
   Entrance, emphasis, exit and motion-path effects; On Click / With Previous /
   After Previous; delays; smooth start/end; text by paragraph, word or letter.
   Written as p:timing with the same node structure PowerPoint uses, plus bldLst.

Design rules live in references/motion.md; this module only makes the XML.
"""
from __future__ import annotations

from lxml import etree
from pptx.oxml.ns import qn

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"

# ---------------------------------------------------------------------------
# Transitions
# ---------------------------------------------------------------------------

# kind -> (namespace tier, element, default attrs). Tier: base | p14 | p15 | p159
TRANSITIONS = {
    # 2007 set (p:)
    "cut": ("base", "cut", {}),
    "fade": ("base", "fade", {}),
    "fade_black": ("base", "fade", {"thruBlk": "1"}),
    "push": ("base", "push", {"dir": "u"}),          # dir: l r u d
    "wipe": ("base", "wipe", {"dir": "r"}),          # dir: l r u d
    "split": ("base", "split", {"orient": "horz", "dir": "out"}),
    "cover": ("base", "cover", {"dir": "l"}),        # 8 directions: l r u d lu ru ld rd
    "uncover": ("base", "pull", {"dir": "l"}),
    "zoom": ("base", "zoom", {"dir": "in"}),
    "circle": ("base", "circle", {}),
    "diamond": ("base", "diamond", {}),
    "plus": ("base", "plus", {}),
    "wedge": ("base", "wedge", {}),
    "wheel": ("base", "wheel", {"spokes": "1"}),
    "blinds": ("base", "blinds", {"dir": "horz"}),
    "checker": ("base", "checker", {"dir": "horz"}),
    "comb": ("base", "comb", {"dir": "horz"}),
    "dissolve": ("base", "dissolve", {}),
    "random_bars": ("base", "randomBar", {"dir": "vert"}),
    "strips": ("base", "strips", {"dir": "ru"}),
    "newsflash": ("base", "newsflash", {}),
    "random": ("base", "random", {}),
    # 2010 set (p14:)
    "reveal": ("p14", "reveal", {"thruBlk": "1", "dir": "r"}),   # smooth reveal; dir l r
    "pan": ("p14", "pan", {"dir": "u"}),                         # dir l r u d
    "vortex": ("p14", "vortex", {"dir": "r"}),
    "ripple": ("p14", "ripple", {"dir": "center"}),
    "honeycomb": ("p14", "honeycomb", {}),
    "glitter": ("p14", "glitter", {"dir": "l", "pattern": "hexagon"}),
    "shred": ("p14", "shred", {"pattern": "strip", "dir": "in"}),
    "flash": ("p14", "flash", {}),
    "switch": ("p14", "switch", {"dir": "r"}),
    "flip": ("p14", "flip", {"dir": "r"}),
    "gallery": ("p14", "gallery", {"dir": "l"}),
    "conveyor": ("p14", "conveyor", {"dir": "l"}),
    "ferris": ("p14", "ferris", {"dir": "l"}),
    "doors": ("p14", "doors", {"dir": "vert"}),
    "window": ("p14", "window", {"dir": "vert"}),
    "warp": ("p14", "warp", {"dir": "in"}),
    "flythrough": ("p14", "flythrough", {"dir": "in"}),
    "cube": ("p14", "prism", {"dir": "l"}),
    "box": ("p14", "prism", {"dir": "l", "isInverted": "1"}),
    "rotate": ("p14", "prism", {"dir": "l", "isContent": "1"}),
    "orbit": ("p14", "prism", {"dir": "l", "isContent": "1", "isInverted": "1"}),
    # 2013 set (p15:prstTrans)
    "curtains": ("p15", "curtains", {}),
    "wind": ("p15", "wind", {}),
    "prestige": ("p15", "prestige", {}),
    "fracture": ("p15", "fracture", {}),
    "crush": ("p15", "crush", {}),
    "peel_off": ("p15", "peelOff", {}),
    "page_curl": ("p15", "pageCurlDouble", {}),
    "page_curl_single": ("p15", "pageCurlSingle", {}),
    "airplane": ("p15", "airplane", {}),
    "origami": ("p15", "origami", {}),
    "fall_over": ("p15", "fallOver", {}),
    "drape": ("p15", "drape", {}),
    # Morph (2016+/365)
    "morph": ("p159", "morph", {"option": "byObject"}),
    "morph_words": ("p159", "morph", {"option": "byWord"}),
    "morph_chars": ("p159", "morph", {"option": "byChar"}),
}

# Grouping used by lint and the docs. "quiet" is what pros use in client decks.
QUIET = {"cut", "fade", "fade_black", "push", "wipe", "split", "uncover", "cover", "reveal", "pan",
         "morph", "morph_words", "morph_chars"}
LOUD = {"vortex", "ripple", "honeycomb", "glitter", "shred", "flash", "switch", "flip", "gallery",
        "conveyor", "ferris", "doors", "window", "warp", "flythrough", "cube", "box", "rotate",
        "orbit", "curtains", "wind", "prestige", "fracture", "crush", "peel_off", "page_curl",
        "page_curl_single", "airplane", "origami", "fall_over", "drape", "blinds", "checker",
        "comb", "dissolve", "random_bars", "strips", "newsflash", "random", "wheel", "wedge",
        "circle", "diamond", "plus", "zoom"}


def _remove_transitions(sld):
    for old in sld.findall(qn("p:transition")) + sld.findall("{%s}AlternateContent" % NS_MC):
        if old.tag == qn("p:transition") or old.find(".//" + qn("p:transition")) is not None:
            sld.remove(old)


def _attrs(d):
    return "".join(f' {k}="{v}"' for k, v in d.items())


def set_transition(slide, kind="fade", duration_ms=None, advance_after_ms=None, click=True, **opts):
    """Set the slide's entry transition. `kind` is a key of TRANSITIONS; opts override its
    attributes (e.g. dir="l"). duration_ms sets p14:dur (PowerPoint 2010+ honours exact ms)."""
    if kind in (None, "none"):
        _remove_transitions(slide._element)
        return
    tier, el, attrs = TRANSITIONS[kind]
    attrs = {**attrs, **{k: v for k, v in opts.items() if v is not None}}
    sld = slide._element
    _remove_transitions(sld)
    dur = int(duration_ms) if duration_ms else None
    spd = "fast" if (dur or 700) <= 500 else ("med" if (dur or 700) <= 750 else "slow")
    adv = ""
    if not click:
        adv += ' advClick="0"'
    if advance_after_ms is not None:
        adv += f' advTm="{int(advance_after_ms)}"'

    if tier == "base" and not dur:
        xml = f'<p:transition xmlns:p="{NS_P}" spd="{spd}"{adv}><p:{el}{_attrs(attrs)}/></p:transition>'
    elif tier == "base":
        xml = (f'<mc:AlternateContent xmlns:mc="{NS_MC}" xmlns:p="{NS_P}">'
               f'<mc:Choice xmlns:p14="{NS_P14}" Requires="p14">'
               f'<p:transition spd="{spd}" p14:dur="{dur}"{adv}><p:{el}{_attrs(attrs)}/></p:transition>'
               f'</mc:Choice><mc:Fallback><p:transition spd="{spd}"{adv}><p:{el}{_attrs(attrs)}/></p:transition>'
               f'</mc:Fallback></mc:AlternateContent>')
    else:
        dur_attr = f' p14:dur="{dur}"' if dur else ""
        if tier == "p14":
            req, decl, inner = "p14", "", f'<p14:{el}{_attrs(attrs)}/>'
        elif tier == "p15":
            inv = {k: attrs[k] for k in ("invX", "invY") if k in attrs}
            req, decl, inner = "p15", f' xmlns:p15="{NS_P15}"', f'<p15:prstTrans prst="{el}"{_attrs(inv)}/>'
        else:  # p159 morph
            req, decl, inner = "p159", f' xmlns:p159="{NS_P159}"', f'<p159:morph{_attrs(attrs)}/>'
        xml = (f'<mc:AlternateContent xmlns:mc="{NS_MC}" xmlns:p="{NS_P}">'
               f'<mc:Choice xmlns:p14="{NS_P14}"{decl} Requires="{req}">'
               f'<p:transition spd="{spd}"{dur_attr}{adv}>{inner}</p:transition></mc:Choice>'
               f'<mc:Fallback><p:transition spd="{spd}"{adv}><p:fade/></p:transition></mc:Fallback>'
               f'</mc:AlternateContent>')
    node = etree.fromstring(xml)
    anchor = sld.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = sld.find(qn("p:cSld"))
    anchor.addnext(node)


def get_transition(slide):
    """Return the transition kind key (best effort) or None."""
    sld = slide._element
    for tr in sld.iter(qn("p:transition")):
        parent = tr.getparent()
        if parent.tag == "{%s}Fallback" % NS_MC:
            continue
        for child in tr:
            tag = etree.QName(child).localname
            ns = etree.QName(child).namespace
            if ns == NS_P159:
                opt = child.get("option", "byObject")
                return {"byWord": "morph_words", "byChar": "morph_chars"}.get(opt, "morph")
            if ns == NS_P15:
                prst = child.get("prst")
                for k, v in TRANSITIONS.items():
                    if v[0] == "p15" and v[1] == prst:
                        return k
                return prst
            if tag == "fade" and child.get("thruBlk") in ("1", "true"):
                return "fade_black"
            if tag == "prism":
                return "cube"
            if tag == "pull":
                return "uncover"
            if tag == "randomBar":
                return "random_bars"
            return tag
        return "cut" if len(tr) == 0 else None
    return None


def transition_ms(slide):
    for tr in slide._element.iter(qn("p:transition")):
        v = tr.get("{%s}dur" % NS_P14)
        if v:
            return int(v)
    return None


# ---------------------------------------------------------------------------
# Animations
# ---------------------------------------------------------------------------

# effect name -> (presetClass, presetID, default subtype, default ms)
EFFECTS = {
    # entrance
    "appear": ("entr", 1, 0, 1),
    "fade": ("entr", 10, 0, 400),
    "wipe": ("entr", 22, 8, 500),           # direction: left (from left), right, up, down
    "fly": ("entr", 2, 4, 600),             # direction: from bottom/left/right/top
    "float_up": ("entr", 42, 0, 600),       # Ascend: fade + short rise, smooth end
    "zoom": ("entr", 53, 16, 500),          # subtle grow from 85% + fade
    "split": ("entr", 16, 21, 500),
    # emphasis
    "pulse": ("emph", 26, 0, 500),
    "grow": ("emph", 6, 0, 600),            # scale by `scale` (default 1.1)
    "ken_burns": ("emph", 6, 0, 9000),      # slow push-in on a photo, with the slide
    "spin": ("emph", 8, 0, 800),
    "dim": ("emph", 9, 0, 400),             # transparency to `opacity`
    # exit
    "fade_out": ("exit", 10, 0, 300),
    "wipe_out": ("exit", 22, 4, 400),
    # motion path
    "move": ("path", 0, 0, 800),            # dx, dy in inches; smooth by default
}

# direction = where the effect comes FROM, as PowerPoint's Effect Options name it
_WIPE = {"left": (8, "wipe(left)"), "right": (2, "wipe(right)"), "top": (1, "wipe(up)"), "bottom": (4, "wipe(down)")}
_WIPE["up"], _WIPE["down"] = _WIPE["bottom"], _WIPE["top"]
_FLY = {"bottom": (4, ("#ppt_x", "#ppt_x"), ("1+#ppt_h/2", "#ppt_y")),
        "top": (1, ("#ppt_x", "#ppt_x"), ("0-#ppt_h/2", "#ppt_y")),
        "left": (8, ("0-#ppt_w/2", "#ppt_x"), ("#ppt_y", "#ppt_y")),
        "right": (2, ("1+#ppt_w/2", "#ppt_x"), ("#ppt_y", "#ppt_y"))}


class _Ids:
    def __init__(self):
        self.n = 0

    def __call__(self):
        self.n += 1
        return self.n


class Timeline:
    """Collect animations for one slide, then `commit()` writes p:timing.

        tl = Timeline(slide, slide_w_in=13.333, slide_h_in=7.5)
        tl.add(shape, "fade", start="click", label="answer appears")
        tl.add(line, "wipe", start="with", direction="left")
        tl.add(label, "fade", start="after", delay=100)

    start: "click" (new click), "with" (with previous), "after" (after previous).
    The very first item may be "with"/"after": it then runs as the slide opens.
    Text options: by="paragraph" (one click per paragraph; start applies to the first),
    by="word" / by="letter" (iterate inside one effect, `iterate_ms` apart).
    """

    def __init__(self, slide, slide_w_in=13.333, slide_h_in=7.5):
        self.slide = slide
        self.W, self.H = slide_w_in, slide_h_in
        self.items = []

    # public -------------------------------------------------------------
    def add(self, shape, effect="fade", start="click", delay=0, duration=None, direction=None,
            by=None, iterate_ms=60, smooth=True, scale=1.1, opacity=0.35, dx=0.0, dy=0.0,
            repeat=None, label=None, paragraphs=None):
        if effect not in EFFECTS:
            raise ValueError(f"unknown effect {effect}; valid: {', '.join(EFFECTS)}")
        if by == "paragraph":
            n = paragraphs or len(shape.text_frame.paragraphs)
            for i in range(n):
                self.items.append(dict(shape=shape, effect=effect, start=start if i == 0 else "click",
                                       delay=delay, duration=duration, direction=direction, para=i,
                                       by=None, iterate_ms=iterate_ms, smooth=smooth, scale=scale,
                                       opacity=opacity, dx=dx, dy=dy, repeat=repeat,
                                       label=(label or "next line") if i else label))
            return self
        self.items.append(dict(shape=shape, effect=effect, start=start, delay=delay, duration=duration,
                               direction=direction, para=None, by=by, iterate_ms=iterate_ms,
                               smooth=smooth, scale=scale, opacity=opacity, dx=dx, dy=dy,
                               repeat=repeat, label=label))
        return self

    @property
    def clicks(self):
        return sum(1 for i, it in enumerate(self.items) if it["start"] == "click")

    def click_labels(self):
        return [it.get("label") or f"{it['effect']}" for it in self.items if it["start"] == "click"]

    # xml ------------------------------------------------------------------
    def _tgt(self, it):
        spid = it["shape"].shape_id
        if it["para"] is not None:
            p = it["para"]
            return f'<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{p}" end="{p}"/></p:txEl></p:spTgt></p:tgtEl>'
        return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'

    def _set_vis(self, ids, tgt, val, delay=0):
        return (f'<p:set><p:cBhvr><p:cTn id="{ids()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="{delay}"/>'
                f'</p:stCondLst></p:cTn>{tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName>'
                f'</p:attrNameLst></p:cBhvr><p:to><p:strVal val="{val}"/></p:to></p:set>')

    def _anim(self, ids, tgt, attr, v0, v1, dur, additive=True):
        add = ' additive="base"' if additive else ""
        return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr{add}><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>'
                f'{tgt}<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:tavLst><p:tav tm="0"><p:val><p:strVal val="{v0}"/></p:val></p:tav>'
                f'<p:tav tm="100000"><p:val><p:strVal val="{v1}"/></p:val></p:tav></p:tavLst></p:anim>')

    def _effect(self, ids, it, node_type):
        cls, pid, sub, dflt = EFFECTS[it["effect"]]
        cid = ids()  # effect node id first, so ids run in document order
        dur = int(it["duration"] or dflt)
        tgt = self._tgt(it)
        eff = it["effect"]
        d = it["direction"]
        body = ""
        if eff == "wipe":
            sub, filt = _WIPE[d or "left"]
            body = self._set_vis(ids, tgt, "visible") + (
                f'<p:animEffect transition="in" filter="{filt}"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}'
                f'</p:cBhvr></p:animEffect>')
        elif eff == "wipe_out":
            sub, filt = _WIPE[d or "bottom"]
            body = (f'<p:animEffect transition="out" filter="{filt}"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}'
                    f'</p:cBhvr></p:animEffect>') + self._set_vis(ids, tgt, "hidden", delay=max(dur - 1, 0))
        elif eff == "appear":
            body = self._set_vis(ids, tgt, "visible")
        elif eff == "fade":
            body = self._set_vis(ids, tgt, "visible") + (
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}'
                f'</p:cBhvr></p:animEffect>')
        elif eff == "fly":
            sub, xs, ys = _FLY[d or "bottom"]
            body = (self._set_vis(ids, tgt, "visible") + self._anim(ids, tgt, "ppt_x", *xs, dur)
                    + self._anim(ids, tgt, "ppt_y", *ys, dur))
        elif eff == "float_up":
            body = (self._set_vis(ids, tgt, "visible")
                    + f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}'
                      f'</p:cBhvr></p:animEffect>'
                    + self._anim(ids, tgt, "ppt_x", "#ppt_x", "#ppt_x", dur)
                    + self._anim(ids, tgt, "ppt_y", "#ppt_y+.06", "#ppt_y", dur))
        elif eff == "zoom":
            body = (self._set_vis(ids, tgt, "visible")
                    + self._anim(ids, tgt, "ppt_w", "#ppt_w*0.85", "#ppt_w", dur)
                    + self._anim(ids, tgt, "ppt_h", "#ppt_h*0.85", "#ppt_h", dur)
                    + f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}'
                      f'</p:cBhvr></p:animEffect>')
        elif eff == "split":
            body = self._set_vis(ids, tgt, "visible") + (
                f'<p:animEffect transition="in" filter="barn(inVertical)"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>'
                f'{tgt}</p:cBhvr></p:animEffect>')
        elif eff == "pulse":
            body = (f'<p:animScale><p:cBhvr><p:cTn id="{ids()}" dur="{dur // 2}" autoRev="1" fill="hold"/>{tgt}'
                    f'</p:cBhvr><p:by x="108000" y="108000"/></p:animScale>')
        elif eff in ("grow", "ken_burns"):
            s = int((it["scale"] if eff == "grow" else max(it["scale"], 1.06)) * 100000)
            body = (f'<p:animScale><p:cBhvr><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>{tgt}'
                    f'</p:cBhvr><p:by x="{s}" y="{s}"/></p:animScale>')
        elif eff == "spin":
            body = (f'<p:animRot by="21600000"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>{tgt}'
                    f'<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>')
        elif eff == "dim":
            body = (f'<p:animEffect filter="image" prLst="opacity: {it["opacity"]}"><p:cBhvr rctx="IE">'
                    f'<p:cTn id="{ids()}" dur="{dur}" fill="hold"/>{tgt}</p:cBhvr></p:animEffect>')
        elif eff == "fade_out":
            body = (f'<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt}'
                    f'</p:cBhvr></p:animEffect>') + self._set_vis(ids, tgt, "hidden", delay=max(dur - 1, 0))
        elif eff == "move":
            px = it["dx"] / self.W
            py = it["dy"] / self.H
            body = (f'<p:animMotion origin="layout" path="M 0 0 L {px:.5f} {py:.5f} E" pathEditMode="relative" '
                    f'ptsTypes="AA"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>{tgt}'
                    f'<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
                    f'</p:cBhvr><p:rCtr x="{int(px * 50000)}" y="{int(py * 50000)}"/></p:animMotion>')
        smooth = ""
        if it["smooth"] and eff in ("fly", "float_up", "move", "zoom", "grow"):
            smooth = ' decel="100000"' if eff != "move" else ' accel="30000" decel="70000"'
        if eff == "ken_burns":
            smooth = ' decel="60000"'
        rep = f' repeatCount="{int(it["repeat"] * 1000)}"' if it["repeat"] else ""
        iterate = ""
        if it["by"] in ("word", "letter"):
            iterate = (f'<p:iterate type="{"wd" if it["by"] == "word" else "lt"}">'
                       f'<p:tmAbs val="{int(it["iterate_ms"])}"/></p:iterate>')
        grp = 0
        return (f'<p:par><p:cTn id="{cid}" presetID="{pid}" presetClass="{cls}" presetSubtype="{sub}"{smooth}{rep} '
                f'fill="hold" grpId="{grp}" nodeType="{node_type}"><p:stCondLst><p:cond delay="{int(it["delay"])}"/>'
                f'</p:stCondLst>{iterate}<p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>'), \
            int(it["delay"]) + dur * (2 if eff == "pulse" else 1) + (
                int(it["iterate_ms"]) * 12 if iterate else 0)

    def commit(self):
        sld = self.slide._element
        old = sld.find(qn("p:timing"))
        if old is not None:
            sld.remove(old)
        if not self.items:
            return 0
        # group: list of click groups; each is a list of sub-groups (after-previous chains);
        # each sub-group is a list of items starting together (with previous)
        groups = []
        for i, it in enumerate(self.items):
            st = it["start"]
            if st == "click" or not groups:
                groups.append({"auto": st != "click", "subs": [[it]]})
            elif st == "with":
                groups[-1]["subs"][-1].append(it)
            else:  # after
                groups[-1]["subs"].append([it])

        # ids must be unique and increasing in document order; build top-down
        ids = _Ids()
        ids()  # 1 tmRoot
        ids()  # 2 mainSeq
        click_xml = []
        for g in groups:
            gid = ids()
            subs_xml = []
            t = 0
            for si, sub in enumerate(g["subs"]):
                sid = ids()
                effs = []
                end = 0
                for k, it in enumerate(sub):
                    if si == 0 and k == 0:
                        nt = "afterEffect" if (g["auto"] and it["start"] == "after") else (
                            "withEffect" if g["auto"] else "clickEffect")
                    elif k == 0:
                        nt = "afterEffect"
                    else:
                        nt = "withEffect"
                    x, e = self._effect(ids, it, nt)
                    effs.append(x)
                    end = max(end, e)
                subs_xml.append(f'<p:par><p:cTn id="{sid}" fill="hold"><p:stCondLst><p:cond delay="{t}"/></p:stCondLst>'
                                f'<p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par>')
                t += end
            st = '<p:cond delay="indefinite"/>'
            if g["auto"]:
                st += '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
            click_xml.append(f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst>{st}</p:stCondLst>'
                             f'<p:childTnLst>{"".join(subs_xml)}</p:childTnLst></p:cTn></p:par>')

        bld = []
        seen = set()
        for it in self.items:
            sh = it["shape"]
            el = sh._element
            if el.tag != qn("p:sp") or el.find(qn("p:txBody")) is None:
                continue
            if sh.shape_id in seen:
                continue
            seen.add(sh.shape_id)
            paras = any(j["para"] is not None for j in self.items if j["shape"] is sh)
            if paras:
                bld.append(f'<p:bldP spid="{sh.shape_id}" grpId="0" build="p"/>')
            else:
                has_fill = el.find(qn("p:spPr")).find(qn("a:solidFill")) is not None
                bld.append(f'<p:bldP spid="{sh.shape_id}" grpId="0"{" animBg=\"1\"" if has_fill else ""}/>')
        bld_xml = f'<p:bldLst>{"".join(bld)}</p:bldLst>' if bld else ""
        xml = (f'<p:timing xmlns:p="{NS_P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
               f'<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
               f'<p:childTnLst>{"".join(click_xml)}</p:childTnLst></p:cTn>'
               f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
               f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
               f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>{bld_xml}</p:timing>')
        node = etree.fromstring(xml)
        ext = sld.find(qn("p:extLst"))
        if ext is not None:
            ext.addprevious(node)
        else:
            sld.append(node)
        return self.clicks


def has_timing(slide):
    return slide._element.find(qn("p:timing")) is not None


def count_clicks(slide):
    """Number of on-click steps in a slide's main sequence (for lint and notes)."""
    tm = slide._element.find(qn("p:timing"))
    if tm is None:
        return 0
    n = 0
    for c in tm.iter(qn("p:cTn")):
        if c.get("nodeType") == "clickEffect":
            n += 1
    return n


def effect_summary(slide):
    """List of (presetClass, presetID, dur_ms) for lint."""
    tm = slide._element.find(qn("p:timing"))
    out = []
    if tm is None:
        return out
    for c in tm.iter(qn("p:cTn")):
        if c.get("presetClass"):
            durs = [int(x.get("dur")) for x in c.iter(qn("p:cTn")) if (x.get("dur") or "").isdigit() and x is not c]
            out.append((c.get("presetClass"), int(c.get("presetID", 0)), max(durs) if durs else 0))
    return out


# ---------------------------------------------------------------------------
# 3-D rotation (Morph animates between two rotations)
# ---------------------------------------------------------------------------

def set_3d_rotation(shape, lat=0.0, lon=0.0, rev=0.0, camera="perspectiveFront", fov=None):
    """Give a shape or picture a 3-D camera rotation in degrees. On a Morph pair, a flat
    shape on slide A and a rotated one on slide B turns in 3-D during the transition."""
    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    spPr = shape._element.find(qn("p:spPr"))
    for old in spPr.findall(qn("a:scene3d")) + spPr.findall(qn("a:sp3d")):
        spPr.remove(old)
    f = f' fov="{int(fov * 60000)}"' if fov else ""
    xml = (f'<a:scene3d xmlns:a="{NS_A}"><a:camera prst="{camera}"{f}>'
           f'<a:rot lat="{int(lat * 60000) % 21600000}" lon="{int(lon * 60000) % 21600000}" rev="{int(rev * 60000) % 21600000}"/>'
           f'</a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>')
    node = etree.fromstring(xml)
    # order in spPr: xfrm, geometry, fill, ln, effectLst/effectDag, scene3d, sp3d, extLst
    ext = spPr.find(qn("a:extLst"))
    if ext is not None:
        ext.addprevious(node)
    else:
        spPr.append(node)
