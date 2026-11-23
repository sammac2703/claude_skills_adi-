"""Write assets/layout_facts.json from the Slide Library: per layout the background (dark or light), whether the logo
and icons should be white, every placeholder (idx, type, position in inches) and a one-line use hint.
Scripts read this file instead of the model re-reading layout tables.   python scripts/extract_layout_facts.py"""
import json
import re
from pathlib import Path

from lxml import etree
from pptx import Presentation

SK = Path(__file__).resolve().parent.parent
LIB = SK / "assets" / "Brand-PPT-Slide-Library.pptx"
E = 914400
SCHEME = {"accent1": "00325C", "tx2": "0067B9", "bg1": "FFFFFF"}
HINT = {
    "Cover Slide with Pattern": "opening; title 10, subtitle 14, date 15",
    "Big Statement/Quote - Primary Blue": "one bold sentence or purpose; statement 27, attribution 28",
    "Big Statement Two Column": "statement plus three supports (messaging pillars)",
    "Strategic Priorities": "three rows of label and text (three wins, priorities)",
    "ToC / Agenda Slide": "agenda or steps, up to six items",
    "KPI Slide (x4)": "four big numbers",
    "KPIs plus Side Content": "four numbers plus commentary panel",
    "Event Roadmap/Timeline": "five or six phases with labels",
    "Progress Status/Timeline": "status grid over time",
    "Project Roadmap 2": "Gantt-style plan",
    "Employee Recognition / Team Success": "people, buddies, three stories with icons",
    "Product Features / Stories": "five items in a 3 plus 2 grid",
    "Two Columns, Dark": "before and after, start and end with arrow",
    "Four Columns": "four items with frames",
    "Next Steps / Action Items (Light)": "actions with a chart",
    "Next Steps / Action Items (Dark)": "actions with a chart",
    "Product or Strategic Dashboard": "portfolio rows",
    "1_Title and Content - Dark Blue": "custom dark composition (icons, rules, hero text)",
    "2_Title Only – Blank": "custom light composition, tables, charts",
    "End Card": "closing",
}


def solid(bg):
    pr = bg.find("{*}bgPr")
    if pr is None:
        return None
    if pr.find("{*}blipFill") is not None:
        return "picture"
    f = pr.find("{*}solidFill")
    if f is None:
        return "gradient" if pr.find("{*}gradFill") is not None else None
    c = f[0]
    if c.tag.endswith("srgbClr"):
        return c.get("val").upper()
    v = SCHEME.get(c.get("val"))
    if v and len(c):
        m = [k for k in c if k.tag.endswith("lumMod")]
        if m:
            x = int(255 * int(m[0].get("val")) / 100000)
            return f"{x:02X}{x:02X}{x:02X}"
    return v


def main():
    prs = Presentation(LIB)
    out = {}
    for l in prs.slide_layouts:
        bg = l._element.find(".//{*}bg")
        col = solid(bg) if bg is not None else "FFFFFF"
        dark = col in ("00325C", "0067B9", "gradient", "picture")
        panels = [{"x": round(s.left / E, 2), "w": round(s.width / E, 2), "fill": None}
                  for s in l.shapes if s.shape_type == 1 and s.width > 3 * E and s.height > 6 * E]
        phs = [{"idx": p.placeholder_format.idx, "type": str(p.placeholder_format.type).split(".")[-1].split(" ")[0],
                "x": round(p.left / E, 2), "y": round(p.top / E, 2), "w": round(p.width / E, 2), "h": round(p.height / E, 2)}
               for p in l.placeholders]
        out[l.name] = {"background": col, "dark": dark, "logo": "white" if dark else "blue",
                       "icons": "white" if dark else "ADI blue 0067B9", "panels": panels,
                       "use": HINT.get(l.name, ""), "placeholders": phs}
    (SK / "assets" / "layout_facts.json").write_text(json.dumps(out, indent=1), encoding="utf-8")
    print(f"{len(out)} layouts -> assets/layout_facts.json")


if __name__ == "__main__":
    main()
