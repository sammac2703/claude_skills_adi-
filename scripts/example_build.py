"""Sample deck v2 on the ADI 2026 standard template: layouts by name, theme colors only, card patterns from the slide library."""
import copy, sys
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt, Inches

BASE, OUT = sys.argv[1], sys.argv[2]
FOOTER = "©2026 Analog Devices, Inc. All Rights Reserved."

prs = Presentation(BASE)
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)
layouts = {l.name: l for l in prs.slide_layouts}


def L(prefix):
    return next(l for n, l in layouts.items() if n.startswith(prefix))


def ph(slide, idx):
    return next(s for s in slide.placeholders if s.placeholder_format.idx == idx)


def drop_ph(slide, idx):
    el = ph(slide, idx)._element
    el.getparent().remove(el)


def add_footer(slide):
    for p in slide.slide_layout.placeholders:
        if p.placeholder_format.idx in (3, 4):
            slide.shapes._spTree.append(copy.deepcopy(p._element))
    ph(slide, 3).text_frame.text = FOOTER


def fill(shape, items):
    tf = shape.text_frame
    for i, it in enumerate(items):
        lvl, txt = it if isinstance(it, tuple) else (0, it)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text, p.level = txt, lvl


# ---- shape helpers (theme colors only) ----
def box(slide, shape, x, y, w, h, color, brightness=0.0, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.theme_color = color
    if brightness:
        s.fill.fore_color.brightness = brightness
    s.line.fill.background()
    s.shadow.inherit = False
    if name:
        s.name = name
    return s


def text(slide, x, y, w, h, s, size, color=TC.TEXT_1, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, name=None):
    t = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = t.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.theme_color = color
    if name:
        t.name = name
    return t


def badge(slide, x, y, d, label, bg, fg, size=18, name=None):
    c = box(slide, MSO_SHAPE.OVAL, x, y, d, d, bg, name=name)
    tf = c.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.theme_color = fg
    return c


# 1 Cover
s = prs.slides.add_slide(L("Cover with Pattern"))
ph(s, 10).text_frame.text = "Quarterly Business Review"
ph(s, 14).text_frame.text = "Samuel Mackay"
ph(s, 15).text_frame.text = "November 2026"

# 2 Agenda: four numbered cards
s = prs.slides.add_slide(L("Title Only"))
s.shapes.title.text = "Agenda"
items = [("Performance Highlights", "Revenue trends by quarter"),
         ("Program Status", "What is on track and what is at risk"),
         ("Key Message", "The one takeaway to remember"),
         ("Next Steps", "Actions, owners and timing")]
cw, gap, y0, ch_ = 2.88, 0.30, 1.9, 4.2
for i, (t, d) in enumerate(items):
    x = 0.46 + i * (cw + gap)
    box(s, MSO_SHAPE.RECTANGLE, x, y0, cw, ch_, TC.ACCENT_2, 0.88, name=f"Agenda card {i+1}")
    badge(s, x + 0.3, y0 + 0.35, 0.7, str(i + 1), TC.ACCENT_2, TC.BACKGROUND_1, 24, name=f"Agenda number {i+1}")
    text(s, x + 0.3, y0 + 1.45, cw - 0.6, 1.0, t, 20, TC.ACCENT_1, True, name=f"Agenda title {i+1}")
    text(s, x + 0.3, y0 + 2.6, cw - 0.6, 1.2, d, 16, name=f"Agenda detail {i+1}")
add_footer(s)

# 3 Section
s = prs.slides.add_slide(L("Section Slide with Pattern"))
ph(s, 10).text_frame.text = "Performance Highlights"

# 4 Chart
s = prs.slides.add_slide(L("Title Only"))
s.shapes.title.text = "Revenue by Quarter (Illustrative)"
cd = CategoryChartData()
cd.categories = ["Q1", "Q2", "Q3", "Q4"]
cd.add_series("FY25", (4.3, 2.5, 3.5, 4.5))
cd.add_series("FY26", (2.4, 4.4, 1.8, 2.8))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.46), Inches(1.5),
                        Inches(12.42), Inches(5.4), cd)
gf.name = "Revenue chart"
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(14)
ch.category_axis.tick_labels.font.size = Pt(14)
ch.value_axis.tick_labels.font.size = Pt(14)
ch.value_axis.major_gridlines.format.line.color.theme_color = TC.BACKGROUND_2
plot = ch.plots[0]
plot.gap_width = 80
plot.has_data_labels = True
plot.data_labels.font.size = Pt(14)
plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
for ser, c in zip(plot.series, (TC.ACCENT_1, TC.ACCENT_2)):
    ser.format.fill.solid()
    ser.format.fill.fore_color.theme_color = c
add_footer(s)

# 5 Program status: two cards with status headers
s = prs.slides.add_slide(L("Title Only"))
s.shapes.title.text = "Program Status"
cols = [("On Track", TC.ACCENT_4, TC.BACKGROUND_1, "✓",
         ["Design review complete", "Samples shipped to lead customers", "Reliability qualification passed"]),
        ("At Risk", TC.ACCENT_5, TC.TEXT_1, "!",
         ["Test capacity constraint in Q2", "Mitigation plan under review", "Customer schedule dependency"])]
for i, (head, c, fg, glyph, rows) in enumerate(cols):
    x = 0.46 + i * 6.42
    box(s, MSO_SHAPE.RECTANGLE, x, 1.7, 6.0, 5.0, TC.ACCENT_2, 0.9, name=f"{head} card")
    box(s, MSO_SHAPE.RECTANGLE, x, 1.7, 6.0, 0.8, TC.ACCENT_1, name=f"{head} header")
    text(s, x + 0.35, 1.7, 5.3, 0.8, head, 24, TC.BACKGROUND_1, True, anchor=MSO_ANCHOR.MIDDLE,
         name=f"{head} header text")
    for j, r in enumerate(rows):
        y = 2.95 + j * 1.2
        badge(s, x + 0.35, y, 0.5, glyph, c, fg, 20, name=f"{head} marker {j+1}")
        text(s, x + 1.1, y, 4.6, 0.5, r, 18, anchor=MSO_ANCHOR.MIDDLE, name=f"{head} item {j+1}")
add_footer(s)

# 6 Big statement
s = prs.slides.add_slide(L("Big Statement - Primary Blue"))
ph(s, 28).text_frame.text = "One clear message per slide"
add_footer(s)

# 7 Next steps: three numbered cards on dark blue
s = prs.slides.add_slide(L("Title and Content - Dark Blue"))
s.shapes.title.text = "Next Steps"
drop_ph(s, 11)
steps = [("Confirm Q2 test capacity", "Align with operations on tester availability and schedule"),
         ("Finalize sampling plan", "Lock the customer list, quantities and ship dates"),
         ("Monthly checkpoint", "Review progress and risks with program leads")]
cw, gap = 3.87, 0.40
for i, (t, d) in enumerate(steps):
    x = 0.46 + i * (cw + gap)
    box(s, MSO_SHAPE.RECTANGLE, x, 1.9, cw, 4.4, TC.TEXT_2, name=f"Step card {i+1}")
    badge(s, x + 0.35, 2.25, 0.7, str(i + 1), TC.BACKGROUND_1, TC.TEXT_2, 24, name=f"Step number {i+1}")
    text(s, x + 0.35, 3.3, cw - 0.7, 1.0, t, 22, TC.BACKGROUND_1, True, name=f"Step title {i+1}")
    text(s, x + 0.35, 4.5, cw - 0.7, 1.5, d, 16, TC.BACKGROUND_1, name=f"Step detail {i+1}")
add_footer(s)

# 8 End card
prs.slides.add_slide(L("End Card"))

prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
