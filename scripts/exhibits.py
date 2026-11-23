"""Consulting-style exhibits in ADI colours, all native editable objects, 14 pt floor, one highlighted insight.

    from exhibits import waterfall, matrix2x2, gantt, rating_table, funnel, source_line, takeaway
    waterfall(s, ["FY25", "Price", "Volume", "Cost", "FY26"], [100, 12, 8, -6, None], 0.46, 1.9, 12.4, 4.3, highlight=2)
    matrix2x2(b, s, 0.9, 1.9, 7.2, 4.6, "Effort", "Impact", ["Quick wins", "Big bets", "Fill-ins", "Avoid"], [("A", .2, .8)])
    gantt(b, s, [("Onboard", 1, 2), ("Build", 3, 7)], 10, 0.46, 1.9, 12.4)
    rating_table(b, s, ["Option", "Cost", "Risk"], [("A", 3, 2), ("B", 1, 4)], 0.46, 1.9, [4, 3, 3])
    funnel(b, s, [("Candidates", 40), ("Interviewed", 12), ("Offers", 4)], 0.46, 1.9, 7.5, 4.3)
    takeaway(b, s, "Price and volume offset cost", y=6.35); source_line(b, s, "Source: ...")

Rules: one accent (ADI light blue) marks the point of the slide, everything else blue or grey; value labels on the
bars; a takeaway sentence and a source line on data slides (references/exhibit-rules.md). `b` is an adi_blocks.Blocks.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from adi_blocks import BLUE, CHAR, GREY, LIGHT, NAVY, WHITE   # noqa: E402
from pptx.chart.data import CategoryChartData               # noqa: E402
from pptx.dml.color import RGBColor                         # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE                       # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN              # noqa: E402
from pptx.util import Inches, Pt                             # noqa: E402


def _ink(dark):
    return WHITE if dark else NAVY


def _sub(dark):
    return GREY if dark else CHAR


def _rect(s, x, y, w, h, fill, name, shape=MSO_SHAPE.RECTANGLE, line=None):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        r.fill.background()
    if line:
        r.line.color.rgb = RGBColor.from_string(line)
        r.line.width = Pt(1)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    r.name = name
    return r


def takeaway(b, s, text, y=6.3, dark=False, x=0.46, w=11.0, size=24):
    """One punchy sentence under the exhibit (the 'so what')."""
    return b.text(s, x, y, w, 0.5, text, size, LIGHT if dark else BLUE, True, name="Takeaway")


def source_line(b, s, text, y=6.82, dark=False, x=0.46, w=11.0):
    """Source and footnote line, 14 pt, above the legal footer."""
    return b.text(s, x, y, w, 0.3, text, 14, _sub(dark), name="Source")


def waterfall(s, labels, values, x, y, w, h, dark=False, highlight=None, number_format="0", name="Waterfall"):
    """Bridge chart (native stacked column with an invisible base). values: first = start total, middle = deltas,
    last = None to compute the end total (or a number). highlight = index of the bar to colour ADI light blue.
    Returns the chart frame."""
    vals = list(values)
    if vals[-1] is None:
        vals[-1] = sum(vals[:-1])
    base, up, down, tot = [], [], [], []
    run = 0
    for i, v in enumerate(vals):
        if i == 0 or i == len(vals) - 1:
            base.append(0), up.append(0), down.append(0), tot.append(v)
            run = v if i == 0 else run
        elif v >= 0:
            base.append(run), up.append(v), down.append(0), tot.append(0)
            run += v
        else:
            run += v
            base.append(run), up.append(0), down.append(-v), tot.append(0)
    cd = CategoryChartData()
    cd.categories = labels
    for nm, ser in (("Base", base), ("Increase", up), ("Decrease", down), ("Total", tot)):
        cd.add_series(nm, ser)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    gf.name = name
    ch = gf.chart
    ink = _ink(dark)
    ch.has_legend = False
    ch.has_title = False
    ch.font.size = Pt(14)
    ch.font.color.rgb = RGBColor.from_string(ink)
    plot = ch.plots[0]
    plot.gap_width = 40
    plot.overlap = 100
    cols = {1: BLUE, 2: GREY, 3: NAVY if not dark else WHITE}
    ser = plot.series
    ser[0].format.fill.background()
    ser[0].format.line.fill.background()
    for k in (1, 2, 3):
        ser[k].format.fill.solid()
        ser[k].format.fill.fore_color.rgb = RGBColor.from_string(cols[k])
    if highlight is not None:
        for k in (1, 2, 3):
            pt = ser[k].points[highlight]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = RGBColor.from_string(LIGHT)
    for k, series_vals in ((1, up), (2, down), (3, tot)):
        for i, v in enumerate(series_vals):
            if v:
                dl = ser[k].points[i].data_label
                dl.has_text_frame = True
                txt = f"+{v:g}" if k == 1 else (f"-{v:g}" if k == 2 else f"{v:g}")
                dl.text_frame.text = txt
                r = dl.text_frame.paragraphs[0].runs[0]
                r.font.size, r.font.bold = Pt(14), True
                r.font.color.rgb = RGBColor.from_string(WHITE if k != 3 or not dark else NAVY)
                dl.position = XL_LABEL_POSITION.CENTER
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    ch.category_axis.tick_labels.font.size = Pt(14)
    ch.category_axis.tick_labels.font.color.rgb = RGBColor.from_string(ink)
    ch.category_axis.format.line.color.rgb = RGBColor.from_string(GREY)
    return gf


def matrix2x2(b, s, x, y, w, h, x_label, y_label, quads, points=(), dark=False, highlight=None):
    """2x2 matrix. quads = [top-left, top-right, bottom-left, bottom-right] titles. points = [(label, px, py)] with px
    and py in 0..1 (py 1 = top). highlight = index of the point drawn in ADI light blue. Axis labels outside."""
    ink, sub = _ink(dark), _sub(dark)
    half_w, half_h = w / 2, h / 2
    for i, q in enumerate(quads):
        cx, cy = x + (i % 2) * half_w, y + (i // 2) * half_h
        fill = ("1A4A78" if dark else "DCE8F4") if i == 1 else ("143E68" if dark else "FFFFFF")
        _rect(s, cx + 0.03, cy + 0.03, half_w - 0.06, half_h - 0.06, fill, f"Quad {i+1}")
        b.text(s, cx + 0.15, cy + 0.1, half_w - 0.3, 0.35, q, 14, sub, True, name=f"Quad label {i+1}")
    for i, (lab, px, py) in enumerate(points):
        d = 0.5
        col = LIGHT if highlight == i else BLUE
        cx, cy = x + px * w - d / 2, y + (1 - py) * h - d / 2
        dot = b.dot(s, cx, cy, d, col, name=f"Point {lab}")
        tf = dot.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(lab)[:3]
        r.font.size, r.font.bold = Pt(14), True
        r.font.color.rgb = RGBColor.from_string(WHITE)
    b.text(s, x, y + h + 0.08, w, 0.3, x_label + " (low to high)", 14, sub, True, align=PP_ALIGN.CENTER, name="X axis")
    ya = b.text(s, x - 0.4 - h / 2 + 0.15, y + h / 2 - 0.15, h, 0.3, y_label + " (low to high)", 14, sub, True,
                align=PP_ALIGN.CENTER, name="Y axis")
    ya.rotation = -90
    return s


def gantt(b, s, tasks, periods, x, y, w, row_h=0.5, label_w=3.0, dark=False, highlight=None, period_label="Week"):
    """Gantt bars. tasks = [(name, start, end)] with 1-based inclusive period numbers. periods = count. Light grid."""
    ink, sub = _ink(dark), _sub(dark)
    gw = (w - label_w) / periods
    gx = x + label_w
    for p in range(periods):
        b.text(s, gx + p * gw, y, gw, 0.3, str(p + 1), 14, sub, align=PP_ALIGN.CENTER, name=f"{period_label} {p+1}")
    b.text(s, x, y, label_w, 0.3, period_label, 14, sub, True, name="Period head")
    top = y + 0.4
    for i, (nm, a, z) in enumerate(tasks):
        ry = top + i * row_h
        b.line(s, x, ry + row_h, x + w, ry + row_h, WHITE if dark else BLUE, 0.75, 0.2 if dark else 0.3, name=f"Row rule {i+1}")
        b.text(s, x, ry, label_w - 0.1, row_h, nm, 16, ink, True, anchor=MSO_ANCHOR.MIDDLE, name=f"Task {i+1}")
        _rect(s, gx + (a - 1) * gw + 0.03, ry + 0.1, (z - a + 1) * gw - 0.06, row_h - 0.2,
              LIGHT if highlight == i else BLUE, f"Bar {i+1}", MSO_SHAPE.ROUNDED_RECTANGLE)
    return s


def rating_table(b, s, header, rows, x, y, col_w, max_rating=4, dark=False, row_h=0.55, highlight_row=None):
    """Comparison table where columns after the first are ratings 0..max shown as filled dots (harvey-ball style).
    rows = [(label, r1, r2, ...)]. Native shapes, so every dot is editable."""
    ink, sub = _ink(dark), _sub(dark)
    cx = [x]
    for wv in col_w[:-1]:
        cx.append(cx[-1] + wv)
    for j, hd in enumerate(header):
        b.text(s, cx[j], y, col_w[j], 0.4, hd, 16, sub, True, name=f"Head {j+1}")
    b.line(s, x, y + 0.45, x + sum(col_w), y + 0.45, WHITE if dark else BLUE, 1.0, 0.5, name="Head rule")
    for i, row in enumerate(rows):
        ry = y + 0.55 + i * row_h
        if highlight_row == i:
            _rect(s, x - 0.1, ry - 0.02, sum(col_w) + 0.2, row_h - 0.04, "1A4A78" if dark else "E6EEF6", f"Highlight {i+1}")
        b.text(s, cx[0], ry, col_w[0], row_h - 0.1, row[0], 18, ink, True, anchor=MSO_ANCHOR.MIDDLE, name=f"Row {i+1}")
        for j, v in enumerate(row[1:], 1):
            d, gap = 0.22, 0.08
            for k in range(max_rating):
                b.dot(s, cx[j] + k * (d + gap), ry + (row_h - 0.1 - d) / 2, d, BLUE if k < v else ("3C4157" if dark else "D8DEE6"),
                      name=f"Rating {i+1}.{j}.{k+1}")
        b.line(s, x, ry + row_h - 0.06, x + sum(col_w), ry + row_h - 0.06, WHITE if dark else BLUE, 0.5, 0.15, name=f"Rule {i+1}")
    return s


def funnel(b, s, stages, x, y, w, h, dark=False, highlight=None):
    """Funnel of centred bars narrowing by value, label inside, value right. stages = [(label, value)]."""
    mx = max(v for _, v in stages)
    n = len(stages)
    gap = 0.12
    bh = (h - gap * (n - 1)) / n
    for i, (lab, v) in enumerate(stages):
        bw = max(1.6, w * 0.8 * v / mx)
        bx = x + (w * 0.8 - bw) / 2
        r = _rect(s, bx, y + i * (bh + gap), bw, bh, LIGHT if highlight == i else (BLUE if i % 2 == 0 else NAVY if not dark else "1A4A78"),
                  f"Stage {i+1}")
        tf = r.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = lab
        run.font.size, run.font.bold = Pt(16), True
        run.font.color.rgb = RGBColor.from_string(WHITE)
        b.text(s, x + w * 0.82, y + i * (bh + gap), w * 0.18, bh, f"{v:g}", 24, LIGHT if dark else BLUE, True,
               anchor=MSO_ANCHOR.MIDDLE, name=f"Stage value {i+1}")
    return s
