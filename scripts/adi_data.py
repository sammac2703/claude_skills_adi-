"""Data blocks for ADI decks: paginated tables and ADI-coloured native charts.

Ideas borrowed from PptxGenJS (autoPage with repeated header rows; chart option sets), rebuilt on python-pptx so they
work on the ADI template and Slide Library layouts.

    from adi_data import paginate_table, add_chart
    slides = paginate_table(b, "2_Title Only – Blank", "Appendix: Plan", header, rows, [0.8, 2.4, 3.3, 3.0, 2.9])
    add_chart(slide, "column", ["Q1", "Q2"], {"2025": [3, 4], "2026": [5, 6]}, 0.46, 1.9, 8, 4.6, dark=False)

Rules kept: 14 pt floor for table text and chart labels, ADI palette only, blues first (one accent), native objects.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from adi_blocks import BLUE, CHAR, GREY, LIGHT, NAVY, WHITE   # noqa: E402
from pptx.chart.data import CategoryChartData               # noqa: E402
from pptx.dml.color import RGBColor                         # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION  # noqa: E402
from pptx.enum.text import MSO_ANCHOR                        # noqa: E402
from pptx.util import Inches, Pt                             # noqa: E402

SERIES = [BLUE, LIGHT, NAVY, GREY, "8637BA", "179963"]       # ADI theme colours, blues first
ZEBRA = "EEEFF2"
KINDS = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
         "bar": XL_CHART_TYPE.BAR_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE,
         "doughnut": XL_CHART_TYPE.DOUGHNUT, "area": XL_CHART_TYPE.AREA}


def _rgb(h):
    return RGBColor.from_string(h)


def rows_per_page(top, max_y, row_h, header_h=None):
    """How many body rows fit between top and max_y when the header row repeats on every page."""
    header_h = row_h if header_h is None else header_h
    return max(1, int((max_y - top - header_h) // row_h))


def _lines(text, width_in, size):
    """Rough wrapped line count for a cell (Barlow averages about 0.5 em per character)."""
    per_line = max(1, int(width_in * 72 / (size * 0.5)))
    return sum(max(1, math.ceil(len(p) / per_line)) for p in str(text).split("\n"))


def paginate_table(b, layout, title, header, rows, col_w, x=0.46, top=1.95, max_y=6.85, size=14, min_row_h=0.38,
                   dark=False, key_title=True, notes=None, footer=True, brand=True):
    """Build one or more slides holding a native table; rows that do not fit move to a new slide with the header row
    repeated and '(2 of 3)' added to the title. Row heights grow for wrapped cells, so no row is split over slides.
    Returns the list of slides (footer, logo and slide number added; the caller sets transitions and notes)."""
    ink, head_fill, band = (WHITE, BLUE, None) if dark else (NAVY, BLUE, ZEBRA)
    heights = [max(min_row_h, max(_lines(c, col_w[i] - 0.2, size) for i, c in enumerate(r)) * size * 1.25 / 72 + 0.14)
               for r in rows]
    head_h = max(min_row_h, max(_lines(c, col_w[i] - 0.2, size) for i, c in enumerate(header)) * size * 1.25 / 72 + 0.14)
    pages, cur, used = [], [], head_h
    for r, h in zip(rows, heights):
        if cur and top + used + h > max_y:
            pages.append(cur)
            cur, used = [], head_h
        cur.append((r, h))
        used += h
    pages.append(cur)
    out = []
    for n, page in enumerate(pages, 1):
        t = title if len(pages) == 1 else f"{title} ({n} of {len(pages)})"
        s = b.slide(layout, t, key_title=key_title)
        total_h = head_h + sum(h for _, h in page)
        gs = s.shapes.add_table(len(page) + 1, len(header), Inches(x), Inches(top), Inches(sum(col_w)), Inches(total_h))
        gs.name = "Table"
        tbl = gs.table
        for ci, w in enumerate(col_w):
            tbl.columns[ci].width = Inches(w)
        tbl.rows[0].height = Inches(head_h)
        for ri, (_, h) in enumerate(page, 1):
            tbl.rows[ri].height = Inches(h)
        for ri, row in enumerate([header] + [r for r, _ in page]):
            for ci, val in enumerate(row):
                c = tbl.cell(ri, ci)
                c.text = str(val)
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
                c.fill.solid()
                c.fill.fore_color.rgb = _rgb(head_fill if ri == 0 else (band if band and ri % 2 else WHITE if not dark else NAVY))
                for p in c.text_frame.paragraphs:
                    for r_ in p.runs:
                        r_.font.size = Pt(size)
                        r_.font.bold = ri == 0
                        r_.font.color.rgb = _rgb(WHITE if ri == 0 else (WHITE if dark else ink))
        b.tidy(s)
        if footer:
            b.footer(s)
        if brand:
            b.brand(s, dark)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        out.append(s)
    return out


def add_chart(slide, kind, categories, series, x, y, w, h, dark=False, number_format="General", labels=True,
              legend=None, title=None, size=14, name="Chart"):
    """Native chart in ADI colours. kind: column, stacked, bar, line, pie, doughnut, area. series: {name: [values]}.
    Labels and axes 14 pt. Pie and doughnut take one series. Combo charts are not supported by python-pptx: place two
    charts or use a line over columns drawn by the user in PowerPoint. Returns the graphic frame."""
    cd = CategoryChartData()
    cd.categories = categories
    for nm, vals in series.items():
        cd.add_series(nm, vals)
    gf = slide.shapes.add_chart(KINDS[kind], Inches(x), Inches(y), Inches(w), Inches(h), cd)
    gf.name = name
    ch = gf.chart
    ink = WHITE if dark else CHAR
    ch.font.size = Pt(size)
    ch.font.color.rgb = _rgb(ink)
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        r = ch.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size, r.font.bold, r.font.color.rgb = Pt(size + 2), True, _rgb(ink)
    multi = len(series) > 1 or kind in ("pie", "doughnut")
    ch.has_legend = multi if legend is None else legend
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(size)
        ch.legend.font.color.rgb = _rgb(ink)
    plot = ch.plots[0]
    if kind in ("pie", "doughnut"):
        pts = plot.series[0].points
        for i in range(len(categories)):
            pts[i].format.fill.solid()
            pts[i].format.fill.fore_color.rgb = _rgb(SERIES[i % len(SERIES)])
    else:
        for i, s in enumerate(plot.series):
            col = _rgb(SERIES[i % len(SERIES)])
            if kind == "line":
                s.format.line.color.rgb = col
                s.format.line.width = Pt(2.5)
                s.marker.format.fill.solid()
                s.marker.format.fill.fore_color.rgb = col
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = col
        if kind in ("column", "bar", "stacked"):
            plot.gap_width = 60
        ch.value_axis.has_major_gridlines = True
        ch.value_axis.major_gridlines.format.line.color.rgb = _rgb("3C4157" if dark else "D8DEE6")
        ch.value_axis.format.line.fill.background()
        ch.value_axis.tick_labels.font.size = Pt(size)
        ch.value_axis.tick_labels.font.color.rgb = _rgb(ink)
        ch.value_axis.tick_labels.number_format = number_format
        ch.value_axis.tick_labels.number_format_is_linked = False
        ch.category_axis.tick_labels.font.size = Pt(size)
        ch.category_axis.tick_labels.font.color.rgb = _rgb(ink)
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(size)
        dl.font.bold = True
        dl.number_format = number_format
        dl.number_format_is_linked = False
        dl.font.color.rgb = _rgb(WHITE if kind in ("stacked", "doughnut") or (kind == "pie") else ink)
        if kind in ("column", "bar"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
    return gf
