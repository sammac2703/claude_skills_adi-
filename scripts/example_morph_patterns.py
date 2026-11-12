"""Demo of the Route C Morph patterns learned from the Gold deck. Build and click through to see each one.

    python example_morph_patterns.py out.pptx

1. make-room pair      two slides share named shapes; on slide B the list shrinks left and new content enters
2. focus row           accordion / spotlight, one slide per item, text only
3. arrows + paragraphs interleaved builds (an arrow lands with the step it points to; paragraphs mix in)
4. slide links         jump to an appendix slide and back (safe stand-in for Slide Zoom)
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from clean_deck import CleanDeck, TC   # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else "morph_patterns_demo.pptx"
d = CleanDeck()
d.cover("Morph Patterns", "Demo of the Route C state-pair helpers", details=["Example only"])

# 1 make-room pair ---------------------------------------------------------------------------
items = ["First area", "Second area", "Third area", "Fourth area"]
a = d.slide("Make Room: Before")
for i, t in enumerate(items):
    d.m_text(a, f"item{i}", 0.46, 2.0 + i * 0.9, 5.9, 0.6, t, 26, TC.ACCENT_1, True)
    d.m_rule(a, f"itemr{i}", 0.46, 2.65 + i * 0.9, 5.9)
d.add_footer(a)
b = d.slide("Make Room: After")
for i, t in enumerate(items):                       # same keys, smaller, tucked left
    d.m_text(b, f"item{i}", 0.46, 2.0 + i * 0.55, 3.8, 0.4, t, 18, TC.ACCENT_1, True)
    d.m_rule(b, f"itemr{i}", 0.46, 2.45 + i * 0.55, 3.8)
d.park(b, "detail", "text", edge="right", y=2.0, w=7.0)      # parked on A so it can enter from the right
d.m_text(b, "detail", 5.0, 2.0, 7.5, 1.2, "New content arrives here while the list steps aside.", 24, TC.TEXT_2, True)
d.park(a, "detail", "text", edge="right", y=2.0, w=7.0)
d.add_footer(b)
d.transition(b, "morph", 1600)

# 2 focus row ----------------------------------------------------------------------------------
d.focus_row("Focus Row", [("Sense", "Capture data in any environment"), ("Measure", "Interpret and analyse data"),
                          ("Compute", "Improve processing efficiency"), ("Connect", "Transfer data reliably")],
            intro="Each slide widens one item; the others narrow.")

# 3 arrows + paragraph interleave, with a link to an appendix --------------------------------------------
s = d.slide("Arrows and Paragraph Builds")
steps = ["Hear", "Meet", "Find", "Apply"]
groups = []
shapes = []
for i, t in enumerate(steps):
    x = 0.46 + i * 3.15
    it = d.numbered(s, x, 2.2, 2.8, i + 1, t, 22)
    shapes.append(it)
    groups.append([([d.arrow(s, x - 0.35, 2.55, x - 0.05, 2.55)] if i else []) + it])
para = d.text(s, 0.46, 4.4, 12.4, 1.6, ["First line of a paragraph build.", "Second line.", "Third line."], 20, TC.TEXT_1, name="Lines")
groups.append([[(para, 0)], [(para, 1)], [(para, 2)]])
more = d.text(s, 0.46, 6.3, 6, 0.4, "Details in the appendix →", 18, TC.ACCENT_3, True, name="Appendix link")
groups.append([[more]])
d.add_footer(s)
d.build(s, groups, steps + ["Paragraphs", "Link"], "Arrows land with the step they point to; paragraphs interleave.")
app = d.slide("Appendix: Detail")
d.text(app, 0.46, 2.0, 12, 0.6, "Appendix content goes here.", 22)
back = d.text(app, 0.46, 6.3, 6, 0.4, "← Back to the slide", 18, TC.ACCENT_3, True, name="Back link")
d.link(more, app)
d.link(back, s)
d.add_footer(app)
d.transition(app, "cut", 0)

d.finish(OUT)
