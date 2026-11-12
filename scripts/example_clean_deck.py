"""Worked Route C example: informal ADI x University of Bristol careers conversation (6 slides + 1 appendix).

    python example_clean_deck.py out.pptx

Techniques shown: click-through builds; a make-room Morph pair (slides 2 to 3: the application list steps aside to
open space for the signal chain); arrows that land with their step; a link to an appendix slide and back.
Copy this file as the starting point for a new conversation deck: keep the structure, replace the content.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from clean_deck import CleanDeck, TC   # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else "clean_deck_example.pptx"
d = CleanDeck(classification="public")

# 1 Cover ---------------------------------------------------------------------------------
d.cover("Analog Devices and the University of Bristol", "Careers in semiconductors: starting a conversation",
        details=["Informal introductory meeting", "[TO CONFIRM: date]  |  [TO CONFIRM: ADI and Bristol attendees]"],
        notes="PURPOSE (30 seconds). An informal first conversation about careers activity, such as a careers talk or event, "
              "and helping students hear about an upcoming internship. Not a proposal; no decision requested. Most of the meeting is discussion.\n\n"
              "Open items: date, attendees, presenters, Bristol contact [TO CONFIRM].")

# 2 Where ADI technology shows up (state A of the make-room pair) ------------------------------------
s2 = d.slide("Where ADI Technology Shows Up")
d.sub(s2, "ADI is a global semiconductor leader that bridges the physical and digital worlds.")
apps = ["Healthcare", "Industrial Automation", "Automotive and Electrification", "Energy", "Communications", "Robotics", "Consumer Technology"]
items = []
for i, a in enumerate(apps):
    col, row = (0, i) if i < 4 else (1, i - 4)
    x, y = 0.46 + col * 6.45, 2.25 + row * 0.75
    items.append([d.m_text(s2, f"app{i}", x, y, 5.9, 0.5, a, 24, TC.ACCENT_1, True),
                  d.m_rule(s2, f"appr{i}", x, y + 0.55, 5.9)])
d.add_footer(s2)
d.build(s2, [items,
             [d.statement(s2, 5.35, "…built by engineers who all started somewhere.", 24)],
             [d.statement(s2, 6.05, "Students can see where engineering theory is used in the real world, and what a career in semiconductors can involve.", 20, TC.TEXT_1, 0.8)]],
        ["The applications ADI technology shows up in", "The engineers behind them", "What this means for students"],
        "KEEP IT SHORT (1 minute). Click 1 lands the seven areas in sequence. Click 2 turns to the people. Click 3 is the point for students. "
        "Then advance: the list steps aside to the left to make room for what those engineers work on.\n\n"
        "Sources: opening line paraphrases the ADI company boilerplate (External Event Guidelines 2024); dated figures deliberately omitted. "
        "Application areas from the meeting brief and older ADI material: confirm against current approved corporate wording [TO CONFIRM].")

# 3 What engineers work on (state B: same named shapes, compact, tucked left) ----------------------------
s3 = d.slide("What Semiconductor Engineers Work On")
d.sub(s3, "ADI's work spans the whole signal chain, from the physical world to the digital one.")
for i, a in enumerate(apps):                       # same keys as slide 2 => Morph glides the list to the left
    d.m_text(s3, f"app{i}", 0.46, 2.3 + i * 0.55, 4.3, 0.4, a, 18, TC.ACCENT_1, True)
    d.m_rule(s3, f"appr{i}", 0.46, 2.72 + i * 0.55, 4.3)
chain = [("Sense", "Capture data in any environment"), ("Measure", "Interpret data for fast, accurate decisions"),
         ("Compute", "Improve data processing efficiency"), ("Connect", "Transfer data reliably"),
         ("Power", "Deliver output with greater efficiency")]
groups = []
for i, (h, t) in enumerate(chain):
    y = 2.3 + i * 0.8
    groups.append([[d.text(s3, 5.5, y, 2.0, 0.6, h, 28, TC.ACCENT_1, True, name=f"Stage {i+1}"),
                    d.text(s3, 7.5, y + 0.1, 5.4, 0.5, t, 18, TC.TEXT_1, name=f"Stage text {i+1}"),
                    d.rule(s3, 5.5, y + 0.66, 7.38, TC.ACCENT_2, name=f"Rule {i+1}")]])
groups.append([d.statement(s3, 6.3, "Many backgrounds and interests fit somewhere on this chain.", 22, h=0.6)])
d.add_footer(s3)
d.build(s3, groups, ["Sense", "Measure", "Compute", "Connect", "Power", "Many routes into the field"],
        "(1 minute). The application list has stepped aside on the left as context. One click per stage, top to bottom; a final click for the closing line. "
        "Ask which stages their students find most interesting.\n\n"
        "Source: 'Who We Are Together' slide, Brunel Lunchtime Careers Talk (2021). Descriptions lightly shortened from the source wording; closing line is our own framing. Confirm against current messaging [TO CONFIRM].")
d.transition(s3, "morph", 1600)

# 4 Ways we could share careers insight -------------------------------------------------------
s = d.slide("Ways We Could Share Careers Insight")
d.sub(s, "Ideas for discussion, not confirmed activities.", TC.ACCENT_3, True)
ideas = ["A careers-in-semiconductors talk", "A presence at a Bristol careers fair", "A meet-an-ADI-engineer session",
         "Graduate and early-career perspectives", "CV and interview tips for engineers", "Joining existing student society events"]
groups = [[d.numbered(s, 0.46 + (i // 3) * 6.45, 2.25 + (i % 3) * 1.25, 5.9, i + 1, t)] for i, t in enumerate(ideas)]
groups.append([d.statement(s, 6.2, "Format, timing and audience would be shaped around what works best for Bristol.", 22)])
d.add_footer(s)
d.build(s, groups, ["Careers talk", "Careers fair", "Meet an engineer", "Early-career perspectives", "CV and interview tips", "Society events", "Shaped around Bristol"],
        "LISTEN FIRST (2 minutes). One click per idea; pause and ask whether it fits how Bristol students already find out about careers. "
        "Nothing here is agreed, scheduled or approved on ADI's side.\n\n"
        "Sources: meeting brief; formats modelled on the Brunel Lunchtime Careers Talk (13 Oct 2021), a UK format reference only. "
        "External speakers, slides or stands need content, brand and legal review and confirmed image rights [TO CONFIRM]. "
        "Collecting student data needs the right privacy notice first [TO CONFIRM with ADI Legal/Privacy].")

# 5 Internship: arrows land with their steps; detail link to the appendix --------------------------------
s5 = d.slide("Helping Students Hear About Our Internship")
d.sub(s5, "Internship details to be confirmed before anything is shared.", TC.ACCENT_3, True)
steps = ["Students hear about ADI", "Talk and meet engineers", "Students find the role", "Apply through ADI"]
groups = []
for i, t in enumerate(steps):
    x = 0.46 + i * 3.15
    item = [d.text(s5, x, 2.2, 0.8, 0.6, f"{i+1:02d}", 28, TC.ACCENT_2, True, name=f"Step number {i+1}"),
            d.rule(s5, x, 2.7, 2.5, name=f"Rule {i+1}"),
            d.text(s5, x, 2.8, 2.6, 0.9, t, 20, TC.ACCENT_1, True, name=f"Step {i+1}")]
    lead = [d.arrow(s5, x - 0.45, 2.5, x - 0.1, 2.5, name=f"Arrow {i}")] if i else []
    groups.append([lead + item])
link = d.text(s5, 0.46, 4.4, 12.4, 0.5, "What students will want to know, and what is still to confirm → appendix", 20, TC.ACCENT_3, True, name="Appendix link")
groups.append([[link]])
groups.append([d.statement(s5, 6.0, "We would only promote roles that are confirmed and approved. Applications would go through ADI's usual careers channels.", 20, h=0.8)])
d.add_footer(s5)
d.build(s5, groups, ["Students hear about ADI", "Talk and meet engineers", "Students find the role", "Apply through ADI",
                     "Details to confirm (links to appendix)", "Only confirmed, approved roles"],
        "(2 minutes). Aim is awareness only. Clicks 1 to 4: how a student might hear about it and apply; each arrow lands with the step it points to. "
        "Click 5 shows the pointer to the appendix: click the line in the show to open the list of details that are still [TO CONFIRM] with ADI Recruiting/HR, then use 'Back' to return. "
        "Click 6: the guardrail. Do not state or imply guaranteed places or a preferential route for Bristol students.\n\n"
        "If the internship is not approved for external promotion, skip click 5 and say only 'an upcoming opportunity we hope to share'. "
        "Vacancy details come only from current approved ADI careers material (careers.analog.com) [TO CONFIRM].")

# 6 Discussion ------------------------------------------------------------------------------------
s = d.slide("What Would Be Most Valuable for Bristol?", dark=True)
qs = ["How do Bristol students currently learn about careers and internships in engineering?",
      "Which courses, year groups or student communities should we think about first?",
      "Would a talk, a careers-fair presence or a meet-an-engineer session be most useful?",
      "What does the academic and careers calendar mean for timing?",
      "What would be a sensible first activity to explore?"]
groups = [[[d.text(s, 0.46, 1.55 + i * 0.88, 0.8, 0.6, f"{i+1:02d}", 28, TC.ACCENT_2, True, brightness=0.4, name=f"Question number {i+1}"),
            d.text(s, 1.4, 1.55 + i * 0.88, 11.4, 0.75, q, 22, TC.BACKGROUND_1, name=f"Question {i+1}")]] for i, q in enumerate(qs)]
groups.append([[d.rule(s, 0.46, 6.1, 12.42, TC.ACCENT_2, name="Rule next"),
                d.text(s, 0.46, 6.25, 12.4, 0.7, "A light next step: identify one possible activity to discuss further. Scope, owners and timing to be confirmed after the meeting.",
                       20, TC.BACKGROUND_1, True, name="Next step")]])
d.add_footer(s)
d.build(s, groups, ["Question 1", "Question 2", "Question 3", "Question 4", "Question 5", "Light next step"],
        "LISTEN (rest of the meeting). One question per click; stop talking after each. Do not steer towards a predetermined ADI programme. "
        "Close with the last click: is there one activity worth exploring? Scope, owners and timing [TO CONFIRM] after the meeting.")

# 7 Appendix (reached from slide 5 by link; skipped in the normal run) -----------------------------------
ap = d.slide("Appendix: Internship Details to Confirm")
d.sub(ap, "Nothing here is final. Every item needs ADI approval before it is shared.", TC.ACCENT_3, True)
facts = ["Role and focus", "Length and start date", "Location", "Who can apply", "How and when to apply", "Approved wording and contacts"]
for i, f in enumerate(facts):
    col, row = divmod(i, 3)
    x, y = 0.46 + col * 6.45, 2.3 + row * 1.2
    d.text(ap, x, y, 5.9, 0.5, f, 24, TC.ACCENT_1, True, name=f"Detail {i+1}")
    d.text(ap, x, y + 0.5, 5.9, 0.4, "[TO CONFIRM]", 18, TC.ACCENT_3, name=f"Detail value {i+1}")
    d.rule(ap, x, y + 0.95, 5.9, name=f"Rule {i+1}")
back = d.text(ap, 0.46, 6.2, 6.0, 0.5, "← Back to the internship slide", 20, TC.ACCENT_2, True, name="Back link")
d.add_footer(ap)
ap.notes_slide.notes_text_frame.text = ("APPENDIX (not part of the main flow). Opened from slide 5 by clicking the link line; use the Back link to return. "
                                         "All six details are [TO CONFIRM] with ADI Recruiting / HR. Do not fill in from older material (the 2021 Brunel talk describes a past scheme).")
d.link(link, ap)
d.link(back, s5)
d.transition(ap, "cut", 0)
ap._element.set("show", "0")   # hidden: skipped in the normal run, still reachable by the link on slide 5

d.finish(OUT)
