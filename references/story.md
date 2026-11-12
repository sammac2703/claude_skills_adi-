# Story layer: how the pitch is structured

Contents: 1 Reference arc · 2 ADI slide sequence · 3 Story techniques · 4 ADI adaptations · 5 Audience tuning · 6 Delivery (speaker notes) · 7 Inputs · 8 Evidence base

## 1. The reference arc (iPhone, Macworld 2007)

Jobs sells one idea in nine beats; the product name lands under three minutes in.

| # | Beat | What he does | Job it does |
|---|---|---|---|
| 1 | Hook | Says he has waited two and a half years for this day | Raises the stakes before any content |
| 2 | Track record | Macintosh (1984), iPod (2001): one product that changed an industry, each time | Earns the right to claim another |
| 3 | Tease | Three products: widescreen iPod, phone, internet communicator, repeated as a trio | Builds anticipation and perceived value |
| 4 | Reveal | "Are you getting it?" The three are one device | The moment everyone remembers |
| 5 | Release valve | Joke iPod with a rotary dial | Laughter after the peak resets attention |
| 6 | Problem | Smartphones are not so smart and not so easy to use | Names a pain the room already feels |
| 7 | Positioning | 2x2 of smart vs easy to use; rivals in one corner, iPhone alone top-right | Makes the gap visible in one picture |
| 8 | Solution | Fixed plastic keyboards are the root cause; a big screen and fingers are the answer | The feature arrives as the answer |
| 9 | Proof | Runs OS X; desktop-class software; a quote from an authority | Backs the claim when the room already wants to believe |

The pattern: credibility and tease before the product, problem before feature, proof last.

## 2. The ADI slide sequence

This is what the skill generates for a customer pitch. Roles in brackets are the lint roles.

| # | Slide (role) | Layout(s) | Content |
|---|---|---|---|
| 1 | Hook (hook) | `mosaic`, `title` | The customer's goal in their words: battery life, noise floor, board space, time to market |
| 2 | Credibility (credibility) | `credibility` | One relevant, cleared-to-share win or reference design in the same application |
| 3 | Objections, optional (objections) | `objections` | For sceptical or post-escalation rooms: name the concern before they do |
| 4 | Problem (problem) | `statement`, `status_quo` | The pain quantified, ideally with their numbers; cost of the status quo built row by row |
| 5 | Criteria (criteria) | `criteria` | Two or three yardsticks the customer agrees matter, before any product appears |
| 6 | Positioning (positioning) | `quadrant`, `bars` | Today's options on the customer's two axes, target corner empty, then filled |
| 7 | Tease (tease) | `tease` | The three wins they want, shown as separate things |
| 8 | Reveal (reveal) | `reveal` | One part or signal chain delivers all three, with a one-sentence definition |
| 9+ | How it works (feature) | `feature`, `product`, `hardware` | One feature per slide, each answering a problem already shown (`answers` field) |
| n | Proof (proof) | `bars`, `big_number`, `side_by_side`, `hardware` | Measured data, a headline number, the eval board, a plot, a demo |
| n+1 | Partners, optional (partners) | `partners` | Reference-design partner, ecosystem tool, cleared customer quote |
| n+2 | Plan, optional (agenda) | `roadmap` | Evaluate, design review, prototype, production |
| last | Ask (ask) | `ask` | The specific next step, owner and date |
| app | Appendix (appendix) | `table`, anything | Datasheet tables, test conditions, full comparisons, backup plots |

Size: 8 to 14 main-flow slides (lint warns above 14) for a 30 to 45 minute meeting, leaving at least a third of the time for questions. Build steps (extra Morph state slides) do not count; they are clicks, not ideas.

## 3. Story techniques

**Building anticipation**
1. Personal stakes first: open on why this meeting matters to them.
2. Borrow credibility from past wins, framed the same way each time, so the audience completes the pattern.
3. Decoy of three: present the value as separate wins the customer already wants, then show they arrive in one solution (`tease` -> `reveal`; the three circles Morph into dots under the product name).
4. Repeat the trio in the notes until the room half-guesses the twist; land it with a question.
5. Release tension after the peak with a light moment before denser material.

**Problem framing and positioning**
- Define, then deflate: what the current approach is, then its weakness.
- Show, don't argue: a 2x2 or a comparison puts options where you say they are and leaves the target space empty.
- Root cause, not symptom (fixed plastic buttons, not "bad phones"; three separate parts in the measurement path, not "high power").
- Pre-empt the obvious alternative in one line (Jobs: "who wants a stylus?"; ADI: "why not just a lower-power MCU?").
- Set the criteria before the product, so the audience agrees the yardstick and then watches one option pass.
- One comparison sentence the audience can check at a glance. MacBook Air: its thickest point is thinner than the rival's thinnest.
- Retire your own old answer before a rival does. If the customer uses an older ADI part, say plainly why the new approach is better.

**Features into benefits**
- Problem before feature: the audience wants the answer before they hear its name.
- Analogy to something proven (the mouse made the Mac, the click wheel made the iPod).
- Jargon only after the benefit; technical depth comes as proof.
- Translate the number into a unit the customer counts in:

| Spec | Translated into |
|---|---|
| 5 GB (iPod) | 1,000 songs in your pocket |
| FireWire (iPod) | A whole CD in under 10 seconds |
| Architecture (2005) | Performance per watt: 15 vs 70, one bar chart |
| 326 ppi (iPhone 4) | Beyond what the eye resolves: text like a printed page |
| Quiescent current (ADI) | Days of battery on their cell |
| Noise density (ADI) | Averaging cycles saved, so power saved, so battery life |
| Integration (ADI) | Parts removed from the BOM, mm² of board freed, calibration steps deleted |
| Accuracy over temperature (ADI) | No field recalibration; one trim at end of line |

- Tie the number to a human or system limit the audience recognises (Retina vs the eye). For ADI: noise floor vs the signal they need, power vs their battery, drift vs their calibration interval.
- Use the chain spec -> system benefit -> business benefit: lower noise -> fewer averaging cycles -> lower power -> longer battery life -> fewer returns.

**Credibility and the close**
- Candour buys credibility: falling revenue (1997), the missed 3 GHz G5 (2005), "not our finest hour" (2011).
- Anchor price against what they already spend (the sum of today's parts, as in `status_quo` with `ours`).
- Partners as validation, only where cleared.
- A clear ask: what, who, by when. "Evaluate" is not an ask; "eval board on your bench by the 13th, review call on the 27th" is.
- Prove readiness by already using it: a demo running on the real eval board, firmware already ported.

## 4. ADI adaptations

| Jobs move | ADI equivalent | Watch out for |
|---|---|---|
| Personal stakes hook | The customer's goal in their words | Keep it about them, not ADI |
| Track record | A relevant win in the same application, or a reference design | Only cite wins cleared to share |
| Three products in one | Functions they now buy separately, delivered in one signal chain or device | Name only functions the part actually integrates |
| "Not so smart" problem | Quantify the pain: power, BOM count, area, calibration effort | Use their numbers where you have them |
| 2x2 positioning | The two axes this customer cares most about | Datasheet-backed, like-for-like data; check ADI rules on naming competitors |
| Feature -> benefit | Spec -> system benefit -> business benefit | Engineers will ask typical vs max, and over what temperature |
| "5 years ahead" | One headline number per slide, measured data in the appendix | Unsupported superlatives cost credibility |
| Live demo | Eval board, LTspice simulation, measured plots, the physical part | Rehearse; have a recorded fallback |

Dial down theatre and length: a customer meeting is a conversation.

### When the evidence is not there yet

Most early decks are missing something. Decide per slide, and say what you decided in the delivery note:
- **No cleared reference win:** drop the credibility slide. Do not write a vague one.
- **No like-for-like competitor data:** either drop the quadrant, or plot qualitative positions with axis labels only (no numbers) and `SOURCE: [TBC: ...]` in the notes. Lint flags the [TBC] source as REVIEW, which means "not presentable until filled".
- **No measured proof, only datasheet typicals:** say so on the slide or in the kicker ("Datasheet, typical") and in CONDITIONS. Lint treats any `proof` slide as proof; you are the one who has to keep "datasheet typical" from being presented as "measured in your system". If the only "proof" is a goal or target, give that slide `"role": "feature"`, not proof, so the superlative check stays on.
- **No product image:** use `glyph` on reveal and feature slides rather than leaving dashed placeholder boxes in a deck that will be shown.

### Where objections go

Two valid places, pick by the room:
- **Slot 3, before the problem:** when the room arrives sceptical (an escalation, a lost socket, a known issue with ADI). Naming it first clears the air so they listen to the rest.
- **After the proof, before the roadmap:** for a normal room, where the objections are technical ("the sensor drifts too", "24 bits is a datasheet number"). Answering them right after the proof is where candour buys most credibility.
Lint does not check which you chose.

## 5. Audience tuning

- **Design engineers:** move faster to the hardware and proof slides; expect interruption; have the appendix table and conditions ready; show the eval board and schematic.
- **Architects / technical leads:** spend longer on criteria and positioning; they choose the approach before the part.
- **Procurement / management:** lead with cost of status quo, supply and roadmap, risk reduction; keep specs to one translated number each.
- **Mixed room:** run the arc for the decision-makers, park deep questions with "let's take that in the appendix after", and keep a backup slide for each predictable question.
- **Remote meetings:** fewer builds (they stutter on screen share), larger type, check the deck at the shared resolution; send the static version afterwards.

## 6. Delivery (goes into speaker notes)

The notes fields map to these habits: `say`, `ask` (question to the room), `pause_before`, `pause_after`, `conditions`, `source`, `fallback`, `transition`.

- Rule of three; repeat the one theme word like a drumbeat.
- Questions to the room ("Are you getting it?", "Did we get that right?").
- Plain words for hard ideas; superlatives rationed until after the evidence.
- Pause before reveals and let reactions finish.
- Pacing: calm start, rising to the reveal, then a steadier analytical pace.
- Recover calmly: when the clicker died in 2007, Jobs told a story until it was fixed. Every demo slide gets a `fallback` line.
- Write notes as speech, not prose: short sentences the presenter can glance at.

## 7. Inputs

Ask for, in one grouped message:
- Customer, application, who is in the room (design engineer, architect, procurement)
- Their goal and current pain, with numbers if known
- The two or three criteria they judge options by
- The ADI part or signal chain being proposed
- Two or three headline specs, each with conditions (typ vs max, temperature, supply, mode)
- What they use today; competitor names (flagged for review)
- Proof available: eval board photo, measured plot, simulation, physical part, demo
- Optional: reference win (cleared to share), partner or ecosystem tool, known objections
- The ask: samples, eval board, design review, with a date
- Assets: application photo, pre-cut transparent product PNGs, icons, eval board photo
- Theme: dark, light or template default

## 8. Evidence base

Keynotes analysed: Macworld Boston 1997 (name the objection first), iMac 1998 (show the real thing, then walk around it), iPod 2001 (spec into a unit the customer counts in; status-quo table built row by row), WWDC 2005 Intel switch (one metric, one bar chart; prove readiness by already using it), iPhone 2007 (the arc above; price anchored to what they already buy), MacBook Air 2008 (one checkable comparison sentence), iPad 2010 (criteria before reveal), WWDC 2010 iPhone 4 (tie the number to a human limit), WWDC 2011 iCloud (retire your own old answer).

Caveats: video observations came from a model watching the videos, timestamps approximate; some uploads were excerpts.
