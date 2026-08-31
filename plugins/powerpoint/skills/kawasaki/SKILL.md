---
name: kawasaki
description: Guy Kawasaki's 10/20/30 rule for presentations - 10 slides, 20 minutes, 30-point minimum font. Use whenever planning, drafting, structuring, cutting or critiquing a presentation, pitch deck, slide deck or talk, in any tool - not only PowerPoint. Trigger on "make me a deck", "pitch deck", "present this", "slides for", "too many slides", "is this deck too long", "review my presentation", or any request to shape what goes on slides and how much.
---

# The 10/20/30 rule

Guy Kawasaki's rule for presentations, from years of sitting through venture
pitches:

> **10 slides. 20 minutes. 30-point font.**

Apply it to any deck unless the user has said otherwise. Say once, briefly, that
you are working to it; then just do it. Don't lecture.

## The rule, and what each part is really for

### 10 slides — the optimal number

Ten is the number of concepts a normal person can absorb in one sitting. The
constraint is not about slide count for its own sake: it forces you to decide
what the argument actually is. A deck that needs 30 slides usually has no
argument, just material.

Kawasaki's canonical **pitch** order, worth following when the deck is a pitch:

| # | Slide | Carries |
|---|---|---|
| 1 | Title | Who you are, what you do, how to reach you |
| 2 | Problem / opportunity | The pain, described as the customer feels it |
| 3 | Value proposition | The change you produce, not the features |
| 4 | Underlying magic | Why this works and is hard to copy |
| 5 | Business model | Who pays, how much, how often |
| 6 | Go-to-market | How you reach them affordably |
| 7 | Competition | Honest landscape; being alone is a warning sign |
| 8 | Team | Why *these* people |
| 9 | Projections & metrics | The numbers and what drives them |
| 10 | Status & timeline | Where you are, what's next, the ask |

For a non-pitch deck, keep the count and drop the specific headings: one slide
per idea, in the order the argument needs.

**When it runs long:** cut, merge, or move detail into speaker notes and the
appendix. Do not shrink the font — that trades a slide-count problem for a
readability one. If the material genuinely needs more, say so and name what the
extra slides buy; an explicit exception beats silent drift.

### 20 minutes — the maximum speaking time

Book an hour, speak for twenty. The rest goes to discussion, and to the
laptop that won't talk to the projector. Twenty minutes over ten slides is
**two minutes a slide** — which is the real reason a slide cannot hold a
paragraph.

Write what you will **say** into the **speaker notes**, and keep the slide to
the headline. That single habit is what makes the other two numbers achievable:
the words have somewhere to go.

### 30-point font — the minimum size

The floor for body text. Its purpose is not typography — it is a **forcing
function**:

- 30-point text means roughly **six lines a slide**, so you must find the point.
- It guarantees the back row can read it.
- It stops you reading your slides aloud, which is what an audience resents
  most. They read faster than you speak; the moment the slide holds your script,
  they finish early and stop listening.

Kawasaki's alternative if 30 feels arbitrary: **find the oldest person in the
audience, halve their age, and use that** — a 60-year-old investor gets 30-point
text.

**Watch the sub-bullets.** Nearly every template shrinks each outline level:
the stock Office template runs 32 / 28 / 24 / 20 point, so a level-2 bullet
breaks the rule *without anyone choosing a font*. Prefer top-level bullets only.

## Applying it

**When drafting**, work in this order — outline first, slides last:

1. Write the **argument** as ten one-line assertions. Each becomes a slide
   title. If you can't get to ten, the deck isn't ready; if you can't get under
   twenty, the scope is too big.
2. Make each title a **claim**, not a label: "Unpriced risk costs us £4m a year",
   not "Risk". The title is the only line most of the room reads.
3. Put at most **three to six short bullets** under it — evidence for the claim,
   not sentences you will read out.
4. Put the **sentences you will actually say** in the speaker notes. Roughly
   250 words per slide is two minutes.
5. Only then build the file.

**When critiquing an existing deck**, report against all three numbers and say
which slides to cut and why. Be specific — "slides 7, 9 and 12 all restate the
value proposition; keep 7" beats "consider tightening".

## With the `powerpoint` MCP server

If the `powerpoint` tools are available, the rule is **measurable**, not just
advice. Use the [`powerpoint`](../powerpoint/SKILL.md) skill for the mechanics,
and:

- **`powerpoint_list_layouts`** *before* writing anything: `level1_font_pt` and
  `deeper_levels` tell you which layouts can hold 30-point text, and
  `layouts_meeting_min_font_at_level_1` lists them outright. Choosing the right
  layout is how you obey the 30-point rule without ever setting a font.
- **`notes`** on every slide — one `powerpoint_add_slides` entry per slide, each
  carrying its own. Without notes the 20-minute estimate has nothing to measure
  and the review says so.
- **`powerpoint_review`** before saving, and again after every trim. It resolves
  each run's *effective* font size through the template inheritance chain, so it
  catches the inherited 28-point sub-bullet that reading the file would miss.
  It also flags slides carrying more than 40 words of body text — a slide can
  pass on font size and still be a document.
- **`powerpoint_delete_slide`** to get back to ten. Delete from the **highest**
  index downwards.

Never fix a font finding by overriding the font — that breaks the template.
Fix it by saying less, nesting less, or choosing a different layout.

## The thresholds are configurable

10 / 20 / 30 are constants in `powerpoint.py`
(`KAWASAKI_MAX_SLIDES`, `KAWASAKI_MAX_MINUTES`, `KAWASAKI_MIN_FONT_PT`). If a
house style differs — a 24-point floor, a 15-minute slot — change them there and
every review follows. Don't quietly apply different numbers in your head.

## When not to apply it

The rule is aimed at **persuasive** presentations to an audience: pitches,
proposals, board updates, conference talks. It is a poor fit for a document
someone will read alone at their desk, a detailed technical walkthrough, or a
training deck meant as a reference afterwards. When the user asks for one of
those, say the rule doesn't fit and why, then write what they actually need —
and consider whether a document would serve them better than slides.
