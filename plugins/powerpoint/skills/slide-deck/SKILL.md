---
name: slide-deck
description: Build a sectioned PowerPoint deck end to end via the powerpoint MCP server - pick up the requested template, open with a title slide, introduce every section with its own divider slide, title each slide inside a section "<Short Section Title> / <Slide Title>", carry the evidence as bullets, and give every slide legible speaker notes written as bold-labelled dot points. Aims at Guy Kawasaki's 10/20/30 rule without treating it as a hard limit. Use when the user asks for a deck, slides, a presentation or a .pptx, especially one with sections, chapters or parts.
---

# Slide deck

The house shape for a deck built with the `powerpoint` MCP server: a title
slide, a divider in front of every section, slides named for the section they
sit in, and speaker notes a presenter can actually read at a lectern.

Three skills, three jobs. This one decides **what the deck is** — the sections,
the running order, the titles, the notes. [`powerpoint`](../powerpoint/SKILL.md)
is the server's mechanics and the sharp edges. [`kawasaki`](../kawasaki/SKILL.md)
is the 10/20/30 rule itself and when it does not apply. Read the mechanics one
before building; this skill assumes it.

If the `powerpoint` tools are not available, say so and stop — this skill is
about driving that server, not about drafting slide text into chat.

## 1. Settle the template first

A deck that ignores the corporate template is a rewrite, so resolve this before
writing a word.

1. `powerpoint_list_presentations` with `location: "templates"`.
2. Then:
   - **The user named one** → pass it straight to `powerpoint_create` as
     `template`. Names resolve forgivingly, so *"the board template"* finds
     `Board Deck Template.pptx`. The result's `template` field says which file
     was actually taken, and `template_fuzzy_matched: true` means it was a
     near-miss rather than an exact name — **confirm that one with the user
     before building**, because the template decides the whole deck's look.
   - **Exactly one template exists** → use it, and say which.
   - **Several exist and the user named none** → ask which. Do not guess: the
     template is the one decision that is expensive to unwind.
   - **None exist** → say the deck will use the stock Office template, and
     carry on. Do not stall.
3. `powerpoint_create` with `filename` and `template`, then
   **`powerpoint_list_layouts`** — never skip it. A branded template names its
   layouts whatever it likes, and this skill needs three roles to exist:
   `title`, `section` and `bullets`.

**Check the template really has a divider layout.** `recommended` (on
`powerpoint_list_layouts`, and as `recommended_layouts` on `powerpoint_create`)
carries a key per role the template offers, so a missing `section` key means
there is no divider layout. Do not pass `layout: "section"` and hope — with no
section role it falls through to a fuzzy name match and can land anywhere. Use
`title_only` instead, or `title` if there is no `title_only` either, and tell
the user which you fell back to. Never fake a divider by leaving a content
slide's body empty.

## 2. The shape

```
1   Title slide                     title       no section prefix
2   Section divider                 section     the section's FULL title
3-5   Section slides                bullets     "<Handle> / <Assertion>"
6   Section divider                 section
7-9   Section slides                bullets
...
n   The ask / what happens next     bullets     the last slide of the last section
```

Budget it like this, and say what you are aiming at once, briefly:

- **2 to 4 sections.** Three is usually right. One section is not a sectioned
  deck — drop the dividers and build it flat.
- **2 to 4 content slides per section.** A section with one slide under it is
  not a section: fold that slide into its neighbour.
- **About 10 content slides all up.** The title slide and the dividers are
  scaffolding and do not spend the budget (see §6).
- **One close.** The last slide carries the ask or the next step, not a
  "Thank you" slide.

## 3. Titles: `<Short Section Title> / <Slide Title>`

Every slide **inside** a section carries its section's handle, a space, a
forward slash, a space, then its own title:

```
Pricing / Unpriced risk costs us $4m a year
Pricing / Broker claims drive three quarters of it
Rollout / Priced at bind from March
```

- Pick the **handle** when you name the section and reuse it **verbatim** on
  every slide in that section. One or two words, ideally under 15 characters.
  `Pricing`, `Rollout`, `The ask`, `Where we are`.
- The **divider** slide carries the section's full title, unprefixed:
  *"Pricing: what the last twelve months cost us"*. The handle is the short
  form of that, not a separate idea.
- The **title slide** and any **appendix** slide take no prefix.
- Keep the whole prefixed title to roughly **60 characters** so it stays on one
  line. If it will not fit, shorten the handle first, then the assertion.
- The part after the slash is still an **assertion, not a label** —
  *"Pricing / Claims up 20%"*, never *"Pricing / Claims"*. The title is the
  only line most of the room reads.

## 4. Bullets: where they earn their place

Bullets are the default for a content slide, not an obligation.

- **Three to six** per slide, each a short phrase carrying evidence for the
  title's claim. Never a sentence you will read aloud.
- **Top level only** wherever you can. Every template shrinks each outline
  level, so a level-2 bullet usually breaks the 30-point rule without anyone
  choosing a font. Nest with the `level` field only when the structure is real.
- **Never type `- `, `* `, `•` or `1. `** into slide bullet text. The layout
  draws the marker. (The notes pane is the opposite case — see §5.)
- **Drop them entirely** when the slide is better as a table
  (`table: {"rows": [...]}` with `bullets` left out), a single number, or one
  sentence. A divider slide takes no bullets at all.

## 5. Speaker notes: on every slide, without exception

Every slide gets notes — the title slide and the dividers included. The slide
carries the headline, the notes carry the argument, and that is what lets the
slide stay a headline at 30 points.

The `notes` field renders a small Markdown-like syntax as **real formatting**
in the notes pane, so write for a presenter glancing down mid-sentence, not for
a reader:

| You write | The pane shows |
|---|---|
| `**Evidence:**` | a **bold** label |
| `__$4m__` | an <u>underlined</u> figure |
| `- Claims up 20%` | a dot point |
| `  - mostly EMEA` | a nested dot point (two spaces per level) |
| *(a blank line)* | an empty line between blocks |
| `\*` `\_` | a literal asterisk or underscore |

**Unlike slide bullets, DO type the `- `.** PowerPoint's notes master draws no
bullet of its own, so the server supplies the glyph.

The house pattern — **bold label, dot points under it, blank line between
blocks.** Prose belongs in notes only where an exact form of words matters:

```
**Say**
- Unpriced risk is costing us __$4m__ a year

**Evidence**
- Claims up __20%__ year on year
  - three brokers, all EMEA
- Nothing repriced since __2023__

**Ask**
- Sign off repricing at bind, effective __March__
```

For a **divider**, the notes are the transition:

```
**Moving to: Pricing**
- What the last twelve months actually cost us
- Three slides, then the ask
```

For the **title slide**, the frame and the housekeeping:

```
**Open**
- Thanks for the time - twenty minutes, then questions

**Frame**
- One decision today: whether we reprice at bind
```

Rules that keep the pane legible:

- **Bold only the labels** and the two or three words that must land. Bold
  everything and nothing is bold.
- **Underline figures, dates and names** — the things that are embarrassing to
  misread aloud.
- **A blank line between every block.** White space is what makes a glance work.
- **Six to ten dot points a slide.** Past that you are writing a document.
- Pass `notes_format: "plain"` for a note that genuinely contains `__init__`
  or a literal `**`, which switches the syntax off for that slide.

## 6. 10/20/30, aspirationally

Work to the rule, say once that you are, and do not let it drive the deck off a
cliff. [`kawasaki`](../kawasaki/SKILL.md) has the reasoning.

- **10 slides** applies to **content** slides. A three-section deck runs to
  roughly 14 slides all up once the title and the dividers are counted, and
  that is fine. `powerpoint_review` counts every slide, so it will report 14
  against 10: read that finding against your content count and **say so** to
  the user rather than deleting dividers to make a number go green.
- **20 minutes** still binds, and the dividers spend it too. Watch the
  `estimated_minutes` in `powerpoint_review`.
- **30 points** is not aspirational. Never fix a font finding by setting a
  font — there is no tool for it, and that is deliberate. Fix it with a
  different layout, fewer words, or less nesting.

**When it runs long,** cut a section before you cut within one: three tight
sections beat five thin ones. Move detail into the notes, or into slides after
the close that you will not present.

**Flag the exception rather than drifting into it.** If the material genuinely
needs 16 content slides, say so and name what the extra slides buy.

## 7. Build it, then review it

Order matters more than it looks. Follow [`powerpoint`](../powerpoint/SKILL.md):

1. Draft the **whole outline first** — sections, handles, every prefixed title,
   the bullets and the notes. Put it to the user before building if the deck is
   for an audience that matters.
2. **ONE `powerpoint_add_slides` call** carrying every slide in order, each with
   its own `notes`. Never a series of `powerpoint_add_slide` calls: they append
   in whatever order they arrive at the server, which comes out as a shuffled
   deck.
3. **`powerpoint_review`**, and report it honestly — including the slide-count
   finding and what it is really counting.
4. **`powerpoint_save`.**

Then tell the user, in a couple of lines: which template it was built from, how
many content slides against how many total, the estimated minutes, and anything
the review flagged that you chose not to fix.

## What this skill is not for

A read-at-your-desk document, a detailed technical walkthrough or a training
reference is not a 10/20/30 deck, and sections plus dividers make it longer
rather than clearer. Say the shape does not fit, then build what the user
actually needs — and consider whether a Word document would serve them better.
