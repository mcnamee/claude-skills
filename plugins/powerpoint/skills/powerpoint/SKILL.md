---
name: powerpoint
description: Build PowerPoint .pptx decks via the powerpoint MCP server, inheriting a corporate template's masters, layouts, theme, fonts and colours; add slides onto named layouts, fill placeholders, write speaker notes, add tables, reorder and delete slides, and audit the result against the 10/20/30 rule. Use when the user asks to create, build, edit, read or review a PowerPoint deck or .pptx file, to make slides from a template, or to check a deck's font sizes, slide count or length.
---

# PowerPoint (via the `powerpoint` MCP server)

Requires the `powerpoint.py` MCP server (sandboxed to its configured
presentations folder). If its tools are not available, tell the user to wire it
in first (see the repo README) and to verify with
`python powerpoint.py --check`.

For *what goes on the slides* — how many, how long, how few words — follow the
[`kawasaki`](../kawasaki/SKILL.md) skill. This one is the mechanics.

## Core workflow: create → layouts → slides → review → save

1. **Find the template.** `powerpoint_list_presentations` with
   `location: "templates"`. If the user named one, pass it straight to
   `powerpoint_create` — names resolve forgivingly (bare name, relative path,
   or a fuzzy near-miss like *"acme deck"* → `Acme Deck Template.pptx`).
2. **`powerpoint_create`** with `filename` and, when there is one, `template`.
   Returns the `session_id` every other call needs. Example slides in the
   template are stripped; its design is kept.
3. **`powerpoint_list_layouts` — do not skip this.** It reports what the
   template actually offers and what size text will really be.
4. **`powerpoint_add_slide`** per slide, with `notes` for what you will say.
5. **`powerpoint_review`** to check the deck against 10/20/30.
6. **`powerpoint_save`.**

## Hard rule: text goes in placeholders, never in a font override

A deck inherits its look from the template's masters, layouts and theme. Put
text into the layout's **placeholders** and the template styles it for you.
There is deliberately **no tool that sets a font, size, colour or position** —
that omission is the feature. If a slide looks wrong, the fix is a different
layout or fewer words, never a font.

- **Never** type `- `, `* `, `• ` or `1. ` into bullet text. The layout supplies
  the marker; a typed one renders on top of it. The server strips them, but
  don't rely on it.
- Nest with the `level` field (`{"text": "...", "level": 1}`), not with typed
  indentation.
- Address placeholders by **type word** (`"title"`, `"body"`, `"subtitle"`) or
  by `idx`. Type words are resolved against the actual layout, so they work on a
  branded template too.

## Reading the template: `powerpoint_list_layouts`

This is the tool that makes template adherence possible, and the one an agent is
most likely to skip. A branded template names its layouts whatever it likes
(`Acme Cover`, `Chapter Opener`, `Standard Content`), so **assuming the stock
Office layout numbering is the single most common way to produce an
off-template deck.**

It gives you, per layout:

- `index`, `name`, and a detected **`role`** — `title`, `section`, `bullets`,
  `two_content`, `title_only`, `picture`, `blank`. Roles are detected from the
  placeholders, not the name, so they survive any naming convention.
- every placeholder's `idx`, `type`, `name`, and **`level1_font_pt`** — the size
  text will *actually* render at, resolved through the whole inheritance chain
  (run → paragraph → shape → layout → master → master txStyles → presentation
  default). Body placeholders also carry **`deeper_levels`**: what sub-bullets
  shrink to.
- `recommended` — the best layout in *this* template for each job.
- `layouts_meeting_min_font_at_level_1` — the layouts whose body text clears
  30 points.

Pass a layout to `powerpoint_add_slide` by **index**, **name**, or **role word**.
Prefer the role word: `layout: "bullets"` works on every template.

## Adding slides

`powerpoint_add_slide` takes `title`, `subtitle`, `bullets`, `notes` and
`layout`, filling whatever the chosen layout supports:

- A **title slide**: `layout: "title"` + `title` + `subtitle`.
- A **content slide**: `layout: "bullets"` + `title` + `bullets`.
- A **two-column slide**: `layout: "two_content"`, then a second
  `powerpoint_set_placeholder` call naming the other placeholder by `idx`
  (`powerpoint_list_layouts` shows both).
- A **section divider**: `layout: "section"` + `title`.

Empty **text** placeholders are removed automatically, so no "Click to add text"
prompt survives. Picture/table/chart placeholders are **kept**, so a human can
drop an image in — this server never inserts images. Pass
`drop_empty_placeholders: false` to keep the text ones too.

Make each `title` the slide's **one assertion** ("Unpriced risk costs us £4m a
year"), not a topic label ("Risk"). The title is the only line most of the room
reads.

## Speaker notes are not optional

`notes` on `powerpoint_add_slide`, or `powerpoint_set_notes` later. Under the
10/20/30 rule the slide carries the headline and the **notes carry the
argument** — that is what keeps text off the slide and lets it stay at 30
points. Notes are also:

- the **only** input to the 20-minute estimate — a deck with no notes cannot be
  timed, and `powerpoint_review` says so rather than scoring it zero;
- **indexed into the knowledge base** along with the slides, so a deck you wrote
  is searchable by what you meant, not just by its headlines.

## Reviewing: `powerpoint_review`

Run it before saving, and again after any trim. It reports each part of the rule
pass/fail with its evidence, and **reports rather than fixes** — every fix is a
content decision.

The font check is the part you cannot do by eye or by reading the file: in a
well-built deck almost no run carries an explicit size, so anything that checks
the obvious property finds nothing wrong with a deck set entirely in 12pt. Each
finding names the slide, the shape, the resolved size and **which inheritance
step set it** (`source`), so *"why is this 28pt?"* has an answer.

- `source: "master_txstyles"` → the **template** sets that size. Do not fight
  it: use a different layout, or stop nesting. Deeper outline levels are smaller
  in nearly every template, so level 2 is usually where the rule breaks.
- `source: "run"` → something set the size explicitly. That should not happen
  with this server; if it does, the text came from the template.
- Table text is reported under **`unmeasured`**, not as a violation: it is sized
  by the deck's table style, which the server does not read. It carries a
  labelled estimate — check those tables by eye if they matter.

## Editing an existing deck

`powerpoint_open` → `powerpoint_get_content` (every slide's index, title, layout,
paragraphs with levels, tables, notes) → edit → `powerpoint_save`.

- `powerpoint_set_placeholder` replaces a placeholder's text;
  `powerpoint_add_bullets` appends without clearing.
- `powerpoint_delete_slide` — **indices shift down after each delete**, so
  delete from the highest index downwards, or re-read between deletes.
- `powerpoint_move_slide` returns the resulting order by title; use it to check
  the sequence reads as an argument rather than a list of topics.

## Notes

- All paths must be inside the configured presentations folder (plus the output
  folder, and the read-only templates folder); requests outside it are refused —
  don't fight the sandbox.
- **Templates are read-only.** Open one to read it or to run
  `powerpoint_list_layouts` on it — that is how you learn a corporate deck's
  layout names before drafting — but every save into that folder is refused.
  Build from it with `powerpoint_create(template=...)` instead.
- `.potx` files can be read as templates; decks are always **saved as `.pptx`**.
- Opening, creating and saving a deck each mirror it to Markdown for the
  knowledge base (on by default, to `C:\Eva\knowledge\powerpoint`).
- Not supported: charts, SmartArt, animations, transitions, images, and editing
  a template's theme.
