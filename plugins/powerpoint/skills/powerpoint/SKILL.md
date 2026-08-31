---
name: powerpoint
description: Build PowerPoint .pptx decks via the powerpoint MCP server, inheriting a corporate template's masters, layouts, theme, fonts and colours; build a whole deck in one ordered call, add slides onto named layouts, fill placeholders, write speaker notes, add tables, reorder and delete slides, and audit the result against the 10/20/30 rule. Use when the user asks to create, build, edit, read or review a PowerPoint deck or .pptx file, to make slides from a template, or to check a deck's font sizes, slide count or length.
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
4. **`powerpoint_add_slides`** — ONE call carrying every slide in order, each
   with its `notes` (see the next section).
5. **`powerpoint_review`** to check the deck against 10/20/30.
6. **`powerpoint_save`.**

## Hard rule: build the deck with ONE `powerpoint_add_slides` call

`powerpoint_add_slide` appends its slide to the **end** of the deck, so the
deck's order is the order the *calls arrive at the server*. Independent tool
calls issued together may be dispatched in parallel, and then they need not
arrive in the order you wrote them — which comes out as a shuffled deck. It gets
worse downstream: `powerpoint_set_notes`, `powerpoint_set_placeholder`,
`powerpoint_add_bullets` and `powerpoint_add_table` all address a slide by
`slide_index`, so assuming *"the third slide I created is index 2"* writes onto
whichever slide actually landed there.

So, for **anything past a single slide**, send one `powerpoint_add_slides` call
with the whole deck in `slides`, in order:

```json
{"session_id": "abc123", "slides": [
  {"layout": "title",   "title": "FY26 plan", "subtitle": "Board review",
   "notes": "Thanks for making the time."},
  {"layout": "bullets", "title": "Unpriced risk costs us £4m a year",
   "bullets": ["Claims up 20%", {"text": "mostly EMEA", "level": 1}],
   "notes": "Walk through where the four million goes."},
  {"layout": "section", "title": "Our answer"},
  {"layout": "bullets", "title": "The numbers",
   "table": {"rows": [["Region", "Q3"], ["APAC", "1.2m"]]},
   "notes": "APAC carries the quarter."},
  {"layout": "two_content", "title": "Now vs next",
   "bullets": ["Today: manual"],
   "placeholders": [{"placeholder": 2, "bullets": ["Next: priced at bind"]}]}
]}
```

- Each entry takes everything `powerpoint_add_slide` takes — `layout`, `title`,
  `subtitle`, `bullets`, `placeholder`, `notes`, `drop_empty_placeholders` —
  **plus** two things that used to need a second call against a `slide_index`:
  - `placeholders`: `[{"placeholder": <idx|name|type word>, "bullets": [...]}]`
    — how a two-content layout's second column gets filled.
  - `table`: `{"rows": [[...], [...]]}` — first row is the header.
- Every entry, **including every layout name**, is validated before any slide is
  added, so a bad entry leaves the deck untouched and the error names its index.
  If a layout/content mismatch only surfaces mid-build, the whole batch unwinds.
- The result returns each slide's real `slide_index` in order — use those for any
  later edit, never an assumption about creation order.
- More slides than fit one call? Send consecutive calls; each appends after the
  last. Never split one deck across parallel calls.
- Keep `powerpoint_add_slide` for a genuinely single slide added on its own. If a
  result comes back with an `order_warning`, the server spotted a burst of single
  appends: check the deck with `powerpoint_get_content` and fix the sequence with
  `powerpoint_move_slide`.

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

Pass a layout by **index**, **name**, or **role word**. Prefer the role word:
`layout: "bullets"` works on every template.

## Adding slides

Each entry in `powerpoint_add_slides` (or a lone `powerpoint_add_slide`) fills
whatever the chosen layout supports:

- A **title slide**: `layout: "title"` + `title` + `subtitle`.
- A **content slide**: `layout: "bullets"` + `title` + `bullets`.
- A **two-column slide**: `layout: "two_content"` + `bullets` for the first
  column and `placeholders: [{"placeholder": <idx>, "bullets": [...]}]` for the
  second (`powerpoint_list_layouts` shows both idx values).
- A **section divider**: `layout: "section"` + `title`.
- A **table slide**: `layout: "bullets"` + `title` + `table: {"rows": [...]}`,
  leaving `bullets` out so the table can take the content placeholder's spot.

Empty **text** placeholders are removed automatically, so no "Click to add text"
prompt survives. Picture/table/chart placeholders are **kept**, so a human can
drop an image in — this server never inserts images. Pass
`drop_empty_placeholders: false` to keep the text ones too.

Make each `title` the slide's **one assertion** ("Unpriced risk costs us £4m a
year"), not a topic label ("Risk"). The title is the only line most of the room
reads.

## Speaker notes are not optional

`notes` on each `powerpoint_add_slides` entry (or `powerpoint_set_notes` later,
against a `slide_index` you actually read back). Under the
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
  `powerpoint_add_bullets` appends without clearing. Both need a `slide_index`
  from `powerpoint_get_content` or from what `powerpoint_add_slides` returned —
  never one inferred from the order slides were created in.
- Adding several slides to an existing deck is still one `powerpoint_add_slides`
  call; they append after the last slide, then `powerpoint_move_slide` if they
  belong somewhere else.
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
