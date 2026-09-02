# Templates

Blank, branded starting files — letterhead, report layout, contract boilerplate,
the deck shell with your title slide. A new document is **created from** one of
these, inheriting its styles, headers/footers, page setup and boilerplate; the
template itself is never touched.

If what you have is a *finished* document showing what good output looks like,
that is an exemplar — put it in [`../exemplars`](../exemplars) instead.

| | |
|---|---|
| **Formats** | `.docx` (used by the `word` plugin), `.pptx` / `.potx` (used by the `powerpoint` plugin) |
| **Configuration** | `word` → templates folder (`--templates-dir` / `MSWORD_TEMPLATES_DIR`)<br>`powerpoint` → templates folder (`--templates-dir` / `POWERPOINT_TEMPLATES_DIR`) |
| **Writable?** | **no** — the server refuses every save into this folder |
| **Committed to git?** | no (see [`../.gitignore`](../../.gitignore)) |

**One folder, two plugins.** Both point here, and each reads only the file
types it understands: `word` reads the `.docx` files and ignores the rest,
`powerpoint` reads the `.pptx`/`.potx` files and ignores the rest. Keep every
template in here and there is one place to look.

## Point the `word` plugin at this folder

The [`word`](../../../plugins/word) plugin (v4.1.0+) takes a **templates folder**
as a read-only second root, alongside its documents folder. Pick whichever
install route you use:

**Plugin install** — answer the *Templates folder* prompt with the absolute path
to this folder:

```
C:\Eva\reference\templates
```

Already installed? `/plugin` → `word` → reconfigure, and set it there.

**`claude mcp add`:**

```powershell
claude mcp add word --scope user -e PYTHONUTF8=1 -- C:\path\to\python.exe C:\path\to\claude-skills\plugins\word\word.py --docs-dir C:\Eva\documents\word --templates-dir C:\Eva\reference\templates
```

**`.mcp.json`** — add to the `word` entry's `args` (see
[`.mcp.json.example`](../../../.mcp.json.example)):

```json
"--templates-dir", "C:\\Eva\\reference\\templates"
```

Confirm it took: the server logs `templates folder (read-only) = ...` at
startup, and `msword_list_documents` with `location: "templates"` lists exactly
what is in here.

## Point the `powerpoint` plugin at the same folder

Identical, with the `powerpoint` names. **Plugin install** — answer its
*Templates folder* prompt with `C:\Eva\reference\templates`.

**`claude mcp add`:**

```powershell
claude mcp add powerpoint --scope user -e PYTHONUTF8=1 -- C:\path\to\python.exe C:\path\to\claude-skills\plugins\powerpoint\powerpoint.py --docs-dir C:\Eva\documents\powerpoint --templates-dir C:\Eva\reference\templates
```

Confirm it took with `powerpoint_list_presentations` and
`location: "templates"`.

> The folder must exist and must be **separate from** the documents folder —
> the server refuses to start otherwise, rather than silently running with no
> templates or with every save blocked.

## What "read-only" buys you

The templates folder is readable everywhere the documents folder is — files can
be listed, opened and inspected (`msword_list_styles` on a template is the
reliable way to find out what your corporate bullet style is actually called) —
but **any save whose target lands in here is refused**, whether it is a save-as
or a save-in-place on a template opened by mistake. New documents always land in
the documents folder. The blanks stay blank.

## Naming

End the name with `Template` so intent is obvious in a listing:

```
Report Template.docx
Letterhead Template.docx
Meeting Agenda Template.docx
Contract Template.docx
Deck Template.pptx
```

Names are matched forgivingly — a bare name, a relative path, or a near-miss
like *"report template"* all resolve — so keep them distinct. If a name here
also exists in your documents folder, the documents folder wins; give the
template a different name to avoid the coin toss.

## Index

| File | Creates | Notes |
|---|---|---|
| _(none yet — add yours here)_ | | |

## Making a template Claude can fill in

Two conventions make the difference between "inherits the styling" and "fills
itself in":

1. **`{{TOKEN}}` placeholders** — `{{TITLE}}`, `{{CLIENT}}`, `{{MEETING_DATE}}`,
   `{{AUTHOR}}`. An explicit marker gives an unambiguous find-and-replace
   target; prose like *"Insert client name here"* is guesswork, and worse, may
   partly match real content.
2. **One styled example row** in any repeating table. Keep the header row and a
   single data row with the right borders, shading and fonts — that row is
   cloned per real item (`copy_from_row`) and then deleted, so the result keeps
   the formatting without you specifying any of it.

Also worth doing:

- Define the **styles** you want used (headings, bullets, the branded ones) in
  the template. Structure comes from Word styles, not typed characters, and a
  template that names its own styles is what makes `msword_list_styles` useful.
- Put the letterhead, footer, page numbering and any standing legal text in the
  template — every document created from it inherits them for free.
- Keep the body **empty**. Left-over sample paragraphs get treated as content to
  edit around, and a half-filled template is worse than a blank one.
- Save as **`.docx`, not `.dotx`**. Word's own template format is not supported;
  a plain `.docx` works exactly as well here, since the file is only ever read.

## How it gets used

- *"Create a Q3 report from my report template."* → `msword_create` with
  `template: "Report Template.docx"`, then the `add_*` tools, then `msword_save`
  — the new file lands in `C:\Eva\documents\word`.
- *"Use the agenda template and fill it out for Monday's meeting, one row per
  item."* → create from the template, `msword_replace_text` for the `{{TOKEN}}`s,
  a cloned table row per item, then drop the example row.
- *"What templates do I have?"* → `msword_list_documents` with
  `location: "templates"`.
- *"What's the bullet style in the letterhead template called?"* →
  `msword_open` on it + `msword_list_styles` (reading is fine; saving is not).

The full workflow, including filling out example tables, is in the
[`word` plugin README](../../../plugins/word/README.md) and its `SKILL.md`.

## PowerPoint templates

`.pptx` (and `.potx`) files here are read by the [`powerpoint`](../../../plugins/powerpoint)
plugin, which creates a new deck **from** one — inheriting its slide masters,
layouts, theme, fonts and colours — and writes the result to
`C:\Eva\documents\powerpoint`. The template itself is never modified, and any
example slides in it are stripped so the new deck starts blank with the styling
intact.

### What makes a good deck template

The `word` advice above is about styles and placeholders; for a deck it is about
**layouts**, because that is the only thing the server writes into.

- **Define a layout for each job** — a cover, a section divider, a
  title-and-bullets content slide, a two-column slide, a title-only slide. The
  server detects each layout's *role* from its placeholders rather than its
  name, so `Chapter Opener` and `Section Header` both work; name them for
  humans.
- **Set your font sizes in the slide master**, not on individual slides. That is
  where they are read from, and it is what lets a whole deck restyle at once.
- **Check the deeper outline levels.** Nearly every template shrinks each level —
  the stock Office one runs 32 / 28 / 24 / 20 point — so a level-2 bullet breaks
  the 30-point rule without anyone choosing a font. If your house style is
  30-point minimum, set level 2 to 30 too.
- **Keep the slides empty.** Example slides are stripped on create, so they cost
  nothing, but a template that is only masters and layouts is clearer.
- **Leave picture placeholders in.** The server writes no images and never drops
  an empty picture placeholder, so one in the layout becomes a correctly
  positioned, correctly styled click target for you to fill in by hand.
- `.potx` works and is only ever read; decks are always saved as `.pptx`.

### How it gets used

- *"Build a pitch deck from our template."* → `powerpoint_create` with
  `template: "Deck Template.pptx"`, then `powerpoint_list_layouts` to see what it
  offers, then a `powerpoint_add_slide` per slide.
- *"What layouts does our template have, and will they hold 30-point text?"* →
  `powerpoint_list_layouts`, which reports each placeholder's *effective* font
  size and lists the layouts that clear the minimum.
- *"Is this deck too long?"* → `powerpoint_review`, which audits slide count,
  estimated speaking time and font sizes against the 10/20/30 rule.

The full workflow is in the [`powerpoint` plugin
README](../../../plugins/powerpoint/README.md) and its two skills.
