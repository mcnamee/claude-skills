# Templates

Blank, branded starting files — letterhead, report layout, contract boilerplate,
the deck shell with your title slide. A new document is **created from** one of
these, inheriting its styles, headers/footers, page setup and boilerplate; the
template itself is never touched.

If what you have is a *finished* document showing what good output looks like,
that is an exemplar — put it in [`../exemplars`](../exemplars) instead.

| | |
|---|---|
| **Formats** | `.docx` (used by the `word` plugin), `.pptx` (see [PowerPoint](#powerpoint-templates)) |
| **Configuration** | `word` plugin → templates folder (`--templates-dir` / `MSWORD_TEMPLATES_DIR`) |
| **Writable?** | **no** — the server refuses every save into this folder |
| **Committed to git?** | no (see [`../.gitignore`](../.gitignore)) |

## Point the `word` plugin at this folder

The [`word`](../../plugins/word) plugin (v4.1.0+) takes a **templates folder**
as a read-only third root, alongside its documents and output folders. Pick
whichever install route you use:

**Plugin install** — answer the *Templates folder* prompt with the absolute path
to this folder:

```
C:\path\to\claude-skills\context\templates
```

Already installed? `/plugin` → `word` → reconfigure, and set it there.

**`claude mcp add`:**

```powershell
claude mcp add word --scope user -e PYTHONUTF8=1 -- C:\path\to\python.exe C:\path\to\claude-skills\plugins\word\word.py --docs-dir C:\Users\me\Documents\ai_docs --output-dir C:\Users\me\Documents\ai_generated --templates-dir C:\path\to\claude-skills\context\templates
```

**`.mcp.json`** — add to the `word` entry's `args` (see
[`.mcp.json.example`](../../.mcp.json.example)):

```json
"--templates-dir", "C:\\path\\to\\claude-skills\\context\\templates"
```

Confirm it took: the server logs `templates folder (read-only) = ...` at
startup, and `msword_list_documents` with `location: "templates"` lists exactly
what is in here.

> The folder must exist and must be **separate from** the documents and output
> folders — the server refuses to start otherwise, rather than silently running
> with no templates or with every save blocked.

## What "read-only" buys you

The templates folder is readable everywhere the documents folder is — files can
be listed, opened and inspected (`msword_list_styles` on a template is the
reliable way to find out what your corporate bullet style is actually called) —
but **any save whose target lands in here is refused**, whether it is a save-as
or a save-in-place on a template opened by mistake. New documents always land in
the output folder. The blanks stay blank.

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
  — the new file lands in the output folder.
- *"Use the agenda template and fill it out for Monday's meeting, one row per
  item."* → create from the template, `msword_replace_text` for the `{{TOKEN}}`s,
  a cloned table row per item, then drop the example row.
- *"What templates do I have?"* → `msword_list_documents` with
  `location: "templates"`.
- *"What's the bullet style in the letterhead template called?"* →
  `msword_open` on it + `msword_list_styles` (reading is fine; saving is not).

The full workflow, including filling out example tables, is in the
[`word` plugin README](../../plugins/word/README.md) and its `SKILL.md`.

## PowerPoint templates

`.pptx` files are welcome here, but note that **this suite has no PowerPoint
server** — nothing in `plugins/` reads or writes `.pptx`, and the `word` server
ignores every file here that is not a `.docx`. They are for Claude Code's own
`pptx` skill where that is available, and for you to open by hand. Keep them
alongside the Word templates so there is one place to look; if a PowerPoint
plugin ever lands in this repo, this is the folder it will point at.
