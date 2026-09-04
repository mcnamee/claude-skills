# templates\word\

Blank `.docx` templates the `word` plugin creates new documents from —
letterhead, report layout, contract boilerplate, a meeting agenda. The template
supplies the styles, headers, footers, page setup and standing text; the new
document lands in [`..\..\documents\word`](../../documents/word).

| | |
|---|---|
| **Setting** | `EVA_TEMPLATES_DIR` (the `word` plugin appends `\word`) |
| **Default** | `C:\Eva\templates\word` |
| **Formats** | `.docx` (not `.dotx` — see below) |
| **Access** | read **only** — every save into this folder is refused |
| **Committed to git?** | no (see [`..\..\.gitignore`](../../.gitignore)) |

If what you have is a *finished* document showing what good output looks like,
that is an exemplar, and it belongs in the `exemplars\` folder of the skill that
reads it — `/exemplar-writer`, `/brief-writer` or `/email-writer`. See
[`..\README.md`](../README.md).

## Naming

End the name with `Template` so intent is obvious in a listing:

```
Report Template.docx
Letterhead Template.docx
Meeting Agenda Template.docx
Contract Template.docx
```

Names are matched forgivingly — a bare name, a relative path, or a near-miss
like *"report template"* all resolve — so keep them distinct. If a name here
also exists in the documents folder, the documents folder wins; give the
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

## What "read-only" buys you

The folder is readable everywhere the documents folder is — files can be listed,
opened and inspected (`msword_list_styles` on a template is the reliable way to
find out what your corporate bullet style is actually called) — but **any save
whose target lands in here is refused**, whether it is a save-as or a
save-in-place on a template opened by mistake. New documents always land in
[`..\..\documents\word`](../../documents/word). The blanks stay blank.

## How it gets used

- *"Create a Q3 report from my report template."* → `msword_create` with
  `template: "Report Template.docx"`, then `msword_add_content`, then
  `msword_save` — the new file lands in `C:\Eva\documents\word`.
- *"Use the agenda template and fill it out for Monday's meeting, one row per
  item."* → create from the template, `msword_replace_text` for the
  `{{TOKEN}}`s, a cloned table row per item, then drop the example row.
- *"What templates do I have?"* → `msword_list_documents` with
  `location: "templates"`.
- *"What's the bullet style in the letterhead template called?"* →
  `msword_open` on it + `msword_list_styles` (reading is fine; saving is not).

The full workflow, including filling out example tables, is in the
[`word` plugin README](../../../plugins/word/README.md) and its `SKILL.md`.
