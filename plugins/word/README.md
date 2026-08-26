# Word (.docx)

Read, edit and create Word documents — including **real Word tracked changes**,
native Word styles, and filling out templates.

| | |
|---|---|
| **Server** | `word.py` v5.0.0 |
| **pip install** | `python-docx` (pulls in `lxml` and `typing_extensions`) |
| **Platform** | any (Word itself is not required) |
| **Writes to disk** | yes — the only write-capable server in the suite |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install word@mcnamee-claude-skills
```

Claude Code prompts for the settings below; the `/word:word` skill is
installed with the server.

Every folder is pre-filled with its place in the [Eva working
tree](../../eva) — copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
you can accept all four as they stand.

| Prompt | Default | Env var | Purpose |
|---|---|---|---|
| Documents folder | `C:\Eva\documents\word` | `MSWORD_DOCS_DIR` | Sandbox: every open/save must be inside this tree. **Required** — the server refuses to start without it |
| Output folder | `C:\Eva\output\word` | `MSWORD_OUTPUT_DIR` | Where `msword_create` writes **new** documents. `off` writes them into the documents folder instead |
| Templates folder | `C:\Eva\reference\templates` | `MSWORD_TEMPLATES_DIR` | Blank `.docx` templates new documents are created from — **read-only**. `off` for no templates |
| Knowledge-base folder | `C:\Eva\knowledge\word` | `MSWORD_KB_DIR` | Mirrors every document opened, created or saved to Markdown, which is what makes it searchable. `off` to disable mirroring |
| Tracked-change author | — | `MSWORD_AUTHOR` | Name stamped on tracked changes |
| Python interpreter | — | — | **Required.** Absolute path to the `python.exe` that has `python-docx` installed |

> **Blank does not mean off.** Leaving a folder prompt empty means "not
> configured", so the default above applies. To switch one off, type `off`
> (`none`, `no`, `false` and `disabled` work too). The documents folder cannot
> be switched off — the server has no sandbox without it.

> The plugin, folder and server file are named `word`, but the environment
> variables keep their `MSWORD_` prefix and the tools keep their `msword_`
> prefix — renaming those would break every existing config and every skill
> that names a tool, for no functional gain.

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| CLI flag | Env var | Purpose |
|---|---|---|
| `--docs-dir` | `MSWORD_DOCS_DIR` | **Required.** Path sandbox: every open/save must be inside this directory tree, and the server refuses to start without one (`--check` is exempt — the self-test sandboxes itself to its own temp folder). Falls back to the `DOCS_DIR` config value, default `C:\Eva\documents\word`. This is the only write-capable server in the suite, and the model chooses the open/save paths |
| `--output-dir` | `MSWORD_OUTPUT_DIR` | Folder where `msword_create` writes **new** `.docx` files, kept **separate** from the knowledge-base folder. Falls back to the `OUTPUT_DIR` config value, default `C:\Eva\output\word`; pass `off` to fall back to the document root instead. Also treated as a permitted open/save location so created documents can be reopened and edited |
| `--templates-dir` | `MSWORD_TEMPLATES_DIR` | Folder of blank `.docx` templates. Falls back to the `TEMPLATES_DIR` config value, default `C:\Eva\reference\templates`; pass `off` for no templates root. A **read-only** third root: its files can be listed, opened and passed as `msword_create`'s `template`, but **every save into it is refused**, so templates stay blank. Must be separate from the documents and output folders — the server refuses to start otherwise. A folder you configured yourself that does not exist is fatal; the built-in default merely not existing yet logs a warning and runs without templates |
| `--kb-dir` | `MSWORD_KB_DIR` | **Every document opened, created or saved** is *also* written out as a Markdown file into this folder for a local RAG knowledge base. Falls back to the `KB_DIR` config value, default `C:\Eva\knowledge\word` — which sits inside the `knowledge-base` server's documents folder, so mirrored documents are actually indexed. Files are named `Word - <name>.md` and overwritten each time; the folder is created if missing. A mirror failure is logged and reported on the result, never allowed to fail the open or the save. Pass `off` to disable mirroring |
| `--author` | `MSWORD_AUTHOR` | Author name stamped on Word tracked changes. Falls back to the `TRACKED_CHANGE_AUTHOR` config value. Can also be overridden per-call via the `author` argument on the editing tools |
| `--check` | — | Run an offline open/edit/save/reopen self-test and exit (no server) |
| `--version` | — | Print version and exit (works even without `python-docx` installed) |

## What it does well

**Finding documents by name.** You don't need absolute paths. `msword_open`
resolves a relative path against the **document root**, so *"edit Policy
103.docx"* opens `<docs-dir>\Policy 103.docx` directly — a bare filename is even
located if it sits in a subfolder of the root. If there's no exact match,
`msword_open` falls back to a **fuzzy name match**, so *"budget policy"* opens
`Budget Policy 2024.docx`; when a fuzzy match is used the result carries
`fuzzy_matched: true` and the requested text so the caller can confirm it got
the right file, and a genuinely ambiguous name returns the tied candidates
rather than guessing. Use **`msword_list_documents`** (optionally with a `query`
substring) to list the `.docx` files under the root — name, relative path, size
and modified time — when you're unsure of the exact name.

**Native Word styles.** Structure lives in paragraph *styles*, not typed
characters, so the tools steer towards real ones: `msword_list_styles` reports
every style a document actually defines (with each one's role and what the
Markdown export turns it into, plus a `recommended` block naming the right style
for *this* template), style names resolve forgivingly (case-insensitive,
`ListBullet`, or an alias like `bullets`) and an unknown name comes back with
the closest real matches instead of a dead end. Text typed as a fake list item
(`"- First point"`) is auto-corrected to a real `List Bullet` paragraph and the
correction is reported in a `warning` — pass `literal_text: true` for text that
must keep a leading marker (`- 5 degrees`), and note a multi-item block in one
call is refused (add one paragraph per item). `msword_insert_paragraph` with no
`style` follows Word's Enter-key rule instead of defaulting to Normal, and
`msword_add_table` defaults to `Table Grid` so tables have visible borders. This
matters beyond appearance: the knowledge-base export is **entirely
style-driven**, so a hand-typed hyphen is mirrored as a plain paragraph and the
list structure is lost.

**Tracked changes**, recorded the way Word itself records them: replacements are
diffed **word-by-word** (only the words that actually change are marked as
deleted/inserted — never "whole paragraph deleted + whole paragraph
reinserted"), and whole-paragraph inserts/deletes include the paragraph mark so
accepting/rejecting adds or removes the paragraph itself. Changes can be
accepted/rejected all at once or individually by id. While changes are pending,
`msword_get_content`/`msword_search` show the final ("No Markup") view.

**Creating documents.** `msword_create` makes a new `.docx` in the
`--output-dir` folder (falling back to the document root) and opens it as a
session; build it up with `msword_add_heading` / `msword_add_paragraph` /
`msword_add_table` and persist with `msword_save` (omit its `path` to save in
place). Any directory part in the requested filename is stripped, so new files
always land inside the output folder.

**From a template.** Pass `template` to `msword_create` to base the new document
on an existing one — a corporate letterhead, report layout or contract
boilerplate. The template is a `.docx` in the **templates folder** or the
document root, named the same forgiving way as `msword_open` (bare name,
relative path, or a fuzzy near-miss); its styles, headers/footers, page setup
and boilerplate are inherited into the new file, which is written to the output
folder. **The template file itself is never modified**, and the result includes
the resolved `template` so you can confirm the right one was used. `.docx`
templates only (Word's `.dotx` is not supported — save the template as `.docx`).

**A templates folder of its own.** `--templates-dir` (by default
`C:\Eva\reference\templates`) stops templates being ordinary documents that
happen to live in the docs folder. The folder becomes a
**read-only** third root: `msword_list_documents` reports each file's `location`
(`docs` / `output` / `templates`) and takes a `location: "templates"` filter to
list just the blanks; templates can be opened and inspected (`msword_list_styles`
on one is how you learn what the corporate bullet style is really called); and
**every save into the folder is refused** — save-as *and* save-in-place on a
template opened by mistake. New documents always land in the output folder.
Pass `off` and templates are resolved from the docs folder like any other
document.

**Filling out an example template.** When the template is a form to fill in —
placeholder text plus an *example* table (e.g. an agenda with a Time/Item/Owner
table) — the table-editing tools let an agent populate it. Read the structure
with `msword_get_content` (`mode: "structured"`) and `msword_get_tables`, swap
placeholder text with `msword_replace_text` (templates that use explicit
`{{TOKEN}}` markers are the most reliable to fill), then for the repeating
table: `msword_add_table_row` (with `copy_from_row` to clone a styled example
row's borders/shading/fonts, and `values` to fill it) once per real item,
`msword_set_cell` to set an individual cell by `(table_index, row, col)`, and
`msword_delete_table_row` to drop leftover example rows (delete highest index
first, since indices shift). These table edits are plain (untracked); a table
row/cell change can't be a Word tracked change.

**Building a RAG knowledge base.** Point `--kb-dir` at the same folder your
`knowledge-base` server indexes. Each time a `.docx` is opened, created or
saved, a Markdown copy is dropped there — headings become `#`/`##`, `List
Bullet`/`List Number` paragraphs become `-`/`1.` lists, and tables become
GitHub-style pipe tables — so Word content lands alongside your Confluence pages
in the same RAG index.

Mirroring on **save** is what captures a document you *wrote*: the copy is
refreshed with every save and named from the saved path, so a save-as lands
under the new name. Without it, a report drafted here would only reach the
knowledge base if somebody later re-opened the `.docx`.

## File access

Open/save only inside the documents folder (`C:\Eva\documents\word`) and the
output folder (`C:\Eva\output\word`); the templates folder
(`C:\Eva\reference\templates`) is **readable but never writable**; new
documents written to the output folder; Markdown mirrored to the knowledge-base
folder (`C:\Eva\knowledge\word`) on open, create and save. Paths are
resolved (symlinks included) before the containment check, so a symlink dropped
inside a configured folder cannot reach files outside it.

## Usage examples

1. "Edit the Word file Policy 103.docx to…" → `msword_open` with `path: "Policy 103.docx"` (resolved against the document root — no absolute path needed) + the editing tools
2. "Open the budget policy doc." (name not exact) → `msword_open` with `path: "budget policy"` (fuzzy-matches `Budget Policy 2024.docx`; the result flags `fuzzy_matched` so you can confirm)
3. "What Word documents do I have?" / "I'm not sure of the exact file name." / "What templates can I start from?" → `msword_list_documents` (optionally with a `query`, or `location: "templates"` for just the blanks), then `msword_open` on the one you want
4. "Open the proposal.docx and show me its full text." → `msword_open` + `msword_get_content`
5. "Open every .docx in my docs folder so it gets mirrored into the RAG knowledge base as Markdown." → `msword_open` with `--kb-dir` set (each open writes `Word - <name>.md`; so does each create and save)
6. "Create a new status report document and draft it with a title, headings and a summary table, then save it to my generated-docs folder." → `msword_create` (writes to `--output-dir`) + `msword_add_heading` + `msword_add_paragraph` + `msword_add_table` + `msword_save`
7. "Create a Q3 report from my report template." → `msword_create` with `template: "Report Template.docx"` (found in the templates folder or the docs folder; inherits the template's styles/headers/boilerplate into a new file, which lands in the output folder — the template is left untouched) + the add_* tools + `msword_save`
8. "Use my agenda template and fill it out for Monday's meeting — one row per item." → `msword_create` with `template: "Agenda Template.docx"` + `msword_replace_text` (placeholders) + `msword_add_table_row` (with `copy_from_row` to clone the example row) per item + `msword_set_cell` + `msword_delete_table_row` (drop leftover example rows) + `msword_save`
9. "Find every mention of 'Acme Corp' in the contract and replace it with 'Acme Corporation'." → `msword_search` + `msword_replace_text`
10. "Add a 'Next Steps' heading and a summary paragraph to the end of the report, then save it." → `msword_add_heading` + `msword_add_paragraph` + `msword_save`
11. "Pull out the data from every table in the document as structured rows." → `msword_get_tables`
12. "Add a 3x4 pricing table to the end of the quote document with these values, using the 'Table Grid' style." → `msword_add_table` + `msword_save`
13. "Fill cell B2 of the second table with 'Approved', and add a row for the new line item." → `msword_set_cell` + `msword_add_table_row`
14. "Change 'DRAFT' to 'FINAL' throughout the report as a tracked change so it shows up as a Word revision for review." → `msword_replace_text` with `track_changes=true`
15. "Rewrite the third paragraph to be more concise, showing your edits as tracked changes — only mark the words you actually changed." → `msword_set_paragraph_text` with `track_changes=true`
16. "Add a new paragraph after the introduction as a tracked insertion, so reviewers can reject it if they disagree." → `msword_insert_paragraph` with `track_changes=true`
17. "Delete the whole limitation-of-liability paragraph as a tracked change — struck out, so legal can accept or reject it." → `msword_delete_paragraph` with `track_changes=true`
18. "What tracked changes are currently in this document, and who made them?" → `msword_list_changes`
19. "Accept Jane's two changes in the pricing section but leave everything else pending." → `msword_list_changes` + `msword_accept_changes` with those change ids
20. "Reject just the change that deleted the warranty sentence." → `msword_list_changes` + `msword_reject_changes` with that change id
21. "Accept all the tracked changes in this document now that legal has signed off." → `msword_accept_all_changes`
22. "Reject all the tracked changes and revert this document to its original wording." → `msword_reject_all_changes`
23. "What styles does this template actually have — I want the corporate bullet style, not a generic one." → `msword_list_styles`
24. "Add these five points as proper Word bullet points, not hyphens." → `msword_add_paragraph` once per point with `style: "List Bullet"` (typing "- " is auto-corrected and flagged in a `warning`)
25. "Turn that hand-typed hyphen list into real Word bullets." → `msword_get_content` (`mode: "structured"`) + `msword_set_paragraph_text` + `msword_set_paragraph_format` with `style: "List Bullet"`

## Troubleshooting

- **"dependency missing" after installing `python-docx`** — the server logs
  `sys.executable` on startup. Almost always the Python interpreter you gave the
  plugin isn't the one you pip-installed into:
```powershell
  & "C:\path\to\python.exe" -m pip install python-docx
  ```
- **Check the config before wiring it in**, which also runs a full offline
  self-test:
```powershell
  & "C:\path\to\python.exe" word.py --check
  ```
- **Server won't start after setting a templates folder** — the log says which
  of the two guards fired: the folder doesn't exist (check the path), or it is
  the same as (or contains) the documents/output folder, which would block every
  save. Give templates a folder of their own.
- **"Saving into the templates folder is not allowed"** — working as intended:
  templates are read-only. Create the document with `msword_create`
  (`template: "..."`), which writes to the output folder, instead of opening the
  template and saving it.
- For airgapped installs, the docstring at the top of `word.py` walks through
  sideloading the wheels.
