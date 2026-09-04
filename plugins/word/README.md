# Word (.docx)

Read, edit and create Word documents — including **real Word tracked changes**,
native Word styles, and filling out templates.

| | |
|---|---|
| **Server** | `word.py` v8.0.0 |
| **pip install** | `python-docx` (pulls in `lxml` and `typing_extensions`) |
| **Platform** | any (Word itself is not required) |
| **Writes to disk** | yes — the only write-capable server in the suite |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install word@mcnamee-claude-skills
```

The `/word:word` skill is installed with the server. The only thing Claude Code
prompts for is the optional tracked-change author - every folder and the Python
interpreter come from the shared environment variables below.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `C:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `C:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `word`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_DOCUMENTS_DIR%\word` | **The one folder of `.docx` files.** Every open and save must be inside it, and **new** documents from `msword_create` are written here too - there is no separate output folder. Searched recursively, so a bare filename finds a file in a sub-folder | **Fatal.** The server is a path sandbox and refuses to start without it |
| `%EVA_TEMPLATES_DIR%\word` | Blank `.docx` templates new documents are created from - letterhead, report layout, contract boilerplate. **Read-only**: they can be listed, opened and passed as `msword_create`'s `template`, but every save into the folder is refused, so the blanks stay blank | Warns on stderr and runs with templates disabled |
| `%EVA_KNOWLEDGE_DIR%\word` | Markdown mirror of every document opened, created or saved (`Word - <name>.md`, overwritten each time) - which is what makes it searchable by the `knowledge-base` plugin. It must stay inside the knowledge root or the mirrors are never indexed | Created on demand. A mirror failure is logged and reported on the result, never allowed to fail the open or the save |

> The templates folder must be **separate from** the documents folder - if one
> contained the other, every save would be refused, so the server refuses to
> start instead.

### Overriding one folder, and this server's own settings

The shared roots are normally all you need. These variables are this
server's own, and a folder variable here beats the matching root - use one
only when an endpoint's layout really differs.

| Variable | Purpose |
|---|---|
| `MSWORD_AUTHOR` | Name stamped on Word tracked changes (default `AI Assistant`). Can also be overridden per call via the `author` argument on the editing tools |
| `MSWORD_DOCS_DIR` | Full path to the documents folder, instead of `%EVA_DOCUMENTS_DIR%\word` |
| `MSWORD_TEMPLATES_DIR` | Full path to the templates folder, instead of `%EVA_TEMPLATES_DIR%\word`. `off` runs with no templates root at all |
| `MSWORD_KB_DIR` | Full path to the mirror folder, instead of `%EVA_KNOWLEDGE_DIR%\word`. `off` disables mirroring |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies - that is what an MCP client substitutes for a prompt left
empty. To switch an optional folder off, set it to `off` (`none`, `no`, `false`
and `disabled` work too). The documents folder cannot be switched off: the
server has no sandbox without it.

A folder you named yourself that does not exist is **fatal** - it is almost
always a typo. The built-in default merely not existing yet is not: the server
warns and carries on without that feature.

> The plugin, folder and server file are named `word`, but these environment
> variables keep their `MSWORD_` prefix and the tools keep their `msword_`
> prefix - renaming those would break every existing config and every skill
> that names a tool, for no functional gain.

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Run an offline open/edit/save/reopen self-test and exit (no server). It sandboxes itself entirely to its own temp folders, so it never touches your real folders |
| `--version` | Print version and exit (works even without `python-docx` installed) |

## What it does well

**Finding documents by name.** You don't need absolute paths. `msword_open`
resolves a relative path against the **document root**, so *"edit Policy
103.docx"* opens `%EVA_DOCUMENTS_DIR%\word\Policy 103.docx` directly — a bare filename is even
located if it sits in a subfolder of the root. If there's no exact match,
`msword_open` falls back to a **fuzzy name match**, so *"budget policy"* opens
`Budget Policy 2024.docx`; when a fuzzy match is used the result carries
`fuzzy_matched: true` and the requested text so the caller can confirm it got
the right file, and a genuinely ambiguous name returns the tied candidates
rather than guessing. Use **`msword_list_documents`** (optionally with a `query`
substring) to list the `.docx` files under the root — name, relative path, size
and modified time — when you're unsure of the exact name.

**Writing a document in one call.** `msword_add_content` takes an ordered list
of **blocks** — headings, paragraphs, bullet/numbered items, tables, page breaks
— and appends them in exactly that order:

```json
{"session_id": "...", "blocks": [
  {"type": "title",     "text": "Quarterly Report"},
  {"type": "heading",   "text": "Overview", "level": 1},
  {"type": "paragraph", "text": "Revenue grew 8% on the prior quarter."},
  {"type": "bullet",    "text": "Widgets led the growth"},
  {"type": "bullet",    "text": "Services were flat"},
  {"type": "table",     "data": [["Item", "Cost"], ["Widget", "5"]]},
  {"type": "page_break"},
  {"type": "heading",   "text": "Appendix", "level": 1}
]}
```

This is the tool to build a document with, and it fixes a real failure mode.
The single-block tools (`msword_add_heading`, `msword_add_paragraph`,
`msword_add_table`) each append to the **end** of the document, so the
document's order is the order the *calls arrive* — and an MCP client may
dispatch independent tool calls in parallel, in which case they need not arrive
in the order the model wrote them. Drafting a report as thirty separate append
calls could therefore come back scrambled, classically with every heading
bunched together at the top and all the body text after them. One call carrying
the whole sequence cannot be reordered.

Details worth knowing:

- `type` is forgiving (`heading`/`h2`/`heading 2`, `paragraph`/`para`/`p`,
  `bullet`, `number`, `table`, `page_break`), a plain string is taken as a
  paragraph, and an omitted `type` is inferred from what's there (`level` → a
  heading, `data`/`rows` → a table).
- `bullet` and `number` use whatever bullet/numbered style **this** document
  defines, so a corporate template's own `DSCO Bullet` is picked up without you
  naming it; pass `style` to force a specific one. One block per item.
- The whole list is validated **before** anything is written, so a bad block
  leaves the document untouched and the error names its index
  (`blocks[7] has unknown type 'sidebar'`).
- `track_changes: true` records every added paragraph and heading as a real
  tracked insertion (tables can't be tracked, and the result says so).
- Headings are echoed back in the result, so the outline can be checked at a
  glance without re-reading the document.
- Still chaining the single-block tools? Three or more of them landing on the
  same session within a second — the signature of a parallel batch — adds an
  `order_warning` to the result, so a scrambled document gets reported instead
  of silently shipped.

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

**Creating documents.** `msword_create` makes a new `.docx` in the **documents
folder** and opens it as a session; build it up with `msword_add_content` and
persist with `msword_save` (omit its `path` to save in place). Any directory
part in the requested filename is stripped, so new files land at the top of the
documents folder — to file one in a sub-folder, save-as with
`msword_save(path: "Reports/Q3.docx")`. There is no separate output folder: what
Eva writes sits in the same library as what you gave it.

**From a template.** Pass `template` to `msword_create` to base the new document
on an existing one — a corporate letterhead, report layout or contract
boilerplate. The template is a `.docx` in the **templates folder** or the
document root, named the same forgiving way as `msword_open` (bare name,
relative path, or a fuzzy near-miss); its styles, headers/footers, page setup
and boilerplate are inherited into the new file, which is written into the
documents folder. **The template file itself is never modified**, and the result includes
the resolved `template` so you can confirm the right one was used. `.docx`
templates only (Word's `.dotx` is not supported — save the template as `.docx`).

**A templates folder of its own.** `%EVA_TEMPLATES_DIR%\word` (by default
`C:\Eva\templates\word`) stops templates being ordinary documents that
happen to live in the docs folder. The folder is a
**read-only** second root: `msword_list_documents` reports each file's `location`
(`docs` / `templates`) and takes a `location: "templates"` filter to
list just the blanks; templates can be opened and inspected (`msword_list_styles`
on one is how you learn what the corporate bullet style is really called); and
**every save into the folder is refused** — save-as *and* save-in-place on a
template opened by mistake. New documents always land in the documents folder.
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

**Building a RAG knowledge base.** The mirror folder sits inside the same
knowledge root the `knowledge-base` server indexes - which is what
`EVA_KNOWLEDGE_DIR` being one shared setting buys you. Each time a `.docx` is opened, created or
saved, a Markdown copy is dropped there — headings become `#`/`##`, `List
Bullet`/`List Number` paragraphs become `-`/`1.` lists, and tables become
GitHub-style pipe tables — so Word content lands alongside your Confluence pages
in the same RAG index.

Mirroring on **save** is what captures a document you *wrote*: the copy is
refreshed with every save and named from the saved path, so a save-as lands
under the new name. Without it, a report drafted here would only reach the
knowledge base if somebody later re-opened the `.docx`.

## File access

Open/save only inside the documents folder (`%EVA_DOCUMENTS_DIR%\word`), which
is also where **new** documents are written; the templates folder
(`%EVA_TEMPLATES_DIR%\word`) is **readable but never writable**; Markdown
mirrored to the knowledge folder (`%EVA_KNOWLEDGE_DIR%\word`) on open, create
and save. Paths are
resolved (symlinks included) before the containment check, so a symlink dropped
inside a configured folder cannot reach files outside it.

## Usage examples

1. "Edit the Word file Policy 103.docx to…" → `msword_open` with `path: "Policy 103.docx"` (resolved against the document root — no absolute path needed) + the editing tools
2. "Open the budget policy doc." (name not exact) → `msword_open` with `path: "budget policy"` (fuzzy-matches `Budget Policy 2024.docx`; the result flags `fuzzy_matched` so you can confirm)
3. "What Word documents do I have?" / "I'm not sure of the exact file name." / "What templates can I start from?" → `msword_list_documents` (optionally with a `query`, or `location: "templates"` for just the blanks), then `msword_open` on the one you want
4. "Open the proposal.docx and show me its full text." → `msword_open` + `msword_get_content`
5. "Open every .docx in my docs folder so it gets mirrored into the RAG knowledge base as Markdown." → `msword_open` (each open writes `Word - <name>.md` into the knowledge folder; so does each create and save)
6. "Create a new status report document and draft it with a title, headings and a summary table." → `msword_create` (writes into the documents folder) + **one** `msword_add_content` call listing the blocks in order + `msword_save`
7. "Create a Q3 report from my report template." → `msword_create` with `template: "Report Template.docx"` (found in the templates folder or the docs folder; inherits the template's styles/headers/boilerplate into a new file, which lands in the documents folder — the template is left untouched) + `msword_add_content` + `msword_save`
8. "Use my agenda template and fill it out for Monday's meeting — one row per item." → `msword_create` with `template: "Agenda Template.docx"` + `msword_replace_text` (placeholders) + `msword_add_table_row` (with `copy_from_row` to clone the example row) per item + `msword_set_cell` + `msword_delete_table_row` (drop leftover example rows) + `msword_save`
9. "Find every mention of 'Acme Corp' in the contract and replace it with 'Acme Corporation'." → `msword_search` + `msword_replace_text`
10. "Add a 'Next Steps' heading and a summary paragraph to the end of the report, then save it." → `msword_add_content` with both blocks in one call (a heading and its paragraph sent as two separate calls can land in either order) + `msword_save`
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
24. "Add these five points as proper Word bullet points, not hyphens." → one `msword_add_content` call with five `{"type": "bullet", "text": "..."}` blocks (typing "- " is auto-corrected and flagged in a `warning`)
25. "Turn that hand-typed hyphen list into real Word bullets." → `msword_get_content` (`mode: "structured"`) + `msword_set_paragraph_text` + `msword_set_paragraph_format` with `style: "List Bullet"`

## Troubleshooting

> **If a server fails with `Executable not found in $PATH: "${EVA_PYTHON}"`**,
> the variable is not set in the environment Claude Code was launched from. Set
> it (see above), then quit Claude Code completely and reopen — `setx` and
> `[Environment]::SetEnvironmentVariable` do not reach a process that is already
> running.

- **The document came out in the wrong order** (headings bunched together, body
  text after them) — that is the signature of a document built with a series of
  separate `msword_add_heading`/`msword_add_paragraph` calls: they each append
  to the end, so the order is whatever order the calls reached the server, and
  parallel dispatch doesn't preserve it. Build the document with **one**
  `msword_add_content` call carrying the blocks in order. Since v5.1.0 a rapid
  burst of single appends also comes back with an `order_warning` saying so.
- **"dependency missing" after installing `python-docx`** — the server logs
  `sys.executable` on startup. Almost always the Python interpreter you gave the
  plugin isn't the one you pip-installed into:
```powershell
  & $env:EVA_PYTHON -m pip install python-docx
  ```
- **Junk `Word - *.md` files in the knowledge folder** (named after test
  documents: `Word - roundtrip.md`, `Word - pdel2.md`, …) — a pre-v5.1.0
  `--check` run mirrored its own scratch documents into the real knowledge
  folder. Delete them and re-index; `--check` now stays inside its temp folders.
- **Check the config before wiring it in**, which also runs a full offline
  self-test:
```powershell
  & $env:EVA_PYTHON word.py --check
  ```
- **Server won't start after setting a templates folder** — the log says which
  of the two guards fired: the folder doesn't exist (check `EVA_TEMPLATES_DIR`
  and that its `word\` sub-folder is there), or it is the same as (or contains)
  the documents folder, which would block every save. Keep the two roots
  separate.
- **"the documents folder does not exist"** — the server prints the path it
  resolved and whether it came from `EVA_DOCUMENTS_DIR`/`MSWORD_DOCS_DIR` or the
  built-in default. Create `%EVA_DOCUMENTS_DIR%\word`, or fix the variable.
- **"Saving into the templates folder is not allowed"** — working as intended:
  templates are read-only. Create the document with `msword_create`
  (`template: "..."`), which writes into the documents folder, instead of
  opening the template and saving it.
- For airgapped installs, the docstring at the top of `word.py` walks through
  sideloading the wheels.
