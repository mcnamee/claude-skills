# Confluence

Search and read Confluence pages across one or two Confluence instances, macro
content included; download a page's attachments into the folder the matching
plugin reads; save a page to Markdown — when you ask for it — so it feeds a
local RAG knowledge base; and, once you switch writing on, create, update and
add to pages.

| | |
|---|---|
| **Server** | `confluence.py` v7.2.0 |
| **pip install** | _none_ — standard library only (HTTP via stdlib `urllib`) |
| **Platform** | any |
| **Writes to disk** | only when you ask: a saved page (Markdown, `H:\Eva\knowledge\confluence`) or a downloaded attachment (`H:\Eva\documents\<type>`) |
| **Writes to Confluence** | only with `CONFLUENCE_ALLOW_WRITE=true` — off by default |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install confluence@mcnamee-claude-skills
```

The plugin prompts for nothing at install. Every setting is an environment
variable the server inherits from Claude Code: set these in the `env` block of
Claude Code's settings.json ([Where to set a variable](../../README.md#where-to-set-a-variable))
or as Windows user environment variables, then fully restart Claude Code. A
blank or missing value means "not set".

| Variable | Required | Purpose |
|---|---|---|
| `CONFLUENCE_BASE_URL` | **yes** | Base URL including any context path, no trailing slash. This is the **default** server |
| `CONFLUENCE_TOKEN` | **yes**\* | Personal Access Token for that server (\*or `CONFLUENCE_USER` + `CONFLUENCE_PASSWORD`) |
| `CONFLUENCE_NAME` | no | e.g. `Green`. Only matters if you configure a second server; defaults to `Primary` |
| `CONFLUENCE_BASE_URL_2` | no | Leave unset for a single-server setup |
| `CONFLUENCE_TOKEN_2` | with `_BASE_URL_2` | The second server's own token |
| `CONFLUENCE_NAME_2` | no | e.g. `Blue` — say it in a prompt to query that server; defaults to `Secondary` |
| `CONFLUENCE_ALLOW_WRITE` | no | `true` to let it create, update and append to pages. Blank = read-only |

The Python interpreter and the folder saved pages go to come from the shared
environment variables in [Configuration](#configuration). Every other setting is
listed under [This server's own settings](#this-servers-own-settings).

```json
{
  "env": {
    "CONFLUENCE_BASE_URL": "https://confluence.internal.example.com",
    "CONFLUENCE_TOKEN":    "token-for-the-first-server"
  }
}
```

**Credentials are env-var only**: there is no flag that could put a token in a
command line, where other local users would see it in a process listing. As
Windows user environment variables instead:

```powershell
setx CONFLUENCE_BASE_URL "https://confluence.internal.example.com"
setx CONFLUENCE_TOKEN    "token-for-the-first-server"
setx CONFLUENCE_TOKEN_2  "token-for-the-second-server"   # only if you have two
```

`setx` does not affect processes that are already running, so quit VS Code /
Claude Code completely (a window reload is not enough) and reopen it. Check it
took in a **new** window with `$env:CONFLUENCE_TOKEN`.

## Macros

Most of what makes a Confluence page useful sits inside a macro, and macros come
in two kinds. The difference decides where the content lives, and so how it has
to be fetched.

| Kind | Examples | Where the content is |
|---|---|---|
| **In the page source** | info / note / warning / tip panels, expand, code, inline task lists, status lozenges, excerpts | Confluence stores the content itself |
| **Generated when the page is displayed** | Task Report, Page Properties Report, Children Display, Page Tree, Jira Issues, Include Page, Excerpt Include, charts, diagrams | The source holds only the macro's **settings**; Confluence builds the content each time the page is opened |

This server reads both. Page bodies come back as **Markdown**:

- inline tasks become `- [ ]` / `- [x]` checkboxes, with the assignee as
  `@username` and the due date kept
- info / note / warning / tip / panel become block quotes with their title
- **expand macros are shown open**, so content hidden behind a toggle is read
- code macros become fenced blocks, with the language
- status lozenges become `**[DONE]**`
- tables stay tables, so a report reads as rows rather than a run-on sentence
- links to pages, users and attachments become readable links and `@mentions`

For the generated kind, the server asks **Confluence** to render the page and
reads the result — the only way that content can be obtained, because Confluence
is what produces it.

So "list all the tasks on page X for Jane" works whether the page carries inline
task checkboxes or a Task Report macro pointed at her.

### Which body gets read

`CONFLUENCE_BODY_FORMAT` (and the per-call `body_format` argument on both page
tools) chooses:

| Value | Reads |
|---|---|
| `auto` **(default)** | The page source for an ordinary page, and the **rendered** page whenever the page uses a macro whose content is generated. Best of both: no display chrome on ordinary pages, full content where it matters |
| `view` | Always the rendered page — what the browser shows |
| `export_view` | The render Confluence uses for PDF/Word export. Try this if a macro is **still** empty under `view`: a Page Tree, and some third-party macros, only render statically on export |
| `storage` | The raw page source only, with no macro output at all (how this server behaved before v5.0.0) |

Every page read reports which one it used on a `Body:` line, and a page read
from the source names each macro whose content was **not** fetched, with its
settings — so an empty section is visible as missing content rather than as an
empty page. If that happens, ask for the page again with `body_format="view"`,
then `body_format="export_view"`.

Unknown macros count as generated on purpose. A third-party macro this server
has never heard of triggers a rendered fetch rather than reading as empty.

## Attachments

`confluence_list_attachments` shows the files on a page — name, type, size,
version — and where each would be downloaded to.
`confluence_download_attachment` fetches one into **the folder the plugin for
that file type reads**, because each document plugin is confined to one folder
and a download anywhere else could never be opened:

| Attachment | Downloaded to | Open it with |
|---|---|---|
| `.xlsx`, `.xlsm` | `%EVA_DOCUMENTS_DIR%\excel` | `excel` (`excel_list_sheets`, `excel_read_range`, `excel_read_table`) |
| `.docx` | `%EVA_DOCUMENTS_DIR%\word` | `word` (`msword_open`) |
| `.pptx` | `%EVA_DOCUMENTS_DIR%\powerpoint` | `powerpoint` (`powerpoint_open`) |
| `.pdf` | `%EVA_DOCUMENTS_DIR%\pdf` | `pdf-to-md` (`convert_pdf_to_markdown`) |
| `.md` | `%EVA_KNOWLEDGE_DIR%\confluence` | `knowledge-base` (`kb_index`) |

So a request like *"get the budget spreadsheet attached to the FY26 Planning
page, then use Excel to total the Travel column"* is two calls: the download,
which reports the saved file name, then an `excel_*` tool on that name. Other
types are listed but not downloaded.

- **A file already in the folder is never replaced** unless the call passes
  `overwrite=true` — ask for "the latest version" and Claude will. `save_as`
  saves under another name; the attachment's extension is always kept, and the
  name cannot climb out of the folder.
- Downloading writes to the local disk only, so it works with page writing
  switched off.
- The documents root must exist (copy `eva\` to `H:\Eva`), or document
  downloads are disabled with a warning at startup; the per-type sub-folder is
  created if missing. `CONFLUENCE_DOCS_DIR=off` switches document downloads off
  outright.

## Writing pages

**Off unless `CONFLUENCE_ALLOW_WRITE=true`.** With it off the write tools are
not offered at all, and the server sends nothing but GET requests. With it on:

| Tool | Does |
|---|---|
| `confluence_create_page` | A new page in a space, optionally under a parent (`parent_id` or `parent_title`) |
| `confluence_update_table` | Changes **cells of one table in place** — finds rows by a column value and sets other columns, and/or adds rows — so the table keeps its widths, colours and merged cells, and any macro it sits in (table filter, column, section, expand) is untouched. See [Tables](#tables) |
| `confluence_update_section` | Changes **one section** — found by its heading text, or a panel/expand title — and sends the rest of the page back byte-for-byte. `mode`: `replace` (default), `append` or `prepend`. The right tool for "update the Director's notes on page 1234" |
| `confluence_append_to_page` | Adds content to the **end** (or `position: "start"`) of a page, keeping everything already on it byte-for-byte — macros, layouts and all. The safe way to add minutes, actions or a new section |
| `confluence_update_page` | **Replaces** a page's whole body and/or its title. Anything left out of the new body — including macros such as a task report — is gone from the page (Confluence keeps the old version in the page history). For a full rewrite only |

**What counts as a section.** A heading and everything under it — sub-headings
included — up to the next heading of the same or a higher level, or the end of
the layout column it sits in; or the body of an info / note / panel / expand
macro whose title matches. Matching ignores case, a trailing colon and curly vs
straight quotes, so `director's notes` finds **Director’s Notes:**. A name that
matches nothing, or more than one section, is refused with the list of sections
the page has. If the section being replaced itself contains a macro (say a Jira
table under **Risks**), the call is refused and names it, unless
`allow_macro_removal=true` — use `mode: "append"` to add beside it instead.

Content is written in **Markdown** and converted to Confluence formatting:

| Markdown | Becomes |
|---|---|
| `# Heading`, paragraphs, `**bold**`, `*italic*`, `~~struck~~`, `` `code` ``, `[link](url)` | the matching Confluence formatting |
| `- item` / `1. item` (nested by indent) | bullet / numbered lists |
| `- [ ] task` / `- [x] done` | real Confluence **inline tasks** (numbering continues after any tasks already on the page) |
| `\| a \| b \|` tables | tables, first row as the header |
| ```` ```python ```` fenced code | the **code** macro, language kept where Confluence knows it |
| `> [!NOTE] Title` quote | an **info** panel (`[!TIP]` tip, `[!IMPORTANT]` note, `[!WARNING]` / `[!CAUTION]` warning) |
| `![alt](diagram.png)` | an image — an `http(s)` URL, or else an attachment on the page |

`content_format: "storage"` sends raw Confluence storage-format XHTML instead,
for a macro Markdown cannot express.

Every write saves a **new page version**, with an optional `version_message`
and `minor_edit` (no watcher notifications). If someone saved the page in the
meantime, Confluence refuses the write (HTTP 409); `confluence_update_page`
also takes `expected_version`, the version you read, and refuses before sending
anything if the page has moved on. Confluence's own permissions still decide
which spaces the account can edit.

## Tables

A table is often wrapped in a macro — a **table filter**, a **column** or
**section** layout, an **expand** — and replacing the section around it would
either delete that macro or rebuild the table from Markdown, losing its
formatting. So tables have their own pair of tools:

- **`confluence_list_tables`** (always available) lists every table stored on a
  page, whatever it is nested in: its **number**, the **heading** above it, the
  **macros around it**, its **column names** and its rows. Merged cells show as
  `^`; a cell holding a macro (a status lozenge, a nested table) is marked
  `[macro]`. Pass `table` to see one table with up to 200 rows.
- **`confluence_update_table`** (writing on) edits one table **in place**. Only
  the text inside the cells you name changes; every tag of the table and every
  macro around it goes back exactly as it was. A cell's own formatting stays
  too: `<strong>100</strong>` in a red-highlighted cell becomes
  `<strong>120</strong>` in the same red cell.

```json
{
  "page_id": "1234",
  "table": "Budget",
  "updates": [
    {"match": {"Item": "Travel"}, "set": {"Q3": 120, "Q4": 95}},
    {"match": {"Item": "Hotels"}, "set": {"Q3": "1,250"}}
  ],
  "add_rows": [{"Item": "Training", "Q3": 40, "Q4": 40}]
}
```

- **Picking the table:** `table` is its number (from `confluence_list_tables`)
  or the heading above it; leave it out when the page has one table. A table
  nested in another's cell never wins a heading match over the outer table.
- **Columns** are named by the header row. A two-row header — **FY26** spanning
  **Q3** and **Q4** — gives `FY26 Q3` / `FY26 Q4`, and `Q3` alone works when
  only one column has it. A 1-based column number works too.
- **Rows** are found by `match` (every column given must match). Text matching
  ignores case and spacing; numbers compare as numbers, so `1250` matches
  `1,250`. More than one matching row is refused unless that update sets
  `"all_matches": true`.
- **All or nothing.** Every change is checked before any is made. An unknown
  column, a row that matches nothing, an ambiguous match, a merged cell (it
  would change two rows), or a cell holding a macro (unless
  `allow_macro_removal: true`) stops the call, and the page is not saved.
- **New rows** go at the bottom and copy the last row's cells, so they pick up
  its formatting; a macro in that row is not copied. If the last row has merged
  cells, a plain row is added instead.
- Values may use inline Markdown (`**bold**`, `[link](url)`); `expected_version`
  refuses the save if the page changed since you read it.
- **Not editable:** a table that Confluence builds when the page is displayed —
  Jira issues, a page properties report, a CSV macro. It has no stored cells;
  `confluence_list_tables` names such macros so this is visible.

## Saving to the knowledge base

**Reading a page does not save it.** Both page tools take a `save_to_kb`
argument, false by default; the page is written to the knowledge-base folder
only when it is true, which Claude sets when you ask for the page to be kept:

| You say | What happens |
|---|---|
| "What does the incident runbook say about escalation?" | Pages are searched and read to answer. Nothing is saved |
| "Save the incident runbook to the knowledge base" | `save_to_kb=true` — one Markdown file, then `kb_index` can pick it up |
| "Pull the whole onboarding tree into the KB" | One saved file per page you asked for |

That is the difference between a knowledge base of pages you chose and one
holding every page skimmed along the way. A search for "retention policy" that
turns up somebody's meeting notes used to file those notes in the index, where
they came back later as an answer.

The saved file is the full page — `CONFLUENCE_MAX_BODY` truncates only what is
returned to the model — under `Confluence - <title>.md`, overwriting any earlier
copy of the same page. If saving is switched off (`off`) and you ask for a page
to be kept anyway, the tool says so instead of failing silently.

To go back to saving every page read, as versions before 3.0.0 did, set
`CONFLUENCE_KB_AUTOSAVE=true`:

```powershell
setx CONFLUENCE_KB_AUTOSAVE "true"
```

## Two Confluence servers

Set the second base URL and the two instances sit behind the same set of tools.
Each server has a name, and the name is how a prompt picks one:

| You say | Server used |
|---|---|
| "Search Confluence for the incident runbook" | **Green** — the first server, always the default |
| "Find content about onboarding on Blue" | **Blue** — because the prompt named it |
| "Check both wikis for the retention policy" | Both — Claude runs the search once per server |

Naming is case-insensitive (`blue` works), and the tools also accept `1`/`2`.
Name a server that is not configured and you get an error listing the ones that
are — it never quietly falls back to the other instance and answers from the
wrong wiki.

Two things change once a second server is configured, and only then:

- **Output says which server it came from** — search results carry
  `server=Blue`, pages carry a `Server: Blue` line. Content IDs are per-instance
  (page `393217` on Green is a different page from `393217` on Blue), so results
  have to be attributable.
- **Knowledge-base files are named `Confluence <server> - <title>.md`** instead
  of `Confluence - <title>.md`, so the same page title on both instances
  produces two files rather than one overwriting the other.

With one server configured, everything behaves exactly as it did before: no
`server` argument on the tools, no labels in the output, and knowledge-base
files keep their original names.

The second server is all-or-nothing. Set `CONFLUENCE_BASE_URL_2` without
credentials and the server refuses to start rather than quietly answering "on
Blue" questions from the first instance. Set the second token *without* the
second base URL and it warns, then runs with one server.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once (in settings.json's `env` block or for your Windows account) and this
plugin has nothing else to configure - it prompts for nothing at install and
takes no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `H:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `H:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `H:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "H:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "H:\Eva\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "H:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

Of the four, this server uses three: `EVA_PYTHON`, `EVA_KNOWLEDGE_DIR`, and
`EVA_DOCUMENTS_DIR` for downloaded attachments. It reads no local folder at all -
pages come over HTTP.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `confluence`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `H:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_KNOWLEDGE_DIR%\confluence` | Where a page is saved as Markdown **when a tool call asks for it** (`save_to_kb=true`) - `Confluence - <title>.md`, or `Confluence <server> - <title>.md` with two instances - for the `knowledge-base` plugin to index. Downloaded `.md` attachments land here too | Created on demand |
| `%EVA_DOCUMENTS_DIR%\excel`, `\word`, `\powerpoint`, `\pdf` | Where a downloaded attachment goes, by file type - the same folders the `excel`, `word`, `powerpoint` and `pdf-to-md` plugins read ([Attachments](#attachments)) | The root must exist, or document downloads are disabled (warned at startup); a sub-folder is created on demand |

### This server's own settings

Everything else is an environment variable of this server's own. **Credentials
are env-var only** - there is no flag that could leak a token into a process
listing.

#### First server (the default one)

| Env var | Purpose |
|---|---|
| `CONFLUENCE_BASE_URL` | Base URL incl. any context path, no trailing slash. **Required** |
| `CONFLUENCE_NAME` | Friendly name used to pick this server in a prompt (default `Primary`) |
| `CONFLUENCE_TOKEN` | Personal Access Token, sent as Bearer (preferred over basic auth) |
| `CONFLUENCE_USER` | Username for basic auth (fallback if no token) |
| `CONFLUENCE_PASSWORD` | Password for basic auth |
| `CONFLUENCE_CA_CERT` | Path to a PEM CA bundle for an internal CA |
| `CONFLUENCE_VERIFY_SSL=false` | Disable TLS certificate verification |

#### Second server (optional)

Setting `CONFLUENCE_BASE_URL_2` is what enables it.

| Env var | Purpose |
|---|---|
| `CONFLUENCE_BASE_URL_2` | Base URL of the second instance |
| `CONFLUENCE_NAME_2` | Friendly name, e.g. `Blue` (default `Secondary`) |
| `CONFLUENCE_TOKEN_2` | Its own Personal Access Token — tokens are per-instance |
| `CONFLUENCE_USER_2` | Username for basic auth on the second server |
| `CONFLUENCE_PASSWORD_2` | Password for basic auth on the second server |
| `CONFLUENCE_CA_CERT_2` | CA bundle for the second server; falls back to `CONFLUENCE_CA_CERT` |
| `CONFLUENCE_VERIFY_SSL_2=false` | TLS verification for the second server; falls back to the first server's setting |

#### Shared by both servers

| Env var | Purpose |
|---|---|
| `CONFLUENCE_TIMEOUT` | Request timeout in seconds (default 30) |
| `CONFLUENCE_BODY_FORMAT` | Which version of a page body to read: `auto` (default), `view`, `export_view` or `storage` — see [Macros](#macros) |
| `CONFLUENCE_MAX_BODY` | Truncate page bodies to N chars, 0 = unlimited (default). Applies only to text returned to the model, not to saved files |
| `CONFLUENCE_KB_DIR` | Full path to the save folder, instead of `%EVA_KNOWLEDGE_DIR%\confluence`. `off` forbids saving outright, after which the server writes no local file at all |
| `CONFLUENCE_KB_AUTOSAVE=true` | Save **every** page read, without being asked (default false). Needs a save folder to be on |
| `CONFLUENCE_ALLOW_WRITE=true` | Offer the page-writing tools (create, update, append) — see [Writing pages](#writing-pages). Default off: read-only. Applies to both servers |
| `CONFLUENCE_DOCS_DIR` | The documents **root** attachments download into, instead of `EVA_DOCUMENTS_DIR`; the per-type sub-folders (`excel`, `word`, `powerpoint`, `pdf`) are still appended, so point it at the same root the document plugins use. `off` forbids document downloads (Markdown attachments still go to the knowledge folder) |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies. To forbid saving outright, set `CONFLUENCE_KB_DIR=off`
(`none`, `no`, `false` and `disabled` work too).

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Connect to every configured server, print who you are authenticated as + visible space count, the attachment folder and whether page writing is on, to stderr, then exit (no server). Non-zero if any server fails |
| `--version` | Print version and exit |

## File access

No local file access until a tool call asks for it: a saved page writes one
Markdown file inside the knowledge-base folder, and a downloaded attachment one
file inside the documents folder for its type (or the knowledge folder, for
Markdown). An existing file is never replaced unless the call says
`overwrite=true`, and a file name can never reach outside those folders.

## Usage examples

1. "Search Confluence for our incident response runbook." → `confluence_search` on the default server
2. "Find pages about the onboarding process on Blue." → `confluence_search` with `server="Blue"`
3. "Find pages in the DOCS space that mention 'release notes' and were updated in the last 30 days." → `confluence_search_cql`
4. "Pull up the full content of Confluence page 393217." → `confluence_get_page`
5. "Open the 'Q3 Roadmap' page in the PROD space and summarise it." → `confluence_get_page_by_title`
6. "List every page under the 'Engineering Handbook' in the DOCS space, direct children only." → `confluence_list_pages_under`
7. "Is the retention policy on Green the same as the one on Blue?" → the same tool called once per server, then compared
8. "Pull the onboarding runbook into our local knowledge base for offline search." → `confluence_get_page` (or `confluence_get_page_by_title`) with `save_to_kb=true`, which writes the Markdown copy the `knowledge-base` plugin's `kb_index`/`kb_ask` can find afterwards
9. "Summarise the release notes page." → the same tools **without** `save_to_kb` — you get the summary and nothing lands in the knowledge base
10. "List all the tasks assigned to Jane on page 393217." → `confluence_get_page`; inline task checkboxes and a Task Report macro both come back as rows
11. "That page looks empty but it has a task report on it." → the same tool again with `body_format="view"`, then `body_format="export_view"`
12. "What files are attached to the FY26 Planning page?" → `confluence_list_attachments`
13. "Get the budget spreadsheet from the FY26 Planning page and total the Travel column." → `confluence_download_attachment` (lands in `documents\excel`), then `excel_read_table` / `excel_column_stats` on the file it reports
14. "Create a page under Team Home called 'Offsite 2026' with this agenda." → `confluence_create_page` with `parent_title` *(writing on)*
15. "Add today's actions to the bottom of the project page." → `confluence_append_to_page`, with the actions as `- [ ]` tasks *(writing on)*
16. "Rewrite the onboarding page with this new text." → `confluence_get_page`, then `confluence_update_page` with `expected_version` *(writing on)*
17. "Update the Director's notes on page 1234 with this." → `confluence_update_section` with `section: "Director's notes"` — the task report and Jira tables elsewhere on the page are untouched *(writing on)*
18. "Update the Q3 figures in the Budget table on page 1234: Travel 120, Hotels 1,250." → `confluence_list_tables`, then `confluence_update_table` — works even though the table sits inside a Table Filter macro *(writing on)*

## Troubleshooting

> **If a server fails with `Executable not found in $PATH: "${EVA_PYTHON}"`**,
> the variable is not set in the environment Claude Code was launched from. Set
> it (see above), then quit Claude Code completely and reopen — `setx` and
> `[Environment]::SetEnvironmentVariable` do not reach a process that is already
> running.

`--check` connects to **every** configured server, authenticates and reports who
you are plus how many spaces you can see. Run it before wiring the server in; it
exits non-zero if any instance fails, so a two-server setup where only one
answers is caught here rather than mid-conversation:

```powershell
& $env:EVA_PYTHON confluence.py --check
```

To try a pair of servers without touching the plugin config (every setting,
tokens included, goes in the environment):

```powershell
$env:CONFLUENCE_NAME       = "Green"
$env:CONFLUENCE_BASE_URL   = "https://green.confluence.example.com"
$env:CONFLUENCE_TOKEN      = "green-token"
$env:CONFLUENCE_NAME_2     = "Blue"
$env:CONFLUENCE_BASE_URL_2 = "https://blue.confluence.example.com"
$env:CONFLUENCE_TOKEN_2    = "blue-token"
& $env:EVA_PYTHON confluence.py --check
```

Behind an internal CA, point `CONFLUENCE_CA_CERT` at the PEM bundle rather than
reaching for `CONFLUENCE_VERIFY_SSL=false`. If the two instances sit behind
different CAs, `CONFLUENCE_CA_CERT_2` covers the second one; leave it unset and
the second server uses the first's bundle.

### A macro's content is missing

Read the `Body:` line at the top of the page output — it says which version of
the body was used, and a page read from the source lists every macro whose
content was not fetched.

1. Ask for the page again with `body_format="view"`. That forces Confluence to
   render the page.
2. Still empty? Try `body_format="export_view"`. A Page Tree, and some
   third-party macros, only render statically on export.
3. Still empty after both? The macro's content is genuinely not reachable
   through the REST API on this instance — usually a macro that renders in the
   browser rather than on the server, or one whose plugin is not licensed. The
   placeholder names the macro, so it can be checked in Confluence itself.

If **every** page reads as source-only, check `CONFLUENCE_BODY_FORMAT` — a value
of `storage` turns macro rendering off entirely. `--check` prints the setting in
use on its `Page body format` line.
