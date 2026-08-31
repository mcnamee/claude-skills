# Claude Skills

Three things, all for Claude Code, all working **entirely offline**:

- **[`plugins/`](plugins)** — eight MCP servers that give an agent hands on the
  things enterprise work actually lives in: Word documents, Excel workbooks,
  Outlook mail, Confluence, Jira, PDFs, plus a local RAG knowledge base to tie
  them together. This repo doubles as a plugin marketplace for them.
- **[`skills/`](skills)** — standalone skills that need no server and no
  install beyond a folder copy.
- **[`agents/`](agents)** — subagents that take a whole job away and hand back
  a finished result, built on top of the servers and skills above.

Plus **[`eva/`](eva)** — the working folder they all read, write and index,
carried here as a scaffold you copy to `C:\Eva`. It holds the knowledge base,
the document library, generated output, and the reference material: **exemplars**
(finished documents showing what good looks like) and **templates** (the blank
`.docx`/`.pptx` new documents are built from).

## Why this exists

- **Built for Windows endpoints in an Enterprise environment.** No internet calls, no telemetry,
  no registry lookups. Copy the folder across, add it as a local marketplace,
  install.
- **One file per server.** Every server is a single `.py` — nothing to build, no
  package tree to transfer. Three of the eight are **standard library only**, and
  the standalone skills have no code at all.
- **Install with prompts, not JSON.** `/plugin install` asks for the folders and
  the Python interpreter instead of you hand-editing absolute paths in eight
  places. The matching skill comes with the server.
- **Confined by default.** Every server that touches the filesystem is locked to
  the folders you name, and refuses to start unconfined rather than falling back
  to "anywhere". `word` and `powerpoint` are the only ones that can change a
  file you already have; the rest either read, or write new Markdown into a
  folder you nominate.
- **Secrets never hit the command line.** Tokens and API keys are environment
  variables only — argv is visible to other local users in process listings.
- **They compose.** Word, Outlook and Confluence can each mirror what they read
  into one Markdown folder; `pdf-to-md` fills the same folder from PDFs; the
  `knowledge-base` server indexes it and answers questions over the lot.
- **And it grows.** `word` mirrors documents it *writes*, not just ones it
  reads, and `knowledge-base` takes a `kb_capture` call — so an analysis or a
  research brief that would otherwise vanish with the chat goes back into the
  index and answers the same question next time.

## The plugins

| Plugin | Version | What it does | pip install |
|---|---|---|---|
| [**word**](plugins/word) | 5.1.0 | Read, edit and create `.docx` — real Word tracked changes, native styles, filling out templates | `python-docx` |
| [**powerpoint**](plugins/powerpoint) | 1.0.0 | Build `.pptx` decks that inherit your own template's layouts and theme, and audit them against the 10/20/30 rule | `python-pptx` |
| [**excel**](plugins/excel) | 4.0.0 | Read and analyse workbooks; parses `.xlsx` directly, so Excel isn't needed | _none_ |
| [**outlook**](plugins/outlook) | 4.0.0 | Read local Outlook mail and calendar via COM, with a content blacklist | `pywin32` |
| [**confluence**](plugins/confluence) | 2.0.0 | Search and read Confluence pages, across one or two instances | _none_ |
| [**jira**](plugins/jira) | 1.1.3 | Query issues, sprints and projects (Jira Data Center v2 API) | _none_ |
| [**knowledge-base**](plugins/knowledge-base) | 3.0.0 | True RAG over your own Markdown: local ChromaDB index + your embeddings API, and capture notes back into it | `chromadb` |
| [**pdf-to-md**](plugins/pdf-to-md) | 5.1.0 | Convert PDFs to Markdown with tables preserved | `pymupdf pymupdf4llm` |

Each plugin's README covers its settings, tools, file access and example
prompts. Every server also carries a semantic version in `__version__`, printed
by `--version` and reported to the MCP client in `serverInfo`.

## Install

Everything below is **Windows / PowerShell**, which is what this suite targets.
Two things to know if you're pasting from elsewhere: a quoted path at the start
of a command needs the call operator (`& "C:\...\python.exe"`), and `%VAR%`
does not expand — it's `$env:VAR`.

**1. Lay out the working folder.** Copy the repo's [`eva/`](eva) folder to
`C:\Eva`:

```powershell
Copy-Item -Recurse C:\path\to\claude-skills\eva C:\Eva
```

Every folder setting in every plugin already defaults to its place in that tree,
so this one step means you can accept each folder prompt as it stands and the
servers are correctly related to each other — mirrors landing inside the indexed
corpus, generated documents kept out of your source library. See
[Folder layout](#folder-layout) for what goes where, and `eva/README.md` for
each folder's own README.

**2. Install the pip dependencies** into the *same* interpreter you'll point the
plugins at:

```powershell
& "C:\path\to\python.exe" -m pip install python-docx pymupdf pymupdf4llm pywin32 chromadb
```

(Drop `pywin32` if you're not on Windows / not using `outlook`. Install only
what the plugins you want need — see the table above.) `word.py`'s docstring
walks through sideloading the wheels.

**3. Add this repo as a marketplace**, then install whichever plugins you want.
These are slash commands, typed inside Claude Code — not shell commands:

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install word@mcnamee-claude-skills
/plugin install excel@mcnamee-claude-skills
```

Claude Code prompts for that server's settings. Every folder prompt is
pre-filled from step 1, so the only answer that is genuinely yours to give is
the **Python interpreter** — the absolute path to the `python.exe` from step 2;
a mismatch here is the most common cause of "dependency missing". The plugins
are independent, so a machine without `pywin32` simply doesn't install
`outlook`.

Leaving a folder prompt **blank** means "not configured", so the default
applies. To switch an optional folder off, type `off`.

`excel` is the simplest to start with: standard library only, one prompt.

**4. Set your secrets** as Windows user environment variables before starting
Claude Code — they're read from the ambient environment, never stored in the
plugin. Only needed for the plugins you actually install:

```powershell
setx CONFLUENCE_TOKEN "your-personal-access-token"
setx JIRA_TOKEN       "your-personal-access-token"
setx KB_EMBED_API_KEY "your-api-key"
```

| Plugin | Environment variable |
|---|---|
| `confluence` | `CONFLUENCE_TOKEN`, plus `CONFLUENCE_TOKEN_2` if you point it at a second Confluence instance |
| `jira` | `JIRA_TOKEN` |
| `knowledge-base` | `KB_EMBED_API_KEY` |

Two `setx` gotchas. It **doesn't affect processes that are already running**, so
quit VS Code completely — a window reload is not enough — and reopen it. And it
**truncates values at 1024 characters**; if you ever hit that, set the variable
through System Properties → Environment Variables instead.

To check a variable took, in a **new** window: `$env:JIRA_TOKEN`. To set one for
the current session only (handy for testing, gone when you close the window):
`$env:JIRA_TOKEN = "..."`.

Useful commands: `/plugin` to browse and manage, `/mcp` to confirm a server
connected, `claude mcp list` to spot an unresolved environment variable, and
`/plugin marketplace update mcnamee-claude-skills` after you transfer a new
version across.

> **Skills are namespaced** by their plugin, so it's `/word:word` rather
> than `/word`. To invoke one by its bare name, install the skill on its own
> instead — see [Skills](#skills) below.
>
> **Plugins are cached on install** (copied under `~/.claude/plugins/`), which is
> why each plugin contains its own server file rather than sharing one — a path
> reaching outside the plugin would break once cached.

### Before you wire anything in

Run the server's `--check` first. It validates config (and connectivity, for the
HTTP servers) without starting the server, and is far easier to read than an MCP
connection failure:

```powershell
& "C:\path\to\python.exe" C:\path\to\claude-skills\plugins\word\word.py --check
```

### Manual install, without plugins

If you'd rather configure a server directly — or want one configured differently
from what its plugin prompts for — register it with `claude mcp add --scope user`
(available in every folder), or copy [`.mcp.json.example`](.mcp.json.example) to
`.mcp.json` in the folder you open Claude Code in (config travels with the
files). Keep `PYTHONUTF8=1`: without it, Windows' legacy codepage can corrupt the
stdio JSON stream on non-ASCII content. Pass secrets with `-e` / the `env` block,
never as flags.

```powershell
claude mcp add excel --scope user -e PYTHONUTF8=1 -- C:\path\to\python.exe C:\path\to\claude-skills\plugins\excel\excel.py --docs-dir C:\path\to\your\workbooks
```

(No `&` needed there — `claude` is unquoted, so PowerShell treats it as a
command already.)

## Configuration conventions

Every server follows the same pattern, so once you know one you know them all.
The per-plugin READMEs list each server's actual settings.

1. **Precedence: CLI flag > environment variable > constant in the file.**
2. **Naming: the env var is the server's prefix + the flag name.** `--docs-dir`
   on `excel.py` is `EXCEL_DOCS_DIR`; on `word.py` it is
   `MSWORD_DOCS_DIR`. Prefixes: `CONFLUENCE_`, `JIRA_`, `KB_`, `EXCEL_`,
   `OUTLOOK_`, `MSWORD_`, `PDF2MD_`. The one deliberate exception: `--insecure`
   pairs with `<PREFIX>_VERIFY_SSL=false`.
3. **Secrets are env-var only.** No `--token`/`--password`/`--*-api-key` flags
   anywhere, because command-line arguments are visible to other local users.
4. **Shared flag vocabulary:** `--docs-dir` (the source folder a server is
   confined to), `--output-dir` (generated files), `--templates-dir` (blank
   templates to create from, read-only), `--kb-dir` (Markdown mirror
   for the RAG knowledge base), `--index-dir` (the vector store),
   `--base-url`/`--ca-cert`/`--insecure`/`--timeout`/`--max-body` (the HTTP
   servers), `--check`, `--version`.
5. **Every folder setting defaults to its place in `C:\Eva`** — see
   [Folder layout](#folder-layout). A **blank** value means "not configured",
   so the default applies; to switch an optional folder off, pass `off`
   (`none`, `no`, `false` and `disabled` also work). A folder you configured
   yourself that does not exist is a fatal error, because it is almost always a
   typo; a built-in default that does not exist yet is a warning, and the
   feature it enables simply stays off.

## Folder layout

One working folder, `C:\Eva`, holds everything the servers read, write and
index. The repo carries it as a scaffold — [`eva/`](eva) is the same tree with a
README in every folder and no content, so `Copy-Item -Recurse ...\eva C:\Eva`
lays it out in one step and every default below is already correct.

```
C:\Eva\
├─ knowledge\        the RAG corpus - Markdown only, the ONE indexed root
│  ├─ notes\           Markdown you write by hand
│  ├─ captures\        notes kb_capture writes back
│  ├─ confluence\      pages the confluence plugin mirrored
│  ├─ email\           emails the outlook plugin mirrored
│  ├─ word\            documents the word plugin mirrored
│  ├─ powerpoint\      decks the powerpoint plugin mirrored
│  └─ pdf\             PDFs the pdf-to-md plugin converted
├─ index\            the ChromaDB vector store - derived, disposable
├─ documents\        the binary library the servers READ
│  ├─ word\            .docx (searched recursively; inbox\ + library\)
│  ├─ powerpoint\      .pptx (searched recursively)
│  ├─ excel\           .xlsx (top level only - excel does not recurse)
│  └─ pdf\             source PDFs
├─ output\           what the assistant creates
│  ├─ word\
│  └─ powerpoint\
└─ reference\        style, not facts - deliberately NOT indexed
   ├─ exemplars\       finished documents showing what good looks like
   └─ templates\       blank branded files new documents/decks start from
```

| Folder | Plugin → setting |
|---|---|
| `knowledge\` | `knowledge-base` → `--docs-dir` |
| `knowledge\captures\` | `knowledge-base` → `--output-dir` |
| `knowledge\confluence\` | `confluence` → `--kb-dir` |
| `knowledge\email\` | `outlook` → `--kb-dir` |
| `knowledge\word\` | `word` → `--kb-dir` |
| `knowledge\powerpoint\` | `powerpoint` → `--kb-dir` |
| `knowledge\pdf\` | `pdf-to-md` → `--output-dir` |
| `index\` | `knowledge-base` → `--index-dir` |
| `documents\word\` | `word` → `--docs-dir` |
| `documents\powerpoint\` | `powerpoint` → `--docs-dir` |
| `documents\excel\` | `excel` → `--docs-dir` |
| `documents\pdf\` | `pdf-to-md` → `--docs-dir` |
| `output\word\` | `word` → `--output-dir` |
| `output\powerpoint\` | `powerpoint` → `--output-dir` |
| `reference\templates\` | `word` and `powerpoint` → `--templates-dir` |

Two rules make it hold together, and both are worth knowing before you move a
folder:

**There is one indexed root.** `knowledge\` is the only folder the RAG index
reads, so every server that produces Markdown writes inside it. Point a mirror
anywhere else and it fills up faithfully while `kb_ask` never sees a word of it.

**Folders are named after what wrote them, not what they are about.** Topic
folders rot — every document belongs to three of them. Provenance maps
one-to-one onto a setting, retrieval is semantic anyway, and cleanup stays
surgical: delete `knowledge\confluence\` and re-mirror, with nothing you wrote
yourself at risk.

[`eva/README.md`](eva/README.md) covers the reasoning, how to put the tree on
another drive, and why nothing in it is committed. Each folder's own README says
what belongs there.

## File access policy

Every server that touches the filesystem is confined to the folder(s) named in
its configuration, and that configuration is **required**:

| Plugin | Local file access |
|---|---|
| `word` | Read/write, confined to the documents folder plus the output and knowledge-base folders; the templates folder is read-only. Opening, creating and saving each mirror to the knowledge-base folder |
| `powerpoint` | Read/write, confined to the presentations folder plus the output and knowledge-base folders; the templates folder is read-only. Opening, creating and saving each mirror to the knowledge-base folder |
| `excel` | Read-only, confined to the workbook folder (top level only) |
| `knowledge-base` | Reads the documents folder; writes its vector index (`C:\Eva\index`) and captured notes (`C:\Eva\knowledge\captures`, always inside the documents folder); never edits or deletes an existing document; network only to the endpoints you configure |
| `pdf-to-md` | Reads the PDF folder, writes the output folder |
| `confluence` | Writes only the knowledge-base folder — every page read is mirrored there. Set it to `off` and the server touches no local file |
| `outlook` | Writes only the knowledge-base folder — every email read is mirrored there, blacklisted messages excepted. Set it to `off` and the server touches no local file |
| `jira` | None — HTTP GET to Jira only |

Paths are resolved (symlinks included) before the containment check, so a
symlink dropped inside a configured folder cannot reach files outside it.

## Skills

There are two kinds, in two places.

**Every plugin ships one**, at `plugins/<name>/skills/<name>/SKILL.md` — it
teaches an agent that server's tools, the right call order, and the sharp
edges (read-only limits, sandboxes, the tracked-changes workflow, when to
reindex). Installing the plugin installs its skill, namespaced as
`/<plugin>:<skill>` — so `/word:word`, not `/word`.

`powerpoint` is the exception that ships **two**: `/powerpoint:powerpoint` for
the server's mechanics, and `/powerpoint:kawasaki` for Guy Kawasaki's 10/20/30
rule — 10 slides, 20 minutes, a 30-point minimum font. They are split because
the rule is a way of thinking about any presentation, in any tool, while the
other is about driving this server; the rule fires when someone asks for a deck
at all, and reaches for the server's audit only when it is there.

**[`skills/`](skills) holds standalone skills**, which need no server and no
plugin. Install one by copying its folder:

| Skill | Invoke | What it does |
|---|---|---|
| [**polish**](skills/polish) | `/polish` | Rewrites a draft into Australian Public Service style — the Australian Government Style Manual — asking who the reader is and what the medium is, then picking the register from them |
| [**unslop**](skills/unslop) | `/unslop` | Strips AI-slop markers from writing — padding, tell-tale vocabulary, stock LLM sentence shapes — leaving meaning and voice intact |

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\unslop $dest
```

Or copy it into `.claude\skills\` inside a project to scope it there. Run
`/doctor` or restart Claude Code if a newly copied skill doesn't appear. See
[`skills/README.md`](skills/README.md) for the details and for how to add
another.

Note the difference in what they give Claude: a **plugin** adds *tools*, a
**skill** adds *instructions*. A server's skill describes how to use it — it
isn't a way to run one. These are MCP servers, launched as long-running stdio
subprocesses, not scripts a skill shells out to. That matters most for `word`
and `powerpoint`, which are session-based (`msword_open` / `powerpoint_create`
return a `session_id` and hold the file in memory until it is saved).

## Agents

**[`agents/`](agents) holds standalone subagents** — a separate context with its
own instructions that the main session hands a whole job to, and gets a finished
result back from. Where a skill steers the conversation you are already in, an
agent goes away and does the work in its own.

| Agent | Invoke | What it does |
|---|---|---|
| [**researcher**](agents/researcher.md) | `@agent-researcher` | Researches a topic across the local knowledge base and Confluence, corroborates what it finds, and returns a cited brief with confidence ratings and named gaps — then offers to capture the brief back into the knowledge base |
| [**report-writer**](agents/report-writer.md) | `@agent-report-writer` | Turns research or notes into the written content of a report or official brief, following the structure of an exemplar in `C:\Eva\reference\exemplars`, then runs `/unslop` and `/polish` over it, and offers to capture the result |

They chain: `researcher` produces the cited brief, `report-writer` writes it up.
Each is one Markdown file, so installing is a file copy:

```powershell
$dest = "$env:USERPROFILE\.claude\agents"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force .\agents\researcher.md $dest
```

Or into `.claude\agents\` inside a project to scope it there. Both are written
for this suite rather than as generic agents, and assume it is installed:
`researcher` searches `knowledge-base` and `confluence` on every question,
`report-writer` finishes through `unslop` and `polish`. `report-writer`
produces content and stops there — hand its Markdown to `word` when it needs to
be a `.docx`, so a wording change doesn't mean rebuilding the document. See
[`agents/README.md`](agents/README.md) for the details, including how
`report-writer` picks an exemplar.

## Reference material

[`eva/reference/`](eva/reference) holds what Claude *reads* for style rather
than for facts, which is why it sits outside the indexed corpus:

| Folder | Holds | Wired up by |
|---|---|---|
| [**`reference/exemplars`**](eva/reference/exemplars) | Finished, good documents (`.md`, `.docx`, `.pptx`, `.pdf`) that show the house style to write in | nothing — point at one in the prompt |
| [**`reference/templates`**](eva/reference/templates) | Blank `.docx`/`.pptx` templates a new document or deck is created from | `word` and `powerpoint` → `--templates-dir` (read-only) |

An exemplar is read for guidance and never becomes the output; a template *is*
the output's first draft. Neither is indexed: add a board paper to the RAG
corpus and its phrasing comes back with the same authority as a policy. Each
folder's README covers naming, conventions and how to ask. Document files there
are **not committed** — see [`eva/.gitignore`](eva/.gitignore).

`report-writer` reads `C:\Eva\reference\exemplars` for the shape of the
document it is writing, so the index table in that folder's README is what lets
it pick the right one without opening every file.

## Versioning

Versions follow semver and are bumped on **every** change (see `CLAUDE.md`):

- **MAJOR** — breaking change to configuration (flag/env-var renames) or to a
  tool's name/arguments/output shape
- **MINOR** — new tools, new flags, new behaviour (backwards compatible)
- **PATCH** — bug fixes, documentation-only or internal changes

A version appears in five places that must stay in sync: the server's
`__version__`, its docstring title, its `plugin.json`, its own README header
and the plugin table above (the marketplace manifest mirrors them too).
Standalone skills under `skills/` and agents under `agents/` are unversioned —
they are prose, not an interface anything depends on.
