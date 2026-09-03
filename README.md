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
the document library, and the reference material: **exemplars**
(finished documents showing what good looks like) and **templates** (the blank
`.docx`/`.pptx` new documents are built from). It also carries
[`eva/CLAUDE.md`](eva/CLAUDE.md), the standing instructions that make the
assistant **Eva** — an executive virtual assistant who writes in Australian
English, without em dashes or the other AI tells, and never fills a gap in the
source material with something invented.

## Why this exists

- **Built for Windows endpoints in an Enterprise environment.** No internet calls, no telemetry,
  no registry lookups. Copy the folder across, add it as a local marketplace,
  install.
- **One file per server.** Every server is a single `.py` — nothing to build, no
  package tree to transfer. Three of the eight are **standard library only**, and
  the standalone skills have no code at all.
- **Four settings, once, for the whole suite.** `EVA_PYTHON`,
  `EVA_DOCUMENTS_DIR`, `EVA_TEMPLATES_DIR` and `EVA_KNOWLEDGE_DIR` are the
  configuration. Each server works in its own sub-folder of those roots, named
  after the plugin, so there is nothing per-plugin to fill in, no folder prompts
  at install, and no folder flags to keep in step across a `.mcp.json`, a
  `claude mcp add` line and eight plugin dialogs. `/plugin install` asks only
  for what is genuinely one server's own — a Confluence URL, an embeddings
  endpoint. The matching skill comes with the server.
- **Confined by default.** Every server that touches the filesystem is locked to
  its sub-folders, and refuses to start unconfined rather than falling back to
  "anywhere". `word` and `powerpoint` are the only ones that can change a file
  you already have; the rest either read, or write new Markdown into the
  knowledge tree.
- **Secrets never hit the command line.** Tokens and API keys are environment
  variables only — argv is visible to other local users in process listings.
- **They compose.** `word` and `powerpoint` mirror what they open into one
  Markdown folder, `confluence` and `outlook` save the pages and emails you ask
  them to keep into the same folder, and `pdf-to-md` fills it from PDFs; the
  `knowledge-base` server indexes it and answers questions over the lot.
- **And it grows.** `word` mirrors documents it *writes*, not just ones it
  reads, and `knowledge-base` takes a `kb_capture` call — so an analysis or a
  research brief that would otherwise vanish with the chat goes back into the
  index and answers the same question next time.

## The plugins

| Plugin | Version | What it does | pip install |
|---|---|---|---|
| [**word**](plugins/word) | 7.0.0 | Read, edit and create `.docx` — real Word tracked changes, native styles, filling out templates | `python-docx` |
| [**powerpoint**](plugins/powerpoint) | 3.0.0 | Build `.pptx` decks that inherit your own template's layouts and theme, and audit them against the 10/20/30 rule | `python-pptx` |
| [**excel**](plugins/excel) | 5.0.0 | Read and analyse workbooks; parses `.xlsx` directly, so Excel isn't needed | _none_ |
| [**outlook**](plugins/outlook) | 6.0.0 | Read local Outlook mail and calendar via COM, with a content blacklist; saves an email to the knowledge base when you ask | `pywin32` |
| [**confluence**](plugins/confluence) | 4.0.0 | Search and read Confluence pages, across one or two instances; saves a page to the knowledge base when you ask | _none_ |
| [**jira**](plugins/jira) | 2.0.0 | Query issues, sprints and projects (Jira Data Center v2 API) | _none_ |
| [**knowledge-base**](plugins/knowledge-base) | 4.0.0 | True RAG over your own Markdown: local ChromaDB index + your embeddings API, and capture notes back into it | `chromadb` |
| [**pdf-to-md**](plugins/pdf-to-md) | 6.0.0 | Convert PDFs to Markdown with tables preserved | `pymupdf pymupdf4llm` |

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

That one step creates every folder the servers need, correctly related to each
other — mirrors landing inside the indexed corpus, one folder per file type with
nothing to file first. **Every folder must exist**: a plugin whose folder is
missing either refuses to start (its documents folder) or runs with that
feature off. See [Folder layout](#folder-layout) for what goes where, and
`eva/README.md` for each folder's own README.

The copy brings [`eva/CLAUDE.md`](eva/CLAUDE.md) with it, which is what turns
Claude into Eva when you run Claude Code in `C:\Eva`. Open it and fill in the
**About me** block — your name, role, who you write to, how you sign off. See
[The assistant's instructions](#the-assistants-instructions).

**2. Install the pip dependencies** into the interpreter you are going to name
in step 3:

```powershell
& "C:\path\to\python.exe" -m pip install python-docx pymupdf pymupdf4llm pywin32 chromadb
```

(Drop `pywin32` if you're not on Windows / not using `outlook`. Install only
what the plugins you want need — see the table above.) `word.py`'s docstring
walks through sideloading the wheels.

**3. Set the four suite-wide environment variables.** This is the whole
configuration story: every plugin reads these, and each works in its own
sub-folder of the three roots, named after the plugin.

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\path\to\python.exe",       "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",            "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\reference\templates",  "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",            "User")
```

| Variable | What it points at | Default if unset |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` from step 2 — every server runs under it. A mismatch here is the most common cause of "dependency missing" | *(none — set it)* |
| `EVA_DOCUMENTS_DIR` | The document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | The template library | `C:\Eva\reference\templates` |
| `EVA_KNOWLEDGE_DIR` | The RAG corpus — the one indexed root | `C:\Eva\knowledge` |

**4. Add this repo as a marketplace**, then install whichever plugins you want.
These are slash commands, typed inside Claude Code — not shell commands:

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install word@mcnamee-claude-skills
/plugin install excel@mcnamee-claude-skills
```

There are **no folder prompts**: `word` asks only for an optional
tracked-change author, `excel`, `powerpoint` and `pdf-to-md` ask for nothing at
all, and the rest ask only for what is genuinely theirs (a Confluence or Jira
URL, an embeddings endpoint). The plugins are independent, so a machine without
`pywin32` simply doesn't install `outlook`.

`excel` is the simplest to start with: standard library only, no prompts.

**5. Set your secrets** the same way, before starting Claude Code — they're read
from the ambient environment, never stored in the plugin. Only needed for the
plugins you actually install:

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
& $env:EVA_PYTHON C:\path\to\claude-skills\plugins\word\word.py --check
```

It prints the folders it resolved and says whether each came from your
environment variables or the built-in default, which is usually enough to spot
a typo without reading anything else.

### Manual install, without plugins

If you'd rather configure a server directly — or want one configured differently
from what its plugin prompts for — register it with `claude mcp add --scope user`
(available in every folder), or copy [`.mcp.json.example`](.mcp.json.example) to
`.mcp.json` in the folder you open Claude Code in (config travels with the
files). Keep `PYTHONUTF8=1`: without it, Windows' legacy codepage can corrupt the
stdio JSON stream on non-ASCII content. Pass secrets with `-e` / the `env` block,
never as flags.

```powershell
claude mcp add excel --scope user -e PYTHONUTF8=1 -e EXCEL_DOCS_DIR=C:\path\to\your\workbooks -- $env:EVA_PYTHON C:\path\to\claude-skills\plugins\excel\excel.py
```

(No `&` needed there — `claude` is unquoted, so PowerShell treats it as a
command already.)

## Configuration conventions

Every server follows the same pattern, so once you know one you know them all.
The per-plugin READMEs list each server's actual settings.

1. **Configuration is environment variables only.** No server takes a folder
   flag, a URL flag or a tuning flag; the only command-line flags are actions —
   `--check`, `--version`, and `knowledge-base`'s `--reindex` / `--search` /
   `--ask` / `--debug`. One way to set a thing means two settings can never
   disagree about it, and nothing has to be re-passed in a `.mcp.json`, a
   `claude mcp add` line and a plugin prompt at once.
2. **Four variables configure the whole suite:** `EVA_PYTHON`,
   `EVA_DOCUMENTS_DIR`, `EVA_TEMPLATES_DIR`, `EVA_KNOWLEDGE_DIR`. Each server
   works in its own **sub-folder** of those roots, named after the plugin —
   `word` reads `documents\word`, writes `knowledge\word`, and takes its blanks
   from `reference\templates\word`. `knowledge-base` is the one exception, and
   necessarily so: it indexes the *whole* knowledge root.
3. **A per-server variable overrides one folder,** for an endpoint whose layout
   differs: `<PREFIX>_DOCS_DIR`, `<PREFIX>_TEMPLATES_DIR`, `<PREFIX>_KB_DIR`.
   It beats the shared root. Prefixes: `CONFLUENCE_`, `JIRA_`, `KB_`, `EXCEL_`,
   `OUTLOOK_`, `MSWORD_` (the `word` server keeps this older prefix), `PDF2MD_`,
   `POWERPOINT_`.
4. **Secrets are env-var only**, and always were — no flag has ever existed
   that could put a token in a command line, where other local users can read
   it out of a process listing.
5. **A blank value means "not configured"**, so the shared root still applies —
   that is what an MCP client substitutes for a prompt left empty. To switch an
   optional folder off, set it to `off` (`none`, `no`, `false` and `disabled`
   also work). A folder you named yourself that does not exist is a fatal
   error, because it is almost always a typo; a built-in default that does not
   exist yet is a warning, and the feature it enables simply stays off.
6. **Every folder must exist.** Copying [`eva/`](eva) to `C:\Eva` creates all of
   them; each server's `--check` reports which are missing and where the path
   came from.

## Folder layout

One working folder, `C:\Eva`, holds everything the servers read, write and
index. The repo carries it as a scaffold — [`eva/`](eva) is the same tree with a
README in every folder and no content, so `Copy-Item -Recurse ...\eva C:\Eva`
lays it out in one step and every default below is already correct.

```
C:\Eva\
├─ CLAUDE.md         who Eva is, and how she writes
├─ knowledge\        the RAG corpus - Markdown only, the ONE indexed root
│  ├─ notes\           Markdown you write by hand
│  ├─ captures\        notes kb_capture writes back
│  ├─ confluence\      pages you asked the confluence plugin to save
│  ├─ email\           emails you asked the outlook plugin to save
│  ├─ word\            documents the word plugin mirrored
│  ├─ powerpoint\      decks the powerpoint plugin mirrored
│  └─ pdf\             PDFs the pdf-to-md plugin converted
├─ index\            the ChromaDB vector store - derived, disposable
├─ documents\        one folder per file type - yours and the assistant's
│  ├─ word\            .docx (searched recursively)
│  ├─ powerpoint\      .pptx (searched recursively)
│  ├─ excel\           .xlsx (top level only - excel does not recurse)
│  └─ pdf\             source PDFs
└─ reference\        style, not facts - deliberately NOT indexed
   ├─ exemplars\       finished documents showing what good looks like
   └─ templates\       blank branded files new documents/decks start from
      ├─ word\           .docx blanks
      └─ powerpoint\     .pptx / .potx deck shells
```

Three environment variables name the three roots; the plugin's own name is the
sub-folder. **Every folder listed here must exist.**

| Plugin | `EVA_DOCUMENTS_DIR\` | `EVA_TEMPLATES_DIR\` | `EVA_KNOWLEDGE_DIR\` |
|---|---|---|---|
| `word` | `word\` | `word\` | `word\` |
| `powerpoint` | `powerpoint\` | `powerpoint\` | `powerpoint\` |
| `excel` | `excel\` | — | — |
| `pdf-to-md` | `pdf\` | — | `pdf\` |
| `outlook` | — | — | `email\` |
| `confluence` | — | — | `confluence\` |
| `knowledge-base` | — | — | the **whole root** it indexes, plus `captures\` |
| `jira` | — | — | — |

`knowledge-base` also keeps its vector store in `C:\Eva\index`
(`KB_INDEX_DIR`) — the one folder in the suite that is not under a shared root,
deliberately outside the corpus so a large binary database does not sit in the
folder you index.

Three rules make it hold together, and all three are worth knowing before you
move a folder:

**One folder per file type.** Every `.docx` lives in `documents\word\`, every
`.pptx` in `documents\powerpoint\`, whether you put it there or the assistant
wrote it. There is no `input\`, no `output\`, and no `inbox\` vs `library\`
split: each plugin can open nothing outside its single folder, so the extras
were either unreachable or one more setting to keep in sync — and each made you
decide where a file belonged before you could ask a question about it. It is
also what lets the plugins have no folder settings of their own: the folder name
*is* the plugin name. Organise inside the folder however you like; `word` and
`powerpoint` search recursively and a bare filename still finds the file.

**There is one indexed root.** `knowledge\` is the only folder the RAG index
reads, so every server that produces Markdown writes inside it. Deriving all of
them from one `EVA_KNOWLEDGE_DIR` is what makes that impossible to get wrong;
point a single mirror anywhere else and it fills up faithfully while `kb_ask`
never sees a word of it.

**Folders are named after what wrote them, not what they are about.** Topic
folders rot — every document belongs to three of them. Provenance maps
one-to-one onto a setting, retrieval is semantic anyway, and cleanup stays
surgical: delete `knowledge\confluence\` and save the pages again, with nothing
you wrote yourself at risk.

[`eva/README.md`](eva/README.md) covers the reasoning, how to put the tree on
another drive, and why nothing in it is committed. Each folder's own README says
what belongs there.

## File access policy

Every server that touches the filesystem is confined to its sub-folders of the
three shared roots, and those folders are **required**:

| Plugin | Local file access |
|---|---|
| `word` | Read/write, confined to the one documents folder (where new documents are created too) plus the knowledge-base folder; the templates folder is read-only. Opening, creating and saving each mirror to the knowledge-base folder |
| `powerpoint` | Read/write, confined to the one presentations folder (where new decks are created too) plus the knowledge-base folder; the templates folder is read-only. Opening, creating and saving each mirror to the knowledge-base folder |
| `excel` | Read-only, confined to the workbook folder (top level only) |
| `knowledge-base` | Reads the documents folder; writes its vector index (`C:\Eva\index`) and captured notes (`C:\Eva\knowledge\captures`, always inside the documents folder); never edits or deletes an existing document; network only to the endpoints you configure |
| `pdf-to-md` | Reads the PDF folder, writes the output folder |
| `confluence` | Writes only the knowledge-base folder, and only for a page you asked to keep (`save_to_kb`); reading a page saves nothing. Set the folder to `off` and the server touches no local file |
| `outlook` | Writes only the knowledge-base folder, and only for an email you asked to keep (`save_to_kb`); reading mail saves nothing, and a blacklisted message is never written at all. Set the folder to `off` and the server touches no local file |
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
| [**brief-writer**](skills/brief-writer) | `/brief-writer` | Drafts a decision or noting brief for a senior executive, following the structure of an exemplar in its own `exemplars/` folder, and finishing with `/polish` |
| [**email-writer**](skills/email-writer) | `/email-writer` | Drafts an email in your voice, classifying what the email is for and matching that intent to your own sent mail in its `exemplars/` folder, then running `/unslop` |
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

The two writers stack on the other two: `brief-writer` finishes through
`/polish`, `email-writer` through `/unslop`, so install the pair each one needs.
Both keep their own `exemplars/` folder inside the skill — your approved briefs,
your sent mail — which is what makes the copy to an endpoint carry the house
structure and your voice with it. Nothing in those folders is committed except
their README.

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

It is one Markdown file, so installing is a file copy:

```powershell
$dest = "$env:USERPROFILE\.claude\agents"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force .\agents\researcher.md $dest
```

Or into `.claude\agents\` inside a project to scope it there. It is written for
this suite rather than as a generic agent and assumes it is installed, searching
`knowledge-base` and `confluence` on every question. What it produces is
research, not a document: hand the brief to [`/brief-writer`](skills/brief-writer)
or [`/email-writer`](skills/email-writer) to write up, which keeps you in the
conversation while the wording is settled. See
[`agents/README.md`](agents/README.md) for the details.

## Reference material

[`eva/reference/`](eva/reference) holds what Claude *reads* for style rather
than for facts, which is why it sits outside the indexed corpus:

| Folder | Holds | Wired up by |
|---|---|---|
| [**`reference/exemplars`**](eva/reference/exemplars) | Finished, good documents (`.md`, `.docx`, `.pptx`, `.pdf`) that show the house style to write in | nothing — point at one in the prompt |
| [**`reference/templates`**](eva/reference/templates) | Blank `.docx`/`.pptx` templates a new document or deck is created from, one sub-folder per plugin (`templates\word`, `templates\powerpoint`) | `EVA_TEMPLATES_DIR` — read-only for both plugins |

An exemplar is read for guidance and never becomes the output; a template *is*
the output's first draft. Neither is indexed: add a board paper to the RAG
corpus and its phrasing comes back with the same authority as a policy. Each
folder's README covers naming, conventions and how to ask. Document files there
are **not committed** — see [`eva/.gitignore`](eva/.gitignore).

Point at one in the prompt and Claude follows its shape, so the index table in
that folder's README is what lets it pick the right file without opening every
one. The [`brief-writer`](skills/brief-writer) and
[`email-writer`](skills/email-writer) skills keep their own exemplars folder
instead, inside the skill, so it travels with the folder copy to an endpoint.

## The assistant's instructions

The plugins give Claude tools and the folders give it somewhere to work.
[`eva/CLAUDE.md`](eva/CLAUDE.md), which the copy in step 1 puts at
`C:\Eva\CLAUDE.md`, gives it a job. Claude Code loads it whenever you run in
that folder, so run there rather than somewhere else.

It defines **Eva**, an executive virtual assistant: correspondence, diary and
meeting prep, minutes and actions, briefs and board papers, decks, research over
the knowledge base, and status summaries out of Jira. Along with the role it
fixes the things you would otherwise correct in every reply:

- **Australian English throughout**, in documents as well as chat — `-ise`
  spellings, `31 August 2026`, `2.30 pm`, financial years as `2025–26`,
  metric, and the Australian Government Style Manual as the authority for the
  finer mechanics.
- **No em dashes and no other AI tells** — the openers, the vocabulary, the
  "it's not X, it's Y" shapes. This is the always-on subset of
  [`unslop`](skills/unslop); the full passes are still `/unslop` then
  `/polish` before anything leaves your desk.
- **Sourced or not stated** — on a network with no internet, an empty search is
  a finding to report rather than a gap to fill from memory, and every fact
  carries where it came from.
- **Boundaries** — Eva drafts, you send. Nothing in the suite can send mail or
  accept a meeting, and the file makes that a rule rather than an accident:
  no committing on your behalf, protective markings carried through and never
  downgraded, personal information kept out of the indexed corpus.

Fill in the **About me** block at the top (name, role, stakeholders, sign-off)
and edit the rest to suit — it is prose, and yours to change. To make Eva the
default in every folder instead of just `C:\Eva`, copy the file to
`%USERPROFILE%\.claude\CLAUDE.md`, remembering it will then shape coding
sessions too.

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
