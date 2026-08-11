# mcp-servers

Seven MCP servers that give an AI agent hands on the things enterprise work
actually lives in — Word documents, Excel workbooks, Outlook mail, Confluence,
Jira, PDFs — plus a local RAG knowledge base to tie them together.

Each one is packaged as a **Claude Code plugin**, so this repo doubles as a
plugin marketplace that works **entirely offline**.

## Why this exists

- **Built for an airgapped Windows endpoint.** No internet calls, no telemetry,
  no registry lookups. Copy the folder across, add it as a local marketplace,
  install.
- **One file per server.** Every server is a single `.py` — nothing to build, no
  package tree to transfer. Four of the seven are **standard library only**.
- **Install with prompts, not JSON.** `/plugin install` asks for the folders and
  the Python interpreter instead of you hand-editing absolute paths in seven
  places. The matching skill comes with the server.
- **Confined by default.** Every server that touches the filesystem is locked to
  the folders you name, and refuses to start unconfined rather than falling back
  to "anywhere". Five of the seven are read-only.
- **Secrets never hit the command line.** Tokens and API keys are environment
  variables only — argv is visible to other local users in process listings.
- **They compose.** Word, Outlook and Confluence can each mirror what they read
  into one Markdown folder; `pdf-to-md` fills the same folder from PDFs; the
  `knowledge-base` server indexes it and answers questions over the lot.

## The plugins

| Plugin | Version | What it does | pip install |
|---|---|---|---|
| [**word**](plugins/word) | 4.0.0 | Read, edit and create `.docx` — real Word tracked changes, native styles, filling out templates | `python-docx` |
| [**excel**](plugins/excel) | 3.0.0 | Read and analyse workbooks; parses `.xlsx` directly, so Excel isn't needed | _none_ |
| [**outlook**](plugins/outlook) | 3.0.0 | Read local Outlook mail and calendar via COM, with a content blacklist | `pywin32` |
| [**confluence**](plugins/confluence) | 1.3.1 | Search and read Confluence pages | _none_ |
| [**jira**](plugins/jira) | 1.1.1 | Query issues, sprints and projects (Jira Data Center v2 API) | _none_ |
| [**knowledge-base**](plugins/knowledge-base) | 2.0.1 | True RAG over your own Markdown: local ChromaDB index + your embeddings API | `chromadb` |
| [**pdf-to-md**](plugins/pdf-to-md) | 4.0.1 | Convert PDFs to Markdown with tables preserved | `pymupdf pymupdf4llm` |

Each plugin's README covers its settings, tools, file access and example
prompts. Every server also carries a semantic version in `__version__`, printed
by `--version` and reported to the MCP client in `serverInfo`.

## Install

**1. Install the pip dependencies** into the *same* interpreter you'll point the
plugins at:

```
"C:\path\to\python.exe" -m pip install python-docx pymupdf pymupdf4llm pywin32 chromadb
```

(Drop `pywin32` if you're not on Windows / not using `outlook`. Install only
what the plugins you want need — see the table above.) `word.py`'s docstring
walks through sideloading the wheels on an airgapped machine.

**2. Add this repo as a marketplace**, then install whichever plugins you want:

```
/plugin marketplace add C:\path\to\mcp-servers
/plugin install word@mcnamee-mcp-servers
/plugin install excel@mcnamee-mcp-servers
```

Claude Code prompts for that server's settings — documents folder, output
folder, and the **Python interpreter** (give the absolute path to the
`python.exe` from step 1; a mismatch here is the most common cause of
"dependency missing"). The plugins are independent, so a machine without
`pywin32` simply doesn't install `outlook`.

`excel` is the simplest to start with: standard library only, one prompt.

**3. Set your secrets** as Windows user environment variables before starting
Claude Code — they're read from the ambient environment, never stored in the
plugin:

| Plugin | Environment variable |
|---|---|
| `confluence` | `CONFLUENCE_TOKEN` |
| `jira` | `JIRA_TOKEN` |
| `knowledge-base` | `KB_EMBED_API_KEY` |

`setx` doesn't affect already-running processes, so quit VS Code completely (not
just a window reload) and reopen it after setting one.

Useful commands: `/plugin` to browse and manage, `/mcp` to confirm a server
connected, `claude mcp list` to spot an unresolved environment variable, and
`/plugin marketplace update mcnamee-mcp-servers` after you transfer a new
version across.

> **Skills are namespaced** by their plugin, so it's `/word:word` rather
> than `/word`.
>
> **Plugins are cached on install** (copied under `~/.claude/plugins/`), which is
> why each plugin contains its own server file rather than sharing one — a path
> reaching outside the plugin would break once cached.

### Before you wire anything in

Run the server's `--check` first. It validates config (and connectivity, for the
HTTP servers) without starting the server, and is far easier to read than an MCP
connection failure:

```
"C:\path\to\python.exe" C:\path\to\mcp-servers\plugins\word\word.py --check
```

### Manual install, without plugins

If you'd rather configure a server directly — or want one configured differently
from what its plugin prompts for — register it with `claude mcp add --scope user`
(available in every folder), or copy [`.mcp.json.example`](.mcp.json.example) to
`.mcp.json` in the folder you open Claude Code in (config travels with the
files). Keep `PYTHONUTF8=1`: without it, Windows' legacy codepage can corrupt the
stdio JSON stream on non-ASCII content. Pass secrets with `-e` / the `env` block,
never as flags.

```
claude mcp add excel --scope user -e PYTHONUTF8=1 -- C:\path\to\python.exe C:\path\to\mcp-servers\plugins\excel\excel.py --docs-dir C:\path\to\your\workbooks
```

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
   confined to), `--output-dir` (generated files), `--kb-dir` (Markdown mirror
   for the RAG knowledge base), `--base-url`/`--ca-cert`/`--insecure`/
   `--timeout`/`--max-body` (the HTTP servers), `--check`, `--version`.

## File access policy

Every server that touches the filesystem is confined to the folder(s) named in
its configuration, and that configuration is **required**:

| Plugin | Local file access |
|---|---|
| `word` | Read/write, confined to the documents folder (plus the output and knowledge-base folders, if set) |
| `excel` | Read-only, confined to the workbook folder |
| `knowledge-base` | Reads the documents folder; writes only its vector index; network only to the endpoints you configure |
| `pdf-to-md` | Reads the PDF folder, writes the output folder |
| `confluence` | None unless a knowledge-base folder is set; then writes only there |
| `outlook` | None unless a knowledge-base folder is set; then writes only there |
| `jira` | None — HTTP GET to Jira only |

Paths are resolved (symlinks included) before the containment check, so a
symlink dropped inside a configured folder cannot reach files outside it.

## Skills

Every server ships with a Claude skill at
`plugins/<name>/skills/<name>/SKILL.md` — it teaches an agent the server's
tools, the right call order, and the sharp edges (read-only limits, sandboxes,
the tracked-changes workflow, when to reindex). **Installing the plugin installs
its skill**, namespaced as `/<plugin>:<skill>`.

To install a skill on its own, without its server:

```
xcopy /E /I plugins\word\skills\word %USERPROFILE%\.claude\skills\word
```

Skills describe *how to use* a server — they aren't a way to run one. These are
MCP servers, launched as long-running stdio subprocesses, not scripts a skill
shells out to. That matters most for `word`, which is session-based
(`msword_open` returns a `session_id` and holds the document in memory until
`msword_save`).

## Versioning

Versions follow semver and are bumped on **every** change (see `CLAUDE.md`):

- **MAJOR** — breaking change to configuration (flag/env-var renames) or to a
  tool's name/arguments/output shape
- **MINOR** — new tools, new flags, new behaviour (backwards compatible)
- **PATCH** — bug fixes, documentation-only or internal changes

A version appears in four places that must stay in sync: the server's
`__version__`, its docstring title, its `plugin.json`, and the table above (the
marketplace manifest mirrors them too).
