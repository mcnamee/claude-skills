You're an expert Python developer.

You are an expert Python 3 programmer. You write scripts for an endpoint that is within an enterprise Microsoft Windows environment, with no internet access. You have access to use Python 3 and pip via a proxy.

## Constraints

- **Windows computer** - the network is a windows enterprise system, and the script is going to be run on a Windows endpoint 
- **Single-file scripts only** — it's difficult to get individual files on the endpoint, so please aim for single files scripts. Each server lives at `plugins/<name>/<name>.py` and is the only `.py` in its plugin; a plugin must never reference a file outside its own folder, because Claude Code copies plugins into a cache on install and an outside path would break there.
- **Standard library as priority** — if a module isn't in the Python 3 standard library, ask for confirmation before adding another library. There is an option to use pip, which has a proxy in the network. 
- **No internet calls** — no requests, urllib calls to external hosts, or network-dependent logic
- **Python 3 compatible** — assume a reasonably modern Python 3 (3.8+), but do not rely on features from very recent releases unless explicitly asked
- **Configuration** - where configuration and testing is needed, please include a doc block at the start of the file with this included, so that its a single file transferred and I can copy/paste from the docblock

## Code quality

- Write complete, runnable scripts — never pseudocode or partial stubs
- Include clear inline comments for non-obvious logic
- Handle likely error conditions explicitly (file not found, bad input, permission errors, etc.)
- Use argparse for any script that accepts arguments, with sensible --help text
- Prefer explicit over clever — readability matters more than brevity
- Ensure documentation (eg. args, usage, requirements, testing) is updated in both the docblock at the top of the server and its own `plugins/<name>/README.md`. The root `README.md` is deliberately kept short — benefits, install, links to each plugin, shared conventions — so per-server detail belongs in the plugin README, not there.

## Before writing code

- If the requirement is ambiguous, ask a clarifying question before proceeding — a wrong assumption costs a full transfer cycle to discover
- State any assumptions you are making at the top of your response
- If a task genuinely cannot be done cleanly with the standard library alone, say so upfront rather than producing a fragile workaround

## Confidence standard

This script will be transferred to an airgapped network, which is time-consuming. Only provide code you are confident is correct and complete. If you are uncertain about any part, flag it explicitly rather than guessing. A caveat is far cheaper than a failed transfer.

## Versioning (MANDATORY on every change)

Every MCP server in this repo carries a semantic version:

- `__version__ = "X.Y.Z"` sits immediately after the module docstring, and the docstring's title line shows the same version, e.g. `excel.py (v2.0.0)`.
- `SERVER_VERSION` / `SERVER_INFO` (whatever the file reports to the MCP client in `serverInfo`) must reference `__version__`, never a duplicate literal.
- Each server exposes a `--version` flag printing `<server-name> <version>`. It must work even when the server's pip dependencies are missing (servers that import heavy/platform deps at module level answer `--version` before that import).

**Whenever you change a server file, bump its version in the same change** — in `__version__`, the docstring title, the `version` field of its `plugins/<name>/.claude-plugin/plugin.json`, the header table of its `plugins/<name>/README.md`, and the plugin table in the root `README.md` (all five must stay in sync; the marketplace at `.claude-plugin/marketplace.json` mirrors the plugin versions too):

- **MAJOR** — anything that breaks an existing integration: renaming/removing a CLI flag, env var or config constant; changing a tool's name, arguments or output shape; changing defaults in a behaviour-altering way.
- **MINOR** — backwards-compatible additions: new tools, new flags/env vars, new behaviour.
- **PATCH** — bug fixes, refactors, comment/docstring-only changes.

## Configuration conventions (all MCP servers)

Keep every server consistent with these rules (documented for users in README.md → "Configuration conventions"):

- **Precedence:** CLI flag > environment variable > constant in the file's CONFIG block. Every non-secret setting should offer at least flag + env var.
- **Naming:** env var = server prefix + upper-snake flag name (`--docs-dir` → `EXCEL_DOCS_DIR`). Prefixes: `CONFLUENCE_`, `JIRA_`, `KB_` (knowledge-base), `EXCEL_`, `OUTLOOK_`, `MSWORD_` (the `word` server keeps this older prefix, as do its `msword_*` tool names — renaming them would break every existing config), `PDF2MD_`. Exception: `--insecure` pairs with `<PREFIX>_VERIFY_SSL=false`.
- **Secrets are env-var ONLY** — never add `--token`/`--password`/`--*-api-key` flags (argv is visible to other local users in process listings).
- **Shared flag vocabulary:** `--docs-dir` (the source-documents folder a server is confined to), `--output-dir` (generated files), `--templates-dir` (blank templates a new file is created from — a READ-ONLY root: readable, but every save into it is refused), `--kb-dir` (Markdown mirror for the RAG knowledge base), `--index-dir` (the vector store), `--base-url`/`--ca-cert`/`--insecure`/`--timeout`/`--max-body` (HTTP servers), `--check` (validate config/connectivity and exit), `--version`. Reuse these names for new servers/settings; do not invent synonyms (no `--folder`, `--input-dir`, `--document-root`).
- **Folder defaults — the `eva/` tree:** every folder setting defaults to its place in `C:\Eva`, the single working folder documented in `eva/README.md`. The repo carries `eva/` as a scaffold: the same tree, a `README.md` in every folder, and no content (`eva/.gitignore` ignores everything else). A new folder setting gets a default inside that tree, and a new folder in the tree gets a `README.md` saying what belongs there — both in the same change. The invariant that matters: **`C:\Eva\knowledge` is the only indexed root**, so anything writing Markdown for the RAG index writes into a sub-folder of it, named after the server that wrote it (`knowledge\confluence`, `knowledge\email`, `knowledge\word`, `knowledge\pdf`, `knowledge\captures`). Never point a mirror outside it. `C:\Eva\reference` is deliberately NOT indexed — exemplars are style, not fact.
- **Folder settings that have a default:** a BLANK value means "not configured" and must fall back to the default, because that is what an MCP client substitutes for a prompt the user left empty — so read env vars as `env("X") or CONSTANT`, never `os.environ.get("X", CONSTANT)`, which lets an empty string win. To let a user switch an optional folder OFF, accept the `DISABLE_KEYWORDS` (`off`, `none`, `no`, `false`, `disabled`); editing the `.py` is not an option for a plugin install, since the file lives in a cache that an update replaces. And distinguish a folder the USER configured from one that came from the default: a missing user-configured folder is fatal (it is almost always a typo), while a missing default just means this endpoint has not created that part of the tree yet — warn, disable the feature it enables, and carry on. A required sandbox (`word`/`excel`/`knowledge-base` → `--docs-dir`) stays fatal either way, but says how to fix it.
- **Reference material:** `eva/reference/exemplars` (finished documents showing what good output looks like) and `eva/reference/templates` (blank `.docx`/`.pptx` files new documents are created from) are reference material, not code — no `plugin.json`, no version. Each carries a `README.md`, not a `CLAUDE.md`: they are consumed on the user's endpoint, not by an agent working in this repo. The document files themselves are gitignored because they are usually the user's real corporate documents and this repo is public; keep it that way unless asked. A server that needs templates points at `C:\Eva\reference\templates` with `--templates-dir` (the `word` server does).
- **Skills:** each server has a matching Claude skill at `plugins/<server-name>/skills/<server-name>/SKILL.md`. When a server's tools or workflow change, update its skill in the same change.
- **Standalone skills:** skills with no MCP server behind them live at `skills/<name>/SKILL.md`, with a `README.md` beside them, and are NOT plugins — no `plugin.json`, no marketplace entry, no version (they are prose, not an interface). The layout mirrors `~/.claude/skills/` so installing is a folder copy, and they are invoked by bare name (`/unslop`). Add new ones there, and add a row to `skills/README.md`. The versioning rules above apply only to `plugins/`.
- **Agents:** subagents live at `agents/<name>.md` — a single Markdown file each, mirroring `~/.claude/agents/` so installing is a file copy. Like standalone skills they are NOT plugins and are unversioned. `name` in the frontmatter must match the filename; the `description` is the only part Claude reads when deciding whether to delegate. Don't set `tools:` — an agent that names MCP tools explicitly loses access silently when a server is registered under a different name, so let it inherit. Add new ones there, and add a row to both `agents/README.md` and the Agents table in the root `README.md`. Agents are written for this suite and may assume its servers and skills are installed and configured — don't pad them with fallback paths for a missing server. That licence covers tooling only: never let an agent fill a gap in the *source material* with something invented; an empty search or an unanswerable section is a finding to report, not a hole to paper over. Keep each agent to one job (`report-writer` writes content and stops; formatting a `.docx` is `word`'s job) so the steps can be redone independently.
- **Claude Code only:** the suite targets Claude Code (plugins, with `claude mcp add` / `.mcp.json` as the manual fallback). Do not add configuration examples for other MCP clients.
