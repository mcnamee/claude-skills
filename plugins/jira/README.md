# Jira (read-only)

Query Jira issues, sprints and projects. **Strictly read-only** — every request
is an HTTP GET, and there is no code path that creates, edits, transitions,
comments on, or deletes anything.

| | |
|---|---|
| **Server** | `jira.py` v2.0.0 |
| **pip install** | _none_ — standard library only (HTTP via stdlib `urllib`) |
| **Platform** | any |
| **Writes to disk** | no |

> Targets Jira **Data Center / Server** (plain-text descriptions via the v2
> API). Jira Cloud's v3 API returns rich-text documents and is not supported.

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install jira@mcnamee-claude-skills
```

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| Jira base URL | **yes** | `JIRA_BASE_URL` | Base URL including any context path, no trailing slash |
| Project allowlist | no | `JIRA_PROJECTS` | Comma-separated project keys, e.g. `ABC,DEF` |

The Python interpreter is **not** prompted for: it comes from the shared
`EVA_PYTHON` environment variable (see [Configuration](#configuration)).

**Your token is not stored in the plugin.** Set `JIRA_TOKEN` as a Windows user
environment variable before starting Claude Code — the plugin reads it from the
ambient environment. Credentials are deliberately env-var only: there is no flag
that could put a token in a command line, where other local users would see it
in a process listing.

```powershell
setx EVA_PYTHON "C:\Python311\python.exe"
setx JIRA_TOKEN "your-personal-access-token"
```

`setx` does not affect processes that are already running, so quit VS Code
completely (a window reload is not enough) and reopen it. Check it took in a
**new** window with `$env:JIRA_TOKEN`.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `C:\Eva\reference\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `C:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\reference\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

This is the one plugin that touches **no local folder at all**, so of the four
it uses only `EVA_PYTHON` - there is nothing here that has to exist on disk.

### This server's own settings

| Env var | Purpose |
|---|---|
| `JIRA_BASE_URL` | Base URL incl. any context path, no trailing slash. **Required** |
| `JIRA_TOKEN` | Personal Access Token, sent as Bearer (preferred; Jira DC 8.14+) |
| `JIRA_USER` | Username for basic auth (fallback if no token) |
| `JIRA_PASSWORD` | Password for basic auth |
| `JIRA_PROJECTS` | Optional comma-separated **project-key allowlist** (e.g. `"ABC,DEF"`). When set, every tool is confined to those projects: searches are scoped with an AND clause, issue keys outside the list are refused, and other projects are hidden from `jira_list_projects` |
| `JIRA_CA_CERT` | Path to a PEM CA bundle for an internal CA |
| `JIRA_VERIFY_SSL=false` | Disable TLS certificate verification |
| `JIRA_TIMEOUT` | Request timeout in seconds (default 30) |
| `JIRA_MAX_BODY` | Truncate issue descriptions to N chars, 0 = unlimited (default) |

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Connect to Jira, print who you are authenticated as + visible project count to stderr, then exit (no server) |
| `--version` | Print version and exit |

## File access

None. HTTP GET to Jira only; the optional `JIRA_CA_CERT` bundle is read once at
startup.

There is no knowledge folder here: nothing you read from
Jira is written to disk or indexed. To keep something from a ticket, ask for it
to be saved and Claude writes a note with the `knowledge-base` plugin's
`kb_capture` — which only ever runs when you ask.

## Usage examples

1. "What's assigned to me right now, highest priority first?" → `jira_my_issues`
2. "Find any tickets mentioning the login timeout bug — has anyone reported this before?" → `jira_search`
3. "Show me ABC-123 in full, including the comments and who changed its status." → `jira_get_issue` with `include_changelog=true`
4. "Everything resolved in project ABC in the last week, for the release notes." → `jira_search_jql` with `project = ABC AND resolved >= -7d`
5. "How healthy is project ABC — what's open, in progress, unassigned?" → `jira_project_status`
6. "Which projects can I see in Jira?" → `jira_list_projects`
7. "Draft a status report from my open tickets as a Word doc with tracked changes." → `jira_my_issues` + the `word` plugin's editing tools

## Troubleshooting

```powershell
& $env:EVA_PYTHON jira.py --check
```

connects, authenticates and reports who you are plus how many projects you can
see. Behind an internal CA, point `JIRA_CA_CERT` at the PEM bundle rather than
reaching for `JIRA_VERIFY_SSL=false`.
