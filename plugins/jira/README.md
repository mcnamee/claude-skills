# Jira

Query Jira issues, sprints and projects — and, once you switch writing on,
create issues, edit them, comment on them and move them through their workflow.
**Read-only by default**: until `JIRA_ALLOW_WRITE=true`, every request is an HTTP
GET and the write tools are not offered at all. Nothing can ever delete an
issue.

| | |
|---|---|
| **Server** | `jira.py` v3.3.0 |
| **pip install** | _none_ — standard library only (HTTP via stdlib `urllib`) |
| **Platform** | any |
| **Writes to disk** | no |
| **Writes to Jira** | only with `JIRA_ALLOW_WRITE=true` — off by default |

> Targets Jira **Data Center / Server** (plain-text descriptions via the v2
> API). Jira Cloud's v3 API returns rich-text documents and is not supported.

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install jira@mcnamee-claude-skills
```

The plugin prompts for nothing at install time. Set these in the `env` block of
Claude Code's settings.json ([Where to set a variable](../../README.md#where-to-set-a-variable))
or as Windows user environment variables, then fully restart Claude Code. A
blank or missing value means "not set".

| Variable | Required | Purpose |
|---|---|---|
| `JIRA_BASE_URL` | **yes** | Base URL including any context path, no trailing slash |
| `JIRA_TOKEN` | **yes** (or `JIRA_USER` + `JIRA_PASSWORD`) | Personal Access Token, sent as Bearer |
| `JIRA_PROJECTS` | no | Comma-separated project keys, e.g. `ABC,DEF` |
| `JIRA_ALLOW_WRITE` | no | `true` to let it create, edit, comment on and transition issues. Blank = read-only |

```json
{
  "env": {
    "JIRA_BASE_URL": "https://jira.internal.example.com",
    "JIRA_TOKEN": "your-personal-access-token",
    "JIRA_PROJECTS": "ABC,DEF"
  }
}
```

The Python interpreter comes from the shared `EVA_PYTHON` environment variable
(see [Configuration](#configuration)).

**Your token is never stored in the plugin.** Credentials are deliberately
env-var only: there is no flag that could put a token in a command line, where
other local users would see it in a process listing. As Windows user
environment variables instead of settings.json:

```powershell
setx EVA_PYTHON "C:\Python311\python.exe"
setx JIRA_TOKEN "your-personal-access-token"
```

`setx` does not affect processes that are already running, so quit VS Code
completely (a window reload is not enough) and reopen it. Check it took in a
**new** window with `$env:JIRA_TOKEN`.

## Tools

| Tool | Does |
|---|---|
| `jira_search` | Free-text search (safely quoted into JQL) |
| `jira_search_jql` | Advanced search with raw JQL |
| `jira_get_issue` | One issue in full: fields, description, comments, optionally the change history, plus any custom fields listed in `JIRA_EXTRA_FIELDS` |
| `jira_my_issues` | Issues assigned to you |
| `jira_project_status` | Health summary of one project |
| `jira_list_projects` | The project keys you can see |
| `jira_list_transitions` | An issue's status and the transitions available from it (and any field each needs) |
| `jira_list_versions` | A project's releases (fix versions): name, id, released/archived, dates. `query` filters by part of the name, ignoring case, spaces, hyphens, underscores and dots; archived ones are hidden unless `include_archived=true` |
| `jira_create_issue` | *(writing on)* New issue: project, type, summary, plus description, priority, assignee, labels, components, fix versions, due date, `parent` for a sub-task, and any custom field via `fields` |
| `jira_update_issue` | *(writing on)* Edit any of those fields; `labels` replaces, `add_labels` / `remove_labels` adjust; `fix_versions` replaces, `add_fix_versions` / `remove_fix_versions` adjust; `assignee: "none"` unassigns; optional `comment` in the same call |
| `jira_add_comment` | *(writing on)* Comment on an issue |
| `jira_transition_issue` | *(writing on)* Move an issue by transition name, id, **or target status** ("move it to Done"), with an optional `resolution` and `comment`. No match lists what is available |

## Writing issues

**Off unless `JIRA_ALLOW_WRITE=true`.** Then:

- **Text is Jira wiki markup**, the format Jira Data Center stores — `h2.
  Heading`, `*bold*`, `_italic_`, `{{code}}`, `* bullet`, `# numbered`,
  `[label|https://url]`, `||head||` / `|cell|` tables, `{code}...{code}`,
  `[~username]` mentions. Not Markdown.
- **Assignees** can be given as a username, an email or a display name
  (`me` for you). A name that matches more than one person is refused with the
  candidates, never guessed.
- **Custom fields** go in `fields` by id, e.g.
  `{"customfield_10010": "value"}` or `{"customfield_10020": {"value": "Option"}}`.
  When Jira refuses a create, its reply names each field it wanted — the tool
  passes that on field by field.
- **`JIRA_PROJECTS` confines writes** exactly as it confines reads: a project
  or issue key outside the allowlist is refused before anything is sent.
- Jira's own permissions still apply: the account can only do what it could do
  in the browser.
- **Fix versions** must already exist in the issue's own project, named
  exactly as Jira has them; `jira_list_versions` shows the names. An unknown
  name is refused by Jira, never created.
- A comment added alongside a transition is posted separately after the move,
  because Jira refuses a comment inside a transition that has no screen.

## Reading custom fields

`jira_get_issue` shows the standard fields only, unless you list others in
`JIRA_EXTRA_FIELDS`. They then appear in an **Extra fields** block above the
description:

```
Extra fields:
  Story Points: 5
  Team: Blue
  Sprint: Sprint 12
```

- **Each entry is a field id or a name.** `customfield_10010` is used as is;
  anything else (`Story Points`, `environment`) is matched against Jira's field
  list by id, then by name ignoring case. The list is fetched once per server
  start, and only if an entry needs it.
- **A name two fields share shows both**, each labelled with its id, so you can
  swap the name for the id you meant.
- **An entry that matches nothing** is noted in the output rather than failing
  the call. Run `--check` after setting the variable: it lists how each entry
  resolved and warns about any that didn't.
- **Values are shown the way the UI shows them**: an option's value (a
  cascading select as `Parent / Child`), a user's display name, a list joined
  with commas, a sprint's name, `5` rather than `5.0`. Anything else is shown
  as compact JSON, and each value is capped at 500 characters.
- Search results stay one line per issue and do not include these fields. To
  *filter* on a custom field, use JQL in `jira_search_jql`, e.g.
  `cf[10010] = "Blue"` or `"Story Points" > 3`.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no install prompts and no folder command-line flags.

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
| `JIRA_ALLOW_WRITE=true` | Offer the write tools (create, update, comment, transition) — see [Writing issues](#writing-issues). Default off: read-only |
| `JIRA_EXTRA_FIELDS` | Optional comma-separated list of extra fields `jira_get_issue` shows, e.g. `"customfield_10010,Story Points,Team"` — see [Reading custom fields](#reading-custom-fields). Blank = none |

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Connect to Jira, print who you are authenticated as + visible project count and whether writing is on, to stderr, then exit (no server) |
| `--version` | Print version and exit |

## File access

None. HTTP to Jira only (GET only, unless `JIRA_ALLOW_WRITE=true`); the optional
`JIRA_CA_CERT` bundle is read once at startup.

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
8. "Raise a bug in ABC: the export button times out on large reports, assign it to Jane Smith." → `jira_create_issue` *(writing on)*
9. "Bump ABC-123 to High and add the label `regression`." → `jira_update_issue` with `priority` and `add_labels` *(writing on)*
10. "Comment on ABC-123 that the fix is in the 2.4 build." → `jira_add_comment` *(writing on)*
11. "Move ABC-123 to In Progress." / "Resolve ABC-123 as Done." → `jira_transition_issue` (with `resolution` for the second) *(writing on)*
12. "What can ABC-123 move to from here?" → `jira_list_transitions`
13. "Which releases does ABC have coming up?" → `jira_list_versions`
14. "For each 'delivers' link on ABC-10, add ABC-10 to the release with the closest name - show me the pairs first." → `jira_get_issue` + `jira_list_versions`, then `jira_update_issue` with `add_fix_versions` *(writing on)*

## Troubleshooting

> **If a server fails with `Executable not found in $PATH: "${EVA_PYTHON}"`**,
> the variable is not set in the environment Claude Code was launched from. Set
> it (see above), then quit Claude Code completely and reopen — `setx` and
> `[Environment]::SetEnvironmentVariable` do not reach a process that is already
> running.

```powershell
& $env:EVA_PYTHON jira.py --check
```

connects, authenticates and reports who you are plus how many projects you can
see. Behind an internal CA, point `JIRA_CA_CERT` at the PEM bundle rather than
reaching for `JIRA_VERIFY_SSL=false`.
