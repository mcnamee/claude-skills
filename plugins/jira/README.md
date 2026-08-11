# Jira (read-only)

Query Jira issues, sprints and projects. **Strictly read-only** — every request
is an HTTP GET, and there is no code path that creates, edits, transitions,
comments on, or deletes anything.

| | |
|---|---|
| **Server** | `jira.py` v1.1.2 |
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
| Python interpreter | **yes** | — | Absolute path to the `python.exe` to launch the server with |

**Your token is not stored in the plugin.** Set `JIRA_TOKEN` as a Windows user
environment variable before starting Claude Code — the plugin reads it from the
ambient environment. Credentials are deliberately env-var only: there are no
`--token`/`--user`/`--password` flags, because command-line arguments are
visible to other local users in process listings.

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| Env var | CLI flag | Purpose |
|---|---|---|
| `JIRA_BASE_URL` | `--base-url` | Base URL incl. any context path, no trailing slash |
| `JIRA_TOKEN` | _(env only)_ | Personal Access Token, sent as Bearer (preferred; Jira DC 8.14+) |
| `JIRA_USER` | _(env only)_ | Username for basic auth (fallback if no token) |
| `JIRA_PASSWORD` | _(env only)_ | Password for basic auth |
| `JIRA_PROJECTS` | `--projects` | Optional comma-separated **project-key allowlist** (e.g. `"ABC,DEF"`). When set, every tool is confined to those projects: searches are scoped with an AND clause, issue keys outside the list are refused, and other projects are hidden from `jira_list_projects` |
| `JIRA_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `JIRA_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification |
| `JIRA_TIMEOUT` | `--timeout` | Request timeout in seconds (default 30) |
| `JIRA_MAX_BODY` | `--max-body` | Truncate issue descriptions to N chars, 0 = unlimited (default) |
| — | `--check` | Connect to Jira, print who you are authenticated as + visible project count to stderr, then exit (no server) |
| — | `--version` | Print version and exit |

## File access

None. HTTP GET to Jira only; the optional `JIRA_CA_CERT` bundle is read once at
startup.

## Usage examples

1. "What's assigned to me right now, highest priority first?" → `jira_my_issues`
2. "Find any tickets mentioning the login timeout bug — has anyone reported this before?" → `jira_search`
3. "Show me ABC-123 in full, including the comments and who changed its status." → `jira_get_issue` with `include_changelog=true`
4. "Everything resolved in project ABC in the last week, for the release notes." → `jira_search_jql` with `project = ABC AND resolved >= -7d`
5. "How healthy is project ABC — what's open, in progress, unassigned?" → `jira_project_status`
6. "Which projects can I see in Jira?" → `jira_list_projects`
7. "Draft a status report from my open tickets as a Word doc with tracked changes." → `jira_my_issues` + the `word` plugin's editing tools

## Troubleshooting

```
"C:\path\to\python.exe" jira.py --check
```

connects, authenticates and reports who you are plus how many projects you can
see. Behind an internal CA, point `--ca-cert` at the PEM bundle rather than
reaching for `--insecure`.
