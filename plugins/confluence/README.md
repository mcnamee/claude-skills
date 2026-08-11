# Confluence (read-only)

Search and read Confluence pages, optionally mirroring every page read to
Markdown so it feeds a local RAG knowledge base.

| | |
|---|---|
| **Server** | `confluence.py` v1.3.1 |
| **pip install** | _none_ — standard library only (HTTP via stdlib `urllib`) |
| **Platform** | any |
| **Writes to disk** | only if the knowledge-base folder is set |

## Install

```
/plugin marketplace add C:\path\to\mcp-servers
/plugin install confluence@mcnamee-mcp-servers
```

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| Confluence base URL | **yes** | `CONFLUENCE_BASE_URL` | Base URL including any context path, no trailing slash |
| Knowledge-base folder | no | `CONFLUENCE_KB_DIR` | If set, every page read is also saved as Markdown here |
| Python interpreter | **yes** | — | Absolute path to the `python.exe` to launch the server with |

**Your token is not stored in the plugin.** Set `CONFLUENCE_TOKEN` as a Windows
user environment variable before starting Claude Code — the plugin reads it from
the ambient environment. Credentials are deliberately env-var only: there are no
`--token`/`--user`/`--password` flags, because command-line arguments are
visible to other local users in process listings.

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| Env var | CLI flag | Purpose |
|---|---|---|
| `CONFLUENCE_BASE_URL` | `--base-url` | Base URL incl. any context path, no trailing slash |
| `CONFLUENCE_TOKEN` | _(env only)_ | Personal Access Token, sent as Bearer (preferred over basic auth) |
| `CONFLUENCE_USER` | _(env only)_ | Username for basic auth (fallback if no token) |
| `CONFLUENCE_PASSWORD` | _(env only)_ | Password for basic auth |
| `CONFLUENCE_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `CONFLUENCE_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification |
| `CONFLUENCE_TIMEOUT` | `--timeout` | Request timeout in seconds (default 30) |
| `CONFLUENCE_MAX_BODY` | `--max-body` | Truncate page bodies to N chars, 0 = unlimited (default). Applies only to text returned to the model, not to files saved via the knowledge-base folder |
| `CONFLUENCE_KB_DIR` | `--kb-dir` | If set, every page read is also saved as a Markdown file (`Confluence - <title>.md`, overwritten each time) into this folder — handy for feeding a local RAG knowledge base, e.g. alongside the `knowledge-base` plugin |
| — | `--check` | Connect to Confluence, print who you are authenticated as + visible space count to stderr, then exit (no server) |
| — | `--version` | Print version and exit |

## File access

No local file access unless the knowledge-base folder is set; then it writes
only inside that folder.

## Usage examples

1. "Search Confluence for our incident response runbook." → `confluence_search`
2. "Find pages in the DOCS space that mention 'release notes' and were updated in the last 30 days." → `confluence_search_cql`
3. "Pull up the full content of Confluence page 393217." → `confluence_get_page`
4. "Open the 'Q3 Roadmap' page in the PROD space and summarise it." → `confluence_get_page_by_title`
5. "List every page under the 'Engineering Handbook' in the DOCS space, direct children only." → `confluence_list_pages_under`
6. "Pull the onboarding runbook into our local knowledge base for offline search." → `confluence_get_page` (or `confluence_get_page_by_title`), automatically mirrored to Markdown when the knowledge-base folder is configured, so the `knowledge-base` plugin's `kb_index`/`kb_ask` can find it afterwards

## Troubleshooting

`--check` connects, authenticates and reports who you are plus how many spaces
you can see — run it before wiring the server in:

```
"C:\path\to\python.exe" confluence.py --check
```

Behind an internal CA, point `--ca-cert` at the PEM bundle rather than reaching
for `--insecure`.
