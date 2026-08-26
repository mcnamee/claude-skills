# Confluence (read-only)

Search and read Confluence pages across one or two Confluence instances,
optionally mirroring every page read to Markdown so it feeds a local RAG
knowledge base.

| | |
|---|---|
| **Server** | `confluence.py` v2.0.0 |
| **pip install** | _none_ — standard library only (HTTP via stdlib `urllib`) |
| **Platform** | any |
| **Writes to disk** | yes — mirrors every page read to `C:\Eva\knowledge\confluence`, unless that setting is `off` |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install confluence@mcnamee-claude-skills
```

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| Confluence base URL | **yes** | `CONFLUENCE_BASE_URL` | Base URL including any context path, no trailing slash. This is the **default** server |
| Name for this Confluence server | no | `CONFLUENCE_NAME` | e.g. `Green`. Only matters if you configure a second server; defaults to `Primary` |
| Second Confluence base URL | no | `CONFLUENCE_BASE_URL_2` | Leave blank for a single-server setup |
| Name for the second Confluence server | no | `CONFLUENCE_NAME_2` | e.g. `Blue` — say it in a prompt to query that server; defaults to `Secondary` |
| Knowledge-base folder | no | `CONFLUENCE_KB_DIR` | Every page read is also saved as Markdown here, for the `knowledge-base` plugin to index. Defaults to `C:\Eva\knowledge\confluence`; `off` to disable |
| Python interpreter | **yes** | — | Absolute path to the `python.exe` to launch the server with |

> **Blank does not mean off.** Leaving the folder prompt empty means "not
> configured", so the default above applies. To switch mirroring off, type `off`
> (`none`, `no`, `false` and `disabled` work too), after which the server writes
> no local file at all.

**Your tokens are not stored in the plugin.** Set them as Windows user
environment variables before starting Claude Code — the plugin reads them from
the ambient environment. Credentials are deliberately env-var only: there are no
`--token`/`--user`/`--password` flags, because command-line arguments are
visible to other local users in process listings.

```powershell
setx CONFLUENCE_TOKEN   "token-for-the-first-server"
setx CONFLUENCE_TOKEN_2 "token-for-the-second-server"   # only if you have two
```

`setx` does not affect processes that are already running, so quit VS Code
completely (a window reload is not enough) and reopen it. Check it took in a
**new** window with `$env:CONFLUENCE_TOKEN`.

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

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

### First server (the default one)

| Env var | CLI flag | Purpose |
|---|---|---|
| `CONFLUENCE_NAME` | `--name` | Friendly name used to pick this server in a prompt (default `Primary`) |
| `CONFLUENCE_BASE_URL` | `--base-url` | Base URL incl. any context path, no trailing slash |
| `CONFLUENCE_TOKEN` | _(env only)_ | Personal Access Token, sent as Bearer (preferred over basic auth) |
| `CONFLUENCE_USER` | _(env only)_ | Username for basic auth (fallback if no token) |
| `CONFLUENCE_PASSWORD` | _(env only)_ | Password for basic auth |
| `CONFLUENCE_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `CONFLUENCE_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification |

### Second server (optional)

Setting `CONFLUENCE_BASE_URL_2` is what enables it.

| Env var | CLI flag | Purpose |
|---|---|---|
| `CONFLUENCE_NAME_2` | `--name-2` | Friendly name, e.g. `Blue` (default `Secondary`) |
| `CONFLUENCE_BASE_URL_2` | `--base-url-2` | Base URL of the second instance |
| `CONFLUENCE_TOKEN_2` | _(env only)_ | Its own Personal Access Token — tokens are per-instance |
| `CONFLUENCE_USER_2` | _(env only)_ | Username for basic auth on the second server |
| `CONFLUENCE_PASSWORD_2` | _(env only)_ | Password for basic auth on the second server |
| `CONFLUENCE_CA_CERT_2` | `--ca-cert-2` | CA bundle for the second server; falls back to `CONFLUENCE_CA_CERT` |
| `CONFLUENCE_VERIFY_SSL_2=false` | `--insecure-2` | TLS verification for the second server; falls back to the first server's setting |

### Shared by both servers

| Env var | CLI flag | Purpose |
|---|---|---|
| `CONFLUENCE_TIMEOUT` | `--timeout` | Request timeout in seconds (default 30) |
| `CONFLUENCE_MAX_BODY` | `--max-body` | Truncate page bodies to N chars, 0 = unlimited (default). Applies only to text returned to the model, not to files saved via the knowledge-base folder |
| `CONFLUENCE_KB_DIR` | `--kb-dir` | If set, every page read is also saved as a Markdown file into this folder — handy for feeding a local RAG knowledge base, e.g. alongside the `knowledge-base` plugin |
| — | `--check` | Connect to every configured server, print who you are authenticated as + visible space count to stderr, then exit (no server) |
| — | `--version` | Print version and exit |

## File access

No local file access unless the knowledge-base folder is set; then it writes
only inside that folder.

## Usage examples

1. "Search Confluence for our incident response runbook." → `confluence_search` on the default server
2. "Find pages about the onboarding process on Blue." → `confluence_search` with `server="Blue"`
3. "Find pages in the DOCS space that mention 'release notes' and were updated in the last 30 days." → `confluence_search_cql`
4. "Pull up the full content of Confluence page 393217." → `confluence_get_page`
5. "Open the 'Q3 Roadmap' page in the PROD space and summarise it." → `confluence_get_page_by_title`
6. "List every page under the 'Engineering Handbook' in the DOCS space, direct children only." → `confluence_list_pages_under`
7. "Is the retention policy on Green the same as the one on Blue?" → the same tool called once per server, then compared
8. "Pull the onboarding runbook into our local knowledge base for offline search." → `confluence_get_page` (or `confluence_get_page_by_title`), automatically mirrored to Markdown when the knowledge-base folder is configured, so the `knowledge-base` plugin's `kb_index`/`kb_ask` can find it afterwards

## Troubleshooting

`--check` connects to **every** configured server, authenticates and reports who
you are plus how many spaces you can see. Run it before wiring the server in; it
exits non-zero if any instance fails, so a two-server setup where only one
answers is caught here rather than mid-conversation:

```powershell
& "C:\path\to\python.exe" confluence.py --check
```

To try a pair of servers without touching the plugin config (tokens go in the
environment, never in the arguments):

```powershell
$env:CONFLUENCE_TOKEN   = "green-token"
$env:CONFLUENCE_TOKEN_2 = "blue-token"
& "C:\path\to\python.exe" confluence.py `
    --name Green --base-url https://green.confluence.example.com `
    --name-2 Blue --base-url-2 https://blue.confluence.example.com --check
```

Behind an internal CA, point `--ca-cert` at the PEM bundle rather than reaching
for `--insecure`. If the two instances sit behind different CAs, `--ca-cert-2`
covers the second one; leave it unset and the second server uses the first's
bundle.
