# Confluence (read-only)

Search and read Confluence pages across one or two Confluence instances, and
save a page to Markdown — when you ask for it — so it feeds a local RAG
knowledge base.

| | |
|---|---|
| **Server** | `confluence.py` v4.0.0 |
| **pip install** | _none_ — standard library only (HTTP via stdlib `urllib`) |
| **Platform** | any |
| **Writes to disk** | only when you ask a page to be saved — then one Markdown file in `C:\Eva\knowledge\confluence` |

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

The Python interpreter and the folder saved pages go to are **not** prompted for:
they come from the shared environment variables in
[Configuration](#configuration).

**Your tokens are not stored in the plugin.** Set them as Windows user
environment variables before starting Claude Code — the plugin reads them from
the ambient environment. Credentials are deliberately env-var only: there is no
flag that could put a token in a command line, where other local users would see
it in a process listing.

```powershell
setx CONFLUENCE_TOKEN   "token-for-the-first-server"
setx CONFLUENCE_TOKEN_2 "token-for-the-second-server"   # only if you have two
```

`setx` does not affect processes that are already running, so quit VS Code
completely (a window reload is not enough) and reopen it. Check it took in a
**new** window with `$env:CONFLUENCE_TOKEN`.

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
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `C:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `C:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

Of the four, this server uses two: `EVA_PYTHON` and `EVA_KNOWLEDGE_DIR`. It
reads no local folder at all - pages come over HTTP - so the knowledge folder is
the only thing it ever writes to.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `confluence`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_KNOWLEDGE_DIR%\confluence` | Where a page is saved as Markdown **when a tool call asks for it** (`save_to_kb=true`) - `Confluence - <title>.md`, or `Confluence <server> - <title>.md` with two instances - for the `knowledge-base` plugin to index | Created on demand |

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
| `CONFLUENCE_MAX_BODY` | Truncate page bodies to N chars, 0 = unlimited (default). Applies only to text returned to the model, not to saved files |
| `CONFLUENCE_KB_DIR` | Full path to the save folder, instead of `%EVA_KNOWLEDGE_DIR%\confluence`. `off` forbids saving outright, after which the server writes no local file at all |
| `CONFLUENCE_KB_AUTOSAVE=true` | Save **every** page read, without being asked (default false). Needs a save folder to be on |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies. To forbid saving outright, set `CONFLUENCE_KB_DIR=off`
(`none`, `no`, `false` and `disabled` work too).

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Connect to every configured server, print who you are authenticated as + visible space count to stderr, then exit (no server). Non-zero if any server fails |
| `--version` | Print version and exit |

## File access

No local file access until a tool call asks for a page to be saved; then it
writes one Markdown file inside the knowledge-base folder, and nowhere else.

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

## Troubleshooting

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
