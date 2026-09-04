# Outlook (mail + calendar)

Read-only access to your local classic Outlook mail and calendar over COM, with
a content blacklist that withholds classified/compliance-marked items from the
AI entirely.

| | |
|---|---|
| **Server** | `outlook.py` v6.0.0 |
| **pip install** | `pywin32` |
| **Platform** | **Windows only** — requires classic Win32 Outlook (not "New Outlook") installed, running, and logged into a profile |
| **Writes to disk** | only when you ask an email to be saved — then one Markdown file in `C:\Eva\knowledge\email` |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install outlook@mcnamee-claude-skills
```

Claude Code prompts only for the two optional settings below; the Python
interpreter and the folder saved email goes to come from the shared environment
variables in [Configuration](#configuration).

| Prompt | Env var | Purpose |
|---|---|---|
| Search folders | `OUTLOOK_SEARCH_FOLDERS` | Comma-separated default folder set for `outlook_search_recent`, e.g. `Inbox,Sent Items,Archive` |
| Blacklist file | `OUTLOOK_BLACKLIST_FILE` | Path to a file of extra content-blacklist terms |

## Saving to the knowledge base

**Reading an email does not save it.** `outlook_get_email` takes a `save_to_kb`
argument, false by default; the message is written to the knowledge-base folder
only when it is true, which Claude sets when you ask for that email to be kept:

| You say | What happens |
|---|---|
| "Summarise the vendor's last email about the renewal" | The message is read to answer. Nothing is saved |
| "Save that email to the knowledge base" | `save_to_kb=true` — one Markdown file, then `kb_index` can pick it up |

This is the plugin where the distinction matters most: saving turns
correspondence into plain text files that are then embedded and quotable in
answers, so it should be a decision, not a side effect of reading your mail.
Blacklisted messages are never written either way — the content filter runs
first, so a blocked message is refused before any save is considered.

The saved file is the full message — `MAX_BODY_CHARS` truncates only what is
returned to the model — named `Email - <date> - <subject> (<id>).md`, so saving
the same message again overwrites its own file rather than piling up copies. If
saving is switched off (`off`) and you ask for a message to be kept anyway, the
tool says so instead of failing silently.

To go back to saving every email read, as versions before 5.0.0 did, set
`OUTLOOK_KB_AUTOSAVE=true`:

```powershell
setx OUTLOOK_KB_AUTOSAVE "true"
```

See [`eva/knowledge/email`](../../eva/knowledge/email).

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

Of the four, this server uses two: `EVA_PYTHON` and `EVA_KNOWLEDGE_DIR`. Mail
comes from Outlook over COM, so it reads no local folder at all.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `email`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_KNOWLEDGE_DIR%\email` | Where `outlook_get_email` saves a message as Markdown **when the call asks for it** (`Email - <date> - <subject> (<id>).md`, overwritten if the same message is saved again), for the `knowledge-base` plugin to index. Blacklisted messages are never written | Created at startup. If it cannot be created the server refuses to start, rather than failing on the first email you ask it to keep |

> This server reads no local folder at all - mail comes from Outlook over COM.
> The knowledge folder is the only thing it ever writes to.

### Overriding one folder, and this server's own settings

The shared roots are normally all you need. These variables are this
server's own, and a folder variable here beats the matching root - use one
only when an endpoint's layout really differs.

| Variable | Purpose |
|---|---|
| `OUTLOOK_KB_DIR` | Full path to the save folder, instead of `%EVA_KNOWLEDGE_DIR%\email`. `off` forbids saving outright, after which the server writes no local file at all |
| `OUTLOOK_KB_AUTOSAVE=true` | Save **every** email read, without being asked (default false). Needs a save folder to be on |
| `OUTLOOK_SEARCH_FOLDERS` | Comma-separated folder names used as the **default** set for `outlook_search_recent`, overriding the `SEARCH_ALL_FOLDERS` value in the file (e.g. `"Inbox,Sent Items,Archive"`). A per-call `folders` argument still takes priority |
| `OUTLOOK_BLACKLIST_FILE` | Path to a file of extra content-blacklist terms (one per line, `#` for comments), added to the built-in list |
| `OUTLOOK_REQUIRE_BLACKLIST=1` | Fail closed: refuse to start unless the content blacklist has at least one active term, so a missing or empty terms file cannot silently disable the compliance filter. Also settable via the `REQUIRE_BLACKLIST` constant in the file |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies. To forbid saving outright, set `OUTLOOK_KB_DIR=off` (`none`,
`no`, `false` and `disabled` work too).

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Connect to Outlook, print diagnostics + blacklist status to stderr, then exit (no server) |
| `--version` | Print version and exit (works even without `pywin32` installed) |

## The content blacklist

Items whose content matches a blacklisted term are withheld from the AI
entirely. The blacklist also applies to **folder names**: folders whose
store/path matches a blacklisted term are withheld from `outlook_list_folders`
and skipped by `outlook_search_recent` (results are labelled with their folder
path, so a marked folder name never appears in output).

Everything else is configured by editing the `USER CONFIGURATION` block at the
top of `outlook.py` directly (there are no CLI flags/env vars for these):

| Setting | Purpose |
|---|---|
| `BLACKLIST_TERMS` | Built-in list of classification/compliance terms that cause an item to be withheld from the AI entirely |
| `BLACKLIST_MATCH_MODE` | `"word"` (default, whole-term match) or `"substring"` (for terms containing punctuation) |
| `MAX_BODY_CHARS` / `CALENDAR_HARD_CAP` / `SEARCH_SCAN_CAP` | Safety caps on body length / items scanned |
| `SEARCH_ALL_FOLDERS` | Folder names (matched across every store) that `outlook_search_recent` searches by default — `["Inbox", "Sent Items", "Archive"]`; use `outlook_list_folders` to see real folder names first. This is only the built-in default: override it with `OUTLOOK_SEARCH_FOLDERS`, or per call by passing a `folders` argument |

## File access

No local file access until a call asks for an email to be saved; then it writes
one Markdown file inside the knowledge-base folder, and nowhere else. The
optional blacklist file is read once at startup.

## Usage examples

1. "Show me my 10 most recent unread emails." → `outlook_list_recent_emails`
2. "Search my inbox for anything from 'Jane Smith' about the contract renewal." → `outlook_search_emails`
3. "Open that email from the vendor and summarise the key dates." → `outlook_get_email` — read only, nothing saved
4. "Save these project emails into my RAG knowledge base as Markdown." → `outlook_get_email` with `save_to_kb=true`
5. "What's on my calendar for the next 7 days?" → `outlook_get_calendar`
6. "What did I send last week?" → `outlook_list_sent_emails`
7. "Find everything about the 'Acme renewal' across my Inbox, Sent Items and Archive from the last month." → `outlook_search_recent`
8. "Search only my 'Projects' and 'Sent Items' folders for anything about the budget review." → `outlook_search_recent` with a `folders` argument overriding the default set
9. "What are my actual Outlook folder names, so I can point the search at the right archive?" → `outlook_list_folders`

The save folder sits inside the same knowledge root the `knowledge-base` server
indexes - which is what `EVA_KNOWLEDGE_DIR` being one shared setting buys you -
so the emails you keep land alongside your Confluence pages and Word documents.

## Troubleshooting

> **If a server fails with `Executable not found in $PATH: "${EVA_PYTHON}"`**,
> the variable is not set in the environment Claude Code was launched from. Set
> it (see above), then quit Claude Code completely and reopen — `setx` and
> `[Environment]::SetEnvironmentVariable` do not reach a process that is already
> running.

```powershell
& $env:EVA_PYTHON outlook.py --check
```

connects to Outlook and prints diagnostics plus the blacklist status. If it
can't connect, confirm classic Outlook (not "New Outlook") is running and logged
into a profile.
