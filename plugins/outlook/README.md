# Outlook (mail + calendar)

Read-only access to your local classic Outlook mail and calendar over COM, with
a content blacklist that withholds classified/compliance-marked items from the
AI entirely — plus a printable PDF day planner for any day.

| | |
|---|---|
| **Server** | `outlook.py` v6.2.0 |
| **pip install** | `pywin32` |
| **Platform** | **Windows only** — requires classic Win32 Outlook (not "New Outlook") installed, running, and logged into a profile |
| **Writes to disk** | only on request: an email saved as Markdown in `C:\Eva\knowledge\email`, or a day planner PDF in `C:\Eva\documents\pdf` |

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
| Printed day hours | `OUTLOOK_CALENDAR_HOURS` | Working day the printed planner's timeline starts from, e.g. `7-19` (default `8-18`) |
| Calendar category colours | `OUTLOOK_CALENDAR_COLOURS` | Usually leave blank — the planner already uses your Outlook category colours. Only overrules one, e.g. `Leadership=purple` |

## The printable day planner

Ask for "today's printable calendar", "tomorrow's planner" or "an agenda I can
take into the meeting" and `outlook_print_calendar` writes an **A4 landscape
PDF** into `C:\Eva\documents\pdf`. Print it single-sided and fold it in half:
the day itself runs down the left panel on an hour-by-hour timeline, the
following days are summarised on the right.

| You say | What you get |
|---|---|
| "Print today's calendar" | `Calendar - 2026-09-14 Monday.pdf` for today, with the next four days that have something on them |
| "Tomorrow's planner" | The same for tomorrow |
| "…without the attendee names" | `show_attendees: false`, for a sheet you can leave on a desk |
| "Show the next week down the side" | `lookahead_days: 6` |

The `date` argument takes `today` (the default), `tomorrow`, `yesterday`, an
offset like `+2`, or `YYYY-MM-DD`. Prefer the words: the server resolves them
against **this machine's** clock, so a planner cannot come out a day wrong
because the model's idea of today was stale.

### What ends up on the page

- **Meetings are blocks**, positioned and sized by their real start and end.
  Two that genuinely clash share the lane in columns the way Outlook's day view
  does; a run of back-to-back short meetings stays one column of thin blocks.
- **The timeline stretches.** It starts from your working day
  (`OUTLOOK_CALENDAR_HOURS`, default 08:00–18:00) and grows to take in anything
  outside it, so a 6 am flight is on the page rather than off the top of it.
- **All-day items** sit above the grid as a strip of chips.
- **Colour comes from your own Outlook categories, with nothing to configure.**
  Outlook already stores a colour against each category; the planner reads it
  off the profile and prints it at full strength. See
  [Category colours](#category-colours) below. Anything uncategorised falls back
  to what Outlook can actually prove: someone outside your SMTP domain is
  invited (External, green), it is internal (Internal, blue), or nobody is
  invited at all (Personal, grey). Nothing is inferred from the subject line,
  and the legend across the top names whichever of these are on the page.
- **Tentative or free-marked time is drawn hollow**, the way a diary pencils
  something in.
- **Empty days are skipped** in the right-hand panel, so a Friday sheet shows
  the week ahead instead of two blank weekend panels. Ask for the literal next
  four days and it passes `skip_empty_days: false`.
- **Blacklisted events never reach the page** — not even as an unlabelled
  block. The footer carries the count.

Re-printing a day overwrites that day's file rather than piling up copies.

### Category colours

Colour-code your diary in Outlook and the planner prints it back to you in the
same scheme. Outlook's swatches are deliberately muted on screen, which is
wrong for paper, so each of its 25 category colours maps to a bold equivalent:

| Outlook colour | On the page | Outlook colour | On the page |
|---|---|---|---|
| Red | `#EF4444` | Dark Red | `#B91C1C` |
| Orange | `#F97316` | Dark Orange | `#C2410C` |
| Peach | `#FDBA74` | Dark Peach | `#FB923C` |
| Yellow | `#FACC15` | Dark Yellow | `#CA8A04` |
| Green | `#10B981` | Dark Green | `#047857` |
| Teal | `#0D9488` | Dark Teal | `#0F766E` |
| Olive | `#84CC16` | Dark Olive | `#4D7C0F` |
| Blue | `#3B82F6` | Dark Blue | `#1D4ED8` |
| Purple | `#8B5CF6` | Dark Purple | `#6D28D9` |
| Maroon | `#9F1239` | Dark Maroon | `#881337` |
| Steel | `#64748B` | Dark Steel | `#475569` |
| Gray | `#9CA3AF` | Dark Gray | `#6B7280` |
| Black | `#1F2937` | *(None)* | falls through to the next category |

The subject on a block goes white or dark by itself, chosen from how bright the
fill is, so a Yellow category prints as dark type on yellow rather than
white-on-white. An appointment with several categories takes the first one that
has a colour.

`OUTLOOK_CALENDAR_COLOURS` is only for overruling one of these — say a category
Outlook has in Peach that you want loud on the page:

```powershell
setx OUTLOOK_CALENDAR_COLOURS "Board=dark-maroon,Leave=silver"
```

The names are those in the table (lower case, hyphenated: `dark-maroon`), plus
`amber`, `rose`, `silver`, `slate` and `grey`, which is the near-white block a
Personal appointment gets.

> **No PDF library is involved.** The planner is drawn straight into the PDF
> imaging model from the standard library alone, so printing adds nothing to the
> pip dependencies and there is nothing extra to transfer. Type is the base-14
> Helvetica every PDF reader carries; the design it follows uses Manrope, which
> would mean shipping a font file.

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

Of the four, this server uses three: `EVA_PYTHON`, `EVA_KNOWLEDGE_DIR` and
`EVA_DOCUMENTS_DIR`. Mail comes from Outlook over COM, so it **reads** no local
folder at all; the two below are where it writes, and only when you ask.

### The folders this plugin uses

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_KNOWLEDGE_DIR%\email` | Where `outlook_get_email` saves a message as Markdown **when the call asks for it** (`Email - <date> - <subject> (<id>).md`, overwritten if the same message is saved again), for the `knowledge-base` plugin to index. Blacklisted messages are never written | Created at startup. If it cannot be created the server refuses to start, rather than failing on the first email you ask it to keep |
| `%EVA_DOCUMENTS_DIR%\pdf` | Where `outlook_print_calendar` writes the day planner (`Calendar - <date> <weekday>.pdf`, overwritten when you re-print that day) | Created at startup. If it cannot be created the server **still starts** and says so; only printing is disabled, so mail and calendar stay readable |

Create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and they
both exist.

> **Why `documents\pdf` and not `documents\outlook`?** The document library is
> organised by **file type**, and a printed planner is a PDF. It lands beside
> your own PDFs rather than in an output folder of its own, which is the same
> rule `word` and `powerpoint` follow for the documents they create. Note the
> `pdf-to-md` plugin also reads that folder: a planner left there will be
> converted into the knowledge base along with everything else if you run a bulk
> conversion, which is rarely what you want, so tidy old sheets out or point
> `OUTLOOK_DOCS_DIR` somewhere of its own.

### Overriding one folder, and this server's own settings

The shared roots are normally all you need. These variables are this
server's own, and a folder variable here beats the matching root - use one
only when an endpoint's layout really differs.

| Variable | Purpose |
|---|---|
| `OUTLOOK_KB_DIR` | Full path to the save folder, instead of `%EVA_KNOWLEDGE_DIR%\email`. `off` forbids saving outright, after which no email is written to disk |
| `OUTLOOK_DOCS_DIR` | Full path to the day-planner folder, instead of `%EVA_DOCUMENTS_DIR%\pdf`. `off` forbids printing outright |
| `OUTLOOK_CALENDAR_HOURS` | The working day the printed timeline starts from, e.g. `7-19` (default `8-18`). It always stretches to fit anything scheduled outside it |
| `OUTLOOK_CALENDAR_COLOURS` | Overrules the colour Outlook already holds against a category, e.g. `Leadership=purple,Client=green` — see [Category colours](#category-colours). An unknown colour name is ignored with a warning |
| `OUTLOOK_KB_AUTOSAVE=true` | Save **every** email read, without being asked (default false). Needs a save folder to be on |
| `OUTLOOK_SEARCH_FOLDERS` | Comma-separated folder names used as the **default** set for `outlook_search_recent`, overriding the `SEARCH_ALL_FOLDERS` value in the file (e.g. `"Inbox,Sent Items,Archive"`). A per-call `folders` argument still takes priority |
| `OUTLOOK_BLACKLIST_FILE` | Path to a file of extra content-blacklist terms (one per line, `#` for comments), added to the built-in list |
| `OUTLOOK_REQUIRE_BLACKLIST=1` | Fail closed: refuse to start unless the content blacklist has at least one active term, so a missing or empty terms file cannot silently disable the compliance filter. Also settable via the `REQUIRE_BLACKLIST` constant in the file |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies. To forbid saving outright, set `OUTLOOK_KB_DIR=off` (`none`,
`no`, `false` and `disabled` work too); `OUTLOOK_DOCS_DIR=off` does the same for
printing. With both off the server writes no local file at all.

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
| `CALENDAR_DAY_START_HOUR` / `CALENDAR_DAY_END_HOUR` | Working day the printed planner's timeline starts from — `OUTLOOK_CALENDAR_HOURS` overrides both |
| `CALENDAR_LOOKAHEAD_DAYS` | How many following days the planner's right-hand panel lists by default (4); a per-call `lookahead_days` still takes priority |
| `CALENDAR_CATEGORY_COLOURS` | Category → colour overrides in the file itself, merged with (and beaten by) `OUTLOOK_CALENDAR_COLOURS`. Both sit on top of the colours read from Outlook |
| `PLANNER_COLOURS` / `OL_CATEGORY_COLOURS` | The bold palette, and which Outlook category colour maps to which name |
| `SEARCH_ALL_FOLDERS` | Folder names (matched across every store) that `outlook_search_recent` searches by default — `["Inbox", "Sent Items", "Archive"]`; use `outlook_list_folders` to see real folder names first. This is only the built-in default: override it with `OUTLOOK_SEARCH_FOLDERS`, or per call by passing a `folders` argument |

## File access

No local file access until a call asks for something to be written: an email
saved as Markdown inside the knowledge-base folder, or a day planner written as
a PDF inside the documents folder. It writes nowhere else, and it reads no local
folder at all (the optional blacklist file is read once at startup).

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
10. "Give me today's calendar to print." → `outlook_print_calendar` — an A4 landscape day planner in `C:\Eva\documents\pdf`
11. "Print tomorrow's planner, without the attendee names on it." → `outlook_print_calendar` with `date: "tomorrow"`, `show_attendees: false`

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

connects to Outlook and prints diagnostics, the folder paths it resolved and the
blacklist status. If it can't connect, confirm classic Outlook (not "New
Outlook") is running and logged into a profile.

> **A planner printed with no meetings on it** usually means the day really is
> empty — `outlook_get_calendar` for the same day is the quick check, and if
> that is empty too its `[debug]` section says what was scanned.
