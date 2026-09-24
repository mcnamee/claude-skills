# Outlook (mail + calendar)

Access to your local classic Outlook mail and calendar over COM, with a content
blacklist that withholds classified/compliance-marked items from the AI entirely
— plus meeting scheduling against the enterprise directory, and a printable PDF
day planner for any day.

| | |
|---|---|
| **Server** | `outlook.py` v10.1.0 |
| **pip install** | `pywin32` |
| **Platform** | **Windows only** — requires classic Win32 Outlook (not "New Outlook") installed, running, and logged into a profile |
| **Sends** | **never.** Not an email, not a reply, not a meeting invitation. It cannot accept, move or delete anything either |
| **Writes to the mailbox** | one thing: an **unsent** meeting draft in your calendar, which you review and send yourself (`OUTLOOK_ALLOW_DRAFTS=false` removes even that) |
| **Writes to disk** | only on request: an email saved as Markdown in `H:\Eva\knowledge\email`, or a day planner PDF in `H:\Eva\documents\pdf` |

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
| Printed page fill | `OUTLOOK_CALENDAR_PAGE_FILL` | How much of the sheet's height the planner uses: `0.8` (default), `80`, or `off` for the whole sheet |
| Calendar category colours | `OUTLOOK_CALENDAR_COLOURS` | Usually leave blank — the planner already uses your Outlook category colours. Only overrules one, e.g. `Leadership=purple` |
| Hidden calendar categories | `OUTLOOK_CALENDAR_HIDE_CATEGORIES` | Comma-separated Outlook categories to leave off the printed planner, e.g. `Birthdays` |
| Meeting booking hours | `OUTLOOK_MEETING_HOURS` | Hours a meeting may be *suggested* in, e.g. `8-18` (default `9-17`) |
| Allow meeting drafts | `OUTLOOK_ALLOW_DRAFTS` | `false` removes `outlook_draft_meeting`, leaving the mailbox read-only |
| Directory scan cap | `OUTLOOK_GAL_SCAN_CAP` | Entries one fuzzy name search may read (default `20000`) |

## Scheduling a meeting

> *"Find a suitable time with Josh Smith in Room R5-3-84."*

Three tools, meant to be used in that order. The `meeting-scheduler` skill
drives them; this is what they actually do.

| Tool | What it does |
|---|---|
| `outlook_find_people` | Fuzzy-finds a person or room across the Global Address List, the All Rooms list and your Contacts |
| `outlook_suggest_meeting_times` | Reads everyone's published free/busy (including yours and the room's) and offers the slots that fit |
| `outlook_draft_meeting` | Saves the meeting into your calendar **unsent** and opens it for you to send |

### Finding people the way you'd say it

You say "Josh Smith". The GAL says "Smith, Joshua P". The room you call
R5-3-84 is filed as "Room R5-3-84 (12 seats)". Matching is token-based and
order-independent, with `difflib` behind it for typos, so all three land — and
so does `josh smtih`.

It searches cheapest-first: Outlook's own ambiguous-name resolution (an
unmistakable name costs no scan at all), then Contacts, then the address
lists. Only the display *name* is read for each directory entry — that is one
COM round trip each, and an enterprise GAL is large — with the expensive
lookups reserved for the handful that score well enough to show.

Pass `kind: "room"` for a room. It searches the room lists first and is both
faster and far more accurate. A room is told apart from a person by
`PR_DISPLAY_TYPE_EX`, because `AddressEntry.DisplayType` has no room in its
enumeration at all; where a cached offline address book withholds that
property, membership of an "All Rooms"-style list decides it, and failing
that the name is read for it.

Each match reports an **`invite_as`** address. That is what the other two
tools take — they will not accept a name.

On a very large GAL the scan stops at `OUTLOOK_GAL_SCAN_CAP` entries and says
so. It runs in the directory's own order rather than by relevance, so a
shorter search doesn't move where it stopped — give the full name or the email
address (resolved directly, no scan) or raise the cap.

### Finding a time

Free/busy comes from `Recipient.FreeBusy`, the same published data Outlook's
own Scheduling Assistant uses. Three judgements are worth knowing about:

- **A tentative *room* is taken; a tentative *person* is not.** A tentatively
  held room is already someone's claim on the space. A tentative person is
  worth asking, and an enterprise calendar is full of them — so that slot is
  offered, with a note.
- **An optional attendee never blocks a slot**, but a slot they can't make is
  labelled with their name, and slots they can make rank higher.
- **Unpublished free/busy is not "free".** External attendees and mailboxes
  that don't publish come back empty, and the reply names those people rather
  than quietly scheduling over them.

Suggestions are spread across days before they're printed. At the default
quarter-hour granularity the five best slots would otherwise be the same
afternoon shifted 15 minutes at a time — five options on paper and one in
practice.

Exchange publishes **30 days** of free/busy and no more, so a longer window is
clamped to it and the reply says where it stopped. `OUTLOOK_MEETING_HOURS`
sets the bookable window (default 09:00–17:00); it is deliberately separate
from `OUTLOOK_CALENDAR_HOURS`, because what you want to *see* on a printed day
and when it's acceptable to *book* you are not the same thing.

### The draft

`outlook_draft_meeting` creates the meeting, attaches the attendees and the
room, **saves** it and opens it on screen. `Send()` is never called anywhere in
the file, so nobody is notified until you press Send yourself.

- **If any attendee won't resolve, nothing at all is saved.** A
  half-addressed draft looks finished and gets sent without the person it was
  for.
- Rooms go in `rooms`, not `location` — that books them as a resource *and*
  fills in the Location. `location` alone is just text and reserves nothing.
- `start` is strictly `YYYY-MM-DD HH:MM`, 24-hour. `16/09/2026` is refused on
  purpose: a meeting written at the wrong time is worse than a rejected
  argument, and locale-dependent date parsing is the same fault that made
  `Restrict()` unusable elsewhere in this server.
- The content blacklist applies in this direction too — a subject or body
  carrying a protective marking is refused rather than written into Outlook.

**The subject and the agenda** are ordinary arguments, and the usual shape of
the request is two steps — *"draft an agenda for the migration review"*, then
*"now find a time with Josh Mann to discuss"*. The second step should carry
the first one's agenda into `body` rather than asking you to retype it, and
should show you the subject and body before drafting: everyone invited reads
them.

`AppointmentItem.Body` is **plain text** and Outlook renders no Markdown, so an
agenda drafted in conversation would otherwise arrive with its asterisks still
on it. The server strips `**bold**`, leading `#` headings, and normalises `*`
and `+` bullets to `-`. It does no more than that on purpose: a lone asterisk
is as likely to be a footnote mark as emphasis, and `C#` is not a heading.

Set `OUTLOOK_ALLOW_DRAFTS=false` and the tool is not offered at all (not
listed, not dispatched), leaving the server read-only on the mailbox. Removing
it beats refusing it later: nothing can promise you a draft it was never
offered.

## The printable day planner

Ask for "today's printable calendar", "tomorrow's planner" or "an agenda I can
take into the meeting" and `outlook_print_calendar` writes an **A4 landscape
PDF** into `H:\Eva\documents\pdf`. Print it single-sided and fold it in half:
the day itself runs down the left panel on an hour-by-hour timeline, the
following days are summarised on the right.

The sheet prints on its **top four fifths**. The bottom fifth comes out blank
with a hairline to fold against, so you can fold it up behind the page and the
planner fits a diary. It is still a full A4 landscape page - print it at 100%,
no scaling.

`OUTLOOK_CALENDAR_PAGE_FILL` changes how much of the sheet is used: a fraction
(`0.8`), a percentage (`80` or `80%`), or `off` for the whole sheet with nothing
left to fold. It accepts `0.5` to `1.0` — below half a page the header, grid and
footer stop fitting, so a smaller value is refused with a warning and the
default kept. Whatever it is set to, the tool's reply says how the sheet folds,
so Claude repeats the right instruction rather than assuming a fifth.

| You say | What you get |
|---|---|
| "Print today's calendar" | `Calendar - 2026-09-14 Monday.pdf` for today, with the next four days that have something on them |
| "Tomorrow's planner" | The same for tomorrow |
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
- **No colour legend across the top.** The blocks carry the colour; a printed
  page is looked at rather than decoded.
- **Attendees are never printed.** Who is in a meeting is in the invitation, it
  crowds out the subject and location the sheet is for, and it puts other
  people's names on something left on a desk. They are still read, since an
  outside-domain attendee is what colours an uncategorised block.
- **One type size for every subject** (7pt), whatever block it sits on. The
  size used to follow the size of the block, which made an hour-long meeting
  shout and a fifteen-minute one whisper for no reason other than its duration.
  It is small on purpose: a planner is read at desk distance, and a subject
  that fits on its block beats large type with an ellipsis.
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
- **Categories you don't want on paper can be left off.** Set
  `OUTLOOK_CALENDAR_HIDE_CATEGORIES="Birthdays"` (comma-separate several) and
  any appointment carrying one of those categories is not printed in either
  panel, does not stretch the timeline, and does not make a day count as busy.
  Names match Outlook's category names, ignoring case. The tool's reply says
  how many were left off. Printing only: `outlook_get_calendar` still shows
  them.
- **Empty days are skipped** in the right-hand panel, so a Friday sheet shows
  the week ahead instead of two blank weekend panels. Ask for the literal next
  four days and it passes `skip_empty_days: false`.
- **Every event is printed, blacklisted or not.** The PDF is drawn on this
  machine and printed on your own paper, and a planner missing the meetings
  that matter most is not worth carrying. What makes that safe is that the
  planner **never reads an appointment's body** — the part that carries marked
  detail — so nothing on the page or in the reply comes from it. The blacklist
  still decides what the AI is told: it is matched against the subject,
  location, organiser and categories, and an event that matches is printed but
  left out of the tool's reply, which reports how many. `outlook_get_calendar`
  is stricter, and unchanged: it reads the body and withholds a matching event
  outright, because its whole output goes to the model.

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

## Recurring meetings

A recurring series is not stored as individual meetings. The folder holds one
master, and the occurrences have to be worked out. This plugin does it **twice
and merges the results**, because either method alone loses meetings:

1. **Outlook's own expansion** — sorting the items by `[Start]` and setting
   `IncludeRecurrences` makes Outlook hand back individual occurrences. Outlook
   holds each series' timezone and daylight-saving rules, so it places every
   occurrence correctly. This is the primary source.
2. **Probing each master's recurrence pattern**, day by day. This only finds an
   occurrence whose exact start time was guessed in advance, so it is a backstop.

**Before v6.3.0 only method 2 ran**, and it took that start time from the
master's own `Start` — which is the series' **first** occurrence, often years
old. Any series whose occurrences no longer begin at that same wall-clock time
was invisible: nothing errored, the meetings simply were not there. That covers
a series set up under a daylight-saving offset that has since drifted, and one
stored in an overseas organiser's timezone, which is why a long-standing weekly
meeting could go missing while a new one showed up fine.

Neither method formats a date string, so this does not reintroduce the
`Restrict()` fault where regional settings silently empty the results.

### Checking your own calendar

```powershell
& $env:EVA_PYTHON outlook.py --check
```

reports, per recurring series, what each method found for today, and names any
series one can see and the other cannot:

```
Recurring-series check, 2026-09-14 to 2026-09-14
  Outlook's own expansion : 12 occurrence(s) after walking 4180 item(s)
  Recurring series in the folder: 34
  Found by probing each series  : 9 occurrence(s)

  3 series would have been MISSED without Outlook's expansion:
    - Weekly Leadership
        probed at 10:00, actually starts 11:00
```

If a meeting is in Outlook and still not in the results, that report is the
thing to send: it distinguishes an empty day from an expansion fault.

## Skills

This plugin ships two, and installing it installs both:

| Skill | For |
|---|---|
| `/outlook:outlook` | Mail, reading the calendar, printing a day planner, saving an email to the knowledge base |
| `/outlook:meeting-scheduler` | Arranging a meeting: finding the person and the room, finding a time, saving the draft |

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

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

Of the four, this server uses three: `EVA_PYTHON`, `EVA_KNOWLEDGE_DIR` and
`EVA_DOCUMENTS_DIR`. Mail comes from Outlook over COM, so it **reads** no local
folder at all; the two below are where it writes, and only when you ask.

### The folders this plugin uses

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_KNOWLEDGE_DIR%\email` | Where `outlook_get_email` saves a message as Markdown **when the call asks for it** (`Email - <date> - <subject> (<id>).md`, overwritten if the same message is saved again), for the `knowledge-base` plugin to index. Blacklisted messages are never written | Created at startup. If it cannot be created the server refuses to start, rather than failing on the first email you ask it to keep |
| `%EVA_DOCUMENTS_DIR%\pdf` | Where `outlook_print_calendar` writes the day planner (`Calendar - <date> <weekday>.pdf`, overwritten when you re-print that day) | Created at startup. If it cannot be created the server **still starts** and says so; only printing is disabled, so mail and calendar stay readable |

Create them, or copy the repo's [`eva/`](../../eva) folder to `H:\Eva` and they
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
| `OUTLOOK_CALENDAR_PAGE_FILL` | How much of the sheet's **height** the printed planner uses: a fraction (`0.8`, the default), a percentage (`80` or `80%`), or `off` (`none`, `no`, `false`, `disabled`) for the whole sheet with nothing left to fold. Accepts `0.5` to `1.0`; anything else is refused with a warning and the default kept |
| `OUTLOOK_CALENDAR_COLOURS` | Overrules the colour Outlook already holds against a category, e.g. `Leadership=purple,Client=green` — see [Category colours](#category-colours). An unknown colour name is ignored with a warning |
| `OUTLOOK_CALENDAR_HIDE_CATEGORIES` | Comma-separated Outlook categories to leave **off** the printed planner, e.g. `Birthdays,Public Holidays`. Case-insensitive, whole category name. Printing only — `outlook_get_calendar` is unaffected |
| `OUTLOOK_KB_AUTOSAVE=true` | Save **every** email read, without being asked (default false). Needs a save folder to be on |
| `OUTLOOK_SEARCH_FOLDERS` | Comma-separated folder names used as the **default** set for `outlook_search_recent`, overriding the `SEARCH_ALL_FOLDERS` value in the file (e.g. `"Inbox,Sent Items,Archive"`). A per-call `folders` argument still takes priority |
| `OUTLOOK_BLACKLIST_FILE` | Path to a file of extra content-blacklist terms (one per line, `#` for comments), added to the built-in list |
| `OUTLOOK_REQUIRE_BLACKLIST=1` | Fail closed: refuse to start unless the content blacklist has at least one active term, so a missing or empty terms file cannot silently disable the compliance filter. Also settable via the `REQUIRE_BLACKLIST` constant in the file |
| `OUTLOOK_MEETING_HOURS` | The hours a meeting may be **suggested** in, e.g. `8-18` (default `9-17`). Separate from `OUTLOOK_CALENDAR_HOURS` above on purpose — see [Finding a time](#finding-a-time) |
| `OUTLOOK_GAL_SCAN_CAP` | How many address-list entries one fuzzy name search may read before giving up (default `20000`). Raise it on a very large GAL if people go missing; a search that stops early says so in its reply |
| `OUTLOOK_ALLOW_DRAFTS=false` | Remove `outlook_draft_meeting` altogether, leaving the server read-only on the mailbox |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies. To forbid saving outright, set `OUTLOOK_KB_DIR=off` (`none`,
`no`, `false` and `disabled` work too); `OUTLOOK_DOCS_DIR=off` does the same for
printing. With both off the server writes no local file at all.

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Connect to Outlook, print diagnostics, folder paths, blacklist status, the address lists and free/busy the scheduling tools can see, and the [recurring-series report](#checking-your-own-calendar), to stderr, then exit (no server) |
| `--version` | Print version and exit (works even without `pywin32` installed) |

## The content blacklist

Items whose content matches a blacklisted term are withheld from the AI
entirely. The blacklist also applies to **folder names**: folders whose
store/path matches a blacklisted term are withheld from `outlook_list_folders`
and skipped by `outlook_search_recent` (results are labelled with their folder
path, so a marked folder name never appears in output).

**The printed planner is the one exception, and it is a narrower one than it
looks.** `outlook_print_calendar` never reads an appointment's body at all, so
the part of an event that carries marked detail cannot reach the model by any
route — the page and the reply are built from the subject, times, location and
categories only. Because nothing sensitive can leak, every event is printed on
your own paper; the blacklist is matched against those visible fields instead,
and an event that matches is left out of the tool's **reply** while still
appearing on the page. `outlook_get_calendar` is unchanged: it reads the body
and withholds a matching event outright.

Everything else is configured by editing the `USER CONFIGURATION` block at the
top of `outlook.py` directly (there are no CLI flags/env vars for these):

| Setting | Purpose |
|---|---|
| `BLACKLIST_TERMS` | Built-in list of classification/compliance terms that cause an item to be withheld from the AI entirely |
| `BLACKLIST_MATCH_MODE` | `"word"` (default, whole-term match) or `"substring"` (for terms containing punctuation) |
| `MAX_BODY_CHARS` / `CALENDAR_HARD_CAP` / `SEARCH_SCAN_CAP` | Safety caps on body length / items scanned |
| `RECURRENCE_SCAN_CAP` | Ceiling on the expanded-occurrence walk. A series with no end date makes the collection unbounded, so the walk stops here and says so rather than reporting an empty day |
| `CALENDAR_DAY_START_HOUR` / `CALENDAR_DAY_END_HOUR` | Working day the printed planner's timeline starts from — `OUTLOOK_CALENDAR_HOURS` overrides both |
| `CALENDAR_LOOKAHEAD_DAYS` | How many following days the planner's right-hand panel lists by default (4); a per-call `lookahead_days` still takes priority |
| `CALENDAR_CATEGORY_COLOURS` | Category → colour overrides in the file itself, merged with (and beaten by) `OUTLOOK_CALENDAR_COLOURS`. Both sit on top of the colours read from Outlook |
| `CALENDAR_HIDDEN_CATEGORIES` | Categories left off the printed planner, in the file itself. Added to `OUTLOOK_CALENDAR_HIDE_CATEGORIES` |
| `PLANNER_COLOURS` / `OL_CATEGORY_COLOURS` | The bold palette, and which Outlook category colour maps to which name |
| `PLANNER_TITLE_PT` | The one type size every event subject on the page is set at (7pt) |
| `PLANNER_CONTENT_FRACTION` | Default for how much of the sheet's height the planner prints on (`0.8` — the top four fifths, leaving the bottom fifth blank to fold up). `OUTLOOK_CALENDAR_PAGE_FILL` overrides it, so the file rarely needs editing |
| `PLANNER_MIN_CONTENT_FRACTION` | The least of the sheet the planner may be squeezed into (`0.5`); a smaller `OUTLOOK_CALENDAR_PAGE_FILL` is refused |
| `SEARCH_ALL_FOLDERS` | Folder names (matched across every store) that `outlook_search_recent` searches by default — `["Inbox", "Sent Items", "Archive"]`; use `outlook_list_folders` to see real folder names first. This is only the built-in default: override it with `OUTLOOK_SEARCH_FOLDERS`, or per call by passing a `folders` argument |

## File access

No local file access until a call asks for something to be written: an email
saved as Markdown inside the knowledge-base folder, or a day planner written as
a PDF inside the documents folder. It writes nowhere else, and it reads no local
folder at all (the optional blacklist file is read once at startup).

In the **mailbox**, the one write is `outlook_draft_meeting`, which saves an
unsent meeting into your calendar. Nothing can send, reply, accept, move or
delete — `Send()` is never called anywhere in the file, and
`OUTLOOK_ALLOW_DRAFTS=false` removes the draft tool as well.

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
10. "Give me today's calendar to print." → `outlook_print_calendar` — an A4 landscape day planner in `H:\Eva\documents\pdf`
11. "Print tomorrow's planner with the whole week down the side." → `outlook_print_calendar` with `date: "tomorrow"`, `lookahead_days: 6`
12. "Who's Josh Smith in the directory?" → `outlook_find_people` — returns the `invite_as` address the scheduling tools need
13. "Is there a room called R5-3-84?" → `outlook_find_people` with `kind: "room"`
14. "Find a suitable time with Josh Smith in Room R5-3-84." → `outlook_find_people` twice, then `outlook_suggest_meeting_times`
15. "When are Josh and Sam both free for an hour next week?" → `outlook_suggest_meeting_times` with `duration_minutes: 60`, `start_date: "+7"`
16. "Book the Tuesday 10am one." → `outlook_draft_meeting` — saved to your calendar **unsent**, opened for you to review and send
17. "Draft an agenda for the migration review." … "Now find a time with Josh Mann to discuss." → the agenda is carried into the meeting body, shown to you, then drafted

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
