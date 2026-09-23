---
name: outlook
description: Read Outlook mail and calendar via the outlook MCP server (COM automation, Windows; cannot send anything), and print a day's calendar as a PDF day planner. Use when the user asks about their emails, wants mail searched or summarised, asks what's on their calendar, asks for a printable/printed calendar, day planner or agenda for today or tomorrow, or asks for an email to be saved into the knowledge base. For scheduling a meeting with someone, use the meeting-scheduler skill instead.
---

# Outlook (via the `outlook` MCP server)

Requires the `outlook.py` MCP server (Windows + classic Outlook running
and logged in). If its tools are not available, tell the user to wire it in
first (see the repo README) and to verify with `python outlook.py --check`.

**Scheduling a meeting is a different skill.** "Find a time with Josh", "book
a room", "set up 30 minutes next week" → use **`meeting-scheduler`**, which
covers `outlook_find_people`, `outlook_suggest_meeting_times` and
`outlook_draft_meeting`. Everything below is mail, calendar reading and
printing.

## Tools

| Tool | Use for |
|---|---|
| `outlook_list_recent_emails` | Recent Inbox messages (optionally unread only) |
| `outlook_search_emails` | Search Inbox by subject/sender |
| `outlook_get_email` | Full body of ONE message by its EntryID |
| `outlook_search_recent` | Search across Inbox/Sent/Archive in a date range (per-call `folders` override) |
| `outlook_list_sent_emails` | What the user sent in a date range |
| `outlook_get_calendar` | Calendar events in a date range (recurring expanded) |
| `outlook_print_calendar` | A **printable** PDF day planner for ONE day |
| `outlook_list_folders` | Real folder names across all stores |
| `outlook_find_people` | Fuzzy-find a person or room in the directory — see `meeting-scheduler` |
| `outlook_suggest_meeting_times` | When everyone is free — see `meeting-scheduler` |
| `outlook_draft_meeting` | Save an **unsent** meeting draft — see `meeting-scheduler` |

## Workflow

1. List/search first (`outlook_list_recent_emails` / `outlook_search_emails`
   / `outlook_search_recent`), then `outlook_get_email` with the EntryID of
   the message the user cares about — bodies are only available via get.
2. If a folder-scoped search misses, `outlook_list_folders` to learn the
   actual folder names, then retry `outlook_search_recent` with `folders`.
3. "What did I do last week?" → `outlook_list_sent_emails` + summarise.
4. "Print today's calendar" / "a planner for tomorrow" → `outlook_print_calendar`,
   **not** `outlook_get_calendar`. Pass `date: "today"` or `date: "tomorrow"`
   rather than working the date out yourself: the server resolves those words
   against the endpoint's own clock, so the sheet cannot come out a day wrong.

## Printing a day planner

`outlook_print_calendar` writes an A4 landscape PDF into `H:\Eva\documents\pdf`
and reports the path plus what is on the page. It is a bifold: the day on an
hour-by-hour timeline down the left half, the following days summarised on the
right. By default the page uses the top four fifths of the sheet, so the blank
bottom fifth folds up behind it and the planner fits a diary - but the endpoint
can set that to anything from half the sheet to all of it, so **take the folding
line from the tool's reply** rather than assuming a fifth. Tell the user the
path, that it prints single-sided at 100% (it is still a full A4 landscape
page), and how the reply says to fold it.

| They say | Call |
|---|---|
| "Print today's calendar" | `date: "today"` |
| "Tomorrow's planner" | `date: "tomorrow"` |
| "A planner for Friday" | `date: "+N"` if you can count the days, else `YYYY-MM-DD` |
| "Show me the next week down the side" | `lookahead_days: 6` |

- **One day per sheet.** There is no week or month layout; for a range, print
  each day or use `outlook_get_calendar` and answer in chat instead.
- **Do not re-read the calendar first.** The tool reads it itself and its reply
  lists what is on the page, so a preceding `outlook_get_calendar` is a wasted
  round trip - and a worse one here, since that tool withholds more (below).
- **Re-printing the same day overwrites that day's file.** That is deliberate,
  so say "updated" rather than warning about a clash.
- **Attendee names are never on the page** and there is no argument to put them
  there. If the user asks for them, say so rather than promising a flag.
- Block colour comes from the endpoint's Outlook **categories** — the colour
  Outlook already holds against each one, printed bolder than Outlook shows it —
  falling back to External / Internal / Personal, which the legend on the page
  names. Do not describe a colour as meaning anything else, and do not offer to
  "set up" colours: they are already the user's own.
- Empty days in the right-hand panel are skipped by default, so a Friday sheet
  shows the week ahead. Pass `skip_empty_days: false` if the user wants the
  literal next four days.

## Saving to the knowledge base

Reading an email does **not** save it. `outlook_get_email` takes `save_to_kb`,
false by default; set it to true **only when the user asks for that message to
be kept** — "save this email to the knowledge base", "add that thread to the
KB". The file lands in `H:\Eva\knowledge\email` and the tool reports the path.

- **Never set it while researching.** Mail you open to answer a question is not
  the user's filing decision. This is correspondence: saving it makes it
  embedded, searchable and quotable in later answers, which is a bigger step
  than saving a wiki page.
- **Asked after the fact** ("save that one") → call `outlook_get_email` again
  for that EntryID with `save_to_kb: true`.
- **Say what you saved** — the subject and the path. Suggest `kb_index` if the
  user wants it searchable immediately.
- Blacklisted messages are never saved: the content filter refuses them before
  any save is considered.

## Notes

- **It cannot send anything** — not mail, not a reply, not a meeting
  invitation — and cannot accept, delete or move anything. Never promise to.
  It writes three things, each only when asked: a saved email, a printed
  planner, and an **unsent** meeting draft (`outlook_draft_meeting`, which
  the user then sends themselves).
- A compliance blacklist may withhold messages/folders entirely; blocked
  items appear only as a withheld count. Do not speculate about their content,
  and never try to work around the filter.
- **The printed planner is the exception.** It never reads an appointment's
  body, so every event goes on the page - the PDF is written locally, on the
  user's own paper. A blocked event is left out of the tool's REPLY only, and
  the reply says how many. So the sheet can hold more meetings than the reply
  lists: say the page is complete, never that a meeting is missing from it,
  and do not ask the user what a withheld meeting was.
- Calendar date filtering is done in Python (locale-independent), so
  regional date settings cannot empty the results. If `outlook_get_calendar`
  finds nothing, its reply includes a `[debug]` section listing the last few
  calendar items scanned (start date + in-range flag) — read it to tell an
  actually-empty window apart from a filtering fault before retrying.
- **A meeting the user says is in Outlook but is not in the results** is a
  recurring-series fault, not an empty day. Recurring series are gathered two
  ways and merged (Outlook's own expansion, plus probing each series), so one
  method failing no longer hides a meeting. Tell the user to run
  `python outlook.py --check`, which reports per series what each method found
  and names any series one can see and the other cannot. Do not re-query with a
  wider date range hoping it appears, and never tell the user their calendar is
  empty when they have said otherwise.
