---
name: outlook
description: Read Outlook mail and calendar via the outlook MCP server (read-only COM automation, Windows), and print a day's calendar as a PDF day planner. Use when the user asks about their emails, wants mail searched or summarised, asks what's on their calendar, asks for a printable/printed calendar, day planner or agenda for today or tomorrow, or asks for an email to be saved into the knowledge base.
---

# Outlook (via the `outlook` MCP server)

Requires the `outlook.py` MCP server (read-only; Windows + classic
Outlook running and logged in). If its tools are not available, tell the
user to wire it in first (see the repo README) and to verify with
`python outlook.py --check`.

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

`outlook_print_calendar` writes an A4 landscape PDF into `C:\Eva\documents\pdf`
and reports the path plus what is on the page. It is a bifold: the day on an
hour-by-hour timeline down the left half, the following days summarised on the
right. Tell the user the path, and that it prints single-sided and folds in half.

| They say | Call |
|---|---|
| "Print today's calendar" | `date: "today"` |
| "Tomorrow's planner" | `date: "tomorrow"` |
| "A planner for Friday" | `date: "+N"` if you can count the days, else `YYYY-MM-DD` |
| "Just the day, no-one else's names on it" | `show_attendees: false` |
| "Show me the next week down the side" | `lookahead_days: 6` |

- **One day per sheet.** There is no week or month layout; for a range, print
  each day or use `outlook_get_calendar` and answer in chat instead.
- **Do not re-read the calendar first.** The tool reads it itself and its reply
  lists everything on the page, so a preceding `outlook_get_calendar` is a
  wasted round trip.
- **Re-printing the same day overwrites that day's file.** That is deliberate,
  so say "updated" rather than warning about a clash.
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
KB". The file lands in `C:\Eva\knowledge\email` and the tool reports the path.

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

- Read-only on the mailbox: it cannot send, reply, delete or move mail, or
  accept a meeting — never promise to. The two things it writes are a saved
  email and a printed planner, both only when asked.
- A compliance blacklist may withhold messages/folders entirely; blocked
  items appear only as a withheld count — including on a printed planner,
  where a blocked meeting is left off the page entirely and counted in the
  footer. Do not speculate about their content, and never try to work around
  the filter.
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
