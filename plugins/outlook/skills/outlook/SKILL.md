---
name: outlook
description: Read Outlook mail and calendar via the outlook MCP server (read-only COM automation, Windows). Use when the user asks about their emails, wants mail searched or summarised, asks what's on their calendar, or asks for an email to be saved into the knowledge base.
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
| `outlook_list_folders` | Real folder names across all stores |

## Workflow

1. List/search first (`outlook_list_recent_emails` / `outlook_search_emails`
   / `outlook_search_recent`), then `outlook_get_email` with the EntryID of
   the message the user cares about — bodies are only available via get.
2. If a folder-scoped search misses, `outlook_list_folders` to learn the
   actual folder names, then retry `outlook_search_recent` with `folders`.
3. "What did I do last week?" → `outlook_list_sent_emails` + summarise.

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

- Read-only: it cannot send, reply, delete or move mail — never promise to.
- A compliance blacklist may withhold messages/folders entirely; blocked
  items appear only as a withheld count. Do not speculate about their
  content, and never try to work around the filter.
- Calendar date filtering is done in Python (locale-independent), so
  regional date settings cannot empty the results. If `outlook_get_calendar`
  finds nothing, its reply includes a `[debug]` section listing the last few
  calendar items scanned (start date + in-range flag) — read it to tell an
  actually-empty window apart from a filtering fault before retrying.
