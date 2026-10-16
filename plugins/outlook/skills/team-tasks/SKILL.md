---
name: team-tasks
description: Print the user's team's open tasks as a double-sided A4 landscape PDF (eight columns, four a side - unassigned and whole-team tasks first, then one column per direct report), gathered from the Confluence page that collates their task lists and the emails the user sent them that have a due date in their Inbox. Use when the user asks for a printable/printed team task list, "print my team's tasks", "what's everyone working on - print it", a task sheet for their direct reports, or to refresh last week's task sheet.
---

# Team task sheet (via the `outlook` MCP server, with `confluence`)

Needs the `outlook` server v11.3.0 or later (`outlook_get_task_team`,
`outlook_print_team_tasks`) and, for the Confluence half, the `confluence`
plugin. If the outlook tools are missing, tell the user to update the plugin.
If the confluence tools are missing, say so, and offer to print from email
alone rather than stopping.

The server lays out the page; **you gather the tasks**. The order matters:

## 1. Get the set-up

Call `outlook_get_task_team` first. It returns:

- the **columns in print order**: column 1 is always "Team / unassigned",
  then the direct reports from `OUTLOOK_TASKS_TEAM`;
- the **Confluence page** (`OUTLOOK_TASKS_PAGE`) - if it says "not set", ask
  the user which page collates the tasks;
- **today's date** on the endpoint - use it for "overdue", not your own idea
  of today.

If it says `OUTLOOK_TASKS_TEAM` is not set, ask the user who their direct
reports are and use those names consistently, then suggest setting the
variable so the columns stay in the same order every time. If it says
printing is off, stop and tell the user before gathering anything.

## 2. Read the Confluence page

- A **numeric ID or URL** → `confluence_get_page` with the ID. A **title** →
  `confluence_get_page_by_title` (ask for the space key if you do not have
  it), or `confluence_search` to find it.
- A collated page is usually a **task report macro** or tables. If the page
  reads as having no tasks where the user expects some, retry with
  `body_format: "view"`, then `"export_view"`, before saying it is empty.
- Take **open tasks only**: skip `- [x]`, "Done", "Complete", "Closed",
  struck-through rows. Keep the assignee, the due date and any status.
- Do **not** pass `save_to_kb`.

## 3. Read the emailed tasks

How the user works: an Outlook rule copies every email they send back into
their **Inbox**. There they either **delete** it (no reminder needed) or
**give it a due date** (a task they want chased). So the emailed tasks are
exactly the ones with a due date - the user has already decided.

- Call `outlook_list_task_emails`. It returns the emails the user sent that are
  in the Inbox with a due date and not marked complete, soonest first, with
  To / Cc, subject, a short preview and EntryID.
- **Every one is a task.** Do not second-guess which ones count, and do not
  add tasks from mail without a due date.
- **Never use `outlook_list_sent_emails` or `outlook_search_recent`** for this:
  Sent Items is not the list, and an Inbox email with no due date is one the
  user has not chosen to track.
- The **due date is the email's**. Use it as given.
- Word the task as a short action ("Send revised costing to Finance"), not the
  subject line. Open the email with `outlook_get_email` only when the subject
  and preview do not make the ask clear.
- Withheld messages stay withheld: never guess what they asked for.

## 4. Build one list

- **Owner**: one direct report → their name exactly as the column list shows
  it. The whole team, two or more people, or nobody → `person: ""` (column 1).
  An email to one person with others cc'd belongs to the person in To.
- **Merge** a task that is on Confluence and in an emailed task into one
  entry: Confluence's wording, `source: "Confluence"`, and the email's due
  date if the page has none.
- **Due** as `YYYY-MM-DD`: the email's due date, the page's date for
  Confluence. Leave it out when there is none - never invent one.
- **source**: `"Confluence"` or `"Email 9 Oct"` (the sent date). **status**
  only when the page or email gives one. **detail** only for a few words that
  change what the task means.
- **sources** (footer): one line on what you read, e.g.
  `"Confluence: Team Task List · emails with a due date"`.

## 5. Print

Call `outlook_print_team_tasks` with `tasks` and `sources`. Then tell the user,
briefly:

- the **path** (`H:\Eva\documents\pdf\Team Tasks - <date>.pdf`);
- **print double-sided, flip on short edge, at 100%** - columns 1-4 on the
  front, 5-8 on the back (Adobe Reader picks this up from the PDF; other
  viewers need it set by hand) - take the exact wording from the reply;
- the per-person counts in one line, and anything the reply flags:
  - a name that **matched no column** (it printed in column 1 as "For
    <name>") - fix the name and re-print if it is a direct report;
  - a column that **ran out of room** - offer to shorten the wording or drop
    low-value items and re-print.

Re-printing the same day overwrites that day's sheet; say "updated".

## Notes

- **Do not print the list into the chat first** unless the user asks to see it.
  The sheet is the deliverable; one line of counts is the summary.
- **Never invent tasks, owners or dates** to fill a quiet column. An empty
  column prints "Nothing open", which is a true answer.
- If you are unsure who owns an emailed task, put it in column 1 rather than
  guessing a person.
- Nothing here sends, changes Confluence, or marks anything complete. When the
  user says a task is done, the fix is theirs: mark the email complete in
  Outlook (or the row done in Confluence) and it drops off the next sheet.
