---
name: jira
description: Query Jira issues via the jira MCP server and, when writing is switched on, create, update, comment on and transition them. Use when the user asks about their tickets, sprint/project status, issue details or history, wants a report drafted from Jira data, or asks to raise/create a ticket, change a ticket's fields or assignee, comment on a ticket, or move a ticket to another status.
---

# Jira (via the `jira` MCP server)

Requires the `jira.py` MCP server (Jira Data Center v2 REST API). It reads
by default; the write tools exist only when the endpoint sets
`JIRA_ALLOW_WRITE=true`. If its tools are not available, tell the user to wire it in first
(see the repo README) and to verify with `python jira.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `jira_my_issues` | "What's assigned to me?" — the user's open issues |
| `jira_search` | Free-text search (safely quoted into JQL) |
| `jira_search_jql` | Advanced search with raw JQL |
| `jira_get_issue` | One issue in full (set `include_changelog=true` for history) |
| `jira_project_status` | Health summary of one project (counts by status, unassigned, top open) |
| `jira_list_projects` | Which project keys are visible |
| `jira_list_transitions` | An issue's status and the moves available from it |
| `jira_list_versions` | A project's releases (fix versions), to find the exact name |
| `jira_create_issue` | Raise a new issue *(writing on)* |
| `jira_update_issue` | Change fields, labels, fix versions, assignee; optional comment *(writing on)* |
| `jira_add_comment` | Comment on an issue *(writing on)* |
| `jira_transition_issue` | Move an issue to another status *(writing on)* |

The last four are only in the tool list when writing is switched on. If the
user asks for a change and they are missing, say that Jira writing is off on
this endpoint and that `JIRA_ALLOW_WRITE=true` (in Claude Code's settings `env`
block, then a restart) turns it on. Do not work around it.

## Workflow

1. "My work" questions → `jira_my_issues`, sort/summarise by priority.
2. "Has anyone seen…" / topic questions → `jira_search` with distinctive
   keywords from the problem description.
3. Time-boxed or precise questions → `jira_search_jql`, e.g.
   `project = ABC AND resolved >= -7d ORDER BY resolved DESC`.
4. Deep-dive on one ticket → `jira_get_issue` (add the changelog only when
   the user asks who changed what).
5. Reports: combine `jira_my_issues`/`jira_search_jql` results, then (if the
   word server is available) draft the report as a .docx with tracked
   changes.

## Changing Jira (only when the write tools are present)

Anything written to Jira is posted as the user and seen by their team, so:

1. **Only change what the user asked for**, and never as a side effect of
   reading. Do not "fix up" a ticket you were only asked to summarise.
2. **Confirm the words first.** Draft a new issue's summary and description,
   or a comment, in the chat and get a yes - unless the user gave you the exact
   text or said to go ahead.
3. **Write in Jira wiki markup, not Markdown:** `h2. Heading`, `*bold*`,
   `_italic_`, `{{code}}`, `* bullet`, `# numbered`, `[label|https://url]`,
   `||head||head||` then `|cell|cell|`, `{code}...{code}`, `[~username]`.
4. **Creating:** `project`, `issue_type` (as the project names it - "Task",
   "Bug", "Story", "Sub-task") and `summary` are required; a sub-task also
   needs `parent`. If Jira refuses, its reply names the missing or invalid
   field; fix that one thing (a custom field goes in `fields` by id) or ask the
   user for the value. Never invent a value for a required field.
5. **Assigning:** pass the person as the user said it. If the tool reports
   several matches, show the candidates and ask; never pick one.
6. **Moving:** `jira_transition_issue` accepts the target status ("Done"). If
   it lists the available transitions instead, choose only if exactly one
   plainly fits the request, otherwise ask. A transition that needs a
   resolution takes `resolution`.
7. **Labels:** `add_labels` / `remove_labels` keep the others; `labels`
   replaces the whole set - use it only when the user wants that.
8. **Fix versions (releases):** look the names up with `jira_list_versions`
   first - Jira only accepts a release that already exists in the issue's
   project, spelled exactly. `add_fix_versions` / `remove_fix_versions` keep
   the others; `fix_versions` replaces the whole set. When matching by a
   "similar" name (e.g. to an issue's "delivers" links), show the user the
   proposed issue → release pairs, and any name with no clear match, before
   changing anything. Never pick between two plausible releases; ask.
9. **Report back** with the issue key and URL the tool returns.

## Notes

- Without `JIRA_ALLOW_WRITE=true` the server cannot create, edit, transition
  or comment on issues — never promise to update Jira then. Nothing can
  delete an issue either way.
- **Nothing from Jira is saved anywhere.** The server writes no files, so
  issues you read never reach the knowledge base. If the user asks for
  something from Jira to be kept ("save this ticket", "remember this
  decision"), write it as a note with the `knowledge-base` server's
  `kb_capture` — quoting the issue keys — and only when they ask.
- A `JIRA_PROJECTS` allowlist may be configured; issues outside it are
  refused and other projects hidden. If a key is refused, say why.
- Jira Cloud is not supported (v3 rich-text API); this targets Data Center.
