---
name: confluence
description: Search, read and (when writing is switched on) create and edit Confluence pages - including changing the figures in a table in place, even inside a table filter or column macro - via the confluence MCP server, across one or two Confluence instances, and download a page's attachments or (writing on) upload one. Use when the user asks to find, read, summarise or pull content from Confluence (runbooks, handbooks, wiki pages, spaces), including when they ask for the tasks, action items, statuses, properties or child pages listed on a Confluence page (macro content), when they name a particular Confluence server, when they ask for a Confluence page to be saved into the local knowledge base, when they want a file attached to a page (a spreadsheet, document, deck, PDF or Markdown file) - often to open it with another plugin next - when they want a local file attached to a page or an image on a page replaced (an org chart SVG, say), or when they ask to create a page, update a page, or add content (minutes, actions, a section) to a page.
---

# Confluence (via the `confluence` MCP server)

Requires the `confluence.py` MCP server (Confluence Data Center). It reads by
default; the page-writing tools exist only when the endpoint sets
`CONFLUENCE_ALLOW_WRITE=true`.
If its tools are not available, tell the user to wire it in first (see the
repo README) and to verify connectivity with `python confluence.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `confluence_search` | Free-text search for pages by topic |
| `confluence_search_cql` | Advanced search with raw CQL (spaces, dates, labels) |
| `confluence_get_page` | Full content of one page by numeric ID |
| `confluence_get_page_by_title` | Full content by exact title + space key |
| `confluence_list_pages_under` | Children of a page (navigate a page tree) |
| `confluence_list_attachments` | The files attached to a page, and where each would download to |
| `confluence_download_attachment` | Save one attachment into the folder the matching plugin reads |
| `confluence_list_tables` | The tables on a page (even inside macros): number, heading, columns, rows |
| `confluence_create_page` | New page in a space, optionally under a parent *(writing on)* |
| `confluence_update_table` | Change cells of one table in place, and/or add rows *(writing on)* |
| `confluence_update_section` | Change one section (by heading or panel title), keeping the rest *(writing on)* |
| `confluence_append_to_page` | Add content to the end/start of a page, keeping the rest *(writing on)* |
| `confluence_update_page` | Replace a page's whole body and/or title *(writing on)* |
| `confluence_upload_attachment` | Attach a local file, or replace an attachment with a new version *(writing on)* |

The last six are only in the tool list when writing is switched on. If the
user asks for a page to be created or changed and they are missing, say that
Confluence writing is off on this endpoint and that `CONFLUENCE_ALLOW_WRITE=true`
(in Claude Code's settings `env` block, then a restart) turns it on. Do not
work around it.

## Picking the server

The environment may have **two Confluence instances**. When it does, every tool
takes an optional `server` argument and the tool schema's `server` enum lists
the configured names (e.g. `Green` and `Blue`). No `server` argument in the
schema means only one instance is configured — ignore this section.

1. **Omit `server` by default.** The first server is the default, and an
   ordinary request ("search Confluence for the incident runbook") belongs
   there.
2. **Pass `server` when the user names one.** "find content about X on Blue",
   "check the Blue wiki", "what does Green say about Y" → `server: "Blue"` /
   `server: "Green"`. Matching is case-insensitive.
3. **Query both when the user asks for both** — "check both wikis", "is the
   policy the same on each?", or when a search on the default server comes back
   empty and the answer plausibly lives on the other one. Call the tool once per
   server and say which findings came from which. Do not silently substitute one
   server for the other.
4. **Content IDs are per-instance.** Page `393217` on Green is a different page
   from `393217` on Blue. Read a page from the same server the search that found
   it used — search results carry `server=<name>` on every line for exactly this
   reason.
5. **Always name the server in your answer** when two are configured, so the
   user knows which wiki a fact came from.

If a tool reports an unknown server, it lists the names that *are* configured —
use one of those rather than guessing, and tell the user if the instance they
asked for is not wired in.

## Macros: where a page's real content lives

A page body comes back as Markdown with its macros rendered, so most of this is
automatic. What matters is knowing when content is **missing** rather than
absent.

Confluence macros come in two kinds:

- **Stored in the page** — info/note/warning panels, expand, code, inline task
  lists, status lozenges. These always read. Inline tasks arrive as `- [ ]` and
  `- [x]` checkboxes with the assignee as `@username`, so counting or filtering
  someone's tasks is straightforward.
- **Generated when the page is displayed** — Task Report, Page Properties
  Report, Children Display, Page Tree, Jira Issues, Include Page, Excerpt
  Include. The page source holds only the macro's settings, so these have to be
  fetched by asking Confluence to render the page.

The tools do that on their own. Two things to watch:

1. **Read the `Body:` line** at the top of the output. `storage (page source)`
   plus a note means generated macros were not included; `view` or
   `export_view` means they were.
2. **A placeholder is a retry, not an answer.** Where the body says
   *"[Confluence task report (macro ...). Its content is generated ... not shown
   here]"*, the content exists and was not fetched. Call the same tool again
   with `body_format: "view"`. If it is still a placeholder, try
   `body_format: "export_view"`. Only after both should you tell the user the
   page has nothing on it.

Never report a page as empty, or a person as having no tasks, on the strength of
a placeholder — that is the one failure this argument exists to prevent.

`body_format` also takes `"storage"` (raw source, no macro output) if the user
explicitly wants to see how a page is built rather than what it shows.

## Workflow

1. Start with `confluence_search` using 2–4 topic keywords. Prefer fewer,
   more distinctive words over full sentences.
2. If the user names a space, date range or label, use `confluence_search_cql`
   instead, e.g. `space = DOCS AND text ~ "release notes" AND lastmodified >= now("-30d")`.
3. Fetch the winning result with `confluence_get_page` (by the ID from the
   search results, on the same server) and answer from the page body. Quote the
   page title and ID — plus the server, if there are two — so the user can find
   it.
4. For "everything under X" requests, walk `confluence_list_pages_under`.
5. **"What tasks are on page X?" / "what is assigned to Jane on X?"** → read the
   page and answer from the checkboxes and the task-report rows. If the body
   carries a macro placeholder instead, retry with `body_format: "view"` before
   answering (see [Macros](#macros-where-a-pages-real-content-lives)). Quote the
   task text, who it is assigned to and any due date, and say whether each one
   is ticked.

## Attachments: getting a file off a page

1. `confluence_list_attachments` on the page (by `page_id`, or `title` +
   `space`). Each line gives the name, type, size, version and the local
   folder it would go to, or why it cannot be downloaded.
2. `confluence_download_attachment` with the exact `filename` (or
   `attachment_id`). The file lands where the plugin for its type looks:

   | Type | Folder | Next tool |
   |---|---|---|
   | `.xlsx` / `.xlsm` | `documents\excel` | `excel_list_sheets`, `excel_read_table`, `excel_read_range` |
   | `.docx` | `documents\word` | `msword_open` |
   | `.pptx` | `documents\powerpoint` | `powerpoint_open` |
   | `.pdf` | `documents\pdf` | `convert_pdf_to_markdown` |
   | `.md` | `knowledge\confluence` | `kb_index` |

3. Carry straight on with the other plugin, passing the **file name** the
   download reported (not the full path - each plugin resolves names inside
   its own folder). "Get the budget from page X and total column D" is the
   download followed by an `excel_*` call, in one turn.

Images (`.svg`, `.png`, `.jpg`, `.jpeg`, `.gif`) download to
`documents\images`. No plugin opens them; it is where `/org-chart` writes.

## Attachments: putting a file on a page *(writing on)*

`confluence_upload_attachment` reads ONLY from the folders in the table above
(plus `documents\images`). Pass `file` as the file name - it is looked up in
the folder for its type - or a full path inside one of them. A file anywhere
else is refused: say so, and ask the user to save it into the right folder
rather than copying it there yourself.

**Replacing an image on a page** ("swap the org chart on page Z for this one")
is an upload, not a page edit. The page shows the attachment by name, so a new
version of it appears in place:

1. `confluence_list_attachments` on page Z to find the image's exact name.
   If more than one image could be the one meant, ask which.
2. `confluence_upload_attachment` with `file` (the local name), `attach_as`
   (the name on the page, when it differs) and `overwrite: true`. Set
   `overwrite` only when the user has asked for the replacement - without it the
   tool refuses and names the existing attachment, which is the check you want.
3. Report the new version number. Do not edit the page body: it already points
   at the attachment.

A **new** image is attached but not displayed. To show it, add
`![](<file name.svg>)` (angle brackets when the name has spaces) with
`confluence_update_section` or `confluence_append_to_page`.

- **Never overwrite by default.** If the file already exists the download is
  refused. When the user wants the newest copy ("get the latest version"),
  retry with `overwrite: true`; when they want both, use `save_as`. If it is not
  clear which, ask in one line.
- Other file types are listed but cannot be downloaded - say so rather than
  suggesting a workaround.
- Downloading changes nothing in Confluence, so it works with writing off.

## Writing pages (only when the write tools are present)

Writing to a wiki publishes as the user, so:

1. **Only write what the user asked for.** Never create or edit a page as a
   side effect of research, and never "tidy up" a page you were reading.
2. **Show the content first** - draft it in the chat, get a yes, then write -
   unless the user has already given you the exact text or said to go ahead.
3. **Pick the narrowest tool.** **Figures or text in a table** ("update the Q3
   numbers", "mark Travel as 120", "add a row for Training") →
   `confluence_list_tables`, then `confluence_update_table` - never
   `update_section` or `update_page`, which would rebuild the table and can
   delete the table filter / column macro around it.
   - Take the table number and exact column names from the listing. `match`
     picks rows by a column value (the row label, or an ID column);
     `set` gives the new values. Put every change for one table in ONE call's
     `updates` list, so they succeed or fail together.
   - If it refuses (ambiguous row, merged cell, unknown column), fix that one
     thing from the error; for a cell holding a macro, tell the user what it
     is and ask before passing `allow_macro_removal: true`.
   - A Jira issues table or page properties report is generated by
     Confluence and cannot be edited - say so.

   "Update the X section / X notes / X panel on page Y" → `confluence_update_section` with `section: "X"`: it changes only
   that part and sends the rest of the page back untouched. Adding to the
   bottom of a page (minutes, actions) → `confluence_append_to_page`. Only a
   full rewrite or a title change needs `confluence_update_page`.
   - `update_section` takes `mode`: `replace` (default), `append` (add to the
     end of that section) or `prepend`. A heading's section runs to the next
     heading of the same or higher level, so its sub-headings go with it on a
     replace - if the user only meant the text directly under the heading and
     there are sub-headings, check first.
   - If `section` matches nothing or several, the error lists the page's
     sections; pick the one the user meant, or ask.
   - If it refuses because the section holds a macro, tell the user which one
     and offer `mode: "append"`; pass `allow_macro_removal: true` only once
     they have agreed to lose it.
4. **`confluence_update_page` replaces the whole body.** Read the page first
   (`confluence_get_page` with `body_format: "storage"` shows what is really
   there), build the complete new body from it, and pass the version you read
   as `expected_version`. If the page carries macros (task reports, Jira
   tables, page properties) that your Markdown cannot reproduce, either use
   `content_format: "storage"` and keep the macro XML, or append instead - and
   tell the user. A refused write because the page moved on means read it again,
   never force it.
5. **Say what you did:** title, page ID, new version and URL, as the tool
   reports them.

Bodies are Markdown by default: headings, bold/italic, `code`, links, nested
lists, tables, block quotes, fenced code (becomes a code macro), `- [ ] task`
(becomes a real Confluence task), and `> [!NOTE] Title` / `[!TIP]` /
`[!WARNING]` quotes (become panels). Write action items as `- [ ]` tasks so they
show up in Confluence task reports. Page titles must be unique within a space;
if a create is refused for that reason, ask the user for another title.

## Saving to the knowledge base

Reading a page does **not** save it. `confluence_get_page` and
`confluence_get_page_by_title` take `save_to_kb`, false by default; set it to
true **only when the user asks for the page to be kept** — "save this to the
knowledge base", "add that page to the KB", "pull the runbook in for offline
search". The file lands in `H:\Eva\knowledge\confluence` and the tool reports
the path.

- **Never set it while researching.** Pages you open to answer a question are
  not the user's filing decision, and a knowledge base full of pages nobody
  chose returns irrelevant answers later.
- **Asked after the fact** ("save that page") → call the same tool again for
  that page with `save_to_kb: true`. Re-reading is cheap; guessing is not.
- **Say what you saved** — the title and the path — so the user knows what is
  now in the index. Suggest `kb_index` if they want it searchable immediately.
- If the user says something is worth keeping and you are not sure they meant
  the KB, ask in one line rather than saving.

## Notes

- Without `CONFLUENCE_ALLOW_WRITE=true` the server is read-only and cannot
  create or edit pages or upload attachments; downloading attachments still
  works.
- With two instances configured, saved files are named
  `Confluence <server> - <title>.md`, so pages that share a title on both
  instances stay separate. Saving is off at the server if
  `CONFLUENCE_KB_DIR` is `off` — the tool says so rather than failing quietly.
- Long pages may be truncated in the returned text if `CONFLUENCE_MAX_BODY` is
  set;
  say so if an answer might sit past the truncation point.
- A rendered body is a snapshot of what Confluence showed at the moment of the
  read. A task report reflects the state of tasks right then, so say when you
  read it if the user is acting on the list.
- Saving a page keeps the same body that was read, so a page saved with its
  macro content keeps that content in the knowledge base.
