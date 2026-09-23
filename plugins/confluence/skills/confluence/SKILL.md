---
name: confluence
description: Search and read Confluence pages via the confluence MCP server, across one or two Confluence instances. Use when the user asks to find, read, summarise or pull content from Confluence (runbooks, handbooks, wiki pages, spaces), including when they ask for the tasks, action items, statuses, properties or child pages listed on a Confluence page (macro content), when they name a particular Confluence server, or when they ask for a Confluence page to be saved into the local knowledge base.
---

# Confluence (via the `confluence` MCP server)

Requires the `confluence.py` MCP server (read-only, Confluence Data Center).
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

- The server is read-only; it cannot create or edit pages.
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
