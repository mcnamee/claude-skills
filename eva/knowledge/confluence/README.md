# knowledge\confluence\

Confluence pages you asked to keep, saved as Markdown. The `confluence` plugin
writes a file here **only when a request says to save the page** — reading a
page to answer a question saves nothing, so pages skimmed on the way to an
answer never reach the index.

| | |
|---|---|
| **Plugin setting** | `confluence` → knowledge-base folder (`--kb-dir` / `CONFLUENCE_KB_DIR`) |
| **Default** | `C:\Eva\knowledge\confluence` |
| **Filenames** | `Confluence - <title>.md`, or `Confluence <server> - <title>.md` with two instances configured |
| **Overwritten** | yes, whenever the same page is saved again — the newest save wins |

Written by the server only; the model never chooses a path in here.

## Say when you want a page kept

"Search Confluence for the retention policy" fills nothing in here. "Save the
retention policy page to the knowledge base" writes one file. That is
deliberate: a folder of pages you chose answers questions better than a folder
of everything that was ever opened.

Set `CONFLUENCE_KB_AUTOSAVE=true` to go back to saving every page read (how the
plugin behaved before v3.0.0) — worth a deliberate decision, because it is the
fastest way to fill this folder with pages nobody asked for.

## Safe to delete

Everything in this folder is derived. If a space is restructured, or pages you
saved months ago are now wrong, delete the files (or the whole folder), read
the pages again with saving on and reindex. Nothing you authored is at risk —
that is the point of keeping copies in their own folder rather than mixed in
with [`..\notes`](../notes).

## A saved page is not a sync

A file here is a snapshot of the page **as read**. Confluence moving on does not
update it, and a page deleted upstream leaves its copy behind. For anything
where currency matters, re-read the page rather than trusting the copy; treat
old files here the way you would a printout.
