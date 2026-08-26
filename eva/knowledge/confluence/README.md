# knowledge\confluence\

Confluence pages, mirrored to Markdown. The `confluence` plugin writes a file
here **every time it reads a page**, so the knowledge base fills up as a
by-product of ordinary use — no bulk export step.

| | |
|---|---|
| **Plugin setting** | `confluence` → knowledge-base folder (`--kb-dir` / `CONFLUENCE_KB_DIR`) |
| **Default** | `C:\Eva\knowledge\confluence` |
| **Filenames** | `Confluence - <title>.md`, or `Confluence <server> - <title>.md` with two instances configured |
| **Overwritten** | yes, on every re-read — the newest read wins |

Written by the server only; the model never chooses a path in here.

## Safe to delete

Everything in this folder is derived. If a space is restructured, or pages you
mirrored months ago are now wrong, delete the files (or the whole folder), read
the pages again and reindex. Nothing you authored is at risk — that is the
point of keeping mirrors in their own folder rather than mixed in with
[`..\notes`](../notes).

## Mirroring is not a sync

A file here is a snapshot of the page **as read**. Confluence moving on does not
update it, and a page deleted upstream leaves its mirror behind. For anything
where currency matters, re-read the page rather than trusting the mirror; treat
old files here the way you would a printout.
