# knowledge\notes\

Markdown **you** write, for the assistant to find later: how a process actually
works, the decision behind a standard, the acronyms your team uses, the notes
from a meeting nobody minuted.

| | |
|---|---|
| **Written by** | you, in any editor |
| **Read by** | the `knowledge-base` plugin's index |
| **Default path** | `C:\Eva\knowledge\notes` |

This is the only folder in [`..`](..) that a plugin never writes to, which makes
it the one place nothing can overwrite. Everything else in `knowledge\` is
regenerated from a source — mirrors get rewritten on every read, and a
re-conversion replaces a converted PDF.

## What works well here

- One topic per file, with a descriptive filename — the filename is cited back
  to you in answers, so `Leave approval - who signs what.md` beats `notes3.md`.
- Real headings (`##`). Chunking is heading-aware, so a well-structured file
  retrieves in useful pieces instead of one undifferentiated block.
- Short is fine. A twenty-line file that answers one question is more useful
  than a hundred-page export.

## After adding a file

Run `kb_index` (or just ask Eva to reindex). Indexing is incremental — only new
and changed files are re-embedded — so it is cheap to run often.
