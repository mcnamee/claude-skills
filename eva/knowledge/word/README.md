# knowledge\word\

Word documents, mirrored to Markdown — headings, lists and tables preserved as
text. The `word` plugin writes a file here whenever it **opens** a document and
whenever it **saves** one.

| | |
|---|---|
| **Plugin setting** | `word` → knowledge-base folder (`--kb-dir` / `MSWORD_KB_DIR`) |
| **Default** | `C:\Eva\knowledge\word` |
| **Filenames** | `Word - <name>.md` |
| **Overwritten** | yes, on every open or save |

Written by the server only.

## Mirroring on save is the useful half

Mirroring on *open* makes the documents in [`..\..\documents\word`](../../documents/word)
searchable. Mirroring on *save* is what captures documents Eva **wrote** — a
report drafted this morning is retrievable this afternoon without anyone
reopening it.

That does mean a document saved repeatedly leaves one file here reflecting the
last save. It is a current-state mirror, not a version history; Word's own
tracked changes are where revision history lives.

## Why the `.docx` files are not in here

This folder holds text extracted for the index. The documents themselves live
in [`..\..\documents\word`](../../documents/word) (source) and
[`..\..\output\word`](../../output/word) (generated) — a `.docx` dropped in here
would simply be ignored by the indexer.
