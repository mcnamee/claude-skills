# knowledge\

The RAG corpus — **the one folder the vector index reads**. Everything in here
is Markdown or plain text, gets split into chunks, embedded and searched by
`kb_ask`, `kb_retrieve` and `kb_index`.

| | |
|---|---|
| **Plugin setting** | `knowledge-base` → documents folder (`--docs-dir` / `KB_DOCS_DIR`) |
| **Default** | `C:\Eva\knowledge` |
| **Indexed extensions** | `.md`, `.markdown`, `.txt` — recursively, all sub-folders |
| **Skipped** | anything else, plus files and folders whose name starts with `.` |

## Sub-folders are provenance, not topic

Each sub-folder is named after **what writes into it**, and maps to exactly one
plugin setting:

| Folder | Filled by | Setting |
|---|---|---|
| [`notes\`](notes) | you, by hand | — |
| [`captures\`](captures) | `kb_capture` | `knowledge-base` → output folder |
| [`confluence\`](confluence) | the `confluence` plugin, on every page read | `--kb-dir` |
| [`email\`](email) | the `outlook` plugin, on every email read | `--kb-dir` |
| [`word\`](word) | the `word` plugin, on every document opened or saved | `--kb-dir` |
| [`pdf\`](pdf) | the `pdf-to-md` plugin, on conversion | `--output-dir` |

Add your own sub-folders freely — the index walks the whole tree, so nesting
costs nothing. Keep naming them after their source.

## Two things to know

**A mirror outside this folder is invisible.** If a plugin's knowledge-base
folder points anywhere that is not inside `knowledge\`, it will mirror perfectly
and the index will never read a line of it. Every default in this repo is set
so that cannot happen; if you move one, move it to another folder in here.

**Captured notes sit beside real documents.** A research brief written by an
agent is indexed with the same authority as a policy you mirrored from
Confluence. Retrieval cannot tell them apart, so anything captured carries a
front-matter stamp saying it was agent-written — leave it in place, and treat
`captures\` as a lead, never a source.

## Binary files do not belong here

`.docx`, `.xlsx` and `.pdf` are not indexable text. They live in
[`..\documents`](../documents); their *converted* Markdown is what lands here.
