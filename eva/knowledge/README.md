# knowledge\

The RAG corpus — **the one folder the vector index reads**. Everything in here
is Markdown or plain text, gets split into chunks, embedded and searched by
`kb_ask`, `kb_retrieve` and `kb_index`.

| | |
|---|---|
| **Setting** | `EVA_KNOWLEDGE_DIR` — one environment variable, shared by every plugin |
| **Default** | `C:\Eva\knowledge` |
| **Must exist?** | **yes**, along with the sub-folder of each plugin you use |
| **Indexed extensions** | `.md`, `.markdown`, `.txt` — recursively, all sub-folders |
| **Skipped** | anything else, plus files and folders whose name starts with `.` |

`knowledge-base` indexes this **whole root**; every other plugin writes into
its own sub-folder of it, named after the plugin. That is the entire wiring:
set `EVA_KNOWLEDGE_DIR` once and a page you save, an email you keep, a document
you open and a PDF you convert all land somewhere the index reads.

## Sub-folders are provenance, not topic

Each sub-folder is named after **what writes into it**, and maps to exactly one
plugin setting:

| Folder | Filled by |
|---|---|
| [`notes\`](notes) | you, by hand |
| [`captures\`](captures) | `kb_capture`, from the `knowledge-base` plugin |
| [`confluence\`](confluence) | the `confluence` plugin, on the pages you ask it to save |
| [`email\`](email) | the `outlook` plugin, on the emails you ask it to save |
| [`word\`](word) | the `word` plugin, on every document opened or saved |
| [`powerpoint\`](powerpoint) | the `powerpoint` plugin, on every deck opened or saved |
| [`pdf\`](pdf) | the `pdf-to-md` plugin, on conversion |

Add your own sub-folders freely — the index walks the whole tree, so nesting
costs nothing. Keep naming them after their source.

## Two things to know

**A mirror outside this folder is invisible.** If a plugin's knowledge folder
points anywhere that is not inside `knowledge\`, it will mirror perfectly and
the index will never read a line of it. Deriving every one of them from
`EVA_KNOWLEDGE_DIR` is what makes that impossible by default; if you override a
single plugin's folder (`OUTLOOK_KB_DIR` and friends), keep it in here.

**Captured notes sit beside real documents.** A research brief written by an
agent is indexed with the same authority as a policy you saved from
Confluence. Retrieval cannot tell them apart, so anything captured carries a
front-matter stamp saying it was agent-written — leave it in place, and treat
`captures\` as a lead, never a source.

## Binary files do not belong here

`.docx`, `.xlsx` and `.pdf` are not indexable text. They live in
[`..\documents`](../documents); their *converted* Markdown is what lands here.
