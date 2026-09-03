# knowledge\captures\

Where `kb_capture` writes notes **back** into the knowledge base — a research
brief, a finished report, a decision recorded during a conversation. The rest of
the suite mirrors what you *read*; this is how what you *produce* survives the
chat.

| | |
|---|---|
| **Setting** | `EVA_KNOWLEDGE_DIR` (the `knowledge-base` plugin appends `\captures`) |
| **Default** | `C:\Eva\knowledge\captures` |
| **Filenames** | `<source> - <title>.md`, e.g. `Research - Licence renewal.md` |
| **Must sit** | inside [`..`](..) — a capture folder outside the indexed root would never be searchable |

Captures are indexed immediately, not on the next `kb_index` run.

## Treat these as leads, not sources

A captured note is indexed beside genuine policy documents and retrieval cannot
tell them apart. Each one is stamped in its front matter as agent-written; that
stamp is what stops a March guess being cited as fact in June. Two habits keep
this folder honest:

- **Trust the citations, not the note.** A captured brief is a summary of the
  documents it cites. Follow those before relying on it.
- **Re-capture, do not duplicate.** Updating something already captured means
  the same title with `overwrite=true`, so you get one current note rather than
  three versions of a stale one.

## Capture is additive, always

Nothing here ever edits or deletes a document elsewhere in `knowledge\`.
Capturing only ever adds a file, so the worst case is a note you delete by hand.
