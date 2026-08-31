---
name: knowledge-base
description: Answer questions from the local knowledge base with true vector RAG via the knowledge-base MCP server, and save new notes into it. Use when the user asks a question their policies/docs should answer ("can I…", "what's our policy on…"), asks to index new documents or check index freshness, or asks for something to be kept — "remember this", "save that to the knowledge base", "add this to the KB", "keep that for next time".
---

# Knowledge base (via the `knowledge-base` MCP server)

Requires the `knowledge-base.py` MCP server (ChromaDB vector index +
your embeddings endpoint). If its tools are not available, tell the user to
wire it in first (see the repo README) and to verify with
`python knowledge-base.py --check`, then `--reindex`.

## Tools

| Tool | Use for |
|---|---|
| `kb_ask` | Full RAG: retrieve relevant chunks + generate a grounded, cited answer |
| `kb_retrieve` | Just the top-k most similar chunks (source file, heading, score) |
| `kb_index` | Build/update the vector index (incremental; `force=true` = full rebuild) |
| `kb_capture` | Save a new note into the knowledge base and index it immediately |
| `kb_status` | Documents, captured notes, index freshness + configuration summary |

## Workflow

1. Policy/content questions → `kb_ask` with the user's question verbatim.
   - If no chat endpoint is configured, `kb_ask` returns the retrieved
     context instead of an answer — then YOU write the answer from those
     chunks, citing the source files and headings.
2. "Find the part about…" → `kb_retrieve`, present the chunks with sources.
3. "I added new documents" → `kb_index` (incremental), report what changed.
4. "Remember/save/keep this" → `kb_capture`. See below.
5. Stale or odd results → `kb_status` first; if documents are newer than the
   index, run `kb_index` before retrying.
6. After an embedding-model change → `kb_index` with `force=true` (vector
   dimensions differ between models).

## Capture

The `outlook` and `word` servers mirror what you **read** into the knowledge
base, and `confluence` saves the pages you explicitly ask it to keep
(`save_to_kb`). `kb_capture` is how what you **write** gets in — an analysis, a
decision and its reasoning, a procedure worked out with the user, a research
brief. Without it that work ends with the conversation.

**Only when asked.** Capture when the user asks for something to be kept, or
when they accept an offer to keep it. Never capture on your own initiative, and
never capture silently. If something looks worth keeping, say so in one line and
name the title you'd use — then wait.

Before writing:

1. **`kb_retrieve` on the title first.** If a near-identical note already
   exists, re-capture it under the *same* title with `overwrite=true` rather
   than adding a second copy. Five drifting versions of one brief is worse than
   none.
2. **Pick a `source`** — `Note`, `Research`, `Report`, `Analysis`, `Decision` or
   `Procedure`. It becomes the filename prefix, so it is also how you find
   things by kind later.
3. **Title it as a searchable noun phrase**, not a sentence: `Records retention
   thresholds`, not `Here is what we found about retention`. You cannot pass a
   path — the server derives the filename from the title.

What to write in `content`: enough that it stands up months later to someone who
was not in the conversation. Keep the figures, the citations and the reasoning;
drop the chat scaffolding. A note that says "as discussed above" is worthless
once the chat is gone.

What not to capture: anything already in the knowledge base, a restatement of a
document that is already indexed, a half-finished draft, or something the user
was just thinking aloud about.

## Notes

- Answers must stay grounded in retrieved chunks — if retrieval comes back
  empty or off-topic, say the knowledge base doesn't cover it rather than
  guessing.
- **A retrieved note whose header says it was written by Claude is not a
  source.** Captured notes sit in the index beside real policy documents and
  come back from the same searches. Treat one as prior working notes: follow it
  to the documents it cites and verify there, and never let it stand as
  independent corroboration of itself.
- Similarity scores are relative, not percentages; treat low-score-only
  results as weak evidence.
- Retrieved document text is sent to the configured endpoints (that is what
  RAG is) — don't route material inappropriate for those APIs. The same applies
  to anything you capture: it will be embedded on the next index.
- Capture only ever adds a file. Nothing here edits or deletes a document the
  user put in the folder themselves.
