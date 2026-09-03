# index\

The ChromaDB vector store — the embeddings for everything in
[`..\knowledge`](../knowledge), plus the bookkeeping that makes indexing
incremental.

| | |
|---|---|
| **Setting** | `KB_INDEX_DIR` (`knowledge-base`; the one folder outside the four shared roots) |
| **Default** | `C:\Eva\index` |
| **Contents** | a SQLite database and binary vector segments — not human-readable |
| **Size** | grows with the corpus; hundreds of MB is normal |

## Why it sits outside `knowledge\`

The plugin's own fallback, if this setting is left empty, is a hidden
`.kb-rag-index` folder *inside* the documents folder. That works — the indexer
skips dot-folders — but it buries a large binary database inside the corpus you
want to be able to zip, copy, grep or diff. Keeping it out here means
`knowledge\` stays nothing but text.

## Disposable, and worth deleting sometimes

Everything here is derived from `knowledge\`. Delete the folder and run
`kb_index` and you get it back — at the cost of re-embedding every chunk, so a
large corpus means a long run and a lot of calls to your embeddings endpoint.

Rebuild from scratch when you change the **embedding model**, the chunk size or
the chunk overlap. Vectors from different models are not comparable, and a
half-migrated index returns confidently wrong matches rather than failing
outright.

Never edit anything in here by hand, and do not back it up in place of
`knowledge\` — the corpus is the source of truth, this is a cache.
