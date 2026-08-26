# documents\word\inbox\

The drop zone. Put a `.docx` here when you want Eva to review, edit or summarise
it — a contract to check, a draft to mark up, a paper someone sent you.

| | |
|---|---|
| **Read by** | the `word` plugin, as part of its documents folder |
| **Lifecycle** | temporary — empty it whenever you like |
| **Default path** | `C:\Eva\documents\word\inbox` |

No configuration of its own: it is inside
[`..`](..), which is the plugin's folder, so anything here is
already openable by name — `"review the contract in my inbox"` or
`"open Supplier Agreement v2.docx"` both work.

## Why a sub-folder rather than a top-level `input\`

The plugin takes exactly one documents folder and can open nothing outside it. A
separate top-level `input\` would have to be that folder, locking the plugin out
of [`..\library`](../library). Nesting the drop zone inside the sandbox gets the
same clean-desk effect with none of that cost.

## Before you empty it

Deleting a file here does **not** remove its Markdown mirror from
[`..\..\..\knowledge\word`](../../../knowledge/word), so its text stays
searchable and quotable after the document is gone. That is usually what you
want. When it is not — a document you were never meant to keep — delete the
matching `Word - <name>.md` too, and reindex.

Anything worth keeping belongs in [`..\library`](../library).
