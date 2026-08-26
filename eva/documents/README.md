# documents\

The binary library — the `.docx`, `.xlsx` and `.pdf` files themselves. This is
what the plugins **read**; the text they extract for searching goes to
[`..\knowledge`](../knowledge), and anything they create goes to
[`..\output`](../output).

| Folder | Plugin | Setting | Sub-folders? |
|---|---|---|---|
| [`word\`](word) | `word` | `--docs-dir` / `MSWORD_DOCS_DIR` | yes — searched recursively |
| [`excel\`](excel) | `excel` | `--docs-dir` / `EXCEL_DOCS_DIR` | **no** — top level only |
| [`pdf\`](pdf) | `pdf-to-md` | `--docs-dir` / `PDF2MD_DOCS_DIR` | only with the recursive option |

Each of these folders is a **sandbox**: the plugin resolves every path against
it (symlinks included) and refuses anything that lands outside. `word` can
therefore open and save inside `documents\word\`, `..\output\word\` and
`..\reference\templates\`, and nowhere else on the machine.

## Where files you are working on go

There is no separate `input\` folder, and that is deliberate. A plugin can only
open files inside its own folder, and each takes exactly one — so a top-level
`input\` would have to *be* that folder, which would lock the plugin out of
everything you keep.

Instead, split by lifecycle **inside** the plugin's folder:
[`word\inbox\`](word/inbox) for what you are working on now,
[`word\library\`](word/library) for what you keep. `word` searches recursively
and matches on filename, so `"open Contract v3.docx"` finds the file wherever it
sits — and `inbox\` stays a folder you can empty without thinking.

`excel` is the exception: it lists only the top level of its folder, so
workbooks must sit directly in [`excel\`](excel).

## Nothing here is indexed

The RAG index reads text, not Office formats. A `.docx` becomes searchable when
`word` opens it (mirroring it to `knowledge\word\`), and a PDF when `pdf-to-md`
converts it. Dropping files here makes them *available*; opening or converting
them makes them *searchable*.
