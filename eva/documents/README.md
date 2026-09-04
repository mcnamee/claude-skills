# documents\

The binary library — the `.docx`, `.pptx`, `.xlsx` and `.pdf` files themselves.
One folder per file type, and each is the **only** folder its plugin touches:
what you put there and what Eva writes live side by side. The text extracted for
searching goes to [`..\knowledge`](../knowledge).

One environment variable, `EVA_DOCUMENTS_DIR`, points at this folder; each
plugin appends its own name, so there is no per-plugin folder setting.

| Folder | Plugin | Sub-folders? |
|---|---|---|
| [`word\`](word) | `word` | yes — searched recursively |
| [`powerpoint\`](powerpoint) | `powerpoint` | yes — searched recursively |
| [`excel\`](excel) | `excel` | **no** — top level only |
| [`pdf\`](pdf) | `pdf-to-md` | only with `PDF2MD_RECURSIVE=1` |

Each of these **must exist** for its plugin to start (`word`, `powerpoint` and
`excel` refuse to run without their folder; `pdf-to-md` reports it on
`--check`).

Each of these folders is a **sandbox**: the plugin resolves every path against
it (symlinks included) and refuses anything that lands outside. `word` can
therefore open and save inside `documents\word\` and read
[`..\templates\word`](../templates/word), and nowhere else on the machine.

## One folder, not three

There is no `input\`, no `output\` and no `inbox\` vs `library\` split. Each
plugin works in exactly **one** folder and can open nothing outside it, so every
extra folder was either unreachable or another setting to keep in sync — and it
made you decide where a file belonged before you could ask a question about it.

So: drop a file in, and ask. A document Eva creates lands in the same folder,
because a draft you are working on and a document you were sent are the same
thing to everyone except a filing system.

Organise inside the folder however suits you — `word` and `powerpoint` search
recursively and match on filename, so `"open Contract v3.docx"` finds the file
whether it sits at the top or in `Contracts\2026\`. `excel` is the exception:
it lists only the top level of its folder, so workbooks must sit directly in
[`excel\`](excel).

## Nothing here is indexed

The RAG index reads text, not Office formats. A `.docx` becomes searchable when
`word` opens or saves it (mirroring it to `knowledge\word\`), a `.pptx` the same
way via `powerpoint`, and a PDF when `pdf-to-md` converts it. Dropping files here
makes them *available*; opening, saving or converting them makes them
*searchable*.
