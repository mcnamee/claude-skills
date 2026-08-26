# reference\

Material about **how to write**, not about what is true. Two folders, doing two
different jobs:

| Folder | Holds | Used for |
|---|---|---|
| [`exemplars\`](exemplars) | Finished, *good* documents — `.md`, `.docx`, `.pptx`, `.pdf` | Showing Claude **what good looks like** — tone, structure, level of detail — before it writes something of the same kind |
| [`templates\`](templates) | Blank/skeleton `.docx` and `.pptx` | The **starting file** a new document is built from — letterhead, styles, headers/footers, boilerplate |

The distinction matters: an exemplar is *read* for guidance and never becomes the
output; a template *is* the output's first draft. A finished board paper belongs
in `exemplars\`; the empty branded shell it was written in belongs in
`templates\`.

Both folders carry their own README explaining what to put in them, how to name
it and how to ask for it. Start there.

## Why this sits outside `knowledge\`

Nothing here is indexed, deliberately. Add a board paper to the RAG corpus and
its phrasing comes back with the same authority as a policy — `kb_ask` starts
citing house style as fact, and a two-year-old example as current practice.
Keeping exemplars out of [`..\knowledge`](../knowledge) means Eva reads them when
you point at one, and never quotes them at you by accident.

If you want a document to be *both* — a genuine reference as well as a model to
write like — put its content in `knowledge\notes\` and keep the formatted file
here. Same document, two jobs, no confusion about which is being cited.

## Wiring it up

Only `templates\` needs configuring, because a server has to open those files:
the `word` plugin's templates folder (`--templates-dir` /
`MSWORD_TEMPLATES_DIR`) defaults to `C:\Eva\reference\templates`, and treats it
as a **read-only** third root — templates can be listed, opened and used as the
base for a new document, but every save into the folder is refused. Full setup in
[`templates/README.md`](templates/README.md).

`exemplars\` needs no configuration: the files are read by whatever tool suits
the format — `.md` directly, `.docx` via the `word` plugin, `.pdf` via
`pdf-to-md`.

## Why `README.md` and not `CLAUDE.md`

`CLAUDE.md` files are auto-loaded instructions for an agent working *in this
repository*. These two folders are consumed on your endpoint, by an assistant
doing your work — so what is written here is documentation for you, not
instructions for a coding agent.
