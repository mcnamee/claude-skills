# reference\templates\

Blank, branded starting files - letterhead, report layout, contract
boilerplate, the deck shell with your title slide. A new document is **created
from** one of these, inheriting its styles, headers/footers, page setup and
boilerplate; the template itself is never touched.

If what you have is a *finished* document showing what good output looks like,
that is an exemplar - put it in [`../exemplars`](../exemplars) instead.

## One folder per plugin

This folder is the **template root**, and each plugin works in its own
sub-folder of it - exactly as they do under `documents\` and `knowledge\`:

| Folder | Holds | Read by |
|---|---|---|
| [`word\`](word) | `.docx` blanks | the `word` plugin |
| [`powerpoint\`](powerpoint) | `.pptx` / `.potx` deck shells | the `powerpoint` plugin |

| | |
|---|---|
| **Setting** | `EVA_TEMPLATES_DIR` - one environment variable for the root; each plugin appends its own sub-folder name |
| **Default** | `C:\Eva\reference\templates` |
| **Must exist?** | **yes** - the sub-folder for each plugin you use. A missing one is not fatal; the plugin simply starts with templates disabled and says so on stderr |
| **Writable?** | **no** - every save into this tree is refused |

A sub-folder per plugin rather than one shared folder means a listing only ever
shows files the asking plugin can actually use, and it matches the rest of the
tree: one variable for the root, the plugin's own name for the leaf.

## Wiring it up

Set the root once for your Windows account, and both plugins find their
sub-folder:

```powershell
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\reference\templates", "User")
```

Confirm it took: the `word` server logs `templates folder (read-only) = ...` at
startup, and `msword_list_documents` with `location: "templates"` lists exactly
what is in `word\`. The `powerpoint` equivalent is
`powerpoint_list_presentations` with `location: "templates"`.

> The template root must be **separate from** the documents root - the server
> refuses to start otherwise, rather than silently blocking every save.

## Why this sits outside `knowledge\`

Nothing here is indexed, deliberately: a template is a shape, not a fact.
`knowledge\` is the only folder the RAG index reads, and neither
`reference\templates\` nor `reference\exemplars\` is inside it.

## Why `README.md` and not `CLAUDE.md`

`CLAUDE.md` files are auto-loaded instructions for an agent working *in this
repository*. These folders are consumed on your endpoint, by an assistant doing
your work - so what is written here is documentation for you, not instructions
for a coding agent.
