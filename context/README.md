# Context

Reference material for Claude — **not code, not a plugin**. Two folders, doing
two different jobs:

| Folder | Holds | Used for |
|---|---|---|
| [**`exemplars/`**](exemplars) | Finished, *good* documents: `.md`, `.docx`, `.pptx`, `.pdf` | Showing Claude **what good looks like** — tone, structure, level of detail — before it writes something of the same kind |
| [**`templates/`**](templates) | Blank/skeleton `.docx` and `.pptx` | The **starting file** a new document is built from — letterhead, styles, headers/footers, boilerplate |

The distinction matters: an exemplar is *read* for guidance and never becomes
the output; a template *is* the output's first draft. A finished board paper
belongs in `exemplars/`; the empty branded shell it was written in belongs in
`templates/`.

Both folders ship with a README explaining what to put in them, how to name it,
and how to ask for it. Start there.

## Wiring it up

Only `templates/` needs configuring, because a server has to open those files.
The `word` plugin takes a **templates folder** setting (`--templates-dir` /
`MSWORD_TEMPLATES_DIR`) — point it at this folder:

```powershell
--templates-dir C:\path\to\claude-skills\context\templates
```

It is a **read-only** third root for that server: templates can be listed,
opened and used as the base for a new document, but nothing can ever be saved
over them. Full setup in [`templates/README.md`](templates/README.md).

`exemplars/` needs no configuration — the files are read by whatever tool suits
the format (`.md` directly, `.docx` via the `word` plugin, `.pdf` via
`pdf-to-md`). If you keep exemplars in a folder the `word` or `knowledge-base`
servers already point at, they are searchable too.

## Why `README.md` and not `CLAUDE.md`

`CLAUDE.md` files are auto-loaded instructions for an agent working *in this
repo*. These folders are reference material used *on your endpoint*, usually
copied somewhere else entirely — so instructions parked here would rarely be
read, and would be read at the wrong time when they were. The per-folder
`README.md` documents the convention for a human; the behaviour Claude follows
lives where it belongs: in each plugin's `SKILL.md`, and in the repo's root
`CLAUDE.md`.

## What is committed

Exemplars and templates are usually your organisation's real documents —
letterhead, board papers, contracts, pricing decks. This repo is public, so
[`.gitignore`](.gitignore) keeps every document file in `context/` **out of git
by default**; only the READMEs are tracked. The folders still work exactly the
same locally.

To publish one deliberately (a genuinely generic sample, say):

```powershell
git add -f context\templates\"Report Template.docx"
```

If none of this material is sensitive and you would rather have it travel with
the repo, delete `context\.gitignore`.
