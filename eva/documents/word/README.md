# documents\word\

The `word` plugin's sandbox — every `.docx` it may open, edit and save.

| | |
|---|---|
| **Plugin setting** | `word` → documents folder (`--docs-dir` / `MSWORD_DOCS_DIR`) |
| **Default** | `C:\Eva\documents\word` |
| **Access** | read **and write** — documents are edited in place here, and new documents are created here |
| **Sub-folders** | yes, searched recursively |

This is also the base for relative paths: `"open Budget Policy 2024.docx"`
resolves against this folder, so nobody has to type an absolute path. An exact
filename always wins; failing that the plugin falls back to a fuzzy match, so
`"budget policy"` still finds the file.

## What belongs here

Every `.docx` in play: the contract you were sent, the policy being revised, the
paper you are marking up, **and** the documents Eva writes for you. One folder,
so nothing has to be filed before it can be asked about and nothing has to be
moved afterwards.

Blank branded templates are the exception. They live in
[`..\..\templates\word`](../../templates/word), which is read-only: the plugin
can create *from* a template but every save into that folder is refused, so the
blanks stay blank.

## Arrange it however you like

Sub-folders cost nothing — the plugin searches recursively and matches on
filename, so `"open Contract v3.docx"` finds the file whether it is at the top
level or in `Policies\2026\`. A `Drafts\` folder works if you want one; so does
no structure at all.

There is deliberately **no** `inbox\` and `library\` split any more, and no
`input\` or `output\` folder. Each was a decision you had to make about a file
before you could do anything with it, and the plugin searched all of them
anyway.

## What the plugin may write

- **New documents** from `msword_create` land here, at the top level. Ask for a
  save-as if one belongs in a sub-folder.
- **Edits** land in the file itself, here, with Word tracked changes if asked.
- **A Markdown mirror** of anything opened, created or saved goes to
  [`..\..\knowledge\word`](../../knowledge/word), which is what makes these
  documents searchable.

## This is not a safe copy

The plugin can save over a document here, in place. For a policy being revised
that is exactly what you want, but it means this folder is not an archive.

- **Ask for tracked changes** on anything you might want to reject. The plugin
  writes real Word revision marks, so the document opens in Word with every edit
  reviewable.
- **Keep the authoritative copy elsewhere** — SharePoint, a document management
  system, wherever it already lives. This is a working folder, not a system of
  record.

## Before you delete from here

Deleting a file does **not** remove its Markdown mirror from
[`..\..\knowledge\word`](../../knowledge/word), so its text stays searchable and
quotable after the document is gone. That is usually what you want. When it is
not — a document that should never have existed — delete the matching
`Word - <name>.md` too, and reindex.

## One caution

Only open `.docx` files from sources you trust. A maliciously crafted document
can use XML entity tricks to pull the contents of other local files into text
that the model then reads.
