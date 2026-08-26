# documents\word\

The `word` plugin's sandbox — every `.docx` it may open, edit and save.

| | |
|---|---|
| **Plugin setting** | `word` → documents folder (`--docs-dir` / `MSWORD_DOCS_DIR`) |
| **Default** | `C:\Eva\documents\word` |
| **Access** | read **and write** — documents are edited in place here |
| **Sub-folders** | yes, searched recursively |

This is also the base for relative paths: `"open Budget Policy 2024.docx"`
resolves against this folder, so nobody has to type an absolute path. An exact
filename always wins; failing that the plugin falls back to a fuzzy match, so
`"budget policy"` still finds the file.

## The two sub-folders

| | |
|---|---|
| [`inbox\`](inbox) | documents you have just dropped in to work on |
| [`library\`](library) | documents you keep |

Both are searched, and a bare filename finds a document in either — the split is
for **your** benefit, so that clearing out finished work is one folder you can
empty rather than a judgement call over a mixed list.

## What the plugin may write

- **Edits** land in the file itself, here, with Word tracked changes if asked.
- **New documents** go to [`..\..\output\word`](../../output/word), never here,
  so generated files never mix with your source library.
- **A Markdown mirror** of anything opened or saved goes to
  [`..\..\knowledge\word`](../../knowledge/word), which is what makes these
  documents searchable.

Templates are a separate read-only root at
[`..\..\reference\templates`](../../reference/templates): the plugin can create
*from* a template but every save into that folder is refused.

## One caution

Only open `.docx` files from sources you trust. A maliciously crafted document
can use XML entity tricks to pull the contents of other local files into text
that the model then reads.
