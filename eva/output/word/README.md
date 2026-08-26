# output\word\

Documents the `word` plugin creates — the report you asked for, the letter built
from a template, the filled-in form.

| | |
|---|---|
| **Plugin setting** | `word` → output folder (`--output-dir` / `MSWORD_OUTPUT_DIR`) |
| **Default** | `C:\Eva\output\word` |
| **Access** | read **and write** — a second permitted root alongside the documents folder |

New documents always land here, never in
[`..\..\documents\word`](../../documents/word). Because this folder is also a
permitted root, a document created here can be reopened and edited later without
moving it anywhere.

## The lifecycle to expect

1. Eva creates a document here — from scratch, or from a template in
   [`..\..\reference\templates`](../../reference/templates).
2. Saving it mirrors its text to
   [`..\..\knowledge\word`](../../knowledge/word), so a report drafted this
   morning is searchable this afternoon without anyone reopening it.
3. When a document graduates from *drafted* to *kept*, move it to
   [`..\..\documents\word\library`](../../documents/word/library) yourself.
   Nothing moves it for you.

That last step is the only manual one, and skipping it is harmless — this folder
just grows. Treat it as a desk rather than a filing cabinet, and clear it when
it gets untidy.

## Deleting from here

Safe, with one wrinkle: the Markdown mirror in `knowledge\word\` survives, so a
deleted draft stays quotable in answers until you delete the matching
`Word - <name>.md` and reindex. For a discarded draft that is usually fine; for
something that should not have existed, delete both.
