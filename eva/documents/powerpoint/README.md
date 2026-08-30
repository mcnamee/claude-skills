# documents\powerpoint\

The `powerpoint` plugin's sandbox — every `.pptx` it may open, edit and save.

| | |
|---|---|
| **Plugin setting** | `powerpoint` → presentations folder (`--docs-dir` / `POWERPOINT_DOCS_DIR`) |
| **Default** | `C:\Eva\documents\powerpoint` |
| **Access** | read **and write** — decks are edited in place here |
| **Sub-folders** | yes, searched recursively |

This is also the base for relative paths: `"open Kickoff.pptx"` resolves against
this folder, so nobody has to type an absolute path. An exact filename always
wins; failing that the plugin falls back to a fuzzy match, so `"kickoff deck"`
still finds the file.

## What belongs here

Decks you already have and want Eva to read, review or edit — last quarter's
board pack, a colleague's draft, the deck you are revising.

Decks Eva **creates** go to [`..\..\output\powerpoint`](../../output/powerpoint)
instead, so generated work never mixes with your source library. Blank branded
templates go to
[`..\..\reference\templates`](../../reference/templates), which is read-only.

## What the plugin may write

- **Edits** land in the file itself, here.
- **A Markdown mirror** of every deck opened or saved goes to
  [`..\..\knowledge\powerpoint`](../../knowledge/powerpoint) — slides *and*
  speaker notes — which is what makes it searchable.

## A note on `.potx`

PowerPoint's own template format is readable, but only as a *template*. Decks
are always saved as `.pptx`. If you keep a `.potx` here it can be opened and
read; if you want to build from it, put it in the templates folder.
