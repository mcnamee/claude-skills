# documents\powerpoint\

The `powerpoint` plugin's sandbox — every `.pptx` it may open, edit and save.

| | |
|---|---|
| **Setting** | `EVA_DOCUMENTS_DIR` (the `powerpoint` plugin appends `\powerpoint`) |
| **Default** | `C:\Eva\documents\powerpoint` |
| **Access** | read **and write** — decks are edited in place here, and new decks are created here |
| **Sub-folders** | yes, searched recursively |

This is also the base for relative paths: `"open Kickoff.pptx"` resolves against
this folder, so nobody has to type an absolute path. An exact filename always
wins; failing that the plugin falls back to a fuzzy match, so `"kickoff deck"`
still finds the file.

## What belongs here

Every `.pptx` in play: decks you already have and want Eva to read, review or
edit (last quarter's board pack, a colleague's draft), **and** the decks Eva
builds for you. This is the one presentations folder, so a deck is never
somewhere else than you expect.

Blank branded templates are the exception. They go to
[`..\..\templates\powerpoint`](../../templates/powerpoint), which is
read-only, so the blanks stay blank.

Sub-folders are searched recursively, so arrange it however suits you (by
client, by year, a `Drafts\` folder if you like one) and a bare filename still
finds the deck.

## What the plugin may write

- **New decks** from `powerpoint_create` land here, at the top level. Ask for a
  save-as to file one in a sub-folder.
- **Edits** land in the file itself, here.
- **A Markdown mirror** of every deck opened or saved goes to
  [`..\..\knowledge\powerpoint`](../../knowledge/powerpoint) — slides *and*
  speaker notes — which is what makes it searchable.

## A note on `.potx`

PowerPoint's own template format is readable, but only as a *template*. Decks
are always saved as `.pptx`. If you keep a `.potx` here it can be opened and
read; if you want to build from it, put it in
[`..\..\templates\powerpoint`](../../templates/powerpoint).
