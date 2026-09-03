# knowledge\powerpoint\

PowerPoint decks, mirrored to Markdown — slide titles as headings, bullets as
nested lists, tables as tables, and **speaker notes as block quotes**. The
`powerpoint` plugin writes a file here whenever it **opens** a deck and whenever
it **saves** one.

| | |
|---|---|
| **Setting** | `EVA_KNOWLEDGE_DIR` (the `powerpoint` plugin appends `\powerpoint`) |
| **Default** | `C:\Eva\knowledge\powerpoint` |
| **Filenames** | `PowerPoint - <name>.md` |
| **Overwritten** | yes, on every open or save |

Written by the server only.

## Why the speaker notes matter here

A deck built to the 10/20/30 rule puts the *headline* on the slide and the
*argument* in the speaker notes. A mirror that captured only the slides would
index the labels and lose the content — you would be able to find that a deck
mentioned pricing, but not what it actually said about it. So the notes are
mirrored too, and they are usually the more useful half.

## Mirroring on save is the useful half

Mirroring on *open* makes the decks in
[`..\..\documents\powerpoint`](../../documents/powerpoint) searchable. Mirroring
on *save* is what captures decks Eva **wrote** — a pitch drafted this morning is
retrievable this afternoon without anyone reopening it.

That does mean a deck saved repeatedly leaves one file here reflecting the last
save. It is a current-state mirror, not a version history.

## Why the `.pptx` files are not in here

This folder holds text extracted for the index. The decks themselves all live
in [`..\..\documents\powerpoint`](../../documents/powerpoint), whether you put
them there or Eva built them — a `.pptx` dropped in here would simply be ignored
by the indexer.
