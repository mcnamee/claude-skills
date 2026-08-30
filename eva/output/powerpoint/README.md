# output\powerpoint\

Decks the `powerpoint` plugin creates — the pitch you asked for, the board
update built from your corporate template.

| | |
|---|---|
| **Plugin setting** | `powerpoint` → output folder (`--output-dir` / `POWERPOINT_OUTPUT_DIR`) |
| **Default** | `C:\Eva\output\powerpoint` |
| **Access** | read **and write** — a second permitted root alongside the presentations folder |

New decks always land here, never in
[`..\..\documents\powerpoint`](../../documents/powerpoint). Because this folder
is also a permitted root, a deck created here can be reopened and edited later
without moving it anywhere.

## The lifecycle to expect

1. Eva creates a deck here, normally from a template in
   [`..\..\reference\templates`](../../reference/templates), inheriting its
   masters, layouts, theme, fonts and colours.
2. Saving it mirrors its slides **and speaker notes** to
   [`..\..\knowledge\powerpoint`](../../knowledge/powerpoint), so a deck drafted
   this morning is searchable this afternoon.
3. When a deck graduates from *drafted* to *kept*, move it to
   [`..\..\documents\powerpoint`](../../documents/powerpoint) yourself. Nothing
   moves it for you.

That last step is the only manual one, and skipping it is harmless — this folder
just grows. Treat it as a desk rather than a filing cabinet, and clear it when
it gets untidy.

## What is still yours to finish

The server writes no images. A layout with a picture placeholder keeps that
placeholder empty rather than dropping it, so opening the deck in PowerPoint
gives you a click target in the right place, styled by the template. Charts,
SmartArt, animations and transitions are the same: add them by hand.

## Deleting from here

Safe, with one wrinkle: the Markdown mirror in `knowledge\powerpoint\` survives,
so a deleted draft stays quotable in answers until you delete the matching
`PowerPoint - <name>.md` and reindex. For a discarded draft that is usually
fine; for something that should not have existed, delete both.
