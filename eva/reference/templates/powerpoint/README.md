# reference\templates\powerpoint\

Blank, branded `.pptx` (or `.potx`) deck shells - your title slide, section
divider and content layouts, with the masters and theme that carry your fonts
and colours. A new deck is **created from** one of these, inheriting all of
that; any example slides in it are stripped so the deck starts blank with the
styling intact, and the template itself is never modified.

| | |
|---|---|
| **Read by** | the [`powerpoint`](../../../../plugins/powerpoint) plugin |
| **Setting** | `EVA_TEMPLATES_DIR` (the plugin appends `\powerpoint`) |
| **Default path** | `C:\Eva\reference\templates\powerpoint` |
| **Must exist?** | **yes** - without it the plugin starts with templates disabled |
| **Writable?** | **no** - the server refuses every save into this folder |
| **Committed to git?** | no (see [`../../../.gitignore`](../../../.gitignore)) |

`.docx` templates go in [`..\word`](../word) instead: each plugin reads only
its own sub-folder, so a listing here is always just the deck shells.

## Naming

End the name with `Template` so intent is obvious in a listing:

```
Deck Template.pptx
Client Pitch Template.pptx
Board Update Template.pptx
```

## What makes a good deck template

For a document it is about styles and placeholders; for a deck it is about
**layouts**, because that is the only thing the server writes into.

- **Define a layout for each job** - a cover, a section divider, a
  title-and-bullets content slide, a two-column slide, a title-only slide. The
  server detects each layout's *role* from its placeholders rather than its
  name, so `Chapter Opener` and `Section Header` both work; name them for
  humans.
- **Set your font sizes in the slide master**, not on individual slides. That is
  where they are read from, and it is what lets a whole deck restyle at once.
- **Check the deeper outline levels.** Nearly every template shrinks each level -
  the stock Office one runs 32 / 28 / 24 / 20 point - so a level-2 bullet breaks
  the 30-point rule without anyone choosing a font. If your house style is
  30-point minimum, set level 2 to 30 too.
- **Keep the slides empty.** Example slides are stripped on create, so they cost
  nothing, but a template that is only masters and layouts is clearer.
- **Leave picture placeholders in.** The server writes no images and never drops
  an empty picture placeholder, so one in the layout becomes a correctly
  positioned, correctly styled click target for you to fill in by hand.
- `.potx` works and is only ever read; decks are always saved as `.pptx`.

## How it gets used

- *"Build a pitch deck from our template."* -> `powerpoint_create` with
  `template: "Deck Template.pptx"`, then `powerpoint_list_layouts` to see what
  it offers, then a `powerpoint_add_slides` call for the whole deck. The new
  file lands in `C:\Eva\documents\powerpoint`.
- *"What layouts does our template have, and will they hold 30-point text?"* ->
  `powerpoint_list_layouts`, which reports each placeholder's *effective* font
  size and lists the layouts that clear the minimum.
- *"Is this deck too long?"* -> `powerpoint_review`, which audits slide count,
  estimated speaking time and font sizes against the 10/20/30 rule.

## Index

| File | Creates | Notes |
|---|---|---|
| _(none yet - add yours here)_ | | |

The full workflow is in the [`powerpoint` plugin
README](../../../../plugins/powerpoint/README.md) and its two skills.
