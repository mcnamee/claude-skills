# templates\powerpoint\

Blank branded deck shells the `powerpoint` plugin creates new presentations
from. A new deck **inherits** the template's slide masters, layouts, theme,
fonts and colours, and lands in
[`..\..\documents\powerpoint`](../../documents/powerpoint). The template itself
is never modified, and any example slides in it are stripped so the new deck
starts blank with the styling intact.

| | |
|---|---|
| **Plugin setting** | `powerpoint` → templates folder (`--templates-dir` / `POWERPOINT_TEMPLATES_DIR`) |
| **Default** | `C:\Eva\templates\powerpoint` |
| **Formats** | `.pptx` and `.potx` (both read-only; decks always save as `.pptx`) |
| **Access** | read **only** — every save into this folder is refused |
| **Committed to git?** | no (see [`..\..\.gitignore`](../../.gitignore)) |

If what you have is a *finished* deck showing what good output looks like, that
is an exemplar, and it belongs in the `exemplars\` folder of the skill that
reads it. See [`..\README.md`](../README.md).

## Naming

End the name with `Template` so intent is obvious in a listing:

```
Deck Template.pptx
Board Pack Template.pptx
Client Pitch Template.potx
```

Names are matched forgivingly, so keep them distinct. If a name here also exists
in the presentations folder, that folder wins.

## Index

| File | Creates | Notes |
|---|---|---|
| _(none yet — add yours here)_ | | |

## What makes a good deck template

For a document it is styles and placeholders; for a deck it is **layouts**,
because that is the only thing the server writes into.

- **Define a layout for each job** — a cover, a section divider, a
  title-and-bullets content slide, a two-column slide, a title-only slide. The
  server detects each layout's *role* from its placeholders rather than its
  name, so `Chapter Opener` and `Section Header` both work; name them for
  humans.
- **Set your font sizes in the slide master**, not on individual slides. That is
  where they are read from, and it is what lets a whole deck restyle at once.
- **Check the deeper outline levels.** Nearly every template shrinks each level —
  the stock Office one runs 32 / 28 / 24 / 20 point — so a level-2 bullet breaks
  the 30-point rule without anyone choosing a font. If your house style is
  30-point minimum, set level 2 to 30 too.
- **Keep the slides empty.** Example slides are stripped on create, so they cost
  nothing, but a template that is only masters and layouts is clearer.
- **Leave picture placeholders in.** The server writes no images and never drops
  an empty picture placeholder, so one in the layout becomes a correctly
  positioned, correctly styled click target for you to fill in by hand.
- `.potx` works and is only ever read; decks are always saved as `.pptx`.

## How it gets used

- *"Build a pitch deck from our template."* → `powerpoint_create` with
  `template: "Deck Template.pptx"`, then `powerpoint_list_layouts` to see what
  it offers, then one `powerpoint_add_slides` call.
- *"What layouts does our template have, and will they hold 30-point text?"* →
  `powerpoint_list_layouts`, which reports each placeholder's *effective* font
  size and lists the layouts that clear the minimum.
- *"What templates do I have?"* → `powerpoint_list_presentations` with
  `location: "templates"`.
- *"Is this deck too long?"* → `powerpoint_review`, which audits slide count,
  estimated speaking time and font sizes against the 10/20/30 rule.

The full workflow is in the
[`powerpoint` plugin README](../../../plugins/powerpoint/README.md) and its two
skills.
