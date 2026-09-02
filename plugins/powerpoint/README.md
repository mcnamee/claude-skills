# PowerPoint (.pptx)

Build PowerPoint decks that **inherit your own template's** masters, layouts,
theme, fonts and colours — and audit them against **Guy Kawasaki's 10/20/30
rule**.

| | |
|---|---|
| **Server** | `powerpoint.py` v2.0.0 |
| **pip install** | `python-pptx` (pulls in `lxml`, `Pillow`, `XlsxWriter`, `typing_extensions`) |
| **Platform** | any (PowerPoint itself is not required) |
| **Writes to disk** | yes — confined to its configured folders |
| **Skills** | `/powerpoint:powerpoint` (mechanics) and `/powerpoint:kawasaki` (the 10/20/30 rule) |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install powerpoint@mcnamee-claude-skills
```

Claude Code prompts for the settings below; both skills are installed with the
server.

Every folder is pre-filled with its place in the [Eva working
tree](../../eva) — copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and you
can accept all three as they stand.

| Prompt | Default | Env var | Purpose |
|---|---|---|---|
| Presentations folder | `C:\Eva\documents\powerpoint` | `POWERPOINT_DOCS_DIR` | The one folder of `.pptx` files: every open/save must be inside this tree, and **new** decks are created here too. **Required** — the server refuses to start without it |
| Templates folder | `C:\Eva\reference\templates` | `POWERPOINT_TEMPLATES_DIR` | Blank `.pptx`/`.potx` templates new decks are created from — **read-only**. Shared with the `word` plugin. `off` for no templates |
| Knowledge-base folder | `C:\Eva\knowledge\powerpoint` | `POWERPOINT_KB_DIR` | Mirrors every deck opened, created or saved to Markdown — slides *and* speaker notes — which is what makes it searchable. `off` to disable mirroring |
| Python interpreter | — | — | **Required.** Absolute path to the `python.exe` that has `python-pptx` installed |

> **Blank does not mean off.** Leaving a folder prompt empty means "not
> configured", so the default above applies. To switch one off, type `off`
> (`none`, `no`, `false` and `disabled` work too). The presentations folder
> cannot be switched off — the server has no sandbox without it.

> **The templates folder is shared with the `word` plugin** on purpose. Each
> server reads only the file types it understands: `word` reads the `.docx`
> files and ignores the rest, this one reads the `.pptx`/`.potx` files and
> ignores the rest. One folder, one place to look.

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| CLI flag | Env var | Purpose |
|---|---|---|
| `--docs-dir` | `POWERPOINT_DOCS_DIR` | **Required.** Path sandbox *and* the folder `powerpoint_create` writes **new** decks into — there is no separate output folder. Every open/save must be inside this directory tree, and the server refuses to start without one (`--check` is exempt — the self-test sandboxes itself to its own temp folder). Falls back to the `DOCS_DIR` config value, default `C:\Eva\documents\powerpoint` |
| `--templates-dir` | `POWERPOINT_TEMPLATES_DIR` | Folder of blank `.pptx`/`.potx` templates. Falls back to the `TEMPLATES_DIR` config value, default `C:\Eva\reference\templates`; pass `off` for no templates root. A **read-only** second root: its files can be listed, opened and passed as `powerpoint_create`'s `template`, but **every save into it is refused**. Must be separate from the presentations folder — the server refuses to start otherwise. A folder you configured yourself that does not exist is fatal; the built-in default merely not existing yet logs a warning and runs without templates |
| `--kb-dir` | `POWERPOINT_KB_DIR` | **Every deck opened, created or saved** is *also* written as Markdown into this folder for a local RAG knowledge base. Falls back to the `KB_DIR` config value, default `C:\Eva\knowledge\powerpoint` — inside the `knowledge-base` server's documents folder, so mirrored decks are actually indexed. Files are named `PowerPoint - <name>.md` and overwritten each time. A mirror failure is logged and reported on the result, never allowed to fail the operation. Pass `off` to disable |
| `--check` | — | Run an offline create/build/save/reopen/audit self-test and exit (no server) |
| `--version` | — | Print version and exit (works even without `python-pptx` installed) |

The 10/20/30 thresholds are **config constants**, not tool arguments, so a house
style is set once and every review follows it:

| Constant | Default | Meaning |
|---|---|---|
| `KAWASAKI_MAX_SLIDES` | `10` | Slide-count guideline |
| `KAWASAKI_MAX_MINUTES` | `20` | Speaking budget |
| `KAWASAKI_MIN_FONT_PT` | `30` | Minimum body font, in points |
| `SPEAKING_WORDS_PER_MINUTE` | `130` | Turns a speaker-note word count into minutes |
| `SPEAKING_SECONDS_PER_SLIDE` | `15` | Per-slide overhead added to the estimate |

## What it does well

**Building a deck in one call.** `powerpoint_add_slides` takes an ordered list of
slides and appends them in exactly that order — each with its layout, title,
subtitle, bullets, extra placeholder fills, a table and speaker notes:

```json
{"session_id": "...", "slides": [
  {"layout": "title",   "title": "FY26 plan", "subtitle": "Board review",
   "notes": "Thanks for making the time."},
  {"layout": "bullets", "title": "Unpriced risk costs us GBP 4m a year",
   "bullets": ["Claims up 20%", {"text": "mostly EMEA", "level": 1}],
   "notes": "Walk through where the four million goes."},
  {"layout": "section", "title": "Our answer"},
  {"layout": "bullets", "title": "The numbers",
   "table": {"rows": [["Region", "Q3"], ["APAC", "1.2m"]]}},
  {"layout": "two_content", "title": "Now vs next",
   "bullets": ["Today: manual"],
   "placeholders": [{"placeholder": 2, "bullets": ["Next: priced at bind"]}]}
]}
```

This is the tool to build a deck with, and it fixes a real failure mode.
`powerpoint_add_slide` appends to the **end** of the deck, so the deck's order is
the order the *calls arrive* — and an MCP client may dispatch independent tool
calls in parallel, in which case they need not arrive in the order the model
wrote them. Building a ten-slide deck as ten separate calls is a race that can
come back shuffled. The follow-up tools made it worse rather than better:
`powerpoint_set_notes`, `powerpoint_set_placeholder`, `powerpoint_add_bullets`
and `powerpoint_add_table` all address a slide by `slide_index`, so a caller that
assumed *"the third slide I created is index 2"* wrote onto whichever slide
actually landed there. One call carrying the whole sequence cannot be reordered.

Details worth knowing:

- Each entry accepts everything `powerpoint_add_slide` accepts, **plus**
  `placeholders` (`[{"placeholder": <idx|name|type word>, "bullets": [...]}]`,
  for a two-content layout's second column) and `table` (`{"rows": [[...]]}`) —
  the two things that previously needed a follow-up call and a `slide_index`.
- Every entry, **including every layout name**, is validated before any slide is
  added, so a bad entry leaves the deck untouched and the error names its index.
  A layout/content mismatch that can only surface mid-build (bullets onto a
  layout with no body placeholder) unwinds the whole batch rather than leaving a
  half-written deck.
- The result reports each slide's real `slide_index` in order, so any later edit
  addresses the right slide.
- Still chaining `powerpoint_add_slide`? Three or more calls landing on the same
  session within a second — the signature of a parallel batch — adds an
  `order_warning` to the result, so a shuffled deck gets reported rather than
  silently shipped.

**Sticking to a template.** This is the whole point of the server. A deck's look
lives in its masters, layouts and theme, so `powerpoint_create` with a
`template` inherits all of it, and every content tool writes into the layout's
**placeholders** rather than setting formatting. There is deliberately **no tool
that sets a font, size, colour or position** — that omission is the feature. Any
example slides in the template are stripped, so you start blank with the styling
intact, and the template file itself is never modified.

**Reading what a template actually offers.** `powerpoint_list_layouts` is the
tool that makes the above usable. A branded template names its layouts whatever
it likes, so assuming the stock Office numbering is the commonest way to produce
an off-template deck. It reports every layout's index, name, a **detected role**
(`title` / `section` / `bullets` / `two_content` / `title_only` / `picture` /
`blank`, worked out from the placeholders rather than the name), and each
placeholder's `idx`, `type` and **effective font size**. Layouts can then be
addressed by index, name, *or* role word — `layout: "bullets"` works on any
template.

**Resolving font sizes properly.** In a well-built deck almost no run carries an
explicit size — it is inherited — so `run.font.size` is `None` nearly
everywhere, and the naive check finds nothing wrong with a deck set entirely in
12pt. This server resolves the size the way PowerPoint does, through the whole
chain: run → paragraph → shape `lstStyle` → layout placeholder → master
placeholder → master `txStyles` → presentation default → PowerPoint's own 18pt
fallback. Every reported size names **which step supplied it**, so *"why is this
28pt?"* has an answer.

**The 10/20/30 audit.** `powerpoint_review` checks slide count against 10,
estimated speaking time against 20 minutes (from the speaker-note word count,
with every input shown so the number can be argued with), and every run below 30
points — with the slide, shape, resolved size and source. It also flags slides
carrying more than 40 words of body text, because a slide can pass on font size
and still be a document. It **reports rather than fixes**: the fixes are content
decisions — say less, cut a slide, choose a different layout — not font
overrides, which would break the template adherence the server exists to
protect. `powerpoint_save` returns the headline automatically.

**Speaker notes as first-class content.** `notes` on each `powerpoint_add_slides`
entry, or `powerpoint_set_notes` against a slide index you read back. Under the rule the slide carries the headline and the
notes carry the argument; the notes are also the only input to the timing
estimate, and they are mirrored into the knowledge base alongside the slides, so
a deck you wrote is searchable by what you meant rather than by its headlines
alone.

**Finding files by name.** You don't need absolute paths. Names resolve against
the presentation root, a bare filename is found in subfolders, and a near-miss
falls back to a **fuzzy match** (*"acme deck"* → `Acme Deck Template.pptx`),
flagged with `fuzzy_matched` so the caller can confirm. A genuinely ambiguous
name returns the tied candidates rather than guessing.

## Tools

| Tool | Does |
|---|---|
| `powerpoint_list_presentations` | List decks and templates, each with its `location` (`docs`/`templates`) |
| `powerpoint_create` | New deck in the presentations folder, optionally inheriting a `template`'s design |
| `powerpoint_open` | Open an existing deck (or inspect a template) |
| `powerpoint_list_layouts` | The template's layouts, placeholders, roles and **effective font sizes** |
| **`powerpoint_add_slides`** | **Append MANY slides in one call, in order** — the tool to build a deck with |
| `powerpoint_add_slide` | Append ONE slide on a layout, filling title / subtitle / bullets / notes |
| `powerpoint_set_placeholder` | Replace one placeholder's text on an existing slide |
| `powerpoint_add_bullets` | Append bullets without clearing |
| `powerpoint_set_notes` | Set speaker notes |
| `powerpoint_add_table` | Add a table, taking over the layout's content placeholder position |
| `powerpoint_get_content` | Read slides, paragraphs (with levels), tables and notes |
| `powerpoint_delete_slide` | Delete a slide — indices shift down |
| `powerpoint_move_slide` | Reorder, returning the resulting order by title |
| `powerpoint_review` | Audit against 10/20/30 with evidence |
| `powerpoint_save` | Save in place or save-as, with a 10/20/30 headline |
| `powerpoint_close` / `powerpoint_list_sessions` | Session management |

## Validate before wiring in

```powershell
& "C:\path\to\python.exe" powerpoint.py --check
```

Builds a temp template, creates a deck from it, adds slides, saves, reopens and
audits — 73 assertions, no network, nothing left behind. Expected tail:

```
[check] ALL CHECKS PASSED
```

### How the font resolution was verified

The inheritance chain is the part most worth distrusting, so it was checked
against an **independent renderer**. A deck built from a template whose master
sets `titleStyle` to 40pt and `bodyStyle` level 1 to 32pt was exported to PDF by
LibreOffice Impress, and the rendered sizes read back out:

| Element | This server | LibreOffice rendered |
|---|---|---|
| Title | 40pt | 40.0pt |
| Bullet, level 1 | 32pt | 32.0pt |
| Bullet, level 2 | 28pt | **28.0pt** — flagged, correctly |
| Table cells | *(unmeasured)* | 18.0pt |

Also confirmed: the saved package carries the template's own master and theme
parts, its slides contain **no hard-coded `sz=` attributes** (the text really
does inherit rather than being baked in), and deleting slides leaves no orphaned
parts in the package.

**Table text is the honest gap.** It is sized by the table style in
`tableStyles.xml`, which this server does not read, so it is reported under
`unmeasured` with a clearly-labelled estimate rather than counted as a
violation. In the stock theme that estimate happened to match the render
exactly; in a template whose table style differs from its `otherStyle` it would
not — which is precisely why it is not asserted as a measurement.

The one thing not verifiable off your network is Microsoft PowerPoint's own
rendering. Open one generated deck in PowerPoint and confirm the template's
fonts and colours look right before relying on this for anything that matters.

## Dependencies (airgapped install)

`python-pptx` needs the **same `lxml`** the `word` plugin does, so if `word.py`
already runs on an interpreter most of the work is done. `lxml` and `Pillow` are
compiled, so their wheels must match your exact Python version and architecture.

```powershell
# On an internet-connected box, ideally running the target Python version:
python -m pip download python-pptx -d .\wheels

# On the endpoint, with the SAME interpreter the MCP client will launch:
& "C:\path\to\python.exe" -m pip install --no-index `
    --find-links .\wheels python-pptx

# Confirm:
& "C:\path\to\python.exe" -c "import pptx; print(pptx.__version__)"
```

A compiled-wheel/interpreter mismatch is the commonest failure; the server logs
`sys.executable` and every dependency version at startup so a mismatch is
obvious.

## Making a template this server can use

Put your branded deck in the templates folder (see
[`eva/reference/templates`](../../eva/reference/templates)) and:

- **Define the layouts you want used**, and name them clearly. Role detection
  reads the placeholders, so a `Chapter Opener` with a title and one body
  placeholder is still found by `layout: "section"` if its name says so, and by
  `"bullets"` otherwise.
- **Set your sizes in the slide master**, not on individual slides. That is
  where `powerpoint_list_layouts` reads them from, and it is what lets the whole
  deck restyle at once.
- **Check the deeper outline levels.** Nearly every template shrinks each level
  (the stock Office one runs 32 / 28 / 24 / 20 point), so a level‑2 bullet
  breaks the 30-point rule without anyone choosing a font.
- **Keep the slides empty.** Example slides are stripped on create, so they cost
  nothing — but a template that is *only* masters and layouts is clearer.
- `.potx` works as a template; decks are always saved as `.pptx`.

## Limitations

- **No charts, SmartArt, animations or transitions.** A chart needs data
  plumbing this server deliberately does not carry.
- **No images.** Nothing here reads image files, by design — the sandbox stays a
  text-only surface. Use a layout with a picture placeholder (empty ones are
  kept, never dropped) and drop the image in by hand.
- **No theme editing.** Changing a template's fonts or colours is the template's
  job; doing it here would defeat the point.
- The **timing estimate is a planning aid, not a stopwatch** — a constant
  words-per-minute never matches a real delivery. It is reported with all its
  inputs so it can be judged.
