# PowerPoint (.pptx)

Build PowerPoint decks that **inherit your own template's** masters, layouts,
theme, fonts and colours — and audit them against **Guy Kawasaki's 10/20/30
rule**.

| | |
|---|---|
| **Server** | `powerpoint.py` v3.0.0 |
| **pip install** | `python-pptx` (pulls in `lxml`, `Pillow`, `XlsxWriter`, `typing_extensions`) |
| **Platform** | any (PowerPoint itself is not required) |
| **Writes to disk** | yes — confined to its configured folders |
| **Skills** | `/powerpoint:powerpoint` (mechanics) and `/powerpoint:kawasaki` (the 10/20/30 rule) |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install powerpoint@mcnamee-claude-skills
```

Both skills are installed with the server. Claude Code prompts for nothing -
every folder and the Python interpreter come from the shared environment
variables below.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `C:\Eva\reference\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `C:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\reference\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `powerpoint`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_DOCUMENTS_DIR%\powerpoint` | **The one folder of `.pptx` files.** Every open and save must be inside it, and **new** decks from `powerpoint_create` are written here too - there is no separate output folder. Searched recursively | **Fatal.** The server is a path sandbox and refuses to start without it |
| `%EVA_TEMPLATES_DIR%\powerpoint` | Blank `.pptx`/`.potx` deck shells new decks are created from, carrying your masters, layouts, theme, fonts and colours. **Read-only**: they can be listed, opened and passed as `powerpoint_create`'s `template`, but every save into the folder is refused | Warns on stderr and runs with templates disabled |
| `%EVA_KNOWLEDGE_DIR%\powerpoint` | Markdown mirror of every deck opened, created or saved - slides *and* speaker notes (`PowerPoint - <name>.md`, overwritten each time) - which is what makes it searchable by the `knowledge-base` plugin | Created on demand. A mirror failure is logged and reported on the result, never allowed to fail the operation |

> The templates folder must be **separate from** the presentations folder - if
> one contained the other, every save would be refused, so the server refuses to
> start instead.

> **`word` and `powerpoint` no longer share one templates folder.** Each reads
> its own sub-folder of `EVA_TEMPLATES_DIR`, so a template listing only ever
> shows files the asking plugin can actually open. If you are upgrading, move
> your `.pptx`/`.potx` files from `reference\templates\` down into
> `reference\templates\powerpoint\`.

### Overriding one folder, and this server's own settings

The shared roots are normally all you need. These variables are this
server's own, and a folder variable here beats the matching root - use one
only when an endpoint's layout really differs.

| Variable | Purpose |
|---|---|
| `POWERPOINT_DOCS_DIR` | Full path to the presentations folder, instead of `%EVA_DOCUMENTS_DIR%\powerpoint` |
| `POWERPOINT_TEMPLATES_DIR` | Full path to the templates folder, instead of `%EVA_TEMPLATES_DIR%\powerpoint`. `off` runs with no templates root at all |
| `POWERPOINT_KB_DIR` | Full path to the mirror folder, instead of `%EVA_KNOWLEDGE_DIR%\powerpoint`. `off` disables mirroring |

**Blank does not mean off.** A blank value means "not configured", so the shared
root still applies. To switch an optional folder off, set it to `off` (`none`,
`no`, `false` and `disabled` work too). The presentations folder cannot be
switched off: the server has no sandbox without it.

A folder you named yourself that does not exist is **fatal** - it is almost
always a typo. The built-in default merely not existing yet is not: the server
warns and carries on without that feature.

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Run an offline create/build/save/reopen/audit self-test and exit (no server). It sandboxes itself to its own temp folders |
| `--version` | Print version and exit (works even without `python-pptx` installed) |

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
& $env:EVA_PYTHON powerpoint.py --check
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
& $env:EVA_PYTHON -m pip install --no-index `
    --find-links .\wheels python-pptx

# Confirm:
& $env:EVA_PYTHON -c "import pptx; print(pptx.__version__)"
```

A compiled-wheel/interpreter mismatch is the commonest failure; the server logs
`sys.executable` and every dependency version at startup so a mismatch is
obvious.

## Making a template this server can use

Put your branded deck in `%EVA_TEMPLATES_DIR%\powerpoint` (see
[`eva/reference/templates/powerpoint`](../../eva/reference/templates/powerpoint))
and:

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
