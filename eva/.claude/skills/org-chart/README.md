# org-chart

Draws an organisation chart as a single **SVG** in the house design: a dark
canvas, the organisation's name in wide capitals, the leader in a box, an
optional row of leadership boxes, and one card per team showing its headcount,
its leads and its members.

You give it staffing data (usually a spreadsheet, often one attached to a
Confluence page). Claude works out which column is which and who reports to
whom, writes that structure to a small JSON file, and runs the bundled
`orgchart.py` to draw it. So every chart gets the same spacing, connectors and
type, and redrawing one after a staffing change is a one-line edit to the JSON.

This is the one standalone skill that ships code: `orgchart.py`, a single file
using only the Python standard library. It runs under the same `EVA_PYTHON`
interpreter as the plugins, and needs no plugin of its own.

## Install

Nothing extra: it ships in the `eva\` scaffold at `eva\.claude\skills\org-chart\`,
so copying `eva\` to `H:\Eva` installs it at `H:\Eva\.claude\skills\org-chart\`
for Claude Code opened in `H:\Eva`. It needs:

- `EVA_PYTHON` set (the plugins need it anyway).
- The `H:\Eva\documents\images` folder, where charts are saved. The `eva\` copy
  creates it.
- For the Confluence round trip: the `confluence` plugin (7.3.0 or later) and the
  `excel` plugin, with `CONFLUENCE_ALLOW_WRITE=true` if the chart is to be
  uploaded.

## Use

```
/org-chart Build the org chart from the staffing spreadsheet on the "Team Structure" page
/org-chart Redraw the Engineering Org Chart - Priya Raman has moved to Flight Software
```

A typical round trip with Confluence:

1. *"Get the staffing spreadsheet attached to the Team Structure page."*
   `confluence_download_attachment` saves it to `documents\excel`.
2. *"Make an org chart from it."* Claude reads the workbook, shows you how it
   mapped the columns and the teams (when anything was a judgement call), then
   writes `documents\images\<name>.json` and `<name>.svg`.
3. *"Replace the org chart on the Team Structure page with it."*
   `confluence_upload_attachment` uploads the SVG as a **new version of the
   existing image**, so the page shows it without being edited.

## Layouts

- **Shared**, the reference design: the leader feeds the leadership row, and
  the leadership row feeds every team together. Used when teams report to the
  leadership jointly, or the data does not say.
- **Grouped**: each leadership box sits over its own teams. Used when every
  team's lead reports to one leadership person.
- **Flat**: no leadership row; the teams hang off the leader.

A **light** theme (the same layout, inverted) is there for a page where a dark
block would look out of place. Ask for it by name.

## What it will not do

- **Invent anyone.** A person the data cannot place is reported to you, not
  guessed into a team. An unfilled position is drawn as `VACANT`, dimmed, and
  left out of the headcount unless you ask otherwise.
- **Overwrite a chart** it was not asked to redraw.
- **Wrap text.** SVG cannot, so a very long position title widens every card.

## Running the script yourself

```powershell
& $env:EVA_PYTHON H:\Eva\.claude\skills\org-chart\orgchart.py --sample H:\Eva\documents\images\sample.json
& $env:EVA_PYTHON H:\Eva\.claude\skills\org-chart\orgchart.py H:\Eva\documents\images\sample.json
start H:\Eva\documents\images\sample.svg
```

`--sample` writes the reference design's data as a worked example. Add
`--overwrite` to replace an existing SVG, or `-o` to name the output. The JSON
format and every option are documented at the top of `orgchart.py`.

Fonts are Segoe UI and Consolas (on every Windows machine); nothing is embedded,
so the file stays small and works offline.
