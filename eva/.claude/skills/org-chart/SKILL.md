---
name: org-chart
description: Draw an organisation chart as an SVG in the house design (dark canvas, boxed leader, a leadership row, one card per team with its leads and members) from staffing data - usually a spreadsheet, often one attached to a Confluence page. Use when asked to make, draw, build, update or redraw an org chart, organisation chart, team structure diagram or staffing chart, or to turn a staff list, team list or reporting-lines spreadsheet into a chart. Maps the columns, confirms the structure, writes a JSON file and runs the bundled orgchart.py to produce the SVG in documents\images, ready to upload with confluence_upload_attachment. Not for a Visio or PowerPoint org chart.
---

# Org chart

Turn staffing data into an SVG org chart in the house design. You do the
thinking (which column is which, who reports to whom); the bundled script does
the drawing, so every chart has the same spacing, connectors and type.

- **Script:** `orgchart.py`, in this skill's folder (normally
  `H:\Eva\.claude\skills\org-chart\orgchart.py`). Standard library only.
- **Run it with** the interpreter in `EVA_PYTHON`:
  PowerShell `& $env:EVA_PYTHON "<skill folder>\orgchart.py" "<json>"`,
  Bash `"$EVA_PYTHON" "<skill folder>/orgchart.py" "<json>"`.
- **Output folder:** `H:\Eva\documents\images` (or `%EVA_DOCUMENTS_DIR%\images`).
  Write the JSON there and the SVG lands beside it. It is the folder
  `confluence_upload_attachment` reads images from, so the chart can go
  straight onto a page.

## 1. Get the data

Take it from wherever the user points:

- **A spreadsheet on a Confluence page:** `confluence_list_attachments` on the
  page, `confluence_download_attachment` for the workbook (it lands in
  `documents\excel`), then `excel_list_sheets` and `excel_read_table` (a named
  Table) or `excel_read_range` (plain cells).
- **A workbook already in `documents\excel`:** the same `excel_*` tools.
- **A CSV, a pasted list, or a table on a page:** read it as given
  (`confluence_list_tables` for a Confluence table).
- **An earlier chart's JSON** in `documents\images`: read it, apply the change
  asked for, and redraw. That is the quickest way to update one.

Read every row. A chart that quietly drops people is worse than no chart.

## 2. Work out the structure

The chart has four tiers. Map the data onto them:

| Tier | What it is | Drawn as |
|---|---|---|
| **Top** | The one person everyone ultimately reports to | The boxed leader at the top |
| **Leadership** | Optional: the people directly under the top who head the teams (technical directors, say) | A row of boxes |
| **Team leads** | The people named as leading each team, with their lead title (Team Lead, Tech Lead) | Listed at the head of the team's card |
| **Members** | Everyone else in the team, with their position title | The tree list in the card |

Columns to find (the names vary, so match on meaning):

- **Name:** `Name`, `Employee`, `Staff member`, or `First name` + `Surname`
  joined.
- **Position:** `Title`, `Position`, `Role`, `Job title`.
- **Team:** `Team`, `Section`, `Unit`, `Branch`, `Group`.
- **Manager:** `Reports to`, `Manager`, `Supervisor`, `Line manager`.
- **Status:** `Vacant`, `Status`, `Filled`, or a name cell reading "Vacant" /
  "TBC". An unfilled position is `"vacant": true`, never a made-up name.

Then decide:

- **Top:** the person with no manager, or whose manager is not in the sheet.
- **Leadership:** the people reporting to the top who are not in a team.
- **Leads:** within each team, people whose position contains "Lead",
  "Manager" or "Head", or whom other team members report to. Order them as
  the sheet does (Team Lead before Tech Lead in the reference design).
- **Which layout:** if each team's lead reports to ONE leadership person, set
  that team's `reports_to` to them and each leadership box sits over its own
  teams. If the teams report to the leadership row jointly, or the sheet does
  not say, leave `reports_to` out on every team: that is the reference
  design's shared layout. It must be all or none; the script refuses a mix.
- **Order:** teams and members in the sheet's order unless the user says
  otherwise.

**Confirm before drawing** when anything above was a judgement call: which
column you took as which, who the top is, the layout, and anyone you could
not place. Show it as one short table (team, leads, member count) plus a
list of the unplaced, and ask once. When the data is unambiguous (clear
columns, one top, every person placed), say what you mapped and go straight on.

Never invent a person, a title or a reporting line. Transcribe names and
titles exactly as the source has them (the script does the capitals). A blank
title stays blank. Someone who fits no tier is reported to the user, not
dropped and not guessed into a team.

## 3. Write the JSON

Save it as `H:\Eva\documents\images\<Chart name>.json`, where the chart name is
what the user would call it (`Engineering Org Chart`). The format:

```json
{
  "title": "Halcyon Orbital",
  "subtitle": "Organisation chart",
  "as_of": "2026-09-30",
  "theme": "dark",
  "count": "filled",
  "top": {"name": "Mara Keller", "role": "Director"},
  "leadership": [
    {"name": "Tomas Lindqvist", "role": "Technical Director"},
    {"name": "Aiko Tanaka", "role": "Technical Director"}
  ],
  "teams": [
    {
      "name": "Propulsion",
      "leads": [
        {"name": "Erik Johansson", "role": "Team Lead"},
        {"name": "Ana Ribeiro", "role": "Tech Lead"}
      ],
      "members": [
        {"name": "Victor Hale", "role": "Turbomachinery Engineer"},
        {"role": "Test Technician", "vacant": true}
      ]
    }
  ]
}
```

- `title` is the organisation or group name, top left. `subtitle` defaults to
  "Organisation chart".
- `as_of` is the date the data is true at, not today's date by reflex: take it
  from the spreadsheet or page (its "last updated" or version date) and say
  which you used. ISO dates print as `30 SEPT 2026`; `""` prints no date;
  leaving it out prints today.
- `theme`: `dark` is the house design. `light` is the same layout inverted,
  for a page or document where a dark block would look out of place. Use it
  only when asked.
- `count` is the number on each team's header: `filled` (people, the default),
  `all` (vacancies included) or `none`.
- `reports_to` on a team names a `leadership` person exactly (see step 2).

## 4. Draw it

Run the script on the JSON. It writes `<Chart name>.svg` beside it and prints
a summary to stderr: the size, how many people and teams, vacancies, the
layout, and whether cards were widened for long text.

- **An SVG of that name already exists:** the script refuses. Re-run with
  `--overwrite` only when the user asked to redraw that chart (updating one is
  asking), otherwise pick another name with `-o`.
- **It reports a data error** (exit code 2): the message names the field. Fix
  the JSON from the source data, or ask the user if the source cannot settle
  it. Do not paper over it.
- **Check the count.** The people in the summary must equal the people in the
  source, less anyone you reported as unplaced. If they differ, find out why
  before handing over.

## 5. Hand it over

Say where the SVG is and what is on it (teams, headcount, vacancies, the as-of
date) and list any assumptions from step 2. Then offer the next step rather
than taking it:

- **Replacing the chart on a Confluence page:** `confluence_list_attachments`
  on that page for the current image's exact name, then
  `confluence_upload_attachment` with `file` = the SVG's name, `attach_as` = the
  name on the page and `overwrite: true`. The page shows the new version with
  no page edit. Only when the user has asked for it to be replaced; if the
  upload tool is missing, Confluence writing is switched off on this endpoint
  (`CONFLUENCE_ALLOW_WRITE=true` turns it on). Say so; do not work around it.
- **A new chart on a page:** upload it (no `overwrite`), then add
  `![](<Chart name.svg>)` with `confluence_update_section` or
  `confluence_append_to_page`.

## Limits

- **One row of teams.** Past about eight teams the chart gets very wide and
  shrinks on the page. Offer one chart per leadership box instead.
- **No text wrapping.** SVG cannot wrap, so a long title widens every card.
  If the summary says cards were widened a lot, suggest shortening the title
  on the chart (with the user's agreement: it is their data).
- **Fonts:** Segoe UI and Consolas, which every Windows machine has; nothing
  is embedded. A viewer on another platform sees Helvetica/Arial and Courier.
- **Showing it on Confluence** depends on the Confluence version rendering
  SVG in an image. If the page already shows an SVG, it does.
