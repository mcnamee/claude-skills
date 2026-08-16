# Exemplars

Finished documents that are **good** — the ones you would hand a new starter and
say *"write it like this"*. Claude reads them to pick up the house style, the
structure, the level of detail and the tone, and then writes something new in
that shape.

An exemplar is read for guidance and **never becomes the output**. If what you
want is the branded, empty file a document is *built from*, that is a template —
put it in [`../templates`](../templates) instead.

| | |
|---|---|
| **Formats** | `.md`, `.docx`, `.pptx`, `.pdf` |
| **Configuration** | none — these are read on request, not by a server |
| **Committed to git?** | no (see [`../.gitignore`](../.gitignore)) |

## What belongs here

- The **best** real example of each kind of thing you produce — one board paper,
  one project status report, one client proposal, one decision record.
- Documents that are **finished and approved**, not drafts. Claude will copy the
  habits it finds, including the bad ones.
- Anything **safe to show an assistant**. Strip or fake client names, pricing and
  personal data first; a redacted exemplar teaches structure just as well.

## What does not

- Blank templates and letterhead → [`../templates`](../templates).
- Twelve near-identical reports. Two or three strong examples beat a folder of
  variations, and a big folder makes "which one?" ambiguous.
- Reference *data* to answer questions from — that is what the
  [`knowledge-base`](../../plugins/knowledge-base) plugin indexes.

## Naming

Lead with the kind of document, because that is what gets matched against a
request. Keep it plain:

```
Board Paper - Quarterly Risk Update.docx
Status Report - Migration Programme.md
Proposal - Managed Service.pdf
Deck - Programme Kickoff.pptx
```

Then add a row to the index below, saying **what each file is good for**. That
one line is what lets Claude (or you) pick the right exemplar without opening
all of them.

## Index

| File | Kind of document | Use it as the model for |
|---|---|---|
| _(none yet — add yours here)_ | | |

## How to use them

Name the exemplar, or name the kind of document and let Claude match:

- *"Write the Q3 status report for the migration programme — follow
  `Status Report - Migration Programme.md` for structure and tone."*
- *"Draft a board paper on the licence renewal. Use the board paper exemplar in
  `context\exemplars` as the model — same sections, same length, same
  register."*
- *"Read the proposal exemplar first, then tell me what sections my draft is
  missing."*

Reading them costs a tool call per file, so point at the one that matters rather
than *"read everything in exemplars"*.

**How each format is read**

| Format | Read by |
|---|---|
| `.md` | directly — cheapest, and the easiest to diff and review |
| `.docx` | the [`word`](../../plugins/word) plugin (`msword_open` → `msword_get_content`) |
| `.pdf` | the [`pdf-to-md`](../../plugins/pdf-to-md) plugin, which converts it to Markdown first |
| `.pptx` | no MCP server in this suite reads PowerPoint; Claude Code's own `pptx` skill can, where it is available |

If an exemplar is one you reach for constantly, keeping a `.md` copy beside the
original is the cheapest option — no conversion step, and the structure is
visible at a glance.

## A note on style guidance

An exemplar shows *what* good output looks like. It does not stop the writing
sounding like an AI wrote it — for that, ask for [`/unslop`](../../skills/unslop)
over the draft.
