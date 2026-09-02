# Exemplars

Finished documents that are **good** - the ones you would hand a new starter and
say *"write it like this"*. [`/exemplar-writer`](../SKILL.md) reads them for
structure, section order, proportions and register, then writes something new in
that shape.

An exemplar is read for guidance and **never becomes the output**. No fact,
figure, date, name or dollar amount is ever carried across from one.

| | |
|---|---|
| **Formats** | `.md`, `.txt` always; `.docx`, `.pptx`, `.pdf` with a caveat (below) |
| **Committed to git?** | no (see [`.gitignore`](.gitignore)) |
| **Where it lands** | `%USERPROFILE%\.claude\skills\exemplar-writer\exemplars\` |

## Naming

**Lead with the kind of document.** That leading phrase is what a request gets
matched against, so it is the part that has to be right:

```
Status Report - Migration Programme.md
Proposal - Managed Service.md
Board Paper - Quarterly Risk Update.md
File Note - Vendor Meeting.md
Terms of Reference - Steering Committee.md
Minutes - Executive Committee.md
Discussion Paper - Cloud Exit Options.md
Position Paper - Records Retention.md
```

Ask for a status report and only the `Status Report - *` files are read. Ask for
a proposal and only the proposals are. Matching is case-insensitive, and obvious
equivalents count as the same kind (`Progress Report` for `Status Report`,
`Statement of Work` for `Proposal`).

## What belongs here

- The **best** real example of each kind of document you produce - one status
  report, one proposal, one file note, one set of minutes. Two or three strong
  ones beat a folder of variations, and a big folder makes the match ambiguous.
- Documents that were **finished and approved**, not drafts. The skill copies
  the habits it finds, including the bad ones.
- Anything **safe to keep in a skill folder**. Strip or fake client names,
  pricing and personal data first - a redacted exemplar teaches shape just as
  well.

A document carrying a classification or handling marking is your call, not the
skill's. It will say the markings are there rather than filing it for you.

## What does not

- **Blank templates and letterhead** - those are what a document is *built
  from*, not read for guidance. They go in `C:\Eva\templates\word` or
  `C:\Eva\templates\powerpoint`, where the `word` and `powerpoint` plugins can
  open them.
- **Briefs** - [`/brief-writer`](../../brief-writer) has its own exemplars
  folder, and its own structure for them.
- **Emails** - same, [`/email-writer`](../../email-writer).
- **Reference data to answer questions from** - that is what the
  [`knowledge-base`](../../../plugins/knowledge-base) plugin indexes, out of
  `C:\Eva\knowledge`. See below for why exemplars deliberately are not indexed.
- **Twelve near-identical reports.** The extras only make the match harder.

## Index

One line per file, saying what it is good for. Fill this in and the skill can
pick the right exemplar without opening all of them.

| File | Kind of document | Use it as the model for |
|---|---|---|
| _(none yet - add yours here)_ | | |

## Formats, and a caveat worth knowing

`.md` and `.txt` are read directly, always work, and are the easiest to review
and diff.

`.docx`, `.pptx` and `.pdf` exemplars need the [`word`](../../../plugins/word),
[`powerpoint`](../../../plugins/powerpoint) or
[`pdf-to-md`](../../../plugins/pdf-to-md) server to read them, and each server
is confined to its own documents folder. This folder sits under
`%USERPROFILE%\.claude\skills\`, which is outside all of them, so a binary
exemplar here may not be readable at all.

Two fixes, either is fine:

1. **Keep a `.md` copy beside the original.** The better arrangement anyway for
   an exemplar you reach for often - no conversion step, and the structure is
   visible at a glance.
2. **Keep the original in the Eva tree** (`C:\Eva\documents\word`,
   `\powerpoint`, `\pdf`) and point the skill at it there in your prompt. That
   folder *is* inside the server's sandbox, and a file you name explicitly beats
   this folder anyway.

## Why none of this is indexed

Exemplars are deliberately kept out of `C:\Eva\knowledge`, the only folder the
RAG index reads. Add a board paper to the indexed corpus and its phrasing comes
back with the same authority as a policy: `kb_ask` starts citing house style as
fact, and a two-year-old example as current practice.

If you want a document to be **both** - a genuine reference as well as a model
to write like - put its content in `C:\Eva\knowledge\notes\` and keep the
formatted copy here. Same document, two jobs, no confusion about which is being
cited.

## How to use them

Name the exemplar, or name the kind of document and let the skill match:

- *"Write the Q3 status report for the migration programme - follow the status
  report exemplar."*
- *"Draft a file note from these meeting notes, same format as the last one."*
- *"Read the proposal exemplar first, then tell me what sections my draft is
  missing."*

Reading them costs a tool call per file, so the skill opens the matched kind
plus one control, not the whole folder.

## A note on style

An exemplar shows what good output looks like. It does not stop the writing
sounding like an AI wrote it - the skill runs
[`/unslop`](../../unslop) over every draft for that, and you can run it again
yourself over anything you have edited since.
