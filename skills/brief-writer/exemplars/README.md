# Exemplar briefs

Finished briefs that were **approved** - the ones you would hand a new starter
and say *"write it like this"*. [`/brief-writer`](../SKILL.md) reads them for
structure, section order, proportions and register, then writes something new
in that shape.

An exemplar is read for guidance and **never becomes the output**. No fact,
figure, date, name or dollar amount is ever carried across from one.

| | |
|---|---|
| **Formats** | `.md`, `.txt`, `.pdf`, `.docx` |
| **Committed to git?** | no (see [`.gitignore`](.gitignore)) |
| **Where it lands** | `%USERPROFILE%\.claude\skills\brief-writer\exemplars\` |

## Naming

**Lead with the kind of brief.** That is what the skill matches against, so it
is the part that has to be right:

```
Decision Brief - Licence Renewal.md
Decision Brief - Regional Office Closure.pdf
Noting Brief - Q3 Programme Status.docx
Ministerial Brief - Senate Estimates Hearing.docx
```

Ask for a decision brief and only the `Decision Brief - *` files are read. Ask
for a noting brief and only the noting ones are. Matching is
case-insensitive, and obvious equivalents count as the same kind
(`Brief for Noting`, `Information Brief`).

## What belongs here

- The **best** real example of each kind of brief you write. Two or three strong
  ones beat a folder of variations, and a big folder makes the match ambiguous.
- Briefs that were **approved**, not drafts. The skill copies the habits it
  finds, including the bad ones.
- Anything **safe to keep in a skill folder**. Strip or fake names, figures and
  personal data first - a redacted brief teaches structure just as well.

A brief carrying a classification or handling marking is your call, not the
skill's. It will say the markings are there rather than filing it for you.

## Index

One line per file, saying what it is good for. Fill this in and the skill can
pick the right exemplar without opening all of them.

| File | Kind | Use it as the model for |
|---|---|---|
| _(none yet - add yours here)_ | | |

## Formats, and a caveat worth knowing

`.md` is the cheapest and the easiest to review, and it always works.

`.docx` and `.pdf` exemplars need the [`word`](../../../plugins/word) or
[`pdf-to-md`](../../../plugins/pdf-to-md) server to read them, and both servers
are confined to their configured folders. This folder sits under
`%USERPROFILE%\.claude\skills\`, which is usually outside those, so a `.docx`
here may not be readable. The fix is a `.md` copy kept beside the original,
which is the better arrangement anyway for an exemplar you reach for often.

## Not to be confused with

[`/exemplar-writer`](../../exemplar-writer)'s own
[`exemplars/`](../../exemplar-writer/exemplars) folder, which holds the
general-purpose ones: status reports, proposals, file notes, minutes. This
folder is the brief-writer skill's own, and only briefs belong in it - the two
skills never read each other's.

Either way, a file you name in the prompt beats the folder, including one in
`C:\Eva\documents\word` (which, unlike a skill folder, the `word` server can
actually open).
