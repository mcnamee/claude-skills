# brief-writer

Draft a **brief**: the formal paper that goes up to a senior executive, either
to get a decision approved or to have a way forward noted.

It matches your request to an exemplar in [`exemplars/`](exemplars), follows
that exemplar's structure, writes in a senior executive register, finishes with
[`/polish`](../polish) for Australian Public Service style, and returns the
Markdown in the chat so you can iterate on it.

A standalone skill: no MCP server, no Python, no dependencies, so it works
anywhere including on an airgapped machine. It uses [`/polish`](../polish),
which should be installed alongside it.

## Install

Copy this folder into your Claude skills directory. From the root of this repo,
in **PowerShell**:

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\brief-writer $dest
Copy-Item -Recurse -Force .\skills\polish $dest
```

The `exemplars\` folder travels with it, so put your exemplars in before you
copy and they land on the endpoint with the skill. For one project only, copy it
to `.claude\skills\brief-writer\` inside that project instead. Run `/doctor` or
restart Claude Code if it doesn't show up.

See [`skills/README.md`](../README.md) for the general install notes.

## Use

```
/brief-writer draft a decision brief on the licence renewal
/brief-writer a noting brief for the deputy secretary on the Q3 result
/brief-writer                (with no argument: the material above in the conversation)
```

It settles three things first, inferring what it can and asking about the rest
in one round: whether it is for a **decision** or for **noting**, who it is
going to, and what exactly is being decided. It never blocks on a question, so
skipping them means it states its assumption and writes.

## Decision or noting

Everything follows from which one it is:

| Kind | The executive is being asked to | Recommendation reads |
|---|---|---|
| **Decision brief** | Approve, agree, choose between options, sign | `That you agree to ...` |
| **Noting brief** | Note a way forward, an outcome or an emerging issue | `That you note ...` |

A decision brief that turns out to be for noting is the wrong document, not a
wordy one, which is why it is the first thing settled.

## Exemplars

[`exemplars/`](exemplars) is the skill's own folder, and the point of it is that
it travels with the skill. Fill it on a machine where you have your documents,
copy the folder across, and the skill arrives on the endpoint already knowing
what your briefs look like.

Name each file with the kind of brief first:

```
Decision Brief - Licence Renewal.md
Noting Brief - Q3 Programme Status.docx
Ministerial Brief - Senate Estimates Hearing.pdf
```

Ask for a decision brief and only the `Decision Brief - *` files are read. `.md`
is cheapest and always readable; `.docx` and `.pdf` need the
[`word`](../../plugins/word) or [`pdf-to-md`](../../plugins/pdf-to-md) server
and may sit outside its sandbox, so a `.md` copy beside the original is worth
keeping. Nothing in the folder is committed to git except its README, so your
real briefs stay yours.

**It follows the exemplar's structure unless you say otherwise.** Section order,
headings as worded, the proportions between sections, the standing blocks a
recommendation and a signature go in. What it never takes is content: no figure,
date, name or dollar amount crosses from an exemplar into a new brief.

With an empty folder it says so and falls back to a standard structure -
recommendation, purpose, background, key issues, options, financials, risks,
consultation, sensitivities, attachments, contact officer - then offers to keep
the approved brief as your first exemplar.

## What you get back

The brief as Markdown in the chat, then one line naming the exemplar it
followed, then flags: sections the material couldn't fill, figures that
contradicted each other, an acronym it couldn't expand, an assumption it had to
make. Sections it can't fill keep their heading with the gap marked in square
brackets, because a brief with a visible hole is fixable and one with a
plausible invention in it isn't.

Then you iterate. Each round returns the whole brief again with the structure
held stable, and only the changed text goes back through `/polish`.

## Turning it into a Word document

Ask, once you're happy with it, and the finished Markdown goes to the
[`/word:word`](../../plugins/word) skill, which builds the `.docx` with native
Word styles or from a blank in `C:\Eva\templates\word`. Writing and
formatting stay two jobs on purpose: a wording change shouldn't mean rebuilding
the document.

## With the other writing skills

[`/polish`](../polish) runs automatically as the last step, told the reader and
medium up front so it doesn't ask you again, and the structure is re-checked
afterwards because a polish pass will happily merge two sections your exemplar
keeps apart.

[`/unslop`](../unslop) is offered rather than run. It's worth taking when the
source material was itself AI-drafted, and it goes **before** `/polish` -
polishing first means applying style rules to padding you're about to delete.

For an email in your own voice rather than a formal paper, that's
[`/email-writer`](../email-writer), which works the same way from its own
exemplars folder.
