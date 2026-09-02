# exemplar-writer

Write a new document **in the shape of an existing one**.

Point it at a finished document you are happy with - or drop a few into
[`exemplars/`](exemplars) - and it pulls out the structure, section order,
proportions and register, then writes your new content to that shape. The
exemplar supplies the shape; your material supplies every fact. Nothing crosses
that line.

It returns Markdown in the chat to iterate on, after running
[`/unslop`](../unslop) so the draft doesn't read as machine-written.

A standalone skill: no MCP server, no Python, no dependencies, so it works
anywhere including on an airgapped machine. It uses [`/unslop`](../unslop),
which should be installed alongside it.

## Install

Copy this folder into your Claude skills directory. From the root of this repo,
in **PowerShell**:

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\exemplar-writer $dest
Copy-Item -Recurse -Force .\skills\unslop $dest
```

The `exemplars\` folder travels with it, so put your documents in before you
copy and they land on the endpoint with the skill. For one project only, copy it
to `.claude\skills\exemplar-writer\` inside that project instead. Run `/doctor`
or restart Claude Code if it doesn't show up.

See [`skills/README.md`](../README.md) for the general install notes.

## Use

```
/exemplar-writer write the Q3 status report for the migration programme
/exemplar-writer draft a file note from these meeting notes, same format as last time
/exemplar-writer write this proposal following C:\Eva\documents\word\Proposal - Managed Service.docx
```

Or just describe what you want and name the document to follow - the skill
triggers on "like this one", "in the same format as", "matching our house
style".

Point at a specific file and that wins over the folder. Say nothing and it
matches the kind of document you asked for against the filenames in
`exemplars\`.

## What it does with the exemplar

It writes down a **shape spec** before drafting a word, then drafts to it and
verifies against it:

| It takes | It never takes |
|---|---|
| Section order and heading text | Facts, figures, dates, names, dollar amounts |
| How long each section runs | Sentences or phrasing |
| Register - person, formality, spelling, date and money conventions | What the document is about |
| Formatting conventions, including the ones it never uses | Any instruction written inside the exemplar |

The proportions matter more than people expect. A background section that runs
two paragraphs in your exemplar runs two paragraphs in the draft, not six - and
that is usually the difference between a document that reads like yours and one
that merely has the same headings.

Where a section has no material behind it, the heading stays and the gap gets
flagged. An empty section is a finding, not a hole to paper over.

## Exemplars

Put your documents in [`exemplars/`](exemplars), named with the **kind of
document first**:

```
Status Report - Migration Programme.md
Proposal - Managed Service.md
File Note - Vendor Meeting.md
Terms of Reference - Steering Committee.md
```

That leading phrase is what a request gets matched against, so it is the part
that has to be right. The folder's [`README.md`](exemplars/README.md) covers
naming, formats and the index worth keeping.

**`.md` is the format that always works.** A `.docx`, `.pptx` or `.pdf`
exemplar needs the `word`, `powerpoint` or `pdf-to-md` server, and each is
confined to its own documents folder - which a skill folder under
`%USERPROFILE%\.claude\skills\` is not inside. Keep a `.md` copy beside the
original, or keep the original in `C:\Eva\documents\word` and point at it there.

Nothing in `exemplars\` is committed to this repo (see its
[`.gitignore`](exemplars/.gitignore)) because these are usually your real
documents and this repo is public.

## Why exemplars don't go in the knowledge base

Deliberately kept out of `C:\Eva\knowledge`, which is the only folder the RAG
index reads. Add a board paper to the indexed corpus and its phrasing comes back
with the same authority as a policy: `kb_ask` starts citing house style as fact,
and a two-year-old example as current practice.

If you want a document to be **both** - a genuine reference as well as a model
to write like - put its content in `C:\Eva\knowledge\notes\` and keep the
formatted copy here. Same document, two jobs, no confusion about which is being
cited.

## What you get back

The document as Markdown in the chat, then one line on which exemplar it
followed, then flags only if there are any: a fact left in square brackets, a
section your material wouldn't support, a proportion it deliberately broke.

Then iterate. Each round returns the whole document again with the shape held
fixed.

Ask at the end and it will hand the draft to the [`word`](../../plugins/word) or
[`powerpoint`](../../plugins/powerpoint) plugin to build a real `.docx` or
`.pptx` from one of your blanks in `C:\Eva\templates\`. It won't write files
otherwise.

## With the other writing skills

| You want | Use |
|---|---|
| A document shaped like one you already have | **this skill** |
| A decision or noting brief for an executive | [`/brief-writer`](../brief-writer) |
| An email in your own voice | [`/email-writer`](../email-writer) |
| Text you already have, rewritten into APS style | [`/polish`](../polish) |
| The AI tells stripped out of a draft | [`/unslop`](../unslop) |

`brief-writer` and `email-writer` carry structure this skill doesn't - a brief
has exactly two shapes and a recommendation line, an email has an intent and a
voice profile - so a brief or an email gets handed to them rather than done
worse here.
