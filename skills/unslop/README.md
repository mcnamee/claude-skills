# unslop

Strip the markers that make writing read as machine-generated — padding
phrases, tell-tale vocabulary, and the stock LLM sentence shapes — **without
changing the meaning or the author's voice**.

A standalone skill: no MCP server, no Python, no dependencies and no
configuration, so it works anywhere — including on an airgapped machine.

## Install

Copy this folder into your Claude skills directory. From the root of this
repo, in **PowerShell**:

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\unslop $dest
```

For one project only, copy it to `.claude\skills\unslop\` inside that project
instead. Run `/doctor` or restart Claude Code if it doesn't show up.

See [`skills/README.md`](../README.md) for the general install notes.

## Use

```
/unslop <paste the text>
/unslop C:\path\to\draft.md
/unslop                    (with no argument: the last substantial piece of writing in the conversation)
```

It returns the cleaned text and nothing else — no preamble, no commentary.
Ask "what changed?" and you get a before/after table with a reason per edit.

## What it removes

Four tiers, applied with decreasing confidence:

1. **Delete outright** — padding that carries no information: *"It's important
   to note that…"*, *"In today's fast-paced world"*, *"Great question!"*,
   *"In conclusion,"*, and closing paragraphs that only restate the piece.
2. **Swap for the plain word** — *delve into* → *look at*, *leverage* → *use*,
   *robust* → *reliable*, *myriad* → *many*, *a testament to* → *shows*.
3. **Restructure** — the stock shapes: the label-colon list (below),
   *"It's not just X — it's Y"*, participial tails (*", ensuring
   reliability"*), rhetorical fragments (*"The result? Faster builds."*),
   empty triads, *"Here's the thing:"*, stacked hedges, and both-sides
   non-endings.
4. **Signals, not targets** — em dashes, bullet lists, bold, raw `<strong>`
   tags, emoji, and *"However"/"Moreover"*. These are ordinary writing as
   often as they are slop, so it only acts when they are clearly mechanical,
   and varies rather than purges.

### The label-colon list

The most recognisable of the structural tells, and the one that survives
longest in business writing:

```
The key benefits include:
- **Better alignment:** Your staff will thrive.
- **Cross organisational synergy:** Making use of each other's strengths.
- **De-duplication of effort:** Ensuring best athlete approach.
```

A stem that praises instead of introducing, labels coined so that there is
something to bold, fragments for bodies, and three of them. The skill's test
is to cover the label and see what is lost: if it is a term you would scan or
search for — a parameter, an option, a product, a date, a step — the list is
doing real work and is left alone. If it only nominalises the sentence beside
it, the label goes and the claim stays:

```
- Staff will thrive
- Teams make use of each other's strengths
- Effort is not duplicated
```

The claims come out as vague as they went in. Stripping the packaging is the
edit; sharpening the content is not, and the skill will not invent a specific
to fill a label's place.

## What it will not do

This is a **subtractive** edit. It doesn't improve the writing, tighten the
argument, restructure the document or make it punchier — those are different
jobs, and you have to ask for them separately.

It never touches facts, numbers, names, citations, quotations (verbatim, even
when the quote is itself sloppy), code, commands, file paths, log output, the
author's opinions and jokes, or your spelling conventions. Terms of art that
look like slop are left alone — *robust* in statistics, *leverage* in finance,
*navigate* about actual navigation.

Two rules keep it honest:

- **When it can't tell whether something is slop or your voice, it leaves it.**
  One surviving marker costs less than a sentence that no longer sounds like
  you. A word you use consistently throughout a piece is treated as voice.
- **If it changed more than about a third of the words, it has to justify each
  edit against one of the four tiers.** Genuinely slop-dense text can warrant
  that much change, but so does an unasked-for rewrite — anything it can't
  point at a rule for gets reverted.

Before returning, it checks that every claim survived, that no claim was
added, that numbers and quotes are byte-identical, and that the word count
didn't grow (if it grew, it rewrote instead of removing).

If your text is already clean it says so and hands it back unchanged, rather
than manufacturing edits to look useful.
