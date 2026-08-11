# unslop

Strip the markers that make writing read as machine-generated — padding
phrases, tell-tale vocabulary, and the stock LLM sentence shapes — **without
changing the meaning or the author's voice**.

| | |
|---|---|
| **Ships** | a skill only — no MCP server |
| **Version** | 1.0.0 |
| **pip install** | _none_ |
| **Platform** | any |
| **Writes to disk** | only if you ask it to edit a file in place |

The only plugin here with no Python behind it: it needs no server, no
dependencies and no configuration, so it works unchanged on an airgapped
machine.

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install unslop@mcnamee-claude-skills
```

Then invoke it as **`/unslop:unslop`** — Claude Code namespaces a plugin's
skills by the plugin name.

If you want the shorter bare `/unslop`, install the skill on its own instead
of as a plugin:

```
xcopy /E /I plugins\unslop\skills\unslop %USERPROFILE%\.claude\skills\unslop
```

Nothing else changes — same skill, same behaviour, just a different name to
type. Run `/doctor` or restart Claude Code if a newly copied skill doesn't
show up.

## Use

```
/unslop:unslop <paste the text>
/unslop:unslop C:\path\to\draft.md
/unslop:unslop            (with no argument: the last substantial piece of writing in the conversation)
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
3. **Restructure** — the stock sentence shapes: *"It's not just X — it's Y"*,
   participial tails (*", ensuring reliability"*), rhetorical fragments
   (*"The result? Faster builds."*), empty triads, *"Here's the thing:"*,
   stacked hedges, and both-sides non-endings.
4. **Signals, not targets** — em dashes, bullet lists, bold, emoji, and
   *"However"/"Moreover"*. These are ordinary writing as often as they are
   slop, so it only acts when they are clearly mechanical, and varies rather
   than purges.

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
