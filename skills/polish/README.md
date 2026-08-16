# polish

Rewrite a draft into **Australian Public Service writing style** — the
[Australian Government Style Manual](https://www.stylemanual.gov.au) and the APSC
[Government writing handbook](https://www.stylemanual.gov.au/style-manual-resources/government-writing-handbook).

Plain language, the reader's need first, active voice, Australian spelling, and the
APS conventions for numbers, dates, capitals, lists and links. The handbook is the
craft half — readers, tone, argument, structure — and the Style Manual is the
mechanics.

A standalone skill: no MCP server, no Python, no dependencies and no
configuration, so it works anywhere — including on an airgapped machine.

## Install

Copy this folder into your Claude skills directory. From the root of this
repo, in **PowerShell**:

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\polish $dest
```

For one project only, copy it to `.claude\skills\polish\` inside that project
instead. Run `/doctor` or restart Claude Code if it doesn't show up.

See [`skills/README.md`](../README.md) for the general install notes.

## Use

```
/polish <paste the draft>
/polish C:\path\to\brief.docx.md
/polish                           (with no argument: the last substantial draft in the conversation)
/polish check <text>              (report only — nothing rewritten)
/polish rewrite this as an email to an external stakeholder
```

It works out the reader, the medium and the purpose from the draft where it can.
Where it can't, it asks — once, up to three questions, with a recommended default
on each so you can skip them. Say "just polish it" and it picks the closest
profile, tells you which, and gets on with it.

## Tone, voice and format profiles

APS style is not one register, and picking the wrong one is the most common reason
a technically correct rewrite still reads wrong. The handbook sets three levels of
formality — **formal** (policies, reports, ministerial and Cabinet briefs; no
contractions or personal pronouns), **standard** (the usual tone for APS writing:
emails, letters, online services, media releases), and **informal** (social media,
blogs). Underneath sits voice: basic government voice is clear, direct, objective
and impartial, shaded to the job — supportive for guidance, expert for research
reports, measured for policy, calm for emergency instructions.

The medium then sets the shape. The skill applies one of eight profiles:

| Profile | Opens with | Paragraphs |
|---|---|---|
| Internal email | The ask or the answer | 1–3 sentences |
| External stakeholder email | Why you're writing, what you need | 2–3 sentences |
| Letter to a member of the public | What this means for them | 1–3 sentences |
| Brief or minute | The recommendation | 2–4 sentences |
| Report or report section | A summary of findings | Up to 6 sentences |
| Web or intranet page | What the reader can do | 2–3 sentences |
| Media release | Who, what, when, where, why | 1–2 sentences |
| Talking points | The bottom line, in one sentence | 1–2 sentences |

The paragraph lengths are the handbook's own, and the profile also sets person
(`you` on a web page, `the department` in a brief), tone, and format-specific
structure — review rights in a decision letter, `Summary` rather than
`Executive summary` in a report, `For action` / `For decision` labels on an
internal email, a media contact block in a release.

## What it changes

Structure first, mechanics last, because commas in a paragraph that's about to
move are wasted work.

1. **Structure** — summary first, then detail. Headings front-loaded, in sentence
   case, under 70 characters and not phrased as questions. Action and due date out
   of the last line. Detail the reader doesn't need moves to an attachment.
2. **Argument** — if the piece asks for something, it needs facts and logic, the
   effect on actual people, and demonstrated credibility. Drafts usually arrive
   with only the first, so the skill flags what's missing.
3. **Sentences** — 15 words on average, 25 as the hard ceiling, active voice, one
   idea each. `It was decided that` becomes `The delegate decided`. Negatives turn
   positive: `We can't send updates if you don't enter your email address` becomes
   `Enter your email address to receive updates`.
4. **Words** — `utilise` → `use`, `prior to` → `before`, `conduct an assessment
   of` → `assess`, `in order to` → `to`. Jargon, corporate filler and noun stacks
   go. `A number of agencies` becomes the number, or gets flagged. Plus the APS
   grammar traps: `the department is`, not `are`.
5. **Mechanics** — Australian spelling (`-ise`, `program`, `licence`/`license`),
   minimal capitals (`the Australian Government` but `the government`), numerals
   for 2 and above with `zero` and `one` in words, `3 March 2026`, `11 am`, `25%`,
   `$50`, no Oxford comma, spaced en dashes rather than em dashes, single quotes,
   the handbook's three list-punctuation patterns, and descriptive link text.
6. **Inclusive and accessible language** — person-first, gender-neutral,
   `older Australians` not `old people`, `they` not `he/she`, and no idiom that
   fails in translation.

## What it won't touch

Facts, figures, dates, dollar amounts and names. Direct quotations, even bad ones.
Legislation and section references. Security classification markings — `OFFICIAL`,
`OFFICIAL: Sensitive`, `PROTECTED` and any caveats stay verbatim and in place.
Defined terms and terms of art.

And it won't change the **force** of a sentence. `must` is not `should`, `may` is
not `will`, and deadlines, entitlements and review rights keep their exact
strength. That rule outranks every style rule in the skill.

Where the draft is missing something the format needs — a recommendation, a due
date, review rights on a decision, a source for a claim — it flags the gap. It
never fills it in.

## Modes

- **rewrite** (default) — the rewritten text, then up to six bullets on what
  changed, then flags.
- **check** — a compliance report with measures (average sentence length,
  sentences over 25 words, passive constructions, undefined acronyms, over-long
  headings), a scan for your writing tics, and findings ordered by severity.
  Nothing is rewritten. Use it on anything already sent.
- **edit** — point it at a file and it edits in place, leaving compliant passages
  alone.

## House style

Agency house styles override the Style Manual wherever they differ. Name yours, or
point the skill at the file, and it follows that instead — except for the two
rules that never bend: don't change a legal obligation to satisfy a style rule,
and don't change a quotation.

If you don't have a house style yet, the flags do double duty: the skill lists the
judgement calls it had to make repeatedly — which spelling, which acronym
expansion, one word or two — which is the beginning of the **word list** the
handbook recommends every writer keep.

## With /unslop

[`/unslop`](../unslop) is subtractive — it strips AI-writing markers and leaves
voice and structure alone. `/polish` converts to a house style and will
restructure freely.

On AI-drafted text, run `/unslop` first, then `/polish`. The other order means
polishing padding you're about to delete.
