# Agents

Standalone Claude Code subagents — no MCP server of their own, no Python, no
dependencies. Each one is a single Markdown file: YAML frontmatter that says
when Claude should hand work to it, then the instructions it follows.

```
agents/
  <agent-name>.md      ← the agent
  README.md            ← this file
```

That layout is deliberately identical to the one Claude Code expects in
`%USERPROFILE%\.claude\agents\`, so installing an agent is a straight file
copy. Every command below is **PowerShell**.

The other two folders in this repo: [`plugins/`](../plugins) holds MCP servers,
which give Claude new *tools*; [`skills/`](../skills) holds standalone skills,
which give Claude new *instructions*. An agent is the third thing — a separate
context with its own instructions and its own conversation, that the main
session hands a whole job to and gets a finished result back from. The
difference that matters in practice: a skill steers the conversation you are
already in, an agent goes away and does the work in its own.

## Available agents

| Agent | Invoke | What it does |
|---|---|---|
| [**researcher**](researcher.md) | `@agent-researcher` | Researches a topic across the local knowledge base and Confluence, corroborates what it finds, and returns a cited brief with confidence ratings and named gaps |
| [**report-writer**](report-writer.md) | `@agent-report-writer` | Turns research or notes into the written content of a report or official brief, following the structure of an exemplar in `./context/exemplars`, then runs `/unslop` and `/polish` over it |

They are built to run back to back: `researcher` produces the cited brief,
`report-writer` writes it up. Neither needs the other, though —
`report-writer` will work from any material you give it.

## Install

Copy the agent's file into your Claude agents directory. From the root of this
repo:

```powershell
$dest = "$env:USERPROFILE\.claude\agents"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force .\agents\researcher.md $dest
Copy-Item -Force .\agents\report-writer.md $dest
```

`New-Item -Force` is there so the first install works before `~\.claude\agents`
exists.

**For one project only** — copy it to `.claude\agents\` inside that project
instead:

```powershell
$dest = ".\.claude\agents"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force C:\path\to\claude-skills\agents\report-writer.md $dest
```

Useful when the agent should travel with the code, or when you want to edit it
for that project without touching the global copy. A project agent wins over a
user one of the same name.

Run `/agents` to confirm Claude Code picked them up, or `/doctor` if one
doesn't appear.

To update, run the same `Copy-Item` again — `-Force` overwrites. To remove one:

```powershell
Remove-Item -Force "$env:USERPROFILE\.claude\agents\researcher.md"
```

## Use

Claude delegates on its own when a request matches an agent's `description`:

```
research our records retention obligations for contractor emails
write this up as a brief for the deputy secretary
```

To force it, `@`-mention the agent:

```
@agent-researcher what does the wiki say about the change freeze?
@agent-report-writer turn the brief above into a Word document
```

Each agent runs in its own context. It sees the job you hand it, not your whole
conversation, so give it the material rather than pointing back at something
you said earlier.

## What they build on

These are written for this suite, not as generic agents. They assume the rest
of it is installed and configured, and they don't carry fallback paths for a
server that isn't there.

| Agent | Assumes | Optional |
|---|---|---|
| `researcher` | [`knowledge-base`](../plugins/knowledge-base) and [`confluence`](../plugins/confluence) plugins | Files you point it at |
| `report-writer` | [`unslop`](../skills/unslop) and [`polish`](../skills/polish) skills | An exemplar in [`context/exemplars`](../context/exemplars); [`word`](../plugins/word) or [`pdf-to-md`](../plugins/pdf-to-md) to read a `.docx` or PDF one |

`researcher` searches both sources on every question — the knowledge base holds
the settled documents, Confluence holds the working knowledge that never became
one. What it will not do is fill a gap from its own memory: on an airgapped
network, a search that comes back empty is a finding, and it reports the
absence and where it looked.

`report-writer` writes content and returns Markdown. It does not produce Word
documents. Hand the finished Markdown to [`word`](../plugins/word) for that,
which keeps writing and formatting as two jobs you can redo independently — a
wording change shouldn't mean rebuilding a `.docx`.

### Exemplars

`report-writer` reads [`context/exemplars`](../context/exemplars) for the shape
of the document it is writing. The exemplar supplies **structure and emphasis —
never content**: section order and headings, how long each section runs, what
that audience always needs to see. Facts, figures and names come only from the
material you hand it.

It is worth setting up. An exemplar tells the agent what your readers actually
expect far better than any generic template, and two or three strong ones — a
brief, a report, a minute — cover most of what gets asked for.

Fill in the **index table** in
[`context/exemplars/README.md`](../context/exemplars/README.md). That one line
per file is what lets the agent pick the right exemplar without opening every
document, which matters because a `.docx` or PDF costs a conversion to read.
`.md` exemplars are the cheapest and the easiest to review, so a Markdown copy
beside the original earns its keep for anything you reach for often.

The [`templates`](../context/templates) folder next door is a different thing —
blank files a document is *built from*, which is the `word` server's business.
`report-writer` will say so rather than treat one as an exemplar.

## Adding an agent

1. Create `agents/<name>.md` with YAML frontmatter:
   ```yaml
   ---
   name: <name>
   description: What it does, and when Claude should hand work to it.
   ---
   ```
   `name` must be lowercase with hyphens, and must match the filename. The
   `description` is the only part Claude reads when deciding whether to
   delegate, so write it to trigger on the right requests and, just as
   importantly, *not* on the wrong ones.
2. Write the instructions below the frontmatter. They are the agent's whole
   system prompt — it has no other context.
3. Add a row to the table above.

Neither agent here sets `tools:`, so both inherit everything the session has.
That is deliberate: naming MCP tools explicitly means an agent silently loses
access when a server is registered under a different name, which is a worse
failure than having a tool it doesn't use. Add `model:` if you want to pin one;
without it, an agent inherits the session's model.

Agents here are unversioned — they are prose, not an interface anything else
depends on. (The MCP servers under `plugins/` are versioned; see the root
[README](../README.md#versioning).)
