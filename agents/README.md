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
| [**report-writer**](report-writer.md) | `@agent-report-writer` | Turns research or notes into a finished report or official brief, following the structure of an exemplar in `./context/exemplars`, then runs `/unslop` and `/polish` over it |

They are built to run back to back: `researcher` produces the cited brief,
`report-writer` turns it into the document. Neither needs the other, though —
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

## What they expect to find

Neither agent requires anything, but both are much more useful with the rest of
the suite installed. Each one degrades honestly — it says what is missing
rather than making something up to fill the space.

| Agent | Wants | Falls back to |
|---|---|---|
| `researcher` | [`knowledge-base`](../plugins/knowledge-base) and [`confluence`](../plugins/confluence) plugins | Saying which source it couldn't reach, and working from files you point it at |
| `report-writer` | [`unslop`](../skills/unslop) and [`polish`](../skills/polish) skills, [`word`](../plugins/word) plugin for `.docx` output, an exemplar in `./context/exemplars` | Markdown output, a standard brief or report structure, and its own plain-language pass |

### `./context/exemplars`

`report-writer` looks here, relative to the folder you are working in, for a
previous report to copy the shape of. The exemplar supplies **structure and
emphasis — never content**: section order and headings, how long each section
runs, what the audience always needs to see. Facts, figures and names come only
from the material you hand it.

It is worth setting up. An exemplar tells the agent what your readers actually
expect far better than any generic template, and a folder that holds one brief,
one report and one minute covers most of what gets asked for.

```
your-project/
  context/
    exemplars/
      brief-to-deputy-secretary.md
      quarterly-report.docx
```

Markdown is the easiest format to read. A `.docx` exemplar works too, but the
`word` server can only open files inside its configured docs folder — keep it
there, or keep a Markdown copy alongside.

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
