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
| [**researcher**](researcher.md) | `@agent-researcher` | Researches a topic across the local knowledge base and Confluence, corroborates what it finds, and returns a cited brief with confidence ratings and named gaps — then offers to capture the brief back into the knowledge base |

`researcher` produces the material; writing it up is a skill's job, not an
agent's. Hand its brief to [`/brief-writer`](../skills/brief-writer) for a paper
going to an executive, [`/exemplar-writer`](../skills/exemplar-writer) for a
report or paper shaped like one you already have, or
[`/email-writer`](../skills/email-writer) for a note
in your own voice. Both keep you in the conversation to iterate, which is where
writing actually gets done.

## Install

Copy the agent's file into your Claude agents directory. From the root of this
repo:

```powershell
$dest = "$env:USERPROFILE\.claude\agents"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force .\agents\researcher.md $dest
```

`New-Item -Force` is there so the first install works before `~\.claude\agents`
exists.

**For one project only** — copy it to `.claude\agents\` inside that project
instead:

```powershell
$dest = ".\.claude\agents"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force C:\path\to\claude-skills\agents\researcher.md $dest
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
```

To force it, `@`-mention the agent:

```
@agent-researcher what does the wiki say about the change freeze?
```

Each agent runs in its own context. It sees the job you hand it, not your whole
conversation, so give it the material rather than pointing back at something
you said earlier.

## What it builds on

`researcher` is written for this suite, not as a generic agent. It assumes the
rest of it is installed and configured, and it doesn't carry fallback paths for
a server that isn't there.

| Agent | Assumes | Optional |
|---|---|---|
| `researcher` | [`knowledge-base`](../plugins/knowledge-base) and [`confluence`](../plugins/confluence) plugins | Files you point it at |

It searches both sources on every question — the knowledge base holds the
settled documents, Confluence holds the working knowledge that never became one.
What it will not do is fill a gap from its own memory: on an airgapped network,
a search that comes back empty is a finding, and it reports the absence and
where it looked.

### Capturing what it produces

`researcher` finishes by **offering** to put its brief into the knowledge base
with `kb_capture`, as `Research - <topic>`. It asks; it doesn't file things on
its own. Say yes and the next person to ask that question finds the work instead
of redoing it.

One thing to watch. A captured note lands in the same index as your real policy
documents and comes back from the same searches, so the agent stamps its own as
agent-written and is told to treat any such file as a **prior brief rather than
a source** — it mines it for leads, then verifies against the documents it
cites. Without that rule a guess captured in March becomes a citation in June.

### Writing it up

The brief is research, not a document. What it becomes depends on the audience:
[`/brief-writer`](../skills/brief-writer) for a paper going to an executive,
[`/email-writer`](../skills/email-writer) for a note in your own voice, or
[`/exemplar-writer`](../skills/exemplar-writer) for anything else that should
follow the shape of a document you already have. All three are skills rather
than agents on purpose — a draft you iterate on in the conversation, and each
has its own exemplars folder for the shape and the voice to follow.

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

The agent here does not set `tools:`, so it inherits everything the session has.
That is deliberate: naming MCP tools explicitly means an agent silently loses
access when a server is registered under a different name, which is a worse
failure than having a tool it doesn't use. Add `model:` if you want to pin one;
without it, an agent inherits the session's model.

Agents here are unversioned — they are prose, not an interface anything else
depends on. (The MCP servers under `plugins/` are versioned; see the root
[README](../README.md#versioning).)
