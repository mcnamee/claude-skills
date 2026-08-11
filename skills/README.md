# Skills

Standalone Claude skills — no MCP server, no Python, no dependencies. Each one
is a folder holding a `SKILL.md` (the instructions Claude follows) and a
`README.md` (what it does, for you).

```
skills/
  <skill-name>/
    SKILL.md      ← the skill itself
    README.md     ← human-facing docs
```

That layout is deliberately identical to the one Claude Code expects in
`~/.claude/skills/`, so installing a skill is a straight folder copy.

The [`plugins/`](../plugins) folder is the other half of this repo: MCP servers
that give Claude new *tools*. Skills here give Claude new *instructions* — they
need nothing installed and work anywhere, including offline.

## Available skills

| Skill | Invoke | What it does |
|---|---|---|
| [**unslop**](unslop) | `/unslop` | Strips AI-slop markers from writing — padding, tell-tale vocabulary, stock LLM sentence shapes — leaving meaning and voice intact |

## Install

Copy the skill's folder into your Claude skills directory.

**For you, in every project** (Windows):

```
xcopy /E /I skills\unslop %USERPROFILE%\.claude\skills\unslop
```

macOS/Linux:

```
cp -r skills/unslop ~/.claude/skills/unslop
```

**For one project only** — copy it to `.claude/skills/<name>/` in that
project's folder instead. Useful when the skill encodes something
project-specific, or when you want it committed alongside the code.

Then invoke it by name: `/unslop`. Claude also applies a skill automatically
when the task matches its `description`, so an explicit slash command isn't
always necessary — though `unslop` is deliberately written to wait until you
ask.

If a newly copied skill doesn't appear, run `/doctor` or restart Claude Code.

To update one, copy the folder over the top again. To remove it, delete the
folder from `~/.claude/skills/`.

## Adding a skill

1. `mkdir skills/<name>` and write `SKILL.md` with YAML frontmatter:
   ```yaml
   ---
   name: <name>
   description: What it does, and when Claude should use it.
   ---
   ```
   `name` must match the folder name — that pairing is what makes the slash
   command work. The `description` is the only part Claude reads when deciding
   whether a skill is relevant, so write it to trigger on the right requests
   and, just as importantly, *not* on the wrong ones.
2. Add a `README.md` beside it for the human-facing explanation.
3. Add a row to the table above.

Skills here are unversioned — they are prose, not an interface anything else
depends on, so there is no version to keep in sync. (The MCP servers under
`plugins/` are versioned; see the root [README](../README.md#versioning).)
