# Skills

Standalone Claude skills — no MCP server, no Python, no dependencies. Each one
is a folder holding a `SKILL.md` (the instructions Claude follows) and a
`README.md` (what it does, for you).

```
skills/
  <skill-name>/
    SKILL.md      ← the skill itself
    README.md     ← human-facing docs
    exemplars/    ← optional: your own documents for the skill to learn from
```

That layout is deliberately identical to the one Claude Code expects in
`%USERPROFILE%\.claude\skills\`, so installing a skill is a straight folder
copy. Every command below is **PowerShell**.

The [`plugins/`](../plugins) folder is the other half of this repo: MCP servers
that give Claude new *tools*. Skills here give Claude new *instructions* — they
need nothing installed and work anywhere, including offline.

## Available skills

| Skill | Invoke | What it does |
|---|---|---|
| [**brief-writer**](brief-writer) | `/brief-writer` | Drafts a decision or noting brief for a senior executive, following the structure of an exemplar in its own `exemplars/` folder, and finishing with `/polish` |
| [**email-writer**](email-writer) | `/email-writer` | Drafts an email in your voice, classifying what the email is for and matching that intent to your own sent mail in its `exemplars/` folder, then running `/unslop` |
| [**polish**](polish) | `/polish` | Rewrites a draft into Australian Public Service style — the Australian Government Style Manual — asking who the reader is and what the medium is, then picking the register from them |
| [**unslop**](unslop) | `/unslop` | Strips AI-slop markers from writing — padding, tell-tale vocabulary, stock LLM sentence shapes — leaving meaning and voice intact |

`brief-writer` and `email-writer` build on the other two: `brief-writer` runs
`/polish` as its last step, `email-writer` runs `/unslop`. Install the pair each
one needs, not just the writer.

## Install

Copy the skill's folder into your Claude skills directory. From the root of
this repo, in **PowerShell**:

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\unslop $dest
```

`New-Item -Force` is there so the first install works before
`~\.claude\skills` exists; `Copy-Item -Recurse` into the parent folder lands
the skill at `%USERPROFILE%\.claude\skills\unslop`.

**For one project only** — copy it to `.claude\skills\<name>\` inside that
project instead:

```powershell
$dest = ".\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force C:\path\to\claude-skills\skills\unslop $dest
```

Useful when the skill encodes something project-specific, or when you want it
committed alongside the code.

Then invoke it by name: `/unslop`. Claude also applies a skill automatically
when the task matches its `description`, so an explicit slash command isn't
always necessary — though `unslop` is deliberately written to wait until you
ask.

If a newly copied skill doesn't appear, run `/doctor` or restart Claude Code.

To update one, run the same `Copy-Item` again — `-Force` overwrites. To
remove it:

```powershell
Remove-Item -Recurse -Force "$env:USERPROFILE\.claude\skills\unslop"
```

## Adding a skill

1. `New-Item -ItemType Directory skills\<name>` and write `SKILL.md` with
   YAML frontmatter:
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

**If the skill learns from your own documents**, give it an `exemplars/`
sub-folder with a `README.md` saying what belongs in it and how to name the
files, plus a `.gitignore` holding

```gitignore
*
!*/
!README.md
!.gitignore
```

so the folder travels with the skill but your real documents never get
committed. That is how `brief-writer` and `email-writer` work: fill the folder
on a machine that has your documents, copy the skill folder to the endpoint, and
it arrives already knowing what your writing looks like.

Skills here are unversioned — they are prose, not an interface anything else
depends on, so there is no version to keep in sync. (The MCP servers under
`plugins/` are versioned; see the root [README](../README.md#versioning).)
