# Skills

Standalone Claude skills — no MCP server, no Python, no dependencies. Each one
is a folder holding a `SKILL.md` (the instructions Claude follows) and a
`README.md` (what it does, for you).

```
eva\.claude\skills\      (H:\Eva\.claude\skills\ on the endpoint)
  <skill-name>/
    SKILL.md      ← the skill itself
    README.md     ← human-facing docs
    exemplars/    ← optional: your own documents for the skill to learn from
```

They live inside the [`eva\`](../..) scaffold, in its `.claude\skills\`
folder - the place Claude Code looks for a project's own skills - so copying
`eva\` to `H:\Eva` installs every one of them, scoped to your working folder
rather than your whole account. Every command below is **PowerShell**.

The [`plugins/`](../../../plugins) folder is the other half of this repo: MCP servers
that give Claude new *tools*. Skills here give Claude new *instructions* — they
need nothing installed and work anywhere, including offline.

## Available skills

| Skill | Invoke | What it does |
|---|---|---|
| [**brief-writer**](brief-writer) | `/brief-writer` | Drafts a decision or noting brief for a senior executive, following the structure of an exemplar in its own `exemplars/` folder, and finishing with `/polish` |
| [**email-writer**](email-writer) | `/email-writer` | Drafts an email in your voice, classifying what the email is for and matching that intent to your own sent mail in its `exemplars/` folder, then running `/unslop` |
| [**exemplar-writer**](exemplar-writer) | `/exemplar-writer` | Writes a document in the shape of one you already have — pulls the structure, section order, proportions and register out of an exemplar, then writes your material to that shape and runs `/unslop` |
| [**polish**](polish) | `/polish` | Rewrites a draft into Australian Public Service style — the Australian Government Style Manual — asking who the reader is and what the medium is, then picking the register from them |
| [**unslop**](unslop) | `/unslop` | Strips AI-slop markers from writing — padding, tell-tale vocabulary, stock LLM sentence shapes — leaving meaning and voice intact |

The three writers build on the other two: `brief-writer` runs `/polish` as its
last step, `email-writer` and `exemplar-writer` run `/unslop`. Install the pair
each one needs, not just the writer.

They divide by what is being written, not by how well they do it: a brief goes
to `brief-writer` (it settles decision versus noting, and writes the
recommendation line), an email to `email-writer` (intent plus a voice profile),
and anything else that should follow the shape of a document you already have to
`exemplar-writer`.

## Install

Nothing extra: they arrive with the `eva\` copy described in the
[main README](../../../README.md#install), at `H:\Eva\.claude\skills\`.

- **Claude Code** loads a project's skills from `.claude\skills\` in the
  folder it is opened in, so start it in `H:\Eva`.
- **OpenCode** reads the same folder when it is opened in `H:\Eva` - see
  [`OPENCODE.md`](../../../OPENCODE.md#7-skills).

Then invoke a skill by name: `/unslop`. Claude also applies a skill
automatically when the task matches its `description`, so an explicit slash
command isn't always necessary - though `unslop` is deliberately written to wait
until you ask.

If a newly copied skill doesn't appear, run `/doctor` or restart Claude Code.

To update one without re-copying the whole tree, copy its folder over the old
one - `-Force` overwrites the skill's files and leaves your exemplars alone,
because the clone has none to copy:

```powershell
Copy-Item -Recurse -Force C:\path\to\claude-skills\eva\.claude\skills\unslop H:\Eva\.claude\skills\
```

To remove one:

```powershell
Remove-Item -Recurse -Force H:\Eva\.claude\skills\unslop
```

**Want a skill everywhere, not just in `H:\Eva`?** Copy its folder to
`%USERPROFILE%\.claude\skills\` instead. Keep one copy or the other, not both:
with two skills of the same name, only one is used.

## Adding a skill

1. `New-Item -ItemType Directory eva\.claude\skills\<name>` and write `SKILL.md` with
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
3. Add a row to the table above, and to the standalone-skills table in the
   root [README](../../../README.md).

`eva/.gitignore` commits each skill's `SKILL.md` and `README.md` and nothing
else, so a skill that needs another file alongside them needs a line there too.

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
committed. That is how `brief-writer`, `email-writer` and `exemplar-writer`
work: fill the folder on a machine that has your documents, copy `eva\` to the
endpoint, and the skill arrives already knowing what your writing looks like.

A skill folder sits outside every MCP server's sandbox, so a `.docx`, `.pptx` or
`.pdf` exemplar in there may not be readable at all. Say so in the folder's
`README.md`, and point at the two fixes: keep a `.md` copy beside the original,
or keep the original in `H:\Eva\documents\<type>` and name it in the prompt.

Skills here are unversioned — they are prose, not an interface anything else
depends on, so there is no version to keep in sync. (The MCP servers under
`plugins/` are versioned; see the root [README](../../../README.md#versioning).)
