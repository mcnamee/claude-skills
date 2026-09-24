# .claude

Claude Code's project folder for `H:\Eva`. Open Claude Code in `H:\Eva` and it
loads everything here, scoped to this working folder rather than your whole
account - which is why the suite ships its standalone skills and agents inside
the `eva\` scaffold instead of asking you to install them globally.

| Folder | Holds | Read by |
|---|---|---|
| [`skills\`](skills) | Standalone skills - `/brief-writer`, `/email-writer`, `/exemplar-writer`, `/polish`, `/unslop` - each with any `exemplars\` it learns from | Claude Code and OpenCode, opened in `H:\Eva` |
| [`agents\`](agents) | Subagents - `researcher` | Claude Code only |

Nothing here is indexed: the knowledge base reads `knowledge\` alone.

Anything else Claude Code writes here (a `settings.local.json`, say) is yours
and never committed to the repo. To update a skill or agent later, copy just
its folder or file over the old one - [`skills\README.md`](skills/README.md)
and [`agents\README.md`](agents/README.md) have the commands. That leaves your
exemplars in place, because the repo carries none.
