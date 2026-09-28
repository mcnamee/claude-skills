# .claude

Claude Code's project folder for `H:\Eva`. Open Claude Code in `H:\Eva` and it
loads everything here, scoped to this working folder rather than your whole
account - which is why the suite ships its standalone skills and agents inside
the `eva\` scaffold instead of asking you to install them globally.

| Folder | Holds | Read by |
|---|---|---|
| [`skills\`](skills) | Standalone skills - `/brief-writer`, `/email-writer`, `/exemplar-writer`, `/polish`, `/unslop` - each with any `exemplars\` it learns from | Claude Code and OpenCode, opened in `H:\Eva` |
| [`agents\`](agents) | Subagents - `researcher` | Claude Code only |
| [`settings.example.json`](settings.example.json) | A settings TEMPLATE: registers the plugin marketplace, enables all eight plugins (each bringing its skills) and lists every environment variable the suite reads | Nothing, until you copy it to `settings.local.json` |

Nothing here is indexed: the knowledge base reads `knowledge\` alone.

## settings.example.json

The whole Claude Code setup in one file. Copy it once, then fill it in:

```powershell
Copy-Item H:\Eva\.claude\settings.example.json H:\Eva\.claude\settings.local.json
```

Re-copying `eva\` later only replaces the template, never your
`settings.local.json`. What it does, key by key:

- **`extraKnownMarketplaces`** registers your clone (`H:\Claude-Skills`) as the
  `mcnamee-claude-skills` marketplace, so there is no `/plugin marketplace add`
  to type. Change the `path` if your clone lives elsewhere.
- **`enabledPlugins`** turns on all eight plugins, and with them every plugin
  skill (`/word:word`, `/outlook:meeting-scheduler`, `/powerpoint:slide-deck`
  and the rest). Set one to `false` if you haven't installed its pip
  dependencies (typically `outlook` without `pywin32`).
- **`env`** lists every environment variable the suite reads: the four
  suite-wide ones first (set `EVA_PYTHON` to your `python.exe`; the three
  folders are already their defaults), then each plugin's settings, tuning and
  folder overrides, blank. **A blank value means "not set"**, so fill in only
  what you need - what each accepts is in the plugin's README.

Claude Code applies the marketplace and most `env` values only once you accept
the workspace trust prompt for `H:\Eva`, so open Claude Code there and say yes.
Then quit and reopen it once, and check with `/plugin` and `/mcp`. If a plugin
shows as enabled but not installed, run `/plugin install <name>@mcnamee-claude-skills`.

Three things to know:

- **It is strict JSON.** No comments, and a backslash is doubled
  (`"C:\\Python312\\python.exe"`) or replaced with a forward slash.
- **A value here beats a Windows environment variable of the same name, blank
  included.** If you already set one with `setx` and want to keep it, delete
  that line from the file.
- **Tokens sit in plain text** on `H:`, which is backed up. Fine if the people
  who run the backups can already reach those accounts; otherwise delete the
  token lines and set them as Windows variables instead.

The standalone skills and agents need no entry: Claude Code loads them from
this folder's `skills\` and `agents\` whenever it is opened in `H:\Eva`.

Anything else Claude Code writes here (a `settings.local.json`, say) is yours
and never committed to the repo. To update a skill or agent later, copy just
its folder or file over the old one - [`skills\README.md`](skills/README.md)
and [`agents\README.md`](agents/README.md) have the commands. That leaves your
exemplars in place, because the repo carries none.
