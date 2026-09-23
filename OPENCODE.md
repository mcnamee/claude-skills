# Using the suite with OpenCode

The rest of this repo is written for Claude Code. This page is the OpenCode
equivalent of the [Install](README.md#install) section: the same servers, the
same skills and the same `C:\Eva` tree, wired in through OpenCode's own config
instead of Claude Code plugins.

Nothing in the servers changes. What changes is the packaging around them:

| Claude Code | OpenCode |
|---|---|
| `/plugin install` per plugin | one project config, [`eva/opencode.json`](eva/opencode.json), registering all eight servers |
| `${CLAUDE_PLUGIN_ROOT}` finds each `.py` | a fifth environment variable, `EVA_REPO_DIR`, points at your clone |
| Plugin install prompts (Confluence URL, author, ...) | plain environment variables - see step 5 |
| Skills come with each plugin | a git hook copies them to OpenCode's skills folder on every `git pull` |
| `CLAUDE.md` loads by convention | named in `opencode.json`'s `instructions` |
| Agents (`agents/`) | not provided for OpenCode |

Everything below is **Windows / PowerShell**.

## 1. Lay out the working folder

Same as for Claude Code - copy the repo's `eva\` folder to `C:\Eva`:

```powershell
Copy-Item -Recurse H:\Claude-Skills\eva C:\Eva
```

This brings `C:\Eva\opencode.json` with it, which OpenCode reads whenever it is
opened in `C:\Eva`. Fill in the **About me** block of `C:\Eva\CLAUDE.md` as the
[main README](README.md#install) describes.

> Copying `eva\` again later overwrites `C:\Eva\opencode.json`. If you edit
> your copy (to switch off a server, say), re-apply the edit after the copy -
> or copy just the README files you need rather than the whole tree.

## 2. Install the pip dependencies

Exactly as in [step 2 of the main install](README.md#install), into the
interpreter you will name as `EVA_PYTHON`.

## 3. Set the environment variables

The four suite-wide variables, plus `EVA_REPO_DIR` for OpenCode:

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:/path/to/python.exe", "User")
[Environment]::SetEnvironmentVariable("EVA_REPO_DIR",      "H:/Claude-Skills",      "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",      "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",      "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",      "User")
```

**`EVA_PYTHON` and `EVA_REPO_DIR` must use forward slashes.** OpenCode pastes
their values into `opencode.json` as raw text before parsing it, so a backslash
becomes a JSON escape: `H:\Claude-Skills` stops OpenCode loading the config at
all, and a folder starting with `b`, `n`, `r` or `t` (`\bin`, `\tools`) silently
turns into a different path. Windows accepts forward slashes, so the same
`EVA_PYTHON` still works for Claude Code if you run both.

The three folder variables are never pasted into the config - each server reads
them straight from the environment - so they can keep their backslashes.

| Variable | What it points at |
|---|---|
| `EVA_PYTHON` | The `python.exe` from step 2 |
| `EVA_REPO_DIR` | Your clone of this repo (OpenCode only) |
| `EVA_DOCUMENTS_DIR` / `EVA_TEMPLATES_DIR` / `EVA_KNOWLEDGE_DIR` | As in the [main README](README.md#install) |

## 4. Choose the model

`opencode.json` defaults to Anthropic, with only that provider offered. Either
set `ANTHROPIC_API_KEY`, or run `/connect` once inside OpenCode:

```powershell
setx ANTHROPIC_API_KEY "your-api-key"
```

For a corporate gateway, uncomment `baseURL` in the `provider` block. For a
different provider, change `enabled_providers`, `provider`, `model` and
`small_model` together.

The config also turns off session sharing (`"share": "disabled"`) and
self-updating (`"autoupdate": false`).

## 5. Set each plugin's own settings

In Claude Code, `/plugin install` prompts for these. OpenCode has no prompts, so
set them as environment variables instead - only for the plugins you use, and
only the ones you need. Each server inherits your environment, so nothing needs
adding to `opencode.json`. What each one does and accepts is in that plugin's
README.

| Plugin | Variables | Secrets |
|---|---|---|
| [word](plugins/word) | `MSWORD_AUTHOR` (tracked-change author) | |
| [outlook](plugins/outlook) | `OUTLOOK_SEARCH_FOLDERS`, `OUTLOOK_BLACKLIST_FILE`, `OUTLOOK_CALENDAR_HOURS`, `OUTLOOK_CALENDAR_PAGE_FILL`, `OUTLOOK_CALENDAR_COLOURS`, `OUTLOOK_MEETING_HOURS`, `OUTLOOK_ALLOW_DRAFTS`, `OUTLOOK_GAL_SCAN_CAP` | |
| [confluence](plugins/confluence) | `CONFLUENCE_NAME`, `CONFLUENCE_BASE_URL`, and `CONFLUENCE_NAME_2`, `CONFLUENCE_BASE_URL_2` for a second instance | `CONFLUENCE_TOKEN`, `CONFLUENCE_TOKEN_2` |
| [jira](plugins/jira) | `JIRA_BASE_URL`, `JIRA_PROJECTS` | `JIRA_TOKEN` |
| [knowledge-base](plugins/knowledge-base) | `KB_EMBED_URL`, `KB_EMBED_MODEL` | `KB_EMBED_API_KEY` |
| [excel](plugins/excel), [powerpoint](plugins/powerpoint), [pdf-to-md](plugins/pdf-to-md) | *(none)* | |

```powershell
setx CONFLUENCE_BASE_URL "https://confluence.example.com"
setx CONFLUENCE_TOKEN    "your-personal-access-token"
```

The per-server folder overrides (`MSWORD_DOCS_DIR` and so on) work exactly as
the [configuration conventions](README.md#configuration-conventions) describe.

`setx` doesn't reach processes that are already running: **quit VS Code
completely** and reopen it after setting anything.

## 6. Switch off what you haven't installed

Every server in `opencode.json` starts enabled. For any plugin whose pip
dependencies you skipped (typically `outlook` without `pywin32`), set
`"enabled": false` on its entry in `C:\Eva\opencode.json`. Otherwise it just
shows as failed.

## 7. Install the skills, and keep them current

OpenCode reads skills from `%USERPROFILE%\.config\opencode\skills`. The plugin
skills live inside each plugin in this repo, so a git hook copies them across
every time you pull. Turn it on once, in your clone:

```powershell
cd H:\Claude-Skills
git config core.hooksPath .githooks
& "C:\Program Files\Git\bin\sh.exe" .githooks/sync-opencode-skills
```

The last line does the first copy by hand, since the hook only runs when a
pull actually brings something in - run it the same way whenever you want to
re-sync without pulling. From then on, every `git pull` (including VS Code's
**Sync** button, and `git pull --rebase`) prints `opencode: synced N plugin skill(s) to ...`. Each
skill folder is replaced whole, so a file removed in the repo disappears from
the copy too. Skills you added yourself are left alone.

`core.hooksPath` applies to this clone only, and replaces `.git\hooks`, so any
hooks you had there stop running. The hooks are in [`.githooks/`](.githooks).

The **standalone skills** (`skills\`) aren't copied by the hook: it replaces a
skill folder whole, which would wipe any exemplars you keep in the installed
copy. Copy the ones you want by hand, once:

```powershell
Copy-Item -Recurse H:\Claude-Skills\skills\brief-writer "$env:USERPROFILE\.config\opencode\skills\"
```

(If you also use Claude Code, OpenCode reads `%USERPROFILE%\.claude\skills`
too, so standalone skills installed there already work.)

## 8. Check it

Run a server's `--check` first. It is far easier to read than an MCP
connection failure:

```powershell
& $env:EVA_PYTHON "$env:EVA_REPO_DIR/plugins/word/word.py" --check
```

Then open `C:\Eva` in VS Code, open the integrated terminal and run `opencode`
(or press `Ctrl+Esc` once the OpenCode extension is installed). From
`C:\Eva`, `opencode mcp list` shows each server and whether it connected.

## Differences to expect

- **Tool names carry the server name.** OpenCode calls `msword_open` from the
  `word` server `word_msword_open`. The skills and `CLAUDE.md` use the short
  names, which the model maps without trouble.
- **Skills are loaded by the model**, through OpenCode's `skill` tool, rather
  than as `/word:word`-style slash commands.
- **No agents.** `agents/researcher.md` is Claude Code only.
- **Timeouts.** OpenCode's default MCP timeout is 5 seconds. `opencode.json`
  sets 5 minutes per server, and 15 for `knowledge-base`, whose first index
  over a large corpus can run for minutes.

## Troubleshooting

| Symptom | Cause |
|---|---|
| OpenCode reports a JSON or config parse error | A backslash in `EVA_PYTHON` or `EVA_REPO_DIR` - use forward slashes (step 3) |
| A server shows as failed | Run its `--check` (step 8); usually a missing pip package or folder |
| A variable seems ignored | VS Code was running when it was set - quit it completely and reopen |
| No `opencode: synced` line after a pull | `git config core.hooksPath` isn't set in this clone (step 7) |
