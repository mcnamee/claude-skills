# Using the suite with OpenCode

The rest of this repo is written for Claude Code. This page is the OpenCode
equivalent of the [Install](README.md#install) section: the same servers, the
same skills and the same `H:\Eva` tree, wired in through OpenCode's own config
instead of Claude Code plugins.

Nothing in the servers changes. What changes is the packaging around them, and
all of it lives in **one file, `H:\Eva\opencode.json`** - no environment
variables to set:

| Claude Code | OpenCode |
|---|---|
| `/plugin install` per plugin | one config registering all eight servers |
| `EVA_PYTHON` and `${CLAUDE_PLUGIN_ROOT}` find each `.py` | the paths are written into the config |
| Plugin install prompts, and secrets as environment variables | each server's `environment` block in the config |
| Skills come with each plugin | the config's `skills.paths` reads them straight from your clone |
| `CLAUDE.md` loads by convention | named in the config's `instructions` |
| Agents (`agents/`) | not provided for OpenCode |

Because the config sits on `H:` with the rest of `H:\Eva`, it survives a
change of endpoint. (OpenCode's global config, under `%USERPROFILE%\.config`,
does not, so nothing here uses it.)

Everything below is **Windows / PowerShell**, and assumes the repo is cloned at
`H:\Claude-Skills`.

## 1. Lay out the working folder

Same as for Claude Code - copy the repo's `eva\` folder to `H:\Eva`, and fill in
the **About me** block of `H:\Eva\CLAUDE.md` as the
[main README](README.md#install) describes.

The copy brings a template, `H:\Eva\opencode.example.jsonc`. OpenCode does not
read it; it is the starting point for your own config.

## 2. Install the pip dependencies

Exactly as in [step 2 of the main install](README.md#install), into the Python
you will name in the config.

## 3. Create your config

```powershell
Copy-Item H:\Eva\opencode.example.jsonc H:\Eva\opencode.json
```

Then edit `H:\Eva\opencode.json`. Copying `eva\` again later only replaces the
template, so your filled-in copy is never overwritten. (Compare the two after a
`git pull` if you want to pick up a new server or setting.)

**Paths.** Find and replace:

- `C:/path/to/python.exe` with the `python.exe` from step 2 (it appears once
  per server)
- `H:/Claude-Skills`, only if your clone lives somewhere else

Write paths with **forward slashes** (`C:/Python312/python.exe`). This is a JSON
file, so a single backslash is an escape character: `C:\Python312` fails to
load, and `\b`, `\n`, `\r` or `\t` silently change the path. Doubled
backslashes (`C:\\Python312\\python.exe`) also work.

**Folders.** Nothing to do. Every server defaults to its sub-folder of
`H:\Eva`. To move one, add its override variable to that server's `environment`
block, e.g. `"MSWORD_DOCS_DIR": "D:/Work/Word"` - the variables are in each
plugin's README.

## 4. Fill in each plugin's settings

Each server's `environment` block lists its settings with blank values. Fill in
the ones you need. **A blank value means "not set"**, so the server's default
applies and there is nothing to delete. What each one accepts is in that
plugin's README.

| Plugin | Settings | Required |
|---|---|---|
| [word](plugins/word) | `MSWORD_AUTHOR` (tracked-change author) | |
| [outlook](plugins/outlook) | `OUTLOOK_SEARCH_FOLDERS`, `OUTLOOK_BLACKLIST_FILE`, `OUTLOOK_CALENDAR_HOURS`, `OUTLOOK_CALENDAR_PAGE_FILL`, `OUTLOOK_CALENDAR_COLOURS`, `OUTLOOK_MEETING_HOURS`, `OUTLOOK_ALLOW_DRAFTS`, `OUTLOOK_GAL_SCAN_CAP` | |
| [confluence](plugins/confluence) | `CONFLUENCE_NAME`, `CONFLUENCE_BASE_URL`, `CONFLUENCE_TOKEN`; the `_2` versions for a second instance; `CONFLUENCE_KB_AUTOSAVE`, `CONFLUENCE_BODY_FORMAT` | `CONFLUENCE_BASE_URL`, `CONFLUENCE_TOKEN` |
| [jira](plugins/jira) | `JIRA_BASE_URL`, `JIRA_PROJECTS`, `JIRA_TOKEN` | `JIRA_BASE_URL`, `JIRA_TOKEN` |
| [knowledge-base](plugins/knowledge-base) | `KB_EMBED_URL`, `KB_EMBED_MODEL`, `KB_EMBED_API_KEY` | |
| [excel](plugins/excel), [powerpoint](plugins/powerpoint), [pdf-to-md](plugins/pdf-to-md) | *(none)* | |

Tokens and API keys go in here too. That keeps them in a plain-text file on
`H:`, which is backed up - fine if the people who run the backups can already
reach those accounts, but worth a thought.

**A value in the config beats a Windows environment variable of the same name,
blank included.** If you already set something as a Windows variable for Claude
Code and want OpenCode to use it, delete that line from the config.

## 5. Choose the model

The config defaults to Anthropic, with only that provider offered. Uncomment
`apiKey` in the `provider` block and fill it in, or leave it commented and run
`/connect` once inside OpenCode.

For a corporate gateway, uncomment `baseURL`. For a different provider, change
`enabled_providers`, `provider`, `model` and `small_model` together.

The config also turns off session sharing (`"share": "disabled"`) and
self-updating (`"autoupdate": false`).

## 6. Switch off what you haven't installed

Every server starts enabled. For any plugin whose pip dependencies you skipped
(typically `outlook` without `pywin32`), or that you don't use, set
`"enabled": false` on its entry. Otherwise it just shows as failed.

## 7. Skills

Nothing to install for the plugin skills. The config points OpenCode's
`skills.paths` at `H:/Claude-Skills/plugins`, so it reads every plugin's skill
straight from your clone, and a `git pull` is all it takes to update them.

This adds to OpenCode's usual skill folders rather than replacing them, so your
own skills keep working from any of:

- `H:\Eva\.claude\skills` (when OpenCode is opened in `H:\Eva`) - the one that
  survives a change of endpoint
- `%USERPROFILE%\.claude\skills`
- `%USERPROFILE%\.config\opencode\skills`

To add a **standalone skill** (`skills\`), copy it into `H:\Eva\.claude\skills`:

```powershell
New-Item -ItemType Directory -Force H:\Eva\.claude\skills | Out-Null
Copy-Item -Recurse H:\Claude-Skills\skills\brief-writer H:\Eva\.claude\skills\
```

Give your own skills names that don't clash with a plugin skill (`word`,
`outlook`, `slide-deck` and so on). With two skills of the same name, OpenCode
keeps only one, and which one isn't guaranteed.

## 8. Check it

Run a server's `--check` to confirm the Python path and its pip packages:

```powershell
& C:\path\to\python.exe H:\Claude-Skills\plugins\word\word.py --check
```

This works as-is for `word`, `powerpoint`, `excel`, `outlook` and `pdf-to-md`.
`confluence`, `jira` and `knowledge-base` need their settings, which live in
the config rather than your environment, so check those through OpenCode below.

Then open `H:\Eva` in VS Code, open the integrated terminal and run `opencode`
(or press `Ctrl+Esc` once the OpenCode extension is installed). From
`H:\Eva`, `opencode mcp list` shows each server and whether it connected.

## Differences to expect

- **Tool names carry the server name.** OpenCode calls `msword_open` from the
  `word` server `word_msword_open`. The skills and `CLAUDE.md` use the short
  names, which the model maps without trouble.
- **Skills are loaded by the model**, through OpenCode's `skill` tool, rather
  than as `/word:word`-style slash commands.
- **No agents.** `agents/researcher.md` is Claude Code only.
- **Timeouts.** OpenCode's default MCP timeout is 5 seconds. The config sets 5
  minutes per server, and 15 for `knowledge-base`, whose first index over a
  large corpus can run for minutes.

## Troubleshooting

| Symptom | Cause |
|---|---|
| OpenCode reports a JSON or config parse error | A single backslash in a path - use forward slashes (step 3) |
| None of the servers appear | OpenCode wasn't opened in `H:\Eva`, or the file is still named `opencode.example.jsonc` |
| A server shows as failed | Run its `--check` (step 8); usually a wrong Python path, a missing pip package, or a required setting left blank |
| A Windows variable seems ignored | The config has a line for it, which wins even when blank - delete the line |
| A plugin skill is missing | The `skills.paths` entry doesn't match your clone - OpenCode skips a folder that doesn't exist |
