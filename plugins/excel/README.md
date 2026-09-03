# Excel (.xlsx)

Read and analyse Excel workbooks. **Read-only** — nothing is ever written to a
workbook — and it parses `.xlsx` directly as a zip of XML, so Excel itself does
not need to be installed.

| | |
|---|---|
| **Server** | `excel.py` v5.0.0 |
| **pip install** | _none_ — standard library only |
| **Platform** | any |
| **Writes to disk** | no |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install excel@mcnamee-claude-skills
```

This is the simplest plugin in the suite - standard library only, no prompts at
all - so it's a good one to install first if you're confirming the flow works.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `C:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `C:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

Of the four, this server uses two: `EVA_PYTHON` and `EVA_DOCUMENTS_DIR`. It
reads no templates and writes nothing at all.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `excel`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_DOCUMENTS_DIR%\excel` | The only folder the server may read workbooks from. **Top level only** - unlike `word`, this server does not search sub-folders | **Fatal.** The server refuses to start without it |

> **Only the top level is listed.** A workbook filed in
> `documents\excel\Finance\` will not appear in `excel_list_workbooks`. Keep
> workbooks directly in the folder and use filename prefixes
> (`Finance - Budget FY26.xlsx`) where you would otherwise want a sub-folder.
> See [`eva/documents/excel`](../../eva/documents/excel).

### Overriding one folder, and this server's own settings

The shared roots are normally all you need. These variables are this
server's own, and a folder variable here beats the matching root - use one
only when an endpoint's layout really differs.

| Variable | Purpose |
|---|---|
| `EXCEL_DOCS_DIR` | Full path to the workbook folder, instead of `%EVA_DOCUMENTS_DIR%\excel` |

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Print environment/config diagnostics and exit (no server). It reports which variable the folder came from |
| `--list` | List readable workbooks in the folder and exit (no server) |
| `--version` | Print version and exit |

## Finding a workbook by name

Every tool takes a `workbook` name, resolved forgivingly against the folder:
exact filename → name without extension → case-insensitive → a unique substring
→ and finally a **fuzzy** name match (the same matcher as `word`), so
*"budgit q3"* or *"q3 budget"* still opens `Budget Q3 2024.xlsx`. A genuinely
ambiguous name returns the candidate list rather than guessing; use
`excel_list_workbooks` to see what's available. Fuzzy fallbacks are logged to
stderr for audit.

## File access

Reads only inside the workbook folder. Paths are resolved (symlinks included)
before the containment check, so a symlink dropped inside the folder cannot
reach files outside it. Nothing is written.

## Usage examples

1. "What Excel workbooks are available for me to look at?" → `excel_list_workbooks`
2. "List the sheets in the 'budget' workbook." → `excel_list_sheets` (the `workbook` name is matched forgivingly — a near-miss like `"q3 budget"` still resolves to `Budget Q3 2024.xlsx`)
3. "What are the column headers on the 'Q3' sheet of the budget workbook?" → `excel_get_headers`
4. "Read rows A1:D50 from the Q3 sheet." → `excel_read_range`
5. "Find every cell in the budget workbook that mentions 'Marketing'." → `excel_search`
6. "Give me the sum, average, min and max of the Revenue column on the Q3 sheet." → `excel_column_stats`

## Troubleshooting

Run the config check before wiring it in — it's far easier to read than an MCP
connection failure:

```powershell
& $env:EVA_PYTHON excel.py --check
& $env:EVA_PYTHON excel.py --list
```
