# Excel (.xlsx)

Read and analyse Excel workbooks — any sheet (tab), any **Table by name**, and
**pivot tables** — by parsing `.xlsx` directly as a zip of XML, so reading
needs neither Excel nor any pip package. Switch writing on and it can also
**change** workbooks — write cells into a tab, add and update Table rows, and
build a real **PivotTable** — by driving the Excel installed on the endpoint.

| | |
|---|---|
| **Server** | `excel.py` v6.1.0 |
| **pip install** | _none_ to read; `pywin32` to write |
| **Platform** | any to read; Windows with desktop Excel to write |
| **Writes to disk** | no — unless `EXCEL_ALLOW_WRITE=true`, and then only workbooks in its folder |

## What it can do

| | Read | Write *(`EXCEL_ALLOW_WRITE=true`)* |
|---|---|---|
| **A sheet (tab)** | `excel_read_range`, `excel_get_headers`, `excel_search`, `excel_column_stats` — all take `sheet` | `excel_write_cells` — a block of values or formulas from any cell; `create_sheet` adds the tab |
| **A Table by name** | `excel_list_tables`, `excel_read_table` — by name wherever it sits, with column pick, a filter and paging | `excel_add_table_rows`, `excel_update_table_rows` (find the row by a column value, set others) |
| **Pivot tables** | `excel_list_pivot_tables`, `excel_read_pivot_table` — layout (source, rows, columns, filters, values) and the figures shown | `excel_create_pivot_table` — a real PivotTable from a Table or range, with row/column/filter fields and sum/count/average/max/min/... values |

A pivot table's figures are read from the cells Excel saved, so they are as of
its **last refresh**; the reply gives that date. Reading can't refresh a pivot
— only Excel can.

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install excel@mcnamee-claude-skills
```

This is the simplest plugin in the suite - standard library only, no prompts at
all - so it's a good one to install first if you're confirming the flow works.

## Writing

**Off unless `EXCEL_ALLOW_WRITE=true`** — then the four write tools appear; until
then they are not offered at all. Writing needs **Windows, desktop Excel and
`pywin32`** in the same Python:

```powershell
& "C:\path\to\python.exe" -m pip install pywin32
```

Why Excel and not a Python library: a workbook is a web of parts that must
agree — formulas, Table ranges, pivot caches, charts — and only Excel keeps
them all consistent. It is also the only thing that can build a real
PivotTable (no Python library can). Each write call:

1. **refuses up front if the workbook is open** — in your Excel or anywhere else.
   Close it and ask again;
2. starts a **private, invisible Excel** (never your own Excel window), with
   alerts, events and **macros switched off**;
3. checks every sheet, Table and column name **before changing anything**, then
   makes the change and lets Excel recalculate;
4. saves — in place, or with **`save_as`** as a new workbook in the same folder,
   leaving the original untouched — and quits that Excel.

Any error closes the workbook **without saving**, so a half-made change never
reaches the file. Values starting with `=` are formulas; `null` clears a cell;
write dates as `YYYY-MM-DD`. A row added to a Table as an object
(`{"Region": "North", "Amount": 5}`) writes only the columns named, so
calculated columns keep their formulas.

> **Verify on the endpoint first.** The Excel automation follows Excel's
> documented object model but cannot run in this repo's (Linux) test
> environment. With writing on, `--check` starts and quits Excel to prove
> automation works; then try `excel_create_pivot_table` with `save_as` on a
> copy before relying on it.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `H:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `H:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `H:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "H:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "H:\Eva\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "H:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

Of the four, this server uses two: `EVA_PYTHON` and `EVA_DOCUMENTS_DIR`. It
reads no templates, and writes nothing unless writing is switched on.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `excel`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `H:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_DOCUMENTS_DIR%\excel` | The only folder the server may read workbooks from - and, with writing on, change them in or save copies into. **Top level only** - unlike `word`, this server does not search sub-folders. The `confluence` plugin downloads spreadsheet attachments here | **Fatal.** The server refuses to start without it |

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
| `EXCEL_ALLOW_WRITE=true` | Offer the write tools — see [Writing](#writing). Default off: read-only |

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Print environment/config diagnostics and exit (no server). It reports which variable the folder came from, whether writing is on and, if it is, whether Excel can be started |
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
reach files outside it. Nothing is written unless `EXCEL_ALLOW_WRITE=true`; then
only a workbook in that folder is changed, and `save_as` is confined to the
same folder (it never replaces an existing file unless `overwrite=true`).

## Usage examples

1. "What Excel workbooks are available for me to look at?" → `excel_list_workbooks`
2. "List the sheets in the 'budget' workbook." → `excel_list_sheets` (the `workbook` name is matched forgivingly — a near-miss like `"q3 budget"` still resolves to `Budget Q3 2024.xlsx`)
3. "What are the column headers on the 'Q3' sheet of the budget workbook?" → `excel_get_headers`
4. "Read rows A1:D50 from the Q3 sheet." → `excel_read_range`
5. "Find every cell in the budget workbook that mentions 'Marketing'." → `excel_search`
6. "Give me the sum, average, min and max of the Revenue column on the Q3 sheet." → `excel_column_stats`
7. "What Tables are in the budget workbook?" → `excel_list_tables`
8. "Show me the rows of tblProjects where Status is Red." → `excel_read_table` with `filter_column` / `filter_value`
9. "What does the pivot on the Summary tab say?" → `excel_read_pivot_table`
10. "Put these three totals into B2:B4 on the Summary tab." → `excel_write_cells` *(writing on)*
11. "Add a row to tblProjects for the new CRM project, owner Jane, status Green." → `excel_add_table_rows` *(writing on)*
12. "Set the status of project P-17 to Amber." → `excel_update_table_rows` with `match` / `set` *(writing on)*
13. "Make a pivot of tblSales: regions down the side, products across, total Amount." → `excel_create_pivot_table` *(writing on)*
14. "Grab the budget spreadsheet off the FY26 Confluence page and pivot it by cost centre." → `confluence_download_attachment`, then `excel_list_tables` and `excel_create_pivot_table` on the downloaded file

## Troubleshooting

> **If a server fails with `Executable not found in $PATH: "${EVA_PYTHON}"`**,
> the variable is not set in the environment Claude Code was launched from. Set
> it (see above), then quit Claude Code completely and reopen — `setx` and
> `[Environment]::SetEnvironmentVariable` do not reach a process that is already
> running.

Run the config check before wiring it in — it's far easier to read than an MCP
connection failure:

```powershell
& $env:EVA_PYTHON excel.py --check
& $env:EVA_PYTHON excel.py --list
```
