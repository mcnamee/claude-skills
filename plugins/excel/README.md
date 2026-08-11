# Excel (.xlsx)

Read and analyse Excel workbooks. **Read-only** — nothing is ever written to a
workbook — and it parses `.xlsx` directly as a zip of XML, so Excel itself does
not need to be installed.

| | |
|---|---|
| **Server** | `excel.py` v3.0.1 |
| **pip install** | _none_ — standard library only |
| **Platform** | any |
| **Writes to disk** | no |

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install excel@mcnamee-claude-skills
```

This is the simplest plugin in the suite — standard library only, one prompt —
so it's a good one to install first if you're confirming the flow works.

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| Workbook folder | **yes** | `EXCEL_DOCS_DIR` | The only folder the server may read workbooks from |
| Python interpreter | **yes** | — | Absolute path to the `python.exe` to launch the server with |

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| CLI flag | Env var | Purpose |
|---|---|---|
| `--docs-dir` | `EXCEL_DOCS_DIR` | **Required.** Folder of `.xlsx`/`.xlsm` workbooks to expose — the server only reads files inside it and refuses to start without one. Falls back to the `DOCS_DIR` constant in the file |
| `--check` | — | Print environment/config diagnostics and exit (no server) |
| `--list` | — | List readable workbooks in the folder and exit (no server) |
| `--version` | — | Print version and exit |

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
& "C:\path\to\python.exe" excel.py --check
& "C:\path\to\python.exe" excel.py --list
```
