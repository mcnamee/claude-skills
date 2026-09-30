---
name: excel
description: Read and analyse Excel workbooks via the excel MCP server (.xlsx/.xlsm) - any sheet/tab, Excel Tables by name, and pivot tables - and, when writing is switched on, write cells into a tab, add or update Table rows, and create real PivotTables. Use when the user asks what's in a spreadsheet, wants data read, searched, filtered or summarised from Excel files, asks about a named Table or a pivot table, or asks to update a spreadsheet, add rows, or build a pivot table.
---

# Excel (via the `excel` MCP server)

Requires the `excel.py` MCP server. Reading parses .xlsx directly (no Excel
needed); the write tools exist only when the endpoint sets
`EXCEL_ALLOW_WRITE=true`, and they drive the installed Excel (Windows,
pywin32). If its tools are not available, tell the user to wire it in first
(see the repo README) and to verify with `python excel.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `excel_list_workbooks` | What workbooks are available |
| `excel_list_sheets` | Sheet names + dimensions, and the Tables / pivot tables on each |
| `excel_get_headers` | The header row of a sheet |
| `excel_read_range` | Read cells (optionally an A1 range like `A1:D50`) |
| `excel_search` | Find cells matching a query across a workbook |
| `excel_column_stats` | Sum/average/min/max of one column |
| `excel_list_tables` | The Excel Tables (ListObjects) in a workbook |
| `excel_read_table` | A Table by name: columns, a filter, paging |
| `excel_list_pivot_tables` | Every pivot table's layout |
| `excel_read_pivot_table` | One pivot's layout plus the figures it shows |
| `excel_write_cells` | Write values/formulas into a tab *(writing on)* |
| `excel_add_table_rows` | Append rows to a Table *(writing on)* |
| `excel_update_table_rows` | Change the Table row(s) matching a value *(writing on)* |
| `excel_create_pivot_table` | Build a real PivotTable *(writing on)* |

The last four are only in the tool list when writing is switched on. If the
user asks for a change and they are missing, say that workbook writing is off
on this endpoint and that `EXCEL_ALLOW_WRITE=true` (in Claude Code's settings
`env` block, then a restart) turns it on - plus `pip install pywin32`. Do not
work around it.

## Workflow

Always orient before reading — workbook names resolve fuzzily, so:

1. `excel_list_workbooks` → confirm the workbook exists.
2. `excel_list_sheets` → pick the sheet, and see which Tables and pivot
   tables it holds.
3. **Data in a Table? Use the Table.** `excel_read_table` by name is better
   than a range: it knows the columns, skips the header and totals rows, can
   filter (`filter_column` / `filter_value`) and pick `columns`, and gives each
   row's sheet row number to cite.
4. Otherwise `excel_get_headers` → learn the columns, then `excel_read_range`
   for data, `excel_search` to locate values, or `excel_column_stats` for
   numeric summaries (prefer it over reading whole columns to compute stats
   yourself).
5. **"What does the pivot say?"** → `excel_read_pivot_table`. Its figures are
   as of the pivot's last refresh (the reply says when); if the user is acting
   on them and the source may have changed, say so.

A file from Confluence: `confluence_download_attachment` puts a spreadsheet in
this plugin's folder and reports its name - pass that name straight to these
tools.

## Changing a workbook (only when the write tools are present)

1. **Only change what the user asked for.** Never write as a side effect of
   reading.
2. **Close it first.** A workbook open in the user's Excel cannot be written;
   the tool says so. Ask them to close it and try again.
3. **Offer `save_as`** when the change is exploratory or the workbook is
   important ("save it as Budget FY26 - pivot.xlsx"), so the original stays
   untouched. Use `overwrite` only when the user wants a file replaced.
4. **Tables:** add rows as objects keyed by column name
   (`{"Project": "CRM", "Owner": "Jane"}`) so unnamed columns - including
   calculated ones - are left alone. To edit, `excel_update_table_rows` with a
   `match` that identifies ONE row (an ID column is best); if several match,
   the tool refuses - narrow it, or confirm with the user before passing
   `all_matches: true`.
5. **Cells:** `values` is a list of rows; `=...` is a formula, `null` clears,
   dates as `YYYY-MM-DD`. Read the area first if it might hold data the user
   wants kept.
6. **Pivot tables:** prefer `source_table` (the pivot then follows the Table
   as it grows) over `source_sheet` + `source_range`. Give `rows` and/or
   `columns`, and `values` as `[{"field": "Amount", "function": "sum"}]`
   (`count`, `average`, `max`, `min` and more). It lands on a new sheet
   `Pivot - <source>` unless `destination_sheet` is given. The reply shows the
   figures; report them.
7. **Report back** with what changed and the file it was saved to, as the tool
   reports it.

## Notes

- Without `EXCEL_ALLOW_WRITE=true` it cannot write cells, create files or run
  macros — never promise to update a spreadsheet then. Macros never run, even
  when writing.
- Reads are capped (rows/columns/search hits) to protect the context
  window; page through large ranges (or `offset` on a Table) rather than
  requesting everything.
- Formula cells return the value Excel last cached, not a re-evaluation.
  Legacy `.xls`/`.xlsb` and password-protected files are unsupported and
  reported as such.
