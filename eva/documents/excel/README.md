# documents\excel\

The `excel` plugin's sandbox — the workbooks it may read and, when writing is
switched on, change.

| | |
|---|---|
| **Setting** | `EVA_DOCUMENTS_DIR` (the `excel` plugin appends `\excel`) |
| **Default** | `H:\Eva\documents\excel` |
| **Access** | **read-only** by default. With `EXCEL_ALLOW_WRITE=true` the plugin can also change a workbook here (cells, Table rows, a new pivot table) or save a copy beside it |
| **Formats** | `.xlsx`, `.xlsm` |

Workbooks are parsed directly for reading, so Excel does not need to be
installed and no file is ever locked or modified by reading it. Writing is
different: it drives the installed Excel, and a workbook open in your own Excel
has to be closed first.

Spreadsheets attached to a Confluence page download into this folder (the
`confluence` plugin's `confluence_download_attachment`), at the top level where
this plugin can see them.

## Keep workbooks at the top level

**Sub-folders are not listed.** Unlike `word`, this plugin lists only the top
level of this folder, so a workbook filed in `documents\excel\Finance\` will not
appear when you ask what workbooks are available.

If you need grouping, prefix the filename — `Finance - Budget FY26.xlsx`,
`Ops - Headcount.xlsx` — which sorts the same way a folder would and stays
visible. Name matching is exact-or-substring first, then fuzzy, so a prefix
costs nothing when asking for a file by name.

## Nothing here is indexed

Spreadsheets are not text and there is no `.xlsx`-to-Markdown mirror in the
suite. The RAG index will never contain a figure from a workbook — Eva reads
them live, on request. If a summary of a workbook needs to be searchable, ask
for it to be captured as a note (`kb_capture`), which puts the *summary* in
[`..\..\knowledge\captures`](../../knowledge/captures) while the numbers stay
here as the source.

## Read-only unless you switch writing on

With writing off (the default) nothing writes here, so this folder can safely
point at a copy of real finance or HR workbooks. Files with formulas, macros or
external links are read for their stored values; opening them here changes
nothing.

Switch writing on and that stops being true, so decide first whether the
workbooks here are ones Eva should be able to change. Asking for the result to
be saved as a new file (`save_as`) keeps the original untouched. Macros never
run when a workbook is opened for a write.
