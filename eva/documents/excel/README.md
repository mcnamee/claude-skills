# documents\excel\

The `excel` plugin's sandbox — the workbooks it may read.

| | |
|---|---|
| **Setting** | `EVA_DOCUMENTS_DIR` (the `excel` plugin appends `\excel`) |
| **Default** | `C:\Eva\documents\excel` |
| **Access** | **read-only** — the plugin never writes a workbook |
| **Formats** | `.xlsx`, `.xlsm` |

Workbooks are parsed directly, so Excel does not need to be installed and no
file is ever locked or modified by reading it.

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

## Read-only, and worth keeping that way

Because nothing writes here, this folder can safely point at a copy of real
finance or HR workbooks. Files with formulas, macros or external links are read
for their stored values; opening them here changes nothing.
