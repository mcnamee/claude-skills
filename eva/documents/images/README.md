# documents\images\

Images: the org charts the `/org-chart` skill draws, and pictures the
`confluence` plugin downloads from a page or uploads to one.

| | |
|---|---|
| **Setting** | `EVA_DOCUMENTS_DIR` (the skill and the `confluence` plugin append `\images`) |
| **Default** | `H:\Eva\documents\images` |
| **Access** | `/org-chart` writes a chart here as `<name>.json` (the structure, which you or Claude can edit) plus `<name>.svg` (the drawing). `confluence` downloads `.svg`, `.png`, `.jpg`, `.jpeg` and `.gif` attachments here, and uploads from here only when asked. Neither replaces an existing file unless asked |
| **Output** | none - nothing here is indexed |

No plugin opens images, so this folder has no sandbox of its own. It exists so
that a chart can be drawn, uploaded to a page, downloaded again and redrawn,
all from one place that the `confluence` plugin is allowed to read.

Keep the `.json` beside each `.svg`: it is what makes the next redraw a small
edit instead of starting from the spreadsheet again.
