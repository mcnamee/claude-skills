# documents\pdf\

Source PDFs, for the `pdf-to-md` plugin to convert into the knowledge base.

| | |
|---|---|
| **Setting** | `EVA_DOCUMENTS_DIR` (the `pdf-to-md` plugin appends `\pdf`) |
| **Default** | `C:\Eva\documents\pdf` |
| **Access** | read-only — conversion never alters a PDF |
| **Output** | [`..\..\knowledge\pdf`](../../knowledge/pdf) |

Conversion preserves reading order, headings and tables, including borderless
tables inferred from column alignment.

## Sub-folders need switching on

By default only the top level of this folder is converted. Turn on the plugin's
recursive option (`PDF2MD_RECURSIVE=1`) and sub-folders are
included, with the same structure mirrored into the output folder.

## Scans convert to nothing

There is no OCR in this suite. A PDF that is an image of a page — a scan, a
photographed signature page — produces an empty or near-empty Markdown file, and
it does so quietly. If a converted file looks suspiciously thin, open the PDF
and check whether its text is selectable. Getting a scan into the knowledge base
means OCRing it elsewhere first.

## Where the PDFs stay

Converted PDFs stay here. The Markdown copy in `knowledge\pdf\` is what gets
indexed and quoted; this folder remains the source, so a conversion can be
redone after a change to chunking or after a bad run, without hunting for the
originals again.
