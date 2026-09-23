# documents\pdf\

The PDF folder: source PDFs for the `pdf-to-md` plugin to convert into the
knowledge base, and the day planners the `outlook` plugin prints.

| | |
|---|---|
| **Setting** | `EVA_DOCUMENTS_DIR` (both the `pdf-to-md` and `outlook` plugins append `\pdf`) |
| **Default** | `H:\Eva\documents\pdf` |
| **Access** | `pdf-to-md` reads only — conversion never alters a PDF. `outlook` writes day planners here and touches nothing else |
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

## Printed day planners

Ask for "today's printable calendar" or "tomorrow's planner" and `outlook` writes
`Calendar - <date> <weekday>.pdf` here: A4 landscape, the day on an hour-by-hour
timeline down one half and the following days down the other. It prints on the
top four fifths of the sheet, so print it single-sided at 100%, fold it in half,
and fold the blank bottom fifth up behind it to fit a diary (set
`OUTLOOK_CALENDAR_PAGE_FILL` to change how much is left to fold, or to `off` for
the whole sheet). Re-printing a day overwrites that day's sheet rather than
piling up copies.

They are **not worth converting**. A bulk `pdf-to-md` run will pick them up along
with everything else and fill the index with last week's meetings, so tidy old
sheets out, or set `OUTLOOK_DOCS_DIR` to a folder of its own if you print often.

## Where the PDFs stay

Converted PDFs stay here. The Markdown copy in `knowledge\pdf\` is what gets
indexed and quoted; this folder remains the source, so a conversion can be
redone after a change to chunking or after a bad run, without hunting for the
originals again.
