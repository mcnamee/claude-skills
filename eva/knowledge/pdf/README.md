# knowledge\pdf\

PDFs converted to Markdown, tables and reading order preserved. This is the
`pdf-to-md` plugin's **output** folder, and it points in here so that converting
a PDF is the same act as adding it to the knowledge base.

| | |
|---|---|
| **Setting** | `EVA_KNOWLEDGE_DIR` (the `pdf-to-md` plugin appends `\pdf`) |
| **Default** | `C:\Eva\knowledge\pdf` |
| **Source PDFs** | [`..\..\documents\pdf`](../../documents/pdf) |
| **Filenames** | the PDF's name with a `.md` extension |

## Conversion quality varies with the PDF

A PDF generated from Word converts cleanly. A scan converts to nothing at all —
there is no OCR here, so an image-only PDF produces an empty or near-empty file.
If a converted file looks thin, check the source before assuming the content is
missing from the knowledge base.

Sub-folder structure under `documents\pdf\` is mirrored here when the plugin
runs with its recursive option enabled.

## Safe to delete

Derived files. Delete and re-convert whenever you want; the source PDFs are
untouched by conversion.
