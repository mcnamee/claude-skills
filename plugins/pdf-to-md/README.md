# PDF to Markdown

Convert PDFs in a folder to Markdown with tables preserved — the front end for a
local knowledge base built from PDF source material.

| | |
|---|---|
| **Server** | `pdf-to-md.py` v5.1.0 |
| **pip install** | `pymupdf pymupdf4llm` |
| **Platform** | any |
| **Writes to disk** | yes — the output folder only (`C:\Eva\knowledge\pdf`) |

OCR of scanned PDFs additionally requires **Tesseract installed on the machine**
(not a pip package).

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install pdf-to-md@mcnamee-claude-skills
```

| Prompt | Default | Env var | Purpose |
|---|---|---|---|
| PDF folder | `C:\Eva\documents\pdf` | `PDF2MD_DOCS_DIR` | Folder containing the source PDFs |
| Output folder | `C:\Eva\knowledge\pdf` | `PDF2MD_OUTPUT_DIR` | Folder the `.md` files are written to |
| Python interpreter | — | — | **Required.** Absolute path to the `python.exe` that has `pymupdf` installed |

Both folders are pre-filled with their place in the [Eva working
tree](../../eva). The output default sits **inside** the `knowledge-base`
plugin's documents folder on purpose: converting a PDF is then the same act as
adding it to the RAG corpus. Point it outside `C:\Eva\knowledge` and the
Markdown piles up unindexed.

To search sub-folders too, set `PDF2MD_RECURSIVE=1` as an environment variable
(sub-folder structure is mirrored in the output).

Image references and "missing image" placeholder text are stripped from the
Markdown by default — see [Images](#images). To keep them, set
`PDF2MD_KEEP_IMAGE_REFS=1`.

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| CLI flag | Env var | Purpose |
|---|---|---|
| `--docs-dir` | `PDF2MD_DOCS_DIR` | **Required.** Folder containing the source PDFs. Falls back to the `DOCS_DIR` constant, default `C:\Eva\documents\pdf` |
| `--output-dir` | `PDF2MD_OUTPUT_DIR` | **Required.** Folder to write `.md` files into. Falls back to the `OUTPUT_DIR` constant, default `C:\Eva\knowledge\pdf` — inside the `knowledge-base` corpus, so conversions are indexed |
| `--recursive` | `PDF2MD_RECURSIVE=1` | Also search sub-folders of the docs folder (sub-folder structure is mirrored in the output) |
| `--keep-image-refs` | `PDF2MD_KEEP_IMAGE_REFS=1` | Keep the image references and "missing image" placeholder text that are [stripped by default](#images) |
| `--check` | — | Print environment/config diagnostics (folders, dependency status, PDFs found) and exit (no server) |
| `--version` | — | Print version and exit (works even without `pymupdf` installed) |

## Images

The server never writes or embeds image files, so **every image reference in the
raw output points at a file that does not exist**. Those dangling "missing
image" references are noise once the Markdown is in a knowledge base, so they
are removed. Stripped by default:

| Removed | Where it comes from |
|---|---|
| `![alt](name.pdf-0003-01.png)`, `![alt][ref]`, `<img …>` | `pymupdf4llm`, pointing at an image file that was never written |
| `<!-- Start of picture text -->` / `<!-- End of picture text -->` | `pymupdf4llm` wrapping text it found *inside* a picture |
| `Image removed by sender.` | Outlook, where a remote image was blocked or stripped |
| `Right-click here to download pictures. To help protect your privacy, Outlook prevented automatic download of this picture from the Internet.` | Outlook |
| `This image cannot currently be displayed.` | Word / PowerPoint placeholder for a picture it could not render |
| `The linked image cannot be displayed. The file may have been moved, renamed, or deleted. …` | Word / PowerPoint |
| `[cid:image001.png@01DA…]` | Inline-image content-ID token left by an email-to-PDF print |

Kept:

- **Text found inside a picture** — that is real page text. Only the markers
  around it go, and the `<br>` separators `pymupdf4llm` uses inside them become
  real line breaks. `<br>` anywhere else (inside a table cell) is left alone.
- **Table rows** that held an image: the row keeps its shape and the cell ends
  up empty.
- A line left holding nothing but its bullet or number collapses to a blank
  line, so no stray markers remain.

Pass `--keep-image-refs` (or `PDF2MD_KEEP_IMAGE_REFS=1`) to turn all of this off
and get the raw `pymupdf4llm` output back. If your PDFs carry a placeholder
phrase that is not in the list above, add it to the `MISSING_IMAGE_PHRASES`
constant in `pdf-to-md.py`.

## File access

Reads only the docs folder, writes only the output folder. Paths are resolved
(symlinks included) before the containment check.

## Usage examples

1. "Convert every PDF in the reference folder to Markdown." → `convert_all_pdfs`
2. "Convert just the 'procurement policy' PDF to Markdown." → `convert_pdf_to_markdown` (PDF names are matched forgivingly, including a fuzzy near-miss)
3. "Reconvert all PDFs to Markdown even though some already have a .md file, since the source PDFs changed." → `convert_all_pdfs` with `force=true`
4. "Convert all our compliance PDFs (including those in sub-folders) to Markdown so the knowledge-base server can search them." → `convert_all_pdfs` (with recursion enabled) feeding into the `knowledge-base` plugin's `kb_index` / `kb_ask`

Point the output folder at the `knowledge-base` plugin's documents folder and
converted PDFs land alongside your Confluence pages, emails and Word documents
in the same RAG index.

## Troubleshooting

- **"dependency missing" after installing `pymupdf`** — the server logs
  `sys.executable` on startup. Almost always the Python interpreter you gave the
  plugin isn't the one you pip-installed into:
```powershell
  & "C:\path\to\python.exe" -m pip install pymupdf pymupdf4llm
  ```
- `--check` reports the folders, dependency status and how many PDFs were found:
```powershell
  & "C:\path\to\python.exe" pdf-to-md.py --check
  ```
