# PDF to Markdown

Convert PDFs in a folder to Markdown with tables preserved — the front end for a
local knowledge base built from PDF source material.

| | |
|---|---|
| **Server** | `pdf-to-md.py` v6.0.0 |
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

Claude Code prompts for nothing - both folders and the Python interpreter come
from the shared environment variables below.

## Configuration

**Four environment variables configure every plugin in this suite.** Set them
once for your Windows account and this plugin has nothing else to configure -
there are no folder prompts at install time and no folder command-line flags.

| Variable | Purpose | Default |
|---|---|---|
| `EVA_PYTHON` | The `python.exe` every server runs under - the same one you installed the pip dependencies into | *(none - you must set it)* |
| `EVA_DOCUMENTS_DIR` | Root of the document library | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | Root of the template library | `C:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | Root of the RAG corpus - the one folder the index reads | `C:\Eva\knowledge` |

```powershell
[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",     "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",             "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",   "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",             "User")
```

`setx NAME "value"` does the same thing from `cmd`. Neither affects processes
that are already running, so quit and reopen your editor afterwards.

Of the four, this server uses three: `EVA_PYTHON`, `EVA_DOCUMENTS_DIR` and
`EVA_KNOWLEDGE_DIR`. It reads no templates.

### The folders this plugin uses

Every server works in its **own sub-folder** of those roots, named after
the plugin. This one uses `pdf`, and **each folder below must exist** -
create them, or copy the repo's [`eva/`](../../eva) folder to `C:\Eva` and
they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_DOCUMENTS_DIR%\pdf` | The source PDFs. **Read-only** - conversion never alters a PDF. Top level only unless `PDF2MD_RECURSIVE=1` | **Fatal.** The server refuses to start without it |
| `%EVA_KNOWLEDGE_DIR%\pdf` | Where the converted Markdown is written. It sits **inside** the `knowledge-base` plugin's corpus on purpose: converting a PDF is then the same act as adding it to the RAG index | **Fatal.** The server refuses to start without it |

> Deriving both folders from the two shared roots is what stops the classic
> mistake here - a Markdown folder outside `%EVA_KNOWLEDGE_DIR%` fills up
> faithfully while the index never reads a line of it.

### Overriding one folder, and this server's own settings

The shared roots are normally all you need. These variables are this
server's own, and a folder variable here beats the matching root - use one
only when an endpoint's layout really differs.

| Variable | Purpose |
|---|---|
| `PDF2MD_DOCS_DIR` | Full path to the PDF folder, instead of `%EVA_DOCUMENTS_DIR%\pdf` |
| `PDF2MD_KB_DIR` | Full path to the Markdown output folder, instead of `%EVA_KNOWLEDGE_DIR%\pdf`. *(Renamed from `PDF2MD_OUTPUT_DIR` in v6.0.0, so every plugin that writes into the knowledge tree now uses the same `_KB_DIR` name.)* |
| `PDF2MD_RECURSIVE=1` | Also convert sub-folders of the PDF folder; the sub-folder structure is mirrored in the output |
| `PDF2MD_KEEP_IMAGE_REFS=1` | Keep the image references and "missing image" placeholder text that are [stripped by default](#images) |

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Print environment/config diagnostics - the folders, which variable each came from, dependency status and how many PDFs were found - and exit (no server) |
| `--version` | Print version and exit (works even without `pymupdf` installed) |

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

Set `PDF2MD_KEEP_IMAGE_REFS=1` to turn all of this off
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
  & $env:EVA_PYTHON -m pip install pymupdf pymupdf4llm
  ```
- `--check` reports the folders, dependency status and how many PDFs were found:
```powershell
  & $env:EVA_PYTHON pdf-to-md.py --check
  ```
