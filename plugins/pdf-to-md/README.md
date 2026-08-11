# PDF to Markdown

Convert PDFs in a folder to Markdown with tables preserved — the front end for a
local knowledge base built from PDF source material.

| | |
|---|---|
| **Server** | `pdf-to-md.py` v4.0.1 |
| **pip install** | `pymupdf pymupdf4llm` |
| **Platform** | any |
| **Writes to disk** | yes — the output folder only |

OCR of scanned PDFs additionally requires **Tesseract installed on the machine**
(not a pip package).

## Install

```
/plugin marketplace add C:\path\to\mcp-servers
/plugin install pdf-to-md@mcnamee-mcp-servers
```

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| PDF folder | **yes** | `PDF2MD_DOCS_DIR` | Folder containing the source PDFs |
| Output folder | **yes** | `PDF2MD_OUTPUT_DIR` | Folder the `.md` files are written to |
| Python interpreter | **yes** | — | Absolute path to the `python.exe` that has `pymupdf` installed |

To search sub-folders too, set `PDF2MD_RECURSIVE=1` as an environment variable
(sub-folder structure is mirrored in the output).

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| CLI flag | Env var | Purpose |
|---|---|---|
| `--docs-dir` | `PDF2MD_DOCS_DIR` | **Required.** Folder containing the source PDFs |
| `--output-dir` | `PDF2MD_OUTPUT_DIR` | **Required.** Folder to write `.md` files into |
| `--recursive` | `PDF2MD_RECURSIVE=1` | Also search sub-folders of the docs folder (sub-folder structure is mirrored in the output) |
| `--check` | — | Print environment/config diagnostics (folders, dependency status, PDFs found) and exit (no server) |
| `--version` | — | Print version and exit (works even without `pymupdf` installed) |

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
  ```
  "C:\path\to\python.exe" -m pip install pymupdf pymupdf4llm
  ```
- `--check` reports the folders, dependency status and how many PDFs were found:
  ```
  "C:\path\to\python.exe" pdf-to-md.py --check
  ```
