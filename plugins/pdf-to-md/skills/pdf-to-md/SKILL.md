---
name: pdf-to-md
description: Convert PDFs to Markdown via the pdf2md MCP server (tables included). Use when the user asks to convert PDF files to Markdown, or to make PDFs searchable by the knowledge-base server.
---

# PDF → Markdown (via the `pdf2md` MCP server)

Requires the `pdf-to-md.py` MCP server (pymupdf/pymupdf4llm). If its tools
are not available, tell the user to wire it in first (see the repo README)
and to verify with `python pdf-to-md.py --check`.

## Tools

| Tool | Use for |
|---|---|
| `convert_all_pdfs` | Convert every PDF in the configured docs folder (`force=true` reconverts existing) |
| `convert_pdf_to_markdown` | Convert ONE PDF by rough name (refuses ambiguous matches) |

## Workflow

1. "Convert everything" → `convert_all_pdfs`. PDFs whose `.md` already
   exists are skipped — add `force=true` only when the user says the source
   PDFs changed.
2. "Convert the X document" → `convert_pdf_to_markdown` with the user's
   rough name as `query`. If it refuses because several PDFs match, show the
   candidates and ask which one.
3. Report the per-file results (converted / skipped / failed) rather than
   just "done" — scanned pages without OCR fail per-file, not silently.
4. Output lands in the configured output folder, which by default is inside
   the knowledge-base corpus (`C:\Eva\knowledge\pdf`) — so converted content
   is searchable as soon as it is indexed (suggest `kb_index` if the RAG server
   is in use).

## Notes

- Both folders come from the environment at server launch
  (`%EVA_DOCUMENTS_DIR%\pdf` → `%EVA_KNOWLEDGE_DIR%\pdf`); the tools cannot
  convert arbitrary paths outside them.
- Sub-folders are only included if the server was started with
  `PDF2MD_RECURSIVE=1` (structure is mirrored in the output).
- Image references and "missing image" placeholder text are stripped from the
  Markdown: the server writes no image files, so every such reference is
  dangling. Text found *inside* a picture is kept. Nothing needs doing about
  images — if the user actually wants those references back, it is
  `PDF2MD_KEEP_IMAGE_REFS=1` at server launch, not a tool argument.
- OCR of scanned PDFs requires Tesseract on the machine; without it, text
  PDFs still convert and scanned ones report a clear per-file failure.
