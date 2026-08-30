# Eva

The working folder for the assistant — everything the plugins read, write and
index. This copy in the repo is a **scaffold**: the folder tree with a README in
every folder explaining what belongs there, and no content.

Copy it to `C:\Eva` on the endpoint and every plugin default in this repo lines
up with it, with nothing left to configure but your Python path, API keys and
Confluence/Jira URLs.

```powershell
Copy-Item -Recurse C:\path\to\claude-skills\eva C:\Eva
```

## The four zones

The tree is organised by **what writes to a folder**, not by topic:

| Zone | Holds | Written by |
|---|---|---|
| [`knowledge\`](knowledge) | The RAG corpus — Markdown only | every server that mirrors what it reads, plus `kb_capture` |
| [`index\`](index) | The ChromaDB vector store | `knowledge-base`, from `knowledge\` |
| [`documents\`](documents) | The binary library — `.docx`, `.pptx`, `.xlsx`, `.pdf` | you |
| [`output\`](output) | Generated files | `word`, `powerpoint` |
| [`reference\`](reference) | Templates and exemplars — style, not facts | you |

```
C:\Eva\
├─ knowledge\        the indexed corpus (.md / .txt only)
│  ├─ notes\           Markdown you write by hand
│  ├─ captures\        notes kb_capture writes back
│  ├─ confluence\      pages the confluence plugin mirrored
│  ├─ email\           emails the outlook plugin mirrored
│  ├─ word\            documents the word plugin mirrored
│  ├─ powerpoint\      decks the powerpoint plugin mirrored
│  └─ pdf\             PDFs the pdf-to-md plugin converted
├─ index\            ChromaDB — derived, disposable
├─ documents\
│  ├─ word\            .docx  (searched recursively)
│  │  ├─ inbox\          drop files here to work on
│  │  └─ library\        documents you keep
│  ├─ powerpoint\      .pptx  (searched recursively)
│  ├─ excel\           .xlsx  (top level only — see its README)
│  └─ pdf\             .pdf   source PDFs
├─ output\
│  ├─ word\            documents the assistant creates
│  └─ powerpoint\      decks the assistant creates
└─ reference\
   ├─ exemplars\       finished documents showing what good looks like
   └─ templates\       blank branded files new documents start from
```

## Which setting points where

Every path below is the **default** baked into the plugin, so a stock install
needs none of these set. Each still takes a flag and an environment variable if
you want the folder somewhere else — flag beats environment variable beats the
default.

| Folder | Plugin → setting | Flag / env var |
|---|---|---|
| `knowledge\` | `knowledge-base` → documents folder | `--docs-dir` / `KB_DOCS_DIR` |
| `knowledge\captures\` | `knowledge-base` → output folder | `--output-dir` / `KB_OUTPUT_DIR` |
| `knowledge\confluence\` | `confluence` → knowledge-base folder | `--kb-dir` / `CONFLUENCE_KB_DIR` |
| `knowledge\email\` | `outlook` → knowledge-base folder | `--kb-dir` / `OUTLOOK_KB_DIR` |
| `knowledge\word\` | `word` → knowledge-base folder | `--kb-dir` / `MSWORD_KB_DIR` |
| `knowledge\powerpoint\` | `powerpoint` → knowledge-base folder | `--kb-dir` / `POWERPOINT_KB_DIR` |
| `knowledge\pdf\` | `pdf-to-md` → output folder | `--output-dir` / `PDF2MD_OUTPUT_DIR` |
| `index\` | `knowledge-base` → index folder | `--index-dir` / `KB_INDEX_DIR` |
| `documents\word\` | `word` → documents folder | `--docs-dir` / `MSWORD_DOCS_DIR` |
| `documents\powerpoint\` | `powerpoint` → presentations folder | `--docs-dir` / `POWERPOINT_DOCS_DIR` |
| `documents\excel\` | `excel` → workbook folder | `--docs-dir` / `EXCEL_DOCS_DIR` |
| `documents\pdf\` | `pdf-to-md` → documents folder | `--docs-dir` / `PDF2MD_DOCS_DIR` |
| `output\word\` | `word` → output folder | `--output-dir` / `MSWORD_OUTPUT_DIR` |
| `output\powerpoint\` | `powerpoint` → output folder | `--output-dir` / `POWERPOINT_OUTPUT_DIR` |
| `reference\templates\` | `word` → templates folder | `--templates-dir` / `MSWORD_TEMPLATES_DIR` |
| `reference\templates\` | `powerpoint` → templates folder | `--templates-dir` / `POWERPOINT_TEMPLATES_DIR` |
| `reference\exemplars\` | none — read as ordinary files | — |

`jira` touches no local folder at all.

## The two rules that keep it tidy

**One indexed root.** `knowledge\` is the only folder the RAG index reads, so
every server that produces Markdown writes *inside* it. Point a mirror somewhere
else and it will fill up faithfully while `kb_ask` never sees a word of it —
which is the single most common way to end up with a knowledge base that
"doesn't know" something you are certain you read.

**Folders are named after what wrote them, not what they are about.** Topic
folders rot: every document belongs to three of them and you spend your time
deciding which. Provenance is unambiguous, maps one-to-one onto a setting, and
retrieval is semantic anyway — folder names do not affect what comes back. It
also makes cleanup surgical: delete `knowledge\confluence\` and re-mirror after
a space is restructured, with nothing you wrote yourself at risk.

## Moving it somewhere else

`C:\Eva` is a literal path in each plugin's config block — short, on the system
drive, the same for every user, and free of the spaces and `%USERPROFILE%`
expansion that break command lines. To put the tree elsewhere, copy it there and
set the environment variables above (or the flags) to match; nothing in the
plugins assumes the drive or the folder name.

## Why nothing here is committed

[`.gitignore`](.gitignore) ignores every file except these READMEs, because in
real use this tree is full of your organisation's documents and this repo is
public. The folder structure still travels with the repo: git stores no empty
directories, so each folder is in the repo *because* of its README.
