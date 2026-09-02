# Eva

The working folder for the assistant — everything the plugins read, write and
index. This copy in the repo is a **scaffold**: the folder tree with a README in
every folder explaining what belongs there, Eva's own instructions in
[`CLAUDE.md`](CLAUDE.md), and no content.

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
| [`knowledge\`](knowledge) | The RAG corpus — Markdown only | every server that mirrors what it opens or saves what you ask it to keep, plus `kb_capture` |
| [`index\`](index) | The ChromaDB vector store | `knowledge-base`, from `knowledge\` |
| [`documents\`](documents) | The binary library — `.docx`, `.pptx`, `.xlsx`, `.pdf` | you **and** the assistant |
| [`templates\`](templates) | Blank branded files a new document is created from — style, not facts | you |

```
C:\Eva\
├─ CLAUDE.md         who Eva is, and how she writes
├─ knowledge\        the indexed corpus (.md / .txt only)
│  ├─ notes\           Markdown you write by hand
│  ├─ captures\        notes kb_capture writes back
│  ├─ confluence\      pages you asked the confluence plugin to save
│  ├─ email\           emails you asked the outlook plugin to save
│  ├─ word\            documents the word plugin mirrored
│  ├─ powerpoint\      decks the powerpoint plugin mirrored
│  └─ pdf\             PDFs the pdf-to-md plugin converted
├─ index\            ChromaDB — derived, disposable
├─ documents\        one folder per file type — yours and Eva's together
│  ├─ word\            .docx  (searched recursively)
│  ├─ powerpoint\      .pptx  (searched recursively)
│  ├─ excel\           .xlsx  (top level only — see its README)
│  └─ pdf\             .pdf   source PDFs
└─ templates\        blank branded files new documents start from
   ├─ word\             .docx templates        (read-only)
   └─ powerpoint\       .pptx / .potx templates (read-only)
```

## The assistant's instructions

[`CLAUDE.md`](CLAUDE.md) at the top of the tree is the other half of the setup:
where the folders tell the plugins *where* to work, that file tells Claude *how*
to. It defines Eva as an executive virtual assistant, fixes Australian spelling
and Australian conventions for dates, times and money, bans em dashes and the
rest of the AI tells, and sets out how she sources facts on a network with no
internet and what she is not allowed to do on your behalf.

Claude Code loads it whenever you run in `C:\Eva`, so open Claude Code *here*
rather than somewhere else and it applies with nothing to configure. To make Eva
the default in every folder, copy it to `%USERPROFILE%\.claude\CLAUDE.md`
instead, remembering it will then shape coding sessions too.

The one part you must fill in is the **About me** block near the top: your name,
role, who you write to, how you sign off. Every line left blank there is
something Eva has to ask about or guess at.

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
| `templates\word\` | `word` → templates folder | `--templates-dir` / `MSWORD_TEMPLATES_DIR` |
| `templates\powerpoint\` | `powerpoint` → templates folder | `--templates-dir` / `POWERPOINT_TEMPLATES_DIR` |

`jira` touches no local folder at all.

## The three rules that keep it tidy

**One folder per file type.** Everything `.docx` lives in `documents\word\`,
everything `.pptx` in `documents\powerpoint\`, and so on — whether you put it
there or Eva wrote it. There is no `input\`, no `output\`, and no `inbox\` vs
`library\` split. Those were the tree's own idea of your workflow, and every one
of them made you decide which folder a file belonged in before you could ask a
question about it. One folder per type has no such decision, and each plugin
takes exactly one folder setting to match. Organise inside it however you like —
`word` and `powerpoint` search recursively, and a bare filename still finds the
file.

**One indexed root.** `knowledge\` is the only folder the RAG index reads, so
every server that produces Markdown writes *inside* it. Point a mirror somewhere
else and it will fill up faithfully while `kb_ask` never sees a word of it —
which is the single most common way to end up with a knowledge base that
"doesn't know" something you are certain you read.

**Folders are named after what wrote them, not what they are about.** Topic
folders rot: every document belongs to three of them and you spend your time
deciding which. Provenance is unambiguous, maps one-to-one onto a setting, and
retrieval is semantic anyway — folder names do not affect what comes back. It
also makes cleanup surgical: delete `knowledge\confluence\` and save the pages
again after a space is restructured, with nothing you wrote yourself at risk.

## Where exemplars live

There used to be a `reference\` zone here holding both templates and
**exemplars** — finished, good documents read for guidance so Claude can write
something in the same shape. Templates stayed (as [`templates\`](templates));
exemplars moved out of the tree entirely, into the `exemplars\` folder of the
skill that reads them:

| Skill | Exemplars folder |
|---|---|
| [`/exemplar-writer`](../skills/exemplar-writer) | `%USERPROFILE%\.claude\skills\exemplar-writer\exemplars\` |
| [`/brief-writer`](../skills/brief-writer) | `%USERPROFILE%\.claude\skills\brief-writer\exemplars\` |
| [`/email-writer`](../skills/email-writer) | `%USERPROFILE%\.claude\skills\email-writer\exemplars\` |

The point is that they travel with the skill: fill the folder on a machine that
has your documents, copy the skill folder to the endpoint, and it arrives
already knowing what your writing looks like. A folder in this tree would have
needed a separate copy and a path in every prompt.

Exemplars were never indexed and still are not, for the same reason: add a board
paper to the RAG corpus and its phrasing comes back with the same authority as a
policy. If you want a document to be both a reference *and* a model to write
like, put its content in `knowledge\notes\` and keep the formatted copy with
the skill.

**Upgrading an existing `C:\Eva`?** Two moves:

```powershell
New-Item -ItemType Directory -Force C:\Eva\templates\word, C:\Eva\templates\powerpoint
Move-Item C:\Eva\reference\templates\*.docx        C:\Eva\templates\word
Move-Item C:\Eva\reference\templates\*.pptx,*.potx C:\Eva\templates\powerpoint
Copy-Item C:\Eva\reference\exemplars\* "$env:USERPROFILE\.claude\skills\exemplar-writer\exemplars"
Remove-Item -Recurse C:\Eva\reference
```

## Moving it somewhere else

`C:\Eva` is a literal path in each plugin's config block — short, on the system
drive, the same for every user, and free of the spaces and `%USERPROFILE%`
expansion that break command lines. To put the tree elsewhere, copy it there and
set the environment variables above (or the flags) to match; nothing in the
plugins assumes the drive or the folder name.

## Why nothing here is committed

[`.gitignore`](.gitignore) ignores every file except these READMEs and
[`CLAUDE.md`](CLAUDE.md), because in real use this tree is full of your
organisation's documents and this repo is public. The folder structure still
travels with the repo: git stores no empty directories, so each folder is in the
repo *because* of its README.
