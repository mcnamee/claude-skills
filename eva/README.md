# Eva

The working folder for the assistant — everything the plugins read, write and
index. This copy in the repo is a **scaffold**: the folder tree with a README in
every folder explaining what belongs there, Eva's own instructions in
[`CLAUDE.md`](CLAUDE.md), and no content.

Copy it to `C:\Eva` on the endpoint, set four environment variables, and every
plugin in this repo lines up with it - with nothing left to configure but your
API keys and Confluence/Jira URLs.

```powershell
Copy-Item -Recurse C:\path\to\claude-skills\eva C:\Eva

[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe",        "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",               "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",             "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",               "User")
```

Those four are the *whole* configuration story for folders. Each plugin appends
its own name as a sub-folder - `word` reads `documents\word`, writes
`knowledge\word` and takes its blanks from `templates\word` - so no
plugin has a folder setting of its own to keep in sync, and there are no folder
command-line flags at all.

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

Four environment variables, set once for your Windows account. Every folder
below is one of those roots plus the plugin's own name, so there is nothing
per-plugin to configure and nothing that can drift out of step.

| Variable | Root | Default |
|---|---|---|
| `EVA_PYTHON` | the `python.exe` every server runs under | *(no default - set it)* |
| `EVA_DOCUMENTS_DIR` | `documents\` | `C:\Eva\documents` |
| `EVA_TEMPLATES_DIR` | `templates\` | `C:\Eva\templates` |
| `EVA_KNOWLEDGE_DIR` | `knowledge\` | `C:\Eva\knowledge` |

Where each plugin lands, and **which folders must exist**:

| Plugin | Documents | Templates | Knowledge |
|---|---|---|---|
| `word` | `documents\word\` | `templates\word\` | `knowledge\word\` |
| `powerpoint` | `documents\powerpoint\` | `templates\powerpoint\` | `knowledge\powerpoint\` |
| `excel` | `documents\excel\` | — | — |
| `pdf-to-md` | `documents\pdf\` | — | `knowledge\pdf\` |
| `outlook` | — | — | `knowledge\email\` |
| `confluence` | — | — | `knowledge\confluence\` |
| `knowledge-base` | `knowledge\` *(the whole root - it indexes everything)* | — | `knowledge\captures\` |
| `jira` | — | — | — |

`knowledge-base` is the one exception to the sub-folder rule, and necessarily
so: it indexes the entire `knowledge\` root, which is exactly the point of
every other plugin writing into a sub-folder of it. Its vector store lives
outside the corpus in `index\` (`KB_INDEX_DIR`, default `C:\Eva\index`).

### Overriding one folder

The suite-wide root is normally all you need. If one endpoint has to put a
single folder somewhere else, each plugin still reads a variable of its own,
which beats the root: `MSWORD_DOCS_DIR`, `MSWORD_TEMPLATES_DIR`,
`MSWORD_KB_DIR`, `POWERPOINT_*` likewise, `EXCEL_DOCS_DIR`, `PDF2MD_DOCS_DIR`,
`PDF2MD_KB_DIR`, `OUTLOOK_KB_DIR`, `CONFLUENCE_KB_DIR`, `KB_DOCS_DIR`,
`KB_INDEX_DIR`, `KB_OUTPUT_DIR`. Each takes a full path, and an optional one
takes `off` to switch that feature off altogether (`OUTLOOK_KB_DIR=off` and no
email is ever written to disk).

There are **no folder command-line flags** on any server - configuration is
environment variables only, so two settings can never disagree about a path.
The only flags are `--check`, `--version`, and the `knowledge-base` actions
(`--reindex`, `--search`, `--ask`).

## The three rules that keep it tidy

**One folder per file type.** Everything `.docx` lives in `documents\word\`,
everything `.pptx` in `documents\powerpoint\`, and so on — whether you put it
there or Eva wrote it. There is no `input\`, no `output\`, and no `inbox\` vs
`library\` split. Those were the tree's own idea of your workflow, and every one
of them made you decide which folder a file belonged in before you could ask a
question about it. One folder per type has no such decision, and it is why the
plugins need no folder settings of their own: the folder name is the plugin
name. Organise inside it however you like — `word` and `powerpoint` search
recursively, and a bare filename still finds the file.

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

**Upgrading an existing `C:\Eva`?** Move the templates out of the old
`reference\` zone into one folder per plugin, hand the exemplars to the skills
that now read them, and set the four environment variables:

```powershell
New-Item -ItemType Directory -Force C:\Eva\templates\word, C:\Eva\templates\powerpoint
Move-Item C:\Eva\reference\templates\*.docx        C:\Eva\templates\word
Move-Item C:\Eva\reference\templates\*.pptx,*.potx C:\Eva\templates\powerpoint
Copy-Item C:\Eva\reference\exemplars\* "$env:USERPROFILE\.claude\skills\exemplar-writer\exemplars"
Remove-Item -Recurse C:\Eva\reference

[Environment]::SetEnvironmentVariable("EVA_PYTHON",        "C:\Python311\python.exe", "User")
[Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\Eva\documents",         "User")
[Environment]::SetEnvironmentVariable("EVA_TEMPLATES_DIR", "C:\Eva\templates",         "User")
[Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\Eva\knowledge",         "User")
```

Then reconfigure the plugins once (`/plugin` → each one) to drop the folder
answers they no longer ask for.

## Moving it somewhere else

`C:\Eva` is the fallback baked into each plugin's config block — short, on the
system drive, the same for every user, and free of the spaces and
`%USERPROFILE%` expansion that break command lines. To put the tree elsewhere,
copy it there and point the four environment variables at the new location;
nothing in the plugins assumes the drive or the folder name.

## Why nothing here is committed

[`.gitignore`](.gitignore) ignores every file except these READMEs and
[`CLAUDE.md`](CLAUDE.md), because in real use this tree is full of your
organisation's documents and this repo is public. The folder structure still
travels with the repo: git stores no empty directories, so each folder is in the
repo *because* of its README.
