# Knowledge base (RAG)

True Retrieval-Augmented Generation over a folder of your own Markdown files: a
local ChromaDB vector index plus your own embeddings API.

| | |
|---|---|
| **Server** | `knowledge-base.py` v4.0.0 |
| **pip install** | `chromadb` (HTTP to your endpoints is stdlib `urllib` — no `requests`) |
| **Platform** | any |
| **Writes to disk** | yes — the vector index folder (`C:\Eva\index`), plus captured notes under `C:\Eva\knowledge\captures` |

## How it works

1. **Index** — documents are split into heading-aware chunks, embedded via your
   embeddings API endpoint, and stored in a local ChromaDB vector database on
   disk. Indexing is incremental: only new/changed files are re-embedded,
   deleted files are removed.
2. **Retrieve** — a question is embedded the same way and the most semantically
   similar chunks come back with source file, section heading and similarity
   score.
3. **Generate** — *(optional)* the retrieved chunks + question go to a
   chat-completions endpoint, which writes a grounded answer citing its sources.
   With no chat endpoint configured, `kb_ask` returns the retrieved context and
   the agent you're already talking to writes the answer — so generation works
   either way.
4. **Capture** — `kb_capture` writes a new Markdown note into the knowledge base
   and indexes it on the spot, so something worked out in a conversation is
   searchable next time instead of disappearing with the chat. See
   [Growing it](#growing-it).

> Retrieved document text **is** sent to the endpoints you configure — that is
> what RAG is — so point it only at material appropriate for those APIs.
> ChromaDB's anonymised telemetry is disabled in the file.

## Install

```
/plugin marketplace add C:\path\to\claude-skills
/plugin install knowledge-base@mcnamee-claude-skills
```

Claude Code prompts only for the two endpoint settings below; the Python
interpreter and every folder come from the shared environment variables in
[Configuration](#configuration).

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| Embeddings endpoint URL | **yes** | `KB_EMBED_URL` | Full URL of your embeddings API |
| Embeddings model | no | `KB_EMBED_MODEL` | Model name sent in embed requests; omit if the endpoint fixes one |

**Your API key is not stored in the plugin.** Set `KB_EMBED_API_KEY` as a
Windows user environment variable before starting Claude Code. API keys are
deliberately env-var only: there is no flag that could put a key in a command
line, where other local users would see it in a process listing.

```powershell
setx EVA_PYTHON        "C:\Python311\python.exe"
setx EVA_KNOWLEDGE_DIR "C:\Eva\knowledge"
setx KB_EMBED_API_KEY  "your-api-key"
```

`setx` does not affect processes that are already running, so quit VS Code
completely (a window reload is not enough) and reopen it. Check it took in a
**new** window with `$env:KB_EMBED_API_KEY`.

Everything else in the reference below is set with the matching `KB_*`
environment variable.

## First run

Before wiring it into the client — the docstring at the top of
`knowledge-base.py` walks through this. Set the configuration in the
environment once, then the action flags do the rest:

```powershell
$env:EVA_KNOWLEDGE_DIR = "C:\Eva\knowledge"
$env:KB_EMBED_URL      = "https://your-gateway/v1/embeddings"
$env:KB_EMBED_API_KEY  = "your-api-key"

& $env:EVA_PYTHON knowledge-base.py --check
& $env:EVA_PYTHON knowledge-base.py --reindex
& $env:EVA_PYTHON knowledge-base.py --search "some topic"
```

If you change embedding model, run `--reindex --force` once — vector dimensions
differ between models.

## Tools

| Tool | Purpose |
|---|---|
| `kb_index` | Build/update the vector index (incremental; `force=true` rebuilds — needed after changing embedding model) |
| `kb_retrieve` | Semantic vector search: top-k most similar chunks, with source file, heading and similarity score |
| `kb_ask` | Full RAG: retrieve, then generate a grounded cited answer (or return context for the agent, if no chat endpoint) |
| `kb_capture` | Save a new Markdown note into the knowledge base and index it immediately |
| `kb_status` | Documents, captured notes, index freshness + configuration summary (never shows keys) |

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

Of the four, this server uses `EVA_PYTHON` and `EVA_KNOWLEDGE_DIR`.

### The folders this plugin uses

Every other server works in its own sub-folder of the knowledge root; this one
is the exception, and indexes the **whole root**. **Each folder below must
exist** - create them, or copy the repo's [`eva/`](../../eva) folder to
`C:\Eva` and they all do.

| Folder | What it is for | Missing? |
|---|---|---|
| `%EVA_KNOWLEDGE_DIR%` | **The indexed corpus** - every `.md`/`.markdown`/`.txt` file under it, recursively. This is the plugin that reads the *whole* knowledge root rather than one sub-folder of it, which is exactly why every other plugin writes into a sub-folder: `knowledge\word`, `knowledge\email`, `knowledge\confluence`, `knowledge\pdf` and the rest are all inside the corpus, so a stock install indexes every one of them with no further wiring | **Fatal.** The server refuses to start without it |
| `%EVA_KNOWLEDGE_DIR%\captures` | Where `kb_capture` writes new notes. It has to sit **inside** the corpus, and must not be a dot-folder or land inside the index folder - all three are pruned from indexing, so a note written there would never be searchable | Created on demand. The server refuses to start if it resolves outside the corpus, rather than let a note be written and never indexed |
| `KB_INDEX_DIR` — default `C:\Eva\index` | The ChromaDB vector store. Deliberately **outside** the corpus, so a multi-hundred-megabyte binary database does not sit inside the folder you want to zip, copy or grep. Derived and disposable: delete it and re-index to rebuild, which you must do after changing embeddings model. This is the one folder in the suite that is not under a shared root | Created on demand. Set `KB_INDEX_DIR=off` to fall back to a hidden `.kb-rag-index` folder inside the corpus |

> See [`eva/knowledge`](../../eva/knowledge) for what belongs in each
> sub-folder of the corpus.

### This server's own settings

Everything else is an environment variable of this server's own. **API keys are
env-var only** - there is no flag that could leak one into a process listing.
`KB_DOCS_DIR` and `KB_OUTPUT_DIR` override the corpus and captures folders with
a full path each, for an endpoint whose layout differs.

| Env var | Purpose |
|---|---|
| `KB_DOCS_DIR` | **Required.** Folder of `.md`/`.markdown`/`.txt` docs, searched recursively. Default `C:\Eva\knowledge` |
| `KB_INDEX_DIR` | ChromaDB folder. Default `C:\Eva\index` — outside the corpus, so a large binary database does not sit inside the folder you want to be able to zip or grep. Clear the `INDEX_DIR` constant to fall back to `<docs-dir>\.kb-rag-index` |
| `KB_OUTPUT_DIR` | Folder `kb_capture` writes new notes into (default `C:\Eva\knowledge\captures`, created on demand). Must resolve **inside** the corpus, and must not be a dot-folder or sit inside the index folder — all three are pruned from indexing, so a note written there would never be searchable. The server refuses to start rather than let that happen |
| `KB_COLLECTION` | ChromaDB collection name (default `kb-rag`) |
| `KB_EMBED_URL` | **Required.** Full URL of the embeddings endpoint |
| `KB_EMBED_MODEL` | Model name sent in embed requests (omit if the endpoint fixes one) |
| `KB_EMBED_API_KEY` | API key for the embeddings endpoint |
| `KB_EMBED_AUTH_HEADER` | Header the key is sent in — default `Authorization` (as `Bearer <key>`); any other name (e.g. Azure's `api-key`) sends the raw key |
| `KB_EMBED_STYLE` | Request format: `openai` (default; batch `{"input": [...]}`), `ollama` (`{"prompt": ...}` one-per-request), `kserve-jina` (KServe V2 Open Inference Protocol — texts sent as a BYTES input tensor `{"inputs": [{"name", "shape", "datatype", "data"}]}`, e.g. a Jina embeddings model served on KServe; the model name is part of the `KB_EMBED_URL` path such as `https://host/v2/models/jina-embeddings/infer`, and flat FP32 output tensors are reshaped via their `shape`; nested data and KServe V1 `predictions` responses also parsed), or `raw-json` (plain `{"texts": [...]}` body — for KServe **custom** predictors and other bespoke wrappers that unpack the raw request body into their pipeline's arguments; the telltale symptom is a server error like `pipeline() missing 1 required positional argument: 'texts'` that doesn't change when the tensor name does). Response parsing additionally accepts bare `embedding`/`embeddings` shapes, so most bespoke internal endpoints work unchanged |
| `KB_EMBED_TENSOR_NAME` | `kserve-jina` style only: name of the input tensor the texts are sent as (default `text`) |
| `KB_EMBED_JSON_KEY` | `raw-json` style only: the JSON key the batch of texts is sent under (default `texts`; e.g. `instances` for a V1-flavoured custom wrapper) |
| `KB_EMBED_TEMPLATE` | **Full request-body control** when none of the styles matches your endpoint: the complete JSON body to POST, with the JSON string `"__TEXTS__"` where the array of texts goes; `"__COUNT__"` becomes the number of texts in the batch (an integer — for strictly-validated tensor `shape` fields) and `"__MODEL__"` the `KB_EMBED_MODEL` name. Overrides the style's request shape; batching still applies. E.g. a FastAPI custom predictor that validates `{"inputs": {"texts": [...]}}` (symptom: HTTP 422 with `loc: [body, inputs]`) is `KB_EMBED_TEMPLATE={"inputs": {"texts": "__TEXTS__"}}`; a strict V2 tensor envelope with a custom `texts` field is `{"inputs": [{"name": "texts", "shape": ["__COUNT__"], "datatype": "BYTES", "data": "__TEXTS__", "texts": "__TEXTS__"}]}`. Tip: FastAPI-wrapped endpoints publish their exact schema at `/openapi.json` (and Swagger UI at `/docs`) — fetch it with your mTLS certs instead of guessing field by field |
| `KB_EMBED_RESPONSE_PATH` | Dotted path to the vectors in the response when auto-detection can't find them, e.g. `outputs.embeddings` or `result.0.vectors` (numeric parts index lists). Applies to every style |
| `KB_DEBUG=1` | Log every request/response body (truncated) to stderr — run `--check` with it to see exactly what is sent and what came back, for matching an unknown endpoint |
| `KB_EMBED_BATCH` | Texts per embeddings request, openai and kserve-jina styles (default 16) |
| `KB_EMBED_QUERY_PREFIX` | Prefix for query embeds, for models that need it (e5-style `"query: "`) |
| `KB_EMBED_DOC_PREFIX` | Prefix for document embeds (`"passage: "`) |
| `KB_EMBED_EXTRA_HEADERS` | JSON object of extra HTTP headers for the embed endpoint |
| `KB_CHAT_URL` | *Optional.* Chat-completions endpoint for the generate step (OpenAI shape; Ollama `/api/chat` and `/api/generate` response shapes also parsed) |
| `KB_CHAT_MODEL` | Generation model name |
| `KB_CHAT_API_KEY` | API key for the chat endpoint (falls back to `KB_EMBED_API_KEY`) |
| `KB_CHAT_AUTH_HEADER` | As per `KB_EMBED_AUTH_HEADER` |
| `KB_CHAT_MAX_TOKENS` | `max_tokens` for generation (default 1024; 0 omits the field) |
| `KB_CHAT_EXTRA_HEADERS` | JSON object of extra HTTP headers for the chat endpoint |
| `KB_CA_CERT` | Path to a PEM CA bundle for an internal CA |
| `KB_CLIENT_CERT` | Path to a PEM client certificate, for gateways that require **mutual TLS (mTLS)** — the fix for errors like `CERTIFICATE_NOT_PROVIDED` / `certificate required`. Presented to both endpoints. Loaded once at startup, so a bad path/passphrase fails immediately |
| `KB_CLIENT_KEY` | Path to the PEM private key for the client certificate; omit if the `KB_CLIENT_CERT` file contains both cert and key |
| `KB_CLIENT_KEY_PASSWORD` | Passphrase, if the client private key is encrypted |
| `KB_VERIFY_SSL=false` | Disable TLS certificate verification. Note this cannot fix an mTLS error — it controls how *you* verify the *server*, not the certificate you present to it (that's `KB_CLIENT_CERT`) |
| `KB_TIMEOUT` | HTTP timeout in seconds (default 120) |
| `KB_CHUNK_CHARS` | Soft max characters per chunk (default 1500) |
| `KB_CHUNK_OVERLAP` | Overlap between adjacent chunks (default 200) |
| `KB_TOP_K` | Default chunks retrieved (default 5) |

### Command-line flags

Configuration is environment variables only, so nothing here sets a path. The
flags are actions:

| Flag | Purpose |
|---|---|
| `--check` | Validate config, self-test the capture path, confirm the captures folder is writable, call the endpoint(s) once, report index status, then exit. The capture checks run first and need no network, so they still tell you something on a machine that can't reach the endpoint |
| `--reindex` | Build/update the vector index, then exit (add `--force` to rebuild from scratch) |
| `--search QUERY` | Test retrieval from the command line, then exit |
| `--ask QUESTION` | Test full RAG (retrieve + generate) from the command line, then exit |
| `--debug` | Log every request/response body (truncated) to stderr - run `--check` with it to see exactly what is sent and what came back. Same as `KB_DEBUG=1` |
| `--version` | Print version and exit (works even without `chromadb` installed) |

## File access

Reads only inside the documents folder (`C:\Eva\knowledge`). Writes the
vector-index folder (`C:\Eva\index`) and captured notes
(`C:\Eva\knowledge\captures`); network only to the endpoint(s) you configure.

Existing documents are never modified or deleted — the only file the server
creates is a new note from `kb_capture`, and the caller supplies a **title, not
a path**. The filename is derived from that title, illegal and separator
characters are stripped, and the resolved path is checked to be inside the
captures folder before anything is written, so no title can reach outside it.

## Usage examples

1. "Using my knowledge base, can I extend my work trip by 2 days, pay for my own accommodation for the weekend, and fly back Monday?" → `kb_ask` (retrieves the travel policy's trip-extension chunks and generates a cited answer — or hands the agent the chunks to answer from, if no chat endpoint is configured)
2. "Find the parts of our policies about accommodation and per diem." → `kb_retrieve`
3. "I've added some new documents to the knowledge base folder — pick them up." → `kb_index` (incremental: only new/changed files are embedded)
4. "Is the knowledge base index up to date? What's actually indexed?" → `kb_status`
5. "Rebuild the whole knowledge base index from scratch (we switched embedding model)." → `kb_index` with `force=true`
6. "Save that as a note in my knowledge base — call it 'Records retention thresholds'." → `kb_capture`
7. "Remember how we worked out the leave accrual calculation, I'll want it again." → `kb_capture` with `source: "Procedure"`

## Feeding it

The `confluence`, `outlook` and `word` plugins each take a
**knowledge-base folder** setting. Point all of them at this plugin's documents
folder and what they collect — wiki pages, emails, Word documents — arrives as
Markdown there, ready for `kb_index`. The `pdf-to-md` plugin fills the same
folder from PDFs. `word` also mirrors on **create and save**, so a document
written for you lands there too.

They differ in *when* they write, and it is worth knowing which is which:
`word` and `powerpoint` mirror **everything they open or save**, while
`confluence` and `outlook` write only what you ask them to keep (`save_to_kb`)
— a search that turns up an unrelated page, or a mail thread skimmed for an
answer, does not land in the index. `CONFLUENCE_KB_AUTOSAVE` and
`OUTLOOK_KB_AUTOSAVE` turn those two back into wholesale mirrors if that is
what you want.

## Growing it

Mirroring captures what you **read**. `kb_capture` captures what you **write** —
the analysis, the research brief, the decision and its reasoning, the procedure
worked out across half an hour of conversation. Without it that work ends with
the chat, and the same question gets researched again in three months.

```
"Save that comparison as a note called 'Retention thresholds by record type'."
  → captures\Analysis - Retention thresholds by record type.md, indexed
```

Notes are written into the captures folder as `<Source> - <Title>.md`, with the
same header block the mirrors use:

```markdown
# Retention thresholds by record type

- Source: Analysis (written by Claude in conversation, not an authoritative document)
- Captured: 2026-08-17 14:32
- Tags: retention, records

---
```

Four things to know:

- **Capture is explicit.** Claude captures when you ask for something to be kept,
  or when you accept an offer to keep it — never on its own initiative. The
  `researcher` agent finishes by *offering* to capture its brief, naming the
  title it'd use.
- **That header stamp matters.** Captured notes come back from later searches
  sitting alongside real policy documents. The stamp is how you — and Claude —
  tell a working note from an authoritative source. Follow a note to the sources
  it cites rather than treating it as evidence in its own right.
- **Re-capturing the same title replaces the note** (with `overwrite=true`);
  without it the call is refused and the existing file reported. That keeps a
  revised brief as one file instead of five near-duplicates.
- **It only ever adds.** No tool here edits or deletes a document you put in the
  folder yourself.
