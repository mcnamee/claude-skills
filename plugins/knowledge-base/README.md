# Knowledge base (RAG)

True Retrieval-Augmented Generation over a folder of your own Markdown files: a
local ChromaDB vector index plus your own embeddings API.

| | |
|---|---|
| **Server** | `knowledge-base.py` v2.0.1 |
| **pip install** | `chromadb` (HTTP to your endpoints is stdlib `urllib` — no `requests`) |
| **Platform** | any |
| **Writes to disk** | yes — the vector index folder only |

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

> Retrieved document text **is** sent to the endpoints you configure — that is
> what RAG is — so point it only at material appropriate for those APIs.
> ChromaDB's anonymised telemetry is disabled in the file.

## Install

```
/plugin marketplace add C:\path\to\mcp-servers
/plugin install knowledge-base@mcnamee-mcp-servers
```

| Prompt | Required | Env var | Purpose |
|---|---|---|---|
| Documents folder | **yes** | `KB_DOCS_DIR` | Folder of `.md`/`.markdown`/`.txt` documents to index |
| Embeddings endpoint URL | **yes** | `KB_EMBED_URL` | Full URL of your embeddings API |
| Embeddings model | no | `KB_EMBED_MODEL` | Model name sent in embed requests; omit if the endpoint fixes one |
| Python interpreter | **yes** | — | Absolute path to the `python.exe` that has `chromadb` installed |

**Your API key is not stored in the plugin.** Set `KB_EMBED_API_KEY` as a
Windows user environment variable before starting Claude Code. API keys are
deliberately env-var only: there are no `--*-api-key` flags, because
command-line arguments are visible to other local users in process listings.

Everything else in the reference below is set with the matching `KB_*`
environment variable (or by launching the server manually with the flag).

## First run

Before wiring it into the client — the docstring at the top of
`knowledge-base.py` walks through this:

```
"C:\path\to\python.exe" knowledge-base.py --check
"C:\path\to\python.exe" knowledge-base.py --reindex
"C:\path\to\python.exe" knowledge-base.py --search "some topic"
```

If you change embedding model, run `--reindex --force` once — vector dimensions
differ between models.

## Tools

| Tool | Purpose |
|---|---|
| `kb_index` | Build/update the vector index (incremental; `force=true` rebuilds — needed after changing embedding model) |
| `kb_retrieve` | Semantic vector search: top-k most similar chunks, with source file, heading and similarity score |
| `kb_ask` | Full RAG: retrieve, then generate a grounded cited answer (or return context for the agent, if no chat endpoint) |
| `kb_status` | Documents vs index freshness + configuration summary (never shows keys) |

## Configuration reference

Precedence is **CLI flag > environment variable > constant in the file**.

| Env var | CLI flag | Purpose |
|---|---|---|
| `KB_DOCS_DIR` | `--docs-dir` | **Required.** Folder of `.md`/`.markdown`/`.txt` docs, searched recursively |
| `KB_INDEX_DIR` | `--index-dir` | ChromaDB folder (default `<docs-dir>\.kb-rag-index`) |
| `KB_COLLECTION` | `--collection` | ChromaDB collection name (default `kb-rag`) |
| `KB_EMBED_URL` | `--embed-url` | **Required.** Full URL of the embeddings endpoint |
| `KB_EMBED_MODEL` | `--embed-model` | Model name sent in embed requests (omit if the endpoint fixes one) |
| `KB_EMBED_API_KEY` | _(env only)_ | API key for the embeddings endpoint |
| `KB_EMBED_AUTH_HEADER` | `--embed-auth-header` | Header the key is sent in — default `Authorization` (as `Bearer <key>`); any other name (e.g. Azure's `api-key`) sends the raw key |
| `KB_EMBED_STYLE` | `--embed-style` | Request format: `openai` (default; batch `{"input": [...]}`), `ollama` (`{"prompt": ...}` one-per-request), `kserve-jina` (KServe V2 Open Inference Protocol — texts sent as a BYTES input tensor `{"inputs": [{"name", "shape", "datatype", "data"}]}`, e.g. a Jina embeddings model served on KServe; the model name is part of the `--embed-url` path such as `https://host/v2/models/jina-embeddings/infer`, and flat FP32 output tensors are reshaped via their `shape`; nested data and KServe V1 `predictions` responses also parsed), or `raw-json` (plain `{"texts": [...]}` body — for KServe **custom** predictors and other bespoke wrappers that unpack the raw request body into their pipeline's arguments; the telltale symptom is a server error like `pipeline() missing 1 required positional argument: 'texts'` that doesn't change when the tensor name does). Response parsing additionally accepts bare `embedding`/`embeddings` shapes, so most bespoke internal endpoints work unchanged |
| `KB_EMBED_TENSOR_NAME` | `--embed-tensor-name` | `kserve-jina` style only: name of the input tensor the texts are sent as (default `text`) |
| `KB_EMBED_JSON_KEY` | `--embed-json-key` | `raw-json` style only: the JSON key the batch of texts is sent under (default `texts`; e.g. `instances` for a V1-flavoured custom wrapper) |
| `KB_EMBED_TEMPLATE` | `--embed-template` | **Full request-body control** when none of the styles matches your endpoint: the complete JSON body to POST, with the JSON string `"__TEXTS__"` where the array of texts goes; `"__COUNT__"` becomes the number of texts in the batch (an integer — for strictly-validated tensor `shape` fields) and `"__MODEL__"` the `--embed-model` name. Overrides the style's request shape; batching still applies. E.g. a FastAPI custom predictor that validates `{"inputs": {"texts": [...]}}` (symptom: HTTP 422 with `loc: [body, inputs]`) is `KB_EMBED_TEMPLATE={"inputs": {"texts": "__TEXTS__"}}`; a strict V2 tensor envelope with a custom `texts` field is `{"inputs": [{"name": "texts", "shape": ["__COUNT__"], "datatype": "BYTES", "data": "__TEXTS__", "texts": "__TEXTS__"}]}`. Tip: FastAPI-wrapped endpoints publish their exact schema at `/openapi.json` (and Swagger UI at `/docs`) — fetch it with your mTLS certs instead of guessing field by field |
| `KB_EMBED_RESPONSE_PATH` | `--embed-response-path` | Dotted path to the vectors in the response when auto-detection can't find them, e.g. `outputs.embeddings` or `result.0.vectors` (numeric parts index lists). Applies to every style |
| `KB_DEBUG=1` | `--debug` | Log every request/response body (truncated) to stderr — run `--check` with it to see exactly what is sent and what came back, for matching an unknown endpoint |
| `KB_EMBED_BATCH` | `--embed-batch` | Texts per embeddings request, openai and kserve-jina styles (default 16) |
| `KB_EMBED_QUERY_PREFIX` | `--embed-query-prefix` | Prefix for query embeds, for models that need it (e5-style `"query: "`) |
| `KB_EMBED_DOC_PREFIX` | `--embed-doc-prefix` | Prefix for document embeds (`"passage: "`) |
| `KB_EMBED_EXTRA_HEADERS` | _(env only)_ | JSON object of extra HTTP headers for the embed endpoint |
| `KB_CHAT_URL` | `--chat-url` | *Optional.* Chat-completions endpoint for the generate step (OpenAI shape; Ollama `/api/chat` and `/api/generate` response shapes also parsed) |
| `KB_CHAT_MODEL` | `--chat-model` | Generation model name |
| `KB_CHAT_API_KEY` | _(env only)_ | API key for the chat endpoint (falls back to `KB_EMBED_API_KEY`) |
| `KB_CHAT_AUTH_HEADER` | `--chat-auth-header` | As per `KB_EMBED_AUTH_HEADER` |
| `KB_CHAT_MAX_TOKENS` | `--chat-max-tokens` | `max_tokens` for generation (default 1024; 0 omits the field) |
| `KB_CHAT_EXTRA_HEADERS` | _(env only)_ | JSON object of extra HTTP headers for the chat endpoint |
| `KB_CA_CERT` | `--ca-cert` | Path to a PEM CA bundle for an internal CA |
| `KB_CLIENT_CERT` | `--client-cert` | Path to a PEM client certificate, for gateways that require **mutual TLS (mTLS)** — the fix for errors like `CERTIFICATE_NOT_PROVIDED` / `certificate required`. Presented to both endpoints. Loaded once at startup, so a bad path/passphrase fails immediately |
| `KB_CLIENT_KEY` | `--client-key` | Path to the PEM private key for the client certificate; omit if the `--client-cert` file contains both cert and key |
| `KB_CLIENT_KEY_PASSWORD` | _(env only)_ | Passphrase, if the client private key is encrypted |
| `KB_VERIFY_SSL=false` | `--insecure` | Disable TLS certificate verification. Note this cannot fix an mTLS error — it controls how *you* verify the *server*, not the certificate you present to it (that's `--client-cert`) |
| `KB_TIMEOUT` | `--timeout` | HTTP timeout in seconds (default 120) |
| `KB_CHUNK_CHARS` | `--chunk-chars` | Soft max characters per chunk (default 1500) |
| `KB_CHUNK_OVERLAP` | `--chunk-overlap` | Overlap between adjacent chunks (default 200) |
| `KB_TOP_K` | `--top-k` | Default chunks retrieved (default 5) |
| — | `--check` | Validate config, call the endpoint(s) once, report index status, then exit |
| — | `--reindex` | Build/update the vector index, then exit (add `--force` to rebuild from scratch) |
| — | `--search QUERY` | Test retrieval from the command line, then exit |
| — | `--ask QUESTION` | Test full RAG (retrieve + generate) from the command line, then exit |
| — | `--version` | Print version and exit (works even without `chromadb` installed) |

## File access

Reads only inside the documents folder; writes only the vector-index folder
(default `<docs-dir>\.kb-rag-index`); network only to the endpoint(s) you
configure.

## Usage examples

1. "Using my knowledge base, can I extend my work trip by 2 days, pay for my own accommodation for the weekend, and fly back Monday?" → `kb_ask` (retrieves the travel policy's trip-extension chunks and generates a cited answer — or hands the agent the chunks to answer from, if no chat endpoint is configured)
2. "Find the parts of our policies about accommodation and per diem." → `kb_retrieve`
3. "I've added some new documents to the knowledge base folder — pick them up." → `kb_index` (incremental: only new/changed files are embedded)
4. "Is the knowledge base index up to date? What's actually indexed?" → `kb_status`
5. "Rebuild the whole knowledge base index from scratch (we switched embedding model)." → `kb_index` with `force=true`

## Feeding it

The `confluence`, `outlook` and `word` plugins each take a
**knowledge-base folder** setting. Point all of them at this plugin's documents
folder and everything you read — wiki pages, emails, Word documents — is
mirrored to Markdown there, ready for `kb_index`. The `pdf-to-md` plugin fills
the same folder from PDFs.
