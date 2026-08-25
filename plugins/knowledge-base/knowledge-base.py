#!/usr/bin/env python3
"""
knowledge-base.py (v2.1.1)
==========================

A single-file MCP (Model Context Protocol) server providing true RAG
(Retrieval-Augmented Generation) over a folder of your own markdown
documents, with real vector retrieval:

  1. INDEX    : documents are split into heading-aware chunks, each chunk is
                embedded via your embeddings API, and the vectors are stored
                in a local ChromaDB database on disk (persistent, incremental
                - only new/changed files are re-embedded).
  2. RETRIEVE : a question is embedded with the same API and the most
                semantically similar chunks are returned, with their source
                file, section heading and similarity score.
  3. GENERATE : (optional) the retrieved chunks + the question are sent to a
                chat-completions API, which writes a grounded answer citing
                its sources. If no chat endpoint is configured, kb_ask
                returns the retrieved context and the agent you're already
                talking to writes the answer instead - so generation works
                either way.
  4. CAPTURE  : kb_capture writes a new markdown note into the knowledge
                base and indexes it immediately, so work produced in a
                conversation (a research brief, a decision, a procedure)
                is searchable next time instead of being lost with the
                chat. The other servers mirror what you READ; this is how
                what you WRITE gets in. Capture only ever adds a file - it
                never edits or deletes an existing document.

Transport: newline-delimited JSON-RPC 2.0 over stdio (the standard MCP stdio
transport).

DEPENDENCY
----------
    pip install chromadb

Everything else is standard library (HTTP to your endpoints is done with
urllib - no requests/httpx needed). chromadb pulls in several packages,
some with compiled wheels; install through your pip proxy with the SAME
interpreter your MCP client launches:

    C:\\path\\to\\python.exe -m pip install chromadb

ChromaDB's anonymised telemetry is explicitly disabled in this file, so the
server makes no network calls other than to the two endpoints you configure.

CONFIGURATION
-------------
CLI flags take priority over environment variables. API keys are env-var
ONLY (command lines are visible to other local users in process listings).

| Env var                 | CLI flag             | Purpose                                                     |
|-------------------------|----------------------|-------------------------------------------------------------|
| KB_DOCS_DIR             | --docs-dir           | REQUIRED. Folder of .md/.markdown/.txt docs (recursive)     |
| KB_INDEX_DIR            | --index-dir          | ChromaDB folder (default: <docs-dir>/.kb-rag-index)         |
| KB_OUTPUT_DIR           | --output-dir         | Folder kb_capture writes new notes into (default:           |
|                         |                      | <docs-dir>/captures). MUST resolve inside --docs-dir, or    |
|                         |                      | captured notes would never be indexed                       |
| KB_COLLECTION           | --collection         | ChromaDB collection name (default: kb-rag; 3-512 chars of   |
|                         |                      | [a-zA-Z0-9._-], starting/ending alphanumeric)               |
| KB_EMBED_URL            | --embed-url          | REQUIRED. Full URL of the embeddings endpoint               |
| KB_EMBED_MODEL          | --embed-model        | Model name sent in the request (omit if endpoint has one)   |
| KB_EMBED_API_KEY        | (env only)           | API key for the embeddings endpoint                         |
| KB_EMBED_AUTH_HEADER    | --embed-auth-header  | Header the key is sent in. Default Authorization (Bearer);  |
|                         |                      | any other name (e.g. api-key for Azure) sends the raw key   |
| KB_EMBED_STYLE          | --embed-style        | Request format: openai (default), ollama, kserve-jina       |
|                         |                      | (KServe V2 inference protocol), or raw-json (plain          |
|                         |                      | {"<key>": [texts]} body, for KServe CUSTOM predictors)      |
| KB_EMBED_TENSOR_NAME    | --embed-tensor-name  | kserve-jina only: name of the input tensor (default: text)  |
| KB_EMBED_JSON_KEY       | --embed-json-key     | raw-json only: key the texts are sent under (default: texts)|
| KB_EMBED_TEMPLATE       | --embed-template     | FULL request-body control: a JSON document with "__TEXTS__" |
|                         |                      | where the array of texts goes; "__COUNT__" becomes the      |
|                         |                      | number of texts in the batch (for tensor "shape" fields)    |
|                         |                      | and "__MODEL__" the --embed-model name. When set it         |
|                         |                      | OVERRIDES --embed-style's request shape. E.g.               |
|                         |                      | {"inputs": {"texts": "__TEXTS__"}} for a FastAPI wrapper    |
|                         |                      | that nests the pipeline arguments under an "inputs" field   |
| KB_EMBED_RESPONSE_PATH  | --embed-response-path| Dotted path to the vectors in the response when the         |
|                         |                      | automatic detection can't find them, e.g.                   |
|                         |                      | "outputs.embeddings" or "result.0.vectors". Applies to      |
|                         |                      | every style; list indexes are numeric path parts            |
| KB_DEBUG=1              | --debug              | Log every request/response body (truncated) to stderr -     |
|                         |                      | shows the exact JSON sent, for matching an unknown endpoint |
| KB_EMBED_BATCH          | --embed-batch        | Texts per embeddings request (default 16; ollama is 1-by-1) |
| KB_EMBED_QUERY_PREFIX   | --embed-query-prefix | Prefix for query embeds (e5-style models: "query: ")        |
| KB_EMBED_DOC_PREFIX     | --embed-doc-prefix   | Prefix for document embeds ("passage: ")                    |
| KB_EMBED_EXTRA_HEADERS  | (env only)           | JSON object of extra HTTP headers for the embed endpoint    |
| KB_CHAT_URL             | --chat-url           | OPTIONAL. Chat-completions endpoint for the generate step   |
| KB_CHAT_MODEL           | --chat-model         | Model name for generation                                   |
| KB_CHAT_API_KEY         | (env only)           | API key for the chat endpoint (falls back to embed key)     |
| KB_CHAT_AUTH_HEADER     | --chat-auth-header   | As per KB_EMBED_AUTH_HEADER                                 |
| KB_CHAT_MAX_TOKENS      | --chat-max-tokens    | max_tokens for generation (default 1024, 0 = omit field)    |
| KB_CHAT_EXTRA_HEADERS   | (env only)           | JSON object of extra HTTP headers for the chat endpoint     |
| KB_CA_CERT              | --ca-cert            | PEM CA bundle for an internal CA                            |
| KB_CLIENT_CERT          | --client-cert        | PEM client certificate, for gateways that require mutual    |
|                         |                      | TLS (mTLS). Presented to BOTH endpoints                     |
| KB_CLIENT_KEY           | --client-key         | PEM private key for the client certificate (omit if the     |
|                         |                      | --client-cert file contains both cert and key)              |
| KB_CLIENT_KEY_PASSWORD  | (env only)           | Passphrase, if the client private key is encrypted          |
| KB_VERIFY_SSL=false     | --insecure           | Disable TLS certificate verification                        |
| KB_TIMEOUT              | --timeout            | HTTP timeout seconds (default 120)                          |
| KB_CHUNK_CHARS          | --chunk-chars        | Soft max characters per chunk (default 1500)                |
| KB_CHUNK_OVERLAP        | --chunk-overlap      | Overlap characters between adjacent chunks (default 200)    |
| KB_TOP_K                | --top-k              | Default number of chunks retrieved (default 5)              |

Endpoint formats ("where you have unknowns"):
  --embed-style openai      : POST {"input": [texts], "model": m}
                              reads response["data"][i]["embedding"]
  --embed-style ollama      : POST {"model": m, "prompt": text} (one per request)
                              reads response["embedding"]
  --embed-style kserve-jina : KServe V2 Open Inference Protocol, as used by a
                              Jina embeddings model served on KServe. POSTs
                                {"inputs": [{"name": <tensor>, "shape": [N],
                                 "datatype": "BYTES", "data": [texts]}]}
                              and reads response["outputs"][...]["data"],
                              reshaping a flat FP32 array via the tensor's
                              "shape" [N, dim] (nested data also accepted).
                              The model name is part of the URL, e.g.
                              https://host/v2/models/jina-embeddings/infer
                              so --embed-model is not sent. If your
                              deployment names its input tensor differently,
                              set --embed-tensor-name (default: text).
                              KServe V1 ({"instances": [...]} ->
                              {"predictions": [...]}) is also parsed.
  --embed-style raw-json    : plain JSON body {"texts": [texts]} (key set by
                              --embed-json-key, plus "model" if --embed-model
                              is given). For KServe CUSTOM predictors and
                              other bespoke wrappers that unpack the raw
                              request body into their pipeline's arguments -
                              the symptom that calls for this style is a
                              server error like "pipeline() missing 1
                              required positional argument: 'texts'" that
                              does NOT change when --embed-tensor-name does
                              (the wrapper reads body keys, not tensors).
                              The response is parsed with the same tolerant
                              reader as every other style (OpenAI "data",
                              "embedding"/"embeddings", KServe "outputs"/
                              "predictions").
  --embed-template          : when none of the styles matches your endpoint,
                              take full control of the request body. The
                              value is the COMPLETE JSON body to send, with
                              the JSON string "__TEXTS__" marking where the
                              array of texts goes ("__MODEL__" likewise for
                              --embed-model). Batching still applies. E.g. a
                              FastAPI custom predictor that validates
                              {"inputs": {"texts": [...]}} (the symptom is an
                              HTTP 422 with loc [body, inputs]) is:
                                KB_EMBED_TEMPLATE={"inputs": {"texts": "__TEXTS__"}}
                              "__COUNT__" is replaced by the number of texts
                              in the batch, so even a strictly-validated V2
                              tensor envelope with extra custom fields can be
                              expressed, e.g.:
                                {"inputs": [{"name": "texts",
                                 "shape": ["__COUNT__"], "datatype": "BYTES",
                                 "data": "__TEXTS__", "texts": "__TEXTS__"}]}
  Response parsing also falls back to top-level "embedding"/"embeddings",
  so most bespoke internal endpoints work with style=openai unchanged. If
  the vectors sit somewhere the auto-detection can't see (e.g. under
  {"outputs": {"embeddings": ...}}), point at them with
  --embed-response-path outputs.embeddings (numeric parts index lists).
  Set KB_DEBUG=1 / --debug to log every request and response body
  (truncated) to stderr - run --check with it to see exactly what is sent,
  and compare with what your endpoint's team expects.
  Generation POSTs OpenAI chat-completions JSON and falls back to Ollama
  /api/chat ("message"."content") and /api/generate ("response") shapes.

INSTALLING INTO CLAUDE CODE
---------------------------
This server ships as the "knowledge-base" Claude Code plugin (its manifest is
.claude-plugin/plugin.json next to this file), so the normal install is:

    /plugin marketplace add C:\\path\\to\\claude-skills
    /plugin install knowledge-base@mcnamee-claude-skills

Claude Code prompts for the documents folder, the embeddings URL, the model
name and the Python interpreter. KB_EMBED_API_KEY is NOT stored in the plugin
- set it as a Windows user environment variable before starting Claude Code,
and the plugin picks it up from there. Every other setting below is available
as its KB_* environment variable. See README.md next to this file for the full
settings reference.

FIRST RUN / TESTING (PowerShell, before wiring into the MCP client)
---------------------------------------------------------------
    $env:KB_EMBED_API_KEY = "..."       (this session only; setx makes it permanent)
    python knowledge-base.py --docs-dir C:\\kb --embed-url https://... --check
        -> validates config, calls the embeddings endpoint once, reports index status
    python knowledge-base.py --docs-dir C:\\kb --embed-url https://... --reindex
        -> builds/updates the vector index (add --force to rebuild from scratch)
    python knowledge-base.py --docs-dir C:\\kb --embed-url https://... --search "trip extension"
        -> test retrieval from the command line
    python knowledge-base.py --docs-dir C:\\kb --embed-url https://... --chat-url https://... --ask "Can I extend my trip?"
        -> test full RAG (retrieve + generate) from the command line

TOOLS EXPOSED
-------------
- kb_index    : build/update the vector index (incremental; force=true rebuilds)
- kb_retrieve : semantic search - top-k most similar chunks for a question
- kb_ask      : full RAG - retrieve, then generate a grounded, cited answer
                (or return the context for the agent to answer from, if no
                chat endpoint is configured)
- kb_capture  : write a new markdown note into the knowledge base and index
                it - for keeping something produced in the conversation
- kb_status   : index freshness + configuration summary (never shows keys)

NOTES
-----
- ALL diagnostic output goes to stderr; stdout carries only JSON-RPC.
- Set PYTHONUTF8=1 in the launching environment so non-ASCII content does not
  crash on the default Windows cp1252 codec.
- File access is confined to --docs-dir (paths are resolved, symlinks
  included, before the containment check). The index folder and dot-folders
  are never indexed. The only network calls are to the endpoints you set.
- Existing documents are never modified or deleted. The ONLY file the server
  writes outside its vector index is a new note created by kb_capture, and
  the caller supplies a TITLE, never a path: the filename is derived from it
  and the result is checked to be inside --output-dir before anything is
  written. Captured notes are the agent's own output, so they are stamped as
  such in their header - treat them as prior working notes when they come
  back from a search, not as an authoritative source.
- Retrieved document text IS sent to the configured endpoints (that's what
  RAG is) - point the server only at material appropriate for those APIs.
- If you change embedding model, the vector dimensions change: run
  --reindex --force once to rebuild the index.
- Mutual TLS: a gateway error like CERTIFICATE_NOT_PROVIDED (or a TLS
  handshake failure / connection reset during --check) means the gateway
  requires a CLIENT certificate. Configure --client-cert/--client-key;
  --insecure cannot fix this, because it only disables YOUR verification of
  the server - it does not change the certificate you present to it. The
  client certificate/key is loaded once at startup, so a bad path or wrong
  passphrase fails immediately with a clear message.
"""

# Semantic version of this server. Bump on EVERY change (see CLAUDE.md):
# MAJOR = breaking config/tool change, MINOR = new feature, PATCH = fix.
__version__ = "2.1.1"

import os
import re
import ssl
import sys
import json
import time
import shutil
import hashlib
import argparse
import datetime
import tempfile
import traceback
import urllib.error
import urllib.request


# ---------------------------------------------------------------------------
# Stream setup: force UTF-8 so non-ASCII content cannot crash output.
# ---------------------------------------------------------------------------
for _stream in ("stdin", "stdout"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8")
    except Exception:
        pass


def log(message):
    """Write a diagnostic line to stderr ONLY. Never touch stdout here."""
    print(message, file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# Configuration (populated in main(); flags take priority over env vars)
# ---------------------------------------------------------------------------

DOC_EXTENSIONS = {".md", ".markdown", ".txt"}

# Hard cap on how much retrieved context is stuffed into a generation prompt.
MAX_CONTEXT_CHARS = 16000


class Config(object):
    """All runtime settings in one place. Filled in by main()."""
    docs_dir = None            # absolute, real path of the documents folder
    index_dir = None           # absolute path of the ChromaDB folder
    output_dir = None          # absolute path kb_capture writes notes into
    collection = "kb-rag"
    embed_url = None
    embed_model = ""
    embed_key = ""
    embed_auth_header = "Authorization"
    embed_style = "openai"
    embed_tensor_name = "text"
    embed_json_key = "texts"
    embed_template = None      # parsed JSON template, or None
    embed_response_path = ""
    debug = False
    embed_batch = 16
    embed_query_prefix = ""
    embed_doc_prefix = ""
    embed_extra_headers = {}
    chat_url = ""
    chat_model = ""
    chat_key = ""
    chat_auth_header = "Authorization"
    chat_max_tokens = 1024
    chat_extra_headers = {}
    ca_cert = ""
    client_cert = ""
    client_key = ""
    client_key_password = ""
    verify_ssl = True
    timeout = 120
    chunk_chars = 1500
    chunk_overlap = 200
    top_k = 5


CFG = Config()


class RagError(Exception):
    """A failure with a message meant to be shown to the caller as-is."""


# ---------------------------------------------------------------------------
# Filesystem helpers (same confinement model as the other servers)
# ---------------------------------------------------------------------------

def is_within(path, base):
    """
    True if `path` resolves to a location inside `base`. Guards against path
    traversal and symlinks pointing outside the configured folder.
    """
    try:
        real_path = os.path.realpath(path)
        real_base = os.path.realpath(base)
        return os.path.commonpath([real_path, real_base]) == real_base
    except Exception:
        return False


def to_rel(path):
    """
    Relative path from docs_dir, using forward slashes for stable display.
    Falls back to the absolute path when the two are on different Windows
    drives, where os.path.relpath raises rather than returning something.
    """
    try:
        return os.path.relpath(path, CFG.docs_dir).replace("\\", "/")
    except ValueError:
        return path.replace("\\", "/")


def scan_documents():
    """
    Return {relative_path: absolute_path} for every document in the folder.
    Dot-folders and the index folder are pruned; anything resolving outside
    the docs folder (e.g. a symlink) is excluded.
    """
    found = {}
    index_real = os.path.realpath(CFG.index_dir) if CFG.index_dir else None
    for root, dirs, files in os.walk(CFG.docs_dir):
        # Prune hidden folders and the vector index itself from the walk.
        dirs[:] = [
            d for d in dirs
            if not d.startswith(".")
            and os.path.realpath(os.path.join(root, d)) != index_real
        ]
        for filename in files:
            if filename.startswith("."):
                continue
            if os.path.splitext(filename)[1].lower() not in DOC_EXTENSIONS:
                continue
            full = os.path.join(root, filename)
            if not is_within(full, CFG.docs_dir):
                log("Excluded (resolves outside the docs folder): {0}".format(full))
                continue
            found[to_rel(full)] = full
    return found


def safe_filename(name, max_len=150):
    """
    Turn a caller-supplied title into a filesystem-safe filename component (no
    extension). Same helper the confluence/outlook/word mirrors use, so captured
    notes and mirrored documents are named consistently.

    Strips characters that are illegal on Windows (< > : " / \\ | ? * and control
    chars), collapses whitespace, removes trailing dots/spaces (also illegal on
    Windows), and caps the length. Returns '' if nothing usable is left - the
    caller decides what to do about that.

    Because the separators are stripped rather than escaped, a title such as
    "..\\..\\secrets" collapses to "secrets": there is no way to walk out of the
    folder with a title. The containment check in capture_note() is the belt to
    this braces.
    """
    if not name:
        return ""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    cleaned = " ".join(cleaned.split()).strip(" .")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" .")
    return cleaned


def capture_note(title, content, source, tags, overwrite):
    """
    Write a new markdown note into the capture folder and return its absolute
    path. Raises RagError with a caller-facing message on any refusal.

    The header block matches the one the confluence/outlook/word mirrors write,
    so a captured note and a mirrored document read the same way once indexed.
    The "written by Claude" stamp is deliberate: these notes come back from
    later searches, and the reader has to be able to tell agent output from an
    authoritative document.
    """
    stem = safe_filename(title)
    if not stem:
        raise RagError(
            "'title' has no characters usable in a filename. Give a short "
            "descriptive title, e.g. 'Records retention thresholds'.")
    prefix = safe_filename(source, max_len=40) or "Note"

    path = os.path.join(CFG.output_dir, "{0} - {1}.md".format(prefix, stem))
    try:
        os.makedirs(CFG.output_dir, exist_ok=True)
    except OSError as exc:
        raise RagError("Could not create the capture folder {0}: {1}".format(
            CFG.output_dir, exc))

    # Belt and braces: the title cannot contain a separator by the time it gets
    # here, but resolve the path and confirm it really is inside the capture
    # folder before writing. This also catches a pre-existing symlink of that
    # name pointing somewhere else.
    if not is_within(path, CFG.output_dir):
        raise RagError("Refused: {0} resolves outside the capture folder.".format(
            os.path.basename(path)))
    if os.path.exists(path) and not overwrite:
        raise RagError(
            "A note already exists at {0}. Pass overwrite=true to replace it, "
            "or use a different title.".format(to_rel(path)))

    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    header = [
        # Collapse whitespace: a newline in the title would otherwise split the
        # H1 in two and push the rest of the header into the body.
        "# {0}".format(" ".join(title.split())),
        "",
        "- Source: {0} (written by Claude in conversation, not an "
        "authoritative document)".format(prefix),
        "- Captured: {0}".format(stamp),
    ]
    if tags:
        header.append("- Tags: {0}".format(", ".join(tags)))
    header += ["", "---", ""]

    try:
        # newline="\n" keeps line endings consistent and avoids CRLF doubling
        # on Windows, matching the other servers' mirrors.
        with open(path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write("\n".join(header) + "\n" + content.strip() + "\n")
    except OSError as exc:
        raise RagError("Could not write the note to {0}: {1}".format(path, exc))
    return path


def read_text(path):
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        return handle.read()


def file_hash(path):
    """SHA-256 of a file's bytes - used to detect changed documents."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(65536), b""):
            digest.update(block)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# HTTP plumbing (stdlib only)
# ---------------------------------------------------------------------------

def build_ssl_context():
    if CFG.ca_cert and CFG.verify_ssl:
        context = ssl.create_default_context(cafile=CFG.ca_cert)
    else:
        context = ssl.create_default_context()
    if not CFG.verify_ssl:
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
    if CFG.client_cert:
        # Mutual TLS: present our certificate to the gateway. Independent of
        # verify_ssl, which only controls how we verify the SERVER.
        try:
            context.load_cert_chain(
                certfile=CFG.client_cert,
                keyfile=CFG.client_key or None,
                password=CFG.client_key_password or None,
            )
        except (ssl.SSLError, OSError) as exc:
            raise RagError(
                "Could not load the mutual-TLS client certificate/key "
                "({0} / {1}): {2}. If the key is encrypted, set "
                "KB_CLIENT_KEY_PASSWORD.".format(
                    CFG.client_cert,
                    CFG.client_key or "key expected in the cert file",
                    exc))
    return context


def auth_headers(key, header_name):
    """Authorization -> 'Bearer <key>'; any other header carries the raw key."""
    if not key:
        return {}
    if header_name.strip().lower() == "authorization":
        return {"Authorization": "Bearer " + key}
    return {header_name.strip(): key}


def _mtls_hint(error):
    """Append advice when a TLS failure looks like a missing client certificate."""
    text = str(error)
    if ("CERTIFICATE_REQUIRED" in text or "certificate required" in text.lower()
            or "CERTIFICATE_NOT_PROVIDED" in text or "handshake failure" in text.lower()):
        return (" This looks like the endpoint requires a CLIENT certificate "
                "(mutual TLS): configure --client-cert / --client-key. "
                "--insecure cannot fix this - it only disables verification "
                "of the server, not the certificate you present.")
    return ""


def _debug_log(label, text):
    """When --debug / KB_DEBUG=1 is on, log a (truncated) payload to stderr."""
    if CFG.debug:
        text = text if len(text) <= 2000 else text[:2000] + " ...[truncated]"
        log("DEBUG {0}: {1}".format(label, text))


def http_post_json(url, payload, headers):
    """POST JSON, return the decoded JSON response. Raises RagError on failure."""
    body = json.dumps(payload).encode("utf-8")
    _debug_log("POST " + url, body.decode("utf-8"))
    request = urllib.request.Request(url, data=body, method="POST")
    request.add_header("Content-Type", "application/json")
    request.add_header("Accept", "application/json")
    for name, value in headers.items():
        request.add_header(name, value)
    try:
        with urllib.request.urlopen(request, timeout=CFG.timeout,
                                    context=build_ssl_context()) as response:
            raw = response.read().decode("utf-8", errors="replace")
            _debug_log("RESPONSE " + url, raw)
            return json.loads(raw)
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
        except Exception:
            pass
        raise RagError("HTTP {0} from {1}: {2}{3}".format(
            exc.code, url, detail or exc.reason, _mtls_hint(detail or exc.reason)))
    except urllib.error.URLError as exc:
        raise RagError("Could not reach {0}: {1}{2}".format(
            url, exc.reason, _mtls_hint(exc.reason)))
    except OSError as exc:
        # TLS alerts can surface on the first read (TLS 1.3), outside
        # urllib's URLError wrapping - e.g. a gateway rejecting the handshake
        # because no client certificate was presented.
        raise RagError("Connection to {0} failed: {1}{2}".format(
            url, exc, _mtls_hint(exc)))
    except json.JSONDecodeError:
        raise RagError("Non-JSON response from {0}.".format(url))


# ---------------------------------------------------------------------------
# Embeddings client
# ---------------------------------------------------------------------------

TEMPLATE_TEXTS = "__TEXTS__"
TEMPLATE_MODEL = "__MODEL__"
TEMPLATE_COUNT = "__COUNT__"


def fill_template(node, texts):
    """
    Deep-copy the request template, substituting the "__TEXTS__" placeholder
    with the list of texts, "__COUNT__" with how many texts there are (an
    integer - for tensor "shape" fields), and "__MODEL__" with the
    configured model name.
    """
    if isinstance(node, dict):
        return {key: fill_template(value, texts) for key, value in node.items()}
    if isinstance(node, list):
        return [fill_template(value, texts) for value in node]
    if node == TEMPLATE_TEXTS:
        return texts
    if node == TEMPLATE_COUNT:
        return len(texts)
    if node == TEMPLATE_MODEL:
        return CFG.embed_model
    return node


def count_template_placeholders(node):
    """How many times the "__TEXTS__" placeholder appears in a template."""
    if isinstance(node, dict):
        return sum(count_template_placeholders(v) for v in node.values())
    if isinstance(node, list):
        return sum(count_template_placeholders(v) for v in node)
    return 1 if node == TEMPLATE_TEXTS else 0


def extract_by_path(data, path):
    """
    Follow a dotted path into a JSON structure: dict parts are keys, numeric
    parts index lists ("outputs.embeddings", "result.0.vectors"). Returns
    None if any step is missing.
    """
    node = data
    for part in path.split("."):
        if isinstance(node, list):
            try:
                node = node[int(part)]
            except (ValueError, IndexError):
                return None
        elif isinstance(node, dict) and part in node:
            node = node[part]
        else:
            return None
    return node


def _vectors_from_tensor(raw, shape, expected):
    """
    Turn a KServe output tensor's `data` into a list of vectors. `raw` is
    either already nested ([[...], [...]]) or a flat float array that must be
    reshaped using the tensor's `shape` (e.g. [N, dim]).
    """
    if not isinstance(raw, list) or not raw:
        return None
    if isinstance(raw[0], list):
        return raw
    # Flat array: work out the vector length, preferring the declared shape.
    if isinstance(shape, list) and len(shape) >= 2 and shape[-1]:
        dim = int(shape[-1])
    elif len(raw) % expected == 0:
        dim = len(raw) // expected
    else:
        return None
    if dim <= 0 or len(raw) % dim != 0:
        return None
    return [raw[i:i + dim] for i in range(0, len(raw), dim)]


def _parse_kserve_response(data, expected):
    """
    Pull `expected` vectors out of a KServe response. V2 Open Inference
    Protocol: {"outputs": [{"name", "shape", "datatype", "data"}]} where
    `data` is typically a flat FP32 array reshaped via `shape`. V1 fallback:
    {"predictions": [[...], ...]}. Returns None if nothing matches (the
    caller raises the descriptive error).
    """
    if not isinstance(data, dict):
        return None
    outputs = data.get("outputs")
    if isinstance(outputs, list):
        # More than one output tensor is possible (e.g. token-level plus
        # pooled embeddings); take the first one that yields `expected` rows.
        for output in outputs:
            if not isinstance(output, dict):
                continue
            vectors = _vectors_from_tensor(
                output.get("data"), output.get("shape"), expected)
            if vectors is not None and len(vectors) == expected:
                return vectors
    predictions = data.get("predictions")
    if isinstance(predictions, list):
        return _vectors_from_tensor(predictions, None, expected)
    return None


def _parse_embedding_response(data, expected):
    """
    Pull `expected` vectors out of an embeddings response, tolerating the
    common shapes: OpenAI {"data":[{"embedding":[...]}]}, bare {"embedding":
    [...]}, {"embeddings":[[...]]}, and KServe V2/V1 tensor responses.
    """
    vectors = None
    if CFG.embed_response_path:
        # Explicit location wins over auto-detection.
        node = extract_by_path(data, CFG.embed_response_path)
        if isinstance(node, list) and node and isinstance(node[0], dict):
            vectors = [item.get("embedding") for item in node]
        else:
            vectors = _vectors_from_tensor(node, None, expected)
    elif isinstance(data, dict):
        if isinstance(data.get("data"), list):
            items = sorted(data["data"], key=lambda item: item.get("index", 0))
            vectors = [item.get("embedding") for item in items]
        elif isinstance(data.get("embeddings"), list):
            vectors = data["embeddings"]
        elif isinstance(data.get("embedding"), list):
            vectors = [data["embedding"]]
        elif "outputs" in data or "predictions" in data:
            vectors = _parse_kserve_response(data, expected)
    if (not vectors or len(vectors) != expected
            or any(not isinstance(v, list) or not v for v in vectors)):
        raise RagError(
            "Unexpected embeddings response shape (expected {0} vector(s)){1}. "
            "Check --embed-style / the endpoint URL, or point at the vectors "
            "with --embed-response-path; run --check with --debug to see the "
            "raw response. Response keys: {2}".format(
                expected,
                " at response path '{0}'".format(CFG.embed_response_path)
                if CFG.embed_response_path else "",
                list(data.keys()) if isinstance(data, dict) else type(data).__name__))
    return vectors


def embed_texts(texts, is_query=False):
    """
    Embed a list of strings via the configured endpoint; returns one vector
    per input, in order. Batching applies to the openai style; the ollama
    style is one request per text (its classic API takes a single prompt).
    """
    prefix = CFG.embed_query_prefix if is_query else CFG.embed_doc_prefix
    inputs = [prefix + text for text in texts]
    headers = dict(CFG.embed_extra_headers)
    headers.update(auth_headers(CFG.embed_key, CFG.embed_auth_header))

    vectors = []
    if CFG.embed_template is not None:
        # A request template overrides the style's request shape entirely.
        for start in range(0, len(inputs), CFG.embed_batch):
            batch = inputs[start:start + CFG.embed_batch]
            payload = fill_template(CFG.embed_template, batch)
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, len(batch)))
    elif CFG.embed_style == "ollama":
        for text in inputs:
            payload = {"prompt": text}
            if CFG.embed_model:
                payload["model"] = CFG.embed_model
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, 1))
    elif CFG.embed_style == "kserve-jina":
        # KServe V2 Open Inference Protocol: texts travel as one BYTES input
        # tensor. The model name lives in the URL (/v2/models/<name>/infer),
        # so no model field is sent. The content_type parameter is the V2
        # string-codec hint (needed by MLServer-based deployments, ignored
        # by others).
        for start in range(0, len(inputs), CFG.embed_batch):
            batch = inputs[start:start + CFG.embed_batch]
            payload = {
                "inputs": [{
                    "name": CFG.embed_tensor_name,
                    "shape": [len(batch)],
                    "datatype": "BYTES",
                    "parameters": {"content_type": "str"},
                    "data": batch,
                }],
            }
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, len(batch)))
    elif CFG.embed_style == "raw-json":
        # Plain JSON body for custom predictors that unpack the request dict
        # straight into their pipeline's arguments (no tensor envelope).
        for start in range(0, len(inputs), CFG.embed_batch):
            batch = inputs[start:start + CFG.embed_batch]
            payload = {CFG.embed_json_key: batch}
            if CFG.embed_model:
                payload["model"] = CFG.embed_model
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, len(batch)))
    else:  # openai-compatible (the default)
        for start in range(0, len(inputs), CFG.embed_batch):
            batch = inputs[start:start + CFG.embed_batch]
            payload = {"input": batch}
            if CFG.embed_model:
                payload["model"] = CFG.embed_model
            data = http_post_json(CFG.embed_url, payload, headers)
            vectors.extend(_parse_embedding_response(data, len(batch)))
    return vectors


# ---------------------------------------------------------------------------
# Generation client (OpenAI chat-completions shape, with Ollama fallbacks)
# ---------------------------------------------------------------------------

GENERATION_SYSTEM_PROMPT = (
    "You are a careful assistant answering questions from a personal knowledge "
    "base. Answer ONLY from the provided context. Quote or paraphrase the "
    "relevant passages and cite the source file for each claim, e.g. "
    "[travel-policy.md]. If the context does not contain the answer, say so "
    "plainly - do not invent one."
)


def generate_answer(question, context):
    """Send the retrieved context + question to the chat endpoint; return the answer text."""
    headers = dict(CFG.chat_extra_headers)
    headers.update(auth_headers(CFG.chat_key, CFG.chat_auth_header))
    payload = {
        "messages": [
            {"role": "system", "content": GENERATION_SYSTEM_PROMPT},
            {"role": "user",
             "content": "Context:\n\n{0}\n\nQuestion: {1}".format(context, question)},
        ],
        "temperature": 0.1,
        "stream": False,
    }
    if CFG.chat_model:
        payload["model"] = CFG.chat_model
    if CFG.chat_max_tokens > 0:
        payload["max_tokens"] = CFG.chat_max_tokens

    data = http_post_json(CFG.chat_url, payload, headers)

    # OpenAI shape, then Ollama /api/chat, then Ollama /api/generate.
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError):
        pass
    try:
        return data["message"]["content"]
    except (KeyError, TypeError):
        pass
    if isinstance(data.get("response"), str):
        return data["response"]
    raise RagError(
        "Unexpected chat response shape. Response keys: {0}".format(
            list(data.keys()) if isinstance(data, dict) else type(data).__name__))


# ---------------------------------------------------------------------------
# Markdown chunking (heading-aware, code-fence-aware)
# ---------------------------------------------------------------------------

# ATX heading, allowing up to 3 leading spaces and trailing '#'s (per CommonMark).
HEADING_RE = re.compile(r"^ {0,3}(#{1,6})\s+(.+?)\s*#*\s*$")


def parse_headings(content):
    """
    Return [{level, text, line}, ...] for every ATX heading, in document
    order. Lines inside fenced code blocks are ignored so a '#' comment in a
    code sample is not mistaken for a heading.
    """
    heads = []
    in_fence = False
    fence = None
    for idx, raw in enumerate(content.splitlines()):
        stripped = raw.strip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            marker = stripped[:3]
            if not in_fence:
                in_fence, fence = True, marker
            elif stripped.startswith(fence):
                in_fence, fence = False, None
            continue
        if in_fence:
            continue
        match = HEADING_RE.match(raw)
        if match:
            heads.append({"level": len(match.group(1)),
                          "text": match.group(2).strip(), "line": idx})
    return heads


def split_text(text, size, overlap):
    """
    Split `text` into pieces of at most ~`size` characters, preferring
    paragraph boundaries, with ~`overlap` characters carried between adjacent
    pieces so a sentence cut at a boundary is still retrievable. `size` is a
    soft target: a carried tail can push a piece slightly over it.
    """
    if len(text) <= size:
        return [text]

    # Break into paragraph units, hard-splitting any single oversized paragraph.
    units = []
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        while len(para) > size:
            units.append(para[:size])
            para = para[size - overlap:]
        if para:
            units.append(para)

    pieces = []
    current = ""
    for unit in units:
        if current and len(current) + 2 + len(unit) > size:
            pieces.append(current)
            tail = current[-overlap:].strip() if overlap else ""
            current = (tail + "\n\n" + unit) if tail else unit
        else:
            current = (current + "\n\n" + unit) if current else unit
    if current:
        pieces.append(current)
    return pieces


def chunk_document(content):
    """
    Split a markdown document into chunks along its heading structure:
    each heading's section becomes one or more chunks, each tagged with the
    full heading path (e.g. 'Travel Policy > Expenses > Per diem') so the
    embedding and the retrieved result both carry that context.

    Returns [{"heading": path, "text": chunk_text}, ...].
    """
    heads = parse_headings(content)
    lines = content.splitlines()

    sections = []  # (heading_path, start_line, end_line)
    if heads:
        if heads[0]["line"] > 0:
            sections.append(("", 0, heads[0]["line"]))
        stack = []  # [(level, text), ...] - the open headings above this point
        for i, head in enumerate(heads):
            while stack and stack[-1][0] >= head["level"]:
                stack.pop()
            stack.append((head["level"], head["text"]))
            end = heads[i + 1]["line"] if i + 1 < len(heads) else len(lines)
            sections.append((" > ".join(t for _lvl, t in stack), head["line"], end))
    else:
        sections.append(("", 0, len(lines)))

    chunks = []
    for heading_path, start, end in sections:
        section_text = "\n".join(lines[start:end]).strip()
        if not section_text:
            continue
        for piece in split_text(section_text, CFG.chunk_chars, CFG.chunk_overlap):
            chunks.append({"heading": heading_path, "text": piece})
    return chunks


# ---------------------------------------------------------------------------
# Vector index (ChromaDB)
# ---------------------------------------------------------------------------

_CHROMA_CLIENT = None


def chroma_client():
    """Create (once) and return the persistent ChromaDB client, telemetry off."""
    global _CHROMA_CLIENT
    if _CHROMA_CLIENT is not None:
        return _CHROMA_CLIENT
    try:
        import chromadb
        from chromadb.config import Settings
    except ImportError:
        raise RagError(
            "chromadb is not installed in this interpreter ({0}). "
            "Install it with:  {0} -m pip install chromadb".format(sys.executable))
    _CHROMA_CLIENT = chromadb.PersistentClient(
        path=CFG.index_dir,
        settings=Settings(anonymized_telemetry=False),
    )
    return _CHROMA_CLIENT


def get_collection(reset=False):
    """Return the vector collection (cosine space), optionally dropping it first."""
    client = chroma_client()
    if reset:
        try:
            client.delete_collection(CFG.collection)
        except Exception:
            pass  # nothing to drop on first run
    return client.get_or_create_collection(
        name=CFG.collection, metadata={"hnsw:space": "cosine"})


def indexed_file_hashes(collection):
    """Return {source_relpath: file_hash} for everything currently indexed."""
    hashes = {}
    offset = 0
    while True:
        page = collection.get(include=["metadatas"], limit=1000, offset=offset)
        metadatas = page.get("metadatas") or []
        if not metadatas:
            break
        for meta in metadatas:
            if meta and "source" in meta:
                hashes[meta["source"]] = meta.get("file_hash", "")
        if len(metadatas) < 1000:
            break
        offset += len(metadatas)
    return hashes


def index_sync(force=False):
    """
    Bring the vector index in line with the docs folder. Only new or changed
    files are re-embedded; files deleted from the folder are removed from the
    index. force=True drops the collection and rebuilds everything (needed
    after changing embedding model, since vector dimensions change).

    Returns a stats dict.
    """
    started = time.time()
    collection = get_collection(reset=force)
    documents = scan_documents()
    indexed = {} if force else indexed_file_hashes(collection)

    removed = [rel for rel in indexed if rel not in documents]
    for rel in removed:
        collection.delete(where={"source": rel})

    unchanged = 0
    updated = []
    errors = []
    total_chunks = 0
    for rel, path in sorted(documents.items()):
        try:
            digest = file_hash(path)
        except OSError as exc:
            errors.append("{0}: {1}".format(rel, exc))
            continue
        if indexed.get(rel) == digest:
            unchanged += 1
            continue

        try:
            content = read_text(path)
        except OSError as exc:
            errors.append("{0}: {1}".format(rel, exc))
            continue
        chunks = chunk_document(content)
        if rel in indexed:
            collection.delete(where={"source": rel})
        if not chunks:
            updated.append(rel)
            continue

        # Embed with the heading path prepended for context; store the raw
        # chunk text so what the model reads back is the document itself.
        embed_inputs = [
            ("{0} — {1}\n\n{2}".format(rel, c["heading"], c["text"])
             if c["heading"] else "{0}\n\n{1}".format(rel, c["text"]))
            for c in chunks
        ]
        try:
            vectors = embed_texts(embed_inputs, is_query=False)
        except RagError as exc:
            # Surface an embedding-dimension mismatch as the fix, not a mystery.
            raise RagError(
                "Embedding failed while indexing '{0}': {1}".format(rel, exc))

        try:
            collection.add(
                ids=["{0}#{1}".format(rel, i) for i in range(len(chunks))],
                embeddings=vectors,
                documents=[c["text"] for c in chunks],
                metadatas=[{
                    "source": rel,
                    "heading": c["heading"],
                    "chunk": i,
                    "file_hash": digest,
                } for i, c in enumerate(chunks)],
            )
        except Exception as exc:
            raise RagError(
                "Storing vectors for '{0}' failed: {1}. If you changed "
                "embedding model, rebuild with kb_index(force=true) or "
                "--reindex --force (vector dimensions differ between models)."
                .format(rel, exc))
        updated.append(rel)
        total_chunks += len(chunks)
        log("Indexed {0} ({1} chunks)".format(rel, len(chunks)))

    return {
        "documents": len(documents),
        "updated": updated,
        "unchanged": unchanged,
        "removed": removed,
        "new_chunks": total_chunks,
        "index_chunks": collection.count(),
        "errors": errors,
        "seconds": round(time.time() - started, 1),
    }


def retrieve(query, top_k):
    """
    Embed `query` and return the top_k most similar chunks:
    [{"source", "heading", "text", "similarity"}, ...] best first.
    """
    collection = get_collection()
    count = collection.count()
    if count == 0:
        raise RagError(
            "The vector index is empty. Run kb_index first (or launch once "
            "with --reindex).")
    vector = embed_texts([query], is_query=True)[0]
    result = collection.query(
        query_embeddings=[vector],
        n_results=min(top_k, count),
        include=["documents", "metadatas", "distances"],
    )
    hits = []
    for text, meta, distance in zip(result["documents"][0],
                                    result["metadatas"][0],
                                    result["distances"][0]):
        hits.append({
            "source": meta.get("source", "?"),
            "heading": meta.get("heading", ""),
            "text": text,
            # cosine distance in [0,2] -> similarity in [-1,1]
            "similarity": round(1.0 - distance, 3),
        })
    return hits


def format_hits(hits):
    """Human/agent-readable rendering of retrieved chunks, best first."""
    blocks = []
    for rank, hit in enumerate(hits, 1):
        where = hit["source"] + (" · " + hit["heading"] if hit["heading"] else "")
        blocks.append("[{0}] {1}  (similarity {2})\n{3}".format(
            rank, where, hit["similarity"], hit["text"]))
    return "\n\n---\n\n".join(blocks)


def build_context(hits):
    """Concatenate hits into a generation context, capped at MAX_CONTEXT_CHARS."""
    parts = []
    total = 0
    for hit in hits:
        where = hit["source"] + (" · " + hit["heading"] if hit["heading"] else "")
        block = "[Source: {0}]\n{1}".format(where, hit["text"])
        if parts and total + len(block) > MAX_CONTEXT_CHARS:
            break
        parts.append(block)
        total += len(block)
    return "\n\n".join(parts), len(parts)


# ---------------------------------------------------------------------------
# Tool implementations (each returns a human-readable text string)
# ---------------------------------------------------------------------------

def _arg_int(args, name, default):
    try:
        return int(args.get(name, default))
    except (TypeError, ValueError):
        return default


def tool_kb_index(args):
    force = bool(args.get("force", False))
    stats = index_sync(force=force)
    lines = [
        "Index {0} in {1}s.".format("rebuilt" if force else "updated", stats["seconds"]),
        "- documents in folder : {0}".format(stats["documents"]),
        "- re-indexed          : {0}".format(len(stats["updated"])),
        "- unchanged (skipped) : {0}".format(stats["unchanged"]),
        "- removed from index  : {0}".format(len(stats["removed"])),
        "- chunks in index     : {0}".format(stats["index_chunks"]),
    ]
    if stats["updated"]:
        lines.append("Re-indexed files:\n" + "\n".join(
            "  - " + rel for rel in stats["updated"][:30]))
        if len(stats["updated"]) > 30:
            lines.append("  ... and {0} more".format(len(stats["updated"]) - 30))
    if stats["errors"]:
        lines.append("Unreadable files (skipped):\n" + "\n".join(
            "  - " + err for err in stats["errors"]))
    return "\n".join(lines)


def tool_kb_retrieve(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required."
    top_k = max(1, min(20, _arg_int(args, "top_k", CFG.top_k)))
    hits = retrieve(query, top_k)
    return (
        "Top {0} chunk(s) for '{1}', most similar first. Cite the source file "
        "when using them.\n\n{2}".format(len(hits), query, format_hits(hits)))


def tool_kb_ask(args):
    question = (args.get("question") or "").strip()
    if not question:
        return "Error: 'question' is required."
    top_k = max(1, min(20, _arg_int(args, "top_k", CFG.top_k)))
    hits = retrieve(question, top_k)
    context, used = build_context(hits)

    if not CFG.chat_url:
        # No generation endpoint: hand the agent the context to answer from.
        # Phrased as a plain instruction rather than a missing-config warning -
        # this is a supported mode (and the better one when the caller is an
        # agent), and the user sees this text on the front of every kb_ask.
        return (
            "Retrieved {0} chunk(s) for this question. Answer it STRICTLY "
            "from the context below, citing the source file for each claim. "
            "If the context doesn't contain the answer, say so.\n\n"
            "Question: {1}\n\n{2}".format(
                len(hits), question, format_hits(hits)))

    answer = generate_answer(question, context)
    sources = []
    for hit in hits[:used]:
        entry = hit["source"] + (" · " + hit["heading"] if hit["heading"] else "")
        if entry not in sources:
            sources.append(entry)
    return "{0}\n\nSources ({1} chunk(s) retrieved):\n{2}".format(
        answer, used, "\n".join("- " + s for s in sources))


def _arg_str_list(args, name):
    """Read an optional list-of-strings argument, tolerating a single string."""
    raw = args.get(name)
    if not raw:
        return []
    if isinstance(raw, str):
        raw = [raw]
    if not isinstance(raw, list):
        return []
    return [str(item).strip() for item in raw if str(item).strip()]


def tool_kb_capture(args):
    title = (args.get("title") or "").strip()
    content = (args.get("content") or "").strip()
    if not title:
        return "Error: 'title' is required."
    if not content:
        return "Error: 'content' is required - there is nothing to capture."

    source = (args.get("source") or "Note").strip() or "Note"
    tags = _arg_str_list(args, "tags")
    overwrite = bool(args.get("overwrite", False))
    do_index = args.get("index", True)
    do_index = True if do_index is None else bool(do_index)

    path = capture_note(title, content, source, tags, overwrite)
    rel = to_rel(path)
    lines = ["Captured to the knowledge base: {0}".format(rel)]

    if not do_index:
        lines.append("Not indexed - run kb_index to make it searchable.")
        return "\n".join(lines)

    # Index straight away so the note is retrievable in this same conversation.
    # A failure here is worth reporting but must not read as "nothing saved":
    # the file is on disk either way, and kb_index will pick it up later.
    try:
        stats = index_sync(force=False)
        lines.append(
            "Indexed. The knowledge base now holds {0} chunk(s) across {1} "
            "document(s).".format(stats["index_chunks"], stats["documents"]))
    except RagError as exc:
        lines.append("Saved, but indexing FAILED ({0}). The file is on disk - "
                     "run kb_index once the endpoint is reachable.".format(exc))
    return "\n".join(lines)


def tool_kb_status(_args):
    documents = scan_documents()
    captured = 0
    if CFG.output_dir and os.path.isdir(CFG.output_dir):
        captured = len([name for name in os.listdir(CFG.output_dir)
                        if os.path.splitext(name)[1].lower() in DOC_EXTENSIONS])
    lines = [
        "Knowledge base folder : {0}".format(CFG.docs_dir),
        "Vector index folder   : {0}".format(CFG.index_dir),
        "Captures folder       : {0} ({1} note(s))".format(CFG.output_dir, captured),
        "Documents in folder   : {0}".format(len(documents)),
    ]
    try:
        collection = get_collection()
        indexed = indexed_file_hashes(collection)
        stale = [rel for rel, path in documents.items()
                 if indexed.get(rel) != file_hash(path)]
        removed = [rel for rel in indexed if rel not in documents]
        lines.append("Files indexed         : {0}".format(len(indexed)))
        lines.append("Chunks in index       : {0}".format(collection.count()))
        if stale or removed:
            lines.append(
                "STALE: {0} file(s) new/changed, {1} deleted - run kb_index "
                "to refresh.".format(len(stale), len(removed)))
        else:
            lines.append("Index is up to date with the folder.")
    except Exception as exc:
        lines.append("Index status          : unavailable ({0})".format(exc))
    style = CFG.embed_style
    if CFG.embed_template is not None:
        style += ", request template set"
    elif style == "kserve-jina":
        style += ", tensor: " + CFG.embed_tensor_name
    elif style == "raw-json":
        style += ", json key: " + CFG.embed_json_key
    if CFG.embed_response_path:
        style += ", response path: " + CFG.embed_response_path
    lines.append("Embeddings endpoint   : {0} (style: {1}, model: {2}, key: {3})".format(
        CFG.embed_url, style, CFG.embed_model or "(none)",
        "set" if CFG.embed_key else "NOT SET"))
    lines.append("Mutual TLS (client)   : {0}".format(
        "cert: {0}, key: {1}, passphrase: {2}".format(
            CFG.client_cert, CFG.client_key or "(in cert file)",
            "set" if CFG.client_key_password else "not set")
        if CFG.client_cert else "(none)"))
    lines.append("Generation endpoint   : {0}".format(
        "{0} (model: {1}, key: {2})".format(
            CFG.chat_url, CFG.chat_model or "(none)",
            "set" if CFG.chat_key else "NOT SET")
        if CFG.chat_url else "(none - kb_ask returns context for the agent to answer from)"))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# MCP tool registry
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "kb_index",
        "description": (
            "Build or update the vector index of the knowledge base. Incremental: "
            "only new or changed files are re-embedded, and deleted files are "
            "removed. Run this after documents change, or when kb_status reports "
            "the index is stale. Set force=true to rebuild from scratch (required "
            "after changing embedding model)."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "force": {
                    "type": "boolean",
                    "description": "Drop the index and re-embed everything (default false).",
                },
            },
        },
    },
    {
        "name": "kb_retrieve",
        "description": (
            "Semantic (vector) search: returns the chunks of the knowledge base "
            "most similar in MEANING to the query, with source file, section "
            "heading and similarity score. Use this to pull the relevant policy/"
            "notes passages for a question, then answer from them, citing the "
            "source files."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "The question or topic, in natural language.",
                },
                "top_k": {
                    "type": "integer",
                    "description": "Number of chunks to return, 1-20 (default {0}).".format(CFG.top_k),
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "kb_ask",
        "description": (
            "Full RAG in one call: retrieves the most relevant chunks for the "
            "question and generates a grounded answer that cites its source files. "
            "If no generation endpoint is configured, it returns the retrieved "
            "context and YOU write the answer from it."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": "The question to answer from the knowledge base.",
                },
                "top_k": {
                    "type": "integer",
                    "description": "Number of chunks to retrieve, 1-20 (default {0}).".format(CFG.top_k),
                },
            },
            "required": ["question"],
        },
    },
    {
        "name": "kb_capture",
        "description": (
            "Save something into the knowledge base as a new markdown note, and "
            "index it so it is searchable immediately. This is how work produced "
            "in a conversation - a research brief, an analysis, a decision and "
            "its reasoning, a procedure worked out with the user - survives past "
            "the chat. Use it when the user asks for something to be remembered, "
            "saved, kept or added to the knowledge base, or when they accept an "
            "offer to capture something. Do NOT capture unasked, and do not "
            "capture what is already in the knowledge base (search first with "
            "kb_retrieve; if a near-duplicate note exists, re-capture under the "
            "same title with overwrite=true rather than adding a second copy). "
            "The server names the file from the title - you cannot choose a path, "
            "and no existing document is ever modified or deleted."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "title": {
                    "type": "string",
                    "description": (
                        "Short descriptive title, used as the heading and the "
                        "filename. Make it a searchable noun phrase, e.g. "
                        "'Records retention thresholds', not a sentence."
                    ),
                },
                "content": {
                    "type": "string",
                    "description": (
                        "The note body, in Markdown. Write it to be understood "
                        "months later by someone without the conversation: keep "
                        "the citations, the figures and the reasoning, and drop "
                        "the chat scaffolding."
                    ),
                },
                "source": {
                    "type": "string",
                    "description": (
                        "What kind of note this is - becomes the filename prefix "
                        "and appears in the header. Use one of: Note, Research, "
                        "Report, Analysis, Decision, Procedure. Default: Note."
                    ),
                },
                "tags": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional keywords, written into the note header.",
                },
                "index": {
                    "type": "boolean",
                    "description": (
                        "Index the note immediately so it is searchable now "
                        "(default true). Set false only when capturing several "
                        "notes in a row, then call kb_index once at the end."
                    ),
                },
                "overwrite": {
                    "type": "boolean",
                    "description": (
                        "Replace an existing note with the same title (default "
                        "false, which refuses and reports the existing file)."
                    ),
                },
            },
            "required": ["title", "content"],
        },
    },
    {
        "name": "kb_status",
        "description": (
            "Report the state of the knowledge base: documents in the folder, "
            "captured notes, files/chunks in the vector index, whether the index "
            "is stale, and which endpoints are configured (never shows keys). Use "
            "to diagnose empty or odd retrieval results."
        ),
        "inputSchema": {"type": "object", "properties": {}},
    },
]

TOOL_DISPATCH = {
    "kb_index": tool_kb_index,
    "kb_retrieve": tool_kb_retrieve,
    "kb_ask": tool_kb_ask,
    "kb_capture": tool_kb_capture,
    "kb_status": tool_kb_status,
}


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------

PROTOCOL_VERSION_DEFAULT = "2024-11-05"
SERVER_INFO = {"name": "knowledge-base", "version": __version__}

SERVER_INSTRUCTIONS = (
    "This server is a RAG pipeline over a personal knowledge base of markdown "
    "documents. For a question, call kb_retrieve (or kb_ask for a generated, "
    "cited answer) and ground your reply in the returned chunks, citing source "
    "files. If retrieval reports an empty or stale index, call kb_index first "
    "(it embeds only new/changed files). kb_status shows index freshness and "
    "configuration. kb_capture adds a new note to the knowledge base - use it "
    "only when the user asks for something to be saved or accepts an offer to "
    "save it, never on your own initiative. Existing documents are read-only: "
    "nothing here modifies or deletes them. A retrieved file whose header says "
    "it was written by Claude is a previous note, not an authoritative source - "
    "follow it to the sources it cites."
)


def rpc_result(req_id, result):
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def rpc_error(req_id, code, message):
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


def text_content(text, is_error=False):
    return {"content": [{"type": "text", "text": text}], "isError": is_error}


def handle_request(req):
    """Process one JSON-RPC request. Return a response dict, or None for notifications."""
    method = req.get("method")
    req_id = req.get("id")
    is_notification = "id" not in req

    if method == "initialize":
        params = req.get("params") or {}
        proto = params.get("protocolVersion", PROTOCOL_VERSION_DEFAULT)
        return rpc_result(req_id, {
            "protocolVersion": proto,
            "capabilities": {"tools": {}},
            "serverInfo": SERVER_INFO,
            "instructions": SERVER_INSTRUCTIONS,
        })

    if method == "notifications/initialized":
        return None

    if method == "ping":
        return rpc_result(req_id, {})

    if method == "tools/list":
        return rpc_result(req_id, {"tools": TOOLS})

    if method == "tools/call":
        params = req.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        func = TOOL_DISPATCH.get(name)
        if func is None:
            return rpc_result(req_id, text_content("Unknown tool: {0}".format(name), is_error=True))
        try:
            output = func(arguments)
            return rpc_result(req_id, text_content(output, is_error=False))
        except RagError as exc:
            return rpc_result(req_id, text_content("Knowledge-base error: {0}".format(exc), is_error=True))
        except Exception as exc:
            log("Tool '{0}' failed:\n{1}".format(name, traceback.format_exc()))
            return rpc_result(req_id, text_content("Knowledge-base tool error: {0}".format(exc), is_error=True))

    if is_notification:
        return None
    return rpc_error(req_id, -32601, "Method not found: {0}".format(method))


def run_server():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, respond."""
    log("knowledge-base server started (stdio). Docs: {0}  Index: {1}".format(
        CFG.docs_dir, CFG.index_dir))
    try:
        for raw_line in sys.stdin:
            line = raw_line.strip()
            if not line:
                continue
            try:
                request = json.loads(line)
            except json.JSONDecodeError:
                log("Ignoring malformed JSON line.")
                continue
            response = handle_request(request)
            if response is not None:
                sys.stdout.write(json.dumps(response) + "\n")
                sys.stdout.flush()
    except KeyboardInterrupt:
        pass
    finally:
        log("knowledge-base server stopped.")


# ---------------------------------------------------------------------------
# CLI modes (--check / --reindex / --search / --ask)
# ---------------------------------------------------------------------------

def check_capture():
    """
    Self-test the capture path against a temporary folder: write a note, assert
    the filename and header, assert a traversal-flavoured title is neutralised,
    then clean up. Runs BEFORE the endpoint calls so it still proves the capture
    half of the server on a machine that cannot reach the embeddings API.

    Returns True on success. CFG.output_dir is restored either way.
    """
    real_output = CFG.output_dir
    tmpdir = tempfile.mkdtemp(prefix="kb-capture-check-")
    try:
        CFG.output_dir = tmpdir
        path = capture_note("Capture self test", "Body line.", "Note", ["a", "b"], False)
        if os.path.basename(path) != "Note - Capture self test.md":
            log("Capture self-test     : FAILED - unexpected filename {0}".format(
                os.path.basename(path)))
            return False
        with open(path, "r", encoding="utf-8") as handle:
            written = handle.read()
        for expected in ("# Capture self test", "- Source: Note", "- Captured: ",
                         "- Tags: a, b", "\n---\n", "Body line."):
            if expected not in written:
                log("Capture self-test     : FAILED - missing {0!r}".format(expected))
                return False
        # A title that looks like a path must stay inside the folder.
        escaped = capture_note(r"..\..\escape", "x", "Note", [], False)
        if os.path.dirname(os.path.realpath(escaped)) != os.path.realpath(tmpdir):
            log("Capture self-test     : FAILED - a path-like title escaped the folder")
            return False
        # And a second capture of the same title must refuse without overwrite.
        try:
            capture_note("Capture self test", "Body line.", "Note", [], False)
        except RagError:
            pass
        else:
            log("Capture self-test     : FAILED - duplicate title was not refused")
            return False
        log("Capture self-test     : OK (writes, sanitises titles, refuses duplicates)")
        return True
    except RagError as exc:
        log("Capture self-test     : FAILED - {0}".format(exc))
        return False
    finally:
        CFG.output_dir = real_output
        shutil.rmtree(tmpdir, ignore_errors=True)


def check_capture_folder():
    """Confirm the real capture folder can be created and written to."""
    probe = os.path.join(CFG.output_dir, ".kb-write-probe")
    try:
        os.makedirs(CFG.output_dir, exist_ok=True)
        with open(probe, "w", encoding="utf-8") as handle:
            handle.write("probe")
        os.remove(probe)
        log("Captures folder       : OK (writable)")
        return True
    except OSError as exc:
        log("Captures folder       : FAILED - {0} is not writable ({1})".format(
            CFG.output_dir, exc))
        return False


def run_check():
    """Validate config, test the endpoints, report index status. Exit code 0/1."""
    ok = True
    log(tool_kb_status({}))
    # Capture is checked first: it needs no network, so it still gives a useful
    # result when the endpoints are unreachable.
    if not check_capture():
        ok = False
    if not check_capture_folder():
        ok = False
    try:
        vector = embed_texts(["connectivity test"], is_query=True)[0]
        log("Embeddings endpoint   : OK ({0} dimensions)".format(len(vector)))
    except RagError as exc:
        log("Embeddings endpoint   : FAILED - {0}".format(exc))
        ok = False
    if CFG.chat_url:
        try:
            reply = generate_answer("Reply with the single word: ok",
                                    "(connectivity test - no context)")
            log("Generation endpoint   : OK (replied: {0})".format(reply.strip()[:60]))
        except RagError as exc:
            log("Generation endpoint   : FAILED - {0}".format(exc))
            ok = False
    log("CHECK OK" if ok else "CHECK FAILED")
    return 0 if ok else 1


def run_reindex(force):
    try:
        log(tool_kb_index({"force": force}))
    except RagError as exc:
        log("REINDEX FAILED: {0}".format(exc))
        return 1
    except Exception:
        log("REINDEX FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    return 0


def run_query(mode, text):
    """CLI retrieval/RAG test: mode is 'search' or 'ask'."""
    try:
        if mode == "search":
            log(tool_kb_retrieve({"query": text}))
        else:
            log(tool_kb_ask({"question": text}))
    except RagError as exc:
        log("FAILED: {0}".format(exc))
        return 1
    except Exception:
        log("FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    return 0


def env_flag_false(name):
    """True if env var `name` is an explicit 'off' value (false/0/no)."""
    return os.environ.get(name, "").strip().lower() in {"false", "0", "no"}


def parse_extra_headers(env_name):
    raw = os.environ.get(env_name, "").strip()
    if not raw:
        return {}
    try:
        headers = json.loads(raw)
        if not isinstance(headers, dict):
            raise ValueError("not a JSON object")
        return {str(k): str(v) for k, v in headers.items()}
    except ValueError as exc:
        log("FATAL: {0} is not a JSON object of headers ({1}).".format(env_name, exc))
        sys.exit(2)


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Single-file MCP server providing RAG (index / retrieve / generate) over "
            "a folder of local markdown documents, using a ChromaDB vector index and "
            "your own embeddings (and optionally chat) API endpoints. With no mode "
            "flag it runs as an stdio MCP server. See the docstring for full "
            "configuration; API keys are env-var only (KB_EMBED_API_KEY, "
            "KB_CHAT_API_KEY)."
        )
    )
    env = os.environ.get
    parser.add_argument("--docs-dir", default=env("KB_DOCS_DIR"),
                        help="Folder of markdown documents (env: KB_DOCS_DIR). Required.")
    parser.add_argument("--index-dir", default=env("KB_INDEX_DIR"),
                        help="ChromaDB folder (env: KB_INDEX_DIR; default: <docs-dir>/.kb-rag-index).")
    parser.add_argument("--output-dir", default=env("KB_OUTPUT_DIR"),
                        help="Folder kb_capture writes new notes into (env: KB_OUTPUT_DIR; "
                             "default: <docs-dir>/captures). Must be inside --docs-dir, "
                             "otherwise captured notes would never be indexed.")
    parser.add_argument("--collection", default=env("KB_COLLECTION", "kb-rag"),
                        help="ChromaDB collection name (env: KB_COLLECTION; default: kb-rag).")
    parser.add_argument("--embed-url", default=env("KB_EMBED_URL"),
                        help="Embeddings endpoint URL (env: KB_EMBED_URL). Required.")
    parser.add_argument("--embed-model", default=env("KB_EMBED_MODEL", ""),
                        help="Embedding model name sent in requests (env: KB_EMBED_MODEL).")
    parser.add_argument("--embed-auth-header", default=env("KB_EMBED_AUTH_HEADER", "Authorization"),
                        help="Header for the embed API key; 'Authorization' sends 'Bearer <key>', "
                             "anything else (e.g. 'api-key') sends the raw key (env: KB_EMBED_AUTH_HEADER).")
    parser.add_argument("--embed-style", choices=["openai", "ollama", "kserve-jina", "raw-json"],
                        default=env("KB_EMBED_STYLE", "openai"),
                        help="Embeddings request format (env: KB_EMBED_STYLE; default: openai). "
                             "kserve-jina speaks the KServe V2 Open Inference Protocol "
                             "(input tensors), e.g. a Jina embeddings model on KServe; "
                             "the model name is part of the --embed-url path. raw-json "
                             "POSTs a plain {\"<key>\": [texts]} body for KServe CUSTOM "
                             "predictors that unpack the request dict into their "
                             "pipeline's arguments.")
    parser.add_argument("--embed-tensor-name", default=env("KB_EMBED_TENSOR_NAME", "text"),
                        help="kserve-jina style only: name of the input tensor the texts are "
                             "sent as (env: KB_EMBED_TENSOR_NAME; default: text).")
    parser.add_argument("--embed-json-key", default=env("KB_EMBED_JSON_KEY", "texts"),
                        help="raw-json style only: the JSON key the batch of texts is sent "
                             "under (env: KB_EMBED_JSON_KEY; default: texts).")
    parser.add_argument("--embed-template", default=env("KB_EMBED_TEMPLATE", ""),
                        help="Full request-body control: the complete JSON body to POST, "
                             "with the string \"__TEXTS__\" where the array of texts goes "
                             "(\"__COUNT__\" becomes the number of texts, for tensor "
                             "shape fields; \"__MODEL__\" the model name). Overrides "
                             "--embed-style's request shape; batching still applies. "
                             "Example: {\"inputs\": {\"texts\": \"__TEXTS__\"}} "
                             "(env: KB_EMBED_TEMPLATE).")
    parser.add_argument("--embed-response-path", default=env("KB_EMBED_RESPONSE_PATH", ""),
                        help="Dotted path to the vectors in the embeddings response when "
                             "auto-detection can't find them, e.g. 'outputs.embeddings' or "
                             "'result.0.vectors'; numeric parts index lists "
                             "(env: KB_EMBED_RESPONSE_PATH).")
    parser.add_argument("--debug", action="store_true",
                        default=os.environ.get("KB_DEBUG", "").strip().lower() in {"1", "true", "yes"},
                        help="Log every request and response body (truncated) to stderr, to "
                             "see exactly what is sent to an unknown endpoint (env: KB_DEBUG=1).")
    parser.add_argument("--embed-batch", type=int, default=int(env("KB_EMBED_BATCH", "16")),
                        help="Texts per embeddings request, openai style (env: KB_EMBED_BATCH; default: 16).")
    parser.add_argument("--embed-query-prefix", default=env("KB_EMBED_QUERY_PREFIX", ""),
                        help="Prefix prepended to query text before embedding, for models "
                             "that require it, e.g. 'query: ' (env: KB_EMBED_QUERY_PREFIX).")
    parser.add_argument("--embed-doc-prefix", default=env("KB_EMBED_DOC_PREFIX", ""),
                        help="Prefix prepended to document chunks before embedding, "
                             "e.g. 'passage: ' (env: KB_EMBED_DOC_PREFIX).")
    parser.add_argument("--chat-url", default=env("KB_CHAT_URL", ""),
                        help="Optional chat-completions endpoint for the generate step "
                             "(env: KB_CHAT_URL). Unset: kb_ask returns context for the agent.")
    parser.add_argument("--chat-model", default=env("KB_CHAT_MODEL", ""),
                        help="Generation model name (env: KB_CHAT_MODEL).")
    parser.add_argument("--chat-auth-header", default=env("KB_CHAT_AUTH_HEADER", "Authorization"),
                        help="Header for the chat API key, as per --embed-auth-header "
                             "(env: KB_CHAT_AUTH_HEADER).")
    parser.add_argument("--chat-max-tokens", type=int, default=int(env("KB_CHAT_MAX_TOKENS", "1024")),
                        help="max_tokens for generation; 0 omits the field (env: KB_CHAT_MAX_TOKENS; default: 1024).")
    parser.add_argument("--ca-cert", default=env("KB_CA_CERT", ""),
                        help="Path to a PEM CA bundle for an internal CA (env: KB_CA_CERT).")
    parser.add_argument("--client-cert", default=env("KB_CLIENT_CERT", ""),
                        help="Path to a PEM client certificate, for gateways requiring "
                             "mutual TLS (env: KB_CLIENT_CERT). Presented to both endpoints.")
    parser.add_argument("--client-key", default=env("KB_CLIENT_KEY", ""),
                        help="Path to the PEM private key for --client-cert; omit if the "
                             "cert file contains both (env: KB_CLIENT_KEY). An encrypted "
                             "key's passphrase goes in KB_CLIENT_KEY_PASSWORD (env only).")
    parser.add_argument("--insecure", action="store_true",
                        default=env_flag_false("KB_VERIFY_SSL"),
                        help="Disable TLS certificate verification (env: KB_VERIFY_SSL=false).")
    parser.add_argument("--timeout", type=int, default=int(env("KB_TIMEOUT", "120")),
                        help="HTTP timeout in seconds (env: KB_TIMEOUT; default: 120).")
    parser.add_argument("--chunk-chars", type=int, default=int(env("KB_CHUNK_CHARS", "1500")),
                        help="Soft max characters per chunk (env: KB_CHUNK_CHARS; default: 1500).")
    parser.add_argument("--chunk-overlap", type=int, default=int(env("KB_CHUNK_OVERLAP", "200")),
                        help="Overlap characters between adjacent chunks (env: KB_CHUNK_OVERLAP; default: 200).")
    parser.add_argument("--top-k", type=int, default=int(env("KB_TOP_K", "5")),
                        help="Default number of chunks retrieved (env: KB_TOP_K; default: 5).")
    parser.add_argument("--check", action="store_true",
                        help="Validate config, test the endpoint(s), report index status, then exit.")
    parser.add_argument("--reindex", action="store_true",
                        help="Build/update the vector index, then exit (no server).")
    parser.add_argument("--force", action="store_true",
                        help="With --reindex: drop the index and re-embed everything.")
    parser.add_argument("--search", metavar="QUERY",
                        help="Test retrieval from the CLI: print the top chunks for QUERY, then exit.")
    parser.add_argument("--ask", metavar="QUESTION",
                        help="Test full RAG from the CLI: retrieve + generate for QUESTION, then exit.")
    parser.add_argument("--version", action="version",
                        version="knowledge-base {0}".format(SERVER_INFO["version"]))
    args = parser.parse_args()

    if not args.docs_dir:
        log("FATAL: no knowledge-base folder set. Pass --docs-dir or set KB_DOCS_DIR.")
        sys.exit(2)
    if not os.path.isdir(args.docs_dir):
        log("FATAL: knowledge-base folder does not exist or is not a directory: {0}".format(args.docs_dir))
        sys.exit(2)
    if not args.embed_url:
        log("FATAL: no embeddings endpoint set. Pass --embed-url or set KB_EMBED_URL.")
        sys.exit(2)
    if not re.match(r"^[a-zA-Z0-9][a-zA-Z0-9._-]{1,510}[a-zA-Z0-9]$", args.collection):
        log("FATAL: --collection must be 3-512 characters of [a-zA-Z0-9._-], "
            "starting and ending alphanumeric: {0}".format(args.collection))
        sys.exit(2)
    if args.chunk_overlap >= args.chunk_chars:
        log("FATAL: --chunk-overlap must be smaller than --chunk-chars.")
        sys.exit(2)
    if args.ca_cert and not os.path.isfile(args.ca_cert):
        log("FATAL: --ca-cert file not found: {0}".format(args.ca_cert))
        sys.exit(2)
    if args.client_key and not args.client_cert:
        log("FATAL: --client-key was given without --client-cert.")
        sys.exit(2)
    if args.client_cert and not os.path.isfile(args.client_cert):
        log("FATAL: --client-cert file not found: {0}".format(args.client_cert))
        sys.exit(2)
    if args.client_key and not os.path.isfile(args.client_key):
        log("FATAL: --client-key file not found: {0}".format(args.client_key))
        sys.exit(2)

    CFG.docs_dir = os.path.realpath(args.docs_dir)
    CFG.index_dir = os.path.realpath(
        args.index_dir or os.path.join(CFG.docs_dir, ".kb-rag-index"))
    CFG.output_dir = os.path.realpath(
        args.output_dir or os.path.join(CFG.docs_dir, "captures"))
    # Captures have to live inside the documents folder or scan_documents()
    # would never see them, so a note would be written and never indexed.
    # Refuse at startup rather than at the first kb_capture call.
    if not is_within(CFG.output_dir, CFG.docs_dir):
        log("FATAL: --output-dir must be inside --docs-dir, otherwise captured "
            "notes would never be indexed.\n  --output-dir: {0}\n  --docs-dir  : "
            "{1}".format(CFG.output_dir, CFG.docs_dir))
        sys.exit(2)
    if CFG.output_dir == CFG.index_dir or is_within(CFG.output_dir, CFG.index_dir):
        log("FATAL: --output-dir must not be inside the vector index folder "
            "({0}), which is pruned from indexing.".format(CFG.index_dir))
        sys.exit(2)
    if any(part.startswith(".") for part in
           os.path.relpath(CFG.output_dir, CFG.docs_dir).split(os.sep)
           if part not in (".", "")):
        log("FATAL: --output-dir must not be a dot-folder ({0}) - dot-folders "
            "are pruned from indexing, so captures would never be "
            "searchable.".format(CFG.output_dir))
        sys.exit(2)
    CFG.collection = args.collection
    CFG.embed_url = args.embed_url
    CFG.embed_model = args.embed_model
    CFG.embed_key = os.environ.get("KB_EMBED_API_KEY", "")
    CFG.embed_auth_header = args.embed_auth_header
    CFG.embed_style = args.embed_style
    CFG.embed_tensor_name = args.embed_tensor_name
    CFG.embed_json_key = args.embed_json_key
    if args.embed_template:
        try:
            CFG.embed_template = json.loads(args.embed_template)
        except ValueError as exc:
            log("FATAL: --embed-template is not valid JSON ({0}): {1}".format(
                exc, args.embed_template))
            sys.exit(2)
        if count_template_placeholders(CFG.embed_template) == 0:
            log("FATAL: --embed-template must contain the JSON string \"{0}\" "
                "where the array of texts goes.".format(TEMPLATE_TEXTS))
            sys.exit(2)
    CFG.embed_response_path = args.embed_response_path
    CFG.debug = args.debug
    CFG.embed_batch = max(1, args.embed_batch)
    CFG.embed_query_prefix = args.embed_query_prefix
    CFG.embed_doc_prefix = args.embed_doc_prefix
    CFG.embed_extra_headers = parse_extra_headers("KB_EMBED_EXTRA_HEADERS")
    CFG.chat_url = args.chat_url
    CFG.chat_model = args.chat_model
    CFG.chat_key = os.environ.get("KB_CHAT_API_KEY", "") or CFG.embed_key
    CFG.chat_auth_header = args.chat_auth_header
    CFG.chat_max_tokens = args.chat_max_tokens
    CFG.chat_extra_headers = parse_extra_headers("KB_CHAT_EXTRA_HEADERS")
    CFG.ca_cert = args.ca_cert
    CFG.client_cert = args.client_cert
    CFG.client_key = args.client_key
    CFG.client_key_password = os.environ.get("KB_CLIENT_KEY_PASSWORD", "")
    CFG.verify_ssl = not args.insecure
    CFG.timeout = max(1, args.timeout)
    CFG.chunk_chars = max(200, args.chunk_chars)
    CFG.chunk_overlap = max(0, args.chunk_overlap)
    CFG.top_k = max(1, min(20, args.top_k))

    if not CFG.verify_ssl:
        log("WARNING: TLS certificate verification is DISABLED.")

    # Fail fast on an unloadable client cert/key (bad file, wrong passphrase)
    # rather than surfacing it on the first tool call.
    if CFG.client_cert:
        try:
            build_ssl_context()
        except RagError as exc:
            log("FATAL: {0}".format(exc))
            sys.exit(2)

    if args.check:
        sys.exit(run_check())
    if args.reindex:
        sys.exit(run_reindex(args.force))
    if args.search:
        sys.exit(run_query("search", args.search))
    if args.ask:
        sys.exit(run_query("ask", args.ask))
    run_server()


if __name__ == "__main__":
    main()
