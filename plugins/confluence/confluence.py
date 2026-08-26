#!/usr/bin/env python3
"""
confluence.py (v2.0.0) - A single-file MCP (Model Context Protocol) server
for querying ONE OR TWO Confluence Data Center instances (tested against the
9.x v1 REST API) using only the Python 3 standard library.

It speaks MCP over stdio (newline-delimited JSON-RPC 2.0), the transport an
MCP client launches for a `type: stdio` server. No third-party packages are
required.

Tools exposed (read-only / query):
  - confluence_search           : free-text search for pages
  - confluence_search_cql       : advanced search using raw CQL
  - confluence_get_page         : fetch one page by numeric ID (with body text)
  - confluence_get_page_by_title: fetch one page by exact title + space key
  - confluence_list_pages_under : list pages beneath a parent page

TWO CONFLUENCE SERVERS
----------------------
A second Confluence instance is optional. Configure it and each server gets a
friendly name (e.g. Green and Blue); every tool then takes an extra optional
'server' argument:

  - 'server' omitted        -> the FIRST server is used. This is the default,
                               so ordinary prompts need not mention a server.
  - 'server': "Blue"        -> the SECOND server is used. Claude passes this
                               when the user names it, e.g. "find content
                               about X on Blue". Matching is case-insensitive.

With two servers configured, output is labelled with the server it came from,
and knowledge-base files are named 'Confluence <name> - <title>.md' so the two
instances cannot overwrite each other. Configure only one server and the
behaviour is exactly as it was before: no 'server' argument, no labels, and
files stay named 'Confluence - <title>.md'.

Content IDs are NOT shared between instances: page 393217 on Green is a
different page from 393217 on Blue.

CONFIGURATION
-------------
Read from environment variables (the natural fit for an MCP client's `env`
block); non-secret settings can be overridden by command-line arguments.
Precedence is CLI flag > environment variable > default.

CREDENTIALS ARE ENV-VAR ONLY - there are no --token/--user/--password flags,
because command-line arguments are visible to other local users in process
listings.

  First server (required):
  CONFLUENCE_NAME       friendly name used to select it in a prompt
                        (--name, default "Primary")
  CONFLUENCE_BASE_URL   e.g. https://confluence.internal.example.com
                        (--base-url; include any context path, no trailing slash)
  CONFLUENCE_TOKEN      Personal Access Token (preferred; sent as Bearer)
  CONFLUENCE_USER       username   } basic-auth fallback if no token is given
  CONFLUENCE_PASSWORD   password   }
  CONFLUENCE_VERIFY_SSL "false" to disable TLS verification (--insecure;
                        default: verify)
  CONFLUENCE_CA_CERT    path to a PEM CA bundle for an internal CA (--ca-cert)

  Second server (optional - set CONFLUENCE_BASE_URL_2 to enable it):
  CONFLUENCE_NAME_2       friendly name (--name-2, default "Secondary")
  CONFLUENCE_BASE_URL_2   base URL of the second instance (--base-url-2)
  CONFLUENCE_TOKEN_2      its own Personal Access Token
  CONFLUENCE_USER_2       username   } basic-auth fallback for the second server
  CONFLUENCE_PASSWORD_2   password   }
  CONFLUENCE_VERIFY_SSL_2 TLS verification for the second server (--insecure-2);
                          falls back to the first server's setting
  CONFLUENCE_CA_CERT_2    CA bundle for the second server (--ca-cert-2);
                          falls back to CONFLUENCE_CA_CERT

  Shared by both servers:
  CONFLUENCE_TIMEOUT    request timeout in seconds (--timeout, default: 30)
  CONFLUENCE_MAX_BODY   truncate page bodies to N chars (--max-body,
                        0 = unlimited, default 0). This limit applies only to
                        the text returned to the model; files saved to
                        CONFLUENCE_KB_DIR are never truncated.
  CONFLUENCE_KB_DIR     every page that is read is also saved as a Markdown
                        file into this folder, for feeding a local RAG
                        knowledge base (--kb-dir). Files are overwritten if
                        they already exist. DEFAULTS to
                        C:\Eva\knowledge\confluence - the Confluence folder of
                        the Eva working tree, which sits inside the
                        knowledge-base plugin's documents folder so mirrored
                        pages are actually indexed. Pass 'off' to disable
                        mirroring, after which the server writes no local file
                        at all.

The second server is all-or-nothing: if CONFLUENCE_BASE_URL_2 is set without
credentials, the server refuses to start rather than quietly answering "Blue"
questions from the first instance.

INSTALLING INTO CLAUDE CODE
---------------------------
This server ships as the "confluence" Claude Code plugin (its manifest is
.claude-plugin/plugin.json next to this file), so the normal install is:

    /plugin marketplace add C:\path\to\claude-skills
    /plugin install confluence@mcnamee-claude-skills

Claude Code prompts for each server's name and base URL, the optional
knowledge-base folder and the Python interpreter. Tokens are NOT stored in the
plugin - set them as Windows user environment variables before starting Claude
Code, and the plugin picks them up from there:

    setx CONFLUENCE_TOKEN   "token-for-the-first-server"
    setx CONFLUENCE_TOKEN_2 "token-for-the-second-server"

(`setx` does not affect processes that are already running - quit VS Code
completely and reopen it.) See README.md next to this file for the full
settings reference.

TESTING
-------
Diagnostic output goes ONLY to stderr. stdout is reserved for the JSON-RPC
stream - writing anything else there would corrupt the protocol.

`--check` connects to EVERY configured server, prints who you are
authenticated as and how many spaces are visible (to stderr), then exits
without starting the server. It exits non-zero if any server fails, so it is
the fastest way to prove a two-server setup before wiring it in:

    & "C:\path\to\python.exe" confluence.py --check

A quick two-server smoke test from PowerShell, without touching the plugin
config (note the tokens go in the environment, never in the arguments):

    $env:CONFLUENCE_TOKEN   = "green-token"
    $env:CONFLUENCE_TOKEN_2 = "blue-token"
    & "C:\path\to\python.exe" confluence.py `
        --name Green --base-url https://green.confluence.example.com `
        --name-2 Blue --base-url-2 https://blue.confluence.example.com --check
"""

# Semantic version of this server. Bump on EVERY change (see CLAUDE.md):
# MAJOR = breaking config/tool change, MINOR = new feature, PATCH = fix.
__version__ = "2.0.0"

import argparse
import base64
import datetime
import html.parser
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request

SERVER_NAME = "confluence-mcp"
SERVER_VERSION = __version__

# Folder every page read is mirrored to as Markdown, for a local RAG index.
# Set it here, or at launch with --kb-dir / the CONFLUENCE_KB_DIR environment
# variable (which take priority over this constant).
# Default: the confluence\ sub-folder of the Eva knowledge base. It MUST stay
# inside the knowledge-base plugin's documents folder (C:\Eva\knowledge) or the
# mirrored pages would never be indexed. The folder is created on demand.
# Set to None here (or pass --kb-dir off) to disable mirroring, after which this
# server touches no local file at all.
KB_DIR = r"C:\Eva\knowledge\confluence"

# Folder-setting values that mean "explicitly turned off". An MCP client can
# only pass strings, and a BLANK string is what it substitutes for a setting the
# user left empty - which means "not configured", falling back to the default
# above. So a keyword is needed to say "definitely off".
DISABLE_KEYWORDS = frozenset(("off", "none", "no", "false", "disabled"))
# Friendly names used when the user does not supply one. They only ever show up
# in output (or in a tool's 'server' enum) when a second server is configured.
DEFAULT_NAME_1 = "Primary"
DEFAULT_NAME_2 = "Secondary"
# Protocol version we default to if the client does not send one. We echo the
# client's requested version when possible (see handle_initialize) so that we
# stay compatible with whatever the host negotiated.
DEFAULT_PROTOCOL_VERSION = "2024-11-05"

# JSON-RPC error codes (subset we use)
PARSE_ERROR = -32700
INVALID_REQUEST = -32600
METHOD_NOT_FOUND = -32601
INTERNAL_ERROR = -32603


def log(*args):
    """Write a diagnostic line to stderr (never stdout)."""
    print("[confluence-mcp]", *args, file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# HTML -> plain text
# ---------------------------------------------------------------------------
class _TextExtractor(html.parser.HTMLParser):
    """
    Minimal HTML/XHTML to plain-text converter.

    Confluence "storage format" bodies are XHTML with extra <ac:...> macro
    tags. We don't try to interpret macros; we just keep the readable text and
    insert line breaks around block-level elements so the result is legible.
    convert_charrefs=True (the default) means entities like &amp; are decoded
    for us and arrive via handle_data.
    """

    _BLOCK_TAGS = {
        "p", "br", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
        "table", "ul", "ol", "blockquote", "pre", "section", "header",
        "footer", "article",
    }
    _SKIP_TAGS = {"script", "style"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._parts = []
        self._skip_depth = 0

    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
        if tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag in self._SKIP_TAGS and self._skip_depth > 0:
            self._skip_depth -= 1
        if tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_data(self, data):
        if self._skip_depth == 0:
            self._parts.append(data)

    def get_text(self):
        text = "".join(self._parts)
        # Collapse runs of blank lines and trim trailing spaces per line.
        lines = [ln.rstrip() for ln in text.splitlines()]
        out = []
        blank = False
        for ln in lines:
            if ln.strip() == "":
                if not blank:
                    out.append("")
                blank = True
            else:
                out.append(ln)
                blank = False
        return "\n".join(out).strip()


def html_to_text(raw):
    """Convert an HTML/XHTML string to plain text, defensively."""
    if not raw:
        return ""
    parser = _TextExtractor()
    try:
        parser.feed(raw)
        parser.close()
        return parser.get_text()
    except Exception:
        # If parsing somehow fails, fall back to returning the raw string
        # rather than losing the content entirely.
        return raw


class _MarkdownExtractor(html.parser.HTMLParser):
    """
    Convert Confluence storage-format XHTML into reasonable Markdown.

    This is a best-effort converter aimed at RAG ingestion, not a pixel-perfect
    renderer. It handles the common structural elements (headings, paragraphs,
    lists, bold/italic, links, inline code, code blocks, block quotes, rules and
    tables). Confluence macros (<ac:...>) are not interpreted, but their inner
    text - including code-macro CDATA bodies - is preserved. Text is not
    Markdown-escaped, so the occasional literal '*' may look like emphasis; that
    is a deliberate trade-off to keep the captured text faithful for search.
    """

    _SKIP_TAGS = {"script", "style"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self._skip_depth = 0
        self._in_pre = 0
        self._list_stack = []      # 'ul' / 'ol' per nesting level
        self._ol_counters = []     # running item number per ordered-list level
        self._href_stack = []      # href per currently-open <a>
        # Table buffering (only the outermost table is rendered as a grid)
        self._table_depth = 0
        self._rows = None          # list of cell-lists for the current table
        self._row = None           # current row (list of cell strings)
        self._cell = None          # buffer for the current cell, or None

    def _emit(self, s):
        # Route text either into the current table cell or the main output.
        if self._cell is not None:
            self._cell.append(s)
        else:
            self.parts.append(s)

    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if tag == "br":
            self._emit("\n" if self._in_pre else "  \n")
        elif tag == "p":
            self._emit("\n\n")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._emit("\n\n" + "#" * int(tag[1]) + " ")
        elif tag in ("strong", "b") and not self._in_pre:
            self._emit("**")
        elif tag in ("em", "i") and not self._in_pre:
            self._emit("*")
        elif tag == "code" and not self._in_pre:
            self._emit("`")
        elif tag == "pre":
            self._in_pre += 1
            self._emit("\n\n```\n")
        elif tag == "blockquote":
            self._emit("\n\n> ")
        elif tag == "hr":
            self._emit("\n\n---\n\n")
        elif tag == "a":
            href = ""
            for key, val in attrs:
                if key == "href":
                    href = val or ""
            self._href_stack.append(href)
            self._emit("[")
        elif tag == "ul":
            self._list_stack.append("ul")
        elif tag == "ol":
            self._list_stack.append("ol")
            self._ol_counters.append(0)
        elif tag == "li":
            indent = "  " * max(0, len(self._list_stack) - 1)
            if self._list_stack and self._list_stack[-1] == "ol":
                self._ol_counters[-1] += 1
                marker = "{}. ".format(self._ol_counters[-1])
            else:
                marker = "- "
            self._emit("\n" + indent + marker)
        elif tag == "table":
            self._table_depth += 1
            if self._table_depth == 1:
                self._rows = []
        elif tag == "tr":
            if self._rows is not None:
                self._row = []
        elif tag in ("td", "th"):
            if self._row is not None:
                self._cell = []

    def handle_endtag(self, tag):
        if tag in self._SKIP_TAGS:
            if self._skip_depth:
                self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        if tag == "p":
            self._emit("\n\n")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._emit("\n\n")
        elif tag in ("strong", "b") and not self._in_pre:
            self._emit("**")
        elif tag in ("em", "i") and not self._in_pre:
            self._emit("*")
        elif tag == "code" and not self._in_pre:
            self._emit("`")
        elif tag == "pre":
            if self._in_pre:
                self._in_pre -= 1
            self._emit("\n```\n\n")
        elif tag == "blockquote":
            self._emit("\n\n")
        elif tag == "a":
            href = self._href_stack.pop() if self._href_stack else ""
            self._emit("]({})".format(href))
        elif tag in ("ul", "ol"):
            if self._list_stack:
                if self._list_stack.pop() == "ol" and self._ol_counters:
                    self._ol_counters.pop()
            self._emit("\n")
        elif tag in ("td", "th"):
            if self._cell is not None and self._row is not None:
                # Markdown cells are single-line: flatten and escape pipes.
                cell_text = " ".join("".join(self._cell).split())
                self._row.append(cell_text.replace("|", "\\|"))
                self._cell = None
        elif tag == "tr":
            if self._row is not None and self._rows is not None:
                self._rows.append(self._row)
                self._row = None
        elif tag == "table":
            if self._table_depth == 1 and self._rows is not None:
                self._emit_table(self._rows)
                self._rows = None
            if self._table_depth:
                self._table_depth -= 1

    def handle_data(self, data):
        if self._skip_depth:
            return
        if self._in_pre:
            self._emit(data)
            return
        if data.strip() == "":
            # Whitespace-only node between tags: keep a single separating space
            # rather than injecting blank lines.
            if data:
                self._emit(" ")
            return
        # Collapse embedded newlines so wrapped source doesn't break paragraphs.
        self._emit(data.replace("\r", " ").replace("\n", " "))

    def unknown_decl(self, data):
        # Capture CDATA content, e.g. Confluence code-macro bodies.
        if self._skip_depth:
            return
        if data.startswith("CDATA["):
            inner = data[6:]
            if inner.endswith("]"):
                inner = inner[:-1]
            self._emit(inner)

    def _emit_table(self, rows):
        if not rows:
            return
        ncols = max((len(r) for r in rows), default=0)
        if ncols == 0:
            return

        def fmt(cells):
            padded = list(cells) + [""] * (ncols - len(cells))
            return "| " + " | ".join(padded) + " |"

        out = [fmt(rows[0]), "| " + " | ".join(["---"] * ncols) + " |"]
        out.extend(fmt(r) for r in rows[1:])
        self.parts.append("\n\n" + "\n".join(out) + "\n\n")

    def get_markdown(self):
        text = "".join(self.parts)
        # Trim trailing spaces and collapse runs of blank lines to a single one.
        lines = [ln.rstrip() for ln in text.split("\n")]
        out = []
        blank = 0
        for ln in lines:
            if ln.strip() == "":
                blank += 1
                if blank <= 1:
                    out.append("")
            else:
                blank = 0
                out.append(ln)
        return "\n".join(out).strip() + "\n"


def html_to_markdown(raw):
    """Convert an HTML/XHTML string to Markdown, falling back to plain text."""
    if not raw:
        return ""
    parser = _MarkdownExtractor()
    try:
        parser.feed(raw)
        parser.close()
        return parser.get_markdown()
    except Exception:
        # Never lose content: fall back to the plain-text extractor.
        return html_to_text(raw)


def safe_filename(name, max_len=150):
    """
    Turn a page title into a filesystem-safe filename component (no extension).

    Strips characters that are illegal on Windows (< > : " / \\ | ? * and control
    chars), collapses whitespace, removes trailing dots/spaces (also illegal on
    Windows), and caps the length. Returns 'untitled' if nothing usable is left.
    """
    if not name:
        return "untitled"
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    cleaned = " ".join(cleaned.split()).strip(" .")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" .")
    return cleaned or "untitled"


def cql_quote(value):
    """
    Escape a string for safe inclusion inside a double-quoted CQL literal.
    Backslashes and double quotes must be escaped. This prevents a value
    containing a quote from breaking out of the literal.
    """
    return value.replace("\\", "\\\\").replace('"', '\\"')


# ---------------------------------------------------------------------------
# Confluence client
# ---------------------------------------------------------------------------
class ConfluenceError(Exception):
    """Raised for any failure talking to Confluence; message is user-facing."""


class ConfluenceClient:
    def __init__(self, name, base_url, token=None, user=None, password=None,
                 verify_ssl=True, ca_cert=None, timeout=30, max_body=0,
                 kb_dir=None):
        if not name:
            raise ValueError("name is required")
        if not base_url:
            raise ValueError("base_url is required")
        self.name = name
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_body = max_body
        # Folder to mirror pages into as Markdown; None/empty disables saving.
        self.kb_dir = kb_dir or None
        # Whether output should say which server it came from. ConfluenceServers
        # turns this on when more than one server is configured; with a single
        # server the output stays exactly as it was before multi-server support.
        self.label_output = False

        # Build auth header. Prefer a Personal Access Token (Bearer) if given.
        self.headers = {"Accept": "application/json"}
        if token:
            self.headers["Authorization"] = "Bearer " + token
        elif user is not None and password is not None:
            raw = "{}:{}".format(user, password).encode("utf-8")
            self.headers["Authorization"] = "Basic " + base64.b64encode(raw).decode("ascii")
        else:
            # main() validates credentials first and prints the exact env-var
            # names for the server in question; this is the backstop.
            raise ValueError(
                "No credentials for Confluence server {!r}.".format(name)
            )

        # Build the TLS context. Only relevant for https:// URLs; ignored for
        # plain http. A custom CA bundle takes precedence; otherwise we either
        # verify normally or, if explicitly asked, disable verification.
        self.verify_ssl = bool(verify_ssl) or bool(ca_cert)
        if ca_cert:
            self.ssl_context = ssl.create_default_context(cafile=ca_cert)
        elif not verify_ssl:
            self.ssl_context = ssl.create_default_context()
            self.ssl_context.check_hostname = False
            self.ssl_context.verify_mode = ssl.CERT_NONE
        else:
            self.ssl_context = ssl.create_default_context()

    def _on_server(self):
        """' on <name>' when more than one server is configured, else ''."""
        return " on {}".format(self.name) if self.label_output else ""

    def _fail(self, message):
        """
        Raise a ConfluenceError, prefixed with this server's name when more than
        one server is configured (so the user can see WHICH instance failed).
        """
        prefix = "[{}] ".format(self.name) if self.label_output else ""
        raise ConfluenceError(prefix + message)

    def _get(self, path, params=None):
        """Perform a GET against the REST API and return parsed JSON."""
        url = self.base_url + path
        if params:
            # urlencode percent-encodes values (including CQL special chars).
            url = url + "?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers=self.headers, method="GET")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout,
                                        context=self.ssl_context) as resp:
                body = resp.read()
        except urllib.error.HTTPError as e:
            # Try to surface Confluence's error message from the response body.
            detail = ""
            try:
                detail = e.read().decode("utf-8", "replace")[:500]
            except Exception:
                pass
            self._fail(
                "HTTP {} from Confluence for {}{}".format(
                    e.code, url, (": " + detail) if detail else ""
                )
            )
        except urllib.error.URLError as e:
            self._fail(
                "Could not reach Confluence at {} ({}). Check the base URL, "
                "network reachability and TLS settings.".format(url, e.reason)
            )
        except ssl.SSLError as e:
            self._fail(
                "TLS error talking to Confluence ({}). For an internal CA, set "
                "CONFLUENCE_CA_CERT, or CONFLUENCE_VERIFY_SSL=false to disable "
                "verification.".format(e)
            )
        try:
            return json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as e:
            self._fail("Confluence returned a non-JSON response: {}".format(e))

    def _abs_link(self, data, link):
        """Build an absolute web URL from a result's webui link."""
        if not link:
            return ""
        base = ""
        links = data.get("_links") if isinstance(data, dict) else None
        if isinstance(links, dict):
            base = links.get("base") or ""
        if not base:
            base = self.base_url
        return base.rstrip("/") + link

    def search(self, cql, limit):
        """Run a CQL query against the content search endpoint."""
        params = {
            "cql": cql,
            "limit": limit,
            "expand": "space,version",
        }
        data = self._get("/rest/api/content/search", params)
        results = data.get("results", []) or []
        # With two servers configured, every result carries the server it came
        # from: content IDs are per-instance, so the follow-up read has to be
        # aimed at the same one.
        server_field = "  server={}".format(self.name) if self.label_output else ""
        lines = []
        for item in results:
            space = (item.get("space") or {}).get("key", "?")
            title = item.get("title", "(untitled)")
            cid = item.get("id", "?")
            ctype = item.get("type", "content")
            link = self._abs_link(
                data, (item.get("_links") or {}).get("webui", "")
            )
            lines.append(
                "- id={id}{server}  type={type}  space={space}\n  title: {title}\n  url: {url}".format(
                    id=cid, server=server_field, type=ctype, space=space,
                    title=title, url=link
                )
            )
        header = "Found {} result(s){} for CQL: {}".format(
            len(lines), self._on_server(), cql)
        if not lines:
            return header + "\n(no matching content)"
        footer = ""
        if self.label_output:
            footer = (
                '\n\n(These content IDs exist on {0} only - pass server="{0}" '
                "when reading or listing them.)".format(self.name)
            )
        return header + "\n\n" + "\n\n".join(lines) + footer

    def _render_page(self, page):
        """Format a single content object (with body.storage) as text."""
        title = page.get("title", "(untitled)")
        cid = page.get("id", "?")
        ctype = page.get("type", "content")
        space = (page.get("space") or {}).get("key", "?")
        version = (page.get("version") or {}).get("number", "?")
        link = self._abs_link(page, (page.get("_links") or {}).get("webui", ""))
        storage = (((page.get("body") or {}).get("storage") or {}).get("value")) or ""
        text = html_to_text(storage)
        truncated_note = ""
        if self.max_body and len(text) > self.max_body:
            text = text[: self.max_body]
            truncated_note = "\n\n[...body truncated to {} characters...]".format(self.max_body)
        fields = [
            ("Title", title),
            ("ID", cid),
            ("Type", ctype),
            ("Space", space),
        ]
        # Only name the server when there is more than one to tell apart.
        if self.label_output:
            fields.append(("Server", self.name))
        fields.extend([("Version", version), ("URL", link)])
        meta = "".join("{}: {}\n".format(k, v) for k, v in fields) + "\n--- Content ---\n"
        rendered = meta + (text if text else "(this page has no readable body content)") + truncated_note

        # If a knowledge-base folder is configured, mirror the page to Markdown.
        # This is a side effect of reading a page; it must never break the tool,
        # so any failure is reported but swallowed.
        if self.kb_dir:
            try:
                path = self._save_to_kb(title, link, space, version, storage)
                log("saved page to knowledge base: {}".format(path))
                rendered += "\n\n[Saved to knowledge base: {}]".format(path)
            except OSError as e:
                log("knowledge-base save failed: {}".format(e))
                rendered += "\n\n[Knowledge-base save FAILED: {}]".format(e)
        return rendered

    def _save_to_kb(self, title, link, space, version, storage):
        """
        Write the page to '<kb_dir>/Confluence - <title>.md', overwriting any
        existing file. Returns the path written; raises OSError on failure.

        With two servers configured the filename becomes
        'Confluence <server> - <title>.md', so a page with the same title on
        both instances produces two files instead of one overwriting the other.

        The FULL body is always saved (the CONFLUENCE_MAX_BODY limit only trims
        what is returned to the model, not what is stored for RAG).
        """
        md_body = html_to_markdown(storage)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        server_line = "- Server: {}\n".format(self.name) if self.label_output else ""
        header = (
            "# {title}\n\n"
            "- Source: {url}\n"
            "{server_line}"
            "- Space: {space}\n"
            "- Version: {version}\n"
            "- Fetched: {stamp}\n\n"
            "---\n\n"
        ).format(title=title, url=link or "(unknown)", server_line=server_line,
                 space=space, version=version, stamp=stamp)
        content = header + (md_body if md_body else "(no readable body content)\n")

        # Create the folder if needed, then write. newline="\n" keeps endings
        # consistent and avoids CRLF doubling on Windows.
        os.makedirs(self.kb_dir, exist_ok=True)
        prefix = "Confluence - "
        if self.label_output:
            prefix = "Confluence {} - ".format(safe_filename(self.name, 40))
        filename = prefix + safe_filename(title) + ".md"
        path = os.path.join(self.kb_dir, filename)
        with open(path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(content)
        return path

    def get_page(self, page_id):
        if page_id is None or str(page_id).strip() == "":
            raise ConfluenceError("'page_id' is required")
        page_id = str(page_id).strip()
        params = {"expand": "body.storage,version,space"}
        page = self._get("/rest/api/content/" + urllib.parse.quote(page_id, safe=""), params)
        return self._render_page(page)

    def get_page_by_title(self, title, space):
        if not title or not space:
            raise ConfluenceError("Both 'title' and 'space' are required")
        params = {
            "title": title,
            "spaceKey": space,
            "expand": "body.storage,version,space",
            "limit": 1,
        }
        data = self._get("/rest/api/content", params)
        results = data.get("results", []) or []
        if not results:
            return "No page titled {!r} found in space {!r}{}.".format(
                title, space, self._on_server())
        return self._render_page(results[0])

    def resolve_page_id(self, title, space):
        """
        Look up a single page's numeric ID from its exact title + space key.
        Returns the ID string, or raises ConfluenceError if not found.
        """
        if not title or not space:
            raise ConfluenceError("Both 'parent_title' and 'space' are required "
                                  "when 'parent_id' is not given")
        params = {"title": title, "spaceKey": space, "limit": 1}
        data = self._get("/rest/api/content", params)
        results = data.get("results", []) or []
        if not results:
            raise ConfluenceError(
                "No page titled {!r} found in space {!r}{}.".format(
                    title, space, self._on_server())
            )
        return str(results[0].get("id"))

    def list_pages_under(self, parent_id, direct_only=False,
                         modified_within_days=None, limit=25):
        """
        List pages beneath a parent page. Builds the CQL internally so the
        caller never has to know CQL.
          - direct_only=True  -> only immediate children (CQL 'parent')
          - direct_only=False -> all descendants at any depth (CQL 'ancestor')
          - modified_within_days -> optionally restrict to pages changed in the
            last N days (CQL 'lastmodified >= now("-Nd")').
        """
        parent_id = str(parent_id).strip()
        if not parent_id:
            raise ConfluenceError("A parent page could not be identified")
        if not parent_id.isdigit():
            # Enforced so the unquoted embed below cannot inject CQL.
            raise ConfluenceError(
                "'parent_id' must be a numeric content ID (got {!r}). Use "
                "'parent_title' plus 'space' if you only know the title.".format(parent_id)
            )
        field = "parent" if direct_only else "ancestor"
        # parent_id is validated as numeric above, so it is safe to embed unquoted.
        clauses = ["{} = {}".format(field, parent_id), "type = page"]
        if modified_within_days is not None:
            try:
                days = int(modified_within_days)
            except (TypeError, ValueError):
                raise ConfluenceError("'modified_within_days' must be a whole number")
            if days > 0:
                clauses.append('lastmodified >= now("-{}d")'.format(days))
        cql = " AND ".join(clauses) + " ORDER BY lastmodified DESC"
        return self.search(cql, limit)


class ConfluenceServers:
    """
    The one or two configured Confluence instances, in configuration order.

    The FIRST entry is the default: a tool call that does not name a server
    goes there, which is what keeps ordinary prompts ("search Confluence for
    X") working without the user thinking about instances.
    """

    def __init__(self, clients):
        if not clients:
            raise ValueError("at least one Confluence server is required")
        self.clients = list(clients)
        # Output is only labelled with a server name when there is more than
        # one server to tell apart - a single-server setup looks exactly as it
        # did before multi-server support was added.
        multi = len(self.clients) > 1
        for client in self.clients:
            client.label_output = multi

    @property
    def multi(self):
        return len(self.clients) > 1

    @property
    def default(self):
        return self.clients[0]

    def names(self):
        return [client.name for client in self.clients]

    def resolve(self, selector):
        """
        Return the client for a tool call's 'server' argument.

        Accepts the server's name (case-insensitive), an unambiguous part of it
        ("blue" for "Blue Wiki"), or its 1-based position ("1"/"2"). An empty or
        missing selector means the default server. Anything else is an error
        naming the servers that ARE configured, so the agent can retry - never
        a silent fall back to the default, which would answer a question about
        one instance with content from the other.
        """
        if selector is None:
            return self.default
        want = str(selector).strip()
        if not want:
            return self.default

        for client in self.clients:
            if client.name.lower() == want.lower():
                return client

        # 1-based position, so 'server: "2"' works even without the name.
        if want.isdigit():
            index = int(want)
            if 1 <= index <= len(self.clients):
                return self.clients[index - 1]

        # Last resort: a unique partial match, e.g. "blue" -> "Blue Wiki".
        partial = [c for c in self.clients if want.lower() in c.name.lower()]
        if len(partial) == 1:
            return partial[0]

        raise ConfluenceError(
            "Unknown Confluence server {!r}. Configured server(s): {}. Omit "
            "'server' to use {}.".format(
                want, ", ".join(repr(n) for n in self.names()), self.default.name
            )
        )


# ---------------------------------------------------------------------------
# Tool definitions and dispatch
# ---------------------------------------------------------------------------
def base_tool_definitions():
    """
    The tools as advertised when a single server is configured (JSON-Schema
    input specs). tool_definitions() adds the 'server' argument on top when a
    second server is configured.
    """
    return [
        {
            "name": "confluence_search",
            "description": (
                "Search Confluence pages by free text. Returns matching pages "
                "with their numeric ID, title, space key and URL. Use "
                "'confluence_get_page' afterwards to read a page's full content. "
                "Optionally restrict to a single space by its space key."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Free-text search terms.",
                    },
                    "space": {
                        "type": "string",
                        "description": "Optional space key to restrict the search (e.g. 'DOCS').",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (1-50, default 25).",
                    },
                },
                "required": ["query"],
            },
        },
        {
            "name": "confluence_search_cql",
            "description": (
                "Search Confluence using a raw CQL (Confluence Query Language) "
                "query for advanced filtering. Examples: "
                "'space = \"DOCS\" AND type = page', "
                "'text ~ \"release notes\" AND lastModified >= now(\"-30d\")', "
                "'title ~ \"runbook\"'. Returns matching content with IDs, "
                "titles, space keys and URLs."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "cql": {
                        "type": "string",
                        "description": "A valid CQL query string.",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (1-50, default 25).",
                    },
                },
                "required": ["cql"],
            },
        },
        {
            "name": "confluence_get_page",
            "description": (
                "Retrieve a single Confluence page by its numeric page ID. "
                "Returns the title, space, version, URL and the page body "
                "converted to plain text."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "page_id": {
                        "type": "string",
                        "description": "The numeric Confluence content ID, e.g. '393217'.",
                    },
                },
                "required": ["page_id"],
            },
        },
        {
            "name": "confluence_get_page_by_title",
            "description": (
                "Retrieve a Confluence page by its exact title within a given "
                "space. Returns the same details as confluence_get_page."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "The exact page title.",
                    },
                    "space": {
                        "type": "string",
                        "description": "The space key the page lives in (e.g. 'DOCS').",
                    },
                },
                "required": ["title", "space"],
            },
        },
        {
            "name": "confluence_list_pages_under",
            "description": (
                "List pages located beneath a parent page in the page tree - "
                "use this for requests like 'pages under X' or 'child pages of "
                "X'. You do NOT need to write CQL; just identify the parent by "
                "its numeric 'parent_id', or by 'parent_title' plus 'space'. "
                "Set 'direct_only' to true for immediate children only, or "
                "leave it false to include all nested descendants. Optionally "
                "set 'modified_within_days' to only return pages changed "
                "recently (e.g. 30 for the past month). Returns IDs, titles, "
                "space keys and URLs, newest first; follow up with "
                "'confluence_get_page' to read each one."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "parent_id": {
                        "type": "string",
                        "description": "Numeric ID of the parent page (preferred if known).",
                    },
                    "parent_title": {
                        "type": "string",
                        "description": "Exact title of the parent page (needs 'space' too).",
                    },
                    "space": {
                        "type": "string",
                        "description": "Space key of the parent page (used with 'parent_title').",
                    },
                    "direct_only": {
                        "type": "boolean",
                        "description": "True = immediate children only; false = all descendants. Default false.",
                    },
                    "modified_within_days": {
                        "type": "integer",
                        "description": "Only include pages modified within this many days (optional).",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of results (1-50, default 25).",
                    },
                },
            },
        },
    ]


def tool_definitions(servers):
    """
    Return the tool list for this configuration.

    With one server it is base_tool_definitions() untouched. With two, every
    tool grows an optional 'server' argument whose enum lists the configured
    names, which is what lets a prompt like "find X on Blue" reach the right
    instance - and, just as importantly, lets Claude see that a second instance
    exists at all.
    """
    tools = base_tool_definitions()
    if not servers.multi:
        return tools

    names = servers.names()
    default_name = servers.default.name
    others = [n for n in names if n != default_name]
    example = others[0] if others else default_name
    server_property = {
        "type": "string",
        "enum": names,
        "description": (
            "Which Confluence server to query. Omit this for the default "
            "server ('{default}'). Set it only when the user names a server, "
            "e.g. \"...on {example}\" -> server=\"{example}\". Content IDs are "
            "NOT shared between servers, so read a page from the same server "
            "the search that found it used."
        ).format(default=default_name, example=example),
    }
    suffix = (
        " Queries the '{default}' server unless 'server' names another one "
        "({names})."
    ).format(default=default_name, names=", ".join(names))

    for tool in tools:
        tool["description"] += suffix
        schema = tool.setdefault("inputSchema", {})
        schema.setdefault("properties", {})["server"] = dict(server_property)
    return tools


def clamp_limit(value, default=25, lo=1, hi=50):
    """Coerce a user-supplied limit into a sane integer range."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, n))


def call_tool(servers, name, arguments):
    """
    Execute a named tool. Returns the text payload on success.
    Raises ConfluenceError (or ValueError) on a tool-domain failure, which the
    caller reports back as an MCP tool error (isError=true).

    The optional 'server' argument picks the Confluence instance; without it
    the call goes to the first configured server.
    """
    arguments = arguments or {}
    client = servers.resolve(arguments.get("server"))
    if name == "confluence_search":
        query = arguments.get("query")
        if not query:
            raise ConfluenceError("'query' is required")
        limit = clamp_limit(arguments.get("limit"))
        cql = 'text ~ "{}"'.format(cql_quote(str(query)))
        space = arguments.get("space")
        if space:
            cql += ' AND space = "{}"'.format(cql_quote(str(space)))
        return client.search(cql, limit)

    if name == "confluence_search_cql":
        cql = arguments.get("cql")
        if not cql:
            raise ConfluenceError("'cql' is required")
        limit = clamp_limit(arguments.get("limit"))
        return client.search(str(cql), limit)

    if name == "confluence_get_page":
        return client.get_page(arguments.get("page_id"))

    if name == "confluence_get_page_by_title":
        return client.get_page_by_title(
            arguments.get("title"), arguments.get("space")
        )

    if name == "confluence_list_pages_under":
        parent_id = arguments.get("parent_id")
        if not parent_id:
            # No explicit ID: resolve it from the parent's title + space.
            parent_id = client.resolve_page_id(
                arguments.get("parent_title"), arguments.get("space")
            )
        limit = clamp_limit(arguments.get("limit"))
        return client.list_pages_under(
            parent_id,
            direct_only=bool(arguments.get("direct_only", False)),
            modified_within_days=arguments.get("modified_within_days"),
            limit=limit,
        )

    raise ConfluenceError("Unknown tool: {}".format(name))


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------
def make_result(msg_id, result):
    return {"jsonrpc": "2.0", "id": msg_id, "result": result}


def make_error(msg_id, code, message):
    return {"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}}


def handle_initialize(params):
    # Echo the client's protocol version when it sends one, for compatibility.
    requested = ""
    if isinstance(params, dict):
        requested = params.get("protocolVersion") or ""
    protocol = requested if isinstance(requested, str) and requested else DEFAULT_PROTOCOL_VERSION
    return {
        "protocolVersion": protocol,
        "capabilities": {"tools": {}},
        "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
    }


def handle_message(servers, msg):
    """
    Process one JSON-RPC message object.
    Returns a response dict, or None for notifications (which get no reply).
    """
    if not isinstance(msg, dict):
        return make_error(None, INVALID_REQUEST, "Invalid Request: not an object")

    method = msg.get("method")
    msg_id = msg.get("id")
    is_request = "id" in msg  # notifications have no id and get no response

    if not isinstance(method, str):
        return make_error(msg_id, INVALID_REQUEST, "Missing method") if is_request else None

    # --- lifecycle / housekeeping ---
    if method == "initialize":
        return make_result(msg_id, handle_initialize(msg.get("params")))

    if method == "ping":
        return make_result(msg_id, {})

    if method.startswith("notifications/"):
        # e.g. notifications/initialized, notifications/cancelled - just ignore.
        return None

    # --- tools ---
    if method == "tools/list":
        return make_result(msg_id, {"tools": tool_definitions(servers)})

    if method == "tools/call":
        params = msg.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        if not name:
            # report as a tool error so the agent can recover
            return make_result(msg_id, {
                "content": [{"type": "text", "text": "Error: no tool name supplied."}],
                "isError": True,
            })
        try:
            text = call_tool(servers, name, arguments)
            return make_result(msg_id, {
                "content": [{"type": "text", "text": text}],
                "isError": False,
            })
        except (ConfluenceError, ValueError) as e:
            log("tool '{}' failed: {}".format(name, e))
            return make_result(msg_id, {
                "content": [{"type": "text", "text": "Error: {}".format(e)}],
                "isError": True,
            })
        except Exception as e:  # never let a tool crash the whole server
            log("unexpected error in tool '{}': {}".format(name, e))
            return make_result(msg_id, {
                "content": [{"type": "text", "text": "Unexpected error: {}".format(e)}],
                "isError": True,
            })

    # --- unknown method ---
    if is_request:
        return make_error(msg_id, METHOD_NOT_FOUND, "Method not found: {}".format(method))
    return None


def serve(servers):
    """
    Main stdio loop. MCP stdio framing is newline-delimited JSON: one JSON-RPC
    message per line, no embedded newlines, responses written the same way.
    """
    log("server started; waiting for JSON-RPC on stdin")
    stdin = sys.stdin
    while True:
        line = stdin.readline()
        if line == "":
            break  # EOF: the client closed the pipe
        line = line.strip()
        if not line:
            continue
        try:
            incoming = json.loads(line)
        except ValueError:
            _write(make_error(None, PARSE_ERROR, "Parse error: invalid JSON"))
            continue

        # JSON-RPC permits a batch (array) of messages. MCP's newer revisions
        # dropped batching, but we handle it defensively.
        if isinstance(incoming, list):
            responses = []
            for item in incoming:
                resp = handle_message(servers, item)
                if resp is not None:
                    responses.append(resp)
            if responses:
                _write(responses)
        else:
            resp = handle_message(servers, incoming)
            if resp is not None:
                _write(resp)

    log("stdin closed; shutting down")


def _write(obj):
    """
    Write a single JSON value as one line to stdout, then flush.

    ensure_ascii=True keeps the output pure ASCII (non-ASCII characters become
    \\uXXXX escapes, which are valid JSON). This is critical on Windows, where
    stdout defaults to a legacy code page (e.g. cp1252) that cannot encode many
    characters found in Confluence page bodies - writing them raw would raise
    UnicodeEncodeError and kill the server. main() also forces the streams to
    UTF-8 as a second layer of defence.
    """
    try:
        sys.stdout.write(json.dumps(obj, ensure_ascii=True) + "\n")
        sys.stdout.flush()
    except (BrokenPipeError, OSError):
        # The client closed the pipe; nothing useful we can do but stop.
        raise SystemExit(0)


# ---------------------------------------------------------------------------
# Entry point / configuration
# ---------------------------------------------------------------------------
def env_str(name):
    """
    Read an environment variable, treating blank as unset.

    A blank value is what an MCP client substitutes for an optional setting the
    user left empty (e.g. "${user_config.base_url_2}" with no second server), so
    it must mean "not configured" rather than "configured as empty". An
    unexpanded "${...}" placeholder - what a client leaves behind when the
    variable it refers to does not exist - means the same thing.
    """
    val = os.environ.get(name)
    if val is None:
        return None
    val = val.strip()
    if val.startswith("${") and val.endswith("}"):
        return None
    return val or None


def env_bool_opt(name):
    """Tri-state boolean env var: True, False, or None when unset/blank."""
    val = env_str(name)
    if val is None:
        return None
    return val.lower() not in ("0", "false", "no", "off")


def env_bool(name, default=True):
    val = env_bool_opt(name)
    return default if val is None else val


def env_int(name, default):
    """Integer env var, falling back to the default if unset or not a number."""
    val = env_str(name)
    if val is None:
        return default
    try:
        return int(val)
    except ValueError:
        log("WARNING: {} is not a whole number ({!r}); using {}.".format(
            name, val, default))
        return default


def clean(value):
    """Strip a CLI/env string value, mapping blank (and None) to None."""
    if value is None:
        return None
    value = str(value).strip()
    return value or None


def build_arg_parser():
    p = argparse.ArgumentParser(
        description="MCP server for querying Confluence Data Center (stdio transport).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    # --- first server (the default one) ---------------------------------
    p.add_argument("--name", default=env_str("CONFLUENCE_NAME"),
                   help="Friendly name for the first Confluence server, used to "
                        "pick it in a prompt (env CONFLUENCE_NAME, default "
                        "'{}'). Only shown when a second server is "
                        "configured.".format(DEFAULT_NAME_1))
    p.add_argument("--base-url", default=env_str("CONFLUENCE_BASE_URL"),
                   help="Confluence base URL incl. any context path, no trailing slash "
                        "(env CONFLUENCE_BASE_URL).")
    # SECURITY: credentials are deliberately env-var ONLY (CONFLUENCE_TOKEN, or
    # CONFLUENCE_USER + CONFLUENCE_PASSWORD; add the _2 suffix for the second
    # server). Command-line arguments are visible to other local users in
    # process listings, so no --token/--user/--password flags are offered.
    p.add_argument("--ca-cert", default=env_str("CONFLUENCE_CA_CERT"),
                   help="Path to a PEM CA bundle for an internal CA "
                        "(env CONFLUENCE_CA_CERT).")
    p.add_argument("--insecure", action="store_true",
                   default=not env_bool("CONFLUENCE_VERIFY_SSL", True),
                   help="Disable TLS certificate verification "
                        "(env CONFLUENCE_VERIFY_SSL=false).")

    # --- second server (optional) ---------------------------------------
    p.add_argument("--name-2", default=env_str("CONFLUENCE_NAME_2"),
                   help="Friendly name for the second Confluence server, e.g. "
                        "'Blue' - say it in a prompt to query that server "
                        "(env CONFLUENCE_NAME_2, default '{}').".format(DEFAULT_NAME_2))
    p.add_argument("--base-url-2", default=env_str("CONFLUENCE_BASE_URL_2"),
                   help="Base URL of a SECOND Confluence instance (env "
                        "CONFLUENCE_BASE_URL_2). Setting this enables the second "
                        "server, which then needs its own CONFLUENCE_TOKEN_2 (or "
                        "CONFLUENCE_USER_2 + CONFLUENCE_PASSWORD_2). Leave unset "
                        "for a single-server setup.")
    p.add_argument("--ca-cert-2", default=env_str("CONFLUENCE_CA_CERT_2"),
                   help="CA bundle for the second server (env "
                        "CONFLUENCE_CA_CERT_2). Falls back to --ca-cert.")
    # default=None so we can tell 'flag not given' from 'flag given': the second
    # server inherits the first server's TLS setting unless told otherwise.
    p.add_argument("--insecure-2", action="store_true", default=None,
                   help="Disable TLS certificate verification for the second "
                        "server only (env CONFLUENCE_VERIFY_SSL_2=false). "
                        "Defaults to whatever the first server uses.")

    # --- shared by both servers ------------------------------------------
    p.add_argument("--timeout", type=int,
                   default=env_int("CONFLUENCE_TIMEOUT", 30),
                   help="HTTP request timeout in seconds, both servers "
                        "(env CONFLUENCE_TIMEOUT).")
    p.add_argument("--max-body", type=int,
                   default=env_int("CONFLUENCE_MAX_BODY", 0),
                   help="Truncate page bodies to N characters, 0 = unlimited "
                        "(env CONFLUENCE_MAX_BODY). Applies to returned text "
                        "only, not to saved knowledge-base files.")
    p.add_argument("--kb-dir", default=env_str("CONFLUENCE_KB_DIR") or KB_DIR,
                   help="Every page read is also saved as a Markdown file into "
                        "this folder for a local RAG knowledge base (env "
                        "CONFLUENCE_KB_DIR, then the KB_DIR config value - "
                        "default C:\\Eva\\knowledge\\confluence). Files "
                        "are named 'Confluence - <title>.md' and overwritten "
                        "each time; with two servers, "
                        "'Confluence <server> - <title>.md'. Pass 'off' to "
                        "disable mirroring entirely.")
    p.add_argument("--check", action="store_true",
                   help="Connect to every configured Confluence server, print "
                        "who you are authenticated as and how many spaces are "
                        "visible (to stderr), then exit (no server).")
    p.add_argument("--version", action="version",
                   version="{0} {1}".format(SERVER_NAME, __version__))
    return p


def check_one(client, servers):
    """Connectivity check for one server: authenticate and count visible spaces."""
    if servers.multi:
        log("--- {}{} ---".format(
            client.name, " (default)" if client is servers.default else ""))
    log("Base URL         : {}".format(client.base_url))
    try:
        me = client._get("/rest/api/user/current")
        log("Authenticated as : {} ({})".format(
            me.get("displayName", "?"), me.get("username") or "?"))
        spaces = client._get("/rest/api/space", {"limit": 100})
        results = spaces.get("results")
        visible = len(results) if isinstance(results, list) else "?"
        log("Spaces visible   : {}{}".format(
            visible, "+" if spaces.get("_links", {}).get("next") else ""))
        return True
    except ConfluenceError as e:
        log("FAILED: {}".format(e))
        return False


def run_check(servers):
    """
    Connectivity check across every configured server. Each one is reported
    separately, and the exit code is non-zero if ANY of them failed - a
    two-server setup where only one instance answers is not a working setup.
    """
    failed = [c.name for c in servers.clients if not check_one(c, servers)]
    if servers.default.kb_dir:
        log("KB mirror folder : {}".format(servers.default.kb_dir))
    if failed:
        log("CHECK FAILED for {} of {} server(s): {}".format(
            len(failed), len(servers.clients), ", ".join(failed)))
        return 1
    log("CHECK OK")
    return 0


def main(argv=None):
    # Force the JSON-RPC streams to UTF-8. On Windows the console/pipe encoding
    # defaults to a legacy code page that cannot represent many characters in
    # Confluence content; without this, reading or writing such characters can
    # crash the server and the client then reports "not connected".
    for stream in (sys.stdin, sys.stdout):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            # reconfigure is unavailable (e.g. a redirected non-text stream);
            # _write still emits ASCII-only JSON, so output stays safe anyway.
            pass

    args = build_arg_parser().parse_args(argv)

    # Credentials come from the environment ONLY (never argv - see
    # build_arg_parser). The second server uses the same names with a _2 suffix.
    token = env_str("CONFLUENCE_TOKEN")
    user = env_str("CONFLUENCE_USER")
    password = env_str("CONFLUENCE_PASSWORD")
    token_2 = env_str("CONFLUENCE_TOKEN_2")
    user_2 = env_str("CONFLUENCE_USER_2")
    password_2 = env_str("CONFLUENCE_PASSWORD_2")

    base_url = clean(args.base_url)
    base_url_2 = clean(args.base_url_2)
    name = clean(args.name) or DEFAULT_NAME_1
    name_2 = clean(args.name_2) or DEFAULT_NAME_2

    # The knowledge-base folder has a real default, so "off" (or any other
    # DISABLE_KEYWORDS value) is how mirroring is switched off from a client
    # that can only pass strings - a blank value means "not configured" and
    # falls back to that default.
    kb_dir = clean(args.kb_dir)
    if kb_dir and kb_dir.strip().lower() in DISABLE_KEYWORDS:
        kb_dir = None

    if not base_url:
        log("FATAL: no base URL. Set CONFLUENCE_BASE_URL or pass --base-url. "
            "(The first server is always the default one; a second server is "
            "configured on top of it with CONFLUENCE_BASE_URL_2.)")
        return 2
    if not token and not (user and password):
        log("FATAL: no credentials. Set the CONFLUENCE_TOKEN environment "
            "variable, or CONFLUENCE_USER and CONFLUENCE_PASSWORD.")
        return 2

    # A second server is enabled by its base URL alone. If the rest of its
    # settings are present without it, say so - silently ignoring them would
    # send "on Blue" questions to the first server and answer them with the
    # wrong wiki's content.
    if not base_url_2 and (token_2 or user_2 or password_2 or clean(args.name_2)):
        log("WARNING: second-server settings are present but "
            "CONFLUENCE_BASE_URL_2 is not set, so only one server is "
            "configured. Set CONFLUENCE_BASE_URL_2 to enable the second one.")

    # spec = (name, base_url, token, user, password, verify_ssl, ca_cert)
    specs = [(name, base_url, token, user, password,
              not args.insecure, clean(args.ca_cert))]

    if base_url_2:
        if not token_2 and not (user_2 and password_2):
            log("FATAL: the second Confluence server ({}) has no credentials. "
                "Set CONFLUENCE_TOKEN_2, or CONFLUENCE_USER_2 and "
                "CONFLUENCE_PASSWORD_2 (each instance needs its own token). "
                "Unset CONFLUENCE_BASE_URL_2 to run with one server."
                .format(name_2))
            return 2
        if name_2.lower() == name.lower():
            log("FATAL: both Confluence servers are named {!r}. Give them "
                "different names via CONFLUENCE_NAME and CONFLUENCE_NAME_2 "
                "(e.g. Green and Blue) so a prompt can pick one.".format(name))
            return 2
        # TLS settings for the second server: flag, then its own env var, then
        # whatever the first server resolved to (usually the same internal CA).
        insecure_2 = args.insecure_2
        if insecure_2 is None:
            verify_env_2 = env_bool_opt("CONFLUENCE_VERIFY_SSL_2")
            insecure_2 = (not verify_env_2) if verify_env_2 is not None else args.insecure
        specs.append((name_2, base_url_2, token_2, user_2, password_2,
                      not insecure_2, clean(args.ca_cert_2) or clean(args.ca_cert)))

    try:
        clients = [
            ConfluenceClient(
                name=spec_name,
                base_url=spec_url,
                token=spec_token,
                user=spec_user,
                password=spec_password,
                verify_ssl=spec_verify,
                ca_cert=spec_ca,
                timeout=args.timeout,
                max_body=args.max_body,
                kb_dir=kb_dir,
            )
            for (spec_name, spec_url, spec_token, spec_user, spec_password,
                 spec_verify, spec_ca) in specs
        ]
        servers = ConfluenceServers(clients)
    except (ValueError, ssl.SSLError, OSError) as e:
        log("FATAL: could not initialise client: {}".format(e))
        return 2

    for client in servers.clients:
        if not client.verify_ssl:
            log("WARNING: TLS verification is disabled for {} ({}).".format(
                client.name, client.base_url))
    if servers.default.kb_dir:
        log("knowledge-base mirroring enabled -> {}".format(servers.default.kb_dir))
    if servers.multi:
        for position, client in enumerate(servers.clients, start=1):
            log("server {}: {} -> {}{}".format(
                position, client.name, client.base_url,
                "  (default)" if position == 1 else ""))
        if not clean(args.name) or not clean(args.name_2):
            log("TIP: set CONFLUENCE_NAME and CONFLUENCE_NAME_2 to memorable "
                "names (e.g. Green and Blue) so a prompt can say which server "
                "to use; currently {}.".format(" and ".join(servers.names())))
    else:
        log("configured for base URL {}".format(servers.default.base_url))

    if args.check:
        return run_check(servers)

    try:
        serve(servers)
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
