#!/usr/bin/env python3
"""
confluence.py (v5.0.0) - A single-file MCP (Model Context Protocol) server
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

MACROS
------
Most of what makes a Confluence page useful is in a macro, and a macro comes in
one of two kinds:

  IN THE PAGE SOURCE - an info/note/warning panel, an expand, a code block, an
    inline task list, a status lozenge. Confluence stores the content itself, so
    it can be read from the page source.
  GENERATED WHEN THE PAGE IS DISPLAYED - a task report, a page properties
    report, a children list, a page tree, Jira issues, an excerpt include. The
    page source holds ONLY the macro's settings; Confluence produces the content
    (the table of tasks, the list of pages) each time someone opens the page.

This server reads both. Page bodies come back as Markdown, with panels as block
quotes, expand macros shown open, code macros as fenced blocks, tables as
tables, and inline tasks as '- [ ]' / '- [x]' checkboxes. For the generated
kind it asks CONFLUENCE to render the page and reads the result, which is the
only way that content can be obtained - see CONFLUENCE_BODY_FORMAT below.

So "list the tasks on page X" now works whether the page carries inline task
checkboxes or a Task Report macro pointed at a person.

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

SAVING TO THE KNOWLEDGE BASE
----------------------------
Reading a page does NOT save it. Both page tools take an optional
'save_to_kb' argument, false by default; only when it is true is the page
written to CONFLUENCE_KB_DIR as Markdown for the local RAG index. That keeps
the knowledge base to pages you asked to keep, instead of every page skimmed
while answering a question - a search that turns up someone's meeting notes
should not put them in the index.

Set CONFLUENCE_KB_AUTOSAVE=true to go back to saving every page that is
read, which is how versions before 3.0.0 behaved.

CONFIGURATION
-------------
EVERY setting is an environment variable - the natural fit for an MCP client's
`env` block, and the reason there are no configuration flags: two settings can
then never disagree. The only command-line flags are --check and --version.

Two variables are shared with every other plugin in this suite, set once for
your Windows account:

  EVA_PYTHON            full path to the python.exe the MCP client launches,
                        e.g. C:\Python311\python.exe (read by the plugin
                        manifest, not by this file)
  EVA_KNOWLEDGE_DIR     root of the RAG corpus (default C:\Eva\knowledge).
                        This server saves into its own "confluence"
                        sub-folder, and THAT FOLDER MUST EXIST:
                        %EVA_KNOWLEDGE_DIR%\confluence

The rest are this server's own. CREDENTIALS ARE ENV-VAR ONLY - there is no
flag that could put a token in a command line, where other local users can
read it out of a process listing.

  First server (required):
  CONFLUENCE_NAME       friendly name used to select it in a prompt
                        (default "Primary")
  CONFLUENCE_BASE_URL   e.g. https://confluence.internal.example.com
                        (include any context path, no trailing slash)
  CONFLUENCE_TOKEN      Personal Access Token (preferred; sent as Bearer)
  CONFLUENCE_USER       username   } basic-auth fallback if no token is given
  CONFLUENCE_PASSWORD   password   }
  CONFLUENCE_VERIFY_SSL "false" to disable TLS verification (default: verify)
  CONFLUENCE_CA_CERT    path to a PEM CA bundle for an internal CA

  Second server (optional - set CONFLUENCE_BASE_URL_2 to enable it):
  CONFLUENCE_NAME_2       friendly name (default "Secondary")
  CONFLUENCE_BASE_URL_2   base URL of the second instance
  CONFLUENCE_TOKEN_2      its own Personal Access Token
  CONFLUENCE_USER_2       username   } basic-auth fallback for the second server
  CONFLUENCE_PASSWORD_2   password   }
  CONFLUENCE_VERIFY_SSL_2 TLS verification for the second server; falls back to
                          the first server's setting
  CONFLUENCE_CA_CERT_2    CA bundle for the second server; falls back to
                          CONFLUENCE_CA_CERT

  Shared by both servers:
  CONFLUENCE_BODY_FORMAT
                        which version of a page body to read (default "auto"):
                          auto        page source for an ordinary page, and the
                                      RENDERED page whenever the page uses a
                                      macro whose content Confluence generates
                                      at display time. This is the setting that
                                      makes task reports, page properties
                                      reports, children lists and Jira tables
                                      readable, while ordinary pages keep the
                                      cleaner source text.
                          view        always the rendered page
                          export_view the render Confluence uses for PDF/Word
                                      export. Try this if a macro is STILL
                                      empty under "view" - a page tree, and
                                      some third-party macros, only render
                                      statically on export
                          storage     the raw page source only, with no macro
                                      output (how this server behaved before
                                      v5.0.0)
                        Each page tool also takes a per-call 'body_format'
                        argument, so a page can be re-read rendered without
                        changing this setting.
  CONFLUENCE_TIMEOUT    request timeout in seconds (default: 30)
  CONFLUENCE_MAX_BODY   truncate page bodies to N chars (0 = unlimited,
                        default 0). This limit applies only to the text
                        returned to the model; saved files are never
                        truncated.
  CONFLUENCE_KB_DIR     override the save folder with a full path of its own,
                        instead of the "confluence" sub-folder of
                        EVA_KNOWLEDGE_DIR. Keep it inside the knowledge-base
                        plugin's corpus or saved pages are never indexed. Set
                        it to "off" to forbid saving entirely, after which the
                        server writes no local file at all and a save_to_kb
                        request is refused.
  CONFLUENCE_KB_AUTOSAVE
                        "true" to save EVERY page that is read, without being
                        asked (default false).

The second server is all-or-nothing: if CONFLUENCE_BASE_URL_2 is set without
credentials, the server refuses to start rather than quietly answering "Blue"
questions from the first instance.

INSTALLING INTO CLAUDE CODE
---------------------------
This server ships as the "confluence" Claude Code plugin (its manifest is
.claude-plugin/plugin.json next to this file), so the normal install is:

    /plugin marketplace add C:\path\to\claude-skills
    /plugin install confluence@mcnamee-claude-skills

Claude Code prompts for each server's name and base URL; the interpreter and
the knowledge folder come from the shared variables above. Tokens are NOT
stored in the plugin - set them as Windows user environment variables before
starting Claude Code, and the plugin picks them up from there:

    setx EVA_PYTHON         "C:\Python311\python.exe"
    setx EVA_KNOWLEDGE_DIR  "C:\Eva\knowledge"
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

If a page reads as empty where you expected a table of tasks or rows, the
'Body:' line in the tool's output says which representation was used, and a
page read from the source names each macro whose content was not fetched. Read
it again with body_format="view", then body_format="export_view".

A quick two-server smoke test from PowerShell, without touching the plugin
config (every setting, tokens included, goes in the environment):

    $env:CONFLUENCE_NAME       = "Green"
    $env:CONFLUENCE_BASE_URL   = "https://green.confluence.example.com"
    $env:CONFLUENCE_TOKEN      = "green-token"
    $env:CONFLUENCE_NAME_2     = "Blue"
    $env:CONFLUENCE_BASE_URL_2 = "https://blue.confluence.example.com"
    $env:CONFLUENCE_TOKEN_2    = "blue-token"
    & "C:\path\to\python.exe" confluence.py --check
"""

# Semantic version of this server. Bump on EVERY change (see CLAUDE.md):
# MAJOR = breaking config/tool change, MINOR = new feature, PATCH = fix.
__version__ = "5.0.0"

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

# Folder a page is saved into as Markdown, for a local RAG index, WHEN a tool
# call asks for it (save_to_kb=true).
# RESOLVED FROM THE ENVIRONMENT in main(): the "confluence" sub-folder of
# %EVA_KNOWLEDGE_DIR% (the suite-wide RAG root), or CONFLUENCE_KB_DIR for a
# full path of its own. The literal here is what a stock C:\Eva install
# resolves to; it MUST stay inside the knowledge-base plugin's corpus
# (C:\Eva\knowledge) or the saved pages would never be indexed. The folder is
# created on demand.
# Set CONFLUENCE_KB_DIR=off to forbid saving altogether, after which this
# server touches no local file at all.
SUBFOLDER = "confluence"                 # this server's knowledge sub-folder
EVA_KNOWLEDGE_DIR = r"C:\Eva\knowledge"  # fallback for the suite-wide root
KB_DIR = r"C:\Eva\knowledge\confluence"

# Which representation of a page body to read. Confluence keeps the page SOURCE
# in "storage" and the RENDERED page in "view"/"export_view"; a macro that
# generates its content at display time (task report, page properties report,
# children list, Jira issues) has NO content in the source, so a reader that
# only looks at storage reports those pages as empty.
#   auto        - storage for a page that carries all its own content, a
#                 rendered body for a page that uses a generated macro. Default.
#   view        - always the rendered page (what the browser shows)
#   export_view - the render Confluence uses for PDF/Word export; try this when
#                 a macro still comes back empty under "view" (the page tree and
#                 some third-party macros only render statically on export)
#   storage     - the raw page source only; no macro output at all
# Override with CONFLUENCE_BODY_FORMAT, or per call with the body_format
# argument on the two page tools.
BODY_FORMAT = "auto"
BODY_FORMATS = ("auto", "view", "export_view", "storage")

# Whether reading a page saves it WITHOUT being asked. False means a page is
# saved only when the caller passes save_to_kb=true, which is the point: a
# search that turns up an unrelated page should not add it to the knowledge
# base just because it was read while answering. Set to True here, or set
# CONFLUENCE_KB_AUTOSAVE=true, to save every page read.
KB_AUTOSAVE = False

# Folder-setting values that mean "explicitly turned off". An MCP client can
# only pass strings, and a BLANK string is what it substitutes for a setting the
# user left empty - which means "not configured", falling back to the
# suite-wide root. So a keyword is needed to say "definitely off".
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
# Confluence macros
# ---------------------------------------------------------------------------
# Confluence stores a macro as <ac:structured-macro ac:name="..."> with its
# settings as <ac:parameter> children. TWO KINDS of macro live in that markup,
# and they need opposite treatment:
#
#   STATIC  - the macro's content IS in the page source, inside
#             <ac:rich-text-body> or <ac:plain-text-body> (an info panel, a
#             code block, an expand). Reading the storage format is enough.
#   DYNAMIC - the macro's content is GENERATED by Confluence when the page is
#             displayed (a task report, a page-properties report, a children
#             list, a Jira issue table). The storage format holds the macro's
#             SETTINGS ONLY, so a reader that only looks at storage sees an
#             empty macro and reports the page as having no tasks/rows/issues.
#
# The fix for the dynamic kind is to ask Confluence for a RENDERED body
# (body.view or body.export_view) instead of body.storage - see BODY_FORMAT.
# The set below is the allowlist of macros known to carry their content in the
# page source; ANY other macro is assumed dynamic, so a third-party macro this
# file has never heard of still triggers a rendered fetch rather than silently
# reading as empty.
STATIC_MACROS = frozenset((
    "code", "noformat", "panel", "info", "note", "tip", "warning", "expand",
    "excerpt", "multiexcerpt", "details", "section", "column", "align",
    "anchor", "status", "toc",
    "toc-zone", "highlight", "quote", "html-bullet", "div", "span",
    "localtabgroup", "localtab", "tabs-page", "tabsetup", "deck", "card",
    "bookmark", "cheese", "color", "font", "note-macro", "hidden-comment",
))

# Macros whose rendered output is worth naming in a placeholder when only the
# storage format is available, mapped to a human phrase. Anything not listed
# falls back to the macro's own name.
DYNAMIC_MACRO_LABELS = {
    "tasks-report-macro": "task report",
    "tasks-report": "task report",
    "detailssummary": "page properties report",
    "details": "page properties",
    "children": "child-page list",
    "pagetree": "page tree",
    "contentbylabel": "content-by-label list",
    "recently-updated": "recently updated list",
    "include": "included page",
    "excerpt-include": "included excerpt",
    "multiexcerpt-include": "included multi-excerpt",
    "jira": "Jira issues",
    "jiraissues": "Jira issues",
    "jirachart": "Jira chart",
    "blog-posts": "blog post list",
    "livesearch": "search box",
    "attachments": "attachment list",
    "gallery": "image gallery",
    "chart": "chart",
    "roadmap": "roadmap planner",
    "calendar": "team calendar",
    "drawio": "draw.io diagram",
    "gliffy": "Gliffy diagram",
    "viewpdf": "embedded PDF",
    "viewxls": "embedded spreadsheet",
    "widget": "embedded widget",
    "profile": "user profile",
    "contributors": "contributor list",
    "spacedetails": "space details",
    "create-from-template": "create-page button",
}

# Matches a macro invocation in storage-format XHTML, capturing its name. Used
# only to decide whether a page needs a rendered body; the real parsing is done
# by the HTML parser below.
_MACRO_NAME_RE = re.compile(
    r"<ac:(?:structured-)?macro\b[^>]*?\bac:name\s*=\s*[\"']([^\"']+)[\"']",
    re.IGNORECASE,
)


def storage_macro_names(raw):
    """Return the set of macro names used in a storage-format body."""
    if not raw:
        return set()
    return {m.lower() for m in _MACRO_NAME_RE.findall(raw)}


def has_dynamic_macro(raw):
    """
    True if a storage-format body uses a macro whose content Confluence
    generates at display time, so reading storage alone would lose it.

    Unknown macros count as dynamic on purpose: guessing "static" for a macro
    we have never seen would silently drop its content, while guessing
    "dynamic" only costs a rendered fetch we were going to be able to use.
    """
    return any(name not in STATIC_MACROS for name in storage_macro_names(raw))


def normalise_body_format(value, default=None):
    """
    Coerce a body-format setting to one of BODY_FORMATS, or the default.

    Accepts the hyphenated spellings a user is likely to type ("export-view").
    Returns the default for anything unrecognised so a typo cannot stop the
    server; the caller warns where a warning is useful.
    """
    if value is None:
        return default
    want = str(value).strip().lower().replace("-", "_")
    if not want:
        return default
    if want in ("rendered", "display"):
        want = "view"
    if want in ("export", "exportview"):
        want = "export_view"
    if want in ("source", "raw"):
        want = "storage"
    return want if want in BODY_FORMATS else default


def _blockquote(text):
    """Prefix every line of text with '> ' so it reads as a Markdown quote."""
    lines = text.strip().split("\n")
    return "\n".join(("> " + ln) if ln.strip() else ">" for ln in lines)


def _format_params(params, skip=()):
    """Render a macro's parameters as 'key=value, key=value' for a placeholder."""
    bits = []
    for key, val in params.items():
        if key in skip or not val:
            continue
        flat = " ".join(str(val).split())
        if len(flat) > 120:
            flat = flat[:117] + "..."
        bits.append("{}={}".format(key or "(unnamed)", flat))
    return ", ".join(bits)


# ---------------------------------------------------------------------------
# HTML -> plain text
# ---------------------------------------------------------------------------
class _TextExtractor(html.parser.HTMLParser):
    """
    Minimal HTML/XHTML to plain-text converter.

    This is the defensive fallback for html_to_markdown; the Markdown extractor
    below is what normally runs. It keeps the readable text and inserts line
    breaks around block-level elements. Macro PARAMETERS are skipped - they are
    settings, not content, and letting them through produced runs of glued-together
    values like 'DEVjsmithtrue' in the middle of a page.
    convert_charrefs=True (the default) means entities like &amp; are decoded
    for us and arrive via handle_data.
    """

    _BLOCK_TAGS = {
        "p", "br", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
        "table", "ul", "ol", "blockquote", "pre", "section", "header",
        "footer", "article", "ac:task", "ac:task-list", "ac:structured-macro",
    }
    _SKIP_TAGS = {"script", "style", "ac:parameter", "ac:task-id"}

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

    def unknown_decl(self, data):
        # CDATA sections (code-macro bodies, link labels) arrive here.
        if self._skip_depth:
            return
        if data.startswith("CDATA["):
            inner = data[6:]
            if inner.endswith("]"):
                inner = inner[:-1]
            self._parts.append(inner)

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
    Convert Confluence XHTML into reasonable Markdown.

    It runs over BOTH shapes of body this server fetches: the storage format
    (the page source, with <ac:...> macro markup) and a rendered body
    (body.view / body.export_view, which is ordinary HTML because Confluence has
    already run the macros). It is a best-effort converter aimed at reading and
    RAG ingestion, not a pixel-perfect renderer.

    Structural HTML is handled directly (headings, paragraphs, lists,
    bold/italic, links, inline code, code blocks, block quotes, rules, tables).
    On top of that it understands the Confluence-specific markup that a generic
    HTML reader turns into mush:

      - <ac:structured-macro> is rendered per macro: code macros become fenced
        blocks, info/note/warning/panel become block quotes, expand shows its
        (otherwise collapsed) contents, status becomes an inline label. A macro
        whose content Confluence generates at display time leaves a placeholder
        naming the macro and its settings, so the reader knows content is
        MISSING rather than absent.
      - <ac:parameter> is captured as a macro setting instead of being emitted
        as loose text.
      - <ac:task-list>/<ac:task> become '- [ ] ' / '- [x] ' checkboxes. This is
        how Confluence stores inline tasks, so "list the tasks on page X" works
        off the page source alone.
      - <ac:link>/<ri:page>/<ri:user>/<ri:attachment> become readable links and
        @mentions, and <ac:image> names its attachment.
      - In a RENDERED body, a task list arrives as <li class="inline-task-list
        checked">; the class is read so the checkbox state survives there too.

    Text is not Markdown-escaped, so the occasional literal '*' may look like
    emphasis; that is a deliberate trade-off to keep the captured text faithful
    for search.
    """

    _SKIP_TAGS = {"script", "style"}
    # Macro wrappers whose children should render normally: they exist to hold
    # content, and add nothing themselves.
    _PASSTHROUGH_TAGS = {
        "ac:rich-text-body", "ac:layout", "ac:layout-section", "ac:layout-cell",
        "ac:inline-comment-marker", "ac:adf-node", "ac:adf-content",
    }

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        # Output is written to the sink on top of this stack. Pushing a sink is
        # how a table cell, a macro body or a macro parameter is captured and
        # re-emitted in a different shape once its closing tag arrives.
        self._sinks = [self.parts]
        self._skip_depth = 0
        self._in_pre = 0
        self._list_stack = []      # 'ul' / 'ol' per nesting level
        self._ol_counters = []     # running item number per ordered-list level
        self._href_stack = []      # href per currently-open <a>
        # Table buffering (only the outermost table is rendered as a grid)
        self._table_depth = 0
        self._rows = None          # list of cell-lists for the current table
        self._row = None           # current row (list of cell strings)
        self._cell_open = False    # True while a <td>/<th> sink is pushed
        # Confluence macro state
        self._macros = []          # stack of open macro dicts
        self._param_names = []     # ac:name per currently-open <ac:parameter>
        self._plain_body = []      # tag name per open <ac:plain-text-*-body>
        self._tasks = []           # stack of open <ac:task> dicts
        self._links = []           # target/anchor per open <ac:link>

    # -- output sinks -------------------------------------------------------
    def _emit(self, s):
        self._sinks[-1].append(s)

    def _push_sink(self):
        self._sinks.append([])

    def _pop_sink(self):
        # Never pop the base sink: a malformed document must not lose the page.
        if len(self._sinks) == 1:
            return ""
        return "".join(self._sinks.pop())

    # -- macro helpers ------------------------------------------------------
    def _macro_param(self, name, default=""):
        return self._macros[-1]["params"].get(name, default) if self._macros else default

    def _render_macro(self, macro):
        """Turn one finished <ac:structured-macro> into Markdown."""
        name = (macro["name"] or "").lower()
        params = macro["params"]
        body = macro["body"].strip()
        plain = macro["plain"]
        title = (params.get("title") or params.get("name") or "").strip()

        if name in ("code", "noformat"):
            lang = (params.get("language") or "").strip()
            if lang.lower() in ("none", "text"):
                lang = ""
            source = plain if plain is not None else body
            head = "\n\n**{}**\n".format(title) if title else "\n\n"
            return "{}```{}\n{}\n```\n\n".format(head, lang, (source or "").strip("\n"))

        if name in ("info", "note", "tip", "warning", "panel", "warn"):
            label = {"warn": "Warning"}.get(name, name.capitalize())
            heading = "**{}:** {}".format(label, title) if title else "**{}**".format(label)
            inner = body if body else "(empty)"
            return "\n\n" + _blockquote(heading + "\n\n" + inner) + "\n\n"

        if name == "expand":
            heading = title or "Click here to expand"
            # Expand macros hide real content behind a toggle; show it.
            return "\n\n**{}**\n\n{}\n\n".format(heading, body)

        if name == "status":
            text = title or (params.get("colour") or params.get("color") or "status")
            return "**[{}]**".format(text.strip().upper())

        if name in ("toc", "toc-zone", "anchor", "bookmark"):
            return "\n" + body + "\n" if body else ""

        if name in ("excerpt", "section", "column", "align", "div", "span",
                    "highlight", "quote", "color", "font"):
            # Layout-only wrappers: keep the content, drop the wrapper.
            return "\n\n" + body + "\n\n" if body else ""

        if body:
            # A macro we have no special rendering for, but which carries its
            # content in the page source: keep the content, name the macro.
            return "\n\n*[{} macro]*\n\n{}\n\n".format(name, body)
        if plain:
            return "\n\n*[{} macro]*\n\n{}\n\n".format(name, plain.strip())

        # No content in the source at all: this is a macro Confluence renders at
        # display time. Say so explicitly, with its settings, so the reader knows
        # the content exists but was not fetched - and how to fetch it.
        label = DYNAMIC_MACRO_LABELS.get(name, "{} macro".format(name))
        settings = _format_params(params)
        return (
            "\n\n[Confluence {label} (macro \"{name}\"{settings}). Its content is "
            "generated by Confluence when the page is displayed and is NOT in the "
            "page source, so it is not shown here. Read this page again with "
            "body_format=\"view\" to get the rendered result.]\n\n"
        ).format(
            label=label, name=name,
            settings=", " + settings if settings else "",
        )

    def _resource(self, text):
        """
        Handle an <ri:...> resource reference. Inside an <ac:link> it names the
        link's TARGET (kept aside so the link body can be used as the label);
        anywhere else it is simply the text to show.
        """
        if not text:
            return
        if self._links and not self._links[-1]["target"]:
            self._links[-1]["target"] = text
        else:
            self._emit(text)

    def _render_task(self, task):
        """Turn one finished <ac:task> into a Markdown checkbox line."""
        status = (task["status"] or "").strip().lower()
        box = "x" if status in ("complete", "completed", "done", "checked") else " "
        text = " ".join(task["body"].split())
        return "\n- [{}] {}".format(box, text)

    # -- parser callbacks ---------------------------------------------------
    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        attrd = {k.lower(): (v or "") for k, v in attrs}

        # --- Confluence-specific markup ---
        if tag in ("ac:structured-macro", "ac:macro"):
            self._macros.append({
                "name": attrd.get("ac:name", ""),
                "params": {},
                "body": "",
                "plain": None,
            })
            self._push_sink()
            return
        if tag == "ac:parameter":
            self._param_names.append(attrd.get("ac:name", ""))
            self._push_sink()
            return
        if tag in ("ac:plain-text-body", "ac:plain-text-link-body"):
            self._plain_body.append(tag)
            self._push_sink()
            return
        if tag == "ac:task-list":
            self._emit("\n\n")
            return
        if tag == "ac:task":
            self._tasks.append({"status": "", "body": ""})
            self._push_sink()
            return
        if tag in ("ac:task-id", "ac:task-uuid", "ac:task-status", "ac:task-body"):
            self._push_sink()
            return
        if tag in self._PASSTHROUGH_TAGS:
            return
        if tag == "ac:link":
            # The link TARGET arrives as a child element (<ri:page>, <ri:user>,
            # <ri:attachment>) while the LABEL is the element's body, so both are
            # collected and combined when the closing tag arrives.
            self._links.append({"target": "", "anchor": attrd.get("ac:anchor", "")})
            self._push_sink()
            return
        if tag == "ac:image":
            self._push_sink()
            return
        if tag == "ri:page":
            self._resource(attrd.get("ri:content-title", ""))
            return
        if tag in ("ri:user", "ri:mention"):
            who = (attrd.get("ri:username") or attrd.get("ri:account-id")
                   or attrd.get("ri:userkey") or "")
            self._resource("@" + who if who else "@user")
            return
        if tag in ("ri:attachment", "ri:blog-post", "ri:space", "ri:content-entity"):
            self._resource(attrd.get("ri:filename") or attrd.get("ri:content-title")
                           or attrd.get("ri:space-key") or "")
            return
        if tag == "ac:emoticon":
            return
        if tag.startswith("ac:") or tag.startswith("ri:"):
            # Unknown Confluence element: render its children, ignore the tag.
            return

        # --- ordinary HTML ---
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
            self._href_stack.append(attrd.get("href", ""))
            self._emit("[")
        elif tag == "ul":
            self._list_stack.append("ul")
        elif tag == "ol":
            self._list_stack.append("ol")
            self._ol_counters.append(0)
        elif tag == "li":
            indent = "  " * max(0, len(self._list_stack) - 1)
            classes = attrd.get("class", "").lower().split()
            if "checked" in classes or "unchecked" in classes:
                # A rendered inline task: keep the tick state the class carries.
                marker = "- [x] " if "checked" in classes else "- [ ] "
            elif self._list_stack and self._list_stack[-1] == "ol":
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
            # Only the outermost table is rendered as a grid; a nested table's
            # rows must not clobber the row being built for the outer one.
            if self._rows is not None and self._table_depth == 1:
                self._row = []
        elif tag in ("td", "th"):
            if (self._row is not None and not self._cell_open
                    and self._table_depth == 1):
                self._cell_open = True
                self._push_sink()
            elif self._table_depth > 1:
                self._emit(" ")     # nested table, flattened into its cell

    def handle_endtag(self, tag):
        if tag in self._SKIP_TAGS:
            if self._skip_depth:
                self._skip_depth -= 1
            return
        if self._skip_depth:
            return

        # --- Confluence-specific markup ---
        if tag in ("ac:structured-macro", "ac:macro"):
            body = self._pop_sink()
            if self._macros:
                macro = self._macros.pop()
                macro["body"] = body
                self._emit(self._render_macro(macro))
            else:
                self._emit(body)
            return
        if tag == "ac:parameter":
            value = self._pop_sink().strip()
            key = self._param_names.pop() if self._param_names else ""
            if self._macros:
                self._macros[-1]["params"][key] = value
            elif value:
                self._emit(value)
            return
        if tag in ("ac:plain-text-body", "ac:plain-text-link-body"):
            value = self._pop_sink()
            if self._plain_body:
                self._plain_body.pop()
            if tag == "ac:plain-text-body" and self._macros:
                # Kept apart from the rich-text body so the code macro can fence
                # it without the macro's own parameters leaking in.
                self._macros[-1]["plain"] = value
            else:
                self._emit(value)
            return
        if tag == "ac:task-list":
            self._emit("\n\n")
            return
        if tag == "ac:task":
            leftover = self._pop_sink()
            if self._tasks:
                task = self._tasks.pop()
                if not task["body"]:
                    task["body"] = leftover
                self._emit(self._render_task(task))
            else:
                self._emit(leftover)
            return
        if tag in ("ac:task-id", "ac:task-uuid"):
            self._pop_sink()        # internal identifiers: not content
            return
        if tag == "ac:task-status":
            value = self._pop_sink().strip()
            if self._tasks:
                self._tasks[-1]["status"] = value
            return
        if tag == "ac:task-body":
            value = self._pop_sink()
            if self._tasks:
                self._tasks[-1]["body"] = value
            else:
                self._emit(value)
            return
        if tag == "ac:link":
            label = " ".join(self._pop_sink().split())
            link = self._links.pop() if self._links else {"target": "", "anchor": ""}
            target = link["target"]
            if link["anchor"]:
                target = (target + "#" + link["anchor"]) if target else ("#" + link["anchor"])
            if label and target and label != target:
                self._emit("[{}]({})".format(label, target))
            else:
                self._emit(label or target)
            return
        if tag == "ac:image":
            inner = " ".join(self._pop_sink().split())
            self._emit("[image: {}]".format(inner) if inner else "[image]")
            return
        if tag in self._PASSTHROUGH_TAGS:
            if tag in ("ac:layout-section", "ac:layout-cell"):
                self._emit("\n\n")
            return
        if tag.startswith("ac:") or tag.startswith("ri:"):
            return

        # --- ordinary HTML ---
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
            if self._table_depth == 1:
                self._close_cell()
            else:
                # A nested table is flattened into the cell that contains it;
                # separate its values so they do not run together.
                self._emit(" ")
        elif tag == "tr":
            if self._table_depth == 1:
                self._close_cell()  # tolerate a row that closes with a cell open
                if self._row is not None and self._rows is not None:
                    self._rows.append(self._row)
                    self._row = None
        elif tag == "table":
            if self._table_depth == 1:
                # Flush whatever is still open. Markup that omits </td> or </tr>
                # (legal in HTML, and seen in some rendered bodies) would
                # otherwise drop the last row on the floor.
                self._close_cell()
                if self._row and self._rows is not None:
                    self._rows.append(self._row)
                self._row = None
                rows = self._rows or []
                self._rows = None
                self._emit_table(rows)
            if self._table_depth:
                self._table_depth -= 1

    def _close_cell(self):
        """Finish the open table cell, if any, appending it to the current row."""
        if not self._cell_open:
            return
        # Markdown cells are single-line: flatten and escape pipes.
        cell_text = " ".join(self._pop_sink().split())
        self._cell_open = False
        if self._row is not None:
            self._row.append(cell_text.replace("|", "\\|"))

    def handle_startendtag(self, tag, attrs):
        # XHTML self-closing element, e.g. <ri:user ri:userkey="..."/>. The base
        # class calls start then end, which is right for every tag we handle.
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_data(self, data):
        if self._skip_depth:
            return
        if self._in_pre or self._plain_body:
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
        self._emit("\n\n" + "\n".join(out) + "\n\n")

    def get_markdown(self):
        # Flush any sink left open by malformed markup, outermost last, so no
        # content is lost when a document ends mid-element.
        while len(self._sinks) > 1:
            leftover = self._pop_sink()
            if leftover:
                self._emit(leftover)
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
                 kb_dir=None, kb_autosave=False, body_format=BODY_FORMAT):
        if not name:
            raise ValueError("name is required")
        if not base_url:
            raise ValueError("base_url is required")
        self.name = name
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_body = max_body
        # Which body representation to read; see BODY_FORMAT at the top of this
        # file. A per-call body_format argument overrides it.
        self.body_format = normalise_body_format(body_format, BODY_FORMAT)
        # Folder to save pages into as Markdown; None/empty forbids saving.
        self.kb_dir = kb_dir or None
        # Whether a page read without an explicit save_to_kb argument is saved
        # anyway. Off by default - saving is something the user asks for.
        self.kb_autosave = bool(kb_autosave)
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

    def _body_expand(self, body_format=None):
        """
        The 'expand' value needed to fetch the body this call wants.

        The storage format is always requested alongside a rendered one: it
        costs nothing (Confluence stores it verbatim), it is the fallback when
        an instance does not return the rendered representation, and in 'auto'
        mode it is what decides whether the rendered body is needed at all.
        """
        fmt = body_format or self.body_format
        if fmt == "storage":
            return "body.storage"
        if fmt == "export_view":
            return "body.export_view,body.storage"
        return "body.view,body.storage"

    def _pick_body(self, page, body_format=None):
        """
        Choose which representation of the body to read, returning
        (html, representation_name).

        Confluence keeps a page's SOURCE in 'storage' and its RENDERED output in
        'view'/'export_view'. Macros that generate their content at display time
        - a task report, a page-properties report, a children list, Jira issues -
        appear in storage as an empty macro tag with settings, so reading storage
        alone reports those pages as having no content. Reading a rendered body
        is the only way to get that content, because Confluence itself is what
        produces it.

        'auto' (the default) picks per page: storage when the page carries all
        its own content (cleaner text, no display chrome), the rendered body when
        the page uses a macro whose output is generated.
        """
        fmt = body_format or self.body_format
        bodies = page.get("body") or {}

        def value(rep):
            return ((bodies.get(rep) or {}).get("value")) or ""

        storage = value("storage")
        if fmt == "storage":
            return storage, "storage"
        preferred = "export_view" if fmt == "export_view" else "view"
        if fmt == "auto" and storage and not has_dynamic_macro(storage):
            return storage, "storage"
        rendered = value(preferred)
        if rendered.strip():
            return rendered, preferred
        # The instance did not return the rendered body (older API, or a render
        # failure). Storage still holds everything except the generated macros,
        # whose placeholders then say what is missing.
        return storage, "storage"

    def _body_note(self, representation, page_html):
        """One line for the output header saying where the body came from."""
        if representation == "storage":
            if has_dynamic_macro(page_html):
                return ("storage (page source) - this page uses macros whose "
                        "content Confluence generates when the page is "
                        "displayed; read it again with body_format=\"view\" to "
                        "include them")
            return "storage (page source)"
        return "{} (rendered by Confluence, so macro content is included)".format(
            representation)

    def _render_page(self, page, save_to_kb=None, body_format=None):
        """
        Format a single content object as Markdown.

        save_to_kb decides whether the page is also written to the
        knowledge-base folder: True on request, False to skip, None ("the
        caller did not say") to follow the kb_autosave setting, which is off
        unless the endpoint deliberately turned it on.

        body_format overrides this server's CONFLUENCE_BODY_FORMAT for one call
        (see _pick_body).
        """
        title = page.get("title", "(untitled)")
        cid = page.get("id", "?")
        ctype = page.get("type", "content")
        space = (page.get("space") or {}).get("key", "?")
        version = (page.get("version") or {}).get("number", "?")
        link = self._abs_link(page, (page.get("_links") or {}).get("webui", ""))
        body_html, representation = self._pick_body(page, body_format)
        # Markdown, not plain text: a task list, a task report and a page
        # properties table are all TABULAR, and flattening them to prose is what
        # made "list the tasks on page X" unanswerable.
        text = html_to_markdown(body_html)
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
        fields.extend([("Version", version), ("URL", link),
                       ("Body", self._body_note(representation, body_html))])
        meta = "".join("{}: {}\n".format(k, v) for k, v in fields) + "\n--- Content ---\n"
        rendered = meta + (text if text else "(this page has no readable body content)") + truncated_note

        # Save the page to Markdown only if this call asked for it (or the
        # endpoint turned autosave on). Saving must never break the read, so
        # any failure is reported but swallowed.
        wanted = self.kb_autosave if save_to_kb is None else bool(save_to_kb)
        if wanted and not self.kb_dir:
            # A save was asked for that this endpoint has switched off. Say so,
            # so it cannot look done. (Autosave with no folder is a startup
            # warning instead - repeating it on every page would be noise.)
            if save_to_kb:
                rendered += (
                    "\n\n[NOT saved: knowledge-base saving is switched off for "
                    "this server. Unset CONFLUENCE_KB_DIR (or point it at "
                    "a folder inside the knowledge base) to enable it.]"
                )
        elif wanted:
            try:
                path = self._save_to_kb(title, link, space, version,
                                        body_html, representation)
                log("saved page to knowledge base: {}".format(path))
                rendered += "\n\n[Saved to knowledge base: {}]".format(path)
            except OSError as e:
                log("knowledge-base save failed: {}".format(e))
                rendered += "\n\n[Knowledge-base save FAILED: {}]".format(e)
        return rendered

    def _save_to_kb(self, title, link, space, version, body_html, representation):
        """
        Write the page to '<kb_dir>/Confluence - <title>.md', overwriting any
        existing file. Returns the path written; raises OSError on failure.

        With two servers configured the filename becomes
        'Confluence <server> - <title>.md', so a page with the same title on
        both instances produces two files instead of one overwriting the other.

        The saved body is the SAME representation that was read (see
        _pick_body), so a page saved for the RAG index carries the macro content
        the reader saw rather than an empty macro tag.

        The FULL body is always saved (the CONFLUENCE_MAX_BODY limit only trims
        what is returned to the model, not what is stored for RAG).
        """
        md_body = html_to_markdown(body_html)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        server_line = "- Server: {}\n".format(self.name) if self.label_output else ""
        header = (
            "# {title}\n\n"
            "- Source: {url}\n"
            "{server_line}"
            "- Space: {space}\n"
            "- Version: {version}\n"
            "- Body: {representation}\n"
            "- Fetched: {stamp}\n\n"
            "---\n\n"
        ).format(title=title, url=link or "(unknown)", server_line=server_line,
                 space=space, version=version, representation=representation,
                 stamp=stamp)
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

    def get_page(self, page_id, save_to_kb=None, body_format=None):
        if page_id is None or str(page_id).strip() == "":
            raise ConfluenceError("'page_id' is required")
        page_id = str(page_id).strip()
        body_format = normalise_body_format(body_format) or self.body_format
        params = {"expand": self._body_expand(body_format) + ",version,space"}
        page = self._get("/rest/api/content/" + urllib.parse.quote(page_id, safe=""), params)
        return self._render_page(page, save_to_kb=save_to_kb,
                                 body_format=body_format)

    def get_page_by_title(self, title, space, save_to_kb=None, body_format=None):
        if not title or not space:
            raise ConfluenceError("Both 'title' and 'space' are required")
        body_format = normalise_body_format(body_format) or self.body_format
        params = {
            "title": title,
            "spaceKey": space,
            "expand": self._body_expand(body_format) + ",version,space",
            "limit": 1,
        }
        data = self._get("/rest/api/content", params)
        results = data.get("results", []) or []
        if not results:
            return "No page titled {!r} found in space {!r}{}.".format(
                title, space, self._on_server())
        return self._render_page(results[0], save_to_kb=save_to_kb,
                                 body_format=body_format)

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
def save_to_kb_property():
    """
    The shared 'save_to_kb' argument for the two page-reading tools.

    Returned fresh each time so a caller mutating one tool's schema cannot
    affect the other's.
    """
    return {
        "type": "boolean",
        "description": (
            "Save this page into the local knowledge base as Markdown, for "
            "the RAG index. Default false: reading a page does NOT save it. "
            "Set it to true ONLY when the user asks for the page to be kept "
            "- 'save this to the knowledge base', 'add that page to the KB', "
            "'keep this for later'. Do not set it while merely reading pages "
            "to answer a question; saving pages nobody asked for is what "
            "makes the knowledge base return irrelevant results later."
        ),
    }


def body_format_property():
    """
    The shared 'body_format' argument for the two page-reading tools.

    It exists for the retry: the server picks a sensible representation on its
    own, but when a macro's content is still missing the model needs a way to
    ask for the rendered page without an environment change on the endpoint.
    """
    return {
        "type": "string",
        "enum": list(BODY_FORMATS),
        "description": (
            "Which version of the page body to read. Leave this unset "
            "('auto'): the server reads the page source for an ordinary page "
            "and asks Confluence to RENDER the page when it uses a macro whose "
            "content is generated at display time (task report, page "
            "properties report, children list, Jira issues). Set 'view' to "
            "force the rendered page, 'export_view' if a macro is STILL empty "
            "under 'view' (page trees and some third-party macros only render "
            "on export), or 'storage' for the raw source with no macro output. "
            "If a page reads as having no tasks/rows/issues where the user "
            "expects some, retry with 'view', then 'export_view', before "
            "telling them the page is empty."
        ),
    }


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
                "Returns the title, space, version, URL and the page body as "
                "Markdown, with Confluence macros rendered: inline task lists "
                "become '- [ ]' / '- [x]' checkboxes, info/note/warning panels "
                "become quotes, expand macros are shown open, code macros "
                "become fenced blocks, and tables stay tables. Macros whose "
                "content Confluence generates when the page is displayed (task "
                "report, page properties report, children list, Jira issues) "
                "are fetched by asking Confluence to render the page - see "
                "'body_format'. Reading a page does not save it anywhere; pass "
                "'save_to_kb' only if the user asks for it to be kept."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "page_id": {
                        "type": "string",
                        "description": "The numeric Confluence content ID, e.g. '393217'.",
                    },
                    "save_to_kb": save_to_kb_property(),
                    "body_format": body_format_property(),
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
                    "save_to_kb": save_to_kb_property(),
                    "body_format": body_format_property(),
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
        return client.get_page(arguments.get("page_id"),
                               save_to_kb=arguments.get("save_to_kb"),
                               body_format=arguments.get("body_format"))

    if name == "confluence_get_page_by_title":
        return client.get_page_by_title(
            arguments.get("title"), arguments.get("space"),
            save_to_kb=arguments.get("save_to_kb"),
            body_format=arguments.get("body_format"),
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


def resolve_kb_dir():
    """
    The knowledge-base folder, from the environment.

    Precedence: CONFLUENCE_KB_DIR (a full path of its own, or one of the
    DISABLE_KEYWORDS to forbid saving), then EVA_KNOWLEDGE_DIR with this
    server's "confluence" sub-folder appended, then that same sub-folder of the
    EVA_KNOWLEDGE_DIR fallback in the config block. Returns None when saving is
    switched off.
    """
    own = env_str("CONFLUENCE_KB_DIR")
    if own:
        return None if own.lower() in DISABLE_KEYWORDS else own
    root = env_str("EVA_KNOWLEDGE_DIR")
    if root:
        return None if root.lower() in DISABLE_KEYWORDS else os.path.join(root, SUBFOLDER)
    return os.path.join(EVA_KNOWLEDGE_DIR, SUBFOLDER)


def build_arg_parser():
    """
    The command line carries no configuration - every setting is an
    environment variable (see the CONFIGURATION section of the docstring), so
    two settings can never disagree and no token can end up in a process
    listing. Only --check and --version are flags.
    """
    p = argparse.ArgumentParser(
        description="MCP server for querying Confluence Data Center (stdio "
                    "transport). Configured entirely by environment variables: "
                    "CONFLUENCE_BASE_URL, CONFLUENCE_TOKEN, and "
                    "EVA_KNOWLEDGE_DIR (this server saves into its "
                    "'confluence' sub-folder). See the CONFIGURATION section "
                    "of this file's docstring.",
    )
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
        log("KB save folder   : {} ({})".format(
            servers.default.kb_dir,
            "every page read (autosave on)" if servers.default.kb_autosave
            else "on request only"))
    else:
        log("KB save folder   : disabled - no page can be saved")
    log("Page body format : {}".format(servers.default.body_format))
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

    base_url = env_str("CONFLUENCE_BASE_URL")
    base_url_2 = env_str("CONFLUENCE_BASE_URL_2")
    name_set = env_str("CONFLUENCE_NAME")
    name_2_set = env_str("CONFLUENCE_NAME_2")
    name = name_set or DEFAULT_NAME_1
    name_2 = name_2_set or DEFAULT_NAME_2
    ca_cert = env_str("CONFLUENCE_CA_CERT")
    ca_cert_2 = env_str("CONFLUENCE_CA_CERT_2")
    insecure = not env_bool("CONFLUENCE_VERIFY_SSL", True)
    timeout = env_int("CONFLUENCE_TIMEOUT", 30)
    max_body = env_int("CONFLUENCE_MAX_BODY", 0)
    kb_autosave = env_bool("CONFLUENCE_KB_AUTOSAVE", KB_AUTOSAVE)
    # Which body representation to read (see BODY_FORMAT). A typo must not stop
    # the server, but it must not be silently honoured either: warn and use the
    # default, because the wrong body means missing macro content.
    body_format_env = env_str("CONFLUENCE_BODY_FORMAT")
    body_format = normalise_body_format(body_format_env, None)
    if body_format is None:
        if body_format_env:
            log("WARNING: CONFLUENCE_BODY_FORMAT={!r} is not one of {}; using "
                "{!r}.".format(body_format_env, ", ".join(BODY_FORMATS),
                               BODY_FORMAT))
        body_format = BODY_FORMAT

    # The knowledge-base folder comes from the suite-wide EVA_KNOWLEDGE_DIR
    # (sub-folder "confluence"), or CONFLUENCE_KB_DIR for a path of its own;
    # "off" is how saving is switched off from a client that can only pass
    # strings, since a blank value means "not configured".
    kb_dir = resolve_kb_dir()

    if not base_url:
        log("FATAL: no base URL. Set CONFLUENCE_BASE_URL. (The first server is "
            "always the default one; a second server is configured on top of "
            "it with CONFLUENCE_BASE_URL_2.)")
        return 2
    if not token and not (user and password):
        log("FATAL: no credentials. Set the CONFLUENCE_TOKEN environment "
            "variable, or CONFLUENCE_USER and CONFLUENCE_PASSWORD.")
        return 2

    # A second server is enabled by its base URL alone. If the rest of its
    # settings are present without it, say so - silently ignoring them would
    # send "on Blue" questions to the first server and answer them with the
    # wrong wiki's content.
    if not base_url_2 and (token_2 or user_2 or password_2 or name_2_set):
        log("WARNING: second-server settings are present but "
            "CONFLUENCE_BASE_URL_2 is not set, so only one server is "
            "configured. Set CONFLUENCE_BASE_URL_2 to enable the second one.")

    # spec = (name, base_url, token, user, password, verify_ssl, ca_cert)
    specs = [(name, base_url, token, user, password, not insecure, ca_cert)]

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
        # TLS settings for the second server: its own env var, else whatever the
        # first server resolved to (usually the same internal CA).
        verify_env_2 = env_bool_opt("CONFLUENCE_VERIFY_SSL_2")
        insecure_2 = (not verify_env_2) if verify_env_2 is not None else insecure
        specs.append((name_2, base_url_2, token_2, user_2, password_2,
                      not insecure_2, ca_cert_2 or ca_cert))

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
                timeout=timeout,
                max_body=max_body,
                kb_dir=kb_dir,
                kb_autosave=kb_autosave,
                body_format=body_format,
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
    if not servers.default.kb_dir:
        log("knowledge-base saving disabled (CONFLUENCE_KB_DIR=off); pages are "
            "never written to disk")
        if servers.default.kb_autosave:
            log("WARNING: CONFLUENCE_KB_AUTOSAVE has no effect while the "
                "knowledge-base folder is off. Unset CONFLUENCE_KB_DIR (or "
                "point it at a folder) to enable saving.")
    elif servers.default.kb_autosave:
        log("knowledge-base AUTOSAVE is on: every page read is saved -> {}"
            .format(servers.default.kb_dir))
    else:
        log("knowledge-base saving on request only (save_to_kb=true) -> {}"
            .format(servers.default.kb_dir))
    if body_format == "storage":
        log("page bodies: storage only - macros that Confluence renders at "
            "display time (task reports, page properties, children lists) will "
            "read as empty. Unset CONFLUENCE_BODY_FORMAT to restore 'auto'.")
    else:
        log("page bodies: {} (macro content included)".format(body_format))
    if servers.multi:
        for position, client in enumerate(servers.clients, start=1):
            log("server {}: {} -> {}{}".format(
                position, client.name, client.base_url,
                "  (default)" if position == 1 else ""))
        if not name_set or not name_2_set:
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
