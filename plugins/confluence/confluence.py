#!/usr/bin/env python3
"""
confluence.py (v7.2.0) - A single-file MCP (Model Context Protocol) server
for querying - and, when switched on, writing to - ONE OR TWO Confluence Data
Center instances (tested against the 9.x v1 REST API) using only the Python 3
standard library.

It speaks MCP over stdio (newline-delimited JSON-RPC 2.0), the transport an
MCP client launches for a `type: stdio` server. No third-party packages are
required.

Tools exposed (read / query):
  - confluence_search           : free-text search for pages
  - confluence_search_cql       : advanced search using raw CQL
  - confluence_get_page         : fetch one page by numeric ID (with body text)
  - confluence_get_page_by_title: fetch one page by exact title + space key
  - confluence_list_pages_under : list pages beneath a parent page
  - confluence_list_tables      : the tables stored on a page - including ones
                                  inside table filter, column or section
                                  macros - with their columns and rows
  - confluence_list_attachments : the files attached to a page
  - confluence_download_attachment
                                : save one attachment into the local folder
                                  the matching plugin reads (see ATTACHMENTS)

Tools exposed ONLY when CONFLUENCE_ALLOW_WRITE=true (see WRITING PAGES):
  - confluence_create_page      : create a page (optionally under a parent)
  - confluence_update_table     : change cells of one table IN PLACE (match a
                                  row by a column value, set other columns),
                                  and/or add rows, keeping its formatting and
                                  every macro around it
  - confluence_update_section   : replace (or add to) ONE section - the
                                  content under a heading, or inside a titled
                                  panel/expand - leaving the rest untouched
  - confluence_append_to_page   : add content to the end/start of a page,
                                  keeping everything already on it
  - confluence_update_page      : replace a page's WHOLE content and/or title

WRITING PAGES
-------------
Off unless CONFLUENCE_ALLOW_WRITE=true. With it off the three write tools are
not offered at all, and the server never sends anything but a GET - exactly
how it behaved before v7.1.0. With it on, the tools take MARKDOWN and convert
it to Confluence's storage format: headings, bold/italic/strikethrough, inline
code, links, images (a bare file name is taken to be an attachment on the
page), nested lists, tables, block quotes and rules, plus
  - '- [ ] item' / '- [x] item'  -> real Confluence inline tasks
  - ```lang fenced code          -> the code macro
  - '> [!NOTE] Title' block      -> an info panel ([!TIP] tip, [!IMPORTANT]
                                    note, [!WARNING] / [!CAUTION] warning)
content_format="storage" passes raw storage-format XHTML through instead, for
a macro Markdown cannot express.

Four ways to change an existing page, from safest to bluntest:
  - confluence_update_table changes the TEXT inside the named cells of one
    table and nothing else, so a table inside a table filter, column,
    section or expand macro keeps the macro, and the table keeps its widths,
    colours and merged cells. Columns are named by the header row; a two-row
    header ("FY26" over "Q3" and "Q4") gives "FY26 Q3", and "Q3" alone works
    when unambiguous. Rows are found by a column's value (numbers compare as
    numbers, so "1,200" matches 1200). Every change is validated before any is
    made: an ambiguous row, an unknown column, a merged cell or a cell holding
    a macro stops the whole call and nothing is saved. New rows copy the last
    row's cells, and so its formatting. A table Confluence builds at display
    time (Jira issues, a page properties report, a CSV macro) has no stored
    cells and cannot be edited this way.
  - confluence_update_section finds ONE part of the page by its heading text
    (everything under it, sub-headings included, up to the next heading of the
    same or a higher level, or the end of its layout column) or by a panel's
    title (its body), and changes only that. The rest of the page goes back
    byte-for-byte, so "update the Director's notes on page 1234" cannot touch
    a task report or Jira table elsewhere on it. Replacing a section that
    itself contains a macro is refused unless allow_macro_removal=true.
  - confluence_append_to_page adds to the end (or start): the existing body
    is sent back byte-for-byte and only the new part is converted.
  - confluence_update_page REPLACES the whole body, so a macro left out of the
    new body is gone from the page (Confluence keeps the old version in the
    page history). For a full rewrite or a title change only.
Every write saves a new page version; a version that moved on in the meantime
is refused by Confluence (HTTP 409), and update_page's expected_version
argument refuses it before anything is sent.

ATTACHMENTS
-----------
confluence_download_attachment saves a file into the folder the plugin that
opens that type reads, because each document plugin is confined to ONE folder
and a download anywhere else could never be opened:

  .xlsx .xlsm  -> %EVA_DOCUMENTS_DIR%\excel       (excel plugin)
  .docx        -> %EVA_DOCUMENTS_DIR%\word        (word plugin)
  .pptx        -> %EVA_DOCUMENTS_DIR%\powerpoint  (powerpoint plugin)
  .pdf         -> %EVA_DOCUMENTS_DIR%\pdf         (pdf-to-md plugin)
  .md          -> the knowledge folder below       (knowledge-base plugin)

So "get the budget spreadsheet from page X, then use excel to total column D"
is two calls: this tool, then excel_read_range on the file name it reports.
Other types are listed but not downloaded. A file already in the folder is
never replaced unless the call passes overwrite=true. Downloading writes only
to the local disk; it needs no write access to Confluence.

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

Three variables are shared with every other plugin in this suite, set once
for your Windows account:

  EVA_PYTHON            full path to the python.exe the MCP client launches,
                        e.g. C:\Python311\python.exe (read by the plugin
                        manifest, not by this file)
  EVA_KNOWLEDGE_DIR     root of the RAG corpus (default H:\Eva\knowledge).
                        This server saves into its own "confluence"
                        sub-folder (created on demand):
                        %EVA_KNOWLEDGE_DIR%\confluence
  EVA_DOCUMENTS_DIR     root of the document library (default
                        H:\Eva\documents). Attachments are downloaded into
                        its per-type sub-folders (see ATTACHMENTS). The root
                        must exist or document downloads are disabled; a
                        sub-folder is created on demand.

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
  CONFLUENCE_ALLOW_WRITE
                        "true" to offer the page-writing tools (create,
                        update, append). Default off: read-only. Applies to
                        both servers; Confluence's own permissions still
                        decide which spaces the account can edit.
  CONFLUENCE_DOCS_DIR   override the documents ROOT attachments are
                        downloaded into, instead of EVA_DOCUMENTS_DIR. The
                        per-type sub-folders (excel, word, powerpoint, pdf)
                        are still appended, so it should be the same root the
                        document plugins use. "off" forbids document downloads
                        (Markdown still goes to the knowledge folder).

The second server is all-or-nothing: if CONFLUENCE_BASE_URL_2 is set without
credentials, the server refuses to start rather than quietly answering "Blue"
questions from the first instance.

INSTALLING INTO CLAUDE CODE
---------------------------
This server ships as the "confluence" Claude Code plugin (its manifest is
.claude-plugin/plugin.json next to this file), so the normal install is:

    /plugin marketplace add C:\path\to\claude-skills
    /plugin install confluence@mcnamee-claude-skills

The plugin prompts for nothing at install. Every setting above is an
environment variable the server inherits from Claude Code, so set them in the
`env` block of Claude Code's settings.json - %USERPROFILE%\.claude\settings.json,
or H:\Eva\.claude\settings.local.json to keep them on H: beside the working
folder - then fully restart Claude Code:

    {
      "env": {
        "EVA_PYTHON":            "C:\\Python311\\python.exe",
        "EVA_KNOWLEDGE_DIR":     "H:\\Eva\\knowledge",
        "CONFLUENCE_BASE_URL":   "https://confluence.internal.example.com",
        "CONFLUENCE_TOKEN":      "token-for-the-first-server",
        "CONFLUENCE_NAME_2":     "Blue",
        "CONFLUENCE_BASE_URL_2": "https://blue.confluence.example.com",
        "CONFLUENCE_TOKEN_2":    "token-for-the-second-server"
      }
    }

(Backslashes are doubled because settings.json is JSON.) A blank or missing
value means "not set", so leave out anything you do not use. Windows user
environment variables work too, as the alternative:

    setx EVA_PYTHON            "C:\Python311\python.exe"
    setx EVA_KNOWLEDGE_DIR     "H:\Eva\knowledge"
    setx CONFLUENCE_BASE_URL   "https://confluence.internal.example.com"
    setx CONFLUENCE_TOKEN      "token-for-the-first-server"
    setx CONFLUENCE_TOKEN_2    "token-for-the-second-server"

(`setx` does not affect processes that are already running - quit VS Code /
Claude Code completely and reopen it.) See README.md next to this file for the
full settings reference.

TESTING
-------
Diagnostic output goes ONLY to stderr. stdout is reserved for the JSON-RPC
stream - writing anything else there would corrupt the protocol.

`--check` connects to EVERY configured server, prints who you are
authenticated as, how many spaces are visible, the download folder and
whether page writing is on (to stderr), then exits
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
__version__ = "7.2.0"

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
# full path of its own. The literal here is what a stock H:\Eva install
# resolves to; it MUST stay inside the knowledge-base plugin's corpus
# (H:\Eva\knowledge) or the saved pages would never be indexed. The folder is
# created on demand.
# Set CONFLUENCE_KB_DIR=off to forbid saving altogether, after which this
# server touches no local file at all.
SUBFOLDER = "confluence"                 # this server's knowledge sub-folder
EVA_KNOWLEDGE_DIR = r"H:\Eva\knowledge"  # fallback for the suite-wide root
KB_DIR = r"H:\Eva\knowledge\confluence"

# Root of the document library that confluence_download_attachment saves into.
# RESOLVED FROM THE ENVIRONMENT in main(): CONFLUENCE_DOCS_DIR, else
# %EVA_DOCUMENTS_DIR%, else this fallback. It is a ROOT, not one folder: a
# download goes into the sub-folder named for its FILE TYPE (documents\excel,
# documents\word, documents\powerpoint, documents\pdf - see ATTACHMENT_FOLDERS),
# because each is the one folder the plugin for that type can open. The root
# must already exist; the type sub-folder is created on demand. Set
# CONFLUENCE_DOCS_DIR=off to forbid document downloads.
EVA_DOCUMENTS_DIR = r"H:\Eva\documents"  # fallback for the suite-wide root

# Whether the page-writing tools (create / update / append) are offered at
# all. Off unless CONFLUENCE_ALLOW_WRITE=true: this server was read-only for
# its first seven major versions, and an endpoint that only ever read
# Confluence must not start editing it because the plugin updated.
ALLOW_WRITE = False

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


# ---------------------------------------------------------------------------
# Markdown -> Confluence storage format (for the write tools)
# ---------------------------------------------------------------------------
# The write tools take Markdown, because that is what the model drafts in, and
# convert it to Confluence's storage format (XHTML with <ac:...> macros). The
# converter covers the Markdown a page is actually written in - headings,
# paragraphs, bold/italic/strikethrough, inline code, links, images, nested
# bullet/numbered lists, task lists, tables, block quotes, rules and fenced code
# - and maps a few constructs onto the Confluence macro a reader expects:
#
#   ```lang ... ```     -> code macro (language kept when Confluence knows it)
#   - [ ] / - [x]       -> inline tasks (a real Confluence task list)
#   > [!NOTE] Title     -> info panel   ([!TIP] tip, [!IMPORTANT] note,
#                                        [!WARNING] / [!CAUTION] warning)
#
# Anything else is escaped and kept as text, so a construct the converter does
# not know can never produce markup Confluence rejects. For a page that needs
# a macro Markdown cannot express, the write tools also accept raw storage
# format (content_format="storage").

# Languages the Confluence code macro accepts, plus the aliases people type. A
# language outside this list is dropped rather than passed on, because an
# unknown language renders as an error box instead of the code.
_CODE_LANGUAGES = {
    "actionscript3": "actionscript3", "applescript": "applescript",
    "bash": "bash", "sh": "bash", "shell": "bash", "zsh": "bash",
    "csharp": "csharp", "cs": "csharp", "c#": "csharp",
    "coldfusion": "coldfusion", "cpp": "cpp", "c++": "cpp", "c": "cpp",
    "css": "css", "delphi": "delphi", "pascal": "delphi", "diff": "diff",
    "patch": "diff", "erlang": "erlang", "groovy": "groovy",
    "java": "java", "javafx": "javafx", "javascript": "javascript",
    "js": "javascript", "json": "javascript", "typescript": "javascript",
    "ts": "javascript", "perl": "perl", "php": "php",
    "powershell": "powershell", "ps1": "powershell", "pwsh": "powershell",
    "python": "python", "py": "python", "ruby": "ruby", "rb": "ruby",
    "sass": "sass", "scss": "sass", "scala": "scala", "sql": "sql",
    "vb": "vb", "vbnet": "vb", "vba": "vb", "xml": "xml", "html": "xml",
    "xhtml": "xml", "yaml": "yaml", "yml": "yaml", "text": "none",
    "txt": "none", "plain": "none", "none": "none",
}

# GitHub-style alert markers in a block quote -> Confluence panel macro.
_ALERT_MACROS = {
    "NOTE": "info", "INFO": "info", "TIP": "tip", "IMPORTANT": "note",
    "WARNING": "warning", "CAUTION": "warning",
}

_FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})\s*([^\s`]*)\s*$")
_HEADING_RE = re.compile(r"^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$")
_HR_RE = re.compile(r"^\s{0,3}([-*_])(\s*\1){2,}\s*$")
_LIST_RE = re.compile(r"^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$")
_TASK_RE = re.compile(r"^\[([ xX])\]\s+(.*)$")
_TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?\s*$")
_ALERT_RE = re.compile(r"^\[!([A-Za-z]+)\]\s*(.*)$")
_TASK_ID_RE = re.compile(r"<ac:task-id>\s*(\d+)\s*</ac:task-id>")


def _xml_text(text):
    """Escape text for XHTML element content."""
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def _xml_attr(text):
    """Escape text for a double-quoted XHTML attribute."""
    return _xml_text(text).replace('"', "&quot;")


def _cdata(text):
    """Wrap text in CDATA, splitting any ']]>' so it cannot end the section."""
    return "<![CDATA[" + text.replace("]]>", "]]]]><![CDATA[>") + "]]>"


def _inline_markdown(text):
    """
    Convert one run of inline Markdown to storage-format XHTML.

    Code spans are lifted out first, so nothing inside backticks is treated as
    emphasis or a link; everything else is escaped BEFORE any markup is added,
    so a literal '<' or '&' in the text can never become a tag.
    """
    stash = []

    def keep(markup):
        stash.append(markup)
        return "\x00{}\x00".format(len(stash) - 1)

    # Inline code spans (`x` or ``x``).
    text = re.sub(r"(`+)(.+?)\1",
                  lambda m: keep("<code>{}</code>".format(_xml_text(m.group(2).strip()))),
                  text)
    # Backslash escapes: \* stays a literal asterisk, and so on.
    text = re.sub(r"\\([\\`*_{}\[\]()#+\-.!|~>])",
                  lambda m: keep(_xml_text(m.group(1))), text)
    # Images: an http(s) URL is an external image, anything else is taken to be
    # the filename of an attachment on the page.
    def image(m):
        alt, src = m.group(1), m.group(2).strip()
        if re.match(r"(?i)^https?://", src):
            ref = '<ri:url ri:value="{}" />'.format(_xml_attr(src))
        else:
            ref = '<ri:attachment ri:filename="{}" />'.format(_xml_attr(src))
        alt_attr = ' ac:alt="{}"'.format(_xml_attr(alt)) if alt else ""
        return keep("<ac:image{}>{}</ac:image>".format(alt_attr, ref))
    text = re.sub(r"!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", image, text)
    # Links: [label](url) and <https://autolink>. Only the tags are lifted out:
    # the label stays in the text, so it is escaped and emphasised with the
    # rest of the line.
    def link(m):
        label, href = m.group(1), m.group(2).strip()
        return (keep('<a href="{}">'.format(_xml_attr(href))) + label
                + keep("</a>"))
    text = re.sub(r"\[([^\]]+)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)", link, text)
    text = re.sub(r"<(https?://[^\s>]+)>",
                  lambda m: keep('<a href="{0}">{1}</a>'.format(
                      _xml_attr(m.group(1)), _xml_text(m.group(1)))), text)

    # Everything left is plain text: escape it, then add emphasis.
    text = _xml_text(text)
    text = re.sub(r"\*\*(?=\S)(.+?)(?<=\S)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<!\w)__(?=\S)(.+?)(?<=\S)__(?!\w)", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<![*\w])\*(?=[^\s*])(.+?)(?<=[^\s*])\*(?![*\w])", r"<em>\1</em>", text)
    text = re.sub(r"(?<!\w)_(?=[^\s_])(.+?)(?<=[^\s_])_(?!\w)", r"<em>\1</em>", text)
    text = re.sub(r"~~(?=\S)(.+?)(?<=\S)~~",
                  r'<span style="text-decoration: line-through;">\1</span>', text)

    # Put the lifted-out pieces back. Stashed markup never contains a marker of
    # its own, so one pass restores everything.
    return re.sub(r"\x00(\d+)\x00", lambda m: stash[int(m.group(1))], text)


def _paragraph_markup(lines):
    """Join a paragraph's lines, honouring Markdown hard breaks."""
    parts = []
    for index, line in enumerate(lines):
        hard = line.endswith("  ") or line.endswith("\\")
        clean = line.rstrip()
        if clean.endswith("\\"):
            clean = clean[:-1].rstrip()
        parts.append(_inline_markdown(clean.strip()))
        if index < len(lines) - 1:
            parts.append("<br />" if hard else " ")
    return "".join(parts)


def _split_table_row(line):
    """Split a Markdown table row on unescaped pipes."""
    row = line.strip()
    if row.startswith("|"):
        row = row[1:]
    if row.endswith("|") and not row.endswith("\\|"):
        row = row[:-1]
    cells = re.split(r"(?<!\\)\|", row)
    return [cell.strip().replace("\\|", "|") for cell in cells]


class _TaskCounter:
    """Hands out Confluence task IDs, continuing from any already on the page."""

    def __init__(self, start=1):
        self.next_id = max(1, int(start))

    def take(self):
        value = self.next_id
        self.next_id += 1
        return value


def _render_list(items, tasks):
    """
    Render a tree of list items (see _parse_list_block) as storage XHTML.

    Consecutive siblings of the same kind share one list element, so a bullet
    list followed directly by a numbered one becomes two lists, as it reads.
    """
    out = []
    index = 0
    while index < len(items):
        kind = items[index]["kind"]
        group = []
        while index < len(items) and items[index]["kind"] == kind:
            group.append(items[index])
            index += 1
        if kind == "task":
            out.append("<ac:task-list>")
            for item in group:
                body = _inline_markdown(item["text"])
                if item["children"]:
                    body += _render_list(item["children"], tasks)
                out.append(
                    "<ac:task><ac:task-id>{}</ac:task-id>"
                    "<ac:task-status>{}</ac:task-status>"
                    "<ac:task-body>{}</ac:task-body></ac:task>".format(
                        tasks.take(),
                        "complete" if item["checked"] else "incomplete",
                        body))
            out.append("</ac:task-list>")
        else:
            out.append("<{}>".format(kind))
            for item in group:
                body = _inline_markdown(item["text"])
                if item["children"]:
                    body += _render_list(item["children"], tasks)
                out.append("<li>{}</li>".format(body))
            out.append("</{}>".format(kind))
    return "".join(out)


def _indent_width(prefix):
    """Width of leading whitespace, a tab counting as four spaces."""
    return len(prefix.replace("\t", "    "))


def _parse_list_block(lines):
    """
    Turn the lines of one Markdown list into a tree of items.

    Each item is {"kind": "ul"|"ol"|"task", "text", "checked", "indent",
    "children"}. Nesting follows indentation: an item indented further than the
    one above it becomes its child. A non-marker line is a continuation of the
    item above it.
    """
    roots = []
    stack = []
    for line in lines:
        m = _LIST_RE.match(line)
        if not m:
            if stack:
                stack[-1]["text"] += " " + line.strip()
            continue
        indent = _indent_width(m.group(1))
        marker, text = m.group(2), m.group(3)
        kind = "ol" if marker[0].isdigit() else "ul"
        checked = False
        task = _TASK_RE.match(text)
        if task and kind == "ul":
            kind = "task"
            checked = task.group(1).lower() == "x"
            text = task.group(2)
        item = {"kind": kind, "text": text, "checked": checked,
                "indent": indent, "children": []}
        while stack and stack[-1]["indent"] >= indent:
            stack.pop()
        (stack[-1]["children"] if stack else roots).append(item)
        stack.append(item)
    return roots


def markdown_to_storage(markdown, task_start=1):
    """
    Convert Markdown to Confluence storage-format XHTML.

    task_start is the first inline-task ID to use; when appending to a page
    that already has tasks, pass one past the highest ID on it so the new tasks
    do not collide with the old ones.
    """
    tasks = _TaskCounter(task_start)
    return _blocks_to_storage(
        (markdown or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"),
        tasks)


def _is_block_start(line, next_line):
    """True when a line starts a block other than a paragraph."""
    return bool(
        _FENCE_RE.match(line) or _HEADING_RE.match(line) or _HR_RE.match(line)
        or _LIST_RE.match(line) or line.lstrip().startswith(">")
        or (line.strip().startswith("|") and next_line is not None
            and _TABLE_SEP_RE.match(next_line))
    )


def _blocks_to_storage(lines, tasks):
    out = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        if not line.strip():
            i += 1
            continue

        fence = _FENCE_RE.match(line)
        if fence:
            marker = fence.group(1)
            lang = _CODE_LANGUAGES.get(fence.group(2).lower().strip())
            body = []
            i += 1
            while i < n and not lines[i].strip().startswith(marker):
                body.append(lines[i])
                i += 1
            i += 1  # skip the closing fence (or run off the end)
            param = ('<ac:parameter ac:name="language">{}</ac:parameter>'
                     .format(lang) if lang else "")
            out.append(
                '<ac:structured-macro ac:name="code">{}'
                "<ac:plain-text-body>{}</ac:plain-text-body>"
                "</ac:structured-macro>".format(param, _cdata("\n".join(body))))
            continue

        heading = _HEADING_RE.match(line)
        if heading:
            level = len(heading.group(1))
            out.append("<h{0}>{1}</h{0}>".format(
                level, _inline_markdown(heading.group(2))))
            i += 1
            continue

        if _HR_RE.match(line):
            out.append("<hr />")
            i += 1
            continue

        next_line = lines[i + 1] if i + 1 < n else None
        if (line.strip().startswith("|") and next_line is not None
                and _TABLE_SEP_RE.match(next_line)):
            header = _split_table_row(line)
            i += 2
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                rows.append(_split_table_row(lines[i]))
                i += 1
            width = len(header)
            cells = ["<tr>" + "".join("<th>{}</th>".format(_inline_markdown(c))
                                      for c in header) + "</tr>"]
            for row in rows:
                row = (row + [""] * width)[:width]
                cells.append("<tr>" + "".join(
                    "<td>{}</td>".format(_inline_markdown(c)) for c in row) + "</tr>")
            out.append("<table><tbody>{}</tbody></table>".format("".join(cells)))
            continue

        if line.lstrip().startswith(">"):
            quoted = []
            while i < n and lines[i].lstrip().startswith(">"):
                inner = lines[i].lstrip()[1:]
                quoted.append(inner[1:] if inner.startswith(" ") else inner)
                i += 1
            alert = _ALERT_RE.match(quoted[0].strip()) if quoted else None
            macro = _ALERT_MACROS.get(alert.group(1).upper()) if alert else None
            if macro:
                title = alert.group(2).strip()
                body = _blocks_to_storage(quoted[1:], tasks)
                title_param = ('<ac:parameter ac:name="title">{}</ac:parameter>'
                               .format(_xml_text(title)) if title else "")
                out.append(
                    '<ac:structured-macro ac:name="{}">{}'
                    "<ac:rich-text-body>{}</ac:rich-text-body>"
                    "</ac:structured-macro>".format(macro, title_param, body))
            else:
                out.append("<blockquote>{}</blockquote>".format(
                    _blocks_to_storage(quoted, tasks)))
            continue

        if _LIST_RE.match(line):
            block = []
            while i < n:
                current = lines[i]
                if not current.strip():
                    # A blank line ends the list unless the list carries on
                    # (another item, or an indented continuation) after it.
                    ahead = i + 1
                    while ahead < n and not lines[ahead].strip():
                        ahead += 1
                    if ahead < n and (_LIST_RE.match(lines[ahead])
                                      or lines[ahead].startswith((" ", "\t"))):
                        i = ahead
                        continue
                    break
                if (block and not _LIST_RE.match(current)
                        and not current.startswith((" ", "\t"))
                        and _is_block_start(current, lines[i + 1] if i + 1 < n else None)):
                    break
                block.append(current)
                i += 1
            out.append(_render_list(_parse_list_block(block), tasks))
            continue

        # Paragraph: consecutive lines up to a blank line or another block.
        para = [line]
        i += 1
        while i < n and lines[i].strip() and not _is_block_start(
                lines[i], lines[i + 1] if i + 1 < n else None):
            para.append(lines[i])
            i += 1
        out.append("<p>{}</p>".format(_paragraph_markup(para)))
    return "".join(out)


def next_task_id(storage):
    """One past the highest inline-task ID in a storage body (1 if none)."""
    ids = [int(x) for x in _TASK_ID_RE.findall(storage or "")]
    return (max(ids) + 1) if ids else 1


# ---------------------------------------------------------------------------
# Section editing: find one part of a storage body by its heading or title
# ---------------------------------------------------------------------------
# confluence_update_section changes ONE part of a page and sends everything
# else back byte-for-byte, so the macros, layouts and tables elsewhere on the
# page cannot be lost. A "section" is either
#   - a HEADING and everything after it up to the next heading of the same or
#     a higher level (or the end of the layout cell / page it sits in), or
#   - a PANEL-LIKE MACRO (info, note, panel, expand...) whose title parameter
#     matches: its rich-text body is the section.
# Locating it needs the exact character offsets of each element, so the
# storage body is indexed with the standard-library HTML parser rather than
# converted to Markdown (which would lose exactly what this protects).

_CDATA_RE = re.compile(r"<!\[CDATA\[.*?\]\]>", re.S)
_TAG_RE = re.compile(r"<[^>]+>")
_HEADING_TAGS = ("h1", "h2", "h3", "h4", "h5", "h6")


def _mask_cdata(raw):
    """
    Blank out CDATA sections (code-macro bodies) with same-length filler, so
    a '<' or '>' inside code cannot be mistaken for markup. Offsets are
    unchanged, so positions found in the masked text apply to the original.
    """
    return _CDATA_RE.sub(lambda m: "x" * len(m.group(0)), raw)


def _norm_label(text):
    """Normalise a heading/title for matching: tags, entities, quotes, case."""
    text = _TAG_RE.sub(" ", text or "")
    text = html.unescape(text)
    for curly, plain in (("\u2018", "'"), ("\u2019", "'"), ("\u201c", '"'),
                         ("\u201d", '"'), ("\u00a0", " ")):
        text = text.replace(curly, plain)
    text = " ".join(text.split()).strip().rstrip(":").strip()
    return text.casefold()


class _StorageIndex(html.parser.HTMLParser):
    """
    Records where every element of a storage body starts and ends, as
    character offsets: start (the '<' of the start tag), inner_start (just
    after it), inner_end (the '<' of the end tag) and end (just after it).
    """

    def __init__(self, raw):
        super().__init__(convert_charrefs=False)
        self.raw = raw
        self._line_starts = [0]
        for index, char in enumerate(raw):
            if char == "\n":
                self._line_starts.append(index + 1)
        self.elements = []
        self._stack = []

    def _offset(self):
        line, col = self.getpos()
        return self._line_starts[line - 1] + col

    def _add(self, tag, attrs, closed):
        start = self._offset()
        text = self.get_starttag_text() or ""
        inner = start + len(text)
        self.elements.append({
            "tag": tag, "attrs": {k.lower(): (v or "") for k, v in attrs},
            "start": start, "inner_start": inner,
            "inner_end": inner if closed else None,
            "end": inner if closed else None,
            "parent": self._stack[-1] if self._stack else None,
        })
        return len(self.elements) - 1

    def handle_starttag(self, tag, attrs):
        self._stack.append(self._add(tag, attrs, closed=False))

    def handle_startendtag(self, tag, attrs):
        self._add(tag, attrs, closed=True)

    def handle_endtag(self, tag):
        start = self._offset()
        end = self.raw.find(">", start) + 1 or len(self.raw)
        for depth in range(len(self._stack) - 1, -1, -1):
            element = self.elements[self._stack[depth]]
            if element["tag"] != tag:
                continue
            # Anything opened inside it and never closed ends here too.
            for index in self._stack[depth + 1:]:
                inner = self.elements[index]
                inner["inner_end"] = inner["end"] = start
            element["inner_end"], element["end"] = start, end
            del self._stack[depth:]
            return
        # A stray end tag with no matching start: ignore it.

    def finish(self):
        self.close()
        for index in self._stack:
            element = self.elements[index]
            element["inner_end"] = element["end"] = len(self.raw)
        self._stack = []
        return self.elements


def _index_storage(raw):
    parser = _StorageIndex(_mask_cdata(raw))
    parser.feed(parser.raw)
    return parser.finish()


def storage_sections(raw):
    """
    Every editable section of a storage body, in page order, as dicts:
    {kind, label, level, start, end} where [start, end) is the CONTENT to
    replace (not the heading or the macro wrapper, which are kept).
    """
    masked = _mask_cdata(raw)
    elements = _index_storage(raw)
    sections = []
    for index, el in enumerate(elements):
        if el["tag"] in _HEADING_TAGS:
            level = int(el["tag"][1])
            end = None
            for later in elements[index + 1:]:
                if later["start"] < el["end"]:
                    continue
                if (later["parent"] == el["parent"] and later["tag"] in _HEADING_TAGS
                        and int(later["tag"][1]) <= level):
                    end = later["start"]
                    break
            if end is None:
                parent = el["parent"]
                end = elements[parent]["inner_end"] if parent is not None else len(raw)
            sections.append({
                "kind": "heading", "level": level,
                "label": masked[el["inner_start"]:el["inner_end"]],
                "start": el["end"], "end": end,
            })
        elif el["tag"] in ("ac:structured-macro", "ac:macro"):
            children = [c for c in elements if c["parent"] == index]
            title = next((c for c in children if c["tag"] == "ac:parameter"
                          and c["attrs"].get("ac:name", "").lower() == "title"), None)
            body = next((c for c in children if c["tag"] == "ac:rich-text-body"), None)
            if title is None or body is None:
                continue
            sections.append({
                "kind": "{} macro".format(el["attrs"].get("ac:name", "?")),
                "level": None,
                "label": masked[title["inner_start"]:title["inner_end"]],
                "start": body["inner_start"], "end": body["inner_end"],
            })
    for section in sections:
        section["display"] = " ".join(
            html.unescape(_TAG_RE.sub(" ", section["label"])).split())
    return sections


def find_section(raw, wanted):
    """
    The one section whose heading/title matches `wanted` (exact after
    normalising, else a unique partial match). Raises ConfluenceError naming
    the sections that exist when there is no single match.
    """
    sections = storage_sections(raw)
    want = _norm_label(wanted)
    if not want:
        raise ConfluenceError("'section' (the heading or panel title) is required.")
    exact = [s for s in sections if _norm_label(s["label"]) == want]
    matches = exact or [s for s in sections if want in _norm_label(s["label"])]

    def listing(items):
        return "; ".join("{!r} ({})".format(s["display"], s["kind"] if s["level"] is None
                                            else "heading h{}".format(s["level"]))
                         for s in items) or "none"
    if not matches:
        raise ConfluenceError(
            "No heading or titled panel matching {!r} on this page. Sections: {}."
            .format(wanted, listing(sections)))
    if len(matches) > 1:
        raise ConfluenceError(
            "{!r} matches more than one section: {}. Use the full heading text."
            .format(wanted, listing(matches)))
    return matches[0]


# ---------------------------------------------------------------------------
# Table editing: change cells in place, wherever the table sits
# ---------------------------------------------------------------------------
# A table is often wrapped in a macro - a table filter, a column or section
# layout, an expand - and replacing the section around it would either delete
# that macro or rebuild the table from Markdown, losing its widths, colours and
# merged cells. confluence_update_table instead rewrites the TEXT inside the
# cells it is asked to change and nothing else: every tag of the table, and
# every macro around it, is sent back as it was.
#
# Rows and columns are mapped onto a grid that honours colspan and rowspan, so
# "the Q3 column" still means the right cell in a table with merged headers.
# The first row is the header row; columns are named by its text.

# Inline wrappers peeled off a cell before its text is replaced, so a figure
# that was bold, coloured or in a paragraph stays that way.
_CELL_WRAPPERS = ("p", "strong", "b", "em", "i", "u", "span", "code", "sub", "sup", "s")
# Markup inside a cell that a text replacement would destroy.
_CELL_RICH_RE = re.compile(r"<(ac:structured-macro|ac:macro|ac:task-list|ac:image|ac:link|"
                           r"table|ul|ol|ri:)", re.I)


def _plain(fragment):
    """Visible text of a storage fragment, whitespace collapsed."""
    return " ".join(html.unescape(_TAG_RE.sub(" ", fragment or "")).split())


def _numeric(text):
    """A cell or match value as a number, if it reads as one ('$1,200' -> 1200)."""
    cleaned = str(text).strip().replace(",", "").replace("$", "").replace("%", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _cell_equal(cell_text, wanted):
    """Loose equality for matching a row: numbers as numbers, else text."""
    a, b = _numeric(cell_text), _numeric(wanted)
    if a is not None and b is not None:
        return a == b
    return _norm_label(cell_text) == _norm_label(str(wanted))


def storage_tables(raw):
    """
    Every table stored in a page body, in page order. Each is a dict:
      number    1-based position on the page
      start/end offsets of the <table> element
      heading   the nearest heading above it ('' if none)
      wrappers  names of the macros it sits inside, outermost first
      rows      the table's own rows (not a nested table's), each a dict
                {start, end, cells}, a cell being {tag, start, end, text,
                rich, colspan, rowspan, row}
      grid      rows x columns of cell references (None for a gap), with a
                spanning cell repeated in every position it covers
    """
    masked = _mask_cdata(raw)
    elements = _index_storage(raw)
    children = [[] for _ in elements]
    for index, el in enumerate(elements):
        if el["parent"] is not None:
            children[el["parent"]].append(index)

    def own_rows(table_index):
        """The table's own <tr>s, in order - never those of a nested table."""
        found, todo = [], list(children[table_index])
        while todo:
            index = todo.pop(0)
            tag = elements[index]["tag"]
            if tag == "tr":
                found.append(index)
            elif tag != "table":
                todo[0:0] = children[index]    # depth-first, keeps page order
        return found

    def span(cell_el, name):
        try:
            return max(1, int(cell_el["attrs"].get(name) or 1))
        except ValueError:
            return 1

    tables = []
    heading_text = ""
    for index, el in enumerate(elements):
        if el["tag"] in _HEADING_TAGS:
            heading_text = _plain(masked[el["inner_start"]:el["inner_end"]])
            continue
        if el["tag"] != "table":
            continue
        wrappers = []
        parent = el["parent"]
        while parent is not None:
            if elements[parent]["tag"] in ("ac:structured-macro", "ac:macro"):
                wrappers.append(elements[parent]["attrs"].get("ac:name", "?"))
            elif elements[parent]["tag"] == "table":
                outer = next((t["number"] for t in tables
                              if t["start"] == elements[parent]["start"]), "?")
                wrappers.append("a cell of table {}".format(outer))
            parent = elements[parent]["parent"]
        rows = []
        for r_index in own_rows(index):
            row_el = elements[r_index]
            cells = []
            for c_index in children[r_index]:
                cell_el = elements[c_index]
                if cell_el["tag"] not in ("td", "th"):
                    continue
                inner = masked[cell_el["inner_start"]:cell_el["inner_end"]]
                cells.append({
                    "tag": cell_el["tag"],
                    "start": cell_el["inner_start"], "end": cell_el["inner_end"],
                    "text": _plain(inner),
                    "rich": bool(_CELL_RICH_RE.search(inner)),
                    "colspan": span(cell_el, "colspan"),
                    "rowspan": span(cell_el, "rowspan"),
                    "row": len(rows),
                })
            rows.append({"start": row_el["start"], "end": row_el["end"], "cells": cells})

        # Lay the cells onto a grid, honouring colspan/rowspan.
        grid = []
        for r, row in enumerate(rows):
            while len(grid) <= r:
                grid.append([])
            col = 0
            for cell in row["cells"]:
                while col < len(grid[r]) and grid[r][col] is not None:
                    col += 1
                for dr in range(cell["rowspan"]):
                    while len(grid) <= r + dr:
                        grid.append([])
                    line = grid[r + dr]
                    while len(line) < col + cell["colspan"]:
                        line.append(None)
                    for dc in range(cell["colspan"]):
                        line[col + dc] = cell
                col += cell["colspan"]
        grid = grid[:len(rows)]
        width = max([len(line) for line in grid] or [0])
        grid = [line + [None] * (width - len(line)) for line in grid]

        # Header rows: the leading rows made only of <th> cells (at least the
        # first row, which is the header even in a table styled without <th>).
        # A two-row header - "FY26" spanning "Q3" and "Q4" - names its columns
        # "FY26 Q3" and "FY26 Q4", and either part can be used on its own.
        header_rows = 1
        while (header_rows < len(grid) and grid[header_rows]
               and all(cell is None or cell["tag"] == "th" for cell in grid[header_rows])
               and any(cell is not None for cell in grid[header_rows])):
            header_rows += 1
        if len(grid) <= 1:
            header_rows = len(grid)
        columns = []
        for col in range(width):
            parts = []
            for r in range(header_rows):
                cell = grid[r][col]
                if cell is not None and cell["text"] and cell["text"] not in parts:
                    parts.append(cell["text"])
            columns.append({"name": " ".join(parts), "parts": parts})
        tables.append({
            "number": len(tables) + 1, "start": el["start"], "end": el["end"],
            "heading": heading_text, "wrappers": list(reversed(wrappers)),
            "rows": rows, "grid": grid, "width": width,
            "header_rows": header_rows, "columns": columns,
        })
    return tables


def table_headers(table):
    """Column names (header rows combined, e.g. 'FY26 Q3')."""
    return [col["name"] for col in table.get("columns") or []]


def find_table(tables, wanted):
    """
    A table by number, by the heading it sits under, or - when the page has
    only one - by omission. Raises ConfluenceError listing the tables.
    """
    def listing():
        return "; ".join(
            "{} (under {!r}{}; columns: {})".format(
                t["number"], t["heading"] or "no heading",
                ", inside " + " > ".join(t["wrappers"]) if t["wrappers"] else "",
                ", ".join(h for h in table_headers(t) if h) or "?")
            for t in tables) or "none"
    if not tables:
        raise ConfluenceError(
            "This page has no stored tables. (A table produced by a macro at "
            "display time - Jira issues, a page properties report, a CSV macro - "
            "has no cells in the page to edit.)")
    text = str(wanted or "").strip()
    if not text:
        if len(tables) == 1:
            return tables[0]
        raise ConfluenceError("This page has {} tables - say which with 'table' "
                              "(its number or the heading above it): {}."
                              .format(len(tables), listing()))
    if text.isdigit():
        number = int(text)
        if 1 <= number <= len(tables):
            return tables[number - 1]
        raise ConfluenceError("There is no table {}. Tables: {}.".format(number, listing()))
    want = _norm_label(text)
    exact = [t for t in tables if _norm_label(t["heading"]) == want]
    matches = exact or [t for t in tables if want and want in _norm_label(t["heading"])]
    # A table nested in another table's cell shares its heading; the heading
    # means the outer one unless the nested table is the only match.
    outer = [t for t in matches
             if not any(w.startswith("a cell of table") for w in t["wrappers"])]
    if outer:
        matches = outer
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ConfluenceError("No table sits under a heading matching {!r}. Tables: {}."
                              .format(text, listing()))
    raise ConfluenceError("{} tables sit under {!r} - use the table number: {}."
                          .format(len(matches), text, listing()))


def _column_index(table, column):
    """
    A column by its full header name ('FY26 Q3'), by one unambiguous part of
    a multi-row header ('Q3'), or by 1-based number.
    """
    columns = table["columns"]
    text = str(column).strip()
    if text.isdigit() and 1 <= int(text) <= len(columns):
        return int(text) - 1
    want = _norm_label(text)
    full = [i for i, col in enumerate(columns) if _norm_label(col["name"]) == want]
    if len(full) == 1:
        return full[0]
    part = [i for i, col in enumerate(columns)
            if any(_norm_label(p) == want for p in col["parts"])]
    if len(part) == 1:
        return part[0]
    names = ", ".join("{} {!r}".format(i + 1, col["name"]) for i, col in enumerate(columns))
    if len(full) > 1 or len(part) > 1:
        raise ConfluenceError("Column {!r} is ambiguous - use the full name or its "
                              "number. Columns: {}.".format(column, names))
    raise ConfluenceError("No column {!r}. Columns: {}.".format(column, names))


def _replace_cell_text(inner, new_markup):
    """
    New content for a cell, keeping the single wrappers around its text: a
    cell holding <p><strong>100</strong></p> becomes <p><strong>120</strong></p>.
    """
    prefix, suffix, core = "", "", inner
    while True:
        m = re.match(r"^(\s*<({})(\s[^>]*)?>)(.*)(</\2>\s*)$".format(
            "|".join(_CELL_WRAPPERS)), core, re.S | re.I)
        if not m or re.search(r"</?{}[\s>]".format(m.group(2)), m.group(4), re.I):
            break
        prefix, suffix, core = prefix + m.group(1), m.group(5) + suffix, m.group(4)
    if not prefix:
        # A bare cell: Confluence's own editor puts cell text in a paragraph.
        return "<p>{}</p>".format(new_markup) if new_markup else ""
    return prefix + new_markup + suffix


def render_tables(tables, only=None, max_rows=50):
    """Tables as text: where each sits, its columns, and its rows."""
    out = []
    for t in tables:
        if only is not None and t is not only:
            continue
        head = "Table {} - under {!r}".format(t["number"], t["heading"] or "no heading")
        if t["wrappers"]:
            head += ", inside " + " > ".join(t["wrappers"])
        head += " - {} row(s) x {} column(s)".format(len(t["rows"]), t["width"])
        out.append(head)
        out.append("Columns: " + " | ".join(
            "{} {}".format(i + 1, col["name"] or "(blank)") for i, col in enumerate(t["columns"])))
        shown = t["grid"][:max_rows + t["header_rows"]]
        for r, line in enumerate(shown):
            cells = []
            for cell in line:
                if cell is None:
                    cells.append("")
                elif cell["row"] != r:
                    cells.append("^")          # covered by a cell spanning from above
                else:
                    cells.append(cell["text"] + (" [macro]" if cell["rich"] else ""))
            out.append("| " + " | ".join(c.replace("|", "\\|") for c in cells) + " |")
            if r == t["header_rows"] - 1:
                out.append("|" + " --- |" * len(cells))
        if len(t["grid"]) > len(shown):
            out.append("[{} more row(s); read this table alone with 'table': {}.]".format(
                len(t["grid"]) - len(shown), t["number"]))
        out.append("")
    return "\n".join(out).rstrip()


# Attachment types the download tool will fetch, by extension, and the plugin
# folder each lands in so the plugin that opens that type can find it. The
# documents library is organised by FILE TYPE, not by where a file came from,
# so a spreadsheet attached to a Confluence page goes to documents\excel beside
# the user's own workbooks - exactly where the excel plugin looks. Markdown
# goes to this server's knowledge folder instead: it is the RAG index's format,
# and knowledge\confluence is where this server already keeps the pages it
# saves.
ATTACHMENT_FOLDERS = {
    ".pdf": "pdf",
    ".docx": "word",
    ".xlsx": "excel",
    ".xlsm": "excel",
    ".pptx": "powerpoint",
    ".md": "knowledge",
    ".markdown": "knowledge",
}

# What to suggest next after a download, per destination.
ATTACHMENT_NEXT_STEP = {
    "pdf": "convert it with the pdf-to-md plugin (convert_pdf_to_markdown)",
    "word": "open it with the word plugin (msword_open)",
    "excel": "open it with the excel plugin (excel_list_sheets, excel_read_range)",
    "powerpoint": "open it with the powerpoint plugin (powerpoint_open)",
    "knowledge": "run kb_index to make it searchable in the knowledge base",
}

# Ceiling on attachments read while looking one up; a page with more than this
# is vanishingly rare, and the cap stops a runaway pagination loop.
MAX_ATTACHMENTS_SCAN = 1000


def _human_size(size):
    """Bytes -> '48 KB' / '3.2 MB'."""
    try:
        size = float(size)
    except (TypeError, ValueError):
        return "?"
    for unit in ("bytes", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return ("{:.0f} {}" if unit in ("bytes", "KB") else "{:.1f} {}").format(size, unit)
        size /= 1024.0
    return "?"


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
                 kb_dir=None, kb_autosave=False, body_format=BODY_FORMAT,
                 allow_write=False, docs_dir=None):
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
        # Whether the page-writing tools exist at all (CONFLUENCE_ALLOW_WRITE).
        # Off by default: an install that has only ever read Confluence must
        # not gain the power to change it just because the plugin updated.
        self.allow_write = bool(allow_write)
        # Root of the document library attachments are downloaded into (each
        # file type in its own sub-folder, see ATTACHMENT_FOLDERS); None
        # forbids downloading documents.
        self.docs_dir = docs_dir or None
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

    def _open(self, req):
        """
        Send a request and return the open response, turning every failure
        into a ConfluenceError with a message the user can act on. The caller
        reads (and closes) the response.
        """
        try:
            return urllib.request.urlopen(req, timeout=self.timeout,
                                          context=self.ssl_context)
        except urllib.error.HTTPError as e:
            # Try to surface Confluence's error message from the response body.
            detail = ""
            try:
                raw = e.read().decode("utf-8", "replace")
                try:
                    parsed = json.loads(raw)
                    detail = parsed.get("message") or raw
                except ValueError:
                    detail = raw
                detail = detail[:500]
            except Exception:
                pass
            hint = ""
            if e.code == 409:
                hint = (" (the page changed since it was read - someone else "
                        "saved a new version. Read it again and retry.)")
            elif e.code in (401, 403) and req.get_method() != "GET":
                hint = (" (this account may not have permission to edit "
                        "that space or page)")
            self._fail(
                "HTTP {} from Confluence for {} {}{}{}".format(
                    e.code, req.get_method(), req.full_url,
                    (": " + detail) if detail else "", hint
                )
            )
        except urllib.error.URLError as e:
            self._fail(
                "Could not reach Confluence at {} ({}). Check the base URL, "
                "network reachability and TLS settings.".format(req.full_url, e.reason)
            )
        except ssl.SSLError as e:
            self._fail(
                "TLS error talking to Confluence ({}). For an internal CA, set "
                "CONFLUENCE_CA_CERT, or CONFLUENCE_VERIFY_SSL=false to disable "
                "verification.".format(e)
            )

    def _request(self, method, path, params=None, payload=None):
        """
        Perform a REST call and return parsed JSON. `payload`, when given, is
        sent as a JSON body (POST/PUT).
        """
        url = self.base_url + path
        if params:
            # urlencode percent-encodes values (including CQL special chars).
            url = url + "?" + urllib.parse.urlencode(params)
        headers = dict(self.headers)
        data = None
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
            # Confluence's XSRF check: a REST write from a non-browser client
            # must say so, or some instances refuse it with a 403.
            headers["X-Atlassian-Token"] = "no-check"
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        with self._open(req) as resp:
            body = resp.read()
        if not body.strip():
            return {}
        try:
            return json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as e:
            self._fail("Confluence returned a non-JSON response: {}".format(e))

    def _get(self, path, params=None):
        """Perform a GET against the REST API and return parsed JSON."""
        return self._request("GET", path, params)

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

    # -- page lookup shared by the write and attachment tools ----------------
    def _target_page_id(self, page_id, title, space):
        """
        The numeric ID of the page a write or attachment call is aimed at:
        'page_id' when given, else the page with exact 'title' in 'space'.
        """
        if page_id is not None and str(page_id).strip():
            page_id = str(page_id).strip()
            if not page_id.isdigit():
                raise ConfluenceError(
                    "'page_id' must be a numeric content ID (got {!r}). Pass "
                    "'title' plus 'space' if you only know the title.".format(page_id))
            return page_id
        if not title or not space:
            raise ConfluenceError(
                "Identify the page with 'page_id', or with 'title' plus 'space'.")
        params = {"title": title, "spaceKey": space, "limit": 1}
        data = self._get("/rest/api/content", params)
        results = data.get("results", []) or []
        if not results:
            raise ConfluenceError(
                "No page titled {!r} found in space {!r}{}.".format(
                    title, space, self._on_server()))
        return str(results[0].get("id"))

    def _fetch_for_edit(self, page_id):
        """A page with its storage body and version, ready to be rewritten."""
        return self._get(
            "/rest/api/content/" + urllib.parse.quote(page_id, safe=""),
            {"expand": "body.storage,version,space"})

    # -- writing pages (only offered when CONFLUENCE_ALLOW_WRITE is on) ------
    def _require_write(self):
        if not self.allow_write:
            self._fail(
                "Writing to Confluence is switched off on this endpoint. Set "
                "CONFLUENCE_ALLOW_WRITE=true and restart the client to enable "
                "creating and editing pages.")

    @staticmethod
    def _to_storage(body, content_format, task_start=1):
        """Convert a tool's 'body' argument to storage format."""
        fmt = str(content_format or "markdown").strip().lower()
        if fmt in ("storage", "xhtml", "html"):
            return body or ""
        if fmt not in ("markdown", "md"):
            raise ConfluenceError(
                "'content_format' must be 'markdown' (default) or 'storage' "
                "(got {!r}).".format(content_format))
        return markdown_to_storage(body or "", task_start=task_start)

    def _write_summary(self, verb, page):
        """One reply for a create/update/append: what changed and where."""
        link = self._abs_link(page, (page.get("_links") or {}).get("webui", ""))
        lines = [
            "{} page{}:".format(verb, self._on_server()),
            "Title: {}".format(page.get("title", "(untitled)")),
            "ID: {}".format(page.get("id", "?")),
            "Space: {}".format((page.get("space") or {}).get("key", "?")),
            "Version: {}".format((page.get("version") or {}).get("number", "?")),
            "URL: {}".format(link or "(unknown)"),
        ]
        return "\n".join(lines)

    def _put_page(self, page, title, storage, version_message=None,
                  minor_edit=False):
        """Save a new version of an existing page."""
        current = (page.get("version") or {}).get("number")
        if not isinstance(current, int):
            self._fail("Could not read the current version of page {}.".format(
                page.get("id", "?")))
        version = {"number": current + 1, "minorEdit": bool(minor_edit)}
        if version_message:
            version["message"] = str(version_message)
        payload = {
            "id": str(page.get("id")),
            "type": page.get("type") or "page",
            "title": title,
            "space": {"key": (page.get("space") or {}).get("key")},
            "version": version,
            "body": {"storage": {"value": storage, "representation": "storage"}},
        }
        return self._request(
            "PUT", "/rest/api/content/" + urllib.parse.quote(str(page.get("id")), safe=""),
            payload=payload)

    def create_page(self, space, title, body, parent_id=None, parent_title=None,
                    content_format=None):
        self._require_write()
        space = str(space or "").strip()
        title = str(title or "").strip()
        if not space or not title:
            raise ConfluenceError("Both 'space' and 'title' are required.")
        payload = {
            "type": "page",
            "title": title,
            "space": {"key": space},
            "body": {"storage": {
                "value": self._to_storage(body, content_format),
                "representation": "storage",
            }},
        }
        if (parent_id is not None and str(parent_id).strip()) or parent_title:
            parent = self._target_page_id(parent_id, parent_title, space)
            payload["ancestors"] = [{"id": parent}]
        page = self._request("POST", "/rest/api/content", payload=payload)
        log("created page {} ({!r}) in {}".format(page.get("id"), title, space))
        return self._write_summary("Created", page)

    def update_page(self, page_id=None, title=None, space=None, body=None,
                    new_title=None, content_format=None, version_message=None,
                    minor_edit=False, expected_version=None):
        self._require_write()
        if body is None and not new_title:
            raise ConfluenceError(
                "Nothing to change: pass 'body' (the page's new content) "
                "and/or 'new_title'.")
        target = self._target_page_id(page_id, title, space)
        page = self._fetch_for_edit(target)
        current = (page.get("version") or {}).get("number")
        if expected_version is not None and str(expected_version).strip():
            try:
                wanted = int(expected_version)
            except (TypeError, ValueError):
                raise ConfluenceError("'expected_version' must be a whole number.")
            if wanted != current:
                raise ConfluenceError(
                    "Page {} is at version {}, not {} - it has changed since "
                    "it was read, so it was NOT updated. Read it again before "
                    "rewriting it.".format(target, current, wanted))
        if body is None:
            storage = ((page.get("body") or {}).get("storage") or {}).get("value") or ""
        else:
            storage = self._to_storage(body, content_format)
        saved = self._put_page(page, str(new_title).strip() if new_title else page.get("title"),
                               storage, version_message, minor_edit)
        log("updated page {} to version {}".format(
            target, (saved.get("version") or {}).get("number")))
        return self._write_summary("Updated", saved)

    def append_to_page(self, page_id=None, title=None, space=None, body=None,
                       position=None, content_format=None, version_message=None,
                       minor_edit=False):
        self._require_write()
        if not body or not str(body).strip():
            raise ConfluenceError("'body' (the content to add) is required.")
        where = str(position or "end").strip().lower()
        if where not in ("end", "start"):
            raise ConfluenceError("'position' must be 'end' (default) or 'start'.")
        target = self._target_page_id(page_id, title, space)
        page = self._fetch_for_edit(target)
        existing = ((page.get("body") or {}).get("storage") or {}).get("value") or ""
        # The existing body is kept byte-for-byte - macros, layouts and all -
        # and only the new content is converted. New inline tasks continue the
        # page's task numbering so they cannot collide with the old ones.
        addition = self._to_storage(body, content_format,
                                    task_start=next_task_id(existing))
        storage = (existing + addition) if where == "end" else (addition + existing)
        saved = self._put_page(page, page.get("title"), storage,
                               version_message, minor_edit)
        log("appended to page {} (version {})".format(
            target, (saved.get("version") or {}).get("number")))
        return self._write_summary(
            "Added content to the {} of".format(where), saved)

    def update_section(self, page_id=None, title=None, space=None, section=None,
                       body=None, mode=None, content_format=None,
                       allow_macro_removal=False, version_message=None,
                       minor_edit=False):
        """
        Replace (or add to) ONE section of a page - the content under a
        heading, or inside a titled panel/expand - sending the rest of the page
        back byte-for-byte.
        """
        self._require_write()
        if body is None or not str(body).strip():
            raise ConfluenceError("'body' (the section's new content) is required.")
        how = str(mode or "replace").strip().lower()
        if how not in ("replace", "append", "prepend"):
            raise ConfluenceError("'mode' must be 'replace' (default), 'append' or 'prepend'.")
        target = self._target_page_id(page_id, title, space)
        page = self._fetch_for_edit(target)
        existing = ((page.get("body") or {}).get("storage") or {}).get("value") or ""
        found = find_section(existing, section)
        start, end = found["start"], found["end"]
        old = existing[start:end]

        # Replacing a section that itself holds a macro would delete that
        # macro. Refuse unless the caller has confirmed it, naming what would go.
        if how == "replace" and not allow_macro_removal:
            doomed = sorted(set(storage_macro_names(old)))
            if doomed:
                raise ConfluenceError(
                    "The {!r} section contains macro(s) that replacing it would "
                    "delete: {}. Nothing was changed. To change figures in a "
                    "table there, use confluence_update_table (it edits cells "
                    "in place and keeps the macro); to add text, "
                    "mode='append'; or pass allow_macro_removal=true once the "
                    "user has agreed to lose them.".format(
                        found["display"], ", ".join(doomed)))

        addition = self._to_storage(body, content_format,
                                    task_start=next_task_id(existing))
        if how == "replace":
            new_part = addition
        elif how == "append":
            new_part = old + addition
        else:
            new_part = addition + old
        storage = existing[:start] + new_part + existing[end:]
        saved = self._put_page(page, page.get("title"), storage,
                               version_message, minor_edit)
        log("{} section {!r} on page {} (version {})".format(
            how, found["display"], target, (saved.get("version") or {}).get("number")))
        verb = {"replace": "Replaced the content of", "append": "Added to the end of",
                "prepend": "Added to the start of"}[how]
        return (self._write_summary(
            "{} the {!r} section ({}) on".format(
                verb, found["display"],
                found["kind"] if found["level"] is None
                else "heading h{}".format(found["level"])), saved)
            + "\nThe rest of the page was left exactly as it was.")

    # -- tables ------------------------------------------------------------
    def list_tables(self, page_id=None, title=None, space=None, table=None):
        target = self._target_page_id(page_id, title, space)
        page = self._fetch_for_edit(target)
        existing = ((page.get("body") or {}).get("storage") or {}).get("value") or ""
        tables = storage_tables(existing)
        head = "Page {} ({!r}){} has {} stored table(s).".format(
            target, page.get("title", "?"), self._on_server(), len(tables))
        generated = sorted(
            name for name in storage_macro_names(existing)
            if name in ("jira", "jiraissues", "detailssummary", "csv", "table-excerpt-include",
                        "tasks-report-macro", "contentbylabel", "children"))
        note = ""
        if generated:
            note = ("\nAlso on the page, built by Confluence when displayed and so "
                    "NOT editable here: {}.".format(", ".join(generated)))
        if not tables:
            return head + note
        if table is not None and str(table).strip():
            chosen = find_table(tables, table)
            body = render_tables(tables, only=chosen, max_rows=200)
        else:
            body = render_tables(tables)
        return (head + note + "\n\n" + body + "\n\n'^' marks a cell merged "
                "from the row above; '[macro]' a cell holding a macro or other rich "
                "content, which confluence_update_table will not overwrite unless "
                "allowed.")

    @staticmethod
    def _cell_markup(value):
        """A tool's cell value as storage markup (inline Markdown allowed)."""
        if value is None:
            return ""
        if isinstance(value, bool):
            value = "Yes" if value else "No"
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        if isinstance(value, (dict, list)):
            raise ConfluenceError("A cell value must be text or a number.")
        return _inline_markdown(str(value).strip())

    def update_table(self, page_id=None, title=None, space=None, table=None,
                     updates=None, add_rows=None, allow_macro_removal=False,
                     expected_version=None, version_message=None, minor_edit=False):
        """
        Change cells of one stored table in place, and/or add rows to it,
        leaving every other byte of the page - including the macros the table
        sits inside - as it was.
        """
        self._require_write()
        updates = updates or []
        add_rows = add_rows or []
        if isinstance(updates, dict):
            updates = [updates]
        if isinstance(add_rows, dict):
            add_rows = [add_rows]
        if not isinstance(updates, list) or not isinstance(add_rows, list):
            raise ConfluenceError("'updates' and 'add_rows' must be lists.")
        if not updates and not add_rows:
            raise ConfluenceError("Nothing to change: pass 'updates' and/or 'add_rows'.")

        target = self._target_page_id(page_id, title, space)
        page = self._fetch_for_edit(target)
        current = (page.get("version") or {}).get("number")
        if expected_version is not None and str(expected_version).strip():
            if str(expected_version).strip() != str(current):
                raise ConfluenceError(
                    "Page {} is at version {}, not {} - it has changed since it "
                    "was read, so nothing was updated.".format(target, current, expected_version))
        existing = ((page.get("body") or {}).get("storage") or {}).get("value") or ""
        chosen = find_table(storage_tables(existing), table)
        grid, first_data = chosen["grid"], chosen["header_rows"]
        label_col = 0

        # Work out EVERY change before touching anything, so one bad column
        # name in the fifth update cannot leave the first four half-applied.
        edits = {}            # cell start offset -> (start, end, markup)
        report = []
        for n, update in enumerate(updates, 1):
            if not isinstance(update, dict) or not isinstance(update.get("match"), dict) \
                    or not isinstance(update.get("set"), dict) or not update["set"]:
                raise ConfluenceError(
                    "Update {} must be {{\"match\": {{column: value}}, \"set\": "
                    "{{column: new value}}}}.".format(n))
            match = [(_column_index(chosen, k), k, v) for k, v in update["match"].items()]
            hits = [r for r in range(first_data, len(grid))
                    if all(grid[r][ci] is not None and _cell_equal(grid[r][ci]["text"], v)
                           for ci, _k, v in match)]
            if not hits:
                raise ConfluenceError(
                    "Update {}: no row of table {} matches {}. Nothing was changed."
                    .format(n, chosen["number"], json.dumps(update["match"], ensure_ascii=False)))
            if len(hits) > 1 and not update.get("all_matches"):
                raise ConfluenceError(
                    "Update {}: {} rows match {}. Nothing was changed - narrow "
                    "'match', or set \"all_matches\": true in that update."
                    .format(n, len(hits), json.dumps(update["match"], ensure_ascii=False)))
            for r in hits:
                for key, value in update["set"].items():
                    ci = _column_index(chosen, key)
                    cell = grid[r][ci]
                    if cell is None:
                        raise ConfluenceError("Update {}: row {} has no cell in column {!r}."
                                              .format(n, r + 1, key))
                    if cell["row"] != r:
                        raise ConfluenceError(
                            "Update {}: the {!r} cell of that row is merged with the "
                            "row above, so changing it would change both. Match the "
                            "first row of the merged cell instead.".format(n, key))
                    if cell["rich"] and not allow_macro_removal:
                        raise ConfluenceError(
                            "Update {}: the {!r} cell holds a macro or other rich "
                            "content ({!r}) that would be replaced. Nothing was "
                            "changed. Pass allow_macro_removal=true once the user "
                            "agrees.".format(n, key, cell["text"]))
                    markup = self._cell_markup(value)
                    new_inner = _replace_cell_text(existing[cell["start"]:cell["end"]], markup)
                    edits[cell["start"]] = (cell["start"], cell["end"], new_inner)
                    row_label = grid[r][label_col]["text"] if grid[r][label_col] else "row {}".format(r + 1)
                    report.append("{} / {}: {!r} -> {!r}".format(
                        row_label, chosen["columns"][ci]["name"], cell["text"], _plain(markup)))

        # New rows copy the last row's cells (and so its formatting); a last
        # row with merged cells is no pattern to copy, so a plain row is built.
        insert_at, new_rows = None, []
        if add_rows:
            if not chosen["rows"]:
                raise ConfluenceError("Table {} has no rows to add after.".format(chosen["number"]))
            last = chosen["rows"][-1]
            plain = (len(chosen["rows"]) <= first_data
                     or any(c["colspan"] > 1 or c["rowspan"] > 1 for c in last["cells"])
                     or len(last["cells"]) != chosen["width"]
                     or any(cell is None or cell["row"] != len(grid) - 1 for cell in grid[-1]))
            for n, row in enumerate(add_rows, 1):
                if isinstance(row, list):
                    if len(row) > chosen["width"]:
                        raise ConfluenceError("New row {} has {} values; the table has {} "
                                              "columns.".format(n, len(row), chosen["width"]))
                    values = {i: v for i, v in enumerate(row)}
                elif isinstance(row, dict):
                    values = {_column_index(chosen, k): v for k, v in row.items()}
                else:
                    raise ConfluenceError("New row {} must be an object or a list.".format(n))
                if plain:
                    cells = "".join("<td>{}</td>".format(
                        _replace_cell_text("", self._cell_markup(values.get(i))))
                        for i in range(chosen["width"]))
                    new_rows.append("<tr>{}</tr>".format(cells))
                else:
                    pieces, cursor = [], last["start"]
                    for i, cell in enumerate(last["cells"]):
                        pieces.append(existing[cursor:cell["start"]])
                        pieces.append(_replace_cell_text(
                            existing[cell["start"]:cell["end"]] if not cell["rich"] else "",
                            self._cell_markup(values.get(i))))
                        cursor = cell["end"]
                    pieces.append(existing[cursor:last["end"]])
                    new_rows.append("".join(pieces))
                report.append("new row: " + " | ".join(
                    "{}={}".format(chosen["columns"][i]["name"] or i + 1, _plain(self._cell_markup(v)))
                    for i, v in sorted(values.items())))
            insert_at = last["end"]

        storage = existing
        if insert_at is not None:
            storage = storage[:insert_at] + "".join(new_rows) + storage[insert_at:]
        for start, end, markup in sorted(edits.values(), reverse=True):
            storage = storage[:start] + markup + storage[end:]
        saved = self._put_page(page, page.get("title"), storage,
                               version_message, minor_edit)
        log("updated table {} on page {} ({} change(s))".format(
            chosen["number"], target, len(report)))
        where = "table {} (under {!r}{})".format(
            chosen["number"], chosen["heading"] or "no heading",
            ", inside " + " > ".join(chosen["wrappers"]) if chosen["wrappers"] else "")
        return (self._write_summary("Updated {} on".format(where), saved)
                + "\nChanges:\n" + "\n".join("  - " + line for line in report)
                + "\nEverything else on the page, including the table's own "
                "formatting, was left as it was.")

    # -- attachments -------------------------------------------------------
    def _fetch_attachments(self, page_id):
        """Every attachment on a page, each carrying an absolute download URL."""
        items = []
        start = 0
        while len(items) < MAX_ATTACHMENTS_SCAN:
            data = self._get(
                "/rest/api/content/{}/child/attachment".format(
                    urllib.parse.quote(page_id, safe="")),
                {"start": start, "limit": 100, "expand": "version"})
            batch = data.get("results", []) or []
            for item in batch:
                item["_download_url"] = self._abs_link(
                    data, (item.get("_links") or {}).get("download", ""))
            items.extend(batch)
            if not batch or not (data.get("_links") or {}).get("next"):
                break
            start += len(batch)
        return items

    def _attachment_destination(self, filename):
        """
        Where an attachment of this name would be downloaded to: returns
        (folder_key, folder, reason). folder is None when it cannot be
        downloaded, and reason then says why.
        """
        ext = os.path.splitext(filename or "")[1].lower()
        key = ATTACHMENT_FOLDERS.get(ext)
        if not key:
            return None, None, (
                "not a type this server downloads (it fetches {})".format(
                    ", ".join(sorted(ATTACHMENT_FOLDERS))))
        if key == "knowledge":
            if not self.kb_dir:
                return key, None, ("Markdown goes to the knowledge folder, and "
                                   "CONFLUENCE_KB_DIR is off")
            return key, self.kb_dir, ""
        if not self.docs_dir:
            return key, None, ("document downloads are off (the documents "
                               "folder is missing, or CONFLUENCE_DOCS_DIR is off)")
        return key, os.path.join(self.docs_dir, key), ""

    def list_attachments(self, page_id=None, title=None, space=None):
        target = self._target_page_id(page_id, title, space)
        items = self._fetch_attachments(target)
        header = "Page {}{} has {} attachment(s)".format(
            target, self._on_server(), len(items))
        if not items:
            return header + "."
        lines = []
        for item in items:
            name = item.get("title", "(unnamed)")
            ext = item.get("extensions") or {}
            version = item.get("version") or {}
            _key, folder, reason = self._attachment_destination(name)
            lines.append(
                "- {name}\n  id={id}  type={mtype}  size={size}  version={ver}  "
                "updated={when}\n  download: {dest}".format(
                    name=name, id=item.get("id", "?"),
                    mtype=ext.get("mediaType") or "?",
                    size=_human_size(ext.get("fileSize")),
                    ver=version.get("number", "?"),
                    when=str(version.get("when") or "?")[:16].replace("T", " "),
                    dest=("-> " + folder) if folder else "not available - " + reason))
        return header + ":\n\n" + "\n\n".join(lines)

    def download_attachment(self, page_id=None, title=None, space=None,
                            filename=None, attachment_id=None, save_as=None,
                            overwrite=False):
        target = self._target_page_id(page_id, title, space)
        if not filename and not attachment_id:
            raise ConfluenceError(
                "Name the attachment with 'filename' (see "
                "confluence_list_attachments) or 'attachment_id'.")
        items = self._fetch_attachments(target)
        match = None
        if attachment_id:
            want = str(attachment_id).strip().lower()
            want = want if want.startswith("att") else "att" + want
            match = next((i for i in items
                          if str(i.get("id", "")).lower() == want), None)
        else:
            want = str(filename).strip()
            match = next((i for i in items if i.get("title") == want), None)
            if match is None:
                folded = [i for i in items
                          if str(i.get("title", "")).lower() == want.lower()]
                match = folded[0] if len(folded) == 1 else None
        if match is None:
            names = ", ".join(repr(i.get("title")) for i in items[:20]) or "none"
            raise ConfluenceError(
                "No attachment {!r} on page {}{}. Attachments: {}.".format(
                    attachment_id or filename, target, self._on_server(), names))

        name = match.get("title") or "attachment"
        key, folder, reason = self._attachment_destination(name)
        if not folder:
            raise ConfluenceError(
                "{!r} cannot be downloaded: {}.".format(name, reason))
        ext = os.path.splitext(name)[1]
        # basename() then safe_filename() so neither the attachment's name nor
        # 'save_as' can climb out of the folder: "..\\x.xlsx" becomes "x.xlsx".
        # The extension always stays the attachment's own, so 'save_as' cannot
        # turn a PDF into something another plugin would misread.
        source = str(save_as).replace("\\", "/") if save_as else name
        stem = os.path.splitext(os.path.basename(source))[0]
        path = os.path.join(folder, safe_filename(stem) + ext)
        if os.path.exists(path) and not overwrite:
            raise ConfluenceError(
                "{} already exists, so nothing was downloaded - a file already "
                "in the folder must not be replaced without asking. Pass "
                "overwrite=true to replace it (for example with a newer version "
                "of the attachment), or 'save_as' for a different name.".format(path))
        url = match.get("_download_url")
        if not url:
            self._fail("Confluence gave no download link for {!r}.".format(name))

        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            raise ConfluenceError("Could not create {} ({}).".format(folder, e))
        headers = dict(self.headers)
        headers["Accept"] = "*/*"
        req = urllib.request.Request(url, headers=headers, method="GET")
        partial = path + ".part"
        written = 0
        try:
            with self._open(req) as resp:
                ctype = (resp.headers.get("Content-Type") or "").lower()
                if ctype.startswith("text/html") and ext.lower() not in (".html", ".htm"):
                    # A login or error page instead of the file: saving it under
                    # the attachment's name would leave a "workbook" that is
                    # really a web page.
                    self._fail(
                        "Confluence returned a web page instead of {!r} - the "
                        "account may not be able to download it.".format(name))
                with open(partial, "wb") as fh:
                    while True:
                        chunk = resp.read(64 * 1024)
                        if not chunk:
                            break
                        fh.write(chunk)
                        written += len(chunk)
            os.replace(partial, path)
        except OSError as e:
            try:
                os.remove(partial)
            except OSError:
                pass
            raise ConfluenceError(
                "Could not write {} ({}). If the file is open in another "
                "program, close it and try again.".format(path, e))
        except ConfluenceError:
            try:
                os.remove(partial)
            except OSError:
                pass
            raise

        expected = (match.get("extensions") or {}).get("fileSize")
        size_note = ""
        if isinstance(expected, int) and expected != written:
            size_note = ("\nWARNING: Confluence reported {} bytes but {} were "
                         "received.".format(expected, written))
        log("downloaded attachment {!r} from page {} -> {}".format(name, target, path))
        version = (match.get("version") or {}).get("number", "?")
        return (
            "Downloaded {name!r} (version {ver}, {size}) from page {page}{srv} to:\n"
            "  {path}\n\nNext: {step}.{note}".format(
                name=name, ver=version, size=_human_size(written), page=target,
                srv=self._on_server(), path=path,
                step=ATTACHMENT_NEXT_STEP.get(key, "open it"), note=size_note))


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


def page_ref_properties():
    """
    The shared 'page_id' / 'title' / 'space' arguments that identify the page a
    write or attachment tool acts on. Returned fresh for each tool.
    """
    return {
        "page_id": {
            "type": "string",
            "description": "Numeric ID of the page (preferred if known).",
        },
        "title": {
            "type": "string",
            "description": "Exact title of the page, if the ID is not known (needs 'space').",
        },
        "space": {
            "type": "string",
            "description": "Space key of the page (used with 'title'), e.g. 'DOCS'.",
        },
    }


def content_format_property():
    """The shared 'content_format' argument for the three write tools."""
    return {
        "type": "string",
        "enum": ["markdown", "storage"],
        "description": (
            "How 'body' is written. 'markdown' (default): headings, "
            "paragraphs, **bold**, *italic*, `code`, links, bullet and "
            "numbered lists, '- [ ]' task lists (become Confluence tasks), "
            "tables, '> quotes', '> [!NOTE] Title' / [!TIP] / [!WARNING] "
            "panels and ```lang fenced code (becomes a code macro). "
            "'storage': raw Confluence storage-format XHTML, for a macro "
            "Markdown cannot express."
        ),
    }


def version_properties():
    """The shared version-note arguments for the edit tools."""
    return {
        "version_message": {
            "type": "string",
            "description": "Optional note for the page history, e.g. 'Added Q3 actions'.",
        },
        "minor_edit": {
            "type": "boolean",
            "description": "Mark as a minor edit, so watchers are not notified (default false).",
        },
    }


def attachment_tool_definitions():
    """The two attachment tools - always offered; downloading writes locally only."""
    list_props = page_ref_properties()
    download_props = page_ref_properties()
    download_props.update({
        "filename": {
            "type": "string",
            "description": "The attachment's file name exactly as listed, e.g. 'Budget FY26.xlsx'.",
        },
        "attachment_id": {
            "type": "string",
            "description": "The attachment's ID from confluence_list_attachments (e.g. 'att123456'), instead of 'filename'.",
        },
        "save_as": {
            "type": "string",
            "description": (
                "Optional file name to save it under (the original "
                "extension is always kept). Default: the attachment's own name."
            ),
        },
        "overwrite": {
            "type": "boolean",
            "description": (
                "Replace a file of the same name that is already in the "
                "folder (default false). Set it only when the user wants "
                "the local copy refreshed."
            ),
        },
    })
    tables_props = page_ref_properties()
    tables_props["table"] = {
        "type": "string",
        "description": "Optional: show just this table (its number, or the heading above it) with up to 200 rows.",
    }
    return [
        {
            "name": "confluence_list_tables",
            "description": (
                "List the tables stored on a Confluence page - including ones "
                "inside a table filter, column, section or expand macro - with "
                "each table's number, the heading above it, the macros around "
                "it, its column names (a two-row header gives names like "
                "'FY26 Q3') and its rows. Read this before "
                "confluence_update_table. Tables that Confluence builds at "
                "display time (Jira issues, page properties reports) are named "
                "but have no cells to edit."
            ),
            "inputSchema": {"type": "object", "properties": tables_props},
        },
        {
            "name": "confluence_list_attachments",
            "description": (
                "List the files attached to a Confluence page: name, ID, "
                "type, size, version, last update, and the local folder "
                "each would be downloaded to. Identify the page by "
                "'page_id', or 'title' plus 'space'."
            ),
            "inputSchema": {"type": "object", "properties": list_props},
        },
        {
            "name": "confluence_download_attachment",
            "description": (
                "Download one attachment from a Confluence page into the "
                "local folder the matching plugin reads, so it can be opened "
                "straight away: .xlsx/.xlsm -> documents\\excel (excel "
                "plugin), .docx -> documents\\word (word plugin), .pptx -> "
                "documents\\powerpoint (powerpoint plugin), .pdf -> "
                "documents\\pdf (pdf-to-md plugin), .md -> "
                "knowledge\\confluence (knowledge base). Returns the saved "
                "path; pass the file name to the other plugin next (e.g. "
                "excel_list_sheets with the workbook name). Never replaces "
                "an existing file unless 'overwrite' is true."
            ),
            "inputSchema": {"type": "object", "properties": download_props},
        },
    ]


def write_tool_definitions():
    """
    The page-writing tools, offered only when CONFLUENCE_ALLOW_WRITE is on.
    """
    create_props = {
        "space": {
            "type": "string",
            "description": "Space key to create the page in, e.g. 'DOCS'.",
        },
        "title": {
            "type": "string",
            "description": "Title of the new page (must be unique in the space).",
        },
        "body": {
            "type": "string",
            "description": "The page content, in Markdown by default (see 'content_format').",
        },
        "parent_id": {
            "type": "string",
            "description": "Optional numeric ID of the page to create it under.",
        },
        "parent_title": {
            "type": "string",
            "description": "Optional exact title of the parent page, in the same space.",
        },
        "content_format": content_format_property(),
    }
    update_props = page_ref_properties()
    update_props.update({
        "body": {
            "type": "string",
            "description": (
                "The page's COMPLETE new content - it replaces the whole "
                "existing body. Omit it to change only the title."
            ),
        },
        "new_title": {
            "type": "string",
            "description": "Optional new title for the page.",
        },
        "expected_version": {
            "type": "integer",
            "description": (
                "The version number you read the page at. If the page has "
                "moved on since, nothing is saved - pass it whenever the new "
                "body was written from an earlier read."
            ),
        },
        "content_format": content_format_property(),
    })
    update_props.update(version_properties())
    append_props = page_ref_properties()
    append_props.update({
        "body": {
            "type": "string",
            "description": "The content to add, in Markdown by default (see 'content_format').",
        },
        "position": {
            "type": "string",
            "enum": ["end", "start"],
            "description": "Add it at the 'end' (default) or the 'start' of the page.",
        },
        "content_format": content_format_property(),
    })
    append_props.update(version_properties())
    section_props = page_ref_properties()
    section_props.update({
        "section": {
            "type": "string",
            "description": (
                "The heading text, or the panel/expand title, of the part to "
                "change, e.g. \"Director's notes\" (case and punctuation "
                "are ignored)."
            ),
        },
        "body": {
            "type": "string",
            "description": "The new content for that section, in Markdown by default.",
        },
        "mode": {
            "type": "string",
            "enum": ["replace", "append", "prepend"],
            "description": "'replace' the section's content (default), or add to its end/start.",
        },
        "allow_macro_removal": {
            "type": "boolean",
            "description": (
                "Only with the user's agreement: let 'replace' delete macros "
                "that sit inside the section (default false)."
            ),
        },
        "content_format": content_format_property(),
    })
    section_props.update(version_properties())
    table_props = page_ref_properties()
    table_props.update({
        "table": {
            "type": "string",
            "description": "The table's number (from confluence_list_tables) or the heading above it.",
        },
        "updates": {
            "type": "array",
            "items": {"type": "object"},
            "description": (
                "Cell changes: each {\"match\": {column: value}, \"set\": "
                "{column: new value}, \"all_matches\": optional true}. Columns "
                "by header name (or one part of a two-row header, e.g. 'Q3') "
                "or 1-based number; numbers match as numbers ('1,200' = 1200)."
            ),
        },
        "add_rows": {
            "type": "array",
            "description": (
                "Rows to add at the bottom: objects of column -> value, or "
                "lists in column order. They copy the last row's formatting."
            ),
        },
        "allow_macro_removal": {
            "type": "boolean",
            "description": "Only with the user's agreement: allow overwriting a cell that holds a macro (default false).",
        },
        "expected_version": {
            "type": "integer",
            "description": "The page version you read; nothing is saved if the page has moved on.",
        },
    })
    table_props.update(version_properties())
    return [
        {
            "name": "confluence_create_page",
            "description": (
                "Create a new Confluence page in a space, optionally under a "
                "parent page, from Markdown (converted to Confluence "
                "formatting: task lists become real tasks, fenced code a code "
                "macro, '> [!NOTE]' an info panel). Returns the new page's "
                "ID and URL. Only create a page the user has asked for, and "
                "show them the content first unless they have already "
                "approved it."
            ),
            "inputSchema": {
                "type": "object",
                "properties": create_props,
                "required": ["space", "title", "body"],
            },
        },
        {
            "name": "confluence_update_page",
            "description": (
                "Replace the WHOLE content (and/or the title) of an existing "
                "Confluence page, saving a new version. The new 'body' "
                "REPLACES the whole page, so anything left out - including "
                "macros such as task reports or Jira tables - is removed. To "
                "change one part of a page use confluence_update_section; to "
                "add to it, confluence_append_to_page. Use this only for a "
                "full rewrite or a title change. Read the page first, and "
                "pass the version you read as 'expected_version'."
            ),
            "inputSchema": {"type": "object", "properties": update_props},
        },
        {
            "name": "confluence_update_section",
            "description": (
                "Change ONE part of an existing Confluence page and leave "
                "everything else on it - macros, tables, layouts - exactly as "
                "it was. The part is found by 'section': a heading's text (the "
                "content under it, up to the next heading of the same or "
                "higher level, sub-headings included) or the title of a "
                "panel/expand/info macro (its body). mode 'replace' "
                "(default) swaps that content for "
                "'body'; 'append' / 'prepend' add to it. Use this, not "
                "confluence_update_page, for 'update the X section/notes on "
                "page Y'. If the name matches nothing, the sections that exist "
                "are listed. Replacing a section that itself contains a macro "
                "is refused unless allow_macro_removal is true."
            ),
            "inputSchema": {
                "type": "object",
                "properties": section_props,
                "required": ["section", "body"],
            },
        },
        {
            "name": "confluence_update_table",
            "description": (
                "Change figures or text in a table on a Confluence page, in "
                "place: only the text inside the cells named changes, so the "
                "table's widths, colours and merged cells - and any macro it "
                "sits in (table filter, column, section, expand) - are kept. "
                "Pick the table by 'table' (number or the heading above it; "
                "omit if the page has one). 'updates' finds rows by column "
                "value and sets others: [{\"match\": {\"Item\": \"Travel\"}, "
                "\"set\": {\"Q3\": 120}}]. 'add_rows' appends rows: "
                "[{\"Item\": \"Hotels\", \"Q3\": 40}]. Every change is "
                "checked before any is made; an ambiguous match, an unknown "
                "column or a cell holding a macro stops the whole call. Run "
                "confluence_list_tables first to see the column names."
            ),
            "inputSchema": {
                "type": "object",
                "properties": table_props,
            },
        },
        {
            "name": "confluence_append_to_page",
            "description": (
                "Add content to the end (or start) of an existing Confluence "
                "page, keeping everything already on it - macros, layouts and "
                "all - untouched. The safe way to add minutes, actions or a "
                "new section. Saves a new version."
            ),
            "inputSchema": {
                "type": "object",
                "properties": append_props,
                "required": ["body"],
            },
        },
    ]


def base_tool_definitions(allow_write=False):
    """
    The tools as advertised when a single server is configured (JSON-Schema
    input specs). tool_definitions() adds the 'server' argument on top when a
    second server is configured. The write tools are only offered when
    writing is switched on.
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
    ] + attachment_tool_definitions() + (
        write_tool_definitions() if allow_write else [])


def tool_definitions(servers):
    """
    Return the tool list for this configuration.

    With one server it is base_tool_definitions() untouched. With two, every
    tool grows an optional 'server' argument whose enum lists the configured
    names, which is what lets a prompt like "find X on Blue" reach the right
    instance - and, just as importantly, lets Claude see that a second instance
    exists at all.
    """
    tools = base_tool_definitions(servers.default.allow_write)
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

    if name == "confluence_list_attachments":
        return client.list_attachments(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"))

    if name == "confluence_download_attachment":
        return client.download_attachment(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"),
            filename=arguments.get("filename"),
            attachment_id=arguments.get("attachment_id"),
            save_as=arguments.get("save_as"),
            overwrite=bool(arguments.get("overwrite", False)),
        )

    if name == "confluence_create_page":
        return client.create_page(
            arguments.get("space"), arguments.get("title"), arguments.get("body"),
            parent_id=arguments.get("parent_id"),
            parent_title=arguments.get("parent_title"),
            content_format=arguments.get("content_format"),
        )

    if name == "confluence_update_page":
        return client.update_page(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"),
            body=arguments.get("body"),
            new_title=arguments.get("new_title"),
            content_format=arguments.get("content_format"),
            version_message=arguments.get("version_message"),
            minor_edit=bool(arguments.get("minor_edit", False)),
            expected_version=arguments.get("expected_version"),
        )

    if name == "confluence_list_tables":
        return client.list_tables(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"),
            table=arguments.get("table"))

    if name == "confluence_update_table":
        return client.update_table(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"),
            table=arguments.get("table"),
            updates=arguments.get("updates"),
            add_rows=arguments.get("add_rows"),
            allow_macro_removal=bool(arguments.get("allow_macro_removal", False)),
            expected_version=arguments.get("expected_version"),
            version_message=arguments.get("version_message"),
            minor_edit=bool(arguments.get("minor_edit", False)),
        )

    if name == "confluence_update_section":
        return client.update_section(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"),
            section=arguments.get("section"),
            body=arguments.get("body"),
            mode=arguments.get("mode"),
            content_format=arguments.get("content_format"),
            allow_macro_removal=bool(arguments.get("allow_macro_removal", False)),
            version_message=arguments.get("version_message"),
            minor_edit=bool(arguments.get("minor_edit", False)),
        )

    if name == "confluence_append_to_page":
        return client.append_to_page(
            arguments.get("page_id"), arguments.get("title"), arguments.get("space"),
            body=arguments.get("body"),
            position=arguments.get("position"),
            content_format=arguments.get("content_format"),
            version_message=arguments.get("version_message"),
            minor_edit=bool(arguments.get("minor_edit", False)),
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

    A blank value is what a settings file or MCP client passes for a setting
    the user left empty (e.g. "CONFLUENCE_BASE_URL_2": "" with no second
    server), so it must mean "not configured" rather than "configured as empty". An
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


def resolve_docs_dir():
    """
    The documents ROOT attachments are downloaded into, from the environment.

    Precedence: CONFLUENCE_DOCS_DIR (a path of its own, or one of the
    DISABLE_KEYWORDS to forbid downloads), then EVA_DOCUMENTS_DIR, then the
    EVA_DOCUMENTS_DIR fallback in the config block. Returns (path, source):
    source names the variable the path came from, or None when it is the
    built-in default - which decides how loudly a missing folder is reported.
    Returns (None, <variable>) when downloads are switched off.
    """
    own = env_str("CONFLUENCE_DOCS_DIR")
    if own:
        return (None if own.lower() in DISABLE_KEYWORDS else own), "CONFLUENCE_DOCS_DIR"
    root = env_str("EVA_DOCUMENTS_DIR")
    if root:
        return (None if root.lower() in DISABLE_KEYWORDS else root), "EVA_DOCUMENTS_DIR"
    return EVA_DOCUMENTS_DIR, None


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
    if servers.default.docs_dir:
        log("Attachments      : {}\\<excel|word|powerpoint|pdf> (.md -> the KB "
            "save folder)".format(servers.default.docs_dir))
    else:
        log("Attachments      : document downloads disabled")
    log("Page writing     : {}".format(
        "ENABLED (create / update / append)" if servers.default.allow_write
        else "off (read-only)"))
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
    allow_write = env_bool("CONFLUENCE_ALLOW_WRITE", ALLOW_WRITE)

    # The documents root for attachment downloads. Unlike a sandboxed
    # document server this is NOT fatal when missing: downloading is two tools,
    # and searching and reading Confluence must still work on an endpoint that
    # has no document library. A folder the user configured is reported loudly,
    # because a missing one is almost always a typo.
    docs_dir, docs_source = resolve_docs_dir()
    docs_off = docs_dir is None
    if docs_dir and not os.path.isdir(docs_dir):
        if docs_source:
            log("WARNING: {} points at {}, which does not exist. Attachment "
                "downloads (other than Markdown) are disabled until it does - "
                "fix the path, or create the folder.".format(docs_source, docs_dir))
        else:
            log("documents folder {} does not exist; attachment downloads "
                "(other than Markdown) are disabled. Copy the eva\\ folder to "
                "H:\\Eva, or set EVA_DOCUMENTS_DIR.".format(docs_dir))
        docs_dir = None

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
                allow_write=allow_write,
                docs_dir=docs_dir,
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
    if servers.default.docs_dir:
        log("attachment downloads -> {} (a sub-folder per file type)".format(
            servers.default.docs_dir))
    elif docs_off:
        log("attachment downloads disabled ({}=off)".format(docs_source))
    if servers.default.allow_write:
        log("WRITE ENABLED: pages can be created and edited "
            "(CONFLUENCE_ALLOW_WRITE=true)")
    else:
        log("read-only: page writing is off (set CONFLUENCE_ALLOW_WRITE=true "
            "to enable it)")
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
