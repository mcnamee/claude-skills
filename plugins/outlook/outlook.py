#!/usr/bin/env python3
"""
outlook.py (v6.2.0)
======================

A single-file MCP (Model Context Protocol) server giving an LLM read-only
access to a locally installed *classic* Microsoft Outlook client (mail +
calendar) on Windows, via COM automation. Read-only means the mailbox: nothing
here can send, reply, accept, move or delete. What it does write is local and
only on request - an email kept as Markdown for the knowledge base, and a
printable PDF day planner.

Designed for an airgapped Windows endpoint where Outlook is installed, running,
and logged into an on-premises Exchange profile. This script makes NO network
calls; all access is local COM to the already-authenticated Outlook process.

Transport: newline-delimited JSON-RPC 2.0 over stdio (the standard MCP stdio
transport).

DEPENDENCY
----------
Requires pywin32 (win32com.client + pythoncom) - the ONLY non-stdlib
dependency. Install via your pip proxy:  pip install pywin32

REQUIREMENTS
------------
- Classic Win32 Outlook (NOT "New Outlook", which has no COM support).
- Outlook installed, running, and logged into a profile.

TOOLS EXPOSED (all read-only on the mailbox itself)
---------------------------------------------------
- outlook_list_recent_emails : recent Inbox messages
- outlook_search_emails      : search Inbox by subject / sender
- outlook_get_email          : full body of one message by EntryID
- outlook_get_calendar       : calendar events in a date range (recurring expanded).
                               Date filtering happens in Python, NOT via Outlook's
                               locale-sensitive Restrict(), so it works identically
                               under any regional date settings. When no events
                               match, the reply includes a [debug] section showing
                               the last 5 calendar items scanned (start date +
                               in-range flag) so a filtering fault stays visible.
- outlook_print_calendar     : a PRINTABLE PDF day planner for one day - see below
- outlook_list_sent_emails   : messages you SENT, in a date range (e.g. "what did I do last week")
- outlook_search_recent      : all mail across Inbox/Archive/Sent in a date range (configurable)
- outlook_list_folders       : list every mail folder across all stores (to configure the above)

Markdown export for a RAG knowledge base (ON REQUEST)
-----------------------------------------------------
Reading an email does NOT save it. outlook_get_email takes an optional
'save_to_kb' argument, false by default; only when it is true is the message
written out as a Markdown file into the knowledge-base folder, so the content
feeds a local RAG index (e.g. alongside knowledge-base.py). Files are named
'Email - <date> - <subject> (<id>).md' and overwritten if the same message is
saved again. Blocked (blacklisted) messages are NEVER written - a save only
runs after a message has cleared the content filter.

Saving on request rather than on read is what keeps a mailbox from becoming
the knowledge base: every message opened while answering a question used to be
embedded and quotable in later answers, whether it deserved to be or not.

The folder is the "email" sub-folder of the suite-wide RAG root
(%EVA_KNOWLEDGE_DIR%\email, C:\Eva\knowledge\email on a stock install), which
sits inside the knowledge-base plugin's corpus so saved mail is actually
indexed. Point OUTLOOK_KB_DIR at a full path of its own to override just this
server.

To make saving impossible, set OUTLOOK_KB_DIR=off: a save_to_kb request is then
refused and the server touches no disk at all. To go the other way and save
EVERY email read, set OUTLOOK_KB_AUTOSAVE=true.
Worth a deliberate decision either way: saving turns correspondence into plain
text files that are then embedded and quotable in answers. See
eva\knowledge\email\README.md.

Printable day planner (outlook_print_calendar)
----------------------------------------------
Asked for "today's printable calendar" or "tomorrow's planner", this writes an
A4 LANDSCAPE PDF laid out as a bifold spread: the day itself on an hour-by-hour
timeline down the left half, and the following days summarised on the right.
Print it single-sided and fold it in half.

The 'date' argument takes 'today' (the default), 'tomorrow', 'yesterday', a day
offset like '+2', or YYYY-MM-DD. The words are there on purpose - the server
resolves them against the ENDPOINT's clock, so a printed planner cannot come out
a day wrong because the model's idea of today was stale.

The PDF is written with NO third-party library: it is drawn straight into the
PDF imaging model from the standard library alone (rectangles, rules, and text
in the base-14 Helvetica every reader carries), so printing adds nothing to the
pip dependencies and nothing to embed. The design it follows uses Manrope;
Helvetica is the closest stand-in that needs no font file.

What ends up on the page:

- Each meeting is a block positioned and sized by its real start and end.
  Overlapping meetings share the lane in columns, the way Outlook's day view
  does; a run of back-to-back short meetings stays one column of thin blocks.
- The timeline covers the working day (CALENDAR_DAY_START_HOUR ..
  CALENDAR_DAY_END_HOUR, or OUTLOOK_CALENDAR_HOURS) and STRETCHES to take in
  anything scheduled outside it - a 6 am flight is on the page, not off it.
- All-day items sit above the grid as a strip of chips.
- Block colour comes from your own Outlook CATEGORIES, and needs nothing
  configured: the colour Outlook already holds against each category is read
  off the profile and printed at full strength. Outlook's own swatches are
  muted on screen and would wash out on paper, so each of its 25 colours maps
  to a bold equivalent (its green becomes a bold green, and so on). Use
  OUTLOOK_CALENDAR_COLOURS="Leadership=purple,Client=green" only to overrule a
  category whose Outlook colour is not the one you want on the page. Anything
  uncategorised falls back to what Outlook can actually prove: somebody outside
  your own SMTP domain is invited (External), it is internal (Internal), or
  nobody is invited at all (Personal). Nothing is guessed from the subject line.
- Type on a block goes white or dark automatically, by how bright the fill is,
  so a yellow category is readable rather than white-on-yellow.
- Tentative or free-marked time is drawn hollow, the way a diary pencils
  something in.
- The right-hand panel skips days with nothing in the diary by default, so a
  Friday planner shows the week ahead rather than two blank weekend panels.
  Pass skip_empty_days=false for strictly consecutive days.

The blacklist applies in full: a withheld event never reaches the page, not even
as an unlabelled block, and the footer carries the count. The file lands in the
documents folder below and an existing file of the same name is overwritten, so
re-printing a day replaces that day's sheet rather than piling up copies.

==============================================================================
CONFIGURATION  -  all editable settings live in the "USER CONFIGURATION" block
just below this docstring. Edit them there; nothing else needs changing.
==============================================================================

1. BLACKLIST_TERMS  (content compliance filter)
   Any email or calendar item whose content contains a blacklisted term is
   WITHHELD - never sent to the AI. Use it for classified / protectively-marked
   material that may not lawfully be processed by the AI.
   - list / search : blocked items are omitted; a count of withheld items is
                     shown (no subject/sender of a blocked item is ever revealed).
   - get_email     : a blocked message returns a generic refusal, not content.
   - calendar      : blocked events are omitted; a withheld count is shown.
   - folders       : folder NAMES matching the blacklist are withheld from
                     outlook_list_folders and skipped by outlook_search_recent
                     (results are labelled with their folder path, so a marked
                     folder name must never appear in output).
   - FAIL CLOSED (optional): set REQUIRE_BLACKLIST = True below, or set
                     OUTLOOK_REQUIRE_BLACKLIST=1, to make the server refuse to
                     start if the blacklist is empty, so a missing terms file
                     cannot silently disable filtering.
   - The matched term is NEVER shown to the AI (that would leak the marking); it
     is logged to STDERR only, for local audit.
   - FAIL-SAFE: if an item's subject or body cannot be read (so it cannot be
     cleared), the item is treated as BLOCKED.
   Seed the list with PLAIN words (PROTECTED, SECRET, CABINET). In "word" mode
   these are still caught inside bracketed markings like [SEC=PROTECTED]. Do NOT
   add very common words (e.g. "OFFICIAL") unless you mean to block almost
   everything. Replace the seeded examples with your real classification scheme.

2. BLACKLIST_MATCH_MODE
   - "word"      : whole-term match. "SECRET" is caught inside "[SEC=SECRET]"
                   but NOT inside "secretary". Best for plain words. (default)
   - "substring" : matches anywhere. Use only when a term itself contains
                   punctuation (e.g. "[SEC=PROTECTED]") that word mode misses.
   Matching is always case-insensitive.

3. Tunable caps (MAX_BODY_CHARS / CALENDAR_HARD_CAP / SEARCH_SCAN_CAP)
   Safety/size limits. Lower MAX_BODY_CHARS if your local model has a small
   context window. The other two are guard rails you can usually leave as-is.

3b. Day-planner defaults (CALENDAR_DAY_START_HOUR / CALENDAR_DAY_END_HOUR /
    CALENDAR_LOOKAHEAD_DAYS / CALENDAR_CATEGORY_COLOURS)
   The working day the printed timeline covers, how many following days the
   right-hand panel lists, and the Outlook category -> colour map. All four have
   an environment variable so a plugin install never needs the file edited.

4. SEARCH_ALL_FOLDERS  (DEFAULT folders the combined outlook_search_recent covers)
   A list of folder NAMES matched across every store in the profile (main
   mailbox, online archive, mounted PST). Default: Inbox, Sent Items, Archive.
   "Archive" is ambiguous - it may be a mailbox folder, an online archive, or a
   PST - so run outlook_list_folders first to see what actually exists and set
   this list to the real names. Note: matching is by name across ALL stores, so
   if you have shared mailboxes with same-named folders they may be included;
   each result is labelled with its store/folder so you can see the source.
   This is only the DEFAULT set, resolved in this order (later wins):
     (a) SEARCH_ALL_FOLDERS here in the file;
     (b) the OUTLOOK_SEARCH_FOLDERS environment variable (comma-separated
         names), if set;
     (c) a per-call "folders" argument to outlook_search_recent.
   Use (b) to configure the default from the MCP client config without editing
   this file.

EXTERNAL BLACKLIST FILE (optional)
----------------------------------
Instead of (or in addition to) editing BLACKLIST_TERMS, supply extra terms in a
file named by the OUTLOOK_BLACKLIST_FILE environment variable. One term per
line; '#' starts a comment. File terms are ADDED to the built-in list (they
never reduce it). Example file contents:

    # outlook-blacklist.txt  - classification terms to withhold from the AI
    PROTECTED
    SECRET
    TOP SECRET
    CABINET
    CABINET-IN-CONFIDENCE

CONFIGURATION  (environment variables, no folder flags)
-------------------------------------------------------
The whole plugin suite is configured by four environment variables, set once
for your Windows account. This server uses three of them:

    EVA_PYTHON          full path to the python.exe that has pywin32
                        installed, e.g. C:\\Python311\\python.exe (read by the
                        plugin manifest, not by this file)
    EVA_KNOWLEDGE_DIR   root of the RAG corpus (default C:\\Eva\\knowledge)
    EVA_DOCUMENTS_DIR   root of the document library (default C:\\Eva\\documents)

This server works in one sub-folder of each:

    %EVA_KNOWLEDGE_DIR%\\email   Where an email is saved as Markdown when you
                                ask for it to be kept, for the knowledge-base
                                plugin to index. Reading an email does NOT save
                                it, and a blacklisted message is never written.
                                THIS FOLDER MUST EXIST - the server will not
                                start without it.

    %EVA_DOCUMENTS_DIR%\\pdf     Where outlook_print_calendar writes the day
                                planner. The document library is organised by
                                file type and a printed planner is a PDF, so it
                                lands beside your own PDFs rather than in an
                                output folder of its own. Missing is NOT fatal:
                                the server says so and disables printing, since
                                mail must stay readable either way.

To set them permanently for your account (PowerShell, one-off):

    [Environment]::SetEnvironmentVariable("EVA_PYTHON", "C:\\Python311\\python.exe", "User")
    [Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "C:\\Eva\\knowledge", "User")
    [Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "C:\\Eva\\documents", "User")

Copy the repo's eva\\ folder to C:\\Eva and the folder exists - see
eva\\README.md.

Server-specific settings, all optional and all environment variables:

    OUTLOOK_KB_DIR              override the save folder with a full path of
                                its own, or "off" to forbid saving entirely,
                                after which no email is written to disk.
    OUTLOOK_KB_AUTOSAVE=true    save EVERY email read, without being asked
                                (off by default - see KB_AUTOSAVE below).
    OUTLOOK_DOCS_DIR            override the day-planner folder with a full
                                path of its own, or "off" to forbid printing
                                entirely.
    OUTLOOK_CALENDAR_HOURS      the working day the printed timeline starts
                                from, e.g. "7-19" (default "8-18"). It always
                                stretches to fit anything outside it.
    OUTLOOK_CALENDAR_COLOURS    Overrule the colour Outlook already holds
                                against a category, e.g.
                                "Leadership=purple,Client=green". Names:
                                red, orange, peach, yellow, amber, green,
                                teal, olive, blue, purple, maroon, rose,
                                steel, silver, slate, black, grey, and a
                                "dark-" prefix on red, orange, peach, yellow,
                                green, teal, olive, blue, purple, maroon.
                                Normally unnecessary - the colours come from
                                Outlook on their own.
    OUTLOOK_SEARCH_FOLDERS      comma-separated folder names for
                                outlook_search_recent, e.g.
                                "Inbox,Sent Items,Archive".
    OUTLOOK_BLACKLIST_FILE      path to a file of extra blacklist terms.
    OUTLOOK_REQUIRE_BLACKLIST=1 refuse to start with an empty blacklist.

There are NO configuration command-line flags: everything above is an
environment variable, so two settings can never disagree. The only flags this
server takes are --check and --version.

INSTALLING INTO CLAUDE CODE
---------------------------
This server ships as the "outlook" Claude Code plugin (its manifest is
.claude-plugin/plugin.json next to this file), so the normal install is:

    /plugin marketplace add C:\\path\\to\\claude-skills
    /plugin install outlook@mcnamee-claude-skills

Claude Code prompts only for the optional search folders and blacklist file;
the interpreter and the knowledge folder come from the environment variables
above. PYTHONUTF8=1 is set for you by the manifest.

To register the server by hand instead (PowerShell):

    claude mcp add outlook --scope user -e PYTHONUTF8=1 -- $env:EVA_PYTHON C:\\path\\to\\outlook.py

See README.md next to this file for the full settings reference.

USAGE / TESTING
---------------
- As an MCP server (normal mode): launched by the MCP client. Run with no
  arguments; every setting comes from the environment.
- Connectivity check (run manually on the endpoint before wiring it in):

      python outlook.py --check

  Connects to Outlook and prints mailbox diagnostics, folder paths and
  blacklist status to stderr, then exits.

IMPORTANT (stdio-on-Windows pitfalls)
-------------------------------------
- ALL diagnostic output goes to stderr. Anything on stdout that is not a
  JSON-RPC message corrupts the protocol stream.
- Set PYTHONUTF8=1 in the launching environment (the plugin manifest does this
  for you) so stdout is UTF-8 and Unicode subjects do not crash on cp1252.
"""

# Semantic version of this server. Bump on EVERY change (see CLAUDE.md):
# MAJOR = breaking config/tool change, MINOR = new feature, PATCH = fix.
__version__ = "6.2.0"

import os
import re
import sys
import zlib
import json
import hashlib
import argparse
import datetime
import traceback
import unicodedata


# ============================================================================
# USER CONFIGURATION  -  EDIT THIS BLOCK  (see the docstring above for detail)
# ============================================================================

# --- 1. Content blacklist: terms that cause an item to be withheld from the AI.
#        Seeded with example AU-style markings; REPLACE with your real scheme.
BLACKLIST_TERMS = [
]

# --- 2. How blacklist terms are matched: "word" (default) or "substring".
BLACKLIST_MATCH_MODE = "word"

# --- 2b. Fail closed if the blacklist ends up empty. When True (or when
#         OUTLOOK_REQUIRE_BLACKLIST=1 is set), the server REFUSES TO START
#         unless at least one blacklist term is active. Turn this on when the
#         compliance filter is mandatory in your environment, so a
#         missing/empty terms file cannot silently disable filtering.
REQUIRE_BLACKLIST = False

# --- 3. Tunable caps.
MAX_BODY_CHARS = 20000      # truncate very long email bodies for the LLM
CALENDAR_HARD_CAP = 1000    # ceiling on calendar events collected (anti-runaway)
SEARCH_SCAN_CAP = 500       # ceiling on raw search hits scanned

# --- 4. Folders included in the combined outlook_search_recent tool. Matched by
#        folder NAME (case-insensitive) across EVERY mailbox/store in the profile,
#        including an online (In-Place) archive or a mounted archive.pst. Run the
#        outlook_list_folders tool to see the exact names available, then edit this
#        list to match your setup. (The separate outlook_list_sent_emails tool is
#        unaffected by this list.)
SEARCH_ALL_FOLDERS = ["Inbox", "Sent Items", "Archive"]

# --- 5. KB_DIR  (Markdown export for a local RAG knowledge base, on request).
#        Where outlook_get_email writes a message as a Markdown file WHEN the
#        call asks for it (save_to_kb=true), so the content can feed a local
#        RAG index (e.g. alongside knowledge-base.py). Files are named
#        'Email - <date> - <subject> (<id>).md' and overwritten if the same
#        message is saved again. Blocked (blacklisted) messages are NEVER
#        written - a save only runs once a message has cleared the content
#        filter.
#        RESOLVED FROM THE ENVIRONMENT in main(): the "email" sub-folder of
#        %EVA_KNOWLEDGE_DIR% (the suite-wide RAG root), or OUTLOOK_KB_DIR for
#        a full path of its own. The literal here is what a stock C:\Eva
#        install resolves to; it MUST stay inside the knowledge-base plugin's
#        corpus (C:\Eva\knowledge) or the saved mail would never be indexed.
#        Set OUTLOOK_KB_DIR=off to forbid saving entirely - no email is then
#        written to disk and a save_to_kb request is refused. (The day planner
#        has its own switch, OUTLOOK_DOCS_DIR, in 5b below.)
SUBFOLDER = "email"                      # this server's knowledge sub-folder
EVA_KNOWLEDGE_DIR = r"C:\Eva\knowledge"  # fallback for the suite-wide root
KB_DIR = r"C:\Eva\knowledge\email"

# --- 5b. PDF_DIR  (where outlook_print_calendar writes the printable planner).
#        The ONE folder this server creates documents in. It is the "pdf"
#        sub-folder of %EVA_DOCUMENTS_DIR%, because the document library is
#        organised by file type and a printed planner is a PDF - it lands beside
#        your own PDFs rather than in an output folder of its own.
#        RESOLVED FROM THE ENVIRONMENT in main(); the literal here is what a
#        stock C:\Eva install resolves to. Set OUTLOOK_DOCS_DIR to a full path
#        of its own, or to "off" to forbid printing entirely.
DOCS_SUBFOLDER = "pdf"                     # this server's documents sub-folder
EVA_DOCUMENTS_DIR = r"C:\Eva\documents"     # fallback for the suite-wide root
PDF_DIR = r"C:\Eva\documents\pdf"

# --- 5c. Printable day planner defaults.
#        CALENDAR_DAY_START_HOUR / CALENDAR_DAY_END_HOUR are the working day the
#        timeline shows. It always STRETCHES to take in anything scheduled
#        outside those hours, so they are a floor and a ceiling on the printed
#        grid, not a filter. Override with OUTLOOK_CALENDAR_HOURS="7-19".
CALENDAR_DAY_START_HOUR = 8
CALENDAR_DAY_END_HOUR = 18

#        How many following days the right-hand panel lists (1-6).
CALENDAR_LOOKAHEAD_DAYS = 4

#        Outlook CATEGORY -> block colour. Usually EMPTY, because the colour
#        Outlook already holds against each category is read off the profile
#        and printed at full strength (see OL_CATEGORY_COLOURS). Put an entry
#        here, or in OUTLOOK_CALENDAR_COLOURS, only to overrule a category whose
#        Outlook colour is not the one you want on the page. Keys are matched
#        case-insensitively against the categories on an appointment; the first
#        match wins. Names are the keys of PLANNER_COLOURS. Anything
#        uncategorised falls back to what Outlook can prove on its own -
#        somebody outside your domain is invited (External), it is an internal
#        meeting (Internal), or nobody is invited at all (Personal).
CALENDAR_CATEGORY_COLOURS = {
}

# --- 6. KB_AUTOSAVE. Whether reading an email saves it WITHOUT being asked.
#        False means a message is saved only when the call passes
#        save_to_kb=true, which is the point: mail skimmed to answer a question
#        is not a filing decision, and a knowledge base full of correspondence
#        nobody chose returns irrelevant answers later. Set to True here, or
#        set OUTLOOK_KB_AUTOSAVE=true, to save every email read.
KB_AUTOSAVE = False

# --- 7. DISABLE_KEYWORDS. Values that mean "explicitly turned off" for a
#        folder setting. An MCP client can only pass strings, and a BLANK
#        string is what it substitutes for a setting the user left empty -
#        which means "not configured", falling back to the suite-wide root. So
#        a keyword is needed to say "definitely off": OUTLOOK_KB_DIR=off
#        forbids saving an email, OUTLOOK_DOCS_DIR=off forbids printing a
#        planner, and with both off the server writes no local file at all.
DISABLE_KEYWORDS = frozenset(("off", "none", "no", "false", "disabled"))

# ============================================================================
# END USER CONFIGURATION  -  you should not need to edit below this line
# ============================================================================


# ---------------------------------------------------------------------------
# Stream setup: force UTF-8 so non-ASCII subjects/bodies cannot crash output.
# ---------------------------------------------------------------------------
for _stream in ("stdin", "stdout"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8")
    except Exception:
        pass


def log(message):
    """Write a diagnostic line to stderr ONLY. Never touch stdout here."""
    print(message, file=sys.stderr, flush=True)


def env(name):
    """
    Read an environment variable, treating blank as unset.

    A BLANK value means "not configured" - that is what an MCP client
    substitutes for a setting the user left empty - as does an unexpanded
    "${...}" placeholder, which is what a client leaves behind when the
    variable it refers to does not exist. Either one returns None so the
    caller's default applies.
    """
    raw = (os.environ.get(name) or "").strip()
    if not raw or (raw.startswith("${") and raw.endswith("}")):
        return None
    return raw


def env_flag(name, default):
    """Read a boolean environment variable, falling back to `default`."""
    raw = env(name)
    if raw is None:
        return default
    return raw.lower() in ("1", "true", "yes", "on")


def resolve_kb_dir():
    """
    The knowledge-base folder, from the environment.

    Precedence: OUTLOOK_KB_DIR (a full path of its own, or one of the
    DISABLE_KEYWORDS to forbid saving), then EVA_KNOWLEDGE_DIR with this
    server's "email" sub-folder appended, then that same sub-folder of the
    EVA_KNOWLEDGE_DIR fallback in the config block. Returns None when saving
    is switched off.
    """
    own = env("OUTLOOK_KB_DIR")
    if own:
        return None if own.lower() in DISABLE_KEYWORDS else own
    root = env("EVA_KNOWLEDGE_DIR")
    if root:
        return None if root.lower() in DISABLE_KEYWORDS else os.path.join(root, SUBFOLDER)
    return os.path.join(EVA_KNOWLEDGE_DIR, SUBFOLDER)


def resolve_pdf_dir():
    """
    The folder outlook_print_calendar writes the planner into.

    Precedence: OUTLOOK_DOCS_DIR (a full path of its own, or one of the
    DISABLE_KEYWORDS to forbid printing), then EVA_DOCUMENTS_DIR with this
    server's "pdf" sub-folder appended, then that same sub-folder of the
    EVA_DOCUMENTS_DIR fallback in the config block. Returns (path, configured):
    `configured` is False only when nothing was set at all, which is what
    decides whether a missing folder is worth complaining about. Returns
    (None, True) when printing is switched off.
    """
    own = env("OUTLOOK_DOCS_DIR")
    if own:
        return (None, True) if own.lower() in DISABLE_KEYWORDS else (own, True)
    root = env("EVA_DOCUMENTS_DIR")
    if root:
        if root.lower() in DISABLE_KEYWORDS:
            return None, True
        return os.path.join(root, DOCS_SUBFOLDER), True
    return os.path.join(EVA_DOCUMENTS_DIR, DOCS_SUBFOLDER), False


def parse_category_colours(raw):
    """
    Parse OUTLOOK_CALENDAR_COLOURS ("Leadership=purple,Client=green") into a
    lower-cased category -> colour map. Unknown colour names are dropped with a
    warning rather than failing the server: a typo in a cosmetic setting must
    not stop mail being read.
    """
    mapping = {}
    for pair in (raw or "").split(","):
        pair = pair.strip()
        if not pair or "=" not in pair:
            continue
        category, _, colour = pair.partition("=")
        category = category.strip().lower()
        colour = colour.strip().lower()
        if not category:
            continue
        if colour not in PLANNER_COLOURS:
            log("WARNING: ignoring OUTLOOK_CALENDAR_COLOURS entry '{0}' - "
                "'{1}' is not one of {2}.".format(
                    pair, colour, ", ".join(sorted(PLANNER_COLOURS))))
            continue
        mapping[category] = colour
    return mapping


# --version must work even when pywin32 is not installed (or off Windows),
# so answer it before the import below can fail.
if "--version" in sys.argv:
    print("outlook-mcp {0}".format(__version__))
    sys.exit(0)

# ---------------------------------------------------------------------------
# pywin32 import. If it is missing the server cannot function, so fail loudly.
# ---------------------------------------------------------------------------
try:
    import pythoncom
    import win32com.client
except ImportError:
    log("FATAL: pywin32 is not installed. Run:  pip install pywin32")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Internal constants (not user configuration)
# ---------------------------------------------------------------------------
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_CALENDAR = 9
OL_CLASS_MAIL = 43  # olMail

# Ceiling on Sent Items scanned when listing by date. Items are sorted
# newest-first and iteration stops as soon as it passes the start of the
# window, so a recent window is reached immediately; this only bounds
# pathological queries for very old windows in a huge Sent folder.
SENT_SCAN_CAP = 5000

# PR_TRANSPORT_MESSAGE_HEADERS (Unicode) - full internet headers, which carry
# the authoritative protective marking (e.g. X-Protective-Marking).
PROP_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"

# Compiled at startup by build_blacklist(); None means no filtering is active.
_BLACKLIST_RE = None

# Outlook category -> planner colour, from CALENDAR_CATEGORY_COLOURS and then
# OUTLOOK_CALENDAR_COLOURS. Keys are lower-cased category names.
_CATEGORY_COLOURS = {}

# Effective default folder set for outlook_search_recent. Initialised from
# SEARCH_ALL_FOLDERS; may be replaced by OUTLOOK_SEARCH_FOLDERS. A per-call
# "folders" argument still overrides this.
_SEARCH_FOLDERS = list(SEARCH_ALL_FOLDERS)


# ---------------------------------------------------------------------------
# Blacklist construction and scanning
# ---------------------------------------------------------------------------

def build_blacklist(extra_terms=None):
    """
    Compile BLACKLIST_TERMS (plus any extra_terms) into a single regex and set
    the module-level _BLACKLIST_RE. Logs the active status to stderr.
    """
    global _BLACKLIST_RE

    terms = list(BLACKLIST_TERMS)
    if extra_terms:
        terms.extend(extra_terms)

    # Clean, de-duplicate (case-insensitively), drop blanks.
    cleaned = []
    seen = set()
    for term in terms:
        term = (term or "").strip()
        key = term.lower()
        if term and key not in seen:
            seen.add(key)
            cleaned.append(term)

    if not cleaned:
        _BLACKLIST_RE = None
        log("WARNING: content blacklist is EMPTY - NO compliance filtering is active.")
        return

    escaped = [re.escape(term) for term in cleaned]
    if BLACKLIST_MATCH_MODE == "substring":
        pattern = "(?:" + "|".join(escaped) + ")"
    else:
        # \b on each side: whole-term match. Catches "SECRET" inside "[SEC=SECRET]"
        # (brackets/equals are non-word chars) but not "secretary".
        pattern = r"\b(?:" + "|".join(escaped) + r")\b"

    _BLACKLIST_RE = re.compile(pattern, re.IGNORECASE)
    log("Content blacklist ACTIVE: {0} term(s), mode='{1}'.".format(
        len(cleaned), BLACKLIST_MATCH_MODE))


def blacklisted_match(text):
    """Return the matched blacklisted term if `text` contains one, else None."""
    if _BLACKLIST_RE is None or not text:
        return None
    found = _BLACKLIST_RE.search(text)
    return found.group(0) if found else None


def load_blacklist_file(path):
    """Read extra blacklist terms from a file (one per line; '#' = comment)."""
    terms = []
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        for line in handle:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            terms.append(line)
    return terms


def email_block_reason(item):
    """
    Decide whether a mail item must be withheld from the AI.

    Returns the matched term (str) if BLOCKED, else None. Fail-safe: if Subject
    or Body cannot be read (so the item cannot be cleared), treat as BLOCKED.
    """
    if _BLACKLIST_RE is None:
        return None

    parts = []
    # Essential fields. If these cannot be read, we cannot verify the item is
    # clean, so we block it.
    try:
        parts.append(item.Subject or "")
    except Exception:
        return "<unreadable subject>"
    try:
        parts.append(item.Body or "")
    except Exception:
        return "<unreadable body>"

    # Supplementary fields - best-effort; a read failure just skips the field.
    for getter in (lambda: item.SenderName,
                   lambda: item.To,
                   lambda: item.CC,
                   lambda: item.Categories):
        try:
            value = getter()
            if value:
                parts.append(str(value))
        except Exception:
            pass

    # Authoritative protective marking lives in the transport headers.
    try:
        headers = item.PropertyAccessor.GetProperty(PROP_TRANSPORT_HEADERS)
        if headers:
            parts.append(str(headers))
    except Exception:
        pass

    return blacklisted_match("\n".join(parts))


def appointment_block_reason(item):
    """As email_block_reason, for a calendar appointment. Fail-safe on Subject/Body."""
    if _BLACKLIST_RE is None:
        return None

    parts = []
    try:
        parts.append(item.Subject or "")
    except Exception:
        return "<unreadable subject>"
    try:
        parts.append(item.Body or "")
    except Exception:
        return "<unreadable body>"

    for getter in (lambda: item.Location,
                   lambda: item.Organizer,
                   lambda: item.Categories):
        try:
            value = getter()
            if value:
                parts.append(str(value))
        except Exception:
            pass

    return blacklisted_match("\n".join(parts))


# ---------------------------------------------------------------------------
# Outlook connection (lazy, cached, with reconnect-on-failure)
# ---------------------------------------------------------------------------
_namespace = None


def get_namespace():
    """Return a cached MAPI namespace, connecting to the running Outlook on first use."""
    global _namespace
    if _namespace is None:
        app = win32com.client.Dispatch("Outlook.Application")
        _namespace = app.GetNamespace("MAPI")
    return _namespace


def reset_namespace():
    """Drop the cached namespace so the next call reconnects from scratch."""
    global _namespace
    _namespace = None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fmt_dt(value):
    """Format a COM/pywintypes datetime as a readable string; fall back to str()."""
    try:
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value)


def sender_smtp(item):
    """Best-effort SMTP address of a sender; degrades to '' rather than prompting."""
    try:
        addr = item.SenderEmailAddress or ""
    except Exception:
        return ""
    if addr.upper().startswith("/O="):
        try:
            exch = item.Sender.GetExchangeUser()
            if exch is not None:
                return exch.PrimarySmtpAddress or addr
        except Exception:
            pass
    return addr


def parse_date(value, fallback):
    """Parse a YYYY-MM-DD string into a date; return fallback if value is empty."""
    if not value:
        return fallback
    return datetime.datetime.strptime(value, "%Y-%m-%d").date()


def _com_to_naive(value):
    """Convert a COM/pywintypes datetime to a naive Python datetime, or None."""
    try:
        return datetime.datetime(
            value.year, value.month, value.day,
            value.hour, value.minute, value.second,
        )
    except Exception:
        return None


def item_received(item):
    """The item's ReceivedTime as a naive datetime, or None."""
    try:
        return _com_to_naive(item.ReceivedTime)
    except Exception:
        return None


def item_best_datetime(item):
    """Best available date for an item: ReceivedTime, else SentOn, else None."""
    dt = item_received(item)
    if dt is None:
        try:
            dt = _com_to_naive(item.SentOn)
        except Exception:
            dt = None
    return dt


def _walk_folders(folder, prefix, out, depth=0):
    """Recursively collect (path, folder_object) pairs beneath `folder`."""
    if depth > 20:  # guard against pathological nesting / loops
        return
    try:
        subfolders = folder.Folders
    except Exception:
        return
    for sub in subfolders:
        try:
            name = sub.Name
        except Exception:
            continue
        path = prefix + "/" + name
        out.append((path, sub))
        _walk_folders(sub, path, out, depth + 1)


def all_mail_folders():
    """
    Return a list of (path, folder_object) for every folder across every store
    in the profile (main mailbox, online archive, mounted PSTs).
    """
    ns = get_namespace()
    out = []
    try:
        stores = ns.Stores
    except Exception:
        stores = None
    if stores is None:
        return out
    for store in stores:
        try:
            root = store.GetRootFolder()
            store_name = store.DisplayName
        except Exception:
            continue
        _walk_folders(root, store_name, out)
    return out


def find_target_folders(names_lower):
    """
    Return (label, folder) for folders whose LEAF name matches names_lower.
    Folders whose path matches the content blacklist are skipped entirely:
    results are labelled with their folder path, so a blacklisted folder name
    must never appear in output.
    """
    matched = []
    for path, folder in all_mail_folders():
        leaf = path.rsplit("/", 1)[-1].lower()
        if leaf in names_lower:
            reason = blacklisted_match(path)
            if reason:
                log("Skipped a search folder (blacklist match: {0}).".format(reason))
                continue
            matched.append((path, folder))
    return matched


# ---------------------------------------------------------------------------
# Tool implementations (each returns a human-readable text string)
# ---------------------------------------------------------------------------

def tool_list_recent_emails(args):
    count = int(args.get("count", 10))
    unread_only = bool(args.get("unread_only", False))

    ns = get_namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # True = descending (newest first)

    lines = []
    withheld = 0
    for item in items:
        if len(lines) >= count:
            break
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            if unread_only and not item.UnRead:
                continue
        except Exception:
            continue

        # Compliance filter: withhold blocked messages entirely.
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld an Inbox message (blacklist match: {0}).".format(reason))
            continue

        try:
            flag = "UNREAD" if item.UnRead else "read"
            lines.append(
                "- [{flag}] {received} | {sender}\n"
                "    Subject : {subject}\n"
                "    EntryID : {eid}".format(
                    flag=flag,
                    received=fmt_dt(item.ReceivedTime),
                    sender=(item.SenderName or "(unknown sender)"),
                    subject=(item.Subject or "(no subject)"),
                    eid=item.EntryID,
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} message(s) withheld by the content blacklist and not shown.]".format(withheld)

    if not lines:
        if withheld:
            return "No viewable messages. {0} message(s) were withheld by the content blacklist.".format(withheld)
        return "No matching messages found in the Inbox."
    header = "Showing {n} message(s) from the Inbox (newest first):".format(n=len(lines))
    return header + "\n" + "\n".join(lines) + note


def tool_search_emails(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required."
    count = int(args.get("count", 10))

    ns = get_namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items

    # DASL @SQL filter using LIKE '%...%'. String filters are NOT locale-sensitive.
    safe = query.replace("'", "''")
    dasl = (
        '@SQL=' +
        '"urn:schemas:httpmail:subject" LIKE \'%' + safe + '%\'' +
        ' OR ' +
        '"urn:schemas:httpmail:fromname" LIKE \'%' + safe + '%\''
    )

    try:
        matches = items.Restrict(dasl)
    except Exception as exc:
        return "Error: Outlook rejected the search filter ({0}).".format(exc)

    # Collect matches (bounded), then sort newest-first in Python.
    collected = []
    scanned = 0
    for item in matches:
        if scanned >= SEARCH_SCAN_CAP:
            break
        scanned += 1
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            collected.append(item)
        except Exception:
            continue

    def _received(it):
        try:
            return it.ReceivedTime
        except Exception:
            return None
    collected.sort(key=lambda it: (_received(it) is not None, _received(it)), reverse=True)

    lines = []
    withheld = 0
    for item in collected:
        if len(lines) >= count:
            break
        # Compliance filter.
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a search hit (blacklist match: {0}).".format(reason))
            continue
        try:
            lines.append(
                "- {received} | {sender}\n"
                "    Subject : {subject}\n"
                "    EntryID : {eid}".format(
                    received=fmt_dt(item.ReceivedTime),
                    sender=(item.SenderName or "(unknown sender)"),
                    subject=(item.Subject or "(no subject)"),
                    eid=item.EntryID,
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} matching message(s) withheld by the content blacklist.]".format(withheld)

    if not lines:
        if withheld:
            return "No viewable matches for '{0}'. {1} match(es) were withheld by the content blacklist.".format(query, withheld)
        return "No Inbox messages matched '{0}' (searched subject and sender name).".format(query)
    header = "Found {n} message(s) matching '{q}' (newest first):".format(n=len(lines), q=query)
    return header + "\n" + "\n".join(lines) + note


# ---------------------------------------------------------------------------
# Markdown export for the RAG knowledge base (mirrors confluence.py / word.py)
# ---------------------------------------------------------------------------

def safe_filename(name, max_len=120):
    """
    Turn an email subject into a filesystem-safe filename component (no
    extension). Strips characters illegal on Windows (< > : " / \\ | ? * and
    control chars), collapses whitespace, removes trailing dots/spaces (also
    illegal on Windows), and caps the length. Returns 'untitled' if nothing
    usable is left.
    """
    if not name:
        return "untitled"
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)
    cleaned = " ".join(cleaned.split()).strip(" .")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" .")
    return cleaned or "untitled"


def email_to_markdown(item, body):
    """
    Render a cleared email as Markdown: a metadata header (subject, sender,
    recipients, received time) followed by the plain-text body. The body is
    already plain text from Outlook, so no HTML conversion is needed.
    """
    def safe(getter, default=""):
        try:
            return getter() or default
        except Exception:
            return default

    smtp = sender_smtp(item)
    sender_line = safe(lambda: item.SenderName, "(unknown)")
    if smtp:
        sender_line += " <{0}>".format(smtp)

    header = (
        "# {subject}\n\n"
        "- From: {sender}\n"
        "- To: {to}\n"
        "- CC: {cc}\n"
        "- Received: {received}\n"
        "- Saved: {stamp}\n\n"
        "---\n\n"
    ).format(
        subject=safe(lambda: item.Subject, "(no subject)"),
        sender=sender_line,
        to=safe(lambda: item.To),
        cc=safe(lambda: item.CC),
        received=fmt_dt(safe(lambda: item.ReceivedTime)),
        stamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    )
    body = body or ""
    return header + (body if body.strip() else "(no body content)") + "\n"


def save_email_to_kb(item, body):
    """
    Write a cleared email to '<KB_DIR>/Email - <date> - <subject> (<id>).md',
    overwriting any existing file, for a local RAG knowledge base. Returns the
    path written; raises OSError on failure.

    The filename embeds the received date, a filesystem-safe subject and a short
    hash of the EntryID, so the same message overwrites its own file while two
    different messages that share a subject and date never collide.

    Only ever called for messages that have already cleared the content
    blacklist, so classified/withheld content is never written to disk.
    """
    def safe(getter, default=""):
        try:
            return getter() or default
        except Exception:
            return default

    subject = safe(lambda: item.Subject, "(no subject)")
    entry_id = safe(lambda: item.EntryID)
    received = safe(lambda: item.ReceivedTime, None)
    try:
        date_part = received.strftime("%Y-%m-%d") if received else "no-date"
    except Exception:
        date_part = "no-date"

    short = (hashlib.sha1(entry_id.encode("utf-8", "replace")).hexdigest()[:8]
             if entry_id else "nohash")
    filename = "Email - {0} - {1} ({2}).md".format(
        date_part, safe_filename(subject), short)

    os.makedirs(KB_DIR, exist_ok=True)
    path = os.path.join(KB_DIR, filename)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(email_to_markdown(item, body))
    return path


def tool_get_email(args):
    entry_id = (args.get("entry_id") or "").strip()
    if not entry_id:
        return "Error: 'entry_id' is required (get it from a list/search result)."

    ns = get_namespace()
    try:
        item = ns.GetItemFromID(entry_id)
    except Exception:
        return "Error: no message found for that EntryID (it may have moved or been deleted)."

    # Compliance filter: refuse blocked messages with a GENERIC message that does
    # not reveal the matched term or any classified content.
    reason = email_block_reason(item)
    if reason:
        log("Refused get_email (blacklist match: {0}) for EntryID {1}.".format(reason, entry_id))
        return ("This message cannot be displayed: its content is withheld under the "
                "content blacklist (classification / compliance policy).")

    try:
        full_body = item.Body or ""
    except Exception:
        return ("This message cannot be displayed: its body could not be read and "
                "therefore cannot be cleared for display.")
    # The full body is what gets saved to the knowledge base; only the copy
    # returned to the model is truncated (mirrors confluence.py's MAX_BODY rule).
    body = full_body
    truncated_note = ""
    if len(body) > MAX_BODY_CHARS:
        body = body[:MAX_BODY_CHARS]
        truncated_note = "\n\n[... body truncated at {0} characters ...]".format(MAX_BODY_CHARS)

    def safe(getter, default=""):
        try:
            return getter() or default
        except Exception:
            return default

    smtp = sender_smtp(item)
    sender_line = safe(lambda: item.SenderName, "(unknown)")
    if smtp:
        sender_line += " <{0}>".format(smtp)

    parts = [
        "Subject : {0}".format(safe(lambda: item.Subject, "(no subject)")),
        "From    : {0}".format(sender_line),
        "To      : {0}".format(safe(lambda: item.To)),
        "CC      : {0}".format(safe(lambda: item.CC)),
        "Received: {0}".format(fmt_dt(safe(lambda: item.ReceivedTime))),
        "",
        body + truncated_note,
    ]
    result_text = "\n".join(parts)

    # Save the (cleared) email to Markdown for RAG ingestion only if this call
    # asked for it, or the endpoint turned autosave on. Saving must never break
    # the read, so any failure is reported but swallowed. The FULL body is saved
    # (MAX_BODY_CHARS only trims the copy returned to the model).
    save_to_kb = args.get("save_to_kb")
    wanted = KB_AUTOSAVE if save_to_kb is None else bool(save_to_kb)
    if wanted and not KB_DIR:
        # A save was asked for that this endpoint has switched off. Say so, so
        # it cannot look done. (Autosave with no folder is a startup warning
        # instead - repeating it on every email would be noise.)
        if save_to_kb:
            result_text += (
                "\n\n[NOT saved: knowledge-base saving is switched off on this "
                "server. Unset OUTLOOK_KB_DIR (or point it at a folder "
                "inside the knowledge base) to enable it.]"
            )
    elif wanted:
        try:
            kb_path = save_email_to_kb(item, full_body)
            log("Saved email to knowledge base: {0}".format(kb_path))
            result_text += "\n\n[Saved to knowledge base: {0}]".format(kb_path)
        except OSError as exc:
            log("Knowledge-base save failed: {0}".format(exc))
            result_text += "\n\n[Knowledge-base save FAILED: {0}]".format(exc)
    return result_text


def tool_list_sent_emails(args):
    """List messages from the Sent Items folder within a date range (newest first)."""
    today = datetime.date.today()
    try:
        start_date = parse_date(args.get("start_date"), today - datetime.timedelta(days=7))
        end_date = parse_date(args.get("end_date"), today)
    except ValueError:
        return "Error: dates must be in YYYY-MM-DD format."
    if end_date < start_date:
        return "Error: end_date is before start_date."
    max_results = int(args.get("max_results", 100))

    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))

    ns = get_namespace()
    sent = ns.GetDefaultFolder(OL_FOLDER_SENT)
    items = sent.Items
    items.Sort("[SentOn]", True)  # True = descending (newest first)

    lines = []
    withheld = 0
    iterated = 0
    for item in items:
        if len(lines) >= max_results or iterated >= SENT_SCAN_CAP:
            break
        iterated += 1
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            sent_on = item.SentOn
            # Represent the COM time as a naive datetime for a day-range comparison.
            sent_naive = datetime.datetime(
                sent_on.year, sent_on.month, sent_on.day,
                sent_on.hour, sent_on.minute, sent_on.second,
            )
        except Exception:
            continue

        # Sorted newest-first: skip anything newer than the window; once we reach
        # an item older than the window, every later item is older too, so stop.
        if sent_naive > end_dt:
            continue
        if sent_naive < start_dt:
            break

        # The content blacklist applies to sent mail as well.
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a sent message (blacklist match: {0}).".format(reason))
            continue

        try:
            lines.append(
                "- {sent} | To: {to}\n"
                "    Subject : {subject}\n"
                "    EntryID : {eid}".format(
                    sent=fmt_dt(item.SentOn),
                    to=(item.To or "(no recipient)"),
                    subject=(item.Subject or "(no subject)"),
                    eid=item.EntryID,
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} sent message(s) withheld by the content blacklist.]".format(withheld)
    if not lines:
        if withheld:
            return "No viewable sent messages between {0} and {1}. {2} withheld by the content blacklist.".format(
                start_date, end_date, withheld)
        return "No sent messages between {0} and {1}.".format(start_date, end_date)
    header = "Sent messages {0} to {1} ({2} shown, newest first):".format(start_date, end_date, len(lines))
    return header + "\n" + "\n".join(lines) + note


def tool_list_folders(args):
    """List every mail folder across all stores, with item counts, for discovery/config."""
    max_lines = int(args.get("max_results", 300))
    folders = all_mail_folders()
    if not folders:
        return "No folders found (could not enumerate Outlook stores)."
    lines = []
    withheld = 0
    for path, folder in folders:
        if len(lines) >= max_lines:
            lines.append("... (list truncated; raise max_results to see more)")
            break
        # The content blacklist applies to folder NAMES too - a folder named
        # after a protective marking must not be revealed to the AI.
        reason = blacklisted_match(path)
        if reason:
            withheld += 1
            log("Withheld a folder name (blacklist match: {0}).".format(reason))
            continue
        try:
            count = folder.Items.Count
        except Exception:
            count = "?"
        lines.append("- {0}  (items: {1})".format(path, count))
    note = ""
    if withheld:
        note = "\n\n[{0} folder(s) withheld by the content blacklist.]".format(withheld)
    return "Outlook folders (store/path, with item counts):\n" + "\n".join(lines) + note


def tool_search_recent(args):
    """List mail across the configured folders (Inbox/Archive/Sent) within a date range."""
    today = datetime.date.today()
    try:
        start_date = parse_date(args.get("start_date"), today - datetime.timedelta(days=7))
        end_date = parse_date(args.get("end_date"), today)
    except ValueError:
        return "Error: dates must be in YYYY-MM-DD format."
    if end_date < start_date:
        return "Error: end_date is before start_date."
    query = (args.get("query") or "").strip().lower()
    max_results = int(args.get("max_results", 100))

    # Which folders to search. The caller may override the SEARCH_ALL_FOLDERS
    # default per call via a "folders" argument: either a JSON list of names or a
    # comma-separated string. Empty/omitted falls back to the configured default.
    folders_arg = args.get("folders")
    if isinstance(folders_arg, str):
        folder_names = [name.strip() for name in folders_arg.split(",") if name.strip()]
    elif isinstance(folders_arg, (list, tuple)):
        folder_names = [str(name).strip() for name in folders_arg if str(name).strip()]
    else:
        folder_names = []
    if not folder_names:
        folder_names = list(_SEARCH_FOLDERS)

    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))

    names_lower = set(name.lower() for name in folder_names)
    targets = find_target_folders(names_lower)
    if not targets:
        return ("No folders matched {0}. Run outlook_list_folders to see the folder "
                "names available, then adjust the 'folders' argument or "
                "SEARCH_ALL_FOLDERS.".format(folder_names))

    collected = []  # tuples of (naive_datetime, folder_label, item)
    withheld = 0
    for label, folder in targets:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # newest first (see break-early note below)
            sortable = True
        except Exception:
            items = folder.Items
            sortable = False

        scanned = 0
        for item in items:
            if scanned >= SENT_SCAN_CAP:
                break
            scanned += 1
            try:
                if item.Class != OL_CLASS_MAIL:
                    continue
            except Exception:
                continue

            received = item_received(item)
            when = item_best_datetime(item)
            if when is None:
                continue
            if when > end_dt:
                continue
            if when < start_dt:
                # Sorted by ReceivedTime desc: if the SORT KEY is past the window,
                # every later item is older too, so stop. Otherwise just skip (the
                # item's usable date came from a fallback field, so ordering is not
                # guaranteed and we must keep scanning).
                if sortable and received is not None and received < start_dt:
                    break
                continue

            # Optional free-text filter on subject / sender / recipient.
            if query:
                def _low(getter):
                    try:
                        return (getter() or "").lower()
                    except Exception:
                        return ""
                if (query not in _low(lambda: item.Subject)
                        and query not in _low(lambda: item.SenderName)
                        and query not in _low(lambda: item.To)):
                    continue

            # Content blacklist applies here too.
            reason = email_block_reason(item)
            if reason:
                withheld += 1
                log("Withheld an item from {0} (blacklist match: {1}).".format(label, reason))
                continue

            collected.append((when, label, item))

    collected.sort(key=lambda row: row[0], reverse=True)

    lines = []
    for when, label, item in collected[:max_results]:
        try:
            subject = item.Subject or "(no subject)"
        except Exception:
            subject = "(no subject)"
        who = ""
        try:
            who = item.SenderName or ""
        except Exception:
            who = ""
        if not who:  # sent items have no sender name; show recipient instead
            try:
                who = "To: " + (item.To or "")
            except Exception:
                who = ""
        try:
            eid = item.EntryID
        except Exception:
            eid = ""
        lines.append(
            "- {when} | {folder} | {who}\n"
            "    Subject : {subject}\n"
            "    EntryID : {eid}".format(
                when=when.strftime("%Y-%m-%d %H:%M"),
                folder=label,
                who=(who or "(unknown)"),
                subject=subject,
                eid=eid,
            )
        )

    note = ""
    if withheld:
        note += "\n\n[{0} item(s) withheld by the content blacklist.]".format(withheld)
    if len(collected) > max_results:
        note += "\n[Showing first {0} of {1} matches; narrow the dates or add a query.]".format(
            max_results, len(collected))

    if not lines:
        if withheld:
            return "No viewable emails between {0} and {1}. {2} withheld by the content blacklist.".format(
                start_date, end_date, withheld)
        return "No emails between {0} and {1} across {2}.".format(
            start_date, end_date, ", ".join(folder_names))
    header = "Emails {0} to {1} across {2} ({3} shown, newest first):".format(
        start_date, end_date, ", ".join(folder_names), len(lines))
    return header + "\n" + "\n".join(lines) + note


def _expand_recurring_occurrences(item, start_dt, end_dt):
    """All occurrences of a recurring appointment starting in [start_dt, end_dt].

    Probes the series' RecurrencePattern day by day at the series' usual start
    time (GetOccurrence raises for days with no/deleted occurrence, which is
    normal and simply skipped), then adds rescheduled occurrences from the
    pattern's Exceptions, since those no longer sit at the usual time. Pure
    COM object access - no locale-sensitive date strings anywhere.

    Returns a list of (naive start datetime, AppointmentItem occurrence).
    """
    try:
        pattern = item.GetRecurrencePattern()
    except Exception:
        return []

    master_start = _com_to_naive(item.Start)
    base_time = master_start.time() if master_start else datetime.time(0, 0)

    found = {}  # start datetime -> occurrence (dedups probe vs exception hits)
    day = start_dt.date()
    while day <= end_dt.date():
        probe = datetime.datetime.combine(day, base_time)
        try:
            occ = pattern.GetOccurrence(probe)
            occ_start = _com_to_naive(occ.Start)
            if occ_start is not None and start_dt <= occ_start <= end_dt:
                found[occ_start] = occ
        except Exception:
            pass  # no occurrence on this day
        day = day + datetime.timedelta(days=1)

    try:
        for exception in pattern.Exceptions:
            try:
                if bool(exception.Deleted):
                    continue
                occ = exception.AppointmentItem
                occ_start = _com_to_naive(occ.Start)
                if occ_start is not None and start_dt <= occ_start <= end_dt:
                    found[occ_start] = occ
            except Exception:
                continue
    except Exception:
        pass

    return list(found.items())


def scan_calendar(start_dt, end_dt):
    """
    Walk the whole calendar folder and return every occurrence inside a window.

    Returns (matches, total, debug_tail): `matches` is a list of (naive start
    datetime, appointment COM object) sorted by start; `total` is how many items
    the folder holds; `debug_tail` is a rolling last-5 of
    (index, start, matched?) so an empty result can prove what was scanned.

    NO Outlook-side filtering happens here. Restrict() formats its date strings
    per the machine's regional settings (US vs AU) and misbehaves with recurring
    appointments, silently returning 0 results on some machines. Instead the
    whole collection is walked by index (Item(i) is more reliable than COM
    enumeration) and filtered in Python, which is locale-independent.
    """
    ns = get_namespace()
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items
    total = int(items.Count)  # raises if the folder cannot be read; caller reports

    matches = []      # (naive start datetime, appointment COM object)
    debug_tail = []   # rolling last-5 raw items: (index, start-or-None, matched?)

    for i in range(1, total + 1):
        try:
            item = items.Item(i)
        except Exception:
            continue

        try:
            item_start = _com_to_naive(item.Start)
        except Exception:
            item_start = None

        try:
            is_recurring = bool(item.IsRecurring)
        except Exception:
            is_recurring = False

        if is_recurring:
            # A recurring master's own Start is its FIRST occurrence, usually
            # far outside the window - expand the series instead of testing it.
            occurrences = _expand_recurring_occurrences(item, start_dt, end_dt)
            matched = bool(occurrences)
            matches.extend(occurrences)
        else:
            matched = item_start is not None and start_dt <= item_start <= end_dt
            if matched:
                matches.append((item_start, item))

        # Rolling tail for the debug section shown when nothing matches.
        # Start date + flag only; subjects are deliberately omitted so the
        # content blacklist cannot leak through this path.
        debug_tail.append((i, item_start, matched))
        if len(debug_tail) > 5:
            debug_tail.pop(0)

        if len(matches) >= CALENDAR_HARD_CAP:
            break

    # Outlook's Sort() is no longer used; order in Python instead.
    matches.sort(key=lambda pair: pair[0])
    return matches, total, debug_tail


def tool_get_calendar(args):
    today = datetime.date.today()
    try:
        start_date = parse_date(args.get("start_date"), today)
        end_date = parse_date(args.get("end_date"), today + datetime.timedelta(days=7))
    except ValueError:
        return "Error: dates must be in YYYY-MM-DD format."

    if end_date < start_date:
        return "Error: end_date is before start_date."

    max_results = int(args.get("max_results", 50))

    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))

    try:
        matches, total, debug_tail = scan_calendar(start_dt, end_dt)
    except Exception as exc:
        return "Error: could not read the calendar folder ({0}).".format(exc)

    lines = []
    withheld = 0
    for item_start, item in matches:
        if len(lines) >= max_results:
            break

        # Compliance filter.
        reason = appointment_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a calendar item (blacklist match: {0}).".format(reason))
            continue

        try:
            all_day = bool(item.AllDayEvent)
            when = "{start} -> {end}".format(start=fmt_dt(item.Start), end=fmt_dt(item.End))
            if all_day:
                when = "{0} (all day)".format(item_start.strftime("%Y-%m-%d"))
            recur = " [recurring]" if bool(item.IsRecurring) else ""
            location = ""
            try:
                location = item.Location or ""
            except Exception:
                pass
            organizer = ""
            try:
                organizer = item.Organizer or ""
            except Exception:
                pass
            lines.append(
                "- {when}{recur}\n"
                "    Subject  : {subject}\n"
                "    Location : {loc}\n"
                "    Organizer: {org}".format(
                    when=when,
                    recur=recur,
                    subject=(item.Subject or "(no subject)"),
                    loc=(location or "-"),
                    org=(organizer or "-"),
                )
            )
        except Exception:
            continue

    note = ""
    if withheld:
        note = "\n\n[{0} event(s) withheld by the content blacklist and not shown.]".format(withheld)

    if not lines:
        if withheld:
            return "No viewable events between {0} and {1}. {2} event(s) were withheld by the content blacklist.".format(
                start_date, end_date, withheld)
        # Debug section: prove what was actually scanned, so a filtering
        # regression is visible instead of looking like an empty calendar.
        msg = "No calendar events between {0} and {1}.".format(start_date, end_date)
        msg += "\n\n[debug] Calendar folder holds {0} item(s); {1} matched the date range.".format(
            total, len(matches))
        if debug_tail:
            msg += "\n[debug] Last {0} item(s) in the folder (start / in range?):".format(
                len(debug_tail))
            for idx, dbg_start, dbg_matched in debug_tail:
                msg += "\n  - item {0}: start={1}, match={2}".format(
                    idx,
                    dbg_start.strftime("%Y-%m-%d %H:%M") if dbg_start else "(unreadable)",
                    "yes" if dbg_matched else "no")
        return msg
    header = "Calendar events {0} to {1} ({2} shown):".format(start_date, end_date, len(lines))
    return header + "\n" + "\n".join(lines) + note


# ---------------------------------------------------------------------------
# Printable day planner - a minimal PDF writer (standard library only)
# ---------------------------------------------------------------------------
# There is no PDF library here on purpose: this endpoint is airgapped and the
# only pip dependency this server has is pywin32, which it needs for COM. The
# planner is rectangles, rules and single lines of text, all of which the PDF
# imaging model does directly, so a few hundred lines of writer is cheaper than
# another package to transfer and keep installed.
#
# Type faces are the base-14 fonts every PDF reader carries (Helvetica and
# Helvetica-Bold), so nothing is embedded and the file opens identically on any
# machine. The design this follows uses Manrope; Helvetica is the closest
# stand-in available without shipping a font file.

PDF_MM = 72.0 / 25.4  # points per millimetre

# Glyph widths (1/1000 em) for codes 32-126 of Helvetica and Helvetica-Bold,
# taken from the Adobe AFM metrics. Needed so text can be measured, and
# therefore truncated with an ellipsis, without a font library.
_HELV_WIDTHS = (
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
    1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
    333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
    556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
)
_HELV_BOLD_WIDTHS = (
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278,
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,
    975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778,
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,
    333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
    611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
)
# Punctuation outside ASCII that the planner actually emits. Everything else
# non-ASCII falls back to its unaccented base letter, whose width is identical
# in these two faces (eacute is exactly as wide as e).
_PDF_EXTRA_WIDTHS = {
    "…": (1000, 1000),  # ellipsis - used by truncation
    "‘": (222, 238),    # left single quote
    "’": (222, 238),    # right single quote / apostrophe
    "“": (333, 500),    # left double quote
    "”": (333, 500),    # right double quote
    "•": (350, 350),    # bullet
    "–": (556, 556),    # en dash
    "—": (1000, 1000),  # em dash
    "·": (278, 278),    # middle dot - the separator used throughout
    " ": (278, 278),    # non-breaking space
}


def _one_line(value):
    """Collapse any string to a single line of text fit to draw."""
    if value is None:
        return ""
    return " ".join(str(value).split())


def _pdf_sanitise(text):
    """
    Reduce text to characters WinAnsiEncoding can represent.

    The base-14 fonts are WinAnsi (cp1252), so a subject containing CJK, emoji
    or symbols has to be transliterated rather than dropped mid-render. Accented
    Latin letters survive as themselves; anything else degrades to its base
    letter, and failing that to '?'.
    """
    try:
        text.encode("cp1252")
        return text
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    out = []
    for ch in text:
        try:
            ch.encode("cp1252")
            out.append(ch)
        except (UnicodeEncodeError, UnicodeDecodeError):
            base = "".join(part for part in unicodedata.normalize("NFD", ch)
                           if not unicodedata.combining(part))
            try:
                out.append(base.encode("cp1252").decode("cp1252") or "?")
            except (UnicodeEncodeError, UnicodeDecodeError):
                out.append("?")
    return "".join(out)


def _pdf_char_width(ch, bold):
    """Width of one character in 1/1000 em, for the chosen weight."""
    table = _HELV_BOLD_WIDTHS if bold else _HELV_WIDTHS
    code = ord(ch)
    if 32 <= code <= 126:
        return table[code - 32]
    extra = _PDF_EXTRA_WIDTHS.get(ch)
    if extra:
        return extra[1 if bold else 0]
    # Accented Latin: identical width to the unaccented letter in both faces.
    base = "".join(part for part in unicodedata.normalize("NFD", ch)
                   if not unicodedata.combining(part))
    if base and 32 <= ord(base[0]) <= 126:
        return table[ord(base[0]) - 32]
    return table[ord("n") - 32]


def _pdf_escape(text):
    """Escape a sanitised string into PDF literal-string bytes."""
    raw = text.encode("cp1252", "replace")
    return (raw.replace(b"\\", b"\\\\")
               .replace(b"(", b"\\(")
               .replace(b")", b"\\)"))


class PdfCanvas:
    """
    A single-page PDF drawn from rectangles, rules and lines of text.

    Coordinates are MILLIMETRES FROM THE TOP-LEFT corner, because that is how a
    page design is described; the PDF's own bottom-left origin in points is an
    internal detail of this class. Font sizes are in points, as in any layout.
    """

    def __init__(self, width_mm, height_mm, title="", author=""):
        self.width_mm = float(width_mm)
        self.height_mm = float(height_mm)
        self.title = title
        self.author = author
        self._ops = []

    # -- internals ----------------------------------------------------------

    def _emit(self, op):
        self._ops.append(op.encode("ascii", "replace") if isinstance(op, str) else op)

    @staticmethod
    def _num(value):
        text = "{0:.3f}".format(float(value)).rstrip("0").rstrip(".")
        return text if text not in ("", "-") else "0"

    @classmethod
    def _rgb(cls, colour):
        value = colour.lstrip("#")
        return " ".join(cls._num(int(value[i:i + 2], 16) / 255.0) for i in (0, 2, 4))

    def _px(self, x_mm):
        return x_mm * PDF_MM

    def _py(self, y_mm):
        """mm from the top -> points from the bottom, which is what PDF wants."""
        return (self.height_mm - y_mm) * PDF_MM

    def _rounded_path(self, x, y, w, h, radius):
        x0, x1 = self._px(x), self._px(x + w)
        y1, y0 = self._py(y), self._py(y + h)
        r = radius * PDF_MM
        k = r * 0.5523
        n = self._num
        parts = [
            "{0} {1} m".format(n(x0 + r), n(y0)),
            "{0} {1} l".format(n(x1 - r), n(y0)),
            "{0} {1} {2} {3} {4} {5} c".format(
                n(x1 - r + k), n(y0), n(x1), n(y0 + r - k), n(x1), n(y0 + r)),
            "{0} {1} l".format(n(x1), n(y1 - r)),
            "{0} {1} {2} {3} {4} {5} c".format(
                n(x1), n(y1 - r + k), n(x1 - r + k), n(y1), n(x1 - r), n(y1)),
            "{0} {1} l".format(n(x0 + r), n(y1)),
            "{0} {1} {2} {3} {4} {5} c".format(
                n(x0 + r - k), n(y1), n(x0), n(y1 - r + k), n(x0), n(y1 - r)),
            "{0} {1} l".format(n(x0), n(y0 + r)),
            "{0} {1} {2} {3} {4} {5} c".format(
                n(x0), n(y0 + r - k), n(x0 + r - k), n(y0), n(x0 + r), n(y0)),
            "h",
        ]
        return " ".join(parts)

    # -- drawing ------------------------------------------------------------

    def rect(self, x, y, w, h, fill=None, stroke=None, line_pt=0.6, radius=0.0):
        """Filled and/or stroked rectangle; `radius` in mm rounds the corners."""
        if w <= 0 or h <= 0 or (fill is None and stroke is None):
            return
        radius = max(0.0, min(float(radius), w / 2.0, h / 2.0))
        if radius > 0:
            path = self._rounded_path(x, y, w, h, radius)
        else:
            n = self._num
            path = "{0} {1} {2} {3} re".format(
                n(self._px(x)), n(self._py(y + h)),
                n(w * PDF_MM), n(h * PDF_MM))

        ops = ["q"]
        if fill is not None:
            ops.append("{0} rg".format(self._rgb(fill)))
        if stroke is not None:
            ops.append("{0} RG".format(self._rgb(stroke)))
            ops.append("{0} w".format(self._num(line_pt)))
        ops.append(path)
        if fill is not None and stroke is not None:
            ops.append("B")
        elif fill is not None:
            ops.append("f")
        else:
            ops.append("S")
        ops.append("Q")
        self._emit(" ".join(ops))

    def line(self, x1, y1, x2, y2, colour="#E5E7EB", line_pt=0.6, dash=None):
        """Straight rule. `dash` is a list of on/off lengths in points."""
        n = self._num
        ops = ["q", "{0} RG".format(self._rgb(colour)), "{0} w".format(n(line_pt))]
        if dash:
            ops.append("[{0}] 0 d".format(" ".join(n(value) for value in dash)))
        ops.append("{0} {1} m {2} {3} l S".format(
            n(self._px(x1)), n(self._py(y1)), n(self._px(x2)), n(self._py(y2))))
        ops.append("Q")
        self._emit(" ".join(ops))

    def text_width(self, string, size, bold=False, tracking=0.0):
        """Width of one line of text, in mm, at `size` points."""
        string = _pdf_sanitise(_one_line(string))
        if not string:
            return 0.0
        total = sum(_pdf_char_width(ch, bold) for ch in string) / 1000.0 * size
        total += len(string) * tracking  # PDF adds Tc after every glyph
        return total / PDF_MM

    def truncate(self, string, size, max_width, bold=False, tracking=0.0):
        """Shorten text with a trailing ellipsis so it fits `max_width` mm."""
        string = _pdf_sanitise(_one_line(string))
        if not string or self.text_width(string, size, bold, tracking) <= max_width:
            return string
        budget = max_width - self.text_width("…", size, bold, tracking)
        if budget <= 0:
            return ""
        used = 0.0
        cut = 0
        for index, ch in enumerate(string):
            step = (_pdf_char_width(ch, bold) / 1000.0 * size + tracking) / PDF_MM
            if used + step > budget:
                break
            used += step
            cut = index + 1
        trimmed = string[:cut].rstrip(" ·-–")
        return (trimmed + "…") if trimmed else ""

    def text(self, x, y, string, size, bold=False, colour="#111827",
             tracking=0.0, align="left", max_width=None):
        """
        Draw one line of text. `y` is the BASELINE in mm from the page top.

        `align` positions the line against `x` ("left", "right" or "centre");
        `max_width` truncates it with an ellipsis first. Returns the width drawn.
        """
        string = _pdf_sanitise(_one_line(string))
        if max_width is not None:
            string = self.truncate(string, size, max_width, bold, tracking)
        if not string:
            return 0.0

        width = self.text_width(string, size, bold, tracking)
        if align == "right":
            x = x - width
        elif align in ("centre", "center"):
            x = x - width / 2.0

        n = self._num
        pieces = [
            b"q ",
            "{0} rg BT /{1} {2} Tf".format(
                self._rgb(colour), "F2" if bold else "F1", n(size)).encode("ascii"),
        ]
        if tracking:
            pieces.append(" {0} Tc".format(n(tracking)).encode("ascii"))
        pieces.append(" 1 0 0 1 {0} {1} Tm ".format(
            n(self._px(x)), n(self._py(y))).encode("ascii"))
        pieces.append(b"(" + _pdf_escape(string) + b") Tj ET Q")
        self._emit(b"".join(pieces))
        return width

    def wrap(self, string, size, max_width, bold=False, max_lines=2):
        """Word-wrap one string into at most `max_lines` lines that each fit."""
        string = _pdf_sanitise(_one_line(string))
        if not string:
            return []
        lines = []
        current = ""
        for word in string.split(" "):
            candidate = (current + " " + word).strip()
            if current and self.text_width(candidate, size, bold) > max_width:
                lines.append(current)
                current = word
                if len(lines) == max_lines:
                    break
            else:
                current = candidate
        if len(lines) < max_lines and current:
            lines.append(current)
        elif len(lines) == max_lines and current:
            # More text than the allowance: fold the tail into the last line so
            # truncation puts an honest ellipsis on the end of it.
            lines[-1] = self.truncate(
                (lines[-1] + " " + current).strip(), size, max_width, bold)
        return [self.truncate(line, size, max_width, bold) for line in lines]

    def text_block(self, x, y, string, size, max_width, bold=False,
                   colour="#111827", leading=None, max_lines=2):
        """
        Word-wrap `string` and draw it. `y` is the baseline of the FIRST line.

        Returns the baseline of the last line drawn, so a caller can carry on
        underneath it.
        """
        leading = leading if leading is not None else size * 1.16 / PDF_MM
        lines = self.wrap(string, size, max_width, bold=bold, max_lines=max_lines)
        baseline = y
        for index, line in enumerate(lines):
            baseline = y + index * leading
            self.text(x, baseline, line, size, bold=bold, colour=colour)
        return baseline

    # -- output -------------------------------------------------------------

    def to_bytes(self):
        """Serialise the page into a complete PDF file."""
        content = zlib.compress(b"\n".join(self._ops))
        stamp = datetime.datetime.now()
        try:
            offset = stamp.astimezone().utcoffset()
            minutes = int(offset.total_seconds() // 60)
            sign = "+" if minutes >= 0 else "-"
            tz = "{0}{1:02d}'{2:02d}'".format(sign, abs(minutes) // 60, abs(minutes) % 60)
        except Exception:
            tz = "Z"
        created = "D:{0}{1}".format(stamp.strftime("%Y%m%d%H%M%S"), tz)

        def literal(value):
            return b"(" + _pdf_escape(_pdf_sanitise(_one_line(value))) + b")"

        objects = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0} {1}] "
             "/Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> "
             "/Contents 4 0 R >>").format(
                self._num(self.width_mm * PDF_MM),
                self._num(self.height_mm * PDF_MM)).encode("ascii"),
            (b"<< /Length " + str(len(content)).encode("ascii") +
             b" /Filter /FlateDecode >>\nstream\n" + content + b"\nendstream"),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
            (b"<< /Title " + literal(self.title) +
             b" /Author " + literal(self.author) +
             b" /Creator (outlook-mcp " + __version__.encode("ascii") + b")" +
             b" /Producer (outlook-mcp " + __version__.encode("ascii") + b")" +
             b" /CreationDate (" + created.encode("ascii") + b") >>"),
        ]

        out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = []
        for number, body in enumerate(objects, start=1):
            offsets.append(len(out))
            out += "{0} 0 obj\n".format(number).encode("ascii")
            out += body
            out += b"\nendobj\n"

        xref_at = len(out)
        out += "xref\n0 {0}\n".format(len(objects) + 1).encode("ascii")
        out += b"0000000000 65535 f \n"
        for position in offsets:
            out += "{0:010d} 00000 n \n".format(position).encode("ascii")
        out += ("trailer\n<< /Size {0} /Root 1 0 R /Info {1} 0 R >>\n"
                "startxref\n{2}\n%%EOF\n").format(
                    len(objects) + 1, len(objects), xref_at).encode("ascii")
        return bytes(out)


# ---------------------------------------------------------------------------
# Printable day planner - laying the calendar out on the page
# ---------------------------------------------------------------------------

# Month and weekday names are spelled out here rather than taken from
# strftime(): %A and %B follow the machine's locale, and this server already
# refuses to let regional settings change what the calendar looks like.
_WEEKDAY_NAMES = ("Monday", "Tuesday", "Wednesday", "Thursday",
                  "Friday", "Saturday", "Sunday")
_WEEKDAY_SHORT = ("MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN")
_MONTH_NAMES = ("January", "February", "March", "April", "May", "June", "July",
                "August", "September", "October", "November", "December")
_MONTH_SHORT = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL",
                "AUG", "SEP", "OCT", "NOV", "DEC")
_NUMBER_WORDS = {2: "two", 3: "three", 4: "four", 5: "five", 6: "six"}

# One entry per colour an event block can take. 'fill' is the solid block,
# 'text'/'sub' the type on it, and 'edge' the colour used when a block is drawn
# hollow (tentative or not-yet-accepted meetings).
# The planner palette: one BOLD hex per colour name, deliberately saturated so a
# block reads across a printed page and a photocopy. The names line up with
# Outlook's own 25 category colours (see OL_CATEGORY_COLOURS below), which are
# muted on screen; printing them at this strength is the point - it keeps your
# own colour-coding, just loud enough to be worth printing.
#
# 'grey' is the odd one out and stays near-white: it is the quiet block a
# personal appointment with nobody invited gets, not Outlook's Gray category
# (that is 'silver').
PLANNER_COLOURS = {
    "red": "#EF4444",
    "dark-red": "#B91C1C",
    "orange": "#F97316",
    "dark-orange": "#C2410C",
    "peach": "#FDBA74",
    "dark-peach": "#FB923C",
    "yellow": "#FACC15",
    "dark-yellow": "#CA8A04",
    "amber": "#F59E0B",
    "green": "#10B981",
    "dark-green": "#047857",
    "teal": "#0D9488",
    "dark-teal": "#0F766E",
    "olive": "#84CC16",
    "dark-olive": "#4D7C0F",
    "blue": "#3B82F6",
    "dark-blue": "#1D4ED8",
    "purple": "#8B5CF6",
    "dark-purple": "#6D28D9",
    "maroon": "#9F1239",
    "dark-maroon": "#881337",
    "rose": "#F43F5E",
    "steel": "#64748B",
    "dark-steel": "#475569",
    "silver": "#9CA3AF",
    "slate": "#6B7280",
    "black": "#1F2937",
    "grey": "#F3F4F6",
    "gray": "#F3F4F6",      # same block, spelled either way
}

# Outlook's OlCategoryColor enum -> the palette name above. This is what turns
# "the category is Outlook green" into a bold green block: Outlook's own colour
# for a category is read straight off the profile, so a diary already colour-
# coded in Outlook prints in the same scheme without configuring anything.
# 0 (olCategoryColorNone) is absent on purpose: a category with no colour set
# falls through to the next one on the appointment, then to the built-in split.
OL_CATEGORY_COLOURS = {
    1: "red",          2: "orange",       3: "peach",       4: "yellow",
    5: "green",        6: "teal",         7: "olive",       8: "blue",
    9: "purple",      10: "maroon",      11: "steel",      12: "dark-steel",
    13: "silver",     14: "slate",       15: "black",      16: "dark-red",
    17: "dark-orange", 18: "dark-peach", 19: "dark-yellow", 20: "dark-green",
    21: "dark-teal",  22: "dark-olive",  23: "dark-blue",  24: "dark-purple",
    25: "dark-maroon",
}

# Fill above which type has to go dark instead of white. Tuned so orange and
# both yellows take dark type (white on any of them is unreadable in print)
# while green, teal and blue keep the white the design uses.
PLANNER_LIGHT_ABOVE = 0.54

_palette_cache = {}


def _hex_rgb(colour):
    value = colour.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def _mix(colour, towards, amount):
    """Blend `colour` `amount` of the way towards another colour, as hex."""
    source, target = _hex_rgb(colour), _hex_rgb(towards)
    return "#{0:02X}{1:02X}{2:02X}".format(*[
        int(round(source[i] + (target[i] - source[i]) * amount)) for i in range(3)
    ])


def _brightness(colour):
    """Perceived brightness, 0 (black) to 1 (white)."""
    red, green, blue = _hex_rgb(colour)
    return (0.299 * red + 0.587 * green + 0.114 * blue) / 255.0


def planner_palette(colour_key):
    """
    The four colours one named block needs, derived from its single hex.

    'fill' is the block, 'text' the subject on it, 'sub' the time/location/people
    line, and 'edge' the darker version used for a hollow block's border and
    type. Deriving them beats writing four hex values per colour by hand: every
    colour in the palette then behaves the same way, and a light fill such as
    yellow gets dark type automatically instead of white-on-white.
    """
    if colour_key in _palette_cache:
        return _palette_cache[colour_key]

    fill = PLANNER_COLOURS.get(colour_key) or PLANNER_COLOURS["blue"]
    light = _brightness(fill) > PLANNER_LIGHT_ABOVE
    palette = {
        "fill": fill,
        "light": light,
        # On a dark fill, white type with a pale tint under it. On a light one,
        # two shades of the fill itself, which keeps the block on-hue.
        "text": _mix(fill, "#000000", 0.78) if light else "#FFFFFF",
        "sub": _mix(fill, "#000000", 0.56) if light else _mix(fill, "#FFFFFF", 0.84),
        # A hollow block puts 'edge' on white, so a pale fill has to be taken
        # much further down or its outline and title vanish on the page.
        "edge": _mix(fill, "#000000", 0.66 if light else 0.34),
        # A near-white block needs an outline or it floats off the page.
        "border": _mix(fill, "#000000", 0.10) if light else None,
    }
    _palette_cache[colour_key] = palette
    return palette


# Accent colour for each day down the right-hand panel, purely so the four days
# are easy to tell apart at a glance.
PLANNER_DAY_ACCENTS = ("blue", "purple", "green", "amber")

# Page ink.
_INK = "#111827"
_MUTED = "#6B7280"
_FAINT = "#9CA3AF"
_RULE = "#E5E7EB"
_HAIRLINE = "#F3F4F6"
_GRID = "#D1D5DB"

# Outlook's OlBusyStatus values.
OL_BUSY_FREE = 0
OL_BUSY_TENTATIVE = 1
OL_BUSY_OUT_OF_OFFICE = 3

CAP_RATIO = 0.716  # Helvetica cap height, as a fraction of the point size


def _cap_mm(size_pt):
    """Cap height of `size_pt` type, in mm - what centring text vertically needs."""
    return size_pt * CAP_RATIO / PDF_MM


def _fmt_day_long(day):
    """'Monday 14 September 2026' - Australian order, no ordinal, no comma."""
    return "{0} {1} {2} {3}".format(
        _WEEKDAY_NAMES[day.weekday()], day.day,
        _MONTH_NAMES[day.month - 1], day.year)


def _fmt_day_short(day):
    """'14 Sep'."""
    return "{0} {1}".format(day.day, _MONTH_SHORT[day.month - 1].title())


def _fmt_time(value):
    """24-hour clock, as the printed grid uses throughout."""
    return "{0:02d}:{1:02d}".format(value.hour, value.minute)


def _fmt_duration(minutes):
    """'5h 45m', '45m', '6h'."""
    hours, mins = divmod(int(max(0, minutes)), 60)
    if hours and mins:
        return "{0}h {1}m".format(hours, mins)
    if hours:
        return "{0}h".format(hours)
    return "{0}m".format(mins)


def _minutes_into_day(value, day):
    """Minutes from midnight on `day`, clamped either side for multi-day events."""
    delta = value - datetime.datetime.combine(day, datetime.time(0, 0))
    return delta.total_seconds() / 60.0


# ---------------------------------------------------------------------------
# Reading one appointment into a plain dict the layout can work from
# ---------------------------------------------------------------------------

def _appointment_people(item, identity, cap=40):
    """
    Attendee display names, and whether anyone is outside our own domain.

    `identity` is the (name, address) of the mailbox owner: they are dropped
    from the list, because a printed planner does not need to tell us we are in
    our own meetings.

    Deliberately reads only Recipients.Address, never AddressEntry: resolving an
    address entry is slow and can touch the directory, and the '/O=' prefix of
    an Exchange address already says 'internal' without asking anyone.
    """
    own_name, own_address = identity
    own_domain = (own_address.rsplit("@", 1)[-1].lower()
                  if "@" in own_address else "")
    names = []
    external = False
    try:
        recipients = item.Recipients
        count = int(recipients.Count)
    except Exception:
        return names, external

    for index in range(1, min(count, cap) + 1):
        try:
            recipient = recipients.Item(index)
        except Exception:
            continue
        try:
            kind = int(recipient.Type)
        except Exception:
            kind = 1
        if kind == 3:  # olResource - a meeting room, not a person
            continue
        try:
            name = _one_line(recipient.Name or "")
        except Exception:
            name = ""
        try:
            address = (recipient.Address or "").strip()
        except Exception:
            address = ""

        if own_domain and address and "@" in address:
            domain = address.rsplit("@", 1)[-1].lower()
            if domain != own_domain:
                external = True

        if not name:
            continue
        is_self = ((own_address and address.lower() == own_address.lower())
                   or (own_name and name.lower() == own_name.lower()))
        if not is_self:
            names.append(name)
    return names, external


def _appointment_to_event(item, occurrence_start, identity, want_people=True):
    """
    Flatten one Outlook appointment into the dict the planner layout uses.

    Every field is read defensively: a calendar full of imported or malformed
    items must not stop the page being printed.
    """
    def safe(getter, default=""):
        try:
            value = getter()
            return default if value is None else value
        except Exception:
            return default

    start = occurrence_start or _com_to_naive(safe(lambda: item.Start, None))
    end = _com_to_naive(safe(lambda: item.End, None))
    if start is None:
        return None
    if end is None or end <= start:
        minutes = 30
        try:
            minutes = int(item.Duration) or 30
        except Exception:
            pass
        end = start + datetime.timedelta(minutes=minutes)

    categories = [part.strip() for part in
                  str(safe(lambda: item.Categories, "")).split(";") if part.strip()]
    try:
        busy = int(item.BusyStatus)
    except Exception:
        busy = 2

    people, external = ([], False)
    if want_people:
        people, external = _appointment_people(item, identity)

    return {
        "start": start,
        "end": end,
        "all_day": bool(safe(lambda: item.AllDayEvent, False)),
        "subject": _one_line(safe(lambda: item.Subject, "")) or "(no subject)",
        "location": _one_line(safe(lambda: item.Location, "")),
        "organiser": _one_line(safe(lambda: item.Organizer, "")),
        "categories": categories,
        "busy": busy,
        "people": people,
        "external": external,
        "recurring": bool(safe(lambda: item.IsRecurring, False)),
    }


def _event_bucket(event, category_colours):
    """
    The legend bucket an event belongs to: (label, colour key).

    An Outlook category the endpoint has mapped to a colour wins, because that
    is the user's own classification of their diary. Failing that the split is
    one Outlook can actually prove: somebody outside our domain is on it, it is
    an internal meeting, or it is time blocked out with nobody invited.
    """
    for category in event["categories"]:
        colour = category_colours.get(category.lower())
        if colour:
            return category, colour
    if event["external"]:
        return "External", "green"
    if event["people"]:
        return "Internal", "blue"
    return "Personal", "grey"


def _bucket_rank(entry):
    """Legend order: mapped categories first, then the built-in three."""
    label = entry[0]
    builtin = {"External": 1, "Internal": 2, "Personal": 3}
    return (builtin.get(label, 0), label.lower())


# ---------------------------------------------------------------------------
# Timeline geometry
# ---------------------------------------------------------------------------

PLANNER_MAX_LEGEND = 5       # colour chips across the header, before Tentative
PLANNER_MIN_BLOCK_MM = 6.4   # a short meeting is grown to this if there is room
PLANNER_MIN_TIGHT_MM = 3.4   # ...and never squeezed below this, so a line still fits
PLANNER_BLOCK_GAP_MM = 1.0   # gutter between two blocks, side by side or stacked


def _place_blocks(events, day, start_min, end_min, top, bottom, lane_x, lane_w):
    """
    Work out where each event's block goes on the timeline.

    Two rules do the work. Events that genuinely overlap share the lane
    Outlook-style, splitting a cluster into as many columns as it needs. And a
    short meeting is grown to a readable height, but only into space nothing
    else is using: a day of back-to-back half-hour meetings stays one honest
    column of thin blocks instead of zig-zagging across two.
    """
    scale = (bottom - top) / float(max(1, end_min - start_min))
    boxes = []
    for event in events:
        from_min = max(start_min, _minutes_into_day(event["start"], day))
        to_min = max(from_min, min(end_min, _minutes_into_day(event["end"], day)))
        boxes.append({
            "event": event,
            "from_min": from_min,
            "to_min": to_min,
            "top": top + (from_min - start_min) * scale,
            "natural": (to_min - from_min) * scale,
        })
    boxes.sort(key=lambda box: (box["from_min"], -box["to_min"]))

    # Clusters are built from the REAL times. Two meetings that only look
    # adjacent once a 15-minute slot has been grown to a readable height are not
    # a clash, and must not be pushed into separate columns as though they were.
    clusters = []
    for box in boxes:
        if clusters and box["from_min"] < clusters[-1]["end"] - 0.01:
            clusters[-1]["boxes"].append(box)
            clusters[-1]["end"] = max(clusters[-1]["end"], box["to_min"])
        else:
            clusters.append({"boxes": [box], "end": box["to_min"]})

    for cluster in clusters:
        columns = []
        for box in cluster["boxes"]:
            for index, column in enumerate(columns):
                if box["from_min"] >= column[-1]["to_min"] - 0.01:
                    column.append(box)
                    box["column"] = index
                    break
            else:
                columns.append([box])
                box["column"] = len(columns) - 1

        width = ((lane_w - PLANNER_BLOCK_GAP_MM * (len(columns) - 1))
                 / len(columns))
        for column in columns:
            for position, box in enumerate(column):
                ceiling = (column[position + 1]["top"]
                           if position + 1 < len(column) else bottom)
                room = ceiling - box["top"] - PLANNER_BLOCK_GAP_MM
                height = max(box["natural"], PLANNER_MIN_BLOCK_MM)
                box["height"] = max(PLANNER_MIN_TIGHT_MM, min(height, max(room, 0.0)))
                box["x"] = lane_x + box["column"] * (width + PLANNER_BLOCK_GAP_MM)
                box["width"] = width
                if box["top"] + box["height"] > bottom:
                    box["top"] = max(top, bottom - box["height"])
    return boxes


def _block_style(event, colour_key):
    """Fill/border/type colours for one block, including the hollow variant."""
    palette = planner_palette(colour_key)
    # Tentative or free-marked time is drawn hollow: an outline says "pencilled
    # in" on a printed page the way a solid block says "booked".
    if event["busy"] in (OL_BUSY_FREE, OL_BUSY_TENTATIVE):
        return {
            "fill": "#FFFFFF",
            # A pale fill makes a pale outline, so a light colour borrows its
            # own darker edge for both the rule and the subject on it.
            "stroke": palette["edge"] if palette["light"] else palette["fill"],
            "line_pt": 1.2,
            "title": palette["edge"],
            "sub": _MUTED,
            "hollow": True,
        }
    return {
        "fill": palette["fill"],
        "stroke": palette["border"],
        "line_pt": 0.6,
        "title": palette["text"],
        "sub": palette["sub"],
        "hollow": False,
    }


def _draw_block(canvas, box, show_people):
    """
    Draw one event block on the timeline.

    Three layouts, picked by how much room the block actually has: the full
    stacked one, a narrow variant for a block sharing its lane with an
    overlapping meeting, and a single centred line for anything too short to
    stack. A printed block that has been squeezed says less rather than
    truncating everything on it into nonsense.
    """
    event = box["event"]
    style = box["style"]
    x, y = box["x"], box["top"]
    width, height = box["width"], box["height"]
    pad_x, pad_y = 2.2, 1.5
    inner = width - pad_x * 2
    if inner <= 4:
        return

    canvas.rect(x, y, width, height, fill=style["fill"], stroke=style["stroke"],
                line_pt=style["line_pt"] or 0.6, radius=1.6)

    span = "{0}–{1}".format(_fmt_time(event["start"]), _fmt_time(event["end"]))
    floor_y = y + height - 1.0

    # --- too short to stack: one centred line ------------------------------
    if height < 8.6:
        size = 9.0 if height >= 5.6 else (8.0 if height >= 4.4 else 7.0)
        baseline = y + height / 2.0 + _cap_mm(size) / 2.0
        if inner < 42.0:
            # Short AND sharing the lane: one line is all there is, so spend it
            # on the subject. Where the block sits already says when it is.
            canvas.text(x + pad_x, baseline, event["subject"], size, bold=True,
                        colour=style["title"], max_width=inner)
            return
        detail = span
        if event["location"]:
            longer = span + " · " + event["location"]
            if canvas.text_width(longer, 7.5) < inner * 0.62:
                detail = longer
        detail_w = canvas.text_width(detail, 7.5)
        title_w = inner - detail_w - 2.4
        if title_w < 14.0:               # no room for both - the time wins
            detail, detail_w = span, canvas.text_width(span, 7.5)
            title_w = inner - detail_w - 2.4
        if title_w < 10.0:               # nor for that - print the title alone
            canvas.text(x + pad_x, baseline, event["subject"], size, bold=True,
                        colour=style["title"], max_width=inner)
            return
        drawn = canvas.text(x + pad_x, baseline, event["subject"], size, bold=True,
                            colour=style["title"], max_width=title_w)
        canvas.text(x + pad_x + drawn + 2.4, baseline, detail, 7.5,
                    colour=style["sub"], max_width=inner - drawn - 2.4)
        return

    # --- narrow, because it is sharing the lane ----------------------------
    if inner < 42.0:
        title_size = 9.0
        baseline = y + pad_y + _cap_mm(title_size)
        lines = canvas.wrap(event["subject"], title_size, inner, bold=True,
                            max_lines=2 if height >= 12.0 else 1)
        for line in lines:
            if baseline > floor_y:
                break
            canvas.text(x + pad_x, baseline, line, title_size, bold=True,
                        colour=style["title"])
            baseline += 3.2
        if baseline - 0.4 <= floor_y:
            # Too narrow for both ends of the range: the start time alone beats
            # a truncated one, and the block's position says how long it runs.
            shown_span = (span if canvas.text_width(span, 7.5, bold=True) <= inner
                          else _fmt_time(event["start"]))
            canvas.text(x + pad_x, baseline, shown_span, 7.5, bold=True,
                        colour=style["sub"], max_width=inner)
            baseline += 2.8
        if event["location"] and baseline <= floor_y:
            canvas.text(x + pad_x, baseline, event["location"], 7.5,
                        colour=style["sub"], max_width=inner)
        return

    # --- the full block ----------------------------------------------------
    title_size = 10.5 if inner >= 62.0 else 9.5
    span_w = canvas.text_width(span, 8.0, bold=True)
    baseline = y + pad_y + _cap_mm(title_size)
    canvas.text(x + pad_x, baseline, event["subject"], title_size, bold=True,
                colour=style["title"], max_width=inner - span_w - 3.0)
    canvas.text(x + width - pad_x, baseline, span, 8.0, bold=True,
                colour=style["sub"], align="right")

    baseline += 3.1
    if event["location"] and baseline <= floor_y:
        canvas.text(x + pad_x, baseline, event["location"], 8.0, bold=True,
                    colour=style["sub"], max_width=inner)
        baseline += 2.8
    if show_people and event["people"] and baseline <= floor_y:
        shown = event["people"][:4]
        line = " · ".join(shown)
        if len(event["people"]) > len(shown):
            line += " +{0}".format(len(event["people"]) - len(shown))
        canvas.text(x + pad_x, baseline, line, 7.5, colour=style["sub"],
                    max_width=inner)


def render_day_planner(day, day_events, ahead, identity, hours,
                       show_people=True, withheld=0, category_colours=None):
    """
    Build the whole A4-landscape planner and return it as PDF bytes.

    `day_events` is today's list, `ahead` a list of (date, events) for the
    right-hand panel, `identity` the (name, address) the header and footer show,
    and `hours` the (first, last) hour the timeline covers.
    """
    category_colours = category_colours or {}
    page_w, page_h = 297.0, 210.0
    pad_l = pad_r = 11.0
    pad_t, pad_b = 10.0, 8.0
    name, address = identity

    canvas = PdfCanvas(page_w, page_h,
                       title="Daily agenda - {0}".format(_fmt_day_long(day)),
                       author=name or address or "")

    timed = [event for event in day_events if not event["all_day"]]
    all_day = [event for event in day_events if event["all_day"]]

    # Buckets first: the legend can only show colours that are on the page.
    for event in day_events:
        label, colour_key = _event_bucket(event, category_colours)
        event["bucket"] = label
        event["colour"] = colour_key

    # ---- header ----------------------------------------------------------
    eyebrow = "DAILY AGENDA" + (" · " + name if name else "")
    canvas.text(pad_l, pad_t + 2.0, eyebrow, 7.5, bold=True, colour=_MUTED,
                tracking=0.85, max_width=120.0)
    date_w = canvas.text(pad_l, pad_t + 8.6, _fmt_day_long(day), 19.0,
                         bold=True, colour=_INK)

    booked = sum((event["end"] - event["start"]).total_seconds() / 60.0
                 for event in timed)
    stats = "{0} meeting{1} · {2} booked".format(
        len(timed), "" if len(timed) == 1 else "s", _fmt_duration(booked))
    if all_day:
        stats += " · {0} all day".format(len(all_day))
    if not timed and not all_day:
        stats = "Nothing scheduled"

    cursor = page_w - pad_r
    cursor -= canvas.text(cursor, pad_t + 8.0, stats, 7.5, bold=True,
                          colour=_FAINT, align="right")
    cursor -= 4.0

    # Legend. With Outlook's own categories driving the colours there can be
    # more of them than the header has room for, so rank by how much of the day
    # each one accounts for, keep what fits, and put the survivors back in the
    # order the legend always uses so it does not reshuffle day to day.
    counts = {}
    for event in day_events:
        key = (event["bucket"], event["colour"])
        counts[key] = counts.get(key, 0) + 1
    legend = sorted(sorted(counts, key=lambda key: (-counts[key], _bucket_rank(key)))
                    [:PLANNER_MAX_LEGEND], key=_bucket_rank)
    if any(event["busy"] in (OL_BUSY_FREE, OL_BUSY_TENTATIVE) for event in day_events):
        legend.append(("Tentative", None))

    def chip_width(text):
        return 2.6 + 2.0 + canvas.text_width(text, 7.5, bold=True) + 6.0

    room = cursor - (pad_l + date_w + 8.0)
    while legend and sum(chip_width(entry[0]) for entry in legend) > room:
        legend.pop()  # the least of the day goes first, never the most of it

    if legend:
        canvas.line(cursor, pad_t + 3.4, cursor, pad_t + 9.2, _RULE, 0.6)
        cursor -= 4.0
    for label, colour_key in reversed(legend):
        cursor -= canvas.text(cursor, pad_t + 8.0, label, 7.5, bold=True,
                              colour="#374151", align="right")
        cursor -= 2.0
        if colour_key is None:  # the hollow swatch
            canvas.rect(cursor - 2.6, pad_t + 5.7, 2.6, 2.6, fill="#FFFFFF",
                        stroke=_INK, line_pt=0.9, radius=0.9)
        else:
            swatch = planner_palette(colour_key)
            canvas.rect(cursor - 2.6, pad_t + 5.7, 2.6, 2.6, fill=swatch["fill"],
                        stroke=swatch["border"], line_pt=0.6, radius=0.9)
        cursor -= 2.6 + 6.0

    header_rule = pad_t + 12.8
    canvas.line(pad_l, header_rule, page_w - pad_r, header_rule, _INK, 1.2)

    # ---- page frame ------------------------------------------------------
    body_top = header_rule + 5.0
    footer_rule = page_h - pad_b - 4.6
    body_bottom = footer_rule - 2.5
    centre_x = pad_l + (page_w - pad_l - pad_r) / 2.0
    left_x = pad_l
    left_w = centre_x - 9.0 - left_x
    right_x = centre_x + 9.0
    right_w = (page_w - pad_r) - right_x
    canvas.line(centre_x, header_rule + 4.0, centre_x, footer_rule - 2.0,
                _GRID, 0.6, dash=[2.2, 2.2])

    # ---- left panel: the day itself --------------------------------------
    today = datetime.date.today()
    if day == today:
        panel_title = "Today"
    elif day == today + datetime.timedelta(days=1):
        panel_title = "Tomorrow"
    else:
        panel_title = _WEEKDAY_NAMES[day.weekday()]
    canvas.text(left_x, body_top + 3.0, panel_title, 10.5, bold=True, colour=_INK)
    canvas.text(left_x + left_w, body_top + 3.0,
                "{0:02d}:00 — {1:02d}:00".format(hours[0], hours[1]), 7.0,
                bold=True, colour=_FAINT, tracking=0.8, align="right")

    cursor_y = body_top + 6.4

    # All-day items have no place on an hour grid, so they sit above it as a
    # strip of chips - which is also where a printed diary puts them.
    if all_day:
        chip_h = 5.0
        chip_x, chip_y = left_x, cursor_y
        rows = 1
        for event in all_day:
            label = event["subject"]
            style = _block_style(event, event["colour"])
            chip_w = min(left_w, canvas.text_width(label, 8.0, bold=True) + 5.0)
            if chip_x + chip_w > left_x + left_w and chip_x > left_x:
                if rows >= 2:
                    canvas.text(chip_x + 1.0, chip_y + chip_h / 2.0 + _cap_mm(7.5) / 2.0,
                                "+{0} more".format(
                                    len(all_day) - all_day.index(event)),
                                7.5, bold=True, colour=_MUTED)
                    break
                rows += 1
                chip_x = left_x
                chip_y += chip_h + 1.2
            canvas.rect(chip_x, chip_y, chip_w, chip_h, fill=style["fill"],
                        stroke=style["stroke"], line_pt=style["line_pt"] or 0.6,
                        radius=1.2)
            canvas.text(chip_x + 2.5, chip_y + chip_h / 2.0 + _cap_mm(8.0) / 2.0,
                        label, 8.0, bold=True, colour=style["title"],
                        max_width=chip_w - 5.0)
            chip_x += chip_w + 1.6
        cursor_y = chip_y + chip_h + 2.6

    # ---- the hour grid ---------------------------------------------------
    gutter = 13.0
    grid_top = cursor_y + 2.2
    grid_bottom = body_bottom - 2.2
    lane_x = left_x + gutter
    lane_w = (left_x + left_w) - lane_x
    start_min, end_min = hours[0] * 60, hours[1] * 60
    scale = (grid_bottom - grid_top) / float(max(1, end_min - start_min))

    label_step = 1 if (60.0 * scale) >= 4.0 else 2
    for hour in range(hours[0], hours[1] + 1):
        y = grid_top + (hour * 60 - start_min) * scale
        canvas.line(lane_x, y, left_x + left_w, y, _GRID, 0.75, dash=[0.6, 1.8])
        if (hour - hours[0]) % label_step == 0:
            canvas.text(left_x, y + _cap_mm(7.5) / 2.0,
                        "{0:02d}:00".format(hour), 7.5, bold=True, colour=_FAINT)

    if timed:
        for box in _place_blocks(timed, day, start_min, end_min,
                                 grid_top, grid_bottom, lane_x, lane_w):
            box["style"] = _block_style(box["event"], box["event"]["colour"])
            _draw_block(canvas, box, show_people)
    elif not all_day:
        canvas.text(lane_x + lane_w / 2.0, (grid_top + grid_bottom) / 2.0,
                    "Nothing in the diary", 10.0, bold=True, colour=_GRID,
                    align="centre")

    # ---- right panel: the days after -------------------------------------
    if ahead:
        first, last = ahead[0][0], ahead[-1][0]
        if len(ahead) == 1:
            ahead_title = "Then {0}".format(_WEEKDAY_NAMES[first.weekday()])
            ahead_range = _fmt_day_short(first)
        else:
            ahead_title = "The next {0} days".format(
                _NUMBER_WORDS.get(len(ahead), len(ahead)))
            ahead_range = "{0} — {1}".format(
                first.day if first.month == last.month else _fmt_day_short(first),
                _fmt_day_short(last))
    else:
        ahead_title, ahead_range = "The days ahead", ""

    canvas.text(right_x, body_top + 3.0, ahead_title, 10.5, bold=True, colour=_INK)
    canvas.text(right_x + right_w, body_top + 3.0, ahead_range, 7.0, bold=True,
                colour=_FAINT, tracking=0.8, align="right")

    rows_top = body_top + 6.4
    time_col = 19.0
    date_col = 17.0
    date_block_mm = 12.4   # the DOW / number / month stack down the left
    row_gap_mm = 2.8
    base_line_mm = 4.3

    # Rows are sized to what is in them and then scaled to fill the panel, so a
    # quiet Wednesday does not take up as much of the page as a full Thursday.
    natural = [max(date_block_mm, 2.6 + max(1, len(events)) * base_line_mm) + row_gap_mm
               for _ahead_day, events in ahead]
    available = body_bottom - rows_top
    factor = available / sum(natural) if sum(natural) else 1.0
    heights = [height * factor for height in natural]

    top = rows_top
    for index, (ahead_day, ahead_events) in enumerate(ahead):
        row_h = heights[index]
        accent = planner_palette(
            PLANNER_DAY_ACCENTS[index % len(PLANNER_DAY_ACCENTS)])
        canvas.text(right_x, top + 2.2, _WEEKDAY_SHORT[ahead_day.weekday()], 7.5,
                    bold=True, colour=accent["fill"], tracking=0.75)
        canvas.text(right_x, top + 7.2, str(ahead_day.day), 14.0, bold=True, colour=_INK)
        canvas.text(right_x, top + 10.6, _MONTH_SHORT[ahead_day.month - 1], 7.0,
                    bold=True, colour=_FAINT)

        list_x = right_x + date_col
        list_w = right_w - date_col
        # Spread the day's lines over the row it was given, within reason: a
        # short list looks deliberate spaced out, and absurd spaced to the floor.
        line_h = base_line_mm
        if ahead_events:
            line_h = max(base_line_mm, min(6.4, (row_h - row_gap_mm - 2.6)
                                           / max(1, len(ahead_events))))
        capacity = max(1, int((row_h - row_gap_mm - 2.2) / line_h))
        baseline = top + 2.6

        if not ahead_events:
            canvas.text(list_x, baseline, "Nothing scheduled", 9.0, colour=_FAINT)
        else:
            shown = ahead_events[:capacity]
            if len(ahead_events) > capacity:
                shown = ahead_events[:max(1, capacity - 1)]
            for event in shown:
                when = ("all day" if event["all_day"]
                        else "{0}–{1}".format(_fmt_time(event["start"]),
                                              _fmt_time(event["end"])))
                canvas.text(list_x, baseline, when, 8.0, bold=True, colour=_MUTED,
                            max_width=time_col - 1.0)
                canvas.text(list_x + time_col + 2.0, baseline, event["subject"], 9.0,
                            bold=True, colour=_INK,
                            max_width=list_w - time_col - 2.0)
                baseline += line_h
            if len(ahead_events) > len(shown):
                canvas.text(list_x + time_col + 2.0, baseline,
                            "+{0} more".format(len(ahead_events) - len(shown)),
                            8.0, bold=True, colour=_FAINT)

        if index < len(ahead) - 1:
            rule_y = top + row_h - row_gap_mm / 2.0
            canvas.line(right_x, rule_y, right_x + right_w, rule_y, _HAIRLINE, 0.5)
        top += row_h

    # ---- footer ----------------------------------------------------------
    canvas.line(pad_l, footer_rule, page_w - pad_r, footer_rule, _RULE, 0.5)
    left_footer = "Outlook Calendar" + (" · " + address if address else "")
    if withheld:
        left_footer += " · {0} item{1} withheld by the content policy".format(
            withheld, "" if withheld == 1 else "s")
    canvas.text(pad_l, footer_rule + 3.4, left_footer, 7.0, bold=True,
                colour=_FAINT, max_width=170.0)
    printed = datetime.datetime.now()
    canvas.text(page_w - pad_r, footer_rule + 3.4,
                "Printed {0} at {1}".format(
                    _fmt_day_short(printed.date()), _fmt_time(printed)),
                7.0, bold=True, colour=_FAINT, align="right")

    return canvas.to_bytes()


# ---------------------------------------------------------------------------
# The outlook_print_calendar tool
# ---------------------------------------------------------------------------

PLANNER_SCAN_HORIZON_DAYS = 14   # how far ahead empty days may be skipped over
PLANNER_MAX_LOOKAHEAD = 6        # more rows than this and the panel is unreadable

_identity_cache = None


def current_user_identity(ns):
    """
    (display name, SMTP address) for the mailbox owner, for the header/footer.

    Tried in order of how cheap and how likely each is to work; every step is
    optional, because a planner with no name on it still prints.
    """
    global _identity_cache
    if _identity_cache is not None:
        return _identity_cache

    name, address = "", ""
    try:
        name = _one_line(ns.CurrentUser.Name or "")
    except Exception:
        pass
    try:
        accounts = ns.Accounts
        for index in range(1, int(accounts.Count) + 1):
            candidate = _one_line(accounts.Item(index).SmtpAddress or "")
            if "@" in candidate:
                address = candidate
                break
    except Exception:
        pass
    if not address:
        try:
            exchange_user = ns.CurrentUser.AddressEntry.GetExchangeUser()
            if exchange_user is not None:
                address = _one_line(exchange_user.PrimarySmtpAddress or "")
        except Exception:
            pass
    _identity_cache = (name, address)
    return _identity_cache


def outlook_category_colours(ns):
    """
    Every Outlook category the profile has, as {lower-case name: palette key}.

    Outlook already stores a colour against each category, which is the user's
    own colour-coding of their diary. Reading it means a calendar coloured in
    Outlook prints in the same scheme with nothing to configure - just at the
    strength the printed design uses, since Outlook's swatches are muted on
    screen and would wash out on paper.

    A category set to "None" is skipped, so it falls through to the next
    category on the appointment and then to the built-in External / Internal /
    Personal split. Any failure returns what was read so far: colour is a
    nicety, and a planner still prints without it.
    """
    mapping = {}
    try:
        categories = ns.Categories
        count = int(categories.Count)
    except Exception:
        log("Could not read the Outlook category list; the planner will fall "
            "back to External / Internal / Personal colours.")
        return mapping

    for index in range(1, count + 1):
        try:
            category = categories.Item(index)
            name = _one_line(category.Name or "")
            colour = int(category.Color)
        except Exception:
            continue
        key = OL_CATEGORY_COLOURS.get(colour)
        if name and key:
            mapping[name.lower()] = key
    return mapping


def parse_planner_date(value, today):
    """
    Turn the tool's 'date' argument into a real date.

    YYYY-MM-DD, or one of 'today' / 'tomorrow' / 'yesterday', or a signed day
    offset like '+2'. The words are accepted on purpose: the machine's clock is
    the authority on an airgapped endpoint, so 'tomorrow' should not depend on
    the model having today's date right. Raises ValueError on anything else.
    """
    text = (value or "").strip().lower()
    if not text or text == "today":
        return today
    if text == "tomorrow":
        return today + datetime.timedelta(days=1)
    if text == "yesterday":
        return today - datetime.timedelta(days=1)
    if re.fullmatch(r"[+-]\d{1,3}", text):
        return today + datetime.timedelta(days=int(text))
    return datetime.datetime.strptime(text, "%Y-%m-%d").date()


def _planner_hours(events, day):
    """
    The first and last hour the timeline covers.

    Starts from the configured working day and stretches, whole hours at a time,
    to take in anything scheduled outside it - an early flight or a late call
    belongs on the page, not off the top of it.
    """
    first, last = CALENDAR_DAY_START_HOUR, CALENDAR_DAY_END_HOUR
    for event in events:
        if event["all_day"]:
            continue
        start_h = int(_minutes_into_day(event["start"], day) // 60)
        end_min = _minutes_into_day(event["end"], day)
        end_h = int(-(-end_min // 60))  # ceiling division
        first = min(first, max(0, start_h))
        last = max(last, min(24, end_h))
    if last <= first:
        last = min(24, first + 1)
    return first, last


def _collect_planner_events(start_date, end_date, with_people):
    """
    Read the calendar between two dates into plain dicts, one per occurrence.

    Returns (events by date, withheld count). Blacklisted items are dropped
    before anything is laid out, so a withheld meeting cannot reach the page
    even as an unlabelled block.
    """
    start_dt = datetime.datetime.combine(start_date, datetime.time(0, 0))
    end_dt = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))
    try:
        matches, _total, _tail = scan_calendar(start_dt, end_dt)
    except Exception as exc:
        raise RuntimeError(
            "could not read the calendar folder ({0})".format(exc))

    identity = current_user_identity(get_namespace())

    by_date = {}
    withheld = 0
    for item_start, item in matches:
        reason = appointment_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a calendar item from the planner (blacklist match: {0}).".format(reason))
            continue
        day = item_start.date()
        event = _appointment_to_event(
            item, item_start, identity,
            want_people=(with_people and day == start_date))
        if event is None:
            continue
        by_date.setdefault(day, []).append(event)

    for events in by_date.values():
        events.sort(key=lambda event: (not event["all_day"], event["start"]))
    return by_date, withheld


def _planner_ahead_days(by_date, day, wanted, skip_empty):
    """
    Which days fill the right-hand panel.

    With `skip_empty` on, days with nothing in the diary are passed over in
    favour of the next day that has something - printing Friday's planner should
    show the week ahead, not two blank weekend panels. If nothing at all is
    scheduled in the horizon, the plain consecutive days are shown instead.
    """
    ahead = []
    if not skip_empty:
        return [(day + datetime.timedelta(days=offset),
                 by_date.get(day + datetime.timedelta(days=offset), []))
                for offset in range(1, wanted + 1)]

    for offset in range(1, PLANNER_SCAN_HORIZON_DAYS + 1):
        candidate = day + datetime.timedelta(days=offset)
        events = by_date.get(candidate)
        if events:
            ahead.append((candidate, events))
        if len(ahead) == wanted:
            break
    if not ahead:
        return [(day + datetime.timedelta(days=offset), [])
                for offset in range(1, wanted + 1)]
    return ahead


def tool_print_calendar(args):
    if not PDF_DIR:
        return ("Error: printing is switched off on this server "
                "(OUTLOOK_DOCS_DIR is set to off), so no PDF was written. "
                "Give OUTLOOK_DOCS_DIR a real folder, or unset it to use "
                "%EVA_DOCUMENTS_DIR%\\{0}.".format(DOCS_SUBFOLDER))

    today = datetime.date.today()
    try:
        day = parse_planner_date(args.get("date"), today)
    except ValueError:
        return ("Error: 'date' must be YYYY-MM-DD, or one of 'today', "
                "'tomorrow', 'yesterday', or a day offset like '+2'.")

    try:
        wanted = int(args.get("lookahead_days", CALENDAR_LOOKAHEAD_DAYS))
    except (TypeError, ValueError):
        wanted = CALENDAR_LOOKAHEAD_DAYS
    wanted = max(1, min(PLANNER_MAX_LOOKAHEAD, wanted))
    show_people = bool(args.get("show_attendees", True))
    skip_empty = bool(args.get("skip_empty_days", True))

    horizon = day + datetime.timedelta(
        days=PLANNER_SCAN_HORIZON_DAYS if skip_empty else wanted)
    try:
        by_date, withheld = _collect_planner_events(day, horizon, show_people)
    except RuntimeError as exc:
        return "Error: {0}.".format(exc)

    day_events = by_date.get(day, [])
    ahead = _planner_ahead_days(by_date, day, wanted, skip_empty)
    hours = _planner_hours(day_events, day)

    # Outlook's own category colours first, then the endpoint's explicit map on
    # top: OUTLOOK_CALENDAR_COLOURS is there to overrule a category whose
    # Outlook colour is not the one you want on paper, so it has to win.
    category_colours = outlook_category_colours(get_namespace())
    category_colours.update(_CATEGORY_COLOURS)

    try:
        pdf = render_day_planner(
            day, day_events, ahead, current_user_identity(get_namespace()),
            hours, show_people=show_people, withheld=withheld,
            category_colours=category_colours)
    except Exception as exc:
        log("Planner render failed:\n{0}".format(traceback.format_exc()))
        return "Error: the planner could not be laid out ({0}).".format(exc)

    # basename() then safe_filename() so a path in 'filename' cannot escape the
    # folder: "..\\elsewhere\\x.pdf" resolves to "x.pdf" inside it.
    requested = _one_line(args.get("filename") or "")
    if requested:
        stem = safe_filename(os.path.splitext(os.path.basename(requested))[0])
    else:
        stem = "Calendar - {0} {1}".format(
            day.isoformat(), _WEEKDAY_NAMES[day.weekday()])
    path = os.path.join(PDF_DIR, stem + ".pdf")

    # Re-printing a day overwrites that day's own sheet, which is the point. A
    # NAMED file is different: the documents folder holds the user's own PDFs
    # too, and silently replacing one of those would be unforgivable.
    if requested and os.path.exists(path):
        return ("Error: {0} already exists and this call named it explicitly, "
                "so nothing was written - a file you already have must not be "
                "replaced without asking. Choose another name, or drop "
                "'filename' to write the dated sheet, which does overwrite "
                "its own previous copy.".format(path))

    try:
        os.makedirs(PDF_DIR, exist_ok=True)
        with open(path, "wb") as handle:
            handle.write(pdf)
    except OSError as exc:
        return ("Error: could not write the planner to {0} ({1}). The folder "
                "comes from OUTLOOK_DOCS_DIR, or %EVA_DOCUMENTS_DIR%\\{2}."
                .format(path, exc, DOCS_SUBFOLDER))

    log("Wrote day planner: {0}".format(path))

    timed = [event for event in day_events if not event["all_day"]]
    booked = sum((event["end"] - event["start"]).total_seconds() / 60.0
                 for event in timed)
    lines = [
        "Printable day planner written to:",
        "  {0}".format(path),
        "",
        "{0}: {1} meeting(s), {2} booked, {3} all-day item(s). Timeline "
        "{4:02d}:00-{5:02d}:00.".format(
            _fmt_day_long(day), len(timed), _fmt_duration(booked),
            len(day_events) - len(timed), hours[0], hours[1]),
    ]
    for event in day_events:
        when = ("all day" if event["all_day"]
                else "{0}-{1}".format(_fmt_time(event["start"]),
                                      _fmt_time(event["end"])))
        lines.append("  - {0}  {1}{2}".format(
            when, event["subject"],
            " ({0})".format(event["location"]) if event["location"] else ""))
    if not day_events:
        lines.append("  (nothing in the diary - the page prints empty)")

    lines.append("")
    lines.append("Following days on the page: " + (", ".join(
        "{0} ({1} item(s))".format(_fmt_day_short(other), len(events))
        for other, events in ahead) or "none"))
    if withheld:
        lines.append("[{0} event(s) withheld by the content blacklist and left "
                     "off the page.]".format(withheld))
    lines.append("It is A4 landscape - print it single-sided and fold it in "
                 "half for a bifold day planner.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# MCP tool registry
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "outlook_list_recent_emails",
        "description": (
            "List the most recent emails in the Outlook Inbox, newest first. "
            "Returns subject, sender, received time, read/unread status, and an "
            "EntryID for each. Use the EntryID with outlook_get_email to read the "
            "full message body. Some messages may be withheld by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "count": {"type": "integer", "description": "How many messages to return (default 10)."},
                "unread_only": {"type": "boolean", "description": "If true, only unread messages (default false)."},
            },
        },
    },
    {
        "name": "outlook_search_emails",
        "description": (
            "Search the Outlook Inbox for messages whose subject OR sender name "
            "contains the query text. Returns subject, sender, received time, and "
            "an EntryID for each match. Some matches may be withheld by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Text to look for in the subject or sender name."},
                "count": {"type": "integer", "description": "Maximum matches to return, newest first (default 10)."},
            },
            "required": ["query"],
        },
    },
    {
        "name": "outlook_get_email",
        "description": (
            "Retrieve the full details and plain-text body of a single email, "
            "identified by the EntryID from a list or search result. The message "
            "may be withheld if it is blocked by a content policy. Reading a "
            "message does not save it anywhere; pass 'save_to_kb' only if the "
            "user asks for it to be kept."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "entry_id": {"type": "string", "description": "The EntryID of the message to read."},
                "save_to_kb": {
                    "type": "boolean",
                    "description": (
                        "Save this email into the local knowledge base as "
                        "Markdown, for the RAG index. Default false: reading a "
                        "message does NOT save it. Set it to true ONLY when the "
                        "user asks for the email to be kept - 'save this to the "
                        "knowledge base', 'add that email to the KB', 'keep this "
                        "for later'. Do not set it while merely reading mail to "
                        "answer a question; saving correspondence nobody asked "
                        "for is what makes the knowledge base return irrelevant "
                        "results later."
                    ),
                },
            },
            "required": ["entry_id"],
        },
    },
    {
        "name": "outlook_get_calendar",
        "description": (
            "List calendar events between two dates (inclusive), sorted by start time. "
            "Recurring meetings are expanded into individual occurrences. Dates are "
            "YYYY-MM-DD; defaults to today through 7 days ahead. Some events may be "
            "withheld by a content policy. If nothing matches, the reply includes a "
            "[debug] section listing the last few calendar items scanned."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date, YYYY-MM-DD (default: today)."},
                "end_date": {"type": "string", "description": "End date, YYYY-MM-DD (default: 7 days from today)."},
                "max_results": {"type": "integer", "description": "Maximum events to return (default 50)."},
            },
        },
    },
    {
        "name": "outlook_print_calendar",
        "description": (
            "Create a PRINTABLE PDF day planner for one day, laid out A4 landscape "
            "as a bifold: an hour-by-hour timeline of that day on the left half, and "
            "a summary of the following days on the right. Use it whenever the user "
            "asks for a printable/printed calendar, a day planner, an agenda to print "
            "or take into a meeting, or 'today's calendar as a PDF'. 'date' accepts "
            "'today' (the default), 'tomorrow', 'yesterday', a day offset like '+2', "
            "or YYYY-MM-DD - prefer the words, because the server uses the endpoint's "
            "own clock. The file is written into the documents folder and the tool "
            "reports the full path plus what is on the page. Reading the calendar is "
            "still read-only; events withheld by the content policy are left off the "
            "page and counted in the footer."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "date": {
                    "type": "string",
                    "description": (
                        "Day to print: 'today' (default), 'tomorrow', 'yesterday', "
                        "an offset like '+2'/'-1', or YYYY-MM-DD."
                    ),
                },
                "lookahead_days": {
                    "type": "integer",
                    "description": (
                        "How many following days the right-hand panel lists, 1-6 "
                        "(default 4)."
                    ),
                },
                "skip_empty_days": {
                    "type": "boolean",
                    "description": (
                        "Default true: days with nothing in the diary are passed over "
                        "in favour of the next day that has something, so a Friday "
                        "planner shows the week ahead instead of two blank weekend "
                        "panels. Set false for strictly consecutive days."
                    ),
                },
                "show_attendees": {
                    "type": "boolean",
                    "description": (
                        "Default true: print attendee names inside the day's larger "
                        "meeting blocks. Set false for a planner that can be left on "
                        "a desk."
                    ),
                },
                "filename": {
                    "type": "string",
                    "description": (
                        "Optional file name for the PDF. Defaults to "
                        "'Calendar - <date> <weekday>.pdf'; an existing file of the "
                        "same name is overwritten."
                    ),
                },
            },
        },
    },
    {
        "name": "outlook_list_sent_emails",
        "description": (
            "List messages you SENT, within a date range (inclusive), newest first. "
            "Returns sent time, recipients, and subject for each, plus an EntryID for "
            "reading the full body with outlook_get_email. Useful for reviewing what "
            "you did over a period, e.g. 'what did I send last week'. Dates are "
            "YYYY-MM-DD; defaults to the last 7 days. Some messages may be withheld by "
            "a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date, YYYY-MM-DD (default: 7 days ago)."},
                "end_date": {"type": "string", "description": "End date, YYYY-MM-DD (default: today)."},
                "max_results": {"type": "integer", "description": "Maximum messages to return (default 100)."},
            },
        },
    },
    {
        "name": "outlook_search_recent",
        "description": (
            "List emails across MULTIPLE folders at once (by default Inbox, Sent Items, "
            "and Archive) within a date range, newest first, merged into one list. Each "
            "result is labelled with the folder it came from. Optionally override which "
            "folders to search with 'folders', or filter by text in the subject, sender, "
            "or recipient. Use this for 'all my email over the last week' style questions. "
            "Dates are YYYY-MM-DD; defaults to the last 7 days. Some items may be withheld "
            "by a content policy."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date, YYYY-MM-DD (default: 7 days ago)."},
                "end_date": {"type": "string", "description": "End date, YYYY-MM-DD (default: today)."},
                "query": {"type": "string", "description": "Optional text to match in subject, sender, or recipient."},
                "folders": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Optional list of folder NAMES to search, matched across all stores "
                        "(e.g. [\"Inbox\", \"Sent Items\"]). Overrides the default set. Run "
                        "outlook_list_folders to see the available names."
                    ),
                },
                "max_results": {"type": "integer", "description": "Maximum emails to return (default 100)."},
            },
        },
    },
    {
        "name": "outlook_list_folders",
        "description": (
            "List every mail folder across all Outlook stores (main mailbox, online "
            "archive, mounted PSTs), with item counts. Use this to discover the exact "
            "folder names available - especially to see what your 'Archive' really is - "
            "so the combined search can be pointed at the right folders."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "max_results": {"type": "integer", "description": "Maximum folders to list (default 300)."},
            },
        },
    },
]

TOOL_DISPATCH = {
    "outlook_list_recent_emails": tool_list_recent_emails,
    "outlook_search_emails": tool_search_emails,
    "outlook_get_email": tool_get_email,
    "outlook_get_calendar": tool_get_calendar,
    "outlook_print_calendar": tool_print_calendar,
    "outlook_list_sent_emails": tool_list_sent_emails,
    "outlook_search_recent": tool_search_recent,
    "outlook_list_folders": tool_list_folders,
}


# ---------------------------------------------------------------------------
# JSON-RPC / MCP plumbing
# ---------------------------------------------------------------------------

PROTOCOL_VERSION_DEFAULT = "2024-11-05"
SERVER_INFO = {"name": "outlook-mcp", "version": __version__}


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
        except Exception as exc:
            log("Tool '{0}' failed:\n{1}".format(name, traceback.format_exc()))
            reset_namespace()  # force a fresh Outlook connection next call
            return rpc_result(req_id, text_content("Outlook tool error: {0}".format(exc), is_error=True))

    if is_notification:
        return None
    return rpc_error(req_id, -32601, "Method not found: {0}".format(method))


def run_server():
    """Main stdio loop: read newline-delimited JSON-RPC, dispatch, respond."""
    pythoncom.CoInitialize()
    log("outlook-mcp server started (stdio). Waiting for requests...")
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
        pythoncom.CoUninitialize()
        log("outlook-mcp server stopped.")


def run_check():
    """Connect to Outlook, print diagnostics + blacklist status to stderr, then exit."""
    pythoncom.CoInitialize()
    try:
        ns = get_namespace()
        log("Connected to Outlook MAPI namespace.")
        try:
            log("Current user        : {0}".format(ns.CurrentUser))
        except Exception as exc:
            log("Could not read CurrentUser: {0}".format(exc))

        inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
        log("Inbox folder        : {0}".format(inbox.Name))
        try:
            log("Inbox item count    : {0}".format(inbox.Items.Count))
        except Exception as exc:
            log("Could not count Inbox items: {0}".format(exc))

        cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
        log("Calendar folder     : {0}".format(cal.Name))
        log("Search folders      : {0}".format(", ".join(_SEARCH_FOLDERS)))
        if PDF_DIR:
            log("Planner PDF folder  : {0}".format(PDF_DIR))
        else:
            log("Planner PDF folder  : disabled - outlook_print_calendar is off")
        if _CATEGORY_COLOURS:
            log("Planner colours     : {0}".format(", ".join(
                "{0}={1}".format(key, value)
                for key, value in sorted(_CATEGORY_COLOURS.items()))))
        if KB_DIR:
            log("KB save folder      : {0} ({1})".format(
                KB_DIR,
                "every email read (autosave on)" if KB_AUTOSAVE
                else "on request only"))
        else:
            log("KB save folder      : disabled - no email can be saved")
        log("CHECK OK - Outlook COM link is working.")
        return 0
    except Exception:
        log("CHECK FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    finally:
        pythoncom.CoUninitialize()


def main():
    global KB_DIR, KB_AUTOSAVE, PDF_DIR, _CATEGORY_COLOURS
    global CALENDAR_DAY_START_HOUR, CALENDAR_DAY_END_HOUR

    parser = argparse.ArgumentParser(
        description=(
            "Read-only MCP server exposing local Outlook mail and calendar via COM, "
            "with a content blacklist that withholds classified/marked items. With "
            "no check flag it runs as an stdio MCP server. Configuration is "
            "environment variables only: EVA_KNOWLEDGE_DIR (this server saves "
            "into its 'email' sub-folder), EVA_DOCUMENTS_DIR (it prints the "
            "day planner into the 'pdf' one), OUTLOOK_SEARCH_FOLDERS, "
            "OUTLOOK_BLACKLIST_FILE - see the CONFIGURATION section of this "
            "file's docstring."
        )
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Connect to Outlook, print diagnostics + blacklist status to stderr, then exit.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="outlook-mcp {0}".format(SERVER_INFO["version"]),
    )
    args = parser.parse_args()

    blacklist_file = env("OUTLOOK_BLACKLIST_FILE")
    search_folders = env("OUTLOOK_SEARCH_FOLDERS")
    require_blacklist = env_flag("OUTLOOK_REQUIRE_BLACKLIST", False)

    # Markdown knowledge-base saving. Resolve the folder now and make sure it is
    # usable so a misconfiguration is caught at startup, not on the first email
    # a user asks to keep.
    KB_AUTOSAVE = env_flag("OUTLOOK_KB_AUTOSAVE", KB_AUTOSAVE)
    KB_DIR = resolve_kb_dir()
    if KB_DIR:
        try:
            os.makedirs(KB_DIR, exist_ok=True)
        except OSError as exc:
            log("FATAL: could not create/use the knowledge-base folder {0}: {1}"
                .format(KB_DIR, exc))
            log("       It resolves from EVA_KNOWLEDGE_DIR (sub-folder '{0}') "
                "or OUTLOOK_KB_DIR.".format(SUBFOLDER))
            sys.exit(2)
        if KB_AUTOSAVE:
            log("Knowledge-base AUTOSAVE is on: every email read is saved -> {0}"
                .format(KB_DIR))
        else:
            log("Knowledge-base saving on request only (save_to_kb=true) -> {0}"
                .format(KB_DIR))
    else:
        log("Knowledge-base saving disabled (OUTLOOK_KB_DIR=off); no email is "
            "written to disk")
        if KB_AUTOSAVE:
            log("WARNING: OUTLOOK_KB_AUTOSAVE has no effect while the "
                "knowledge-base folder is off. Unset OUTLOOK_KB_DIR (or give "
                "it a real path) to enable saving.")

    # Printable day planner. Unlike the knowledge-base folder this is NOT fatal
    # when it cannot be used: printing is one tool, and mail and calendar must
    # still be readable on an endpoint where the document library is missing.
    PDF_DIR, docs_configured = resolve_pdf_dir()
    if PDF_DIR:
        try:
            os.makedirs(PDF_DIR, exist_ok=True)
            log("Day planner PDFs -> {0}".format(PDF_DIR))
        except OSError as exc:
            log("WARNING: cannot use the planner folder {0}: {1}".format(PDF_DIR, exc))
            log("         It resolves from {0}. outlook_print_calendar is "
                "disabled until it exists.".format(
                    "OUTLOOK_DOCS_DIR / EVA_DOCUMENTS_DIR (sub-folder '{0}')".format(
                        DOCS_SUBFOLDER)
                    if docs_configured
                    else "the built-in default {0}".format(PDF_DIR)))
            PDF_DIR = None
    else:
        log("Day planner printing disabled (OUTLOOK_DOCS_DIR=off); no PDF is "
            "written")

    # Category -> colour map for the planner: the file's own value first, then
    # the environment on top of it.
    _CATEGORY_COLOURS = {key.lower(): value.lower()
                         for key, value in CALENDAR_CATEGORY_COLOURS.items()}
    _CATEGORY_COLOURS.update(parse_category_colours(env("OUTLOOK_CALENDAR_COLOURS")))

    # Working-day window for the printed timeline, e.g. OUTLOOK_CALENDAR_HOURS="7-19".
    hours_raw = env("OUTLOOK_CALENDAR_HOURS")
    if hours_raw:
        match = re.fullmatch(r"\s*(\d{1,2})\s*[-:to]+\s*(\d{1,2})\s*", hours_raw)
        if match and 0 <= int(match.group(1)) < int(match.group(2)) <= 24:
            CALENDAR_DAY_START_HOUR = int(match.group(1))
            CALENDAR_DAY_END_HOUR = int(match.group(2))
            log("Planner timeline hours set to {0:02d}:00-{1:02d}:00.".format(
                CALENDAR_DAY_START_HOUR, CALENDAR_DAY_END_HOUR))
        else:
            log("WARNING: OUTLOOK_CALENDAR_HOURS='{0}' is not a range like '7-19'; "
                "keeping {1:02d}:00-{2:02d}:00.".format(
                    hours_raw, CALENDAR_DAY_START_HOUR, CALENDAR_DAY_END_HOUR))

    # Load any external blacklist terms and compile the filter BEFORE serving.
    extra_terms = []
    if blacklist_file:
        if not os.path.isfile(blacklist_file):
            log("FATAL: OUTLOOK_BLACKLIST_FILE not found: {0}".format(blacklist_file))
            sys.exit(2)
        try:
            extra_terms = load_blacklist_file(blacklist_file)
        except Exception as exc:
            log("FATAL: could not read OUTLOOK_BLACKLIST_FILE: {0}".format(exc))
            sys.exit(2)
    build_blacklist(extra_terms)

    # Fail closed when the blacklist is mandatory but empty.
    if (require_blacklist or REQUIRE_BLACKLIST) and _BLACKLIST_RE is None:
        log("FATAL: OUTLOOK_REQUIRE_BLACKLIST is set but the content blacklist "
            "is EMPTY. Add terms to BLACKLIST_TERMS or point "
            "OUTLOOK_BLACKLIST_FILE at a terms file. Refusing to start without "
            "the compliance filter.")
        sys.exit(2)

    # Optional override of the default outlook_search_recent folders.
    if search_folders:
        global _SEARCH_FOLDERS
        names = [name.strip() for name in search_folders.split(",") if name.strip()]
        if names:
            _SEARCH_FOLDERS = names
            log("Default outlook_search_recent folders set to: {0}".format(names))
        else:
            log("WARNING: OUTLOOK_SEARCH_FOLDERS was empty; keeping default {0}."
                .format(_SEARCH_FOLDERS))

    if args.check:
        sys.exit(run_check())
    run_server()


if __name__ == "__main__":
    main()
