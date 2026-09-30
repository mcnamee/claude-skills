#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
excel.py (v6.1.0) -- Excel (.xlsx) MCP server: reads directly, writes
through Excel.

PURPOSE
    A single-file MCP (Model Context Protocol) stdio server that lets a local
    model query, ANALYSE and change Excel workbooks. READING is standard library only: it
    parses .xlsx directly (a .xlsx file is just a ZIP of XML), never opens
    Excel and never touches the network. WRITING drives the installed desktop
    Excel through COM (pywin32), because Excel is the only thing that keeps a
    workbook's formulas, Tables, charts and pivot caches consistent - and the
    only thing that can build a real PivotTable.

    Tools exposed (always):
      * excel_list_workbooks    -- list .xlsx files in the configured folder
      * excel_list_sheets       -- sheet names + declared dimensions, and the
                                   Tables / pivot tables on each sheet
      * excel_get_headers       -- the first (header) row of a sheet
      * excel_read_range        -- read cells from a sheet, optionally an A1 range
      * excel_search            -- find cells whose value matches a query
      * excel_column_stats      -- summary statistics for one column
      * excel_list_tables       -- every Table (Insert > Table / ListObject):
                                   name, sheet, range, columns
      * excel_read_table        -- a Table by NAME, as a Markdown table, with
                                   column pick, a simple filter and paging
      * excel_list_pivot_tables -- every pivot table's layout: source, row /
                                   column / filter fields, value fields
      * excel_read_pivot_table  -- one pivot's layout plus the figures it shows

    Write tools (see WRITING - need Windows, desktop Excel and pywin32):
      * excel_write_cells        -- write a block of values/formulas into a
                                    sheet (tab), optionally adding the sheet
      * excel_add_table_rows     -- append rows to a Table by name
      * excel_update_table_rows  -- change the Table row(s) matching a value
      * excel_create_pivot_table -- build a real PivotTable from a Table or a
                                    range, on a new or existing sheet

    Design mirrors the other servers in this suite: stdout is reserved for
    JSON-RPC only, all diagnostics go to stderr, config lives in one fenced
    block below, and a --check flag validates the environment before wiring
    the server in. Workbook names are resolved forgivingly: exact filename,
    then name without extension, then case-insensitive, then a unique
    substring, and finally a FUZZY name match (difflib) - so "budgit q3" or
    "q3 budget" still opens "Budget Q3 2024.xlsx". A genuinely ambiguous name
    lists the candidates instead of guessing; use excel_list_workbooks to see
    what is available. (The fuzzy fallback is logged to stderr for audit.)

SUPPORTED / NOT SUPPORTED
    * Supported: .xlsx (and macro-enabled .xlsm, same underlying format).
    * NOT supported: legacy .xls (old binary BIFF format), .xlsb (binary),
      and password-encrypted workbooks. These are detected and reported
      clearly rather than mis-parsed.
    * Formula cells: the LAST VALUE CACHED BY EXCEL is returned (the same
      value you would see on screen). Formulas are not re-evaluated, so a
      workbook saved by a tool that did not cache values may show blanks for
      formula cells. This is inherent to reading without a calc engine.
    * Dates: Excel stores dates as serial numbers. Cells whose STYLE marks
      them as a date/time are converted to ISO-8601 text. Dates before
      1900-03-01 may be off by one day due to Excel's historical 1900
      leap-year bug; this affects almost no real enterprise data.

REQUIREMENTS
    * Python 3.8+ (standard library only) for everything but writing.
    * Writing only: Windows, desktop Microsoft Excel, and pywin32.

CONFIGURATION
    The workbook folder is REQUIRED. It is the "excel" sub-folder of
    %EVA_DOCUMENTS_DIR%, the suite-wide document root shared by every plugin in
    this repo, so H:\Eva\documents\excel on a stock install (copy the repo's
    eva\ folder to H:\Eva and it exists). THE FOLDER MUST EXIST: the server
    refuses to start otherwise, and only ever reads files inside it (symlinks
    that resolve outside the folder are excluded). Set EXCEL_DOCS_DIR to
    override just this server with a full path of its own. There are no folder
    command-line flags: configuration is environment variables only, so two
    settings can never disagree about a path. The only flags are --check,
    --list and --version. The size caps below the imports are constants.
    Precedence: EXCEL_DOCS_DIR > EVA_DOCUMENTS_DIR + "excel" > the
    EVA_DOCUMENTS_DIR constant + "excel".

    NOTE: only the TOP LEVEL of the workbook folder is listed - sub-folders are
    not searched (word.py does search recursively; this server does not). Keep
    workbooks directly in the folder, and use filename prefixes
    ("Finance - Budget FY26.xlsx") where you would otherwise want a sub-folder.

TABLES AND PIVOT TABLES (reading)
    A Table is read by its name wherever it sits, using the column names the
    Table itself defines; a totals row is reported separately. A pivot table's
    FIGURES are read from the cells Excel saved into the sheet, so they are as
    of the pivot's last refresh-and-save - the reply says when that was, and a
    pivot whose source has changed since is stale until refreshed in Excel.

WRITING  (Windows + desktop Excel + pywin32)
    Like the word and powerpoint plugins, editing documents in its folder is
    this plugin's job, so the write tools are always offered; on an endpoint
    without Excel or pywin32 they answer with what to install, and every read
    tool keeps working. Each write call:
      1. refuses up front if the workbook is open elsewhere (its ~$ lock file
         exists, or the file is locked) - close it in Excel first;
      2. starts a PRIVATE, invisible Excel (DispatchEx - your own Excel window
         is never touched), with alerts, events and macros switched off;
      3. opens the one workbook, checks every sheet / Table / column name
         BEFORE changing anything, makes the change and lets Excel
         recalculate;
      4. saves - in place, or with 'save_as' as a NEW workbook in the same
         folder, leaving the original untouched - and quits that Excel.
    Any error closes the workbook WITHOUT saving, so a half-made change never
    reaches the file. Only files inside the workbook folder can be written,
    and 'save_as' is confined to it too.
    pywin32 is imported only when a write tool runs, so reading still needs
    nothing beyond the standard library:
      & "C:\path\to\python.exe" -m pip install pywin32
    UNTESTED ON THIS REPO'S CI: the COM calls follow Excel's documented object
    model but can only run on a Windows endpoint with Excel. Run --check (it
    starts and quits Excel to prove automation works), then try excel_create_pivot_table with 'save_as' on a copy
    before relying on it.

STANDALONE TESTING (before wiring the server in)
    1) Environment / config sanity check (prints interpreter + folder state):
         python excel.py --check

    2) List available workbooks without starting the server loop:
         python excel.py --list

    3) Drive the JSON-RPC protocol by hand. On Windows PowerShell, create a
       file "probe.txt" with these three lines (each a complete JSON object
       on ONE line) and pipe it in:

         {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"probe","version":"0"}}}
         {"jsonrpc":"2.0","method":"notifications/initialized"}
         {"jsonrpc":"2.0","id":2,"method":"tools/list"}

       Then:
         Get-Content probe.txt | python excel.py
       You should see three JSON lines back on stdout (the third listing the
       tools). Any diagnostic text appears on stderr and does NOT corrupt the
       protocol stream.

    4) Call a tool by hand (adjust the workbook/sheet names):
         {"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"excel_list_sheets","arguments":{"workbook":"budget"}}}

CONFIGURATION  (environment variables, no folder flags)
    The whole plugin suite is configured by four environment variables, set
    once for your Windows account. This server uses two of them:

      EVA_PYTHON          full path to the python.exe the MCP client launches,
                          e.g. C:\Python311\python.exe (read by the plugin
                          manifest, not by this file)
      EVA_DOCUMENTS_DIR   root of the document library (default H:\Eva\documents)

    This server works in the "excel" sub-folder of the document root, and THAT
    FOLDER MUST EXIST:

      %EVA_DOCUMENTS_DIR%\excel   REQUIRED. The .xlsx/.xlsm workbooks this
                                  server may read (and the write tools
                                  change or save copies into) - top level
                                  only. The server refuses to start if it is
                                  missing.

    To set them permanently for your account (PowerShell, one-off):

      [Environment]::SetEnvironmentVariable("EVA_PYTHON", "C:\Python311\python.exe", "User")
      [Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "H:\Eva\documents", "User")

    Copy the repo's eva\ folder to H:\Eva and the folder exists - see
    eva\README.md.

    EXCEL_DOCS_DIR overrides the workbook folder with a full path of its own,
    for an endpoint whose layout differs.

    There are NO folder command-line flags: configuration is environment
    variables only, so two settings can never disagree about a path.

INSTALLING INTO CLAUDE CODE
    This server ships as the "excel" Claude Code plugin (its manifest is
    .claude-plugin/plugin.json next to this file), so the normal install is:

      /plugin marketplace add C:\path\to\claude-skills
      /plugin install excel@mcnamee-claude-skills

    It prompts for nothing: the interpreter and the workbook folder come from
    the environment variables above. PYTHONUTF8=1 is set for you by the
    manifest, so Windows cp1252 cannot corrupt output.

    To register the server by hand instead (PowerShell):

      claude mcp add excel --scope user -e PYTHONUTF8=1 -- $env:EVA_PYTHON C:\path\to\excel.py

    See README.md next to this file for the full settings reference.

PROTOCOL NOTE
    Transport is newline-delimited JSON-RPC 2.0 over stdio (one JSON object
    per line, no embedded newlines), which is the MCP stdio convention.
    Advertised protocolVersion is "2024-11-05".
"""

# Semantic version of this server. Bump on EVERY change (see CLAUDE.md):
# MAJOR = breaking config/tool change, MINOR = new feature, PATCH = fix.
__version__ = "6.1.0"

import sys
import os
import io
import re
import json
import difflib
import argparse
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

# ===========================================================================
# CONFIG  -- folders come from the environment; these are the fallbacks.
# ===========================================================================
# Four environment variables configure the whole plugin suite. Set them once
# for your Windows account and every plugin picks them up; each server works in
# its OWN sub-folder of each root, named after the plugin. This server uses one
# root, and its sub-folder is "excel":
#
#   EVA_DOCUMENTS_DIR   -> %EVA_DOCUMENTS_DIR%\excel   workbooks (read, and
#                                                  changed by the write tools)
#
# EVA_DOCUMENTS_DIR below is the fallback when the variable is not set, and
# matches the Eva working tree: copy the repo's eva\ folder to H:\Eva and the
# folder exists. There are NO folder command-line flags - configuration is
# environment variables only, so two settings can never disagree about a path.
# To point the workbook folder somewhere else on an endpoint whose layout
# differs, set EXCEL_DOCS_DIR to a full path; it beats the suite-wide root.
SUBFOLDER = "excel"                # this server's sub-folder in each root
EVA_DOCUMENTS_DIR = r"H:\Eva\documents"

# REQUIRED: folder containing the .xlsx workbooks the model is allowed to read,
# resolved from the environment in main(); the literal here is what a stock
# H:\Eva install resolves to. The server only ever reads files inside this
# folder (symlinks that resolve outside it are excluded) and REFUSES TO START
# if it is missing. NOTE this server lists only the TOP LEVEL of the folder
# (unlike word.py, which searches recursively), so workbooks must sit directly
# in it - see eva\documents\excel\README.md.
DOCS_DIR = r"H:\Eva\documents\excel"

# File extensions treated as readable workbooks (lower-case, incl. dot).
ALLOWED_EXTENSIONS = (".xlsx", ".xlsm")

# Safety caps so a huge sheet cannot flood the model's context window.
MAX_ROWS_PER_READ = 200      # max rows returned by excel_read_range in one call
MAX_COLS_PER_READ = 64       # max columns returned per row
MAX_SEARCH_HITS = 100        # max matches returned by excel_search
MAX_CELL_TEXT_LEN = 500      # long cell text is truncated to this many chars

# Workbook-name matching. An exact/substring match always wins; only when none
# is found does resolve_workbook_path fall back to a FUZZY match on the name, so
# a near-miss like "budgit" or "q3 budget" still finds "Budget Q3 2024.xlsx".
# These tune that fallback (same values as word.py / pdf-to-md.py):
#   FUZZY_MIN_RATIO       - below this similarity, AND with no shared words, a
#                           name is treated as "no match" rather than opened.
#   FUZZY_AMBIGUITY_DELTA - if a runner-up scores within this of the best, the
#                           match is too close to call and the candidates are
#                           listed instead of one being opened silently.
FUZZY_MIN_RATIO = 0.40
FUZZY_AMBIGUITY_DELTA = 0.05

# Server identity reported to the client.
SERVER_NAME = "excel-mcp"
SERVER_VERSION = __version__
PROTOCOL_VERSION = "2024-11-05"

# ===========================================================================
# End of CONFIG
# ===========================================================================


# --- stdout/stderr hygiene --------------------------------------------------
# stdout MUST carry only JSON-RPC. Reconfigure both streams to UTF-8 so that
# Windows' default cp1252 codec cannot raise on non-ASCII content. We use
# line buffering on stdout so each JSON reply is flushed promptly.
try:
    sys.stdout.reconfigure(encoding="utf-8", newline="\n")
    sys.stderr.reconfigure(encoding="utf-8")
except AttributeError:
    # Python < 3.7 fallback (not expected on 3.8+, kept defensive).
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", newline="\n")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")


def log(*parts):
    """Diagnostics to stderr ONLY. Never write diagnostics to stdout."""
    print("[excel_mcp]", *parts, file=sys.stderr, flush=True)


# ===========================================================================
# XLSX parsing (pure standard library)
# ===========================================================================

# OOXML uses XML namespaces; we work in local names to stay namespace-agnostic.
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Built-in numFmt ids that always denote a date/time (per the OOXML spec).
_BUILTIN_DATE_FMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}


def _lname(tag):
    """Return the local part of a namespaced XML tag ('{ns}row' -> 'row')."""
    return tag.rsplit("}", 1)[-1]


def _attr_local(elem, local):
    """Fetch an attribute by its local name, ignoring any namespace prefix."""
    for key, val in elem.attrib.items():
        if key.rsplit("}", 1)[-1] == local:
            return val
    return None


def _col_letters_to_index(letters):
    """'A' -> 0, 'B' -> 1, 'Z' -> 25, 'AA' -> 26 ..."""
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _col_index_to_letters(idx):
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA' ..."""
    idx += 1
    out = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def _split_cell_ref(ref):
    """'B12' -> ('B', 12). Returns (col_letters, row_number)."""
    col = []
    i = 0
    while i < len(ref) and ref[i].isalpha():
        col.append(ref[i])
        i += 1
    row_part = ref[i:]
    row = int(row_part) if row_part.isdigit() else None
    return "".join(col), row


def _looks_like_date_format(code):
    """
    Heuristic: does a custom number-format code represent a date/time?
    Strip quoted literals, escaped chars, and [colour]/[condition] brackets,
    then look for any date/time token (y m d h s). 'm' is ambiguous
    (month vs minute) but its presence still implies a date/time format.
    """
    if not code:
        return False
    cleaned = []
    i = 0
    while i < len(code):
        ch = code[i]
        if ch == '"':                      # skip quoted literal
            i += 1
            while i < len(code) and code[i] != '"':
                i += 1
            i += 1
            continue
        if ch == "\\":                     # skip escaped char
            i += 2
            continue
        if ch == "[":                      # skip [Red], [$-409], [>0] etc.
            i += 1
            while i < len(code) and code[i] != "]":
                i += 1
            i += 1
            continue
        cleaned.append(ch)
        i += 1
    text = "".join(cleaned).lower()
    return any(tok in text for tok in ("y", "m", "d", "h", "s"))


def _serial_to_isoformat(serial, date1904):
    """
    Convert an Excel date serial to an ISO-8601 string.
    See module docstring re: the pre-1900-03-01 leap-year quirk.
    """
    try:
        serial = float(serial)
    except (TypeError, ValueError):
        return None
    if date1904:
        base = datetime(1904, 1, 1)
    else:
        base = datetime(1899, 12, 30)
    dt = base + timedelta(days=serial)
    # If there is no fractional part, it's a pure date; render date-only.
    if abs(serial - round(serial)) < 1e-9 and dt.time() == datetime.min.time():
        return dt.date().isoformat()
    return dt.isoformat(sep=" ")


class WorkbookError(Exception):
    """Raised for unreadable / unsupported / corrupt workbooks."""


class Workbook:
    """
    Minimal .xlsx reader (reading only - writes go through Excel, below). Loads shared strings, styles, and the
    sheet index up front; sheet data is parsed on demand.
    """

    def __init__(self, path):
        self.path = path
        if not os.path.isfile(path):
            raise WorkbookError("File not found: %s" % path)
        try:
            self._zip = zipfile.ZipFile(path, "r")
        except zipfile.BadZipFile:
            # Legacy .xls and encrypted files are not ZIP containers.
            raise WorkbookError(
                "Not a readable .xlsx. Legacy .xls, .xlsb, or password-"
                "protected files are not supported: %s" % os.path.basename(path)
            )
        names = set(self._zip.namelist())
        if "xl/workbook.xml" not in names:
            raise WorkbookError(
                "Missing xl/workbook.xml -- not a valid .xlsx: %s"
                % os.path.basename(path)
            )
        self.date1904 = False
        self._shared = []            # shared strings (index -> text)
        self._date_style = []        # style index -> bool (is a date format)
        self._sheets = []            # list of {"name", "path"}
        self._load_workbook_index()
        self._load_shared_strings()
        self._load_styles()

    # -- loading helpers ----------------------------------------------------

    def _read_xml(self, member):
        """Parse a member of the zip into an ElementTree root, or None."""
        try:
            data = self._zip.read(member)
        except KeyError:
            return None
        try:
            return ET.fromstring(data)
        except ET.ParseError as exc:
            raise WorkbookError("Corrupt XML in %s: %s" % (member, exc))

    def _load_workbook_index(self):
        root = self._read_xml("xl/workbook.xml")
        if root is None:
            raise WorkbookError("Could not read xl/workbook.xml")

        # Detect the 1904 date system (rare; old Mac-authored files).
        for el in root.iter():
            if _lname(el.tag) == "workbookPr":
                d1904 = _attr_local(el, "date1904")
                if d1904 in ("1", "true", "True"):
                    self.date1904 = True
                break

        # Map relationship id -> target path.
        rels = {}
        rroot = self._read_xml("xl/_rels/workbook.xml.rels")
        if rroot is not None:
            for rel in rroot:
                if _lname(rel.tag) != "Relationship":
                    continue
                rid = rel.get("Id")
                target = rel.get("Target")
                if not rid or not target:
                    continue
                # Normalise target into a full path within the zip.
                if target.startswith("/"):
                    full = target.lstrip("/")
                else:
                    full = "xl/" + target.lstrip("./")
                # Collapse any '..' the rare Target might contain.
                full = os.path.normpath(full).replace(os.sep, "/")
                rels[rid] = full

        # Ordered list of sheets from workbook.xml.
        for el in root.iter():
            if _lname(el.tag) != "sheet":
                continue
            name = el.get("name") or "Sheet"
            rid = _attr_local(el, "id")   # r:id
            spath = rels.get(rid)
            if spath is None:
                # Fallback guess if the relationship is missing.
                log("warning: sheet '%s' has no resolvable relationship" % name)
                continue
            self._sheets.append({"name": name, "path": spath})

        if not self._sheets:
            raise WorkbookError("No worksheets found in %s"
                                % os.path.basename(self.path))

    def _load_shared_strings(self):
        root = self._read_xml("xl/sharedStrings.xml")
        if root is None:
            return
        for si in root:
            if _lname(si.tag) != "si":
                continue
            # Concatenate every <t> descendant (covers plain and rich text).
            parts = [t.text or "" for t in si.iter() if _lname(t.tag) == "t"]
            self._shared.append("".join(parts))

    def _load_styles(self):
        root = self._read_xml("xl/styles.xml")
        if root is None:
            return
        # Custom number formats: numFmtId -> format code.
        custom = {}
        for el in root.iter():
            if _lname(el.tag) == "numFmt":
                fid = el.get("numFmtId")
                code = el.get("formatCode")
                if fid is not None:
                    try:
                        custom[int(fid)] = code or ""
                    except ValueError:
                        pass
        # cellXfs: ordered list; a cell's s="N" indexes into this.
        for el in root.iter():
            if _lname(el.tag) != "cellXfs":
                continue
            for xf in el:
                if _lname(xf.tag) != "xf":
                    continue
                fid_txt = xf.get("numFmtId")
                is_date = False
                if fid_txt is not None:
                    try:
                        fid = int(fid_txt)
                    except ValueError:
                        fid = -1
                    if fid in _BUILTIN_DATE_FMT_IDS:
                        is_date = True
                    elif fid in custom:
                        is_date = _looks_like_date_format(custom[fid])
                self._date_style.append(is_date)
            break

    # -- public API ---------------------------------------------------------

    def sheet_names(self):
        return [s["name"] for s in self._sheets]

    def _resolve_sheet(self, sheet):
        """Resolve a sheet by exact name, case-insensitive name, or 1-based index."""
        if sheet is None:
            return self._sheets[0]
        # Exact match first.
        for s in self._sheets:
            if s["name"] == sheet:
                return s
        # Case-insensitive.
        low = str(sheet).strip().lower()
        for s in self._sheets:
            if s["name"].lower() == low:
                return s
        # 1-based numeric index.
        if str(sheet).isdigit():
            i = int(sheet) - 1
            if 0 <= i < len(self._sheets):
                return self._sheets[i]
        raise WorkbookError(
            "Sheet '%s' not found. Available: %s"
            % (sheet, ", ".join(self.sheet_names()))
        )

    def _cell_value(self, c):
        """Resolve a <c> element to a Python value (str/int/float/bool/None)."""
        ctype = c.get("t")          # cell type
        style = c.get("s")          # style index (for date detection)
        v_el = None
        is_el = None
        for child in c:
            ln = _lname(child.tag)
            if ln == "v":
                v_el = child
            elif ln == "is":
                is_el = child
        v = v_el.text if v_el is not None else None

        if ctype == "s":                       # shared string (v = index)
            try:
                return self._trunc(self._shared[int(v)])
            except (ValueError, IndexError, TypeError):
                return None
        if ctype == "inlineStr":               # inline string in <is>
            if is_el is not None:
                parts = [t.text or "" for t in is_el.iter()
                         if _lname(t.tag) == "t"]
                return self._trunc("".join(parts))
            return None
        if ctype == "str":                     # formula string result
            return self._trunc(v) if v is not None else None
        if ctype == "b":                       # boolean
            return v == "1"
        if ctype == "e":                       # error text, e.g. #DIV/0!
            return self._trunc(v) if v is not None else None

        # Default: numeric (ctype None or "n"). Could be a date by style.
        if v is None or v == "":
            return None
        if style is not None:
            try:
                if self._date_style[int(style)]:
                    iso = _serial_to_isoformat(v, self.date1904)
                    if iso is not None:
                        return iso
            except (ValueError, IndexError):
                pass
        # Plain number: keep ints as ints for clean output.
        try:
            f = float(v)
        except ValueError:
            return self._trunc(v)
        if f.is_integer():
            return int(f)
        return f

    @staticmethod
    def _trunc(text):
        if text is None:
            return None
        text = str(text)
        if len(text) > MAX_CELL_TEXT_LEN:
            return text[:MAX_CELL_TEXT_LEN] + " ...[truncated]"
        return text

    def declared_dimension(self, sheet):
        """Return the sheet's declared used-range ref (e.g. 'A1:H100') or None."""
        s = self._resolve_sheet(sheet)
        root = self._read_xml(s["path"])
        if root is None:
            return None
        for el in root.iter():
            if _lname(el.tag) == "dimension":
                return el.get("ref")
        return None

    def iter_rows(self, sheet, min_row=1, max_row=None, max_col=None):
        """
        Yield (row_number, [values...]) for a sheet. Sparse cells are filled
        with None so column alignment is preserved. Rows are yielded in
        document order; empty trailing cells are trimmed per row.
        """
        s = self._resolve_sheet(sheet)
        root = self._read_xml(s["path"])
        if root is None:
            return
        # Find <sheetData>.
        sheet_data = None
        for el in root:
            if _lname(el.tag) == "sheetData":
                sheet_data = el
                break
        if sheet_data is None:
            return

        for row_el in sheet_data:
            if _lname(row_el.tag) != "row":
                continue
            r_attr = row_el.get("r")
            row_num = int(r_attr) if r_attr and r_attr.isdigit() else None
            if row_num is None:
                continue
            if row_num < min_row:
                continue
            if max_row is not None and row_num > max_row:
                break

            # Place cells by column index into a dict, then flatten.
            cells = {}
            highest = -1
            for c in row_el:
                if _lname(c.tag) != "c":
                    continue
                ref = c.get("r")
                if ref:
                    col_letters, _ = _split_cell_ref(ref)
                    col_idx = _col_letters_to_index(col_letters)
                else:
                    col_idx = highest + 1  # positional fallback
                if max_col is not None and col_idx >= max_col:
                    continue
                cells[col_idx] = self._cell_value(c)
                if col_idx > highest:
                    highest = col_idx

            if highest < 0:
                yield row_num, []
                continue
            width = highest + 1
            values = [cells.get(i) for i in range(width)]
            yield row_num, values

    # -- relationships, Tables and pivot tables ------------------------------
    #
    # A Table (Insert > Table, a "ListObject" to VBA) and a pivot table are
    # separate XML parts linked from their sheet by a relationship:
    #
    #   xl/worksheets/sheet1.xml
    #     -> xl/worksheets/_rels/sheet1.xml.rels
    #          -> ../tables/table1.xml            (the Table: name, range,
    #                                              column names)
    #          -> ../pivotTables/pivotTable1.xml  (the pivot: layout, fields)
    #               -> ../pivotCache/pivotCacheDefinition1.xml
    #                                             (source range, field names)
    #
    # The CELLS a pivot table shows are ordinary cells in the sheet (Excel
    # saves the rendered pivot into the grid), so its values are read like any
    # other range - as of the last time Excel refreshed and saved it.

    def _relationships(self, part):
        """
        The relationships of a part: a list of (type_suffix, full_path). The
        type is shortened to its last path segment ('table', 'pivotTable',
        'pivotCacheDefinition').
        """
        folder, name = part.rsplit("/", 1) if "/" in part else ("", part)
        rels_path = (folder + "/" if folder else "") + "_rels/" + name + ".rels"
        root = self._read_xml(rels_path)
        out = []
        if root is None:
            return out
        for rel in root:
            if _lname(rel.tag) != "Relationship":
                continue
            target = rel.get("Target") or ""
            rtype = (rel.get("Type") or "").rsplit("/", 1)[-1]
            if not target or rel.get("TargetMode") == "External":
                continue
            if target.startswith("/"):
                full = target.lstrip("/")
            else:
                full = os.path.normpath(
                    os.path.join(folder, target)).replace(os.sep, "/")
            out.append((rtype, full))
        return out

    def tables(self):
        """
        Every Table in the workbook, in sheet order, as dicts:
        {name, sheet, ref, columns, header_rows, totals_rows}.
        """
        found = []
        for sheet in self._sheets:
            for rtype, target in self._relationships(sheet["path"]):
                if rtype != "table":
                    continue
                root = self._read_xml(target)
                if root is None:
                    continue
                columns = [col.get("name") or "" for col in root.iter()
                           if _lname(col.tag) == "tableColumn"]
                try:
                    header_rows = int(root.get("headerRowCount", "1"))
                    totals_rows = int(root.get("totalsRowCount", "0"))
                except ValueError:
                    header_rows, totals_rows = 1, 0
                found.append({
                    "name": root.get("displayName") or root.get("name") or "?",
                    "sheet": sheet["name"],
                    "ref": root.get("ref") or "",
                    "columns": columns,
                    "header_rows": header_rows,
                    "totals_rows": totals_rows,
                })
        return found

    def find_table(self, name):
        """A Table by name (exact, then case-insensitive), or WorkbookError."""
        tables = self.tables()
        want = str(name or "").strip()
        for t in tables:
            if t["name"] == want:
                return t
        for t in tables:
            if t["name"].lower() == want.lower():
                return t
        raise WorkbookError(
            "No Table named '%s' in %s. Tables: %s"
            % (want, os.path.basename(self.path),
               ", ".join("%s (sheet '%s')" % (t["name"], t["sheet"])
                         for t in tables) or "none - this workbook has no "
               "Tables (Insert > Table); read its sheets instead"))

    @staticmethod
    def _shared_items(cache_field):
        """The values of a pivot cache field's <sharedItems>, in order."""
        items = []
        for el in cache_field:
            if _lname(el.tag) != "sharedItems":
                continue
            for item in el:
                kind = _lname(item.tag)
                if kind == "m":
                    items.append("(blank)")
                else:
                    items.append(item.get("v"))
        return items

    def pivot_tables(self):
        """
        Every pivot table in the workbook, as dicts describing its layout:
        {name, sheet, location, source, rows, columns, filters, values,
        refreshed, hidden_items}. Field names come from the pivot cache.
        """
        found = []
        for sheet in self._sheets:
            for rtype, target in self._relationships(sheet["path"]):
                if rtype != "pivotTable":
                    continue
                root = self._read_xml(target)
                if root is None:
                    continue
                cache_root = None
                for ctype, ctarget in self._relationships(target):
                    if ctype == "pivotCacheDefinition":
                        cache_root = self._read_xml(ctarget)
                        break
                found.append(self._describe_pivot(sheet["name"], root, cache_root))
        return found

    def _describe_pivot(self, sheet_name, root, cache_root):
        # Field names and each field's distinct values, from the cache.
        cache_fields = []
        source = "(unknown source)"
        refreshed = None
        if cache_root is not None:
            for el in cache_root.iter():
                if _lname(el.tag) == "cacheField":
                    cache_fields.append({"name": el.get("name") or "?",
                                         "items": self._shared_items(el)})
            for el in cache_root.iter():
                if _lname(el.tag) == "cacheSource":
                    kind = el.get("type") or "worksheet"
                    ws_src = next((c for c in el if _lname(c.tag) == "worksheetSource"), None)
                    if ws_src is not None and ws_src.get("name"):
                        source = "Table/name '%s'" % ws_src.get("name")
                    elif ws_src is not None:
                        source = "'%s'!%s" % (ws_src.get("sheet") or sheet_name,
                                              ws_src.get("ref") or "?")
                    else:
                        source = "%s data source" % kind
                    break
            stamp = cache_root.get("refreshedDate")
            if stamp:
                refreshed = _serial_to_isoformat(stamp, False)

        def field_name(index):
            try:
                index = int(index)
            except (TypeError, ValueError):
                return "?"
            if index == -2:
                return "(Values)"      # the pseudo-field holding the data fields
            if 0 <= index < len(cache_fields):
                return cache_fields[index]["name"]
            return "field %d" % index

        # Per pivot field: the cache item index behind each of its items, and
        # how many items are hidden (unticked in a row/column filter).
        pivot_fields = []
        for el in root.iter():
            if _lname(el.tag) != "pivotFields":
                continue
            for pf in el:
                if _lname(pf.tag) != "pivotField":
                    continue
                item_x, hidden = [], 0
                for items in pf:
                    if _lname(items.tag) != "items":
                        continue
                    for it in items:
                        item_x.append(it.get("x"))
                        if it.get("h") in ("1", "true"):
                            hidden += 1
                pivot_fields.append({"items": item_x, "hidden": hidden})
            break

        def fields_of(container, attr):
            names = []
            for el in root.iter():
                if _lname(el.tag) != container:
                    continue
                for f in el:
                    names.append(f.get(attr))
                break
            return names

        rows = [field_name(x) for x in fields_of("rowFields", "x")]
        cols = [field_name(x) for x in fields_of("colFields", "x")]

        filters = []
        for el in root.iter():
            if _lname(el.tag) != "pageFields":
                continue
            for pf in el:
                fld = pf.get("fld")
                label = field_name(fld)
                chosen = "(All)"
                item = pf.get("item")
                try:
                    fidx, iidx = int(fld), int(item)
                    x = pivot_fields[fidx]["items"][iidx]
                    chosen = cache_fields[fidx]["items"][int(x)]
                except (TypeError, ValueError, IndexError):
                    if item is not None:
                        chosen = "(one item selected)"
                filters.append("%s = %s" % (label, chosen))
            break

        values = []
        for el in root.iter():
            if _lname(el.tag) != "dataFields":
                continue
            for df in el:
                values.append("%s (%s of %s)" % (
                    df.get("name") or "?", df.get("subtotal") or "sum",
                    field_name(df.get("fld"))))
            break

        location = ""
        for el in root.iter():
            if _lname(el.tag) == "location":
                location = el.get("ref") or ""
                break

        hidden = []
        for index, pf in enumerate(pivot_fields):
            if pf["hidden"]:
                hidden.append("%s: %d item(s) hidden" % (field_name(index), pf["hidden"]))

        return {
            "name": root.get("name") or "?",
            "sheet": sheet_name,
            "location": location,
            "source": source,
            "rows": rows, "columns": cols, "filters": filters, "values": values,
            "refreshed": refreshed, "hidden_items": hidden,
        }

    def find_pivot(self, name):
        """A pivot table by name, or the only one when name is omitted."""
        pivots = self.pivot_tables()
        if not pivots:
            raise WorkbookError("%s has no pivot tables."
                                % os.path.basename(self.path))
        want = str(name or "").strip()
        if not want:
            if len(pivots) == 1:
                return pivots[0]
            raise WorkbookError(
                "This workbook has %d pivot tables - name one: %s"
                % (len(pivots), ", ".join("%s (sheet '%s')" % (p["name"], p["sheet"])
                                          for p in pivots)))
        for p in pivots:
            if p["name"] == want or p["name"].lower() == want.lower():
                return p
        raise WorkbookError(
            "No pivot table named '%s'. Pivot tables: %s"
            % (want, ", ".join("%s (sheet '%s')" % (p["name"], p["sheet"])
                               for p in pivots)))

    def read_grid(self, sheet, ref, max_rows=None):
        """
        The cells of an A1 range as a list of rows (lists of equal width),
        blanks as None. Returns (rows, truncated).
        """
        parsed = _parse_a1_range(ref)
        if not parsed or None in parsed:
            raise WorkbookError("Not a complete A1 range: %r" % (ref,))
        min_row, max_row, min_col, max_col = parsed
        width = max_col - min_col + 1
        by_row = {}
        stop = max_row if max_rows is None else min(max_row, min_row + max_rows - 1)
        for row_num, values in self.iter_rows(sheet, min_row=min_row,
                                              max_row=stop, max_col=max_col + 1):
            by_row[row_num] = (values[min_col:] + [None] * width)[:width]
        rows = [by_row.get(r, [None] * width) for r in range(min_row, stop + 1)]
        return rows, stop < max_row

    def close(self):
        try:
            self._zip.close()
        except Exception:
            pass


# ===========================================================================
# Workbook folder resolution
# ===========================================================================

def list_workbook_files(folder):
    """
    Return sorted list of readable workbook file names in the folder.
    Files that RESOLVE outside the folder (e.g. a symlink pointing elsewhere)
    are excluded, so a link dropped into the folder cannot expose workbooks
    beyond it.
    """
    if not folder:
        raise WorkbookError(
            "No workbook folder configured. Set EVA_DOCUMENTS_DIR (the "
            "server reads its 'excel' sub-folder) or EXCEL_DOCS_DIR."
        )
    if not os.path.isdir(folder):
        raise WorkbookError("Workbook folder does not exist: %s" % folder)
    real_base = os.path.realpath(folder)
    out = []
    for name in os.listdir(folder):
        if name.startswith("~$"):          # Excel lock/temp files
            continue
        ext = os.path.splitext(name)[1].lower()
        full = os.path.join(folder, name)
        if ext in ALLOWED_EXTENSIONS and os.path.isfile(full):
            try:
                real = os.path.realpath(full)
                contained = os.path.commonpath([real, real_base]) == real_base
            except ValueError:  # different drives on Windows
                contained = False
            if not contained:
                log("excluded (resolves outside the workbook folder): %s" % full)
                continue
            out.append(name)
    return sorted(out)


def _normalise_name(text):
    """Lowercase, turn separators into spaces and collapse whitespace, so fuzzy
    matching ignores punctuation/case differences between a query and a name."""
    text = text.lower()
    for ch in ("_", "-", ".", "(", ")", "[", "]", ",", "&", "+"):
        text = text.replace(ch, " ")
    return " ".join(text.split())


def _score_name(query_norm, name_norm):
    """Score a query against a name as (token-containment, sequence-ratio): the
    fraction of query words present in the name, then difflib's overall
    similarity. Ranking on the pair prefers names that contain the query words
    and, among those, the closest overall string."""
    q_tokens = query_norm.split()
    n_tokens = set(name_norm.split())
    contained = (sum(1 for t in q_tokens if t in n_tokens) / len(q_tokens)) if q_tokens else 0.0
    ratio = difflib.SequenceMatcher(None, query_norm, name_norm).ratio()
    return (contained, ratio)


def _fuzzy_match_workbook(files, requested):
    """
    Fuzzy-match `requested` against workbook file names. Returns (best, tied):
    best is the single closest name, or None when nothing is close enough;
    `tied` lists the near-equal candidates when the match is too ambiguous to
    pick one. Matching is on the file stem (name without extension).
    """
    query_norm = _normalise_name(os.path.splitext(os.path.basename(str(requested)))[0])
    if not query_norm or not files:
        return None, []
    scored = [
        (f, _score_name(query_norm, _normalise_name(os.path.splitext(f)[0])))
        for f in files
    ]
    scored.sort(key=lambda item: item[1], reverse=True)
    best, (best_contained, best_ratio) = scored[0]
    # Nothing shares a word and the closest name is still too different.
    if best_contained == 0.0 and best_ratio < FUZZY_MIN_RATIO:
        return None, []
    tied = []
    for f, (contained, ratio) in scored:
        if contained == best_contained and (best_ratio - ratio) <= FUZZY_AMBIGUITY_DELTA:
            tied.append(f)
        else:
            break
    if len(tied) > 1:
        return None, tied
    return best, []


def resolve_workbook_path(folder, requested):
    """
    Resolve a loosely specified workbook name to a full path.
    Matching order: exact -> exact+ext -> case-insensitive -> substring ->
    FUZZY (difflib name match, last resort). Rejects any path that escapes the
    configured folder.
    """
    files = list_workbook_files(folder)
    if not files:
        raise WorkbookError("No workbooks found in %s" % folder)

    req = os.path.basename(str(requested).strip())  # strip any path parts
    req_low = req.lower()

    # Exact filename match.
    for f in files:
        if f == req:
            return os.path.join(folder, f)
    # Match ignoring a missing extension.
    for f in files:
        if os.path.splitext(f)[0].lower() == req_low:
            return os.path.join(folder, f)
    for f in files:
        if f.lower() == req_low:
            return os.path.join(folder, f)
    # Substring (unique) match.
    hits = [f for f in files if req_low in f.lower()]
    if len(hits) == 1:
        return os.path.join(folder, hits[0])
    if len(hits) > 1:
        raise WorkbookError(
            "Workbook '%s' is ambiguous. Candidates: %s"
            % (requested, ", ".join(hits))
        )

    # No exact/substring match - fall back to a fuzzy name match, so a near-miss
    # like "budgit" or "q3 budget" still finds "Budget Q3 2024.xlsx".
    best, tied = _fuzzy_match_workbook(files, requested)
    if best is not None:
        log("fuzzy-matched workbook '%s' -> %s" % (requested, best))
        return os.path.join(folder, best)
    if tied:
        raise WorkbookError(
            "Workbook '%s' matches several about equally well: %s. Be more "
            "specific or use excel_list_workbooks." % (requested, ", ".join(tied))
        )
    raise WorkbookError(
        "Workbook '%s' not found. Available: %s"
        % (requested, ", ".join(files))
    )


# ===========================================================================
# Tool implementations. Each returns a plain string (rendered to the model).
# ===========================================================================

def _open(folder, workbook):
    return Workbook(resolve_workbook_path(folder, workbook))


def tool_list_workbooks(folder, args):
    files = list_workbook_files(folder)
    if not files:
        return "No workbooks found in %s" % folder
    lines = ["Workbooks in %s:" % folder]
    for f in files:
        full = os.path.join(folder, f)
        try:
            size = os.path.getsize(full)
            lines.append("  - %s (%d bytes)" % (f, size))
        except OSError:
            lines.append("  - %s" % f)
    return "\n".join(lines)


def tool_list_sheets(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        lines = ["Workbook: %s" % os.path.basename(wb.path)]
        if wb.date1904:
            lines.append("(uses the 1904 date system)")
        lines.append("Sheets:")
        # Tables and pivots are extras here: a malformed part must not stop
        # the sheets themselves being listed.
        try:
            tables = wb.tables()
        except WorkbookError as exc:
            tables = []
            lines.append("(Tables could not be read: %s)" % exc)
        try:
            pivots = wb.pivot_tables()
        except WorkbookError as exc:
            pivots = []
            lines.append("(Pivot tables could not be read: %s)" % exc)
        for name in wb.sheet_names():
            dim = wb.declared_dimension(name)
            dim_txt = " declared range %s" % dim if dim else " (no declared range)"
            lines.append("  - %s:%s" % (name, dim_txt))
            here = [t["name"] for t in tables if t["sheet"] == name]
            if here:
                lines.append("      Tables: %s" % ", ".join(here))
            here = [p["name"] for p in pivots if p["sheet"] == name]
            if here:
                lines.append("      Pivot tables: %s" % ", ".join(here))
        return "\n".join(lines)
    finally:
        wb.close()


def tool_get_headers(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        sheet = args.get("sheet")
        header_row = int(args.get("header_row", 1) or 1)
        for row_num, values in wb.iter_rows(sheet, min_row=header_row,
                                             max_row=header_row,
                                             max_col=MAX_COLS_PER_READ):
            cols = []
            for i, val in enumerate(values):
                letter = _col_index_to_letters(i)
                cols.append("%s=%s" % (letter, "" if val is None else val))
            resolved = wb._resolve_sheet(sheet)["name"]
            if not cols:
                return "Header row %d of sheet '%s' is empty." % (header_row, resolved)
            return ("Headers for sheet '%s' (row %d):\n  "
                    % (resolved, header_row)) + "\n  ".join(cols)
        resolved = wb._resolve_sheet(sheet)["name"]
        return "Sheet '%s' has no row %d." % (resolved, header_row)
    finally:
        wb.close()


def _parse_a1_range(rng):
    """
    Parse 'A1:D20' (or 'A1') into (min_row, max_row, min_col, max_col),
    using None for open ends. Returns None if rng is falsy.
    """
    if not rng:
        return None
    rng = str(rng).replace(" ", "")
    if ":" in rng:
        start, end = rng.split(":", 1)
    else:
        start = end = rng
    sc, sr = _split_cell_ref(start)
    ec, er = _split_cell_ref(end)
    min_col = _col_letters_to_index(sc) if sc else None
    max_col = _col_letters_to_index(ec) if ec else None
    if (min_col is not None and max_col is not None and min_col > max_col):
        min_col, max_col = max_col, min_col
    if (sr is not None and er is not None and sr > er):
        sr, er = er, sr
    return sr, er, min_col, max_col


def tool_read_range(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        sheet = args.get("sheet")
        parsed = _parse_a1_range(args.get("range"))
        if parsed:
            min_row, max_row, min_col, max_col = parsed
            min_row = min_row or 1
        else:
            min_row, max_row, min_col, max_col = 1, None, None, None

        # Apply safety caps.
        eff_max_col = MAX_COLS_PER_READ
        if max_col is not None:
            eff_max_col = min(MAX_COLS_PER_READ, max_col + 1)
        row_cap = MAX_ROWS_PER_READ

        out_rows = []
        truncated = False
        resolved = wb._resolve_sheet(sheet)["name"]
        for row_num, values in wb.iter_rows(sheet, min_row=min_row,
                                            max_row=max_row,
                                            max_col=eff_max_col):
            # Trim to requested left column if a range was given.
            if min_col is not None:
                values = values[min_col:] if min_col < len(values) else []
                start_letter = _col_index_to_letters(min_col)
            else:
                start_letter = "A"
            cell_texts = []
            for i, val in enumerate(values):
                letter = _col_index_to_letters(
                    (_col_letters_to_index(start_letter) + i))
                cell_texts.append("%s%d=%s"
                                  % (letter, row_num,
                                     "" if val is None else val))
            out_rows.append("  " + " | ".join(cell_texts) if cell_texts
                            else "  (row %d empty)" % row_num)
            if len(out_rows) >= row_cap:
                truncated = True
                break

        if not out_rows:
            return "No data found in sheet '%s' for the requested range." % resolved
        header = "Sheet '%s'" % resolved
        if args.get("range"):
            header += " range %s" % args.get("range")
        if truncated:
            header += " (truncated to %d rows)" % row_cap
        return header + ":\n" + "\n".join(out_rows)
    finally:
        wb.close()


def tool_search(folder, args):
    query = args.get("query")
    if query is None or str(query) == "":
        return "Error: 'query' is required."
    query_low = str(query).lower()
    case_sensitive = bool(args.get("case_sensitive", False))
    target_sheet = args.get("sheet")   # optional: restrict to one sheet

    wb = _open(folder, args.get("workbook"))
    try:
        sheets = ([wb._resolve_sheet(target_sheet)["name"]]
                  if target_sheet else wb.sheet_names())
        hits = []
        for sname in sheets:
            for row_num, values in wb.iter_rows(sname,
                                                max_col=MAX_COLS_PER_READ):
                for i, val in enumerate(values):
                    if val is None:
                        continue
                    hay = str(val)
                    needle = str(query)
                    found = (needle in hay) if case_sensitive \
                        else (query_low in hay.lower())
                    if found:
                        ref = "%s%d" % (_col_index_to_letters(i), row_num)
                        hits.append("  [%s] %s = %s" % (sname, ref, hay))
                        if len(hits) >= MAX_SEARCH_HITS:
                            break
                if len(hits) >= MAX_SEARCH_HITS:
                    break
            if len(hits) >= MAX_SEARCH_HITS:
                break
        if not hits:
            return "No cells matched '%s'." % query
        head = "Found %d match(es) for '%s'" % (len(hits), query)
        if len(hits) >= MAX_SEARCH_HITS:
            head += " (stopped at cap %d)" % MAX_SEARCH_HITS
        return head + ":\n" + "\n".join(hits)
    finally:
        wb.close()


def _resolve_column_index(wb, sheet, column, header_row):
    """
    Turn a column spec into a 0-based index. Accepts a column letter
    ('C'), a 1-based number ('3'), or a header name matched in header_row.
    """
    spec = str(column).strip()

    # 1-based column number, e.g. "3".
    if spec.isdigit():
        return int(spec) - 1

    header_map = _header_map(wb, sheet, header_row)

    # Header name match (case-insensitive) takes priority over letter parsing,
    # so a column literally named "C" still resolves to that header.
    if spec.lower() in header_map:
        return header_map[spec.lower()]

    # Otherwise, only treat it as a column letter if it plausibly IS one
    # (short and all letters). This avoids silently mapping an unmatched
    # header name like "Cost" onto some far-off column.
    if spec.isalpha() and len(spec) <= 3:
        return _col_letters_to_index(spec)

    raise WorkbookError(
        "Column '%s' not found. Provide a column letter (e.g. C), a 1-based "
        "number (e.g. 3), or one of these headers: %s"
        % (column, ", ".join(sorted(header_map.keys())) or "(none)")
    )


def _header_map(wb, sheet, header_row):
    """Map lower-cased header text -> column index for the given header row."""
    mapping = {}
    for _rn, values in wb.iter_rows(sheet, min_row=header_row,
                                    max_row=header_row,
                                    max_col=MAX_COLS_PER_READ):
        for i, val in enumerate(values):
            if val is not None and str(val).strip() != "":
                mapping[str(val).strip().lower()] = i
    return mapping


def tool_column_stats(folder, args):
    if "column" not in args:
        return "Error: 'column' is required (letter, number, or header name)."
    wb = _open(folder, args.get("workbook"))
    try:
        sheet = args.get("sheet")
        header_row = int(args.get("header_row", 1) or 1)
        data_start = int(args.get("data_start_row", header_row + 1)
                         or header_row + 1)
        col_idx = _resolve_column_index(wb, sheet, args["column"], header_row)
        resolved = wb._resolve_sheet(sheet)["name"]

        total = 0
        non_empty = 0
        numbers = []
        distinct = {}
        for row_num, values in wb.iter_rows(sheet, min_row=data_start):
            if col_idx >= len(values):
                total += 1
                continue
            val = values[col_idx]
            total += 1
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue
            non_empty += 1
            if isinstance(val, bool):
                key = str(val)
            elif isinstance(val, (int, float)):
                numbers.append(float(val))
                key = repr(val)
            else:
                key = str(val)
            distinct[key] = distinct.get(key, 0) + 1

        lines = ["Column stats for '%s' in sheet '%s' (data from row %d):"
                 % (args["column"], resolved, data_start)]
        lines.append("  rows scanned : %d" % total)
        lines.append("  non-empty    : %d" % non_empty)
        lines.append("  empty        : %d" % (total - non_empty))
        lines.append("  distinct     : %d" % len(distinct))

        if numbers:
            n = len(numbers)
            s = sum(numbers)
            mean = s / n
            srt = sorted(numbers)
            mid = n // 2
            median = srt[mid] if n % 2 else (srt[mid - 1] + srt[mid]) / 2.0
            lines.append("  numeric count: %d" % n)
            lines.append("  sum          : %s" % _fmt_num(s))
            lines.append("  mean         : %s" % _fmt_num(mean))
            lines.append("  median       : %s" % _fmt_num(median))
            lines.append("  min          : %s" % _fmt_num(srt[0]))
            lines.append("  max          : %s" % _fmt_num(srt[-1]))
        else:
            lines.append("  (no numeric values found in this column)")

        # Top categorical values (handy for non-numeric columns).
        if distinct and not numbers:
            top = sorted(distinct.items(), key=lambda kv: kv[1], reverse=True)[:10]
            lines.append("  top values   :")
            for k, cnt in top:
                lines.append("      %s (%d)" % (k, cnt))
        return "\n".join(lines)
    finally:
        wb.close()


def _md_cell(value):
    """One Markdown table cell: blank for None, single line, pipes escaped."""
    if value is None:
        return ""
    return " ".join(str(value).split()).replace("|", "\\|")


def _md_table(header, rows):
    """Render a header + rows as a Markdown table."""
    width = max([len(header)] + [len(r) for r in rows]) if (header or rows) else 0
    header = (list(header) + [""] * width)[:width]
    out = ["| " + " | ".join(_md_cell(h) for h in header) + " |",
           "| " + " | ".join("---" for _ in header) + " |"]
    for row in rows:
        row = (list(row) + [None] * width)[:width]
        out.append("| " + " | ".join(_md_cell(v) for v in row) + " |")
    return "\n".join(out)


def _values_equal(cell, wanted):
    """
    Loose equality for filtering/matching rows: numbers compare as numbers
    (so 3 matches 3.0), everything else as trimmed, case-insensitive text.
    """
    if cell is None:
        return wanted is None or str(wanted).strip() == ""
    try:
        return float(cell) == float(wanted)
    except (TypeError, ValueError):
        pass
    return str(cell).strip().lower() == str(wanted).strip().lower()


def _table_layout(table):
    """(first_row, last_row, first_col, last_col) of a Table's full range."""
    parsed = _parse_a1_range(table["ref"])
    if not parsed or None in parsed:
        raise WorkbookError("Table '%s' has an unreadable range %r."
                            % (table["name"], table["ref"]))
    return parsed


def tool_list_tables(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        tables = wb.tables()
        name = os.path.basename(wb.path)
        if not tables:
            return ("%s has no Tables (Insert > Table). Read its sheets with "
                    "excel_read_range instead." % name)
        lines = ["Tables in %s:" % name]
        for t in tables:
            first, last, _c1, _c2 = _table_layout(t)
            data_rows = max(0, (last - first + 1) - t["header_rows"] - t["totals_rows"])
            lines.append("  - %s  (sheet '%s', range %s, %d data row(s)%s)"
                         % (t["name"], t["sheet"], t["ref"], data_rows,
                            ", totals row" if t["totals_rows"] else ""))
            lines.append("      columns: %s" % ", ".join(t["columns"]))
        return "\n".join(lines)
    finally:
        wb.close()


def tool_read_table(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        table = wb.find_table(args.get("table"))
        first, last, c1, c2 = _table_layout(table)
        grid, _ = wb.read_grid(table["sheet"], table["ref"])
        header = table["columns"] or (grid[0] if grid else [])
        body = grid[table["header_rows"]:]
        totals = []
        if table["totals_rows"]:
            totals = body[-table["totals_rows"]:]
            body = body[:-table["totals_rows"]]
        # Sheet row number of each data row, so an answer can cite the cell.
        first_data = first + table["header_rows"]
        numbered = [(first_data + i, row) for i, row in enumerate(body)]

        # Optional filter: keep rows whose column equals a value.
        fcol, fval = args.get("filter_column"), args.get("filter_value")
        if fcol:
            lookup = [str(h).strip().lower() for h in header]
            if str(fcol).strip().lower() not in lookup:
                raise WorkbookError("Table '%s' has no column '%s'. Columns: %s"
                                    % (table["name"], fcol, ", ".join(header)))
            idx = lookup.index(str(fcol).strip().lower())
            numbered = [(r, row) for r, row in numbered
                        if idx < len(row) and _values_equal(row[idx], fval)]

        # Optional column subset, in the order asked for.
        wanted = args.get("columns")
        if wanted:
            if isinstance(wanted, str):
                wanted = [w for w in wanted.split(",")]
            lookup = [str(h).strip().lower() for h in header]
            picks = []
            for w in wanted:
                key = str(w).strip().lower()
                if key not in lookup:
                    raise WorkbookError("Table '%s' has no column '%s'. Columns: %s"
                                        % (table["name"], w, ", ".join(header)))
                picks.append(lookup.index(key))
            header = [header[i] for i in picks]
            numbered = [(r, [row[i] if i < len(row) else None for i in picks])
                        for r, row in numbered]
            totals = [[row[i] if i < len(row) else None for i in picks] for row in totals]

        try:
            offset = max(0, int(args.get("offset") or 0))
        except (TypeError, ValueError):
            offset = 0
        total = len(numbered)
        page = numbered[offset:offset + MAX_ROWS_PER_READ]

        head = ("Table '%s' (sheet '%s', range %s): %d data row(s)%s"
                % (table["name"], table["sheet"], table["ref"], total,
                   " matching %s = %s" % (fcol, fval) if fcol else ""))
        if not page:
            return head + ("\n(no rows)" if not offset else
                           "\n(no rows from offset %d)" % offset)
        shown = "showing rows %d-%d" % (offset + 1, offset + len(page))
        lines = [head + ", " + shown + ". 'Row' is the sheet row number.", ""]
        lines.append(_md_table(["Row"] + list(header),
                               [[r] + list(row) for r, row in page]))
        if totals and offset + len(page) >= total:
            lines.append("")
            lines.append("Totals row: " + " | ".join(_md_cell(v) for v in totals[0]))
        if offset + len(page) < total:
            lines.append("")
            lines.append("[%d more row(s): call again with offset=%d.]"
                         % (total - offset - len(page), offset + len(page)))
        return "\n".join(lines)
    finally:
        wb.close()


def _pivot_summary(p):
    lines = ["Pivot table '%s' (sheet '%s', cells %s)"
             % (p["name"], p["sheet"], p["location"] or "?"),
             "  source : %s" % p["source"],
             "  rows   : %s" % (", ".join(p["rows"]) or "-"),
             "  columns: %s" % (", ".join(p["columns"]) or "-"),
             "  values : %s" % (", ".join(p["values"]) or "-"),
             "  filters: %s" % (", ".join(p["filters"]) or "-")]
    if p["hidden_items"]:
        lines.append("  hidden : %s" % "; ".join(p["hidden_items"]))
    if p["refreshed"]:
        lines.append("  last refreshed: %s" % p["refreshed"])
    return lines


def tool_list_pivot_tables(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        pivots = wb.pivot_tables()
        name = os.path.basename(wb.path)
        if not pivots:
            return "%s has no pivot tables." % name
        lines = ["Pivot tables in %s:" % name]
        for p in pivots:
            lines.append("")
            lines.extend(_pivot_summary(p))
        return "\n".join(lines)
    finally:
        wb.close()


def tool_read_pivot_table(folder, args):
    wb = _open(folder, args.get("workbook"))
    try:
        p = wb.find_pivot(args.get("pivot"))
        lines = _pivot_summary(p)
        if not p["location"]:
            lines.append("")
            lines.append("(The pivot table has no saved location, so there are "
                         "no cells to read.)")
            return "\n".join(lines)
        # The page-filter area sits ABOVE the location ref; the ref itself is
        # the body of the pivot, header row first.
        grid, truncated = wb.read_grid(p["sheet"], p["location"],
                                       max_rows=MAX_ROWS_PER_READ)
        lines.append("")
        lines.append("Values as Excel last calculated and saved them%s:"
                     % (" (first %d rows)" % MAX_ROWS_PER_READ if truncated else ""))
        if grid:
            lines.append(_md_table(grid[0], grid[1:]))
        else:
            lines.append("(empty)")
        lines.append("")
        lines.append("If the source data has changed since the last refresh, "
                     "these figures are stale until the pivot is refreshed in "
                     "Excel.")
        return "\n".join(lines)
    finally:
        wb.close()


def _fmt_num(x):
    """Render a float cleanly (drop trailing .0 for whole numbers)."""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return "%.6g" % x if isinstance(x, float) else str(x)


# ===========================================================================
# Writing: driving the installed Excel through COM
# ===========================================================================
# Everything above reads the .xlsx file directly. Writing does NOT: a workbook
# is a web of parts that must agree (shared strings, styles, the calc chain,
# Table ranges, pivot caches, charts), and Excel itself is the only thing that
# keeps them all consistent - and the only thing that can build a real
# PivotTable. So each write starts a private, invisible Excel (DispatchEx, never
# the user's own window), opens the one workbook, makes the change, lets Excel
# recalculate, saves, and quits. A failure part-way closes the workbook
# WITHOUT saving, so a half-made change never reaches the file.
#
# Needs Windows, desktop Excel, and pywin32 (pip install pywin32), imported
# only when a write tool runs - reading still needs nothing but the standard
# library, and --version works without it.

# Excel constants used below (from the Excel type library; late binding has no
# access to the names).
XL_DATABASE = 1                  # PivotCaches.Create SourceType
XL_ROW_FIELD, XL_COLUMN_FIELD, XL_PAGE_FIELD = 1, 2, 3
XL_OPENXML_WORKBOOK = 51         # .xlsx
XL_OPENXML_MACRO_WORKBOOK = 52   # .xlsm
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3   # never run macros on open
PIVOT_FUNCTIONS = {
    "sum": -4157, "count": -4112, "average": -4106, "avg": -4106,
    "mean": -4106, "max": -4136, "min": -4139, "product": -4149,
    "countnums": -4113, "count_numbers": -4113, "stdev": -4155,
    "stdevp": -4156, "var": -4164, "varp": -4165,
}
MAX_WRITE_CELLS = 20000          # cells one write call may set
MAX_PIVOT_PREVIEW_ROWS = 60      # rows of a new pivot echoed back



def _com_modules():
    """Import pywin32 on demand, with an actionable error when it is missing."""
    if os.name != "nt":
        raise WorkbookError("Writing to a workbook drives Microsoft Excel, "
                            "which needs Windows.")
    try:
        import pythoncom                 # noqa: F401  (pywin32)
        import pywintypes                # noqa: F401
        import win32com.client           # noqa: F401
    except ImportError:
        raise WorkbookError("Writing to a workbook needs pywin32. Install it "
                            "into this Python:  pip install pywin32  - then "
                            "restart the client.")
    return pythoncom, pywintypes, win32com.client


def _com_message(exc):
    """The readable part of a pywintypes.com_error."""
    try:
        info = exc.excepinfo
        if info and info[2]:
            return str(info[2]).strip()
    except AttributeError:
        pass
    try:
        return str(exc.args[1])
    except (AttributeError, IndexError):
        return str(exc)


def _check_not_locked(path):
    """
    Refuse early when the workbook is open somewhere (usually in the user's
    own Excel): a second Excel would only get a read-only copy, and the change
    could not be saved.
    """
    folder, name = os.path.split(path)
    if os.path.exists(os.path.join(folder, "~$" + name)):
        raise WorkbookError("%s appears to be open in Excel (its ~$ lock file "
                            "exists). Close it there, then try again." % name)
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        raise WorkbookError("%s is locked by another program (probably open in "
                            "Excel). Close it, then try again." % name)
    except OSError as exc:
        raise WorkbookError("Cannot open %s for writing: %s" % (name, exc))


def _save_as_path(folder, source_path, save_as, overwrite):
    """
    Validate a 'save_as' name: inside the workbook folder, the same kind of
    file as the source, and not replacing an existing workbook unless asked.
    Returns the full path, or None when the source is to be saved in place.
    """
    if not save_as:
        return None
    # basename() so a path in 'save_as' cannot leave the folder.
    name = os.path.basename(str(save_as).replace("\\", "/")).strip()
    if not name or name in (".", ".."):
        raise WorkbookError("'save_as' must be a file name.")
    src_ext = os.path.splitext(source_path)[1].lower()
    root, ext = os.path.splitext(name)
    if not ext:
        name, ext = name + src_ext, src_ext
    if ext.lower() not in ALLOWED_EXTENSIONS:
        raise WorkbookError("'save_as' must end in .xlsx or .xlsm.")
    if src_ext == ".xlsm" and ext.lower() == ".xlsx":
        raise WorkbookError("Saving a macro workbook (.xlsm) as .xlsx would "
                            "strip its macros - keep the .xlsm extension.")
    path = os.path.join(folder, name)
    if os.path.abspath(path) == os.path.abspath(source_path):
        return None
    if os.path.exists(path) and not overwrite:
        raise WorkbookError("%s already exists. Choose another 'save_as' name, "
                            "or pass overwrite=true to replace it." % name)
    return path


class ExcelSession:
    """
    One workbook open in a private, invisible Excel, for one tool call.

        with ExcelSession(path) as xl:
            ws = xl.sheet("Data")
            ...
            xl.commit(save_as_path)      # save; without it nothing is kept

    Leaving the block without commit() closes the workbook unsaved.
    """

    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.app = None
        self.wb = None
        self._pythoncom = None
        self._com_error = Exception

    def __enter__(self):
        pythoncom, pywintypes, client = _com_modules()
        self._pythoncom = pythoncom
        self._com_error = pywintypes.com_error
        _check_not_locked(self.path)
        pythoncom.CoInitialize()
        try:
            # DispatchEx: a NEW Excel process, so the user's own open window
            # (and whatever they have in it) is never touched.
            self.app = client.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.ScreenUpdating = False
            self.app.EnableEvents = False
            self.app.AskToUpdateLinks = False
            try:
                self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except self._com_error:
                pass
            # Empty passwords make a protected workbook fail fast instead of
            # waiting on a password prompt nobody can see.
            self.wb = self.app.Workbooks.Open(
                self.path, UpdateLinks=0, ReadOnly=False, Password="",
                WriteResPassword="", IgnoreReadOnlyRecommended=True,
                Notify=False, AddToMru=False)
            if self.wb.ReadOnly:
                raise WorkbookError("%s opened read-only (it is open elsewhere, "
                                    "or write-protected), so it cannot be "
                                    "changed." % os.path.basename(self.path))
        except self._com_error as exc:
            self._shutdown()
            raise WorkbookError("Excel could not open %s: %s"
                                % (os.path.basename(self.path), _com_message(exc)))
        except Exception:
            self._shutdown()
            raise
        return self

    def __exit__(self, exc_type, exc, tb):
        self._shutdown()
        if exc_type is not None and issubclass(exc_type, self._com_error):
            raise WorkbookError("Excel reported an error, and nothing was saved: %s"
                                % _com_message(exc))
        return False

    def _shutdown(self):
        if self.wb is not None:
            try:
                self.wb.Close(SaveChanges=False)
            except Exception:
                pass
            self.wb = None
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._pythoncom = None

    # -- helpers used by the write tools ------------------------------------

    def sheet_names(self):
        return [self.wb.Worksheets(i).Name
                for i in range(1, self.wb.Worksheets.Count + 1)]

    def sheet(self, name, create=False):
        """A worksheet by name (case-insensitive) or 1-based index."""
        want = str(name or "").strip()
        if not want:
            return self.wb.Worksheets(1)
        for i in range(1, self.wb.Worksheets.Count + 1):
            ws = self.wb.Worksheets(i)
            if ws.Name.lower() == want.lower():
                return ws
        if want.isdigit() and 1 <= int(want) <= self.wb.Worksheets.Count:
            return self.wb.Worksheets(int(want))
        if create:
            return self.add_sheet(want)
        raise WorkbookError("Sheet '%s' not found. Available: %s. Pass "
                            "create_sheet=true to add it."
                            % (want, ", ".join(self.sheet_names())))

    def add_sheet(self, name):
        clean = _clean_sheet_name(name)
        if clean.lower() in [n.lower() for n in self.sheet_names()]:
            raise WorkbookError("A sheet named '%s' already exists." % clean)
        last = self.wb.Worksheets(self.wb.Worksheets.Count)
        ws = self.wb.Worksheets.Add(After=last)
        ws.Name = clean
        return ws

    def table(self, name):
        """A ListObject (Table) by name, searching every sheet."""
        want = str(name or "").strip().lower()
        found = []
        for i in range(1, self.wb.Worksheets.Count + 1):
            ws = self.wb.Worksheets(i)
            for j in range(1, ws.ListObjects.Count + 1):
                lo = ws.ListObjects(j)
                found.append(lo.Name)
                if lo.Name.lower() == want:
                    return lo
        raise WorkbookError("No Table named '%s'. Tables: %s"
                            % (name, ", ".join(found) or "none"))

    def commit(self, save_as=None):
        """Recalculate and save (in place, or as a new file in the folder)."""
        try:
            self.app.CalculateFull()
        except self._com_error:
            pass
        if save_as:
            fmt = (XL_OPENXML_MACRO_WORKBOOK
                   if save_as.lower().endswith(".xlsm") else XL_OPENXML_WORKBOOK)
            self.wb.SaveAs(os.path.abspath(save_as), FileFormat=fmt)
        else:
            self.wb.Save()


def _clean_sheet_name(name):
    """Excel's sheet-name rules: <= 31 chars, none of []:*?/\\ ."""
    clean = re.sub(r"[\[\]:*?/\\]", " ", str(name)).strip().strip("'")
    clean = " ".join(clean.split())[:31].strip()
    if not clean:
        raise WorkbookError("'%s' is not a usable sheet name." % name)
    return clean


def _as_grid(value):
    """A COM Range.Value as a list of lists (a single cell comes back bare)."""
    if value is None:
        return [[None]]
    if not isinstance(value, (tuple, list)):
        return [[value]]
    return [list(row) if isinstance(row, (tuple, list)) else [row] for row in value]


def _com_value(value):
    """A JSON value -> what to hand Excel: None clears, the rest as given."""
    if isinstance(value, (dict, list)):
        raise WorkbookError("A cell value must be text, a number, true/false or "
                            "null - got %s." % type(value).__name__)
    return value


def _display(value):
    """A COM value as text for a reply (dates come back as pywintypes times)."""
    if value is None:
        return None
    if hasattr(value, "strftime"):
        try:
            if value.hour == 0 and value.minute == 0 and value.second == 0:
                return value.strftime("%Y-%m-%d")
            return value.strftime("%Y-%m-%d %H:%M:%S")
        except (AttributeError, ValueError):
            pass
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _cell_matches(cell, wanted):
    """_values_equal, plus a date cell matching its YYYY-MM-DD text."""
    if hasattr(cell, "strftime"):
        text, want = str(_display(cell)), str(wanted).strip()
        return text == want or text[:10] == want
    return _values_equal(cell, wanted)


def _write_path(folder, args):
    """The workbook to change and the optional save-as target."""
    path = resolve_workbook_path(folder, args.get("workbook"))
    save_as = _save_as_path(folder, path, args.get("save_as"),
                            bool(args.get("overwrite", False)))
    return path, save_as


def _saved_note(path, save_as):
    return "Saved to %s." % (save_as or path)


def tool_write_cells(folder, args):
    values = args.get("values")
    if not isinstance(values, list) or not values:
        raise WorkbookError("'values' must be a list of rows, e.g. "
                            "[[\"Name\", 5], [\"Other\", 6]].")
    rows = [row if isinstance(row, list) else [row] for row in values]
    width = max(len(r) for r in rows)
    if width == 0:
        raise WorkbookError("'values' has no cells.")
    if len(rows) * width > MAX_WRITE_CELLS:
        raise WorkbookError("That is %d cells; one call may write up to %d."
                            % (len(rows) * width, MAX_WRITE_CELLS))
    grid = tuple(tuple(_com_value(v) for v in (r + [None] * width)[:width])
                 for r in rows)
    start = str(args.get("cell") or "A1").strip().upper()
    col_letters, row_num = _split_cell_ref(start)
    if not col_letters or not row_num:
        raise WorkbookError("'cell' must be one A1 cell reference, e.g. 'B2'.")
    c0 = _col_letters_to_index(col_letters) + 1
    path, save_as = _write_path(folder, args)
    with ExcelSession(path) as xl:
        ws = xl.sheet(args.get("sheet"), create=bool(args.get("create_sheet", False)))
        target = ws.Range(ws.Cells(row_num, c0),
                          ws.Cells(row_num + len(rows) - 1, c0 + width - 1))
        target.Value = grid
        address = target.Address.replace("$", "")
        sheet_name = ws.Name
        xl.commit(save_as)
    return ("Wrote %d row(s) x %d column(s) to '%s'!%s in %s. %s"
            % (len(rows), width, sheet_name, address,
               os.path.basename(path), _saved_note(path, save_as)))


def _table_headers(lo):
    return [str(h) if h is not None else "" for h in _as_grid(lo.HeaderRowRange.Value)[0]]


def _header_index(headers, column, table_name):
    lookup = [h.strip().lower() for h in headers]
    key = str(column).strip().lower()
    if key not in lookup:
        raise WorkbookError("Table '%s' has no column '%s'. Columns: %s"
                            % (table_name, column, ", ".join(headers)))
    return lookup.index(key)


def tool_add_table_rows(folder, args):
    rows = args.get("rows")
    if not isinstance(rows, list) or not rows:
        raise WorkbookError("'rows' must be a list: each row an object of "
                            "column name -> value, or a list of values in "
                            "column order.")
    path, save_as = _write_path(folder, args)
    with ExcelSession(path) as xl:
        lo = xl.table(args.get("table"))
        headers = _table_headers(lo)
        # Validate EVERY row before adding any, so a bad column name in row 5
        # cannot leave rows 1-4 half-added.
        plan = []
        for n, row in enumerate(rows, 1):
            if isinstance(row, dict):
                plan.append([(_header_index(headers, k, lo.Name) + 1, _com_value(v))
                             for k, v in row.items()])
            elif isinstance(row, list):
                if len(row) > len(headers):
                    raise WorkbookError("Row %d has %d values but Table '%s' has "
                                        "%d columns." % (n, len(row), lo.Name,
                                                         len(headers)))
                plan.append([(i + 1, _com_value(v)) for i, v in enumerate(row)])
            else:
                raise WorkbookError("Row %d must be an object or a list." % n)
        for cells in plan:
            new_row = lo.ListRows.Add()
            # Only the columns given are written, so a calculated column's
            # formula (which Excel fills into the new row) is left alone.
            for col, value in cells:
                new_row.Range.Cells(1, col).Value = value
        name, ref = lo.Name, lo.Range.Address.replace("$", "")
        xl.commit(save_as)
    return ("Added %d row(s) to Table '%s' (now %s) in %s. %s"
            % (len(plan), name, ref, os.path.basename(path), _saved_note(path, save_as)))


def tool_update_table_rows(folder, args):
    match = args.get("match")
    changes = args.get("set")
    if not isinstance(match, dict) or not match:
        raise WorkbookError("'match' must be an object of column -> value "
                            "identifying the row(s), e.g. {\"ID\": \"A-17\"}.")
    if not isinstance(changes, dict) or not changes:
        raise WorkbookError("'set' must be an object of column -> new value.")
    all_matches = bool(args.get("all_matches", False))
    path, save_as = _write_path(folder, args)
    with ExcelSession(path) as xl:
        lo = xl.table(args.get("table"))
        headers = _table_headers(lo)
        match_idx = [(_header_index(headers, k, lo.Name), v) for k, v in match.items()]
        set_idx = [(_header_index(headers, k, lo.Name), _com_value(v))
                   for k, v in changes.items()]
        body = lo.DataBodyRange
        grid = _as_grid(body.Value) if body is not None else []
        hits = [r for r, row in enumerate(grid)
                if all(_cell_matches(row[i], v) for i, v in match_idx)]
        if not hits:
            raise WorkbookError("No row of Table '%s' matches %s; nothing changed."
                                % (lo.Name, json.dumps(match, ensure_ascii=False)))
        if len(hits) > 1 and not all_matches:
            raise WorkbookError("%d rows of Table '%s' match %s; nothing changed. "
                                "Narrow 'match', or pass all_matches=true to "
                                "change all of them."
                                % (len(hits), lo.Name,
                                   json.dumps(match, ensure_ascii=False)))
        first_row = body.Row
        for r in hits:
            for i, value in set_idx:
                body.Cells(r + 1, i + 1).Value = value
        name = lo.Name
        sheet_rows = ", ".join(str(first_row + r) for r in hits[:20])
        xl.commit(save_as)
    return ("Updated %d row(s) of Table '%s' (sheet row(s) %s%s) in %s. %s"
            % (len(hits), name, sheet_rows, ", ..." if len(hits) > 20 else "",
               os.path.basename(path), _saved_note(path, save_as)))


def _field_list(value, label):
    if value is None:
        return []
    if isinstance(value, str):
        value = [v for v in value.split(",")]
    if not isinstance(value, list):
        raise WorkbookError("'%s' must be a list of column names." % label)
    return [str(v).strip() for v in value if str(v).strip()]


def tool_create_pivot_table(folder, args):
    rows = _field_list(args.get("rows"), "rows")
    cols = _field_list(args.get("columns"), "columns")
    filters = _field_list(args.get("filters"), "filters")
    values = args.get("values")
    if isinstance(values, (str, dict)):
        values = [values]
    if not isinstance(values, list) or not values:
        raise WorkbookError("'values' is required: the column(s) to summarise, "
                            "e.g. [{\"field\": \"Amount\", \"function\": \"sum\"}].")
    specs = []
    for v in values:
        if isinstance(v, str):
            v = {"field": v}
        if not isinstance(v, dict) or not str(v.get("field") or "").strip():
            raise WorkbookError("Each entry of 'values' needs a 'field'.")
        func = str(v.get("function") or "sum").strip().lower()
        if func not in PIVOT_FUNCTIONS:
            raise WorkbookError("Unknown summary function '%s'. Use one of: %s"
                                % (func, ", ".join(sorted(set(PIVOT_FUNCTIONS)))))
        specs.append((str(v["field"]).strip(), func,
                      str(v.get("caption") or "").strip(),
                      str(v.get("number_format") or "").strip()))
    if not rows and not cols:
        raise WorkbookError("Give at least one 'rows' or 'columns' field to "
                            "group by.")

    source_table = str(args.get("source_table") or "").strip()
    source_range = str(args.get("source_range") or "").strip()
    if not source_table and not source_range:
        raise WorkbookError("Name the data: 'source_table' (a Table name) or "
                            "'source_sheet' plus 'source_range' (e.g. 'A1:F500', "
                            "header row first).")

    path, save_as = _write_path(folder, args)
    with ExcelSession(path) as xl:
        if source_table:
            lo = xl.table(source_table)
            # The Table's NAME, not its current cells, so the pivot keeps up
            # when rows are added to the Table and it is refreshed.
            source_data = lo.Name
            source_label = "Table '%s'" % lo.Name
            headers = _table_headers(lo)
        else:
            src_ws = xl.sheet(args.get("source_sheet"))
            source_data = src_ws.Range(source_range)
            source_label = "'%s'!%s" % (src_ws.Name, source_range.upper())
            headers = [str(h) if h is not None else ""
                       for h in _as_grid(source_data.Rows(1).Value)[0]]
        # Check every field against the source headers before building
        # anything: Excel's own error for a bad field name says nothing useful.
        lookup = {h.strip().lower(): h for h in headers if h.strip()}
        for field in rows + cols + filters + [s[0] for s in specs]:
            if field.lower() not in lookup:
                raise WorkbookError("'%s' is not a column of %s. Columns: %s"
                                    % (field, source_label, ", ".join(headers)))

        dest_name = str(args.get("destination_sheet") or "").strip()
        if dest_name and dest_name.lower() in [n.lower() for n in xl.sheet_names()]:
            dest_ws = xl.sheet(dest_name)
        else:
            base = dest_name or ("Pivot - " + (source_table or src_ws.Name))
            candidate, n = _clean_sheet_name(base), 2
            existing = [x.lower() for x in xl.sheet_names()]
            while candidate.lower() in existing:
                candidate = _clean_sheet_name("%s (%d)" % (base[:26], n))
                n += 1
            dest_ws = xl.add_sheet(candidate)
        dest_cell = str(args.get("destination_cell") or "A3").strip().upper()
        name = str(args.get("name") or "").strip() or "Pivot%s" % re.sub(
            r"[^A-Za-z0-9]", "", source_table or src_ws.Name)[:40]

        cache = xl.wb.PivotCaches().Create(SourceType=XL_DATABASE,
                                           SourceData=source_data)
        pt = cache.CreatePivotTable(TableDestination=dest_ws.Range(dest_cell),
                                    TableName=name)
        for field in rows:
            pt.PivotFields(lookup[field.lower()]).Orientation = XL_ROW_FIELD
        for field in cols:
            pt.PivotFields(lookup[field.lower()]).Orientation = XL_COLUMN_FIELD
        for field in filters:
            pt.PivotFields(lookup[field.lower()]).Orientation = XL_PAGE_FIELD
        for field, func, caption, number_format in specs:
            real = lookup[field.lower()]
            label = caption or "%s of %s" % (
                {"countnums": "Count", "count_numbers": "Count", "avg": "Average",
                 "mean": "Average"}.get(func, func.capitalize()), real)
            data_field = pt.AddDataField(pt.PivotFields(real), label,
                                         PIVOT_FUNCTIONS[func])
            if number_format:
                data_field.NumberFormat = number_format
        pt_name, sheet_name = pt.Name, dest_ws.Name
        address = pt.TableRange2.Address.replace("$", "")
        preview = _as_grid(pt.TableRange1.Value)
        xl.commit(save_as)

    lines = ["Created pivot table '%s' on sheet '%s' (%s) from %s in %s. %s"
             % (pt_name, sheet_name, address, source_label,
                os.path.basename(save_as or path), _saved_note(path, save_as)),
             "", "It shows:"]
    shown = [[_display(v) for v in row] for row in preview[:MAX_PIVOT_PREVIEW_ROWS]]
    if shown:
        lines.append(_md_table(shown[0], shown[1:]))
    if len(preview) > MAX_PIVOT_PREVIEW_ROWS:
        lines.append("[%d more row(s) - read it with excel_read_pivot_table.]"
                     % (len(preview) - MAX_PIVOT_PREVIEW_ROWS))
    return "\n".join(lines)


def _check_excel_automation():
    """--check helper: can this endpoint start Excel? Returns a status line."""
    try:
        pythoncom, _pywintypes, client = _com_modules()
    except WorkbookError as exc:
        return "NOT available - %s" % exc
    pythoncom.CoInitialize()
    app = None
    try:
        app = client.DispatchEx("Excel.Application")
        return "OK - Excel %s" % app.Version
    except Exception as exc:
        return "NOT available - Excel could not be started (%s)" % _com_message(exc)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ===========================================================================
# MCP tool registry (name -> (handler, description, input schema))
# ===========================================================================

TOOLS = {
    "excel_list_workbooks": {
        "handler": tool_list_workbooks,
        "description": "List the Excel workbooks (.xlsx/.xlsm) available to read "
                       "in the configured folder, with file sizes.",
        "schema": {
            "type": "object",
            "properties": {},
        },
    },
    "excel_list_sheets": {
        "handler": tool_list_sheets,
        "description": "List the sheet names and declared used-ranges in a workbook.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string",
                             "description": "Workbook name (loose match; "
                                            "extension optional)."},
            },
            "required": ["workbook"],
        },
    },
    "excel_get_headers": {
        "handler": tool_get_headers,
        "description": "Return the header row of a sheet, one entry per column, "
                       "with its column letter. Use this before reading data so "
                       "you know which columns exist.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index. "
                                         "Defaults to the first sheet."},
                "header_row": {"type": "integer",
                               "description": "Row number of the header "
                                              "(default 1)."},
            },
            "required": ["workbook"],
        },
    },
    "excel_read_range": {
        "handler": tool_read_range,
        "description": "Read cell values from a sheet. Optionally pass an A1 "
                       "range like 'A1:D50'. Output is capped for safety; use a "
                       "range to page through large sheets.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index. "
                                         "Defaults to the first sheet."},
                "range": {"type": "string",
                          "description": "Optional A1 range, e.g. 'A1:D50'. "
                                         "Omit to read from the top."},
            },
            "required": ["workbook"],
        },
    },
    "excel_search": {
        "handler": tool_search,
        "description": "Find cells whose value contains a query string. Searches "
                       "all sheets unless 'sheet' is given. Returns cell "
                       "references and values.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "query": {"type": "string",
                          "description": "Text to search for."},
                "sheet": {"type": "string",
                          "description": "Optional: restrict to one sheet."},
                "case_sensitive": {"type": "boolean",
                                   "description": "Default false."},
            },
            "required": ["workbook", "query"],
        },
    },
    "excel_column_stats": {
        "handler": tool_column_stats,
        "description": "Summarise one column: count, sum, mean, median, min, "
                       "max for numeric data, or top values for categorical "
                       "data. Column may be a letter (C), a number (3), or a "
                       "header name.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "column": {"type": "string",
                           "description": "Column letter, 1-based number, or "
                                          "header name."},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index. "
                                         "Defaults to the first sheet."},
                "header_row": {"type": "integer",
                               "description": "Header row number (default 1)."},
                "data_start_row": {"type": "integer",
                                   "description": "First data row "
                                                  "(default header_row + 1)."},
            },
            "required": ["workbook", "column"],
        },
    },
    "excel_list_tables": {
        "handler": tool_list_tables,
        "description": "List the Excel Tables (Insert > Table, also called "
                       "ListObjects) in a workbook: name, sheet, range, data "
                       "row count and column names. Use a Table's name with "
                       "excel_read_table.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
            },
            "required": ["workbook"],
        },
    },
    "excel_read_table": {
        "handler": tool_read_table,
        "description": "Read an Excel Table by its name (e.g. 'tblBudget'), "
                       "wherever it sits in the workbook, as a Markdown table "
                       "with the sheet row number of each row. Optionally "
                       "pick columns, filter to rows where one column equals "
                       "a value, and page with 'offset'.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "table": {"type": "string",
                          "description": "The Table's name (case-insensitive)."},
                "columns": {"type": "array", "items": {"type": "string"},
                            "description": "Optional: only these columns, by header name."},
                "filter_column": {"type": "string",
                                  "description": "Optional: keep rows where this column..."},
                "filter_value": {"type": "string",
                                 "description": "...equals this value (case-insensitive; "
                                                "numbers compare as numbers)."},
                "offset": {"type": "integer",
                           "description": "Skip this many data rows (paging; default 0)."},
            },
            "required": ["workbook", "table"],
        },
    },
    "excel_list_pivot_tables": {
        "handler": tool_list_pivot_tables,
        "description": "List the pivot tables in a workbook with their layout: "
                       "sheet and cells, source data, row/column/filter fields "
                       "and value fields (e.g. 'Sum of Amount').",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
            },
            "required": ["workbook"],
        },
    },
    "excel_read_pivot_table": {
        "handler": tool_read_pivot_table,
        "description": "Read one pivot table: its layout plus the figures it "
                       "shows, as Excel last calculated and saved them. "
                       "'pivot' may be omitted when the workbook has only one.",
        "schema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "pivot": {"type": "string",
                          "description": "Pivot table name, e.g. 'PivotTable1'."},
            },
            "required": ["workbook"],
        },
    },
}


_SAVE_AS_PROPS = {
    "save_as": {"type": "string",
                "description": "Optional: save the result as a NEW workbook "
                               "of this name in the same folder, leaving the "
                               "original untouched."},
    "overwrite": {"type": "boolean",
                  "description": "Allow 'save_as' to replace an existing file "
                                 "(default false)."},
}


def _with_save_as(props):
    merged = dict(props)
    merged.update(_SAVE_AS_PROPS)
    return merged


# Each call starts a private, invisible Excel, so these need Windows, desktop
# Excel and pywin32; without them they say what to install.
WRITE_TOOLS = {
    "excel_write_cells": {
        "handler": tool_write_cells,
        "description": "Write values into a sheet (tab), starting at a cell: "
                       "'values' is a list of rows, e.g. [[\"Region\", "
                       "\"Total\"], [\"North\", 1200]]. A string starting "
                       "with '=' is a formula; null clears a cell; dates as "
                       "YYYY-MM-DD. Excel recalculates and saves the workbook "
                       "(or a copy, with 'save_as'). The workbook must not be "
                       "open in Excel.",
        "schema": {
            "type": "object",
            "properties": _with_save_as({
                "workbook": {"type": "string"},
                "sheet": {"type": "string",
                          "description": "Sheet name or 1-based index "
                                         "(default: the first sheet)."},
                "cell": {"type": "string",
                         "description": "Top-left cell to write from, e.g. 'B2' "
                                        "(default A1)."},
                "values": {"type": "array", "items": {"type": "array"},
                           "description": "Rows of cell values."},
                "create_sheet": {"type": "boolean",
                                 "description": "Add the sheet if it does not "
                                                "exist (default false)."},
            }),
            "required": ["workbook", "values"],
        },
    },
    "excel_add_table_rows": {
        "handler": tool_add_table_rows,
        "description": "Append rows to an Excel Table by name. Each row is an "
                       "object of column name -> value (unnamed columns keep "
                       "their calculated formulas), or a list of values in "
                       "column order. The Table grows to take them.",
        "schema": {
            "type": "object",
            "properties": _with_save_as({
                "workbook": {"type": "string"},
                "table": {"type": "string", "description": "The Table's name."},
                "rows": {"type": "array",
                         "description": "Rows to add: objects or lists."},
            }),
            "required": ["workbook", "table", "rows"],
        },
    },
    "excel_update_table_rows": {
        "handler": tool_update_table_rows,
        "description": "Change values in the row(s) of an Excel Table that "
                       "match: 'match' is column -> value identifying the "
                       "row (e.g. {\"ID\": \"A-17\"}), 'set' is column -> "
                       "new value. Refuses when several rows match unless "
                       "all_matches is true.",
        "schema": {
            "type": "object",
            "properties": _with_save_as({
                "workbook": {"type": "string"},
                "table": {"type": "string", "description": "The Table's name."},
                "match": {"type": "object",
                          "description": "Column -> value the row must have "
                                         "(case-insensitive)."},
                "set": {"type": "object",
                        "description": "Column -> new value."},
                "all_matches": {"type": "boolean",
                                "description": "Change every matching row "
                                               "(default false)."},
            }),
            "required": ["workbook", "table", "match", "set"],
        },
    },
    "excel_create_pivot_table": {
        "handler": tool_create_pivot_table,
        "description": "Create a real Excel PivotTable from a Table "
                       "('source_table') or a range with a header row "
                       "('source_sheet' + 'source_range'). Group by 'rows' "
                       "and/or 'columns', summarise 'values' (each "
                       "{field, function: sum|count|average|max|min|...}), "
                       "optionally with 'filters'. It goes on a new sheet "
                       "unless 'destination_sheet' names an existing one. "
                       "Returns the figures it shows.",
        "schema": {
            "type": "object",
            "properties": _with_save_as({
                "workbook": {"type": "string"},
                "source_table": {"type": "string",
                                 "description": "Name of the Table to summarise (preferred)."},
                "source_sheet": {"type": "string",
                                 "description": "Sheet of the source range (with 'source_range')."},
                "source_range": {"type": "string",
                                 "description": "A1 range including the header row, e.g. 'A1:F500'."},
                "rows": {"type": "array", "items": {"type": "string"},
                         "description": "Column names to group down the side."},
                "columns": {"type": "array", "items": {"type": "string"},
                            "description": "Column names to group across the top."},
                "values": {"type": "array",
                           "items": {"type": "object"},
                           "description": "What to summarise: [{\"field\": "
                                          "\"Amount\", \"function\": \"sum\", "
                                          "\"caption\": optional, "
                                          "\"number_format\": optional e.g. "
                                          "\"#,##0\"}]."},
                "filters": {"type": "array", "items": {"type": "string"},
                            "description": "Column names to add as report filters."},
                "destination_sheet": {"type": "string",
                                      "description": "Sheet to put it on (default: a "
                                                     "new sheet 'Pivot - <source>')."},
                "destination_cell": {"type": "string",
                                     "description": "Top-left cell (default A3)."},
                "name": {"type": "string",
                         "description": "PivotTable name (default derived from the source)."},
            }),
            "required": ["workbook", "values"],
        },
    },
}


def active_tools():
    """Every tool on offer: the read tools, then the write tools."""
    tools = dict(TOOLS)
    tools.update(WRITE_TOOLS)
    return tools


# ===========================================================================
# JSON-RPC / MCP stdio server
# ===========================================================================

def _make_result(req_id, result):
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def _make_error(req_id, code, message):
    return {"jsonrpc": "2.0", "id": req_id,
            "error": {"code": code, "message": message}}


def _tool_text_result(text):
    return {"content": [{"type": "text", "text": text}], "isError": False}


def _tool_error_result(text):
    return {"content": [{"type": "text", "text": text}], "isError": True}


def handle_request(msg, folder):
    """
    Handle one parsed JSON-RPC message. Returns a response dict, or None for
    notifications (which must not be answered).
    """
    method = msg.get("method")
    req_id = msg.get("id")
    is_notification = "id" not in msg

    if method == "initialize":
        result = {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {"tools": {}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
        }
        return _make_result(req_id, result)

    if method in ("notifications/initialized", "initialized", "notifications/cancelled"):
        return None  # notification: no response

    if method == "ping":
        return _make_result(req_id, {})

    if method == "tools/list":
        tool_list = []
        for name, spec in active_tools().items():
            tool_list.append({
                "name": name,
                "description": spec["description"],
                "inputSchema": spec["schema"],
            })
        return _make_result(req_id, {"tools": tool_list})

    if method == "tools/call":
        params = msg.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        spec = active_tools().get(name)
        if spec is None:
            return _make_result(
                req_id,
                _tool_error_result("Unknown tool: %s" % name))
        try:
            text = spec["handler"](folder, arguments)
            log("tool '%s' args=%s -> %d chars"
                % (name, json.dumps(arguments, ensure_ascii=False), len(text)))
            return _make_result(req_id, _tool_text_result(text))
        except WorkbookError as exc:
            log("tool '%s' workbook error: %s" % (name, exc))
            return _make_result(req_id, _tool_error_result(str(exc)))
        except Exception as exc:   # never crash the server on one bad call
            log("tool '%s' unexpected error: %r" % (name, exc))
            return _make_result(
                req_id,
                _tool_error_result("Internal error running %s: %s"
                                   % (name, exc)))

    # Unknown method.
    if is_notification:
        return None
    return _make_error(req_id, -32601, "Method not found: %s" % method)


def serve(folder):
    """Main stdio loop: read newline-delimited JSON-RPC, write responses."""
    log("starting; interpreter=%s" % sys.executable)
    log("workbook folder=%s" % folder)
    if not os.path.isdir(folder):
        log("WARNING: workbook folder does not exist yet: %s" % folder)

    for raw in sys.stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError as exc:
            log("could not parse line as JSON: %s" % exc)
            # Cannot know the id; emit a parse error with null id.
            sys.stdout.write(json.dumps(
                _make_error(None, -32700, "Parse error")) + "\n")
            sys.stdout.flush()
            continue

        # A batch (list) is technically valid JSON-RPC; handle defensively.
        messages = msg if isinstance(msg, list) else [msg]
        for m in messages:
            if not isinstance(m, dict):
                continue
            try:
                response = handle_request(m, folder)
            except Exception as exc:
                log("fatal handler error: %r" % exc)
                response = _make_error(m.get("id"), -32603,
                                       "Internal error: %s" % exc)
            if response is not None:
                sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
                sys.stdout.flush()

    log("stdin closed; exiting")


# ===========================================================================
# CLI
# ===========================================================================

def env(name):
    """Read an environment variable, treating blank as unset.

    A blank value is what an MCP client substitutes for a setting the user left
    empty, and an unexpanded "${...}" placeholder is what it leaves behind when
    the variable it refers to does not exist. Both mean "not configured", so
    the default still applies.
    """
    value = (os.environ.get(name) or "").strip()
    if not value or (value.startswith("${") and value.endswith("}")):
        return None
    return value


def resolve_docs_dir():
    """The workbook folder, from the environment.

    Precedence: EXCEL_DOCS_DIR (a full path of its own), then
    EVA_DOCUMENTS_DIR with this server's "excel" sub-folder appended, then the
    same sub-folder of the EVA_DOCUMENTS_DIR default in the CONFIG block.
    Returns (path, user_configured) - user_configured is False only when
    nothing was set at all, which decides how a missing folder is reported.
    """
    own = env("EXCEL_DOCS_DIR")
    if own:
        return own, True
    root = env("EVA_DOCUMENTS_DIR")
    if root:
        return os.path.join(root, SUBFOLDER), True
    return os.path.join(EVA_DOCUMENTS_DIR, SUBFOLDER), False


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Excel (.xlsx) MCP server: reads workbooks directly, and "
                    "changes them through the installed Excel. "
                    "Configuration is environment variables only: "
                    "EVA_DOCUMENTS_DIR (this server works in its 'excel' "
                    "sub-folder), or EXCEL_DOCS_DIR to override that one "
                    "folder. See the CONFIGURATION section of this file's "
                    "docstring.")
    parser.add_argument("--check", action="store_true",
                        help="Print environment/config diagnostics and exit.")
    parser.add_argument("--list", action="store_true",
                        help="List readable workbooks in the folder and exit.")
    parser.add_argument("--version", action="version",
                        version="{0} {1}".format(SERVER_NAME, __version__))
    args = parser.parse_args(argv)

    global DOCS_DIR
    DOCS_DIR, folder_chosen = resolve_docs_dir()
    folder = DOCS_DIR

    if args.check:
        print("excel_mcp environment check")
        print("  python executable : %s" % sys.executable)
        print("  python version    : %s" % sys.version.split()[0])
        print("  workbook folder   : %s" % folder)
        print("  came from         : %s"
              % ("EXCEL_DOCS_DIR / EVA_DOCUMENTS_DIR" if folder_chosen
                 else "built-in default (no EVA_DOCUMENTS_DIR set)"))
        print("  folder exists     : %s" % os.path.isdir(folder))
        if os.path.isdir(folder):
            try:
                files = list_workbook_files(folder)
                print("  workbooks found   : %d" % len(files))
                for f in files:
                    print("      - %s" % f)
            except WorkbookError as exc:
                print("  error listing     : %s" % exc)
        # Starts (and quits) a private Excel, proving the write tools can run
        # here before the model finds out the hard way.
        print("  Excel automation  : %s (needed by the write tools only)"
              % _check_excel_automation())
        tools = active_tools()
        print("  tools registered  : %d (%s)"
              % (len(tools), ", ".join(tools.keys())))
        return 0

    # The workbook folder is REQUIRED: the server only reads inside it and
    # must not start unconfined.
    if not os.path.isdir(folder):
        log("FATAL: the workbook folder does not exist or is not a "
            "directory: %s" % folder)
        if folder_chosen:
            log("       That path came from EVA_DOCUMENTS_DIR or "
                "EXCEL_DOCS_DIR. Create the folder, or fix the variable.")
        else:
            log("       That is the built-in default. Create the folder, set "
                "EVA_DOCUMENTS_DIR to your own document root, or copy the "
                "repo's eva\\ folder to H:\\Eva to lay out the whole tree.")
        return 2

    if args.list:
        try:
            print(tool_list_workbooks(folder, {}))
        except WorkbookError as exc:
            print("Error: %s" % exc, file=sys.stderr)
            return 1
        return 0

    serve(folder)
    return 0


if __name__ == "__main__":
    sys.exit(main())
