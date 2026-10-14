#!/usr/bin/env python3
"""
outlook.py (v11.3.0)
======================

A single-file MCP (Model Context Protocol) server giving an LLM near-read-only
access to a locally installed *classic* Microsoft Outlook client (mail +
calendar) on Windows, via COM automation, plus enough of Outlook's scheduling
machinery to get a meeting into the diary: fuzzy directory search, free/busy,
and an UNSENT meeting draft - and an UNSENT email draft, opened in Outlook's
own compose window.

NOTHING HERE CAN SEND. Not an email, not a reply, not a meeting invitation.
Nor can it accept, move or delete anything in the mailbox. There are exactly
TWO writes to the mailbox, both drafts the user sends themselves:
outlook_draft_meeting saves a meeting into the Calendar with its invitations
unsent (the state Outlook labels "invitations have not been sent for this
meeting"), and outlook_draft_email (v11.1.0) opens an email, reply or forward
in Outlook's compose window with the agreed text in it and saves it to Drafts.
Set OUTLOOK_ALLOW_DRAFTS=false and both are gone, at which point the mailbox is
strictly read-only again. The other two things it
writes are local files and only on request: an email kept as Markdown for the
knowledge base, a printable PDF day planner, and a printable PDF of the
team's open tasks.

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

TOOLS EXPOSED (read-only on the mailbox, apart from the two draft tools)
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
                               Recurring series are gathered two ways and merged
                               (see RECURRING SERIES below), so a series does not
                               go missing when one method cannot see it.
- outlook_print_calendar     : a PRINTABLE PDF day planner for one day, with an
                               optional team view of up to four colleagues on
                               the back (v11.2.0) - see below
- outlook_get_task_team      : the team task sheet's columns and Confluence
                               page (v11.3.0) - see below
- outlook_list_task_emails   : emails you sent, flagged in your Inbox with a due
                               date - the emailed tasks (v11.3.0)
- outlook_print_team_tasks   : a PRINTABLE PDF of the team's open tasks, eight
                               columns over two sides of A4 (v11.3.0)
- outlook_list_sent_emails   : messages you SENT, in a date range (e.g. "what did I do last week")
- outlook_search_recent      : all mail across Inbox/Archive/Sent in a date range (configurable)
- outlook_list_folders       : list every mail folder across all stores (to configure the above)

TOOLS FOR SCHEDULING A MEETING  (new in v8.0.0 - see MEETING SCHEDULING below)
------------------------------------------------------------------------------
- outlook_find_people           : fuzzy-find a person or room in the directory
- outlook_suggest_meeting_times : times everyone (and the room) is free
- outlook_draft_meeting         : save an UNSENT meeting into the calendar

TOOL FOR DRAFTING AN EMAIL  (new in v11.1.0 - see EMAIL DRAFTS below)
---------------------------------------------------------------------
- outlook_draft_email           : open an UNSENT email / reply / forward in
                                  Outlook's compose window, saved to Drafts

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
(%EVA_KNOWLEDGE_DIR%\email, H:\Eva\knowledge\email on a stock install), which
sits inside the knowledge-base plugin's corpus so saved mail is actually
indexed. Point OUTLOOK_KB_DIR at a full path of its own to override just this
server.

To make saving impossible, set OUTLOOK_KB_DIR=off: a save_to_kb request is then
refused and the server touches no disk at all. To go the other way and save
EVERY email read, set OUTLOOK_KB_AUTOSAVE=true.
Worth a deliberate decision either way: saving turns correspondence into plain
text files that are then embedded and quotable in answers. See
eva\knowledge\email\README.md.

MEETING SCHEDULING
------------------
Three tools that together answer "find a suitable time with Josh Smith in
Room R5-3-84, then save it so I can send it". They are meant to be used in
that order, and the middle one will not take a name - only an address.

1. outlook_find_people - fuzzy directory search
   ...........................................
   An enterprise directory is unusable without one: the user says "Josh
   Smith", the GAL says "Smith, Joshua P", and the room they call R5-3-84 is
   filed as "Room R5-3-84 (12 seats)". Matching is token-based and
   order-independent, with difflib behind it for typos, so all three land.

   Searched, cheapest first:
     a. Outlook's own ambiguous-name resolution. An unmistakable name is
        answered instantly and NOTHING is scanned. Pass deep=true to scan
        anyway when the user says it found the wrong person.
     b. The Contacts folders.
     c. The address lists - Global Address List, All Rooms, and any others.
        Only the display NAME is read for each entry, because that is one COM
        round trip per entry; the expensive reads (SMTP address, job title,
        room-or-person) happen only for the few that score well enough to be
        shown. The scan stops at OUTLOOK_GAL_SCAN_CAP entries and says so.

   A room is told from a person by PR_DISPLAY_TYPE_EX (DT_ROOM / DT_EQUIPMENT)
   - AddressEntry.DisplayType has no room in its enumeration at all. Where a
   cached offline address book withholds that property, an entry found in an
   "All Rooms"-style list is a room, and failing that the name is read for it.

   Each match reports an 'invite_as' value. That is what the other two tools
   take.

2. outlook_suggest_meeting_times - when is everyone actually free
   ..............................................................
   Reads Recipient.FreeBusy for every attendee and room - the same published
   free/busy Outlook's own Scheduling Assistant uses - and walks the window in
   slot_minutes steps looking for a gap that fits. The user's own calendar is
   included by default.

   Three judgements worth knowing about:

   - TENTATIVE blocks a ROOM but not a PERSON. A tentatively-held room is
     already somebody's claim on the space; a tentative person is worth
     asking, and an enterprise calendar is full of them. A person's tentative
     time is offered with a note saying so.
   - AN OPTIONAL ATTENDEE NEVER BLOCKS a slot, but a slot they cannot make is
     labelled with their name, and slots they CAN make rank higher.
   - UNPUBLISHED FREE/BUSY IS NOT FREE. An external attendee, a mailbox that
     does not publish, or an unreachable Availability service all come back
     empty, and the reply names those people rather than quietly scheduling
     over them.

   Suggestions are spread out before they are printed: at the default
   quarter-hour granularity the five best slots are otherwise the same
   afternoon shifted 15 minutes at a time, which is five options on paper and
   one in practice. Overlapping slots and a third slot on the same day are
   skipped in favour of a different day.

   Exchange publishes 30 days of free/busy and no more, so a longer window is
   clamped to it and the reply says where it stopped.

3. outlook_draft_meeting - save it, do not send it
   ................................................
   Creates an AppointmentItem with MeetingStatus = olMeeting, attaches the
   attendees (required / optional / resource), Saves it, and opens it on the
   user's screen. Save() only - Send() is never called anywhere in this file
   and must not be added.

   If ANY attendee will not resolve, nothing at all is saved: a half-addressed
   draft looks finished and gets sent without the person it was for. The
   content blacklist applies in this direction too - a subject or body
   carrying a protective marking is refused rather than written into Outlook.

   THE SUBJECT AND THE AGENDA. Both are ordinary arguments, and the usual
   shape of the request is two steps: write an agenda, then "find a time with
   Josh to discuss it". The agenda belongs in 'body' - the caller should carry
   it across rather than making the user type it again, having shown them what
   is being attached, because a meeting body is read by everyone invited.

   AppointmentItem.Body is PLAIN TEXT and Outlook renders no Markdown, so an
   agenda drafted in conversation would otherwise reach the invitees with its
   asterisks still on it. plain_text_body() strips **bold**, leading
   #-headings, and normalises * and + bullets to -. It does no more than that
   on purpose: a lone asterisk is as likely to be a footnote mark as emphasis,
   and "C#" is not a heading.

EMAIL DRAFTS
------------
The wording of an email is agreed in the CONVERSATION; outlook_draft_email is
the hand-off, called only when the user says to put it in Outlook ("ok, send
that to Outlook"). Its tool description tells the model so, because a compose
window popping open on every round of a draft is noise.

- mode 'new', or 'reply' / 'reply_all' / 'forward' with the EntryID of the
  message being answered. Outlook itself builds the reply, so the thread, the
  quoted original, the RE:/FW: subject and the reply signature are Outlook's.
- THE SIGNATURE AND THE FONT ARE OUTLOOK'S. The compose window is created
  first (that is the moment Outlook adds the signature for the account and the
  kind of message), then the text is written through the window's Word editor
  at the top of the message, where the cursor sits. Text put there takes the
  formatting of that line: the user's own new-mail font, or their reply font.
  Writing HTMLBody instead makes a reply pick up the ORIGINAL message's style
  (the familiar "why is my reply in Times New Roman"), so it is only a
  fallback for when the editor cannot be reached, and the reply says so.
- Markdown is converted, not pasted: **bold**, *italic*, "- " bullets, "1. "
  numbered lists, [text](url) and bare URLs as links, "# heading" as a bold
  line. A plain-text email (the user's setting, or a reply to a plain-text
  message) drops the formatting and writes bullets as "- ".
- The window is left OPEN for the user to edit, and the draft is also Saved
  to Drafts so closing Outlook cannot lose it. Save() only - Send() is never
  called anywhere in this file and must not be added.
- Any To/Cc/Bcc that will not resolve means NOTHING is put into Outlook, as
  for a meeting. The content blacklist applies on the way in (subject, body,
  recipients), and a message the blacklist withholds cannot be replied to or
  forwarded from here.

RECURRING SERIES  (why a meeting can be in Outlook but not in the results)
--------------------------------------------------------------------------
Recurring appointments are not stored as individual meetings: the folder holds
one master per series, and the occurrences have to be worked out. This server
does it TWICE and merges the results, because either method alone loses
meetings.

1. Outlook's own expansion. Sorting the Items collection by [Start] and setting
   IncludeRecurrences makes Outlook hand back individual occurrences. Outlook
   holds each series' timezone and daylight-saving rules, so it places an
   occurrence correctly. This is the primary source.

2. Probing each master's RecurrencePattern with GetOccurrence, day by day. This
   only finds an occurrence whose exact start time was guessed in advance, so it
   is a backstop rather than the primary.

Method 2 was on its own until v6.3.0, and it guessed that time from the MASTER's
Start - which is the FIRST occurrence of the series, often years old. Any series
whose occurrences no longer start at that same wall-clock time was invisible:
nothing errored, the meetings simply were not there. That covers a series set up
under a daylight-saving offset that has since drifted, and one stored in an
overseas organiser's timezone. It is why a long-standing weekly meeting could be
missing while a new one showed up fine. Method 2 now probes the pattern's own
StartTime as well as the master's, and method 1 catches the rest.

Neither method uses a formatted date string, so this does not reintroduce the
Restrict() fault where regional settings silently empty the results.

To see what the two methods make of your own calendar:

    python outlook.py --check

prints, per recurring series, what each one found for today, and names any
series one can see and the other cannot.

Printable day planner (outlook_print_calendar)
----------------------------------------------
Asked for "today's printable calendar" or "tomorrow's planner", this writes an
A4 LANDSCAPE PDF laid out as a bifold spread: the day itself on an hour-by-hour
timeline down the left half, and the following days summarised on the right.
Print it single-sided and fold it in half.

The page prints on the TOP FOUR FIFTHS of the sheet: the bottom fifth comes out
blank, with a hairline to fold it back on, so the folded sheet sits in a diary.
It is still a full A4 landscape page, so it prints on ordinary paper with no
scaling - only the ink stops early. OUTLOOK_CALENDAR_PAGE_FILL changes how much
of the sheet is used, as a fraction ("0.8"), a percentage ("80" or "80%"), or
"off" for the whole sheet with nothing left to fold.

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
  There is no colour legend across the top: a printed page is looked at, not
  decoded, and the header reads better given back to the date.
- Whole categories can be left OFF the page: set
  OUTLOOK_CALENDAR_HIDE_CATEGORIES="Birthdays,Holidays" and any appointment
  carrying one of them is not printed at all, in either panel, and does not
  stretch the timeline or count as a busy day. Names match Outlook's category
  names case-insensitively. Only the printed planner honours this;
  outlook_get_calendar still shows everything the blacklist allows.
- Type on a block goes white or dark automatically, by how bright the fill is,
  so a yellow category is readable rather than white-on-yellow. Every subject
  on the page is set at ONE size (PLANNER_TITLE_PT, 7pt) whatever the block it
  sits on: the type used to follow the size of the block, which made an
  hour-long meeting shout and a fifteen-minute one whisper for no reason other
  than its duration. It is small on purpose - a planner is read at desk
  distance, and the subject of a meeting fitting on its block matters more than
  large type.
- ATTENDEES ARE NEVER PRINTED. Who is in a meeting is in the invitation, it
  crowds out the subject and location the sheet is actually for, and it puts
  other people's names on something left on a desk. They are still read, since
  an outside-domain attendee is what colours an uncategorised block.
- Tentative or free-marked time is drawn hollow, the way a diary pencils
  something in.
- The right-hand panel skips days with nothing in the diary by default, so a
  Friday planner shows the week ahead rather than two blank weekend panels.
  Pass skip_empty_days=false for strictly consecutive days.

TEAM VIEW ON THE BACK (v11.2.0). Set OUTLOOK_CALENDAR_TEAM to up to four
people - your manager and colleagues - and the PDF gets a SECOND page: their
calendars for the same day as side-by-side columns, each with the person's
name at the top, in the same design as the front (same hour grid, blocks,
chips and footer, over one shared timeline so a time reads straight across).
It has no page header (v11.2.1): the date is on the front, and that room goes
to the columns, with the date repeated small in the footer. It is meant for
the back of the sheet:

    PRINT DOUBLE-SIDED, FLIP ON SHORT EDGE.

The PDF carries that as a print-dialog preset (the PDF 1.7 /Duplex viewer
preference, DuplexFlipShortEdge). Adobe Reader and Acrobat honour it; other
viewers may not, so the tool's reply says to check the setting. Flipping on
the short edge keeps the top of the back at the top of the front, so the
front's blank bottom strip, once folded up behind the page, lies across the
back from (2 x page fill - 1) down to the fold - 60% to 80% of the sheet at
the default 0.8. The team view is therefore laid out in the top 60% of the
back, where the folded strip cannot cover it (the whole back when
OUTLOOK_CALENDAR_PAGE_FILL=off; never less than 45%, below which part of it is
under the strip until it is unfolded).

    OUTLOOK_CALENDAR_TEAM="Jane Smith=jane.smith@agency.gov.au; Bob=bob.jones@agency.gov.au"

Entries are separated by semicolons (a directory name like "Smith, Jane" has
a comma in it); commas also work between plain email addresses. "Label=" is
optional and is the name printed over the column; without it the directory's
display name is used, "Smith, Jane" turned round to "Jane Smith". Fewer than
four entries make fewer, wider columns; more than four are cut to four with a
warning. Pass include_team=false for the front page alone.

How each person's column is read, independently of the others:

1. Their calendar itself (Namespace.GetSharedDefaultFolder), when they have
   shared it with you or you are a delegate: subjects, locations, colours,
   tentative drawn hollow - the same reading as your own calendar, recurring
   series included. An appointment that started up to TEAM_LOOKBACK_DAYS
   earlier and is still running (a fortnight's leave booked as one item) is
   included, and anything running past either end of the day is an all-day
   chip rather than stretching the timeline to midnight. A PRIVATE
   appointment prints as "Private appointment", with no location.
2. Their FREE/BUSY, when the calendar cannot be opened - or opens but shows
   nothing while free/busy shows them booked, which is what free/busy-only
   permission can look like. Blocks then say Busy, Tentative (hollow), Out
   of office or Working elsewhere, and nothing about what. The column
   heading says "free/busy only".
3. Neither: the column prints "Calendar not available"; a name that does not
   resolve prints "Not found in the directory". One person failing never
   stops the rest of the sheet printing.

The tool's reply reports the team view as COUNTS per person, never subjects:
those are other people's diaries, and their content is for the sheet on your
desk rather than for the AI. Run --check to see, per person, whether their
calendar or only their free/busy is readable.

THE BLACKLIST AND THE PRINTED PAGE. Every event is printed, blacklisted or
not. The PDF is drawn here and written to this machine's own disk; a planner
missing the meetings that matter most is not worth carrying, and the blacklist
is there to keep marked material away from the AI, not away from the person
whose diary it is. What makes that safe is that this path NEVER READS AN
APPOINTMENT'S BODY - the body is the part that carries marked detail, and
nothing on the page or in the tool's reply comes from it. A block shows the
subject, the times and the location, and that is all.

The blacklist still decides what the AI is told. It is applied to the fields
the reply is built from (subject, location, organiser, categories), and an
event that matches is printed but left out of the reply, which reports the
count. This differs from outlook_get_calendar, which reads the body and
withholds a matching event outright - that tool's whole output goes to the
model, so it has more to protect.

The file lands in the documents folder below and an existing file of the same
name is overwritten, so re-printing a day replaces that day's sheet rather than
piling up copies.

Printable team task sheet (v11.3.0)
-----------------------------------
"Print my team's tasks" writes an A4 LANDSCAPE PDF of your direct reports'
open tasks: EIGHT COLUMNS, FOUR A SIDE, meant to be

    PRINTED DOUBLE-SIDED, FLIP ON SHORT EDGE

(the same print-dialog preset as the planner's team view). Column 1 is
"Team / unassigned" - a task emailed to the whole team, to several people, or
listed in Confluence with nobody against it - and columns 2-8 are the people in
OUTLOOK_TASKS_TEAM, in that order. Header and footer match the day planner,
with the date you printed it as the title. Each task gets an empty tick box,
its due date (red once overdue, amber when due today), and a short status /
source line; each column is sorted soonest-due first, undated tasks after. A
column with more than fits ends in "+N more not shown" rather than shrinking
its type, and the reply says which. The whole sheet is used - there is no
strip to fold, since a task list is pinned up or filed, not kept in a diary.

The tasks come from the Confluence page that collates them (which this server
cannot read) and from emails you flagged, so the work is split between three
tools and the model in the middle:

1. outlook_get_task_team returns the columns and the Confluence page that
   collates the tasks (OUTLOOK_TASKS_PAGE). It reads configuration only.
2. outlook_list_task_emails returns the emailed tasks. An Outlook rule copies
   every email you send back into your Inbox, where you either delete it or
   FLAG it with a due date to be reminded. So an emailed task is an Inbox item
   that you sent, flagged (not complete) with a real due date - nothing else.
   Sent Items is never read, there is no date window, and marking the flag
   complete in Outlook takes the task off the next sheet. The flag is found
   with a DASL filter on PR_FLAG_STATUS (an integer, so no locale fault); if
   Outlook rejects it the Inbox is walked instead, up to 5000 items. The
   content blacklist applies as usual.
3. The model reads the Confluence page through the confluence plugin, opens
   an email with outlook_get_email where the subject and preview are not
   enough, de-duplicates the two sources, and decides who owns each task.
4. outlook_print_team_tasks takes that list as {person, task, due, status,
   source, detail} and lays it out. A 'person' is matched case-insensitively
   against the roster by full name, email address, or a first name or surname
   only one person has; a name that matches nobody goes into column 1 marked
   "For <name>" and is reported back, never dropped. Without OUTLOOK_TASKS_TEAM
   the columns are made from the names on the tasks, in the order they first
   appear.

The PDF lands in the same documents folder as the planner, as
"Team Tasks - <date>.pdf"; re-printing the same day overwrites it. The skill
outlook:team-tasks carries the whole workflow.

    OUTLOOK_TASKS_TEAM="Jane Smith=jane.smith@agency.gov.au; Bob Jones=bob.jones@agency.gov.au"
    OUTLOOK_TASKS_PAGE="Team Task List"

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
   - day planner   : every event is PRINTED (the PDF is written locally, and
                     the planner never reads an appointment's body at all);
                     a blocked event is left out of the tool's reply instead.
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

3. Tunable caps (MAX_BODY_CHARS / CALENDAR_HARD_CAP / SEARCH_SCAN_CAP /
   RECURRENCE_SCAN_CAP)
   Safety/size limits. Lower MAX_BODY_CHARS if your local model has a small
   context window. The other two are guard rails you can usually leave as-is.

3b. Day-planner defaults (CALENDAR_DAY_START_HOUR / CALENDAR_DAY_END_HOUR /
    CALENDAR_LOOKAHEAD_DAYS / CALENDAR_CATEGORY_COLOURS /
    CALENDAR_HIDDEN_CATEGORIES / CALENDAR_TEAM /
    PLANNER_CONTENT_FRACTION, which OUTLOOK_CALENDAR_PAGE_FILL overrides)
   The working day the printed timeline covers, how many following days the
   right-hand panel lists, the Outlook category -> colour map, the categories
   left off the page altogether, the colleagues printed on the back, and how
   much of
   the sheet's height the page uses. All of them have an environment variable
   so a plugin install never needs the file edited.

3c. Team task sheet (TASKS_TEAM / TASKS_CONFLUENCE_PAGE, overridden by
    OUTLOOK_TASKS_TEAM / OUTLOOK_TASKS_PAGE)
   Your direct reports in column order, and the Confluence page that collates
   their tasks. See "Printable team task sheet" above.

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
for your account. The plugin prompts for nothing at install: every setting
below is an environment variable the server inherits from Claude Code. Set
them in the "env" block of Claude Code's settings.json
(%USERPROFILE%\\.claude\\settings.json, or H:\\Eva\\.claude\\settings.local.json
to keep it on H: beside the working folder), or as Windows user environment
variables, then fully restart Claude Code. For example:

    {
      "env": {
        "EVA_PYTHON": "C:\\Python311\\python.exe",
        "EVA_KNOWLEDGE_DIR": "H:\\Eva\\knowledge",
        "EVA_DOCUMENTS_DIR": "H:\\Eva\\documents",
        "OUTLOOK_SEARCH_FOLDERS": "Inbox,Sent Items,Archive",
        "OUTLOOK_MEETING_HOURS": "8-18"
      }
    }

(Backslashes are doubled inside JSON.) A blank or missing value means "not
set", and the default applies.

This server uses three of the four suite-wide variables:

    EVA_PYTHON          full path to the python.exe that has pywin32
                        installed, e.g. C:\\Python311\\python.exe (read by the
                        plugin manifest, not by this file)
    EVA_KNOWLEDGE_DIR   root of the RAG corpus (default H:\\Eva\\knowledge)
    EVA_DOCUMENTS_DIR   root of the document library (default H:\\Eva\\documents)

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

Or, as Windows user environment variables instead of settings.json
(PowerShell, one-off):

    [Environment]::SetEnvironmentVariable("EVA_PYTHON", "C:\\Python311\\python.exe", "User")
    [Environment]::SetEnvironmentVariable("EVA_KNOWLEDGE_DIR", "H:\\Eva\\knowledge", "User")
    [Environment]::SetEnvironmentVariable("EVA_DOCUMENTS_DIR", "H:\\Eva\\documents", "User")

Copy the repo's eva\\ folder to H:\\Eva and the folder exists - see
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
    OUTLOOK_CALENDAR_PAGE_FILL  how much of the sheet's HEIGHT the printed
                                planner uses: a fraction ("0.8", the default),
                                a percentage ("80" or "80%"), or "off" (or
                                none/no/false/disabled) for the whole sheet.
                                Below 1 the rest of the page prints blank with
                                a fold hairline, so it can be turned up behind
                                the planner to fit a diary. Accepts 0.5 to 1.0;
                                anything else is refused with a warning and the
                                default kept.
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
    OUTLOOK_CALENDAR_HIDE_CATEGORIES
                                comma-separated Outlook categories to leave
                                OFF the printed planner, e.g.
                                "Birthdays,Public Holidays". Matched
                                case-insensitively against the whole category
                                name. Printing only - outlook_get_calendar is
                                unaffected.
    OUTLOOK_CALENDAR_TEAM       up to four colleagues' calendars printed on
                                the BACK of the planner (double-sided, flip
                                on short edge), e.g. "Jane Smith=jane.smith@
                                agency.gov.au; bob.jones@agency.gov.au".
                                Semicolon-separated; "Label=" is the name
                                printed over the column and is optional.
                                Unset means a one-page planner. See TEAM VIEW
                                ON THE BACK above.
    OUTLOOK_TASKS_TEAM          your direct reports, up to seven, as columns
                                2-8 of the team task sheet, in print order.
                                Same form as OUTLOOK_CALENDAR_TEAM. Unset
                                means the columns come from the task names.
    OUTLOOK_TASKS_PAGE          the Confluence page (title, ID or URL) that
                                collates the team's tasks. Passed to the model,
                                which reads it with the confluence plugin.
    OUTLOOK_SEARCH_FOLDERS      comma-separated folder names for
                                outlook_search_recent, e.g.
                                "Inbox,Sent Items,Archive".
    OUTLOOK_BLACKLIST_FILE      path to a file of extra blacklist terms.
    OUTLOOK_REQUIRE_BLACKLIST=1 refuse to start with an empty blacklist.
    OUTLOOK_MEETING_HOURS       the hours a meeting may be SUGGESTED in, e.g.
                                "8-18" (default "9-17"). Deliberately separate
                                from OUTLOOK_CALENDAR_HOURS: what you want to
                                SEE on a printed day and when it is acceptable
                                to BOOK somebody are not the same window.
    OUTLOOK_GAL_SCAN_CAP        how many directory entries one fuzzy name
                                search may read before giving up (default
                                20000). Raise it on a very large GAL if people
                                go missing; a search that stops early says so.
    OUTLOOK_ALLOW_DRAFTS=false  remove outlook_draft_meeting and
                                outlook_draft_email altogether. The server is
                                then read-only on the mailbox again and neither
                                tool is offered at all, so nothing can promise
                                the user a draft.

There are NO configuration command-line flags: everything above is an
environment variable, so two settings can never disagree. The only flags this
server takes are --check and --version.

INSTALLING INTO CLAUDE CODE
---------------------------
This server ships as the "outlook" Claude Code plugin (its manifest is
.claude-plugin/plugin.json next to this file), so the normal install is:

    /plugin marketplace add C:\\path\\to\\claude-skills
    /plugin install outlook@mcnamee-claude-skills

The plugin prompts for nothing at install: the interpreter, the folders and
every server-specific setting come from the environment variables above (the
settings.json "env" block, or Windows user environment variables). Fully
restart Claude Code after changing one. PYTHONUTF8=1 is set for you by the
manifest.

To register the server by hand instead (PowerShell):

    claude mcp add outlook --scope user -e PYTHONUTF8=1 -- $env:EVA_PYTHON C:\\path\\to\\outlook.py

See README.md next to this file for the full settings reference.

USAGE / TESTING
---------------
- As an MCP server (normal mode): launched by the MCP client. Run with no
  arguments; every setting comes from the environment.
- Connectivity check (run manually on the endpoint before wiring it in):

      python outlook.py --check

  Connects to Outlook and prints mailbox diagnostics, folder paths, blacklist
  status and the recurring-series report described above to stderr, then exits.
  With OUTLOOK_CALENDAR_TEAM set it also reports, per person, whether their
  calendar or only their free/busy can be read, and today's item count.
- Testing the task sheet: set OUTLOOK_TASKS_TEAM and OUTLOOK_TASKS_PAGE, run
  --check (it prints both), then ask for "a printable list of my team's
  tasks". Print double-sided, flip on short edge: columns 1-4 on the front,
  5-8 on the back, both the right way up.
- Testing the team view: set OUTLOOK_CALENDAR_TEAM, run --check and confirm
  each person resolves, then ask for "today's printable calendar" and print
  the PDF double-sided, flip on short edge. Hold the sheet up to the light:
  the back's header should sit behind the front's header, the right way up.

IMPORTANT (stdio-on-Windows pitfalls)
-------------------------------------
- ALL diagnostic output goes to stderr. Anything on stdout that is not a
  JSON-RPC message corrupts the protocol stream.
- Set PYTHONUTF8=1 in the launching environment (the plugin manifest does this
  for you) so stdout is UTF-8 and Unicode subjects do not crash on cp1252.
"""

# Semantic version of this server. Bump on EVERY change (see CLAUDE.md):
# MAJOR = breaking config/tool change, MINOR = new feature, PATCH = fix.
__version__ = "11.3.0"

import os
import re
import sys
import zlib
import json
import difflib
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
RECURRENCE_SCAN_CAP = 30000  # ceiling on the expanded-occurrence walk (see below)
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
#        a full path of its own. The literal here is what a stock H:\Eva
#        install resolves to; it MUST stay inside the knowledge-base plugin's
#        corpus (H:\Eva\knowledge) or the saved mail would never be indexed.
#        Set OUTLOOK_KB_DIR=off to forbid saving entirely - no email is then
#        written to disk and a save_to_kb request is refused. (The day planner
#        has its own switch, OUTLOOK_DOCS_DIR, in 5b below.)
SUBFOLDER = "email"                      # this server's knowledge sub-folder
EVA_KNOWLEDGE_DIR = r"H:\Eva\knowledge"  # fallback for the suite-wide root
KB_DIR = r"H:\Eva\knowledge\email"

# --- 5b. PDF_DIR  (where outlook_print_calendar writes the printable planner).
#        The ONE folder this server creates documents in. It is the "pdf"
#        sub-folder of %EVA_DOCUMENTS_DIR%, because the document library is
#        organised by file type and a printed planner is a PDF - it lands beside
#        your own PDFs rather than in an output folder of its own.
#        RESOLVED FROM THE ENVIRONMENT in main(); the literal here is what a
#        stock H:\Eva install resolves to. Set OUTLOOK_DOCS_DIR to a full path
#        of its own, or to "off" to forbid printing entirely.
DOCS_SUBFOLDER = "pdf"                     # this server's documents sub-folder
EVA_DOCUMENTS_DIR = r"H:\Eva\documents"     # fallback for the suite-wide root
PDF_DIR = r"H:\Eva\documents\pdf"

# --- 5c. Printable day planner defaults.
#        CALENDAR_DAY_START_HOUR / CALENDAR_DAY_END_HOUR are the working day the
#        timeline shows. It always STRETCHES to take in anything scheduled
#        outside those hours, so they are a floor and a ceiling on the printed
#        grid, not a filter. Override with OUTLOOK_CALENDAR_HOURS="7-19".
CALENDAR_DAY_START_HOUR = 8
CALENDAR_DAY_END_HOUR = 18

#        How many following days the right-hand panel lists (1-6).
CALENDAR_LOOKAHEAD_DAYS = 4

#        How much of the sheet's HEIGHT the page uses lives with the rest of
#        the planner's geometry further down, as PLANNER_CONTENT_FRACTION
#        (0.8 - the bottom fifth prints blank to fold up behind the page).
#        Override with OUTLOOK_CALENDAR_PAGE_FILL="0.8" / "80" / "off".

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

#        Outlook CATEGORIES to leave OFF the printed planner altogether, e.g.
#        ["Birthdays"]. An appointment carrying ANY of them is not drawn in
#        either panel, does not stretch the timeline and does not make a day
#        count as busy. Matched case-insensitively against the whole category
#        name. Merged with OUTLOOK_CALENDAR_HIDE_CATEGORIES (comma-separated).
#        Only the printed planner honours this - outlook_get_calendar still
#        reports everything the blacklist allows.
CALENDAR_HIDDEN_CATEGORIES = [
]

# --- 5d. Team view on the BACK of the planner (new in v11.2.0).
#        Up to four other people's calendars - your manager and colleagues -
#        printed as four side-by-side day columns on a second page, so the
#        planner prints double-sided (flip on short edge) with your own day on
#        the front and theirs on the back. Each entry is an email address or
#        a directory name, optionally with the name to print over its column
#        in front of an "=":
#            ["Jane Smith=jane.smith@agency.gov.au", "bob.jones@agency.gov.au"]
#        Without a label the directory's display name is printed ("Smith,
#        Jane" is turned round to "Jane Smith"). Empty means no back page.
#        OUTLOOK_CALENDAR_TEAM REPLACES this list (it does not add to it):
#        semicolon-separated entries of the same form.
CALENDAR_TEAM = [
]

# --- 5e. Team task sheet (outlook_print_team_tasks, new in v11.3.0).
#        Your direct reports, up to seven, in the order their columns print
#        after the "Team / unassigned" column - same "Label=address" form as
#        CALENDAR_TEAM. Empty means the columns are made from the names on the
#        tasks instead, in the order they first appear.
#        OUTLOOK_TASKS_TEAM REPLACES this list (semicolon-separated entries).
TASKS_TEAM = [
]
#        The Confluence page that collates the team's tasks (title, page ID or
#        URL). This server cannot read Confluence; it only hands the value to
#        the model, which reads the page through the confluence plugin.
#        Override with OUTLOOK_TASKS_PAGE. The emailed tasks need no setting:
#        they are the emails you sent that sit flagged, with a due date, in
#        your Inbox (see outlook_list_task_emails).
TASKS_CONFLUENCE_PAGE = ""

# --- 6. KB_AUTOSAVE. Whether reading an email saves it WITHOUT being asked.
#        False means a message is saved only when the call passes
#        save_to_kb=true, which is the point: mail skimmed to answer a question
#        is not a filing decision, and a knowledge base full of correspondence
#        nobody chose returns irrelevant answers later. Set to True here, or
#        set OUTLOOK_KB_AUTOSAVE=true, to save every email read.
KB_AUTOSAVE = False

# --- 6b. Meeting scheduling (outlook_find_people, outlook_suggest_meeting_times
#         and outlook_draft_meeting).
# The hours a meeting may be SUGGESTED in. Deliberately separate from the
# printed planner's CALENDAR_DAY_START_HOUR/END_HOUR above: the planner's
# window is what you want to SEE (start it at 6 and a dawn flight is on the
# page), while this one is when it is acceptable to BOOK somebody. Override
# with OUTLOOK_MEETING_HOURS, e.g. "8-18".
MEETING_START_HOUR = 9
MEETING_END_HOUR = 17

# Ceiling on directory entries read during one fuzzy name search. An
# enterprise Global Address List can hold six figures of entries and each one
# costs a COM round trip, so the scan is bounded; a search that stops early
# says so in its reply. Override with OUTLOOK_GAL_SCAN_CAP.
GAL_SCAN_CAP = 20000

# Whether outlook_draft_meeting and outlook_draft_email exist at all. They are
# the only tools in this server that write to the MAILBOX - an unsent meeting
# saved into the Calendar, and an unsent email opened in Outlook and saved to
# Drafts - which the user then reviews and sends themselves. Set it False (or
# OUTLOOK_ALLOW_DRAFTS=false) and neither is offered, leaving the server
# read-only on the mailbox. Nothing here can ever SEND.
ALLOW_DRAFTS = True

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


def parse_category_list(raw):
    """
    Parse OUTLOOK_CALENDAR_HIDE_CATEGORIES ("Birthdays, Public Holidays") into a
    set of lower-cased category names. Blank entries are skipped, so a trailing
    comma is harmless.
    """
    return set(part.strip().lower() for part in (raw or "").split(",")
               if part.strip())


def parse_team(raw, limit=None, what="team calendars", page="the back page"):
    """
    Parse OUTLOOK_CALENDAR_TEAM into a list of (label, spec) pairs.

    OUTLOOK_TASKS_TEAM uses the same form, and passes its own `limit` (seven
    people) and wording for the warning.

    Entries are separated by semicolons, the separator Outlook itself uses for
    a list of people, because a directory name such as "Smith, Jane" has a
    comma IN it. Commas are still accepted between entries when every piece
    holds an "@" (a plain list of email addresses), since that is how most
    people will type one. Each entry is "Label=address" or just "address";
    the label is the name printed over the column. More than
    TEAM_MAX_COLUMNS entries are cut to that many with a warning - the page
    has room for four columns and no more.
    """
    entries = []
    for chunk in (raw or "").split(";"):
        chunk = chunk.strip()
        if not chunk:
            continue
        parts = [part.strip() for part in chunk.split(",") if part.strip()]
        if len(parts) > 1 and all("@" in part for part in parts):
            pieces = parts
        else:
            pieces = [chunk]
        for piece in pieces:
            label, sep, spec = piece.partition("=")
            if not sep:
                label, spec = "", piece
            label, spec = _one_line(label), _one_line(spec)
            if not spec:
                log("WARNING: ignoring an entry in {0}, '{1}' - there is "
                    "nothing after the '='.".format(what, piece))
                continue
            entries.append((label, spec))
    limit = TEAM_MAX_COLUMNS if limit is None else limit
    if len(entries) > limit:
        log("WARNING: {0} {1} are configured but {2} holds {3}; printing the "
            "first {3} only.".format(len(entries), what, page, limit))
        entries = entries[:limit]
    return entries


def parse_page_fill(raw):
    """
    OUTLOOK_CALENDAR_PAGE_FILL -> the fraction of the sheet's height to use.

    Accepts a fraction ("0.8"), a percentage ("80" or "80%") and, like every
    other optional setting in this suite, a DISABLE_KEYWORD - here meaning "no
    strip to fold up", so the planner takes the whole sheet. A value below
    PLANNER_MIN_CONTENT_FRACTION or above 1 is refused: a planner squeezed into
    a third of a page is not a planner, and anything over the sheet is off it.

    Returns the fraction, or None to keep the built-in default (which is also
    what a bad value gets, after a warning - a typo here must not stop the
    server printing).
    """
    if raw is None:
        return None
    text = raw.strip().lower().rstrip("%").strip()
    if text in DISABLE_KEYWORDS or text == "full":
        return 1.0
    try:
        value = float(text)
    except ValueError:
        log("WARNING: OUTLOOK_CALENDAR_PAGE_FILL='{0}' is not a number like "
            "'0.8' or '80'; keeping {1}.".format(raw, PLANNER_CONTENT_FRACTION))
        return None
    if value > 1.0:            # a percentage, written the way people say it
        value /= 100.0
    if not PLANNER_MIN_CONTENT_FRACTION <= value <= 1.0:
        log("WARNING: OUTLOOK_CALENDAR_PAGE_FILL='{0}' is outside {1}-1.0 "
            "({2:.0f}%-100%); keeping {3}.".format(
                raw, PLANNER_MIN_CONTENT_FRACTION,
                PLANNER_MIN_CONTENT_FRACTION * 100, PLANNER_CONTENT_FRACTION))
        return None
    return value


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
    import pywintypes
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

# Lower-cased Outlook categories left off the printed planner, from
# CALENDAR_HIDDEN_CATEGORIES and OUTLOOK_CALENDAR_HIDE_CATEGORIES.
_HIDDEN_CATEGORIES = set()

# The (label, spec) pairs printed as columns on the back of the planner, from
# CALENDAR_TEAM or OUTLOOK_CALENDAR_TEAM. Empty means a one-page planner.
_CALENDAR_TEAM = []

# The team task sheet's roster and Confluence page, from the TASKS_* settings
# above or their OUTLOOK_TASKS_* environment variables.
_TASKS_TEAM = []
_TASKS_PAGE = ""

# Effective default folder set for outlook_search_recent. Initialised from
# SEARCH_ALL_FOLDERS; may be replaced by OUTLOOK_SEARCH_FOLDERS. A per-call
# "folders" argument still overrides this.
_SEARCH_FOLDERS = list(SEARCH_ALL_FOLDERS)

# Effective meeting-scheduling settings. Initialised from the configuration
# block above; main() may replace them from OUTLOOK_MEETING_HOURS,
# OUTLOOK_GAL_SCAN_CAP and OUTLOOK_ALLOW_DRAFTS.
_MEETING_START_HOUR = MEETING_START_HOUR
_MEETING_END_HOUR = MEETING_END_HOUR
_GAL_SCAN_CAP = GAL_SCAN_CAP
_ALLOW_DRAFTS = ALLOW_DRAFTS


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


def appointment_visible_block_reason(item):
    """
    As appointment_block_reason, but over ONLY the fields that leave this server.

    The printed planner never reads an appointment's Body. The body is the part
    that carries protectively-marked detail, and nothing on the page or in the
    tool's reply is drawn from it - a block shows the subject, the times and the
    location, and that is all. What can reach the AI is therefore the subject,
    location, organiser and categories, so those are what this clears.

    An event blocked here is still PRINTED: the page is written locally, on the
    user's own paper, and a diary with holes in it is not worth printing. It is
    left out of the text handed back to the model, which is the only part of
    this tool the AI ever sees.

    Fail-safe: a subject that cannot be read cannot be cleared, so it counts as
    blocked.
    """
    if _BLACKLIST_RE is None:
        return None

    parts = []
    try:
        parts.append(item.Subject or "")
    except Exception:
        return "<unreadable subject>"

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


def _int_arg(args, name, default):
    """
    Read an integer tool argument, tolerating a missing or null value.

    An MCP client happily sends "duration_minutes": null for a field the model
    left out, and int(None) would surface as an unhelpful "Outlook tool error".
    """
    value = args.get(name)
    if value is None or value == "":
        return default
    return int(value)


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


def _series_times(item, pattern):
    """
    The candidate times of day a series' occurrences can start at.

    RecurrencePattern.StartTime is the authority: it is the pattern's own time
    of day. The master's Start is only its FIRST occurrence, which for a series
    set up years ago can sit at a different wall-clock time from today's - a
    daylight-saving offset that has since drifted, or a series stored in the
    organiser's timezone rather than ours. Both are tried, because probing a
    time nothing starts at costs one failed call and can never invent an
    occurrence that is not really there.
    """
    times = []
    for getter in (lambda: pattern.StartTime, lambda: item.Start):
        try:
            value = _com_to_naive(getter())
        except Exception:
            value = None
        if value is not None and value.time() not in times:
            times.append(value.time())
    return times or [datetime.time(0, 0)]


def _expand_recurring_occurrences(item, start_dt, end_dt, report=None):
    """All occurrences of a recurring appointment starting in [start_dt, end_dt].

    Probes the series' RecurrencePattern day by day at each time the series
    could start (GetOccurrence raises for days with no/deleted occurrence, which
    is normal and simply skipped), then adds rescheduled occurrences from the
    pattern's Exceptions, since those no longer sit at the usual time. Pure COM
    object access - no locale-sensitive date strings anywhere.

    This can only find an occurrence whose start time it guessed in advance, so
    it is the BACKSTOP, not the primary: scan_calendar also asks Outlook to
    expand the series itself and takes the union. `report`, when given a dict,
    collects what happened for the --check diagnostic.

    Returns a list of (naive start datetime, AppointmentItem occurrence).
    """
    try:
        pattern = item.GetRecurrencePattern()
    except Exception as exc:
        if report is not None:
            report["error"] = "GetRecurrencePattern failed: {0}".format(exc)
        return []

    times = _series_times(item, pattern)
    if report is not None:
        report["probed_times"] = [value.strftime("%H:%M") for value in times]

    found = {}  # start datetime -> occurrence (dedups probe vs exception hits)
    day = start_dt.date()
    while day <= end_dt.date():
        for base_time in times:
            probe = datetime.datetime.combine(day, base_time)
            try:
                occ = pattern.GetOccurrence(probe)
                occ_start = _com_to_naive(occ.Start)
                if occ_start is not None and start_dt <= occ_start <= end_dt:
                    found[occ_start] = occ
            except Exception:
                pass  # no occurrence at this time on this day
        day = day + datetime.timedelta(days=1)

    if report is not None:
        report["by_probe"] = len(found)

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
    except Exception as exc:
        if report is not None:
            report["error"] = "Exceptions unreadable: {0}".format(exc)

    if report is not None:
        report["total"] = len(found)
    return list(found.items())


def _expanded_occurrences(start_dt, end_dt, report=None, folder=None,
                          overlap_from=None):
    """
    Ask OUTLOOK to expand every recurring series across the window.

    Setting IncludeRecurrences on a Start-sorted Items collection makes Outlook
    hand back individual occurrences rather than series masters. That matters
    because Outlook holds each series' timezone and daylight-saving rules, so it
    places an occurrence correctly where a hand-rolled probe has to guess the
    time of day and silently finds nothing when the guess is wrong. That guess
    is exactly what goes stale on a long-standing series, or one whose organiser
    sits in another timezone.

    The collection is unbounded for a series with no end date, so it is walked
    in Start order and abandoned as soon as it passes the window (an ascending
    sort means everything after that point is later still). RECURRENCE_SCAN_CAP
    stops a pathological calendar walking forever; hitting it is reported rather
    than passed off as an empty day.

    Still NO locale-sensitive date strings: "[Start]" is a property name, not a
    formatted date, so this does not reintroduce the Restrict() fault that makes
    filtering depend on regional settings. Returns (naive start, item) pairs.

    `folder` is the calendar to read (your own when None). With
    `overlap_from`, an item that STARTED between then and start_dt is kept too
    when it is still running at start_dt - a week of leave booked as one
    appointment, say. Those items are walked past anyway, so the only extra
    cost is reading End on the few that start inside that lookback.
    """
    if folder is None:
        folder = get_namespace().GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = folder.Items
    try:
        # Order matters: Outlook requires the ascending Start sort to be in
        # place before it will expand recurrences into the collection.
        items.Sort("[Start]")
        items.IncludeRecurrences = True
    except Exception as exc:
        if report is not None:
            report["error"] = "could not enable recurrence expansion: {0}".format(exc)
        return []

    found = []
    walked = 0
    try:
        item = items.GetFirst()
    except Exception as exc:
        if report is not None:
            report["error"] = "GetFirst failed: {0}".format(exc)
        return []

    while item is not None and walked < RECURRENCE_SCAN_CAP:
        walked += 1
        try:
            item_start = _com_to_naive(item.Start)
        except Exception:
            item_start = None

        if item_start is not None:
            if item_start > end_dt:
                break  # sorted ascending: everything from here is later still
            if item_start >= start_dt:
                found.append((item_start, item))
            elif overlap_from is not None and item_start >= overlap_from:
                if _still_running(item, start_dt):
                    found.append((item_start, item))

        try:
            item = items.GetNext()
        except Exception:
            break

    if report is not None:
        report["walked"] = walked
        report["found"] = len(found)
        if walked >= RECURRENCE_SCAN_CAP:
            report["error"] = ("hit the {0}-item scan cap before reaching the "
                               "window".format(RECURRENCE_SCAN_CAP))
    return found


def _still_running(item, when):
    """True when an appointment's End is after `when` (False if unreadable)."""
    try:
        item_end = _com_to_naive(item.End)
    except Exception:
        return False
    return item_end is not None and item_end > when


def _occurrence_key(start, item):
    """Identity of one occurrence, for merging the two expansion passes."""
    try:
        entry_id = item.EntryID or ""
    except Exception:
        entry_id = ""
    return (start, entry_id)


def scan_calendar(start_dt, end_dt, folder=None, overlap_from=None):
    """
    Walk the whole calendar folder and return every occurrence inside a window.

    Returns (matches, total, debug_tail): `matches` is a list of (naive start
    datetime, appointment COM object) sorted by start; `total` is how many items
    the folder holds; `debug_tail` is a rolling last-5 of
    (index, start, matched?) so an empty result can prove what was scanned.

    Recurring series are gathered TWICE and merged, because either pass alone
    drops meetings. Outlook's own expansion (_expanded_occurrences) knows each
    series' timezone and daylight-saving rules; the per-master probe
    (_expand_recurring_occurrences) has to guess the time of day an occurrence
    starts at, and finds nothing when the guess is wrong, which is what hides a
    long-standing series whose offset has drifted since it was set up. Taking
    the union means a fault in either one costs nothing.

    NO Outlook-side filtering happens here. Restrict() formats its date strings
    per the machine's regional settings (US vs AU) and misbehaves with recurring
    appointments, silently returning 0 results on some machines. Instead the
    whole collection is walked by index (Item(i) is more reliable than COM
    enumeration) and filtered in Python, which is locale-independent.

    `folder` is the calendar to scan - your own when None, or a colleague's
    shared calendar for the team view. `overlap_from` also keeps a single
    (non-recurring) appointment that started between then and start_dt and
    is still running at start_dt, so a multi-day block of leave shows on the
    days after its first. See _expanded_occurrences.
    """
    if folder is None:
        folder = get_namespace().GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = folder.Items
    total = int(items.Count)  # raises if the folder cannot be read; caller reports

    matches = []      # (naive start datetime, appointment COM object)
    seen = set()      # occurrence keys already collected, across both passes
    debug_tail = []   # rolling last-5 raw items: (index, start-or-None, matched?)

    def collect(start, item):
        """Add one occurrence unless the other pass already found it."""
        key = _occurrence_key(start, item)
        if key in seen:
            return False
        seen.add(key)
        matches.append((start, item))
        return True

    # Pass 1: Outlook expands the recurring series itself.
    expansion = {}
    for occ_start, occ in _expanded_occurrences(start_dt, end_dt, expansion,
                                                folder=folder,
                                                overlap_from=overlap_from):
        collect(occ_start, occ)
        if len(matches) >= CALENDAR_HARD_CAP:
            break
    if expansion.get("error"):
        log("Outlook's own recurrence expansion did not complete ({0}); "
            "falling back to probing each series.".format(expansion["error"]))

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
            for occ_start, occ in occurrences:
                collect(occ_start, occ)
        else:
            matched = item_start is not None and start_dt <= item_start <= end_dt
            if (not matched and overlap_from is not None and item_start is not None
                    and overlap_from <= item_start < start_dt):
                matched = _still_running(item, start_dt)
            if matched:
                collect(item_start, item)

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
    A PDF of one or more same-sized pages, drawn from rectangles, rules and
    lines of text.

    Coordinates are MILLIMETRES FROM THE TOP-LEFT corner, because that is how a
    page design is described; the PDF's own bottom-left origin in points is an
    internal detail of this class. Font sizes are in points, as in any layout.
    Drawing goes onto the latest page; new_page() starts another.

    `duplex_short_edge` asks the reader's print dialog to default to
    double-sided, flipped on the short edge (the PDF 1.7 /Duplex viewer
    preference). It is a PRESET, not a command: Adobe Reader and Acrobat
    honour it, other viewers may ignore it and print with their own setting.
    """

    def __init__(self, width_mm, height_mm, title="", author=""):
        self.width_mm = float(width_mm)
        self.height_mm = float(height_mm)
        self.title = title
        self.author = author
        self.duplex_short_edge = False
        self._pages = [[]]
        self._ops = self._pages[0]

    def new_page(self):
        """Start a fresh page; everything drawn from now on lands on it."""
        self._ops = []
        self._pages.append(self._ops)

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
        """Serialise every page into a complete PDF file."""
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

        # Objects 1-5 are fixed; each page then takes two more (the page and
        # its content stream), so page i is object 6 + 2i.
        first_page = 6
        kids = " ".join("{0} 0 R".format(first_page + 2 * index)
                        for index in range(len(self._pages)))
        catalog = b"<< /Type /Catalog /Pages 2 0 R"
        if self.duplex_short_edge:
            catalog += b" /ViewerPreferences << /Duplex /DuplexFlipShortEdge >>"
        catalog += b" >>"

        objects = [
            catalog,
            "<< /Type /Pages /Kids [{0}] /Count {1} >>".format(
                kids, len(self._pages)).encode("ascii"),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
            (b"<< /Title " + literal(self.title) +
             b" /Author " + literal(self.author) +
             b" /Creator (outlook-mcp " + __version__.encode("ascii") + b")" +
             b" /Producer (outlook-mcp " + __version__.encode("ascii") + b")" +
             b" /CreationDate (" + created.encode("ascii") + b") >>"),
        ]
        for index, ops in enumerate(self._pages):
            content = zlib.compress(b"\n".join(ops))
            objects.append(
                ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0} {1}] "
                 "/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                 "/Contents {2} 0 R >>").format(
                    self._num(self.width_mm * PDF_MM),
                    self._num(self.height_mm * PDF_MM),
                    first_page + 2 * index + 1).encode("ascii"))
            objects.append(
                b"<< /Length " + str(len(content)).encode("ascii") +
                b" /Filter /FlateDecode >>\nstream\n" + content + b"\nendstream")

        # /Duplex is a PDF 1.7 viewer preference, so a file that carries it
        # says 1.7; everything else here is plain 1.4.
        version = b"1.7" if self.duplex_short_edge else b"1.4"
        out = bytearray(b"%PDF-" + version + b"\n%\xe2\xe3\xcf\xd3\n")
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
        out += ("trailer\n<< /Size {0} /Root 1 0 R /Info 5 0 R >>\n"
                "startxref\n{1}\n%%EOF\n").format(
                    len(objects) + 1, xref_at).encode("ascii")
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


# ---------------------------------------------------------------------------
# Timeline geometry
# ---------------------------------------------------------------------------

PLANNER_MIN_BLOCK_MM = 6.4   # a short meeting is grown to this if there is room
PLANNER_MIN_TIGHT_MM = 3.4   # ...and never squeezed below this, so a line still fits
PLANNER_BLOCK_GAP_MM = 1.6   # white space between two blocks, side by side or stacked

# ONE size for every event subject on the page, whatever the block it sits on.
# The size a subject is set at used to follow the size of its block, which made
# an hour-long meeting shout and a fifteen-minute one whisper for no reason
# other than its duration; a printed page reads better when the type is even.
# 7pt is the smaller of the sizes it used to pick from, and its cap height
# (about 1.8 mm) still fits the thinnest block PLANNER_MIN_TIGHT_MM allows.
PLANNER_TITLE_PT = 7.0

# How much of the sheet's HEIGHT the planner prints on, as a fraction. At 0.8
# the page is laid out in the top four fifths of an A4 landscape sheet and the
# bottom fifth comes out blank, so it can be folded up behind the rest and the
# planner fits a diary. The PDF is still a full A4 landscape page - only the
# ink stops early - so it prints on ordinary paper with no scaling.
# OUTLOOK_CALENDAR_PAGE_FILL overrides this: a fraction (0.8), a percentage
# (80 or 80%), or one of the DISABLE_KEYWORDS for the whole sheet.
PLANNER_CONTENT_FRACTION = 0.8

# How little of the sheet the planner may be squeezed into. Below this the
# header, the hour grid and the footer stop fitting, so a smaller value is
# refused rather than printed as a mess.
PLANNER_MIN_CONTENT_FRACTION = 0.5


def _place_blocks(events, day, start_min, end_min, top, bottom, lane_x, lane_w):
    """
    Work out where each event's block goes on the timeline.

    Two rules do the work. Events that genuinely overlap share the lane
    Outlook-style, splitting a cluster into as many columns as it needs. And a
    short meeting is grown to a readable height, but only into space nothing
    else is using: a day of back-to-back half-hour meetings stays one honest
    column of thin blocks instead of zig-zagging across two.

    Growth is worked out LAST, against every other block on the lane rather
    than only the ones in the same cluster. A 15-minute meeting at 09:45 and
    the hour at 10:00 are separate clusters - they do not overlap - so a
    cluster-local ceiling let the grown short block print straight over the
    real one.
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
        for box in cluster["boxes"]:
            box["x"] = lane_x + box["column"] * (width + PLANNER_BLOCK_GAP_MM)
            box["width"] = width

    # Heights come last, once every block knows where it sits across the lane,
    # so the ceiling a block may grow up to is the next block ANYWHERE in its
    # way - same cluster or not.
    for box in boxes:
        ceiling = bottom
        for other in boxes:
            if other is box or other["top"] <= box["top"] + 0.01:
                continue
            # Only a block sharing some of this one's width is in its way; two
            # columns of the same cluster sit side by side and are not.
            if (other["x"] < box["x"] + box["width"] - 0.01
                    and box["x"] < other["x"] + other["width"] - 0.01):
                ceiling = min(ceiling, other["top"])
        room = ceiling - box["top"] - PLANNER_BLOCK_GAP_MM
        # The gap comes out of the block itself, not just out of the overflow
        # case, so two back-to-back meetings read as two blocks with white
        # between them rather than one continuous slab.
        natural = max(box["natural"] - PLANNER_BLOCK_GAP_MM, 0.0)
        height = max(natural, PLANNER_MIN_BLOCK_MM)
        box["height"] = max(PLANNER_MIN_TIGHT_MM, min(height, max(room, 0.0)))
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


def _draw_block(canvas, box):
    """
    Draw one event block on the timeline.

    Three layouts, picked by how much room the block actually has: the full
    stacked one, a narrow variant for a block sharing its lane with an
    overlapping meeting, and a single centred line for anything too short to
    stack. A printed block that has been squeezed says less rather than
    truncating everything on it into nonsense.

    Type is deliberately small and, for the subject, ALWAYS PLANNER_TITLE_PT -
    a block does not get bigger type for being a longer meeting. A day planner
    is read at desk distance, and the subject fitting on the block matters more
    than the type being large enough to read across a room.
    """
    event = box["event"]
    style = box["style"]
    x, y = box["x"], box["top"]
    width, height = box["width"], box["height"]
    pad_x, pad_y = 2.0, 1.2
    inner = width - pad_x * 2
    if inner <= 4:
        return

    canvas.rect(x, y, width, height, fill=style["fill"], stroke=style["stroke"],
                line_pt=style["line_pt"] or 0.6, radius=1.6)

    span = "{0}–{1}".format(_fmt_time(event["start"]), _fmt_time(event["end"]))
    floor_y = y + height - 0.8

    # --- too short to stack: one centred line ------------------------------
    if height < 7.4:
        size = PLANNER_TITLE_PT
        baseline = y + height / 2.0 + _cap_mm(size) / 2.0
        if inner < 36.0:
            # Short AND sharing the lane: one line is all there is, so spend it
            # on the subject. Where the block sits already says when it is.
            canvas.text(x + pad_x, baseline, event["subject"], size, bold=True,
                        colour=style["title"], max_width=inner)
            return
        detail = span
        if event["location"]:
            longer = span + " · " + event["location"]
            if canvas.text_width(longer, 6.0) < inner * 0.62:
                detail = longer
        detail_w = canvas.text_width(detail, 6.0)
        title_w = inner - detail_w - 2.0
        if title_w < 12.0:               # no room for both - the time wins
            detail, detail_w = span, canvas.text_width(span, 6.0)
            title_w = inner - detail_w - 2.0
        if title_w < 9.0:                # nor for that - print the title alone
            canvas.text(x + pad_x, baseline, event["subject"], size, bold=True,
                        colour=style["title"], max_width=inner)
            return
        drawn = canvas.text(x + pad_x, baseline, event["subject"], size, bold=True,
                            colour=style["title"], max_width=title_w)
        canvas.text(x + pad_x + drawn + 2.0, baseline, detail, 6.0,
                    colour=style["sub"], max_width=inner - drawn - 2.0)
        return

    # --- narrow, because it is sharing the lane ----------------------------
    if inner < 36.0:
        title_size = PLANNER_TITLE_PT
        baseline = y + pad_y + _cap_mm(title_size)
        # Two lines of subject beat one truncated line plus a time: the block's
        # position on the grid already says when it is, and what the meeting is
        # cannot be read off anything else.
        lines = canvas.wrap(event["subject"], title_size, inner, bold=True,
                            max_lines=2 if height >= 7.8 else 1)
        for line in lines:
            if baseline > floor_y:
                break
            canvas.text(x + pad_x, baseline, line, title_size, bold=True,
                        colour=style["title"])
            baseline += 2.6
        if baseline - 0.4 <= floor_y:
            # Too narrow for both ends of the range: the start time alone beats
            # a truncated one, and the block's position says how long it runs.
            shown_span = (span if canvas.text_width(span, 6.0, bold=True) <= inner
                          else _fmt_time(event["start"]))
            canvas.text(x + pad_x, baseline, shown_span, 6.0, bold=True,
                        colour=style["sub"], max_width=inner)
            baseline += 2.3
        if event["location"] and baseline <= floor_y:
            canvas.text(x + pad_x, baseline, event["location"], 6.0,
                        colour=style["sub"], max_width=inner)
        return

    # --- the full block ----------------------------------------------------
    title_size = PLANNER_TITLE_PT
    span_w = canvas.text_width(span, 6.5, bold=True)
    baseline = y + pad_y + _cap_mm(title_size)
    canvas.text(x + pad_x, baseline, event["subject"], title_size, bold=True,
                colour=style["title"], max_width=inner - span_w - 2.5)
    canvas.text(x + width - pad_x, baseline, span, 6.5, bold=True,
                colour=style["sub"], align="right")

    baseline += 2.6
    if event["location"] and baseline <= floor_y:
        canvas.text(x + pad_x, baseline, event["location"], 6.5, bold=True,
                    colour=style["sub"], max_width=inner)


def render_day_planner(day, day_events, ahead, identity, hours,
                       category_colours=None, team=None):
    """
    Build the whole A4-landscape planner and return it as PDF bytes.

    `day_events` is today's list, `ahead` a list of (date, events) for the
    right-hand panel, `identity` the (name, address) the header and footer show,
    and `hours` the (first, last) hour the timeline covers.

    `team`, when given, is the list of columns from _team_columns(): a SECOND
    page is added with those calendars side by side for the same day, and the
    PDF asks the print dialog for double-sided, flip on short edge, so the team
    view lands on the back of the sheet.

    Every event given is drawn. Nothing is withheld from the page: the PDF is
    built here and written to the endpoint's own disk, so the content blacklist
    has nothing to protect at this point - it governs the text the tool hands
    back to the AI instead.

    Attendees are never printed. Who is in a meeting is in the invitation, it
    crowds out the subject and location that a printed planner is actually for,
    and it puts other people's names on a sheet that gets left on a desk. They
    are still READ, because whether anyone outside our own domain is invited is
    what colours an uncategorised block.
    """
    category_colours = category_colours or {}
    page_w, page_h = 297.0, 210.0
    # Everything is laid out inside content_h, the top slice of the sheet, so
    # the rest comes off the printer blank and folds up behind the page.
    content_h = page_h * PLANNER_CONTENT_FRACTION
    pad_l = pad_r = 11.0
    pad_t, pad_b = 10.0, 8.0
    name, address = identity

    canvas = PdfCanvas(page_w, page_h,
                       title="Daily agenda - {0}".format(_fmt_day_long(day)),
                       author=name or address or "")

    timed = [event for event in day_events if not event["all_day"]]
    all_day = [event for event in day_events if event["all_day"]]

    for event in day_events:
        _label, event["colour"] = _event_bucket(event, category_colours)

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

    # No colour legend across the top. The blocks carry the colour, a printed
    # page is looked at rather than decoded, and the space reads better given
    # back to the date.
    canvas.text(page_w - pad_r, pad_t + 8.0, stats, 7.5, bold=True,
                colour=_FAINT, align="right", max_width=page_w - pad_l - pad_r
                - date_w - 8.0)

    header_rule = pad_t + 12.8
    canvas.line(pad_l, header_rule, page_w - pad_r, header_rule, _INK, 1.2)

    # ---- page frame ------------------------------------------------------
    body_top = header_rule + 5.0
    footer_rule = content_h - pad_b - 4.6
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
            chip_w = min(left_w, canvas.text_width(label, PLANNER_TITLE_PT,
                                                   bold=True) + 5.0)
            if chip_x + chip_w > left_x + left_w and chip_x > left_x:
                if rows >= 2:
                    canvas.text(chip_x + 1.0,
                                chip_y + chip_h / 2.0
                                + _cap_mm(PLANNER_TITLE_PT) / 2.0,
                                "+{0} more".format(
                                    len(all_day) - all_day.index(event)),
                                PLANNER_TITLE_PT, bold=True, colour=_MUTED)
                    break
                rows += 1
                chip_x = left_x
                chip_y += chip_h + 1.2
            canvas.rect(chip_x, chip_y, chip_w, chip_h, fill=style["fill"],
                        stroke=style["stroke"], line_pt=style["line_pt"] or 0.6,
                        radius=1.2)
            canvas.text(chip_x + 2.5,
                        chip_y + chip_h / 2.0 + _cap_mm(PLANNER_TITLE_PT) / 2.0,
                        label, PLANNER_TITLE_PT, bold=True,
                        colour=style["title"], max_width=chip_w - 5.0)
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
            _draw_block(canvas, box)
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
    time_col = 16.0
    date_col = 17.0
    date_block_mm = 12.4   # the DOW / number / month stack down the left
    row_gap_mm = 2.8
    base_line_mm = 3.6

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
            line_h = max(base_line_mm, min(5.4, (row_h - row_gap_mm - 2.6)
                                           / max(1, len(ahead_events))))
        capacity = max(1, int((row_h - row_gap_mm - 2.2) / line_h))
        baseline = top + 2.6

        if not ahead_events:
            canvas.text(list_x, baseline, "Nothing scheduled", 7.0, colour=_FAINT)
        else:
            shown = ahead_events[:capacity]
            if len(ahead_events) > capacity:
                shown = ahead_events[:max(1, capacity - 1)]
            for event in shown:
                when = ("all day" if event["all_day"]
                        else "{0}–{1}".format(_fmt_time(event["start"]),
                                              _fmt_time(event["end"])))
                # The time stays bold: it is the column the eye runs down.
                # The subject is set regular, which also buys back the width
                # bold was costing it.
                canvas.text(list_x, baseline, when, 6.5, bold=True, colour=_MUTED,
                            max_width=time_col - 1.0)
                canvas.text(list_x + time_col + 2.0, baseline, event["subject"], 7.0,
                            colour=_INK, max_width=list_w - time_col - 2.0)
                baseline += line_h
            if len(ahead_events) > len(shown):
                canvas.text(list_x + time_col + 2.0, baseline,
                            "+{0} more".format(len(ahead_events) - len(shown)),
                            6.5, bold=True, colour=_FAINT)

        if index < len(ahead) - 1:
            rule_y = top + row_h - row_gap_mm / 2.0
            canvas.line(right_x, rule_y, right_x + right_w, rule_y, _HAIRLINE, 0.5)
        top += row_h

    # ---- footer ----------------------------------------------------------
    canvas.line(pad_l, footer_rule, page_w - pad_r, footer_rule, _RULE, 0.5)
    left_footer = "Outlook Calendar" + (" · " + address if address else "")
    canvas.text(pad_l, footer_rule + 3.4, left_footer, 7.0, bold=True,
                colour=_FAINT, max_width=170.0)
    printed = datetime.datetime.now()
    canvas.text(page_w - pad_r, footer_rule + 3.4,
                "Printed {0} at {1}".format(
                    _fmt_day_short(printed.date()), _fmt_time(printed)),
                7.0, bold=True, colour=_FAINT, align="right")

    # ---- fold guide ------------------------------------------------------
    # The line the blank bottom of the sheet folds back on. It is a hairline in
    # the palest grey the page uses: enough to fold against, faint enough that
    # a sheet nobody folds does not look like it is missing something.
    if content_h < page_h - 1.0:
        canvas.line(pad_l, content_h, page_w - pad_r, content_h, _RULE, 0.4,
                    dash=[1.4, 2.6])

    if team:
        canvas.new_page()
        canvas.duplex_short_edge = True
        _draw_team_page(canvas, day, team, category_colours)

    return canvas.to_bytes()


# ---------------------------------------------------------------------------
# The team view - other people's calendars on the back of the planner
# ---------------------------------------------------------------------------

# The back page holds this many columns at most: four is what fits A4
# landscape with a meeting subject still readable on each block.
TEAM_MAX_COLUMNS = 4

# How far back a colleague's calendar is searched for an appointment that is
# still running on the day - a fortnight's leave booked as one item, say.
TEAM_LOOKBACK_DAYS = 31

# The back page never shrinks below this fraction of the sheet, however much
# of the front is left blank to fold up (see team_page_fraction()).
TEAM_MIN_CONTENT_FRACTION = 0.45

# Outlook's OlSensitivity value for a private appointment.
OL_SENSITIVITY_PRIVATE = 2

# Free/busy digit -> (block subject, Outlook busy status, palette colour) for
# a colleague whose calendar can only be seen as free/busy. Busy status drives
# the block style exactly as it does for a real appointment: tentative is
# drawn hollow. Out of office is purple because that is Outlook's own colour
# for it.
TEAM_FREE_BUSY_BLOCKS = {
    "1": ("Tentative", 1, "steel"),
    "2": ("Busy", 2, "steel"),
    "3": ("Out of office", 3, "purple"),
    "4": ("Working elsewhere", 4, "teal"),
}


def team_page_fraction():
    """
    How much of the sheet's HEIGHT the team page on the back may use.

    The sheet prints double-sided and flips on the SHORT edge, so the top of
    the back is the top of the front. When the blank bottom strip of the front
    (1 - PLANNER_CONTENT_FRACTION of the sheet) is folded up behind the page,
    it lies over the back from (2 x fill - 1) down to the fold - 60% to 80% at
    the default 0.8. So the team view is laid out in the top (2 x fill - 1)
    of the back, where the folded strip cannot cover it. With no strip at all
    it takes the whole sheet. TEAM_MIN_CONTENT_FRACTION is the floor: a very
    small page fill would otherwise leave the back nothing at all, and then
    the bottom of the team view is under the strip until it is unfolded.
    """
    if PLANNER_CONTENT_FRACTION >= 1.0:
        return 1.0
    return max(TEAM_MIN_CONTENT_FRACTION, 2.0 * PLANNER_CONTENT_FRACTION - 1.0)


def _person_label(name):
    """
    The name to print over a column, from a directory display name.

    Enterprise directories file people as "Smith, Jane" or "Smith, Jane
    (Branch)"; a column header reads better as "Jane Smith". Anything not in
    that shape is printed as it is.
    """
    name = _one_line(name)
    match = re.fullmatch(r"([^,()]+),\s*([^,()]+?)\s*(\([^()]*\))?", name)
    if match:
        return "{0} {1}".format(match.group(2).strip(), match.group(1).strip())
    return name


def _free_busy_events(free_busy, day):
    """
    Turn one day of a free/busy string into blocks the planner can draw.

    Used for a colleague whose calendar cannot be opened, only seen as
    free/busy: the blocks say when they are busy, tentative, away or working
    elsewhere, with nothing about what. Runs of the same state are merged
    into one block. A run covering the whole day becomes an all-day chip, so
    a day out of office does not stretch the timeline to midnight.
    """
    slots = (24 * 60) // FREEBUSY_MIN_PER_CHAR
    digits = (free_busy or "")[:slots]
    midnight = datetime.datetime.combine(day, datetime.time(0, 0))
    events = []
    index = 0
    while index < len(digits):
        state = digits[index]
        end = index
        while end < len(digits) and digits[end] == state:
            end += 1
        block = TEAM_FREE_BUSY_BLOCKS.get(state)
        if block:
            subject, busy, colour = block
            events.append({
                "start": midnight + datetime.timedelta(
                    minutes=index * FREEBUSY_MIN_PER_CHAR),
                "end": midnight + datetime.timedelta(
                    minutes=end * FREEBUSY_MIN_PER_CHAR),
                "all_day": index == 0 and end >= slots,
                "subject": subject,
                "location": "",
                "organiser": "",
                "categories": [],
                "busy": busy,
                "people": [],
                "external": False,
                "recurring": False,
                "colour": colour,
            })
        index = end
    return events


def _collect_team_events(folder, day, identity):
    """
    One colleague's appointments on `day`, read from their shared calendar.

    The same reading as your own calendar (both recurrence passes), plus
    anything that started up to TEAM_LOOKBACK_DAYS earlier and is still
    running, so a week of leave is on its fourth day too. A timed appointment
    that runs past either end of the day is shown as an all-day chip rather
    than stretching the shared timeline to midnight. A PRIVATE appointment
    prints as "Private appointment" with no location, whatever the access
    allows: it was marked private for a reason, and this sheet gets left on
    a desk. Hidden categories are dropped exactly as on the front. Raises
    when the folder cannot be read, which the caller takes as "fall back to
    free/busy".
    """
    start_dt = datetime.datetime.combine(day, datetime.time(0, 0))
    end_dt = datetime.datetime.combine(day, datetime.time(23, 59, 59))
    day_end = start_dt + datetime.timedelta(days=1)
    matches, _total, _tail = scan_calendar(
        start_dt, end_dt, folder=folder,
        overlap_from=start_dt - datetime.timedelta(days=TEAM_LOOKBACK_DAYS))

    events = []
    for item_start, item in matches:
        event = _appointment_to_event(item, item_start, identity)
        if event is None or event["end"] <= start_dt:
            continue
        if _HIDDEN_CATEGORIES and any(
                category.lower() in _HIDDEN_CATEGORIES
                for category in event["categories"]):
            continue
        try:
            private = int(item.Sensitivity) == OL_SENSITIVITY_PRIVATE
        except Exception:
            private = False
        if private:
            event["subject"] = "Private appointment"
            event["location"] = ""
        if not event["all_day"] and (event["start"] < start_dt
                                     or event["end"] > day_end):
            event["all_day"] = True
        events.append(event)
    events.sort(key=lambda event: (not event["all_day"], event["start"]))
    return events


def _team_column(ns, label, spec, day, identity):
    """
    Everything one column of the team page needs, for one person.

    Tries their calendar first (subjects and locations, when you have been
    given that much access), and falls back to their free/busy when it cannot
    be opened. It ALSO falls back when the calendar opens but shows nothing
    that day while free/busy shows them busy: with free/busy-only permission
    Outlook can hand back an empty folder rather than an error, and an empty
    column for somebody in back-to-back meetings is worse than a column of
    plain "Busy" blocks.

    Returns {"label", "source", "events"}, where source is "calendar",
    "free/busy", "unresolved" (not in the directory) or "unavailable" (neither
    the calendar nor free/busy could be read). Never raises: one person's
    calendar failing must not stop the other three, or the front, printing.
    """
    column = {"label": label or spec, "source": "unavailable", "events": []}
    try:
        recipient = ns.CreateRecipient(spec)
        recipient.Resolve()
        resolved = bool(recipient.Resolved)
    except Exception as exc:
        log("Team calendar: could not resolve '{0}': {1}".format(spec, exc))
        resolved = False
    if not resolved:
        column["source"] = "unresolved"
        return column
    if not label:
        try:
            column["label"] = _person_label(recipient.Name or spec) or spec
        except Exception:
            column["label"] = spec

    events = None
    try:
        folder = ns.GetSharedDefaultFolder(recipient, OL_FOLDER_CALENDAR)
        events = _collect_team_events(folder, day, identity)
    except Exception as exc:
        log("Team calendar: cannot open the calendar of '{0}' ({1}); trying "
            "free/busy.".format(spec, exc))

    if events:
        column["source"], column["events"] = "calendar", events
        return column

    free_busy = _free_busy(recipient, day)
    blocks = _free_busy_events(free_busy, day) if free_busy else None
    if blocks:
        if events is not None:
            log("Team calendar: the calendar of '{0}' opened but showed nothing "
                "on {1}, while free/busy shows booked time - printing "
                "free/busy instead.".format(spec, day))
        column["source"], column["events"] = "free/busy", blocks
    elif events is not None:
        column["source"] = "calendar"       # opened, and genuinely empty
    elif free_busy is not None:
        column["source"] = "free/busy"      # free all day
    return column


def _team_columns(day, identity):
    """Read every configured team calendar into a list of columns."""
    ns = get_namespace()
    return [_team_column(ns, label, spec, day, identity)
            for label, spec in _CALENDAR_TEAM]


def _draw_team_page(canvas, day, team, category_colours):
    """
    Draw the team view: one day column per colleague, side by side.

    The same design as the front - the hour grid, the blocks, the chips and
    the footer are all drawn the same way - over a shared timeline, so a time
    reads straight across every column. The page is laid out in the top
    team_page_fraction() of the sheet, above where the front's folded strip
    lies on the back.

    There is NO page header (v11.2.1): the date is on the front of the same
    sheet, and the room it took goes to the columns instead, which are short
    enough already above the fold. The person's name heading each column is
    the first thing on the page.
    """
    page_w, page_h = canvas.width_mm, canvas.height_mm
    content_h = page_h * team_page_fraction()
    pad_l = pad_r = 11.0
    pad_t, pad_b = 10.0, 8.0

    all_events = []
    for column in team:
        for event in column["events"]:
            if "colour" not in event:
                _label, event["colour"] = _event_bucket(event, category_colours)
            all_events.append(event)
    hours = _planner_hours(all_events, day)

    # ---- page frame ------------------------------------------------------
    # No header: the column headings start at the top margin.
    body_top = pad_t
    footer_rule = content_h - pad_b - 4.6
    body_bottom = footer_rule - 2.5
    gutter = 13.0
    col_gap = 6.0
    cols_x = pad_l + gutter
    cols_w = (page_w - pad_r) - cols_x
    col_w = (cols_w - col_gap * (len(team) - 1)) / float(len(team))
    col_xs = [cols_x + index * (col_w + col_gap) for index in range(len(team))]

    for col_x in col_xs[1:]:
        divider = col_x - col_gap / 2.0
        canvas.line(divider, body_top - 1.0, divider, footer_rule - 2.0,
                    _GRID, 0.6, dash=[2.2, 2.2])

    # ---- column headings: the person's name, and what was readable -------
    for column, col_x in zip(team, col_xs):
        timed = [event for event in column["events"] if not event["all_day"]]
        if column["source"] == "unresolved":
            note = "not found"
        elif column["source"] == "unavailable":
            note = "no access"
        elif column["source"] == "free/busy":
            note = "free/busy only"
        elif timed:
            note = "{0} meeting{1}".format(len(timed),
                                          "" if len(timed) == 1 else "s")
        else:
            note = ""
        note_w = canvas.text(col_x + col_w, body_top + 3.0, note, 7.0,
                             bold=True, colour=_FAINT, tracking=0.4,
                             align="right") if note else 0.0
        canvas.text(col_x, body_top + 3.0, column["label"], 10.5, bold=True,
                    colour=_INK, max_width=col_w - note_w - 2.0)

    cursor_y = body_top + 6.4

    # ---- all-day chips: one row per column, the same height across -------
    chip_h = 5.0
    if any(event["all_day"] for column in team for event in column["events"]):
        for column, col_x in zip(team, col_xs):
            chips = [event for event in column["events"] if event["all_day"]]
            chip_x = col_x
            for position, event in enumerate(chips):
                style = _block_style(event, event["colour"])
                chip_w = min(col_w, canvas.text_width(
                    event["subject"], PLANNER_TITLE_PT, bold=True) + 5.0)
                left = len(chips) - position
                more = "+{0} more".format(left)
                more_w = canvas.text_width(more, PLANNER_TITLE_PT, bold=True)
                # The last chip that fits leaves room for "+n more" when
                # anything is still to come after it.
                room = col_x + col_w - chip_x - (more_w + 2.0 if left > 1 else 0.0)
                if chip_w > room and chip_x > col_x:
                    canvas.text(chip_x + 1.0, cursor_y + chip_h / 2.0
                                + _cap_mm(PLANNER_TITLE_PT) / 2.0, more,
                                PLANNER_TITLE_PT, bold=True, colour=_MUTED)
                    break
                chip_w = min(chip_w, max(room, 8.0))
                canvas.rect(chip_x, cursor_y, chip_w, chip_h, fill=style["fill"],
                            stroke=style["stroke"],
                            line_pt=style["line_pt"] or 0.6, radius=1.2)
                canvas.text(chip_x + 2.5, cursor_y + chip_h / 2.0
                            + _cap_mm(PLANNER_TITLE_PT) / 2.0,
                            event["subject"], PLANNER_TITLE_PT, bold=True,
                            colour=style["title"], max_width=chip_w - 5.0)
                chip_x += chip_w + 1.6
        cursor_y += chip_h + 2.6

    # ---- the shared hour grid ---------------------------------------------
    grid_top = cursor_y + 2.2
    grid_bottom = body_bottom - 2.2
    start_min, end_min = hours[0] * 60, hours[1] * 60
    scale = (grid_bottom - grid_top) / float(max(1, end_min - start_min))
    label_step = 1 if (60.0 * scale) >= 4.0 else 2
    for hour in range(hours[0], hours[1] + 1):
        y = grid_top + (hour * 60 - start_min) * scale
        for col_x in col_xs:
            canvas.line(col_x, y, col_x + col_w, y, _GRID, 0.75, dash=[0.6, 1.8])
        if (hour - hours[0]) % label_step == 0:
            canvas.text(pad_l, y + _cap_mm(7.5) / 2.0, "{0:02d}:00".format(hour),
                        7.5, bold=True, colour=_FAINT)

    for column, col_x in zip(team, col_xs):
        timed = [event for event in column["events"] if not event["all_day"]]
        if timed:
            for box in _place_blocks(timed, day, start_min, end_min,
                                     grid_top, grid_bottom, col_x, col_w):
                box["style"] = _block_style(box["event"], box["event"]["colour"])
                _draw_block(canvas, box)
            continue
        if column["source"] == "unresolved":
            message = "Not found in the directory"
        elif column["source"] == "unavailable":
            message = "Calendar not available"
        elif column["events"]:
            continue                       # all-day only: the chip says it
        else:
            message = "Nothing in the diary"
        canvas.text(col_x + col_w / 2.0, (grid_top + grid_bottom) / 2.0,
                    message, 9.0, bold=True, colour=_GRID, align="centre",
                    max_width=col_w - 2.0)

    # ---- footer ----------------------------------------------------------
    canvas.line(pad_l, footer_rule, page_w - pad_r, footer_rule, _RULE, 0.5)
    canvas.text(pad_l, footer_rule + 3.4,
                "Outlook Calendar · Team view · " + _fmt_day_long(day), 7.0,
                bold=True, colour=_FAINT, max_width=170.0)
    printed = datetime.datetime.now()
    canvas.text(page_w - pad_r, footer_rule + 3.4,
                "Printed {0} at {1}".format(
                    _fmt_day_short(printed.date()), _fmt_time(printed)),
                7.0, bold=True, colour=_FAINT, align="right")

    # The same fold hairline as the front, on the same line of the sheet, so
    # the fold can be made from either side.
    if PLANNER_CONTENT_FRACTION < 1.0:
        fold_y = page_h * PLANNER_CONTENT_FRACTION
        canvas.line(pad_l, fold_y, page_w - pad_r, fold_y, _RULE, 0.4,
                    dash=[1.4, 2.6])


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


def _collect_planner_events(start_date, end_date):
    """
    Read the calendar between two dates into plain dicts, one per occurrence.

    Returns the events by date. EVERY event is laid out on the page,
    blacklisted or not: the PDF is drawn locally and printed on the user's own
    paper, and a planner missing the meetings that matter most is not worth
    carrying. The blacklist still governs what reaches the AI, and it is
    applied to the fields the reply is built from - the appointment's Body is
    never read at all on this path, so nothing in it can leak either way. Each
    event carries "show_to_ai": False means printed, but not named in the reply.

    The one thing that IS left off is an appointment in a category the user
    chose to hide (OUTLOOK_CALENDAR_HIDE_CATEGORIES) - dropped here, before
    layout, so it cannot stretch the timeline or make an empty day look busy.
    Returns (by_date, hidden_by_date), the second counting those per day.
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
    hidden_by_date = {}
    for item_start, item in matches:
        day = item_start.date()
        # Only the printed day needs its recipients read: that is the only
        # panel whose colour depends on who is invited, and reading them is a
        # COM round trip per meeting.
        event = _appointment_to_event(
            item, item_start, identity, want_people=(day == start_date))
        if event is None:
            continue
        # A category the user asked to keep off paper (birthdays, say). Tested
        # before the blacklist so a hidden item is not logged as printed.
        if _HIDDEN_CATEGORIES and any(
                category.lower() in _HIDDEN_CATEGORIES
                for category in event["categories"]):
            hidden_by_date[day] = hidden_by_date.get(day, 0) + 1
            continue
        # The blacklist decides what the AI is told about, not what is printed.
        # Only the fields the reply is built from are tested, because they are
        # the only ones that leave the endpoint.
        reason = appointment_visible_block_reason(item)
        event["show_to_ai"] = reason is None
        if reason:
            log("Printed a calendar item but withheld it from the reply "
                "(blacklist match: {0}).".format(reason))
        by_date.setdefault(day, []).append(event)

    for events in by_date.values():
        events.sort(key=lambda event: (not event["all_day"], event["start"]))
    return by_date, hidden_by_date


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
    skip_empty = bool(args.get("skip_empty_days", True))

    horizon = day + datetime.timedelta(
        days=PLANNER_SCAN_HORIZON_DAYS if skip_empty else wanted)
    try:
        by_date, hidden_by_date = _collect_planner_events(day, horizon)
    except RuntimeError as exc:
        return "Error: {0}.".format(exc)

    day_events = by_date.get(day, [])
    ahead = _planner_ahead_days(by_date, day, wanted, skip_empty)
    hours = _planner_hours(day_events, day)

    # Count the blacklist matches that are actually ON THE SHEET: the scan
    # reaches further ahead than the right-hand panel shows, so the collection
    # holds days the page never gets to.
    withheld = sum(1 for event in
                   day_events + [item for _other, events in ahead
                                 for item in events]
                   if not event.get("show_to_ai", True))

    # Outlook's own category colours first, then the endpoint's explicit map on
    # top: OUTLOOK_CALENDAR_COLOURS is there to overrule a category whose
    # Outlook colour is not the one you want on paper, so it has to win.
    category_colours = outlook_category_colours(get_namespace())
    category_colours.update(_CATEGORY_COLOURS)

    # The team view on the back: the same day, from the calendars in
    # OUTLOOK_CALENDAR_TEAM. Each column falls back on its own, so one
    # colleague who has not shared their calendar costs only their column.
    team = None
    if _CALENDAR_TEAM and args.get("include_team", True) is not False:
        team = _team_columns(day, current_user_identity(get_namespace()))

    try:
        pdf = render_day_planner(
            day, day_events, ahead, current_user_identity(get_namespace()),
            hours, category_colours=category_colours, team=team)
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
    listed = 0
    for event in day_events:
        # Everything above is a count, and counts were always reported. The
        # per-event lines are the only place a subject is spelled out, so a
        # blacklisted one stops here - it is on the printed page, not in this
        # reply.
        if not event.get("show_to_ai", True):
            continue
        listed += 1
        when = ("all day" if event["all_day"]
                else "{0}-{1}".format(_fmt_time(event["start"]),
                                      _fmt_time(event["end"])))
        lines.append("  - {0}  {1}{2}".format(
            when, event["subject"],
            " ({0})".format(event["location"]) if event["location"] else ""))
    if not day_events:
        lines.append("  (nothing in the diary - the page prints empty)")
    elif not listed:
        lines.append("  (every event that day is withheld from this summary by "
                     "the content blacklist - all of them are on the page)")

    lines.append("")
    lines.append("Following days on the page: " + (", ".join(
        "{0} ({1} item(s))".format(_fmt_day_short(other), len(events))
        for other, events in ahead) or "none"))
    # Only the days actually on the sheet: the scan reaches further ahead than
    # the right-hand panel shows.
    hidden = sum(hidden_by_date.get(when, 0)
                 for when in [day] + [other for other, _events in ahead])
    if hidden:
        lines.append("{0} event(s) in a hidden category (OUTLOOK_CALENDAR_HIDE_"
                     "CATEGORIES) were left off the page.".format(hidden))
    if withheld:
        lines.append("[{0} event(s) matched the content blacklist. They ARE on "
                     "the printed page - only their subjects are withheld from "
                     "this summary.]".format(withheld))

    # The team view is reported as counts per person and never by subject:
    # these are other people's diaries, and what is in them is for the sheet
    # on the desk, not for this reply.
    if team:
        lines.append("")
        lines.append("Page 2 (the back): team view for the same day.")
        for column in team:
            timed = [event for event in column["events"] if not event["all_day"]]
            all_day = len(column["events"]) - len(timed)
            if column["source"] == "unresolved":
                detail = ("NOT FOUND in the directory - check the entry in "
                          "OUTLOOK_CALENDAR_TEAM; the column prints empty")
            elif column["source"] == "unavailable":
                detail = ("neither the calendar nor free/busy could be read - "
                          "the column prints 'Calendar not available'")
            elif column["source"] == "free/busy":
                detail = ("free/busy only (no access to the calendar itself): "
                          "{0} busy block(s), {1} all-day".format(
                              len(timed), all_day))
            else:
                detail = "{0} meeting(s), {1} all-day".format(len(timed), all_day)
            lines.append("  - {0}: {1}".format(column["label"], detail))
        lines.append(
            "PRINT IT DOUBLE-SIDED, FLIP ON SHORT EDGE. The PDF asks the print "
            "dialog for that and Adobe Reader honours it; in any other viewer "
            "check the setting before printing, or the back comes out on a "
            "second sheet or upside down.")
        if PLANNER_CONTENT_FRACTION < 1.0:
            covered = 2.0 * PLANNER_CONTENT_FRACTION - 1.0 < team_page_fraction()
            lines.append(
                "The team view sits in the top {0:.0f}% of the back{1}.".format(
                    team_page_fraction() * 100,
                    " - part of it is under the folded strip at this page "
                    "fill, so unfold the strip to read the bottom of it"
                    if covered else
                    ", clear of the bottom strip once that is folded up behind"))
    elif _CALENDAR_TEAM:
        lines.append("")
        lines.append("Team view left off (include_team=false): one page only.")
    elif args.get("include_team") is True:
        lines.append("")
        lines.append("No team view: OUTLOOK_CALENDAR_TEAM is not set, so there "
                     "is no back page. Set it to up to four email addresses, "
                     "semicolon-separated, and restart Claude Code.")

    # The folding advice follows the configured page fill, so it cannot tell
    # the user to fold a strip that this endpoint does not leave blank.
    if team:
        if PLANNER_CONTENT_FRACTION >= 1.0:
            lines.append("It is A4 landscape - fold it in half for a bifold "
                         "day planner.")
        else:
            lines.append(
                "It is A4 landscape, printed on the top {0:.0f}% of the front - "
                "fold it in half for a bifold day planner, and fold the blank "
                "bottom {1:.0f}% up behind it.".format(
                    PLANNER_CONTENT_FRACTION * 100,
                    (1.0 - PLANNER_CONTENT_FRACTION) * 100))
    elif PLANNER_CONTENT_FRACTION >= 1.0:
        lines.append("It is A4 landscape - print it single-sided and fold it "
                     "in half for a bifold day planner.")
    else:
        lines.append(
            "It is A4 landscape, printed on the top {0:.0f}% of the sheet - "
            "print it single-sided, fold it in half for a bifold day planner, "
            "and fold the blank bottom {1:.0f}% up behind it.".format(
                PLANNER_CONTENT_FRACTION * 100,
                (1.0 - PLANNER_CONTENT_FRACTION) * 100))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Printable team task sheet (outlook_get_task_team / outlook_print_team_tasks)
# ---------------------------------------------------------------------------
#
# A double-sided A4 landscape sheet of the user's direct reports' open tasks:
# eight columns, four a side. Column one is everything nobody in particular
# owns (an email to the whole team, a Confluence row with no assignee); the
# other seven are the people in OUTLOOK_TASKS_TEAM, in that order.
#
# This server cannot read Confluence, and should not try to word a task from
# email prose with regular expressions. So the work is split.
# outlook_get_task_team hands the model the roster and the Confluence page;
# outlook_list_task_emails finds the emailed tasks by their follow-up flag
# (the user's own decision, so no guessing there); the model reads the page
# (through the confluence plugin), turns both into one list of
# {person, task, due, ...}; and
# outlook_print_team_tasks lays that list out. Everything about the PAGE - the
# columns, their order, sorting, overdue colouring, overflow - is decided here,
# so the sheet comes out the same way every time.
#
# The page reuses the day planner's PdfCanvas, header and footer, so the two
# sheets look like a set.

TASKS_PEOPLE_MAX = 7           # seven people + the unassigned column = 8
TASKS_COLUMNS_PER_PAGE = 4     # four columns a side of A4 landscape
TASKS_MAX_ITEMS = 400          # guard rail on one call's task list
TASKS_MAX_TEXT = 400           # characters kept of one task's text
TASKS_TEXT_PT = 7.5            # the task itself
TASKS_META_PT = 6.0            # the due / status / source line under it
TASKS_TEXT_LINES = 3           # a task wraps to at most this many lines
TASKS_UNASSIGNED_LABEL = "Team / unassigned"

# What a 'person' value may say to mean "nobody in particular". Anything else
# that matches no column also lands in column one, but is reported back as
# unmatched, because it is far more likely a misspelt name than a team task.
TASKS_UNASSIGNED_WORDS = frozenset((
    "", "team", "the team", "whole team", "all", "everyone", "everybody",
    "unassigned", "nobody", "none", "n/a", "-", "tbc", "tbd",
))

# Accent bar under each column heading. Column one is slate so the
# "nobody owns this" column reads as different from the people.
TASKS_ACCENTS = ("slate", "blue", "purple", "green", "amber", "teal", "rose",
                 "dark-orange")


def _tasks_tokens(value):
    """Lower-case name tokens: 'Smith, Jane (Branch)' -> {'smith', 'jane', 'branch'}."""
    return set(re.findall(r"[a-z0-9']+", _one_line(value).lower()))


def _tasks_column_label(label, spec):
    """
    The name printed over a person's column.

    The configured label wins. Without one, a directory name is turned round
    ('Smith, Jane' -> 'Jane Smith'), and an email address becomes its local
    part in title case ('jane.smith@...' -> 'Jane Smith'): there is no
    directory lookup here, so the sheet prints without Outlook's help.
    """
    if label:
        return _one_line(label)
    spec = _one_line(spec)
    if "@" in spec:
        local = spec.split("@", 1)[0]
        return " ".join(part.capitalize()
                        for part in re.split(r"[._-]+", local) if part)
    return _person_label(spec)


def _tasks_person_column(person, columns, loose=False):
    """
    Index into `columns` (people only) that `person` names, or None.

    Matched, in order: the label or configured entry exactly (case-insensitive),
    the email address or its local part, and then every word of `person` being
    a word of exactly ONE column's name - so 'Jane', 'Smith' and 'Smith, Jane'
    all find 'Jane Smith', while a first name two people share finds neither.

    `loose` also accepts the reverse - every word of a column's name being in
    `person` - for columns made from the tasks themselves, so 'Jane Smith'
    joins a column already started as 'Jane' rather than starting another.
    """
    wanted = _one_line(person).lower()
    if not wanted:
        return None
    for index, column in enumerate(columns):
        if wanted in column["keys"]:
            return index
    tokens = _tasks_tokens(wanted)
    if not tokens:
        return None
    hits = [index for index, column in enumerate(columns)
            if tokens <= column["tokens"]]
    if not hits and loose:
        hits = [index for index, column in enumerate(columns)
                if column["tokens"] and column["tokens"] <= tokens]
    return hits[0] if len(hits) == 1 else None


def _tasks_new_column(label, spec=""):
    """One person's column, with every key a task's 'person' may match it by."""
    printed = _tasks_column_label(label, spec)
    keys = set(value for value in (_one_line(label).lower(),
                                   _one_line(spec).lower(),
                                   printed.lower()) if value)
    tokens = _tasks_tokens(printed) | _tasks_tokens(label)
    if "@" in spec:
        local = spec.split("@", 1)[0].lower()
        keys.add(local)
        tokens |= _tasks_tokens(local.replace(".", " "))
    else:
        tokens |= _tasks_tokens(spec)
    return {"label": printed, "spec": _one_line(spec), "keys": keys,
            "tokens": tokens, "items": [], "unassigned": False}


def _parse_task_due(value):
    """
    A task's due date as a date, or None when it is not one.

    YYYY-MM-DD is what the tool asks for; D/M/YYYY is accepted as well because
    that is how an Australian Confluence page writes a date. Anything else
    ('end of month', 'ASAP') is not an error - it is printed as written.
    """
    value = _one_line(value)
    match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", value)
    if match:
        parts = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    else:
        match = re.fullmatch(r"(\d{1,2})[/.](\d{1,2})[/.](\d{4})", value)
        if not match:
            return None
        parts = (int(match.group(3)), int(match.group(2)), int(match.group(1)))
    try:
        return datetime.date(*parts)
    except ValueError:
        return None


def _clean_tasks(raw):
    """
    Validate the tool's 'tasks' argument into a list of plain dicts.

    Raises ValueError with a message the model can act on. Every string is
    collapsed to one line, so nothing in a task can break the page layout.
    """
    if not isinstance(raw, list):
        raise ValueError("'tasks' must be a list of objects, each with at "
                         "least a 'task' string")
    if len(raw) > TASKS_MAX_ITEMS:
        raise ValueError("{0} tasks is more than one sheet can hold (limit "
                         "{1}) - leave out what is done or not worth printing"
                         .format(len(raw), TASKS_MAX_ITEMS))
    tasks = []
    for position, item in enumerate(raw, start=1):
        if not isinstance(item, dict):
            raise ValueError("task {0} is not an object".format(position))
        text = _one_line(item.get("task"))
        if not text:
            raise ValueError("task {0} has no 'task' text".format(position))
        due_raw = _one_line(item.get("due"))
        tasks.append({
            "order": position,
            "task": text[:TASKS_MAX_TEXT],
            "person": _one_line(item.get("person"))[:120],
            "due_raw": due_raw[:40],
            "due": _parse_task_due(due_raw),
            "status": _one_line(item.get("status"))[:40],
            "source": _one_line(item.get("source"))[:60],
            "detail": _one_line(item.get("detail"))[:160],
            "for": "",
        })
    return tasks


def _build_task_columns(tasks, roster):
    """
    Sort the tasks into columns: unassigned first, then one per person.

    With a roster (OUTLOOK_TASKS_TEAM) the people and their order are fixed,
    and a name that matches nobody goes into column one marked "For <name>"
    rather than being dropped. Without one, columns are made from the names on
    the tasks in the order they first appear, up to TASKS_PEOPLE_MAX; past
    that, the rest go into column one the same way.

    Returns (columns, unmatched) where unmatched is {name: count}.
    """
    unassigned = {"label": TASKS_UNASSIGNED_LABEL, "spec": "", "keys": set(),
                  "tokens": set(), "items": [], "unassigned": True}
    people = [_tasks_new_column(label, spec) for label, spec in roster]
    unmatched = {}

    for task in tasks:
        person = task["person"]
        if person.lower() in TASKS_UNASSIGNED_WORDS:
            unassigned["items"].append(task)
            continue
        index = _tasks_person_column(person, people, loose=not roster)
        if index is None and not roster and len(people) < TASKS_PEOPLE_MAX:
            # As spec, not label, so an email address or 'Smith, Jane' is
            # turned into a printable name the way a roster entry would be.
            people.append(_tasks_new_column("", person))
            index = len(people) - 1
        if index is None:
            task["for"] = person
            unassigned["items"].append(task)
            unmatched[person] = unmatched.get(person, 0) + 1
        else:
            people[index]["items"].append(task)

    today = datetime.date.today()
    columns = [unassigned] + people
    for column in columns:
        # Dated tasks first, soonest (so overdue) at the top; undated after,
        # in the order given - which is the page's or the email's own order.
        column["items"].sort(key=lambda task: (
            0 if task["due"] else 1,
            task["due"] or today,
            task["order"]))
    return columns, unmatched


def _task_meta(task, today):
    """
    The small line under a task: (due text, due colour, rest of the line).

    The due date is drawn separately, bold and coloured, because it is what
    the eye runs down the column for: red once overdue, amber when due today.
    """
    due_text, due_colour = "", _MUTED
    if task["due"]:
        if task["due"] < today:
            due_text = "Overdue " + _fmt_day_short(task["due"])
            due_colour = PLANNER_COLOURS["dark-red"]
        elif task["due"] == today:
            due_text = "Due today"
            due_colour = PLANNER_COLOURS["dark-yellow"]
        else:
            due_text = "Due " + _fmt_day_short(task["due"])
    elif task["due_raw"]:
        due_text = "Due " + task["due_raw"]

    rest = []
    if task["for"]:
        rest.append("For " + task["for"])
    for key in ("status", "source", "detail"):
        if task[key]:
            rest.append(task[key])
    return due_text, due_colour, " · ".join(rest)


def _measure_task(canvas, task, width, today):
    """(wrapped lines, height in mm) of one task at the given text width."""
    lines = canvas.wrap(task["task"], TASKS_TEXT_PT, width,
                        max_lines=TASKS_TEXT_LINES)
    due_text, _colour, rest = _task_meta(task, today)
    height = 1.2 + len(lines) * 3.2
    if due_text or rest:
        height += 2.7
    return lines, height + 2.2


def _draw_task_header(canvas, identity, stats, page_no, pages,
                      page_w, pad_l, pad_r, pad_t):
    """The day planner's header, with the printed date as its title."""
    name, _address = identity
    eyebrow = "TEAM TASKS" + (" · " + name if name else "")
    if pages > 1:
        eyebrow += " · {0} OF {1}".format(page_no, pages)
    canvas.text(pad_l, pad_t + 2.0, eyebrow, 7.5, bold=True, colour=_MUTED,
                tracking=0.85, max_width=140.0)
    date_w = canvas.text(pad_l, pad_t + 8.6,
                         _fmt_day_long(datetime.date.today()), 19.0,
                         bold=True, colour=_INK)
    canvas.text(page_w - pad_r, pad_t + 8.0, stats, 7.5, bold=True,
                colour=_FAINT, align="right",
                max_width=page_w - pad_l - pad_r - date_w - 8.0)
    header_rule = pad_t + 12.8
    canvas.line(pad_l, header_rule, page_w - pad_r, header_rule, _INK, 1.2)
    return header_rule


def _draw_task_footer(canvas, identity, sources, page_no, pages,
                      page_w, page_h, pad_l, pad_r, pad_b):
    """The day planner's footer: what this is on the left, when it printed on the right."""
    _name, address = identity
    footer_rule = page_h - pad_b - 4.6
    canvas.line(pad_l, footer_rule, page_w - pad_r, footer_rule, _RULE, 0.5)
    left = "Team Tasks" + (" · " + sources if sources else "")
    if address:
        left += " · " + address
    printed = datetime.datetime.now()
    right = "Printed {0} at {1}".format(_fmt_day_short(printed.date()),
                                        _fmt_time(printed))
    if pages > 1:
        right = "Page {0} of {1} · {2}".format(page_no, pages, right)
    right_w = canvas.text(page_w - pad_r, footer_rule + 3.4, right, 7.0,
                          bold=True, colour=_FAINT, align="right")
    canvas.text(pad_l, footer_rule + 3.4, left, 7.0, bold=True, colour=_FAINT,
                max_width=page_w - pad_l - pad_r - right_w - 8.0)


def _draw_task_column(canvas, column, accent_key, col_x, col_w, top, bottom,
                      today):
    """
    One column: the person's name, an accent bar, then their tasks.

    Each task has an empty tick box, because this is a sheet that gets ticked
    off with a pen. A column with more than fits ends in '+N more' rather than
    shrinking its type; the count is returned so the tool can say so.
    """
    items = column["items"]
    note = "{0} open".format(len(items)) if items else ""
    note_w = canvas.text(col_x + col_w, top + 3.0, note, 7.0, bold=True,
                         colour=_FAINT, tracking=0.4,
                         align="right") if note else 0.0
    canvas.text(col_x, top + 3.0, column["label"], 10.5, bold=True,
                colour=_INK, max_width=col_w - note_w - 2.0)
    canvas.rect(col_x, top + 5.2, col_w, 1.0,
                fill=planner_palette(accent_key)["fill"], radius=0.5)

    y = top + 9.0
    if not items:
        canvas.text(col_x + col_w / 2.0, (y + bottom) / 2.0, "Nothing open",
                    9.0, bold=True, colour=_GRID, align="centre",
                    max_width=col_w - 2.0)
        return 0

    box = 2.6
    text_x = col_x + box + 1.8
    text_w = col_w - box - 1.8
    more_h = 4.0
    for index, task in enumerate(items):
        lines, height = _measure_task(canvas, task, text_w, today)
        remaining = len(items) - index - 1
        if y + height + (more_h if remaining else 0.0) > bottom:
            canvas.text(text_x, y + 3.0,
                        "+{0} more not shown".format(len(items) - index),
                        7.0, bold=True, colour=_MUTED)
            return len(items) - index

        baseline = y + 1.2 + _cap_mm(TASKS_TEXT_PT)
        canvas.rect(col_x, baseline - _cap_mm(TASKS_TEXT_PT) / 2.0 - box / 2.0,
                    box, box, stroke=_MUTED, line_pt=0.6, radius=0.4)
        for line in lines:
            canvas.text(text_x, baseline, line, TASKS_TEXT_PT, colour=_INK)
            baseline += 3.2
        due_text, due_colour, rest = _task_meta(task, today)
        if due_text or rest:
            meta_y = baseline - 3.2 + 2.7
            drawn = 0.0
            if due_text:
                drawn = canvas.text(text_x, meta_y, due_text, TASKS_META_PT,
                                    bold=True, colour=due_colour,
                                    max_width=text_w) + 1.6
            if rest and text_w - drawn > 6.0:
                canvas.text(text_x + drawn, meta_y,
                            ("· " if due_text else "") + rest, TASKS_META_PT,
                            colour=_MUTED, max_width=text_w - drawn)
        y += height
        if remaining:
            canvas.line(text_x, y - 1.1, col_x + col_w, y - 1.1, _HAIRLINE, 0.5)
    return 0


def render_team_tasks(columns, identity, sources=""):
    """
    Build the task sheet and return (PDF bytes, {column label: count not shown}).

    Four columns a side of A4 landscape, column one being the unassigned
    tasks. Eight columns make two pages, and the PDF asks the print dialog
    for double-sided, flip on short edge - the same preset as the day
    planner's team view - so they land back to back on one sheet. The whole
    sheet is used: there is no blank strip to fold, because a task list is
    filed or pinned up rather than kept in a diary.
    """
    page_w, page_h = 297.0, 210.0
    pad_l = pad_r = 11.0
    pad_t, pad_b = 10.0, 8.0
    name, address = identity
    today = datetime.date.today()

    canvas = PdfCanvas(page_w, page_h,
                       title="Team tasks - {0}".format(_fmt_day_long(today)),
                       author=name or address or "")

    total = sum(len(column["items"]) for column in columns)
    overdue = sum(1 for column in columns for task in column["items"]
                  if task["due"] and task["due"] < today)
    stats = "{0} open task{1}".format(total, "" if total == 1 else "s")
    if overdue:
        stats += " · {0} overdue".format(overdue)
    if columns[0]["items"]:
        stats += " · {0} unassigned".format(len(columns[0]["items"]))
    if not total:
        stats = "Nothing open"

    pages = max(1, -(-len(columns) // TASKS_COLUMNS_PER_PAGE))
    col_gap = 6.0
    col_w = ((page_w - pad_l - pad_r - col_gap * (TASKS_COLUMNS_PER_PAGE - 1))
             / float(TASKS_COLUMNS_PER_PAGE))
    hidden = {}

    for page_index in range(pages):
        if page_index:
            canvas.new_page()
        header_rule = _draw_task_header(canvas, identity, stats,
                                        page_index + 1, pages, page_w,
                                        pad_l, pad_r, pad_t)
        body_top = header_rule + 5.0
        footer_rule = page_h - pad_b - 4.6
        body_bottom = footer_rule - 2.5

        first = page_index * TASKS_COLUMNS_PER_PAGE
        on_page = columns[first:first + TASKS_COLUMNS_PER_PAGE]
        for slot, column in enumerate(on_page):
            col_x = pad_l + slot * (col_w + col_gap)
            if slot:
                divider = col_x - col_gap / 2.0
                canvas.line(divider, body_top - 1.0, divider, footer_rule - 2.0,
                            _GRID, 0.6, dash=[2.2, 2.2])
            accent = TASKS_ACCENTS[(first + slot) % len(TASKS_ACCENTS)]
            not_shown = _draw_task_column(canvas, column, accent, col_x, col_w,
                                          body_top, body_bottom, today)
            if not_shown:
                hidden[column["label"]] = not_shown

        _draw_task_footer(canvas, identity, sources, page_index + 1, pages,
                          page_w, page_h, pad_l, pad_r, pad_b)

    if pages > 1:
        canvas.duplex_short_edge = True
    return canvas.to_bytes(), hidden


def tool_get_task_team(args):
    """
    Everything the model needs before it gathers tasks: the columns in print
    order, the Confluence page to read, and where the emailed tasks come from,
    with today's date from the endpoint's own clock.
    """
    today = datetime.date.today()
    lines = ["Team task sheet - set-up on this endpoint", ""]
    lines.append("Columns, in print order (four a side, column 1 first):")
    lines.append("  1. {0} - tasks for the whole team, for several people, or "
                 "with nobody assigned (person: \"\")".format(
                     TASKS_UNASSIGNED_LABEL))
    if _TASKS_TEAM:
        for number, (label, spec) in enumerate(_TASKS_TEAM, start=2):
            printed = _tasks_column_label(label, spec)
            lines.append("  {0}. {1}{2}".format(
                number, printed,
                "" if spec.lower() == printed.lower()
                else " <{0}>".format(spec)))
        lines.append("Use these names as each task's 'person'. A first name or "
                     "surname alone also matches when only one person has it.")
    else:
        lines.append("  (OUTLOOK_TASKS_TEAM is not set, so the remaining columns "
                     "are made from the names on the tasks, in the order they "
                     "first appear, up to {0} people. Ask the user who their "
                     "direct reports are, and suggest setting "
                     "OUTLOOK_TASKS_TEAM for a fixed order.)".format(
                         TASKS_PEOPLE_MAX))
    lines.append("")
    if _TASKS_PAGE:
        lines.append("Confluence task page: {0}  (OUTLOOK_TASKS_PAGE)".format(
            _TASKS_PAGE))
    else:
        lines.append("Confluence task page: not set (OUTLOOK_TASKS_PAGE) - ask "
                     "the user which page collates the team's tasks.")
    lines.append("Emailed tasks: outlook_list_task_emails - emails YOU sent "
                 "that are in your Inbox (your rule copies them back) and that "
                 "you flagged with a due date. Unflagged, completed and "
                 "undated ones are not tasks. Do not read Sent Items.")
    lines.append("Today on this endpoint: {0}.".format(_fmt_day_long(today)))
    if not PDF_DIR:
        lines.append("")
        lines.append("PRINTING IS OFF on this endpoint (OUTLOOK_DOCS_DIR=off): "
                     "outlook_print_team_tasks will refuse. Say so before "
                     "gathering anything.")
    return "\n".join(lines)


# The emailed half of the task sheet. The user's Outlook rule copies every email
# they send back into the Inbox; there they either DELETE it (no reminder
# needed) or FLAG it with a due date (a task they want chased). So the emailed
# tasks are exactly the Inbox items that the user sent, that are flagged and
# not complete, and that carry a real due date - no date window, no guessing
# from the wording, and marking the flag complete in Outlook takes the task off
# the next sheet.
#
# MAPI PR_FLAG_STATUS (0x1090, a long): 2 = flagged, 1 = complete. A DASL
# filter on an integer formats no date, so it is safe from the regional-
# settings fault that keeps Restrict() off dates elsewhere in this file. If
# Outlook rejects it anyway, the Inbox is walked and FlagStatus checked in
# Python instead, up to TASK_EMAIL_SCAN_CAP items.
TASK_EMAIL_FLAG_FILTER = (
    '@SQL="http://schemas.microsoft.com/mapi/proptag/0x10900003" = 2')
OL_FLAG_MARKED = 2
TASK_EMAIL_SCAN_CAP = 5000
TASK_EMAIL_PREVIEW_CHARS = 400
# Outlook stores "no date" as 1 January 4501; anything from 4500 on is none.
OUTLOOK_NO_DATE_YEAR = 4500


def _sent_by_me(item, identity):
    """
    Whether `item` was sent by the mailbox owner.

    The SMTP address is compared first (an Exchange sender is resolved to its
    SMTP form by sender_smtp), and the display name only when no address is
    known, because two people can share a name but not an address.
    """
    name, address = identity
    if address:
        return sender_smtp(item).strip().lower() == address.strip().lower()
    if name:
        try:
            return _one_line(item.SenderName or "").lower() == name.lower()
        except Exception:
            return False
    return False


def tool_list_task_emails(args):
    ns = get_namespace()
    identity = current_user_identity(ns)
    if not identity[0] and not identity[1]:
        return ("Error: could not work out the mailbox owner's name or address, "
                "so emails you sent cannot be told apart from the rest of the "
                "Inbox. Run python outlook.py --check on the endpoint.")

    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    filtered = True
    try:
        candidates = inbox.Items.Restrict(TASK_EMAIL_FLAG_FILTER)
    except Exception as exc:
        log("Flag filter rejected ({0}); walking the Inbox instead.".format(exc))
        candidates = inbox.Items
        filtered = False

    found = []
    no_date = 0
    not_mine = 0
    withheld = 0
    scanned = 0
    capped = False
    for item in candidates:
        if scanned >= TASK_EMAIL_SCAN_CAP:
            capped = True
            break
        scanned += 1
        try:
            if item.Class != OL_CLASS_MAIL:
                continue
            if int(item.FlagStatus) != OL_FLAG_MARKED:
                continue                     # unflagged, or flag completed
        except Exception:
            continue
        if not _sent_by_me(item, identity):
            not_mine += 1                    # someone else's email you flagged
            continue
        try:
            due = _com_to_naive(item.TaskDueDate)
        except Exception:
            due = None
        if due is None or due.year >= OUTLOOK_NO_DATE_YEAR:
            no_date += 1
            continue
        reason = email_block_reason(item)
        if reason:
            withheld += 1
            log("Withheld a flagged email (blacklist match: {0}).".format(reason))
            continue
        try:
            preview = _one_line(item.Body or "")[:TASK_EMAIL_PREVIEW_CHARS]
        except Exception:
            preview = ""
        try:
            found.append({
                "due": due.date(),
                "sent": _com_to_naive(item.SentOn),
                "to": _one_line(item.To or ""),
                "cc": _one_line(item.CC or ""),
                "subject": _one_line(item.Subject or "") or "(no subject)",
                "preview": preview,
                "entry_id": item.EntryID,
            })
        except Exception:
            continue

    found.sort(key=lambda entry: (entry["due"], entry["sent"]
                                  or datetime.datetime.min))
    today = datetime.date.today()
    lines = ["Flagged emails you sent, in your Inbox, with a due date: {0} "
             "(soonest due first; today is {1}).".format(
                 len(found), today.isoformat())]
    for entry in found:
        lines.append("- Due {0}{1} | Sent {2} | To: {3}{4}".format(
            entry["due"].isoformat(),
            " (OVERDUE)" if entry["due"] < today else "",
            fmt_dt(entry["sent"]) if entry["sent"] else "?",
            entry["to"] or "(no recipient)",
            " | Cc: " + entry["cc"] if entry["cc"] else ""))
        lines.append("    Subject : {0}".format(entry["subject"]))
        if entry["preview"]:
            lines.append("    Preview : {0}".format(entry["preview"]))
        lines.append("    EntryID : {0}".format(entry["entry_id"]))
    notes = []
    if no_date:
        notes.append("{0} flagged email(s) you sent have no due date and are "
                     "not tasks for the sheet".format(no_date))
    if not_mine:
        notes.append("{0} flagged email(s) were sent by someone else and are "
                     "left out".format(not_mine))
    if withheld:
        notes.append("{0} withheld by the content blacklist".format(withheld))
    if capped:
        notes.append("the scan stopped after {0} items{1}, so older flagged "
                     "emails may be missing".format(
                         TASK_EMAIL_SCAN_CAP,
                         "" if filtered else " (Outlook rejected the flag "
                         "filter, so the whole Inbox was walked)"))
    if notes:
        lines.append("")
        lines.append("[" + "; ".join(notes) + ".]")
    return "\n".join(lines)


def tool_print_team_tasks(args):
    if not PDF_DIR:
        return ("Error: printing is switched off on this server "
                "(OUTLOOK_DOCS_DIR is set to off), so no PDF was written. "
                "Give OUTLOOK_DOCS_DIR a real folder, or unset it to use "
                "%EVA_DOCUMENTS_DIR%\\{0}.".format(DOCS_SUBFOLDER))
    try:
        tasks = _clean_tasks(args.get("tasks"))
    except ValueError as exc:
        return "Error: {0}. Nothing was written.".format(exc)

    columns, unmatched = _build_task_columns(tasks, _TASKS_TEAM)
    sources = _one_line(args.get("sources"))[:160]
    try:
        identity = current_user_identity(get_namespace())
    except Exception:
        # The header and footer name the mailbox owner as a nicety; a sheet
        # without a name on it still prints.
        identity = ("", "")

    try:
        pdf, hidden = render_team_tasks(columns, identity, sources)
    except Exception as exc:
        log("Task sheet render failed:\n{0}".format(traceback.format_exc()))
        return "Error: the task sheet could not be laid out ({0}).".format(exc)

    today = datetime.date.today()
    requested = _one_line(args.get("filename") or "")
    if requested:
        stem = safe_filename(os.path.splitext(os.path.basename(requested))[0])
    else:
        stem = "Team Tasks - {0}".format(today.isoformat())
    path = os.path.join(PDF_DIR, stem + ".pdf")

    # Same rule as the day planner: re-printing today's sheet replaces it, but
    # a file the call NAMED might be one of the user's own PDFs.
    if requested and os.path.exists(path):
        return ("Error: {0} already exists and this call named it explicitly, "
                "so nothing was written. Choose another name, or drop "
                "'filename' to write today's sheet, which does overwrite its "
                "own previous copy.".format(path))
    try:
        os.makedirs(PDF_DIR, exist_ok=True)
        with open(path, "wb") as handle:
            handle.write(pdf)
    except OSError as exc:
        return ("Error: could not write the task sheet to {0} ({1}). The "
                "folder comes from OUTLOOK_DOCS_DIR, or %EVA_DOCUMENTS_DIR%\\{2}."
                .format(path, exc, DOCS_SUBFOLDER))
    log("Wrote team task sheet: {0}".format(path))

    pages = max(1, -(-len(columns) // TASKS_COLUMNS_PER_PAGE))
    overdue = sum(1 for task in tasks if task["due"] and task["due"] < today)
    lines = [
        "Team task sheet written to:",
        "  {0}".format(path),
        "",
        "{0} task(s), {1} overdue, over {2} column(s) on {3} page(s):".format(
            len(tasks), overdue, len(columns), pages),
    ]
    for number, column in enumerate(columns, start=1):
        side = ("front" if (number - 1) // TASKS_COLUMNS_PER_PAGE == 0
                else "back")
        detail = "{0} task(s)".format(len(column["items"]))
        if column["label"] in hidden:
            detail += (", {0} DID NOT FIT and print as '+{0} more not shown'"
                       .format(hidden[column["label"]]))
        lines.append("  {0}. {1} ({2}): {3}".format(
            number, column["label"], side, detail))

    if unmatched:
        lines.append("")
        lines.append("These names matched no column, so their tasks are in "
                     "column 1 marked 'For <name>': " + ", ".join(
                         "{0} ({1})".format(person, count)
                         for person, count in sorted(unmatched.items())))
        if _TASKS_TEAM:
            lines.append("If one is a direct report, re-print with the name "
                         "as it appears in the column list.")
    if hidden:
        lines.append("")
        lines.append("Some columns ran out of room. Shorten those tasks or "
                     "leave out ones that are done, then print again.")
    lines.append("")
    if pages > 1:
        lines.append(
            "PRINT IT DOUBLE-SIDED, FLIP ON SHORT EDGE, at 100% on A4 "
            "landscape: columns 1-4 on the front, 5-{0} on the back. The PDF "
            "asks the print dialog for that and Adobe Reader honours it; in "
            "any other viewer check the setting before printing.".format(
                len(columns)))
    else:
        lines.append("It is one A4 landscape page - print it single-sided "
                     "at 100%.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Meeting scheduling - part 1: finding people and rooms in the directory
# ---------------------------------------------------------------------------
#
# Everything below serves one job: "find a suitable time with Josh Smith in
# Room R5-3-84, then save it as a draft I can send". That needs three things
# Outlook can do but MAPI makes awkward - resolve a half-remembered name
# against an enterprise directory, read everyone's free/busy, and save a
# meeting WITHOUT sending it.

OL_FOLDER_CONTACTS = 10
OL_CLASS_CONTACT = 40       # olContact
OL_APPOINTMENT_ITEM = 1     # CreateItem(olAppointmentItem)
OL_MEETING = 1              # AppointmentItem.MeetingStatus = olMeeting

# Recipient.Type on an AppointmentItem.
OL_RECIP_REQUIRED = 1
OL_RECIP_OPTIONAL = 2
OL_RECIP_RESOURCE = 3

# AppointmentItem.BusyStatus, by the word a user would say.
BUSY_STATUS_NAMES = {
    "free": 0,
    "tentative": 1,
    "busy": 2,
    "oof": 3,
    "out of office": 3,
    "outofoffice": 3,
    "elsewhere": 4,
    "working elsewhere": 4,
}

# PR_SMTP_ADDRESS (Unicode) - the real mailbox address behind a directory
# entry, whose AddressEntry.Address is an X.500 legacyExchangeDN instead.
PROP_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

# PR_DISPLAY_TYPE_EX - the only property that tells a ROOM apart from a
# person. AddressEntry.DisplayType has no room in its enumeration at all, so
# without this a boardroom looks exactly like a colleague.
PROP_DISPLAY_TYPE_EX = "http://schemas.microsoft.com/mapi/proptag/0x39050003"
DT_MAILUSER = 0
DT_DISTLIST = 1
DT_PRIVATE_DISTLIST = 5
DT_REMOTE_MAILUSER = 6
DT_ROOM = 7
DT_EQUIPMENT = 8
DT_SEC_DISTLIST = 9

# Address lists whose name contains one of these hold rooms and equipment.
# Exchange creates "All Rooms"; the rest cover the names an admin may have
# used instead.
ROOM_LIST_HINTS = ("room", "resource", "equipment", "facilit")

# A directory entry whose name reads like a room, used only when
# PR_DISPLAY_TYPE_EX cannot be read (a cached-mode OAB sometimes omits it).
ROOM_NAME_HINTS = ("room", "rm ", "boardroom", "meeting", "training",
                   "conference", "conf ", "suite", "theatre", "theater")

# A directory match has to reach this to be shown at all. Low enough for a
# typo, high enough that "Josh Smith" does not return every Smith.
MIN_NAME_SCORE = 0.62

# Contact folders are small, but a mounted PST can hold a silly number.
CONTACT_SCAN_CAP = 5000


def _com_datetime_variants(value):
    """
    Yield the forms a COM date argument may need, most likely first.

    pywin32 converts a naive datetime.datetime on most builds; where it does
    not, a pywintypes.Time does. Callers try each in turn because there is no
    safe string alternative - a formatted date string is read per the
    machine's regional settings, which is exactly the fault that made
    Restrict() unusable elsewhere in this file.
    """
    yield value
    try:
        yield pywintypes.Time(value)
    except Exception:
        pass


def _set_com_datetime(obj, attr, value):
    """Set a COM date property, trying each accepted date form."""
    last = None
    for candidate in _com_datetime_variants(value):
        try:
            setattr(obj, attr, candidate)
            return
        except Exception as exc:
            last = exc
    raise RuntimeError("Outlook would not accept {0}={1} ({2}).".format(
        attr, value, last))


def _prop(obj, schema, default=None):
    """Read a MAPI property through the PropertyAccessor, or return default."""
    try:
        return obj.PropertyAccessor.GetProperty(schema)
    except Exception:
        return default


# ---------------------------------------------------------------------------
# Fuzzy name matching (standard library only - difflib plus token scoring)
# ---------------------------------------------------------------------------

def _norm_text(text):
    """Lower-case, replace punctuation with spaces, collapse whitespace."""
    lowered = (text or "").lower()
    return " ".join(re.split(r"[^a-z0-9]+", lowered)).strip()


def _tokens(text):
    """The comparable words of a name: 'Smith, Joshua' -> ['smith', 'joshua']."""
    return [part for part in _norm_text(text).split(" ") if part]


def _token_score(query_token, candidate_token):
    """How well one typed word matches one word of a directory entry, 0.0-1.0."""
    if query_token == candidate_token:
        return 1.0
    if candidate_token.startswith(query_token):
        return 0.92                      # "josh" -> "joshua", "r5" -> "r5"
    if query_token.startswith(candidate_token):
        return 0.82                      # "joshua" typed, "josh" in the GAL
    if query_token in candidate_token:
        return 0.70
    ratio = difflib.SequenceMatcher(None, query_token, candidate_token).ratio()
    return ratio * 0.8 if ratio >= 0.8 else 0.0   # typos: "smtih" -> "smith"


def name_score(query_tokens, candidate):
    """
    Score a directory entry against the words the user typed, 0.0-1.0.

    Every typed word has to find a home in the candidate, and the score is the
    mean of each word's best match - so word ORDER does not matter ("Josh
    Smith" scores full marks against "Smith, Joshua") while a wrong surname
    still drags the score down ("Josh Taylor" does not match "Josh Smith").
    A whole-string similarity is taken as a floor so a run-together code like
    "r5-3-84" still lands on "Room R5-3-84 (12 seats)".
    """
    candidate_tokens = _tokens(candidate)
    if not candidate_tokens or not query_tokens:
        return 0.0
    per_token = [max(_token_score(q, c) for c in candidate_tokens)
                 for q in query_tokens]
    score = sum(per_token) / len(per_token)
    if min(per_token) >= 0.7:
        # Every word matched something: lift it clear of partial matches.
        score = min(1.0, score + 0.05)
    whole = difflib.SequenceMatcher(
        None, " ".join(query_tokens), " ".join(candidate_tokens)).ratio()
    return max(score, whole)


# ---------------------------------------------------------------------------
# Reading the directory
# ---------------------------------------------------------------------------

def _entry_display_type(entry):
    """
    PR_DISPLAY_TYPE_EX for an address entry, masked to its type value.

    The high bits are capability flags (ACL-capable and so on), so only the
    low word is the type: 7 is a room and 8 a piece of equipment. Returns None
    when the property cannot be read, which a cached offline address book does
    sometimes do - callers fall back to the address list the entry came from.
    """
    raw = _prop(entry, PROP_DISPLAY_TYPE_EX)
    try:
        return int(raw) & 0x0000FFFF
    except (TypeError, ValueError):
        return None


def _entry_smtp(entry):
    """Primary SMTP address of an address entry, or '' - never raises."""
    for getter in ("GetExchangeUser", "GetExchangeDistributionList"):
        try:
            obj = getattr(entry, getter)()
        except Exception:
            continue
        if obj is None:
            continue
        try:
            if obj.PrimarySmtpAddress:
                return obj.PrimarySmtpAddress
        except Exception:
            pass
    smtp = _prop(entry, PROP_SMTP_ADDRESS) or ""
    if smtp:
        return str(smtp)
    try:
        raw = entry.Address or ""
    except Exception:
        return ""
    # An Exchange entry's Address is the X.500 legacyExchangeDN. It resolves
    # for an invitation but is meaningless to show a user, so hide it here and
    # let the display name carry the identification.
    return "" if raw.upper().startswith("/O=") else raw


def _entry_details(entry):
    """Job title / department / office for an address entry, as a dict."""
    detail = {}
    try:
        user = entry.GetExchangeUser()
    except Exception:
        user = None
    if user is None:
        return detail
    for key, attribute in (("title", "JobTitle"),
                           ("department", "Department"),
                           ("office", "OfficeLocation"),
                           ("company", "CompanyName"),
                           ("phone", "BusinessTelephoneNumber")):
        try:
            value = getattr(user, attribute)
        except Exception:
            continue
        if value:
            detail[key] = str(value).strip()
    return detail


def _looks_like_room(name):
    """Name heuristic, used only when PR_DISPLAY_TYPE_EX is unreadable."""
    lowered = " {0} ".format((name or "").lower())
    return any(hint in lowered for hint in ROOM_NAME_HINTS)


def _address_lists():
    """Every address list in the profile as (name, AddressList) pairs."""
    found = []
    try:
        lists = get_namespace().AddressLists
        for index in range(1, int(lists.Count) + 1):
            try:
                address_list = lists.Item(index)
                found.append(((address_list.Name or ""), address_list))
            except Exception:
                continue
    except Exception as exc:
        log("Could not read the profile's address lists: {0}".format(exc))
    return found


def _ordered_address_lists(kind):
    """
    Address lists in the order worth scanning for this kind of search.

    Looking for a room, the room lists come first and everything in them is a
    room by definition, so no per-entry property read is needed. Looking for a
    person, the room lists are skipped entirely: they are pure noise and they
    cost a full scan each.
    """
    room_lists, other_lists = [], []
    for name, address_list in _address_lists():
        lowered = name.lower()
        if any(hint in lowered for hint in ROOM_LIST_HINTS):
            room_lists.append((name, address_list, True))
        else:
            other_lists.append((name, address_list, False))

    # The Global Address List is where a person is found, so scan it first.
    other_lists.sort(key=lambda row: 0 if "global address list" in row[0].lower() else 1)
    if kind == "room":
        return room_lists + other_lists
    if kind == "person":
        return other_lists
    return other_lists + room_lists


def _match_record(name, smtp, kind, source, score, detail=None):
    """One search hit, in the shape the tool output and the caller both use."""
    return {
        "name": (name or "").strip(),
        "smtp": (smtp or "").strip(),
        "kind": kind,
        "source": source,
        "score": score,
        "detail": detail or {},
        # What to hand outlook_suggest_meeting_times / outlook_draft_meeting.
        # An SMTP address is unambiguous; a display name is what is left when
        # the directory will not give one up, and Outlook resolves it by name.
        "invite_as": (smtp or name or "").strip(),
    }


def _dedupe_key(record):
    """Two hits are the same person if they share an address, else a name."""
    return (record["smtp"] or record["name"]).lower()


def _scan_contacts(query_tokens, kind, results, scanned):
    """
    Search the default Contacts folder and its immediate sub-folders.

    Not every store's contacts - just the user's own, which is where a name
    the directory does not hold (an external counterpart, a supplier) lives.
    They are small, so this costs nothing next to an address-list scan.
    """
    ns = get_namespace()
    folders = []
    try:
        root = ns.GetDefaultFolder(OL_FOLDER_CONTACTS)
        folders.append(root)
        for index in range(1, int(root.Folders.Count) + 1):
            try:
                folders.append(root.Folders.Item(index))
            except Exception:
                continue
    except Exception as exc:
        log("No Contacts folder could be read: {0}".format(exc))
        return scanned

    for folder in folders:
        try:
            items = folder.Items
            count = int(items.Count)
        except Exception:
            continue
        for index in range(1, count + 1):
            if scanned >= CONTACT_SCAN_CAP:
                return scanned
            scanned += 1
            try:
                item = items.Item(index)
                if item.Class != OL_CLASS_CONTACT:
                    continue
                name = (item.FullName or item.FileAs or "").strip()
                if not name:
                    continue
                company = (item.CompanyName or "").strip()
                try:
                    email = (item.Email1Address or "").strip()
                except Exception:
                    email = ""
            except Exception:
                continue

            # Match against the name, the company, and the address, so
            # "someone at Contoso" and a part-remembered address both work.
            score = max(name_score(query_tokens, name),
                        name_score(query_tokens, "{0} {1}".format(name, company)),
                        name_score(query_tokens, email))
            if score < MIN_NAME_SCORE:
                continue
            is_room = _looks_like_room(name)
            if kind == "room" and not is_room:
                continue
            if kind == "person" and is_room:
                continue
            detail = {}
            if company:
                detail["company"] = company
            try:
                if item.JobTitle:
                    detail["title"] = item.JobTitle.strip()
            except Exception:
                pass
            results.append(_match_record(
                name, email, "room" if is_room else "person",
                "Contacts ({0})".format(folder.Name), score, detail))
    return scanned


def _scan_address_lists(query_tokens, kind, results, cap):
    """
    Search the profile's address lists, newest-match-wins, bounded by `cap`.

    Only the display NAME is read for every entry - that is one COM round trip
    each and an enterprise GAL has a lot of entries. The expensive reads (SMTP
    address, job title, room-or-person) happen only for the handful that score
    well enough to be shown.

    Returns (entries_scanned, hit_cap).
    """
    scanned = 0
    for list_name, address_list, is_room_list in _ordered_address_lists(kind):
        try:
            entries = address_list.AddressEntries
            total = int(entries.Count)
        except Exception:
            continue
        for index in range(1, total + 1):
            if scanned >= cap:
                return scanned, True
            scanned += 1
            try:
                entry = entries.Item(index)
                name = (entry.Name or "").strip()
            except Exception:
                continue
            if not name:
                continue
            score = name_score(query_tokens, name)
            if score < MIN_NAME_SCORE:
                continue

            # Only now is it worth paying for the detail properties.
            display_type = _entry_display_type(entry)
            if display_type in (DT_DISTLIST, DT_PRIVATE_DISTLIST, DT_SEC_DISTLIST):
                entry_kind = "group"
            elif display_type in (DT_ROOM, DT_EQUIPMENT):
                entry_kind = "room"
            elif is_room_list:
                entry_kind = "room"          # everything in "All Rooms" is one
            elif display_type is None and _looks_like_room(name):
                entry_kind = "room"          # OAB withheld the property
            else:
                entry_kind = "person"

            if kind == "room" and entry_kind != "room":
                continue
            if kind == "person" and entry_kind == "room":
                continue

            results.append(_match_record(
                name, _entry_smtp(entry), entry_kind, list_name, score,
                _entry_details(entry)))
    return scanned, False


def _resolve_exactly(query, kind):
    """
    Ask Outlook to resolve the name the way its own Check Names does.

    Exchange ambiguous-name resolution is instant and knows the directory far
    better than any scan, so an unmistakable name costs no scan at all. It
    fails (rather than choosing) when the name is ambiguous, which is the
    signal to fall back to the fuzzy scan.
    """
    try:
        recipient = get_namespace().CreateRecipient(query)
        recipient.Resolve()
        if not recipient.Resolved:
            return None
        entry = recipient.AddressEntry
        name = (entry.Name or query).strip()
    except Exception:
        return None

    display_type = _entry_display_type(entry)
    if display_type in (DT_DISTLIST, DT_PRIVATE_DISTLIST, DT_SEC_DISTLIST):
        entry_kind = "group"
    elif display_type in (DT_ROOM, DT_EQUIPMENT) or _looks_like_room(name):
        entry_kind = "room"
    else:
        entry_kind = "person"
    if kind != "any" and entry_kind != kind:
        return None
    return _match_record(name, _entry_smtp(entry), entry_kind,
                         "resolved directly by Outlook", 1.0,
                         _entry_details(entry))


def find_directory_matches(query, kind="any", count=8, deep=False):
    """
    Fuzzy-find people and rooms. Returns (matches, notes).

    Order of work, fastest first:
      1. Outlook's own ambiguous-name resolution. An exact, unambiguous name
         is answered immediately and nothing is scanned.
      2. The Contacts folders - small, and where the user's own people are.
      3. The address lists (Global Address List, All Rooms, anything else),
         bounded by GAL_SCAN_CAP entries.
    """
    query = (query or "").strip()
    notes = []
    if not query:
        return [], ["No search text was given."]
    query_tokens = _tokens(query)
    if not query_tokens:
        return [], ["'{0}' has no searchable words in it.".format(query)]

    exact = _resolve_exactly(query, kind)
    if exact is not None and not deep:
        notes.append("Outlook resolved '{0}' directly, so the directory was "
                     "not scanned. Ask again with deep=true to list "
                     "alternatives.".format(query))
        return [exact], notes

    results = []
    if exact is not None:
        results.append(exact)

    contacts_scanned = _scan_contacts(query_tokens, kind, results, 0)
    listed, hit_cap = _scan_address_lists(query_tokens, kind, results,
                                          _GAL_SCAN_CAP)
    if hit_cap:
        notes.append(
            "Stopped after {0} entries (OUTLOOK_GAL_SCAN_CAP) without reaching "
            "the end of the directory. The scan runs in the directory's own "
            "order, so shortening the search does NOT move where it stopped: "
            "if the person is missing, give their full name or email address "
            "(Outlook then resolves it directly, without scanning), or raise "
            "OUTLOOK_GAL_SCAN_CAP.".format(_GAL_SCAN_CAP))
    log("Directory search '{0}': {1} contact(s) and {2} address-list entr(ies) "
        "scanned, {3} match(es).".format(query, contacts_scanned, listed,
                                         len(results)))

    # Keep the best-scoring copy of each person, then rank.
    best = {}
    for record in results:
        key = _dedupe_key(record)
        if key not in best or record["score"] > best[key]["score"]:
            best[key] = record
    ranked = sorted(best.values(),
                    key=lambda row: (-row["score"], row["name"].lower()))
    return ranked[:max(1, count)], notes


def tool_find_people(args):
    query = (args.get("query") or "").strip()
    if not query:
        return "Error: 'query' is required - the name, room or address to look for."
    kind = (args.get("kind") or "any").strip().lower()
    if kind not in ("any", "person", "room"):
        return "Error: 'kind' must be 'any', 'person' or 'room'."
    count = max(1, min(int(args.get("count", 8)), 25))
    deep = bool(args.get("deep", False))

    matches, notes = find_directory_matches(query, kind, count, deep)
    if not matches:
        lines = ["No {0} in Outlook's directory matches '{1}'.".format(
            {"any": "person or room", "person": "person", "room": "room"}[kind],
            query)]
        lines.extend(notes)
        lines.append("Try fewer words, a surname on its own, or part of the "
                     "email address. Rooms are often named by code "
                     "(\"R5-3-84\") rather than in words.")
        return "\n".join(lines)

    lines = ["{0} match(es) for '{1}':".format(len(matches), query), ""]
    for position, record in enumerate(matches, start=1):
        lines.append("{0}. {1}  [{2}]  match {3:.2f}".format(
            position, record["name"] or "(unnamed)", record["kind"],
            record["score"]))
        if record["smtp"]:
            lines.append("     address   : {0}".format(record["smtp"]))
        detail = record["detail"]
        job = ", ".join(part for part in (detail.get("title"),
                                          detail.get("department")) if part)
        if job:
            lines.append("     role      : {0}".format(job))
        for key, label in (("company", "company   "), ("office", "office    "),
                           ("phone", "phone     ")):
            if detail.get(key):
                lines.append("     {0}: {1}".format(label, detail[key]))
        lines.append("     found in  : {0}".format(record["source"]))
        lines.append("     invite_as : {0}".format(record["invite_as"]))
    if notes:
        lines.append("")
        lines.extend(notes)
    lines.append("")
    lines.append("Pass the invite_as value to outlook_suggest_meeting_times "
                 "(attendees / rooms) and then to outlook_draft_meeting. "
                 "Confirm the person with the user before inviting anyone.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Meeting scheduling - part 2: free/busy and suggesting a time
# ---------------------------------------------------------------------------

# Recipient.FreeBusy always answers for 30 days from midnight of the date it
# is given. Nothing can widen that, so a search window is clamped to it.
FREEBUSY_DAYS = 30

# Minutes per character in the free/busy string. 15 lets a suggestion land on
# a quarter hour; Exchange rounds up to whatever it actually stores.
FREEBUSY_MIN_PER_CHAR = 15

# The detail digits CompleteFormat=True returns, and what each means for a
# suggested slot.
FB_FREE = "0"
FB_TENTATIVE = "1"
FB_BUSY = "2"
FB_OOF = "3"
FB_ELSEWHERE = "4"

FB_LABELS = {
    FB_FREE: "free",
    FB_TENTATIVE: "tentative",
    FB_BUSY: "busy",
    FB_OOF: "out of office",
    FB_ELSEWHERE: "working elsewhere",
}

# Digits that make a slot unusable. Tentative is deliberately NOT here for a
# person - a pencilled-in hold is worth offering with a note, and an
# enterprise calendar is full of them. For a ROOM it IS blocking: a tentative
# room booking is somebody else's hold on the same space.
HARD_BUSY = (FB_BUSY, FB_OOF)

# Working elsewhere is not a clash: they are working, just not at their desk.
SOFT_FLAGS = (FB_TENTATIVE, FB_ELSEWHERE)

MEETING_DEFAULT_MINUTES = 30
MEETING_DEFAULT_DAYS = 10       # how far ahead a search looks by default
MEETING_SLOT_MINUTES = 15       # granularity of the candidate start times
MEETING_MAX_SUGGESTIONS = 5


def _free_busy(recipient, start_date, min_per_char=FREEBUSY_MIN_PER_CHAR):
    """
    A recipient's free/busy string, or None when Exchange has none to give.

    CompleteFormat=True asks for the detail digits (free / tentative / busy /
    out of office / working elsewhere) rather than a plain busy flag, which is
    what lets a suggestion say somebody is only pencilled in.

    None means UNKNOWN, never "free". An external attendee, a mailbox that
    does not publish, or an Availability service that cannot be reached all
    land here, and the caller has to say so rather than quietly offering a
    time nobody has checked.
    """
    origin = datetime.datetime.combine(start_date, datetime.time(0, 0))
    last = None
    for when in _com_datetime_variants(origin):
        try:
            value = recipient.FreeBusy(when, min_per_char, True)
        except Exception as exc:
            last = exc
            continue
        return str(value) if value else None
    log("Free/busy unavailable for '{0}': {1}".format(
        getattr(recipient, "Name", "?"), last))
    return None


def _fb_state(free_busy, origin, when, min_per_char=FREEBUSY_MIN_PER_CHAR):
    """The free/busy digit covering one moment, or None if it is not covered."""
    if not free_busy:
        return None
    index = int((when - origin).total_seconds() // 60) // min_per_char
    if index < 0 or index >= len(free_busy):
        return None
    return free_busy[index]


def _slot_states(participant, origin, start, end):
    """Every distinct free/busy digit a participant shows across one slot."""
    states = set()
    cursor = start
    while cursor < end:
        states.add(_fb_state(participant["free_busy"], origin, cursor))
        cursor += datetime.timedelta(minutes=FREEBUSY_MIN_PER_CHAR)
    # The final minute matters too: a slot ending at 10:30 must not be judged
    # on 10:15 alone when the block from 10:15 is busy from 10:29.
    states.add(_fb_state(participant["free_busy"], origin,
                         end - datetime.timedelta(minutes=1)))
    return states


def _blocking_state(participant, states):
    """
    The digit that rules a participant out of a slot, or None.

    A room is held to a stricter rule than a person: anything other than free
    blocks it, because a tentative room is already somebody's hold on the
    space, while a tentative person is worth asking.
    """
    if participant["kind"] == "room":
        for state in states:
            if state is not None and state != FB_FREE:
                return state
        return None
    for state in states:
        if state in HARD_BUSY:
            return state
    return None


def _build_participants(specs, role, kind):
    """
    Resolve a list of names/addresses into participants with free/busy loaded.

    A spec that will not resolve is an error the caller must see, not one to
    skip: quietly dropping an attendee is how a meeting gets scheduled without
    the person it was for.
    """
    participants, unresolved = [], []
    ns = get_namespace()
    for spec in specs:
        spec = (spec or "").strip()
        if not spec:
            continue
        try:
            recipient = ns.CreateRecipient(spec)
            recipient.Resolve()
            resolved = bool(recipient.Resolved)
        except Exception as exc:
            log("Could not resolve '{0}': {1}".format(spec, exc))
            resolved = False
        if not resolved:
            unresolved.append(spec)
            continue
        try:
            name = (recipient.Name or spec).strip()
        except Exception:
            name = spec
        participants.append({
            "spec": spec,
            "name": name,
            "kind": kind,
            "role": role,
            "recipient": recipient,
            "free_busy": None,
        })
    return participants, unresolved


def _load_free_busy(participants, start_date, warnings):
    """Fill in each participant's free/busy, noting anyone Exchange cannot answer for."""
    blind = []
    for participant in participants:
        participant["free_busy"] = _free_busy(participant["recipient"], start_date)
        if participant["free_busy"] is None:
            blind.append(participant["name"])
    if blind:
        warnings.append(
            "No free/busy published for: {0}. Their diary was NOT checked, "
            "so any suggestion here may still clash for them - say so when "
            "you offer the times.".format(", ".join(blind)))


def _spread_options(candidates, limit, max_per_day=2):
    """
    Reduce ranked slots to options a person can actually choose between.

    Candidates come at the slot granularity, so the five best are usually the
    same afternoon shifted a quarter of an hour at a time - five options on
    paper and one in practice. Take the best, then skip anything overlapping a
    slot already taken or making a third on the same day, so the list reads as
    real alternatives. A second pass without the per-day cap fills the list
    when only one day has anything free in it.
    """
    chosen, per_day = [], {}
    for cap in (max_per_day, None):
        for slot in candidates:
            if len(chosen) >= limit:
                return chosen
            day = slot["start"].date()
            if cap is not None and per_day.get(day, 0) >= cap:
                continue
            if any(slot["start"] < taken["end"] and taken["start"] < slot["end"]
                   for taken in chosen):
                continue
            chosen.append(slot)
            per_day[day] = per_day.get(day, 0) + 1
    return chosen


def _describe_span(start, end):
    """'Tue 16 Sep 2026, 10:00-10:30'."""
    return "{0} {1:02d} {2} {3}, {4:%H:%M}-{5:%H:%M}".format(
        _WEEKDAY_SHORT[start.weekday()].title(), start.day,
        _MONTH_SHORT[start.month - 1].title(), start.year, start, end)


def tool_suggest_meeting_times(args):
    required_specs = args.get("attendees") or []
    optional_specs = args.get("optional_attendees") or []
    room_specs = args.get("rooms") or []
    if isinstance(required_specs, str):
        required_specs = [required_specs]
    if isinstance(optional_specs, str):
        optional_specs = [optional_specs]
    if isinstance(room_specs, str):
        room_specs = [room_specs]
    if not (required_specs or optional_specs or room_specs):
        return ("Error: give at least one of 'attendees', 'optional_attendees' "
                "or 'rooms'. Use outlook_find_people first to turn a name like "
                "'Josh Smith' into an address.")

    duration = _int_arg(args, "duration_minutes", MEETING_DEFAULT_MINUTES)
    if duration < 5 or duration > 24 * 60:
        return "Error: 'duration_minutes' must be between 5 and 1440."
    slot_minutes = _int_arg(args, "slot_minutes", MEETING_SLOT_MINUTES)
    if slot_minutes < 5 or slot_minutes > 120:
        return "Error: 'slot_minutes' must be between 5 and 120."
    max_suggestions = max(1, min(_int_arg(args, "max_suggestions",
                                          MEETING_MAX_SUGGESTIONS), 20))
    include_weekends = bool(args.get("include_weekends", False))
    include_self = bool(args.get("include_self", True))

    today = datetime.date.today()
    try:
        start_date = parse_planner_date(args.get("start_date"), today)
    except ValueError:
        return ("Error: 'start_date' must be YYYY-MM-DD, 'today', 'tomorrow', "
                "or a day offset like '+2'.")
    days = _int_arg(args, "days", MEETING_DEFAULT_DAYS)
    if days < 1:
        return "Error: 'days' must be 1 or more."
    capped_days = min(days, FREEBUSY_DAYS)

    earliest = _int_arg(args, "earliest_hour", _MEETING_START_HOUR)
    latest = _int_arg(args, "latest_hour", _MEETING_END_HOUR)
    if not 0 <= earliest < latest <= 24:
        return ("Error: 'earliest_hour' and 'latest_hour' must be hours 0-24 "
                "with the earliest first.")

    warnings = []
    required, missing_required = _build_participants(
        required_specs, "required", "person")
    optional, missing_optional = _build_participants(
        optional_specs, "optional", "person")
    rooms, missing_rooms = _build_participants(
        room_specs, "resource", "room")
    unresolved = missing_required + missing_optional + missing_rooms
    if unresolved:
        return ("Error: Outlook could not resolve these against the directory: "
                "{0}.\nRun outlook_find_people on each one and pass the "
                "'invite_as' value it returns. Nothing was checked."
                .format(", ".join(unresolved)))

    if include_self:
        try:
            me = get_namespace().CurrentUser
            required.insert(0, {
                "spec": "(you)",
                "name": "{0} (you)".format(me.Name),
                "kind": "person",
                "role": "required",
                "recipient": me,
                "free_busy": None,
            })
        except Exception as exc:
            warnings.append("Could not read your own calendar for the check "
                            "({0}); your time was not considered.".format(exc))

    everyone = required + optional + rooms
    _load_free_busy(everyone, start_date, warnings)

    origin = datetime.datetime.combine(start_date, datetime.time(0, 0))
    now = datetime.datetime.now()
    step = datetime.timedelta(minutes=slot_minutes)
    length = datetime.timedelta(minutes=duration)

    clear, near_miss = [], []
    for offset in range(capped_days):
        day = start_date + datetime.timedelta(days=offset)
        if not include_weekends and day.weekday() >= 5:
            continue
        day_start = datetime.datetime.combine(day, datetime.time(earliest, 0))
        if latest == 24:
            day_end = datetime.datetime.combine(
                day, datetime.time(0, 0)) + datetime.timedelta(days=1)
        else:
            day_end = datetime.datetime.combine(day, datetime.time(latest, 0))

        slot_start = day_start
        while slot_start + length <= day_end:
            slot_end = slot_start + length
            if slot_start <= now:
                slot_start += step
                continue

            blockers, flags, optional_out, optional_free = [], [], [], 0
            for participant in everyone:
                states = _slot_states(participant, origin, slot_start, slot_end)
                blocked = _blocking_state(participant, states)
                if blocked is not None:
                    if participant["role"] == "optional":
                        # An optional clash never rules a slot out, but the
                        # user still has to be told who would miss it.
                        optional_out.append(participant["name"])
                        continue
                    blockers.append((participant["name"], FB_LABELS.get(
                        blocked, blocked)))
                    continue
                if participant["role"] == "optional":
                    optional_free += 1
                soft = [state for state in SOFT_FLAGS if state in states]
                if soft:
                    flags.append("{0} is {1}".format(
                        participant["name"], FB_LABELS[soft[0]]))

            candidate = {
                "start": slot_start,
                "end": slot_end,
                "flags": flags + ["{0} can't make it (optional)".format(name)
                                  for name in optional_out],
                "blockers": blockers,
                "optional_free": optional_free,
            }
            if not blockers:
                clear.append(candidate)
            elif len(blockers) == 1:
                near_miss.append(candidate)
            slot_start += step

    lines = []
    window_end = start_date + datetime.timedelta(days=capped_days - 1)
    lines.append(
        "Looking for {0} minutes between {1} and {2}, {3:02d}:00-{4:02d}:00{5}."
        .format(duration, start_date.isoformat(), window_end.isoformat(),
                earliest, latest,
                "" if include_weekends else ", weekdays only"))
    lines.append("Checked: {0}".format(", ".join(
        "{0} [{1}]".format(person["name"],
                           "room" if person["kind"] == "room" else person["role"])
        for person in everyone) or "nobody"))
    if days > capped_days:
        lines.append("Outlook only publishes {0} days of free/busy, so the "
                     "search stopped at {1}.".format(FREEBUSY_DAYS, window_end))
    lines.append("")

    if clear:
        # Fewest caveats first, then the most optional people who can make it,
        # then earliest - which is what "a suitable time" usually means.
        clear.sort(key=lambda slot: (len(slot["flags"]),
                                     -slot["optional_free"], slot["start"]))
        shown = _spread_options(clear, max_suggestions)
        lines.append("{0} suggestion(s), best first:".format(len(shown)))
        for position, slot in enumerate(shown, start=1):
            note = "; ".join(slot["flags"]) if slot["flags"] else "everyone free"
            lines.append("{0}. {1}   {2}".format(
                position, _describe_span(slot["start"], slot["end"]), note))
            lines.append("     start: \"{0:%Y-%m-%d %H:%M}\"  duration_minutes: {1}"
                         .format(slot["start"], duration))
        # Count only genuinely different times, not the same slot nudged along
        # by a quarter of an hour, or the offer of "more" is a lie.
        spare = len(_spread_options(clear, len(clear))) - len(shown)
        if spare > 0:
            lines.append("({0} other non-overlapping option(s) available - ask "
                         "for more, or for a particular day.)".format(spare))
    else:
        lines.append("Nothing is clear for everyone in that window.")
        if near_miss:
            near_miss.sort(key=lambda slot: slot["start"])
            lines.append("")
            lines.append("Closest, each with ONE person in the way:")
            for slot in _spread_options(near_miss, max_suggestions):
                who, why = slot["blockers"][0]
                lines.append("  {0}   {1} is {2}".format(
                    _describe_span(slot["start"], slot["end"]), who, why))
            lines.append("")
            lines.append("Widen the hours, move the person in the way to "
                         "optional_attendees, or search further ahead.")
        else:
            lines.append("Try a longer window ('days'), wider hours, or a "
                         "shorter meeting.")

    if warnings:
        lines.append("")
        lines.extend("Note: {0}".format(warning) for warning in warnings)
    lines.append("")
    if clear:
        lines.append("Free/busy is what Exchange publishes - the same source "
                     "as Outlook's Scheduling Assistant. Put these to the user, "
                     "and once they pick one call outlook_draft_meeting with "
                     "the start and duration shown.")
    else:
        lines.append("Free/busy is what Exchange publishes - the same source "
                     "as Outlook's Scheduling Assistant. Do not invent a time: "
                     "ask the user which constraint to relax.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Meeting scheduling - part 3: saving an UNSENT meeting draft
# ---------------------------------------------------------------------------
#
# One of the two things in this server that write to the mailbox (the other
# is the email draft below), and it is
# deliberately the smallest write that is useful: a meeting saved into the
# Calendar with its invitations NOT sent - the state Outlook labels
# "Invitations have not been sent for this meeting". The user opens it, edits
# whatever they like, and presses Send themselves.
#
# Send() is never called anywhere in this file, and must not be added. Set
# OUTLOOK_ALLOW_DRAFTS=false to remove the tool altogether, after which the
# server is read-only on the mailbox again.

MEETING_MAX_BODY_CHARS = 20000


def parse_meeting_start(value):
    """
    Parse the 'start' argument into a naive datetime.

    Accepts 'YYYY-MM-DD HH:MM' (what outlook_suggest_meeting_times prints),
    'YYYY-MM-DDTHH:MM', and the same with seconds. Deliberately strict: a
    meeting written at the wrong time is worse than a rejected argument, and
    there is no locale-dependent parsing here for the same reason Restrict()
    is avoided elsewhere in this file.
    """
    text = (value or "").strip().replace("T", " ")
    for pattern in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.datetime.strptime(text, pattern)
        except ValueError:
            continue
    raise ValueError(
        "'start' must be 'YYYY-MM-DD HH:MM' (24-hour), e.g. "
        "'2026-09-16 10:00'. Got: {0!r}".format(value))


# Markdown that must not survive into a meeting invitation. An agenda drafted
# in conversation arrives as Markdown by habit, and AppointmentItem.Body is
# PLAIN TEXT - Outlook renders none of it, so "**Agenda**" reaches everyone
# invited with the asterisks still on it.
_MD_STRONG = re.compile(r"(\*\*|__)(?=\S)(.+?)(?<=\S)\1", re.DOTALL)
_MD_HEADING = re.compile(r"^[ \t]{0,3}#{1,6}[ \t]+", re.MULTILINE)
_MD_BULLET = re.compile(r"^([ \t]*)[*+][ \t]+", re.MULTILINE)


def plain_text_body(text):
    """
    Strip the Markdown an agenda arrives with, leaving plain text.

    Deliberately narrow: **bold** and __bold__, a leading #-heading, and * or
    + bullets normalised to -. Nothing else.

    A LONE asterisk is left alone, because in an agenda it is as likely to be
    a footnote mark as emphasis. "C#" and "#3" survive too, because only a #
    followed by a SPACE at the start of a line is a heading. Getting this
    wrong is visible to everyone invited, so it does less rather than more.
    """
    cleaned = _MD_STRONG.sub(r"\2", text)
    cleaned = _MD_HEADING.sub("", cleaned)
    return _MD_BULLET.sub(r"\1- ", cleaned)


def _add_meeting_recipients(appointment, specs, recipient_type, label, failures):
    """Attach one class of recipient, recording anything that will not resolve."""
    added = []
    for spec in specs:
        spec = (spec or "").strip()
        if not spec:
            continue
        try:
            recipient = appointment.Recipients.Add(spec)
            recipient.Type = recipient_type
            recipient.Resolve()
            if not recipient.Resolved:
                failures.append("{0} ({1})".format(spec, label))
                continue
            added.append((recipient.Name or spec, spec))
        except Exception as exc:
            log("Could not add {0} '{1}': {2}".format(label, spec, exc))
            failures.append("{0} ({1})".format(spec, label))
    return added


def tool_draft_meeting(args):
    if not _ALLOW_DRAFTS:
        return ("Meeting drafting is switched off on this endpoint "
                "(OUTLOOK_ALLOW_DRAFTS=false), so nothing can be written to "
                "the calendar. Tell the user the suggested time and let them "
                "create the meeting in Outlook themselves.")

    subject = (args.get("subject") or "").strip()
    if not subject:
        return "Error: 'subject' is required."

    try:
        start = parse_meeting_start(args.get("start"))
    except ValueError as exc:
        return "Error: {0}".format(exc)

    duration = _int_arg(args, "duration_minutes", MEETING_DEFAULT_MINUTES)
    end_text = (args.get("end") or "").strip()
    if end_text:
        try:
            end = parse_meeting_start(end_text)
        except ValueError as exc:
            return "Error: {0}".format(exc)
        duration = int((end - start).total_seconds() // 60)
    if duration < 5 or duration > 24 * 60:
        return ("Error: the meeting must be between 5 minutes and 24 hours "
                "long (got {0} minutes).".format(duration))
    end = start + datetime.timedelta(minutes=duration)

    required_specs = args.get("attendees") or []
    optional_specs = args.get("optional_attendees") or []
    room_specs = args.get("rooms") or []
    for name, value in (("attendees", required_specs),
                        ("optional_attendees", optional_specs),
                        ("rooms", room_specs)):
        if isinstance(value, str):
            value = [value]
        if name == "attendees":
            required_specs = value
        elif name == "optional_attendees":
            optional_specs = value
        else:
            room_specs = value
    if not (required_specs or optional_specs or room_specs):
        return ("Error: a meeting needs somebody in it - give 'attendees', "
                "'optional_attendees' or 'rooms'. Use outlook_find_people to "
                "turn a name into an address first.")

    raw_body = (args.get("body") or "").strip()
    if len(raw_body) > MEETING_MAX_BODY_CHARS:
        return "Error: 'body' is longer than {0} characters.".format(
            MEETING_MAX_BODY_CHARS)
    body = plain_text_body(raw_body)

    # The content blacklist applies in BOTH directions. A subject or body
    # carrying a protective marking must not be routed through the AI on its
    # way into Outlook any more than on its way out. Both forms of the body
    # are scanned, so stripping markup can never open a gap in the filter.
    marking = blacklisted_match("{0}\n{1}\n{2}".format(subject, raw_body, body))
    if marking:
        log("Refused to draft a meeting (blacklist match: {0}).".format(marking))
        return ("Refused: the subject or body of this meeting contains "
                "content the local content policy withholds from the AI. "
                "Create this meeting in Outlook directly.")

    busy_word = (args.get("busy_status") or "busy").strip().lower()
    if busy_word not in BUSY_STATUS_NAMES:
        return "Error: 'busy_status' must be one of: {0}.".format(
            ", ".join(sorted(set(BUSY_STATUS_NAMES))))
    reminder = _int_arg(args, "reminder_minutes", 15)
    if reminder < 0 or reminder > 40320:
        return "Error: 'reminder_minutes' must be 0 (no reminder) to 40320."

    categories = args.get("categories") or []
    if isinstance(categories, str):
        categories = [categories]

    ns = get_namespace()
    appointment = ns.Application.CreateItem(OL_APPOINTMENT_ITEM)
    # MeetingStatus has to be set BEFORE the recipients go on, or Outlook
    # keeps the item as a private appointment and silently drops them.
    appointment.MeetingStatus = OL_MEETING
    appointment.Subject = subject
    _set_com_datetime(appointment, "Start", start)
    appointment.Duration = duration
    appointment.BusyStatus = BUSY_STATUS_NAMES[busy_word]
    if body:
        appointment.Body = body
    if reminder:
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = reminder
    else:
        appointment.ReminderSet = False
    if categories:
        appointment.Categories = "; ".join(str(one).strip()
                                           for one in categories if str(one).strip())

    failures = []
    added_required = _add_meeting_recipients(
        appointment, required_specs, OL_RECIP_REQUIRED, "required", failures)
    added_optional = _add_meeting_recipients(
        appointment, optional_specs, OL_RECIP_OPTIONAL, "optional", failures)
    added_rooms = _add_meeting_recipients(
        appointment, room_specs, OL_RECIP_RESOURCE, "room", failures)

    if failures:
        # Never save a meeting that is missing somebody. A half-addressed
        # draft looks finished and gets sent without the person it was for.
        try:
            appointment.Delete()
        except Exception:
            pass
        return ("Error: Outlook could not resolve {0}. NOTHING was saved.\n"
                "Run outlook_find_people on each one and pass the 'invite_as' "
                "value it returns.".format(", ".join(failures)))

    location = (args.get("location") or "").strip()
    if not location and added_rooms:
        location = "; ".join(name for name, _ in added_rooms)
    if location:
        appointment.Location = location

    appointment.Save()   # Save, never Send. See the note at the top of part 3.

    try:
        entry_id = appointment.EntryID
    except Exception:
        entry_id = "(unavailable)"

    opened = False
    if bool(args.get("open_in_outlook", True)):
        try:
            appointment.Display(False)   # modeless: returns straight away
            opened = True
        except Exception as exc:
            log("Could not open the draft in Outlook: {0}".format(exc))

    lines = ["Meeting DRAFT saved to your calendar - NOT sent.", ""]
    lines.append("  Subject  : {0}".format(subject))
    lines.append("  When     : {0} {1} {2} {3}, {4:%H:%M}-{5:%H:%M} "
                 "({6} minutes)".format(
                     _WEEKDAY_NAMES[start.weekday()], start.day,
                     _MONTH_NAMES[start.month - 1], start.year,
                     start, end, duration))
    if location:
        lines.append("  Location : {0}".format(location))
    for label, people in (("Required", added_required),
                          ("Optional", added_optional),
                          ("Room    ", added_rooms)):
        if people:
            lines.append("  {0} : {1}".format(label, ", ".join(
                "{0} <{1}>".format(name, spec) if "@" in spec else name
                for name, spec in people)))
    lines.append("  Reminder : {0}".format(
        "{0} minutes before".format(reminder) if reminder else "none"))
    lines.append("  Shows as : {0}".format(busy_word))
    lines.append("  EntryID  : {0}".format(entry_id))
    if start < datetime.datetime.now():
        lines.append("")
        lines.append("  NOTE: that start time is in the past.")
    lines.append("")
    if opened:
        lines.append("It is open in Outlook now. Edit anything you want, then "
                     "press Send to invite everyone.")
    else:
        lines.append("Open it from your calendar on {0}, edit anything you "
                     "want, then press Send to invite everyone.".format(
                         start.date().isoformat()))
    lines.append("No invitation has gone out: this server cannot send, and "
                 "nobody has been notified.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Email drafting: an UNSENT email, opened in Outlook's own compose window
# ---------------------------------------------------------------------------
#
# The second write to the mailbox, and the same kind as the meeting draft: the
# smallest write that is useful. A new email, a reply, a reply-all or a forward
# is opened in Outlook's compose window with the agreed text in it, and saved
# to Drafts so it cannot be lost. The user reads it, edits it and presses Send
# themselves. Send() is never called anywhere in this file, and must not be
# added. OUTLOOK_ALLOW_DRAFTS=false removes this tool along with the meeting
# draft.
#
# The wording is agreed in the CONVERSATION first; this tool is only the
# hand-off, called when the user says to put it in Outlook. Its description
# says so, because a tool that pops a window open on every iteration of a
# draft is noise.
#
# HOW THE TEXT LOOKS LIKE THE USER'S OWN EMAIL
# Outlook adds the user's signature when a compose window is created for an
# item - the new-message signature on a new email, the replies/forwards one on
# a reply or forward, and for the account the item belongs to. So the item is
# created by Outlook itself (CreateItem, or Reply()/ReplyAll()/Forward() on the
# original, which also keeps the quoted thread and the conversation threading),
# its inspector is created to trigger the signature, and only THEN is the text
# written in.
#
# The text goes in through the compose window's Word editor
# (Inspector.WordEditor), at the very top of the message, where the cursor sits
# when the user starts typing. Text inserted there takes the formatting of that
# empty first line - the user's own compose font for a new email, their reply
# font for a reply - which is exactly what typing it would have done. Writing
# HTMLBody instead is the common approach and the wrong one for a reply: the
# inserted paragraphs pick up the ORIGINAL message's Normal style, which is how
# replies come out in Times New Roman. HTMLBody is kept only as a fallback for
# when the Word editor cannot be reached, and the reply says which was used.
#
# Markdown from the conversation is translated, not pasted: **bold**, *italic*,
# "- " / "* " bullets and "1. " numbered lists become real Outlook formatting,
# a "# heading" becomes a bold line, `code` loses its backticks, and
# [text](url) and bare http(s) URLs become links. Each source line becomes one
# line in the email and a blank line stays a blank line, so the spacing is the
# spacing that was agreed. In a PLAIN TEXT email (the user's setting, or a
# reply to a plain-text message) the formatting is dropped and bullets are
# written as "- ".

OL_MAIL_ITEM = 0            # CreateItem(olMailItem)
OL_FORMAT_PLAIN = 1         # MailItem.BodyFormat
OL_FORMAT_HTML = 2
OL_FORMAT_RICH_TEXT = 3
OL_DISCARD = 1              # MailItem.Close(olDiscard)

# Recipient.Type on a MailItem (same numbers as the meeting ones, different
# meaning, so they get their own names).
OL_MAIL_TO = 1
OL_MAIL_CC = 2
OL_MAIL_BCC = 3

EMAIL_MAX_BODY_CHARS = 20000

EMAIL_DRAFT_MODES = ("new", "reply", "reply_all", "forward")

_MD_EMAIL_HEADING = re.compile(r"^[ \t]{0,3}#{1,6}[ \t]+(.*?)[ \t]*#*[ \t]*$")
_MD_EMAIL_BULLET = re.compile(r"^[ \t]*[-*+][ \t]+(.*)$")
_MD_EMAIL_NUMBER = re.compile(r"^[ \t]*\d{1,3}[.)][ \t]+(.*)$")

# Inline Markdown, tried left to right. A lone asterisk (a footnote, "5 * 3")
# is left alone: emphasis needs a non-space on the inside of both marks and
# no letter on the outside, so "file*name*" is not italic either.
_MD_EMAIL_INLINE = re.compile(
    r"\[(?P<ltext>[^\]\n]+)\]\((?P<lurl>[^)\s]+)\)"
    r"|(?P<bmark>\*\*|__)(?=\S)(?P<bold>.+?)(?<=\S)(?P=bmark)"
    r"|(?<![\w*])\*(?=[^\s*])(?P<ital>.+?)(?<=[^\s*])\*(?![\w*])"
    r"|`(?P<code>[^`\n]+)`"
    r"|(?P<url>https?://[^\s<>()\[\]]*[^\s<>()\[\].,;:!?'\"])"
)


def _md_runs(text, bold=False, italic=False):
    """
    Split one line of Markdown into formatted runs.

    Returns a list of dicts: text, bold, italic, url (None unless a link).
    Bold and italic content is parsed again so a link inside **bold** still
    becomes a link; anything else is kept as typed.
    """
    runs = []
    pos = 0
    for match in _MD_EMAIL_INLINE.finditer(text):
        if match.start() > pos:
            runs.append({"text": text[pos:match.start()], "bold": bold,
                         "italic": italic, "url": None})
        if match.group("ltext") is not None:
            runs.append({"text": match.group("ltext"), "bold": bold,
                         "italic": italic, "url": match.group("lurl")})
        elif match.group("bold") is not None:
            runs.extend(_md_runs(match.group("bold"), True, italic))
        elif match.group("ital") is not None:
            runs.extend(_md_runs(match.group("ital"), bold, True))
        elif match.group("code") is not None:
            runs.append({"text": match.group("code"), "bold": bold,
                         "italic": italic, "url": None})
        else:
            url = match.group("url")
            runs.append({"text": url, "bold": bold, "italic": italic,
                         "url": url})
        pos = match.end()
    if pos < len(text):
        runs.append({"text": text[pos:], "bold": bold, "italic": italic,
                     "url": None})
    return [run for run in runs if run["text"]]


def markdown_to_email_blocks(text):
    """
    Turn a Markdown email body into one block per line.

    Each block is {"kind": "para" | "bullet" | "number" | "blank", "runs": [...]}.
    Leading and trailing blank lines are dropped (Outlook already leaves a gap
    above the signature); blank lines in between are kept, because that is how
    the paragraphs of an email are spaced.
    """
    blocks = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            blocks.append({"kind": "blank", "runs": []})
            continue
        heading = _MD_EMAIL_HEADING.match(line)
        bullet = _MD_EMAIL_BULLET.match(line)
        number = _MD_EMAIL_NUMBER.match(line)
        if heading:
            blocks.append({"kind": "para",
                           "runs": _md_runs(heading.group(1), bold=True)})
        elif bullet:
            blocks.append({"kind": "bullet", "runs": _md_runs(bullet.group(1))})
        elif number:
            blocks.append({"kind": "number", "runs": _md_runs(number.group(1))})
        else:
            blocks.append({"kind": "para", "runs": _md_runs(line.strip())})
    while blocks and blocks[0]["kind"] == "blank":
        blocks.pop(0)
    while blocks and blocks[-1]["kind"] == "blank":
        blocks.pop()
    return blocks


def _run_plain_text(run):
    """A run as plain text: a link keeps its address beside its words."""
    if run["url"] and run["text"] != run["url"]:
        return "{0} ({1})".format(run["text"], run["url"])
    return run["text"]


def email_blocks_plain_lines(blocks):
    """The blocks as plain-text lines, bullets as '- ' and numbers counted."""
    lines = []
    counter = 0
    for block in blocks:
        counter = counter + 1 if block["kind"] == "number" else 0
        text = "".join(_run_plain_text(run) for run in block["runs"])
        if block["kind"] == "bullet":
            text = "- " + text
        elif block["kind"] == "number":
            text = "{0}. {1}".format(counter, text)
        lines.append(text)
    return lines


def _utf16_len(text):
    """Length as Word counts it: UTF-16 code units, so an emoji counts as 2."""
    return len(text.encode("utf-16-le")) // 2


def _list_groups(blocks):
    """Runs of consecutive bullet / number blocks: (kind, first, last) indexes."""
    groups = []
    index = 0
    while index < len(blocks):
        kind = blocks[index]["kind"]
        if kind in ("bullet", "number"):
            first = index
            while index + 1 < len(blocks) and blocks[index + 1]["kind"] == kind:
                index += 1
            groups.append((kind, first, index))
        index += 1
    return groups


def insert_email_via_word(inspector, blocks, plain):
    """
    Write the blocks into the top of an open compose window's Word editor.

    Everything is inserted in ONE call at position 0 and then formatted by
    offset. Bold, italic and list formatting add no characters, so the offsets
    stay true; links DO change the document (a hyperlink is a field), so they
    go in last and from the end backwards, which leaves every earlier offset
    untouched. The final line is not followed by a paragraph mark: it joins
    the empty line the cursor was sitting on, so the gap Outlook leaves above
    the signature survives exactly as it was.

    On a failure part-way through, the inserted text is removed again before
    the error is raised, so the caller's fallback never writes the email twice.
    """
    document = inspector.WordEditor
    if document is None:
        raise RuntimeError("the compose window has no Word editor")

    if plain:
        texts = email_blocks_plain_lines(blocks)
    else:
        texts = ["".join(run["text"] for run in block["runs"])
                 for block in blocks]
    content = "\r".join(texts)

    document.Range(0, 0).InsertBefore(content)
    inserted_end = _utf16_len(content)
    try:
        # Where each line starts, in Word's character positions.
        starts = []
        offset = 0
        for text in texts:
            starts.append(offset)
            offset += _utf16_len(text) + 1        # +1 for the "\r"

        links = []
        if not plain:
            for block, start in zip(blocks, starts):
                at = start
                for run in block["runs"]:
                    length = _utf16_len(run["text"])
                    if run["bold"] or run["italic"]:
                        span = document.Range(at, at + length)
                        if run["bold"]:
                            span.Font.Bold = True
                        if run["italic"]:
                            span.Font.Italic = True
                    if run["url"]:
                        links.append((at, at + length, run["url"]))
                    at += length

            for kind, first, last in _list_groups(blocks):
                span = document.Range(
                    starts[first],
                    starts[last] + _utf16_len(texts[last]))
                if kind == "bullet":
                    span.ListFormat.ApplyBulletDefault()
                else:
                    span.ListFormat.ApplyNumberDefault()

        # Links last, from the end backwards (see the docstring). A link that
        # will not go in keeps its address as text, so nothing is lost.
        for start, end, url in reversed(links):
            try:
                document.Hyperlinks.Add(document.Range(start, end), url)
            except Exception as exc:
                log("Could not add a link to {0}: {1}".format(url, exc))
                document.Range(end, end).InsertAfter(" ({0})".format(url))
    except Exception:
        try:
            document.Range(0, inserted_end).Delete()
        except Exception:
            pass
        raise

    # Leave the cursor at the end of the text, where the user would carry on.
    try:
        document.Range(inserted_end, inserted_end).Select()
    except Exception:
        pass


def _html_escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def email_blocks_html(blocks):
    """The blocks as Outlook-flavoured HTML, for the HTMLBody fallback."""
    def runs_html(runs):
        out = []
        for run in runs:
            piece = _html_escape(run["text"])
            if run["url"]:
                piece = '<a href="{0}">{1}</a>'.format(
                    _html_escape(run["url"]), piece)
            if run["italic"]:
                piece = "<i>" + piece + "</i>"
            if run["bold"]:
                piece = "<b>" + piece + "</b>"
            out.append(piece)
        return "".join(out)

    html = []
    groups = {first: (kind, last) for kind, first, last in _list_groups(blocks)}
    index = 0
    while index < len(blocks):
        if index in groups:
            kind, last = groups[index]
            tag = "ul" if kind == "bullet" else "ol"
            items = "".join("<li class=MsoNormal>{0}</li>".format(
                runs_html(blocks[i]["runs"])) for i in range(index, last + 1))
            html.append("<{0} style='margin-top:0cm;margin-bottom:0cm'>{1}"
                        "</{0}>".format(tag, items))
            index = last + 1
            continue
        block = blocks[index]
        if block["kind"] == "blank":
            html.append("<p class=MsoNormal><o:p>&nbsp;</o:p></p>")
        else:
            html.append("<p class=MsoNormal>{0}</p>".format(
                runs_html(block["runs"])))
        index += 1
    return "".join(html)


_HTML_SECTION_OPEN = re.compile(r"<div\s+class=[\"']?WordSection1[\"']?[^>]*>",
                                re.IGNORECASE)
_HTML_BODY_OPEN = re.compile(r"<body[^>]*>", re.IGNORECASE)


def insert_email_via_body(mail, blocks, plain):
    """
    Fallback when the Word editor cannot be reached: write Body / HTMLBody.

    The text goes in ahead of whatever Outlook already put there (the
    signature, and on a reply the quoted thread), at the top of the message
    section so it sits where the user would have typed it.
    """
    if plain:
        existing = mail.Body or ""
        content = "\r\n".join(email_blocks_plain_lines(blocks))
        if not existing.startswith("\r\n"):
            content += "\r\n\r\n"
        mail.Body = content + existing
        return
    existing = mail.HTMLBody or ""
    fragment = email_blocks_html(blocks)
    anchor = _HTML_SECTION_OPEN.search(existing) or _HTML_BODY_OPEN.search(existing)
    if anchor:
        mail.HTMLBody = (existing[:anchor.end()] + fragment
                         + existing[anchor.end():])
    else:
        mail.HTMLBody = fragment + existing


def _email_recipient_list(args, name):
    value = args.get(name) or []
    if isinstance(value, str):
        # ";" only: a directory name is often "Smith, Josh".
        value = value.split(";")
    return [str(one).strip() for one in value if str(one).strip()]


def _add_email_recipients(mail, specs, recipient_type, label, failures):
    """Attach one class of recipient, recording anything that will not resolve."""
    for spec in specs:
        try:
            recipient = mail.Recipients.Add(spec)
            recipient.Type = recipient_type
            recipient.Resolve()
            if not recipient.Resolved:
                failures.append("{0} ({1})".format(spec, label))
        except Exception as exc:
            log("Could not add {0} '{1}': {2}".format(label, spec, exc))
            failures.append("{0} ({1})".format(spec, label))


def _discard_mail(mail):
    """Throw away an unsaved draft without leaving anything in Drafts."""
    try:
        mail.Close(OL_DISCARD)
    except Exception:
        try:
            mail.Delete()
        except Exception:
            pass


def tool_draft_email(args):
    if not _ALLOW_DRAFTS:
        return ("Email drafting is switched off on this endpoint "
                "(OUTLOOK_ALLOW_DRAFTS=false), so nothing can be put into "
                "Outlook. Give the user the agreed subject and text to copy "
                "into Outlook themselves.")

    mode = (args.get("mode") or "new").strip().lower().replace("-", "_")
    if mode not in EMAIL_DRAFT_MODES:
        return "Error: 'mode' must be one of: {0}.".format(
            ", ".join(EMAIL_DRAFT_MODES))
    entry_id = (args.get("entry_id") or "").strip()
    if mode != "new" and not entry_id:
        return ("Error: mode '{0}' needs 'entry_id' - the EntryID of the "
                "message being answered, from a list/search result.".format(mode))
    if mode == "new" and entry_id:
        return ("Error: 'entry_id' was given with mode 'new'. Use mode "
                "'reply', 'reply_all' or 'forward' to answer that message.")

    raw_body = args.get("body") or ""
    if not raw_body.strip():
        return "Error: 'body' is required - the text agreed in the conversation."
    if len(raw_body) > EMAIL_MAX_BODY_CHARS:
        return "Error: 'body' is longer than {0} characters.".format(
            EMAIL_MAX_BODY_CHARS)
    blocks = markdown_to_email_blocks(raw_body)

    subject = (args.get("subject") or "").strip()
    if mode == "new" and not subject:
        return "Error: 'subject' is required for a new email."

    to_specs = _email_recipient_list(args, "to")
    cc_specs = _email_recipient_list(args, "cc")
    bcc_specs = _email_recipient_list(args, "bcc")

    # The content blacklist applies in BOTH directions, as for a meeting: a
    # marking must not be routed through the AI on its way into Outlook.
    outgoing = "\n".join([subject, raw_body] + to_specs + cc_specs + bcc_specs)
    marking = blacklisted_match(outgoing)
    if marking:
        log("Refused to draft an email (blacklist match: {0}).".format(marking))
        return ("Refused: this email contains content the local content "
                "policy withholds from the AI. Write it in Outlook directly.")

    ns = get_namespace()
    original_subject = ""
    if mode == "new":
        mail = ns.Application.CreateItem(OL_MAIL_ITEM)
    else:
        try:
            original = ns.GetItemFromID(entry_id)
        except Exception:
            return ("Error: no message found for that EntryID (it may have "
                    "moved or been deleted).")
        try:
            is_mail = original.Class == OL_CLASS_MAIL
        except Exception:
            is_mail = False
        if not is_mail:
            return ("Error: that EntryID is not an email, so it cannot be "
                    "replied to or forwarded from here.")
        # A message the AI may not read may not be answered through it either:
        # its reply carries the original underneath.
        reason = email_block_reason(original)
        if reason:
            log("Refused draft_email {0} (blacklist match: {1}) for EntryID "
                "{2}.".format(mode, reason, entry_id))
            return ("Refused: the message being answered is withheld under "
                    "the content blacklist. Reply to it in Outlook directly.")
        try:
            original_subject = original.Subject or ""
        except Exception:
            original_subject = ""
        if mode == "reply":
            mail = original.Reply()
        elif mode == "reply_all":
            mail = original.ReplyAll()
        else:
            mail = original.Forward()

    failures = []
    _add_email_recipients(mail, to_specs, OL_MAIL_TO, "to", failures)
    _add_email_recipients(mail, cc_specs, OL_MAIL_CC, "cc", failures)
    _add_email_recipients(mail, bcc_specs, OL_MAIL_BCC, "bcc", failures)
    if failures:
        # As for a meeting: a half-addressed draft looks finished and goes out
        # without the person it was for.
        _discard_mail(mail)
        return ("Error: Outlook could not resolve {0}. Nothing was put into "
                "Outlook.\nRun outlook_find_people on each one and pass the "
                "'invite_as' value it returns, or use a full email "
                "address.".format(", ".join(failures)))

    if subject:
        mail.Subject = subject

    # Creating the inspector is what makes Outlook add the signature, so it
    # comes BEFORE any text goes in. Display() is modeless: it returns at once
    # and leaves the window open on the user's screen.
    inspector = None
    try:
        inspector = mail.GetInspector
    except Exception as exc:
        log("Could not create the compose window: {0}".format(exc))
    opened = False
    if bool(args.get("open_in_outlook", True)):
        try:
            mail.Display(False)
            opened = True
        except Exception as exc:
            log("Could not open the draft in Outlook: {0}".format(exc))

    try:
        plain = mail.BodyFormat == OL_FORMAT_PLAIN
    except Exception:
        plain = False

    method = None
    if inspector is not None:
        try:
            insert_email_via_word(inspector, blocks, plain)
            method = "word"
        except Exception as exc:
            log("Word editor insert failed, falling back to the message body: "
                "{0}".format(exc))
    if method is None:
        try:
            insert_email_via_body(mail, blocks, plain)
            method = "body"
        except Exception as exc:
            log("Could not write the email body: {0}".format(exc))
            if opened:
                return ("Error: the email opened in Outlook but its text could "
                        "not be written in ({0}). Nothing was saved to Drafts; "
                        "close that window or paste the agreed text into it."
                        .format(exc))
            _discard_mail(mail)
            return ("Error: the text could not be written into Outlook ({0}). "
                    "Nothing was saved.".format(exc))

    saved = True
    try:
        mail.Save()   # Save, never Send. See the note at the top of this part.
    except Exception as exc:
        saved = False
        log("Could not save the draft: {0}".format(exc))

    if opened and inspector is not None:
        try:
            inspector.Activate()     # bring it to the front
        except Exception:
            pass

    def safe(getter):
        try:
            return getter() or ""
        except Exception:
            return ""

    kind = {"new": "New email",
            "reply": "Reply",
            "reply_all": "Reply all",
            "forward": "Forward"}[mode]
    lines = ["Email DRAFT {0} - NOT sent.".format(
        "open in Outlook" if opened else "saved to Drafts"), ""]
    if mode == "new":
        lines.append("  Type     : {0}".format(kind))
    else:
        lines.append("  Type     : {0} to \"{1}\"".format(kind, original_subject))
    lines.append("  Subject  : {0}".format(safe(lambda: mail.Subject)))
    for label, getter in (("To      ", lambda: mail.To),
                          ("Cc      ", lambda: mail.CC),
                          ("Bcc     ", lambda: mail.BCC)):
        value = safe(getter)
        if value:
            lines.append("  {0} : {1}".format(label, value))
    if method == "word":
        lines.append("  Written  : into Outlook's editor above your signature, "
                     "in your own {0} font".format(
                         "compose" if mode == "new" else "reply"))
    else:
        lines.append("  Written  : into the message body (the editor could not "
                     "be reached, so check the font matches your usual one)")
    lines.append("  Format   : {0}".format(
        "plain text (formatting dropped)" if plain else "HTML"))
    if saved:
        lines.append("  Saved    : in Drafts as well, so closing the window "
                     "does not lose it")
    else:
        lines.append("  Saved    : NOT saved to Drafts - Outlook will ask when "
                     "the window is closed")
    lines.append("")
    if not opened:
        lines.append("It could not be opened on screen. Open it from Drafts, "
                     "edit anything you want, then press Send.")
    else:
        lines.append("It is open in Outlook now. Edit anything you want, then "
                     "press Send yourself.")
    if mode == "forward" and not to_specs:
        lines.append("Nobody is addressed yet - add recipients before sending.")
    lines.append("Nothing has gone out: this server cannot send.")
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
            "reports the full path plus what is on the page, including how the sheet "
            "is meant to be folded and printed - repeat that rather than assuming "
            "it. When OUTLOOK_CALENDAR_TEAM is set, the PDF has a second page: up "
            "to four colleagues' calendars for the same day in side-by-side "
            "columns, meant for the back of the sheet, so it must be printed "
            "double-sided, flipped on the short edge. Reading the "
            "calendar is still read-only; every event is on the page, and the planner "
            "never reads an event's body - an event whose subject, location or "
            "category matches the content blacklist is printed but left out of this "
            "tool's reply, which reports how many."
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
                "filename": {
                    "type": "string",
                    "description": (
                        "Optional file name for the PDF. Defaults to "
                        "'Calendar - <date> <weekday>.pdf'; an existing file of the "
                        "same name is overwritten."
                    ),
                },
                "include_team": {
                    "type": "boolean",
                    "description": (
                        "Default true: when OUTLOOK_CALENDAR_TEAM names up to four "
                        "colleagues, a second page puts their calendars for the "
                        "same day side by side, for the back of the sheet "
                        "(double-sided, flip on short edge). Set false for the "
                        "front page only."
                    ),
                },
            },
        },
    },
    {
        "name": "outlook_get_task_team",
        "description": (
            "First step of a printable TEAM TASK SHEET: returns the columns in "
            "print order (column 1 'Team / unassigned', then the user's direct "
            "reports from OUTLOOK_TASKS_TEAM), the Confluence page that collates "
            "their tasks (OUTLOOK_TASKS_PAGE), and today's date from the "
            "endpoint's own clock. Reads configuration only - no mail. Call it "
            "before gathering tasks for outlook_print_team_tasks."
        ),
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "outlook_list_task_emails",
        "description": (
            "The EMAILED half of the team task sheet: emails the user SENT that "
            "sit in their Inbox (an Outlook rule copies sent mail back there) "
            "FLAGGED for follow-up with a due date - the user flags the ones "
            "they want chased and deletes the rest. Returns each one's due "
            "date (the flag's), sent time, To, Cc, subject, a short body "
            "preview and EntryID, soonest due first. Unflagged, completed, "
            "undated and other people's emails are left out and counted. Use "
            "this, not outlook_list_sent_emails, for team tasks."
        ),
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "outlook_print_team_tasks",
        "description": (
            "Print the user's team's open tasks as a PDF: A4 landscape, four "
            "columns a side, double-sided - column 1 is tasks with no single "
            "owner (whole team, several people, or nobody assigned), then one "
            "column per direct report in OUTLOOK_TASKS_TEAM order, with the "
            "printed date in the header. YOU supply the tasks, already gathered "
            "from the Confluence task page and outlook_list_task_emails and "
            "de-duplicated, with done items left out; the server sorts each "
            "column by due date, marks overdue ones, puts a tick box on each, "
            "and reports the path, the per-column counts, any column that ran "
            "out of room, and any name that matched no column. Re-printing on "
            "the same day overwrites that day's sheet. Use the outlook "
            "team-tasks skill for the whole workflow."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "tasks": {
                    "type": "array",
                    "description": "Every open task to print, in any order.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "task": {
                                "type": "string",
                                "description": (
                                    "The task itself, short and actionable "
                                    "(wraps to at most three lines)."
                                ),
                            },
                            "person": {
                                "type": "string",
                                "description": (
                                    "Who owns it, as named in "
                                    "outlook_get_task_team's column list. Empty "
                                    "for the whole team, several people, or "
                                    "nobody assigned - those go in column 1."
                                ),
                            },
                            "due": {
                                "type": "string",
                                "description": (
                                    "Due date as YYYY-MM-DD (sorted, and "
                                    "marked overdue when past). Other text, "
                                    "e.g. 'end of month', prints as written."
                                ),
                            },
                            "status": {
                                "type": "string",
                                "description": (
                                    "Optional short status, e.g. 'In progress', "
                                    "'Blocked'."
                                ),
                            },
                            "source": {
                                "type": "string",
                                "description": (
                                    "Where it came from, short: 'Confluence' or "
                                    "'Email 9 Oct'."
                                ),
                            },
                            "detail": {
                                "type": "string",
                                "description": "Optional few words of context.",
                            },
                        },
                        "required": ["task"],
                    },
                },
                "sources": {
                    "type": "string",
                    "description": (
                        "Optional one-line note for the footer saying what was "
                        "read, e.g. 'Confluence: Team Tasks · flagged emails'."
                    ),
                },
                "filename": {
                    "type": "string",
                    "description": (
                        "Optional file name. Defaults to 'Team Tasks - "
                        "<date>.pdf', which a re-print the same day overwrites; "
                        "a named file that already exists is refused."
                    ),
                },
            },
            "required": ["tasks"],
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
        "name": "outlook_find_people",
        "description": (
            "Fuzzy-find a person or a meeting room in Outlook's directory - the "
            "Global Address List, the All Rooms list, and the user's own Contacts. "
            "Use it FIRST whenever the user names somebody or somewhere in words "
            "('Josh Smith', 'Room R5-3-84', 'someone in Payroll'), because the "
            "scheduling tools need a real address, not a name. Matching tolerates "
            "word order, nicknames and typos: 'Josh Smith' finds 'Smith, Joshua'. "
            "Each match reports an 'invite_as' value - pass THAT to "
            "outlook_suggest_meeting_times and outlook_draft_meeting. Always "
            "confirm which match the user meant before inviting anyone."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": (
                        "The name, room name/code, or part of an email address to "
                        "look for, e.g. 'Josh Smith', 'R5-3-84', 'j.smith@'."
                    ),
                },
                "kind": {
                    "type": "string",
                    "enum": ["any", "person", "room"],
                    "description": (
                        "Narrow the search: 'person' skips the room lists, 'room' "
                        "searches them first. Default 'any'. Use 'room' when the "
                        "user says room, space, boardroom or gives a room code - it "
                        "is both faster and far more accurate."
                    ),
                },
                "count": {"type": "integer", "description": "Maximum matches to return, best first (default 8, max 25)."},
                "deep": {
                    "type": "boolean",
                    "description": (
                        "Default false. When a name resolves unambiguously, Outlook "
                        "answers it directly and no directory scan happens at all. "
                        "Set true to scan the whole directory anyway and list "
                        "alternatives - use it when the user says the match was the "
                        "wrong person."
                    ),
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "outlook_suggest_meeting_times",
        "description": (
            "Suggest times when everybody - and the room - is actually free, from "
            "the free/busy Exchange publishes (the same source as Outlook's "
            "Scheduling Assistant). Use it for 'find a time with X', 'when can we "
            "meet', 'book an hour with the team next week'. Attendees and rooms "
            "must be addresses, so run outlook_find_people first on any name given "
            "in words. The user's own calendar is included automatically. Returns "
            "ranked slots with an exact start time to hand to outlook_draft_meeting; "
            "when nothing is clear it names who is in the way. Anyone whose "
            "free/busy is unpublished is reported, never assumed free."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "attendees": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Required attendees, as the 'invite_as' addresses from "
                        "outlook_find_people. A clash for any of these rules the "
                        "slot out."
                    ),
                },
                "optional_attendees": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Optional attendees. Their clashes never rule a slot out; "
                        "slots where they can also make it are ranked higher."
                    ),
                },
                "rooms": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Meeting rooms or equipment to book. Held to a stricter "
                        "test than people: a tentatively-held room counts as taken."
                    ),
                },
                "duration_minutes": {"type": "integer", "description": "Meeting length in minutes (default 30, 5-1440)."},
                "start_date": {
                    "type": "string",
                    "description": (
                        "First day to consider: 'today' (the default), 'tomorrow', "
                        "an offset like '+7', or YYYY-MM-DD. Prefer the words - the "
                        "server uses the endpoint's own clock."
                    ),
                },
                "days": {"type": "integer", "description": "How many days from start_date to search (default 10; Exchange publishes at most 30)."},
                "earliest_hour": {"type": "integer", "description": "Earliest hour a meeting may start, 0-24 (default 9, or OUTLOOK_MEETING_HOURS)."},
                "latest_hour": {"type": "integer", "description": "Hour by which a meeting must have finished, 0-24 (default 17, or OUTLOOK_MEETING_HOURS)."},
                "include_weekends": {"type": "boolean", "description": "Include Saturday and Sunday (default false)."},
                "include_self": {"type": "boolean", "description": "Check the user's own calendar too (default true). Leave it on unless booking a meeting they are not in."},
                "slot_minutes": {"type": "integer", "description": "Granularity of the candidate start times, 5-120 (default 15)."},
                "max_suggestions": {"type": "integer", "description": "How many slots to list, 1-20 (default 5)."},
            },
        },
    },
    {
        "name": "outlook_draft_meeting",
        "description": (
            "Save a meeting invitation into the user's calendar WITHOUT sending it - "
            "Outlook's 'invitations have not been sent' state - and open it for them "
            "to review, edit and send. It cannot send: nobody is notified until "
            "the user presses Send themselves. Use it once they have chosen a time, with the exact "
            "start and duration from outlook_suggest_meeting_times and the "
            "'invite_as' addresses from outlook_find_people. If any attendee will "
            "not resolve, nothing at all is saved. Never invent attendees, a time, "
            "or a subject: confirm all three with the user first."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "subject": {
                    "type": "string",
                    "description": (
                        "Meeting subject line - the first thing every invitee "
                        "reads. Use the user's own words for it. Where they have "
                        "not given any, name what the meeting is about from the "
                        "conversation (the agenda just drafted, the document just "
                        "written) and put that to them before drafting; never "
                        "invent a subject and never leave it vague."
                    ),
                },
                "start": {
                    "type": "string",
                    "description": (
                        "Start time as 'YYYY-MM-DD HH:MM' in 24-hour local time, "
                        "exactly as outlook_suggest_meeting_times prints it."
                    ),
                },
                "duration_minutes": {"type": "integer", "description": "Length in minutes (default 30, 5-1440). Ignored if 'end' is given."},
                "end": {"type": "string", "description": "Optional end time, 'YYYY-MM-DD HH:MM'. Overrides duration_minutes."},
                "attendees": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Required attendees, as 'invite_as' addresses from outlook_find_people.",
                },
                "optional_attendees": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional attendees, as 'invite_as' addresses.",
                },
                "rooms": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "Rooms/equipment to request, as 'invite_as' addresses. They "
                        "are added as resources and become the Location unless "
                        "'location' says otherwise."
                    ),
                },
                "location": {"type": "string", "description": "Free-text location. Defaults to the room names when rooms are booked."},
                "body": {
                    "type": "string",
                    "description": (
                        "Meeting body - where the agenda goes. When the user has "
                        "just written or agreed an agenda and then asks for a time "
                        "'to discuss it', put THAT agenda here rather than asking "
                        "them to retype it; show them what you are attaching first, "
                        "because everyone invited reads it. Never invent agenda "
                        "items, and leave out your own working notes and any "
                        "options they discarded. Write plain text: Outlook renders "
                        "no Markdown, so headings and bullets arrive as typed "
                        "(the server strips ** and normalises * bullets, and does "
                        "nothing more)."
                    ),
                },
                "reminder_minutes": {"type": "integer", "description": "Reminder lead time in minutes (default 15; 0 for no reminder)."},
                "busy_status": {
                    "type": "string",
                    "enum": ["free", "tentative", "busy", "oof", "elsewhere"],
                    "description": "How the time shows on the calendar (default 'busy').",
                },
                "categories": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Outlook categories to apply, e.g. ['Leadership'].",
                },
                "open_in_outlook": {
                    "type": "boolean",
                    "description": (
                        "Default true: pop the draft open on the user's screen so "
                        "they can edit and send it. Set false to leave it saved in "
                        "the calendar without stealing focus."
                    ),
                },
            },
            "required": ["subject", "start"],
        },
    },
    {
        "name": "outlook_draft_email",
        "description": (
            "Open an UNSENT email in Outlook's own compose window - a new email, a "
            "reply, a reply-all or a forward - with the agreed text written in above "
            "the user's signature, in their own Outlook font, and leave it open for "
            "them to review, edit and send. Also saved to Drafts. It cannot send. "
            "ONLY call this when the user asks for the draft to go into Outlook "
            "('put that in Outlook', 'ok, send that to Outlook', 'open it as a "
            "draft'). Asked to draft, write or reply to an email, write it in the "
            "CHAT and iterate there; do not call this tool on each round. When they "
            "hand it over, pass the latest agreed text exactly, without changing a "
            "word. If any recipient will not resolve, nothing is put into Outlook."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "mode": {
                    "type": "string",
                    "enum": list(EMAIL_DRAFT_MODES),
                    "description": (
                        "'new' (default) for a fresh email; 'reply', 'reply_all' or "
                        "'forward' to answer the message named by 'entry_id'. A "
                        "reply keeps the thread, the quoted original and Outlook's "
                        "reply signature - use it whenever the email answers one "
                        "the user received, rather than a new email."
                    ),
                },
                "entry_id": {
                    "type": "string",
                    "description": "EntryID of the message being answered (reply / reply_all / forward), from a list or search result.",
                },
                "subject": {
                    "type": "string",
                    "description": (
                        "Subject line. Required for a new email. Leave it out on a "
                        "reply or forward to keep Outlook's own 'RE:' / 'FW:' "
                        "subject; set it only if the user changed it."
                    ),
                },
                "body": {
                    "type": "string",
                    "description": (
                        "The email text as agreed in the conversation - the body "
                        "only: no 'Subject:' line, and no sign-off name or "
                        "signature block, because Outlook adds the user's own "
                        "signature. Markdown is converted to Outlook formatting: "
                        "**bold**, *italic*, '- ' bullets, '1. ' numbered lists, "
                        "[text](url) links. One line per line and blank lines "
                        "between paragraphs, exactly as shown to the user."
                    ),
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": (
                        "To recipients: email addresses, or 'invite_as' values from "
                        "outlook_find_people. On a reply these are ADDED to the "
                        "people Outlook already addresses."
                    ),
                },
                "cc": {"type": "array", "items": {"type": "string"}, "description": "Cc recipients, as for 'to'."},
                "bcc": {"type": "array", "items": {"type": "string"}, "description": "Bcc recipients, as for 'to'."},
                "open_in_outlook": {
                    "type": "boolean",
                    "description": (
                        "Default true: open the compose window on the user's "
                        "screen. Set false only if they ask for it to go straight "
                        "to Drafts."
                    ),
                },
            },
            "required": ["body"],
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
    "outlook_get_task_team": tool_get_task_team,
    "outlook_list_task_emails": tool_list_task_emails,
    "outlook_print_team_tasks": tool_print_team_tasks,
    "outlook_list_sent_emails": tool_list_sent_emails,
    "outlook_search_recent": tool_search_recent,
    "outlook_list_folders": tool_list_folders,
    "outlook_find_people": tool_find_people,
    "outlook_suggest_meeting_times": tool_suggest_meeting_times,
    "outlook_draft_meeting": tool_draft_meeting,
    "outlook_draft_email": tool_draft_email,
}

# The two draft tools are the only ones that write to the mailbox, so they are
# removed outright - from the tool list as well as the dispatch table - when
# drafting is switched off. Hiding them is better than refusing them later: a
# model cannot promise the user a draft it was never offered.
DRAFT_TOOLS = ("outlook_draft_meeting", "outlook_draft_email")


def apply_draft_setting():
    """Drop the draft tools when OUTLOOK_ALLOW_DRAFTS turns them off."""
    global TOOLS
    if _ALLOW_DRAFTS:
        return
    TOOLS = [tool for tool in TOOLS if tool["name"] not in DRAFT_TOOLS]
    for name in DRAFT_TOOLS:
        TOOL_DISPATCH.pop(name, None)


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


def report_directory_health():
    """
    Print what the scheduling tools have to work with on this endpoint.

    Three things decide whether "find a time with Josh in R5-3-84" can work,
    and all three fail quietly rather than loudly: the profile may expose no
    Global Address List, it may have no All Rooms list (so rooms are found
    only by name), and Exchange may publish no free/busy (so every suggestion
    would be a guess). This names each one.
    """
    log("")
    log("Directory and free/busy check")
    log("-----------------------------")

    room_lists, other_lists = [], []
    for name, _ in _address_lists():
        if any(hint in name.lower() for hint in ROOM_LIST_HINTS):
            room_lists.append(name)
        else:
            other_lists.append(name)
    if other_lists:
        log("Address lists       : {0}".format(", ".join(other_lists)))
    else:
        log("Address lists       : NONE - outlook_find_people can only search "
            "Contacts, so a colleague will not be found by name.")
    if room_lists:
        log("Room lists          : {0}".format(", ".join(room_lists)))
    else:
        log("Room lists          : none found. Rooms are still findable in the "
            "GAL, but only by how their name reads.")

    log("Meeting hours       : {0:02d}:00-{1:02d}:00".format(
        _MEETING_START_HOUR, _MEETING_END_HOUR))
    log("Directory scan cap  : {0} entries".format(_GAL_SCAN_CAP))
    log("Drafting            : {0}".format(
        "ON - can save an UNSENT meeting or email draft (never send one)"
        if _ALLOW_DRAFTS else "OFF - read-only on the mailbox"))

    try:
        me = get_namespace().CurrentUser
        free_busy = _free_busy(me, datetime.date.today())
    except Exception as exc:
        log("Own free/busy       : could not be read ({0})".format(exc))
        return
    if not free_busy:
        log("Own free/busy       : EMPTY. Exchange published nothing, so "
            "outlook_suggest_meeting_times cannot see anyone's diary and will "
            "say so rather than guess.")
        return
    today_chars = (24 * 60) // FREEBUSY_MIN_PER_CHAR
    booked = sum(1 for ch in free_busy[:today_chars] if ch != FB_FREE)
    log("Own free/busy       : {0} chars at {1} min each; {2} busy today"
        .format(len(free_busy), FREEBUSY_MIN_PER_CHAR,
                "{0:.1f} hour(s)".format(booked * FREEBUSY_MIN_PER_CHAR / 60.0)))


def report_recurrence_health(days=1):
    """
    Print, per recurring series, what each expansion pass found for today.

    This is the diagnostic for "a meeting is in Outlook but not on the page".
    A series that Outlook expands but the probe misses names the fault exactly:
    the probe guessed the wrong time of day, and the line shows both the time it
    probed and the time the occurrence really starts.
    """
    today = datetime.date.today()
    start_dt = datetime.datetime.combine(today, datetime.time(0, 0))
    end_dt = datetime.datetime.combine(
        today + datetime.timedelta(days=max(0, days - 1)),
        datetime.time(23, 59, 59))

    log("")
    log("Recurring-series check, {0} to {1}".format(today, end_dt.date()))

    expansion = {}
    by_outlook = _expanded_occurrences(start_dt, end_dt, expansion)
    log("  Outlook's own expansion : {0} occurrence(s) after walking {1} item(s)"
        .format(expansion.get("found", 0), expansion.get("walked", 0)))
    if expansion.get("error"):
        log("    PROBLEM: {0}".format(expansion["error"]))

    # Index the Outlook pass by start time so each series can be compared.
    outlook_starts = {}
    for occ_start, occ in by_outlook:
        try:
            subject = occ.Subject or "(no subject)"
        except Exception:
            subject = "(unreadable)"
        outlook_starts.setdefault(subject, []).append(occ_start)

    ns = get_namespace()
    items = ns.GetDefaultFolder(OL_FOLDER_CALENDAR).Items
    total = int(items.Count)
    masters = 0
    probed_total = 0
    missed = []

    for index in range(1, total + 1):
        try:
            item = items.Item(index)
            if not bool(item.IsRecurring):
                continue
        except Exception:
            continue
        masters += 1
        report = {}
        occurrences = _expand_recurring_occurrences(item, start_dt, end_dt, report)
        probed_total += len(occurrences)
        try:
            subject = item.Subject or "(no subject)"
        except Exception:
            subject = "(unreadable)"

        # Blacklisted series are named nowhere, check output included.
        if appointment_block_reason(item):
            continue

        outlook_hits = outlook_starts.get(subject, [])
        if outlook_hits and not occurrences:
            missed.append((subject, report.get("probed_times", []),
                           [when.strftime("%H:%M") for when in outlook_hits],
                           report.get("error", "")))

    log("  Recurring series in the folder: {0}".format(masters))
    log("  Found by probing each series  : {0} occurrence(s)".format(probed_total))

    if missed:
        log("")
        log("  {0} series would have been MISSED without Outlook's expansion:"
            .format(len(missed)))
        for subject, probed, actual, error in missed[:20]:
            log("    - {0}".format(subject))
            log("        probed at {0}, actually starts {1}{2}".format(
                ", ".join(probed) or "(none)", ", ".join(actual),
                " [{0}]".format(error) if error else ""))
        if len(missed) > 20:
            log("    ... and {0} more".format(len(missed) - 20))
        log("  Both passes run, so these DO appear on the planner. The gap is")
        log("  what used to hide them: the time of day each series starts at is")
        log("  not the time its first occurrence started at.")
    elif masters:
        log("  Both passes agree - no series is visible to one and not the other.")
    log("")


def report_team_health():
    """
    Print, per OUTLOOK_CALENDAR_TEAM entry, what the back page would show today.

    Says whether each person resolved, whether their calendar opened or only
    their free/busy could be read, and how many items today has - counts
    only, never subjects, since these are other people's diaries.
    """
    today = datetime.date.today()
    log("Planner team view   : {0} calendar(s), back page uses the top {1:.0f}% "
        "of the sheet".format(len(_CALENDAR_TEAM), team_page_fraction() * 100))
    for column, (_label, spec) in zip(
            _team_columns(today, current_user_identity(get_namespace())),
            _CALENDAR_TEAM):
        timed = sum(1 for event in column["events"] if not event["all_day"])
        all_day = len(column["events"]) - timed
        if column["source"] == "unresolved":
            state = "NOT FOUND in the directory - fix this entry"
        elif column["source"] == "unavailable":
            state = "NO ACCESS - neither the calendar nor free/busy could be read"
        elif column["source"] == "free/busy":
            state = ("free/busy only (the calendar itself is not shared with "
                     "you): {0} busy block(s), {1} all-day today".format(
                         timed, all_day))
        else:
            state = "calendar readable: {0} meeting(s), {1} all-day today".format(
                timed, all_day)
        log("  - {0} <{1}>: {2}".format(column["label"], spec, state))


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
        log("Planner page fill   : {0}".format(
            "the whole sheet" if PLANNER_CONTENT_FRACTION >= 1.0
            else "top {0:.0f}%, bottom {1:.0f}% blank to fold up".format(
                PLANNER_CONTENT_FRACTION * 100,
                (1.0 - PLANNER_CONTENT_FRACTION) * 100)))
        if _CATEGORY_COLOURS:
            log("Planner colours     : {0}".format(", ".join(
                "{0}={1}".format(key, value)
                for key, value in sorted(_CATEGORY_COLOURS.items()))))
        if _HIDDEN_CATEGORIES:
            log("Planner hides       : {0}".format(
                ", ".join(sorted(_HIDDEN_CATEGORIES))))
        if _CALENDAR_TEAM:
            try:
                report_team_health()
            except Exception:
                log("Team calendar check could not run:\n{0}".format(
                    traceback.format_exc()))
        else:
            log("Planner team view   : off (OUTLOOK_CALENDAR_TEAM not set)")
        log("Task sheet people   : {0}".format(
            ", ".join(_tasks_column_label(label, spec)
                      for label, spec in _TASKS_TEAM)
            or "not set (OUTLOOK_TASKS_TEAM) - columns come from the task names"))
        log("Task sheet sources  : Confluence page {0}; flagged emails you "
            "sent, in the Inbox, with a due date".format(
                _TASKS_PAGE or "not set (OUTLOOK_TASKS_PAGE)"))
        if KB_DIR:
            log("KB save folder      : {0} ({1})".format(
                KB_DIR,
                "every email read (autosave on)" if KB_AUTOSAVE
                else "on request only"))
        else:
            log("KB save folder      : disabled - no email can be saved")
        try:
            report_directory_health()
        except Exception:
            log("Directory/free-busy check could not run:\n{0}".format(
                traceback.format_exc()))

        try:
            report_recurrence_health()
        except Exception:
            log("Recurring-series check could not run:\n{0}".format(
                traceback.format_exc()))
        log("CHECK OK - Outlook COM link is working.")
        return 0
    except Exception:
        log("CHECK FAILED:\n{0}".format(traceback.format_exc()))
        return 1
    finally:
        pythoncom.CoUninitialize()


def main():
    global KB_DIR, KB_AUTOSAVE, PDF_DIR, _CATEGORY_COLOURS, _HIDDEN_CATEGORIES
    global _CALENDAR_TEAM, _TASKS_TEAM, _TASKS_PAGE
    global CALENDAR_DAY_START_HOUR, CALENDAR_DAY_END_HOUR
    global PLANNER_CONTENT_FRACTION
    global _MEETING_START_HOUR, _MEETING_END_HOUR, _GAL_SCAN_CAP, _ALLOW_DRAFTS

    parser = argparse.ArgumentParser(
        description=(
            "MCP server exposing local Outlook mail and calendar via COM, with a "
            "content blacklist that withholds classified/marked items. It can "
            "never send anything; its only writes to the mailbox are UNSENT "
            "meeting and email drafts (OUTLOOK_ALLOW_DRAFTS=false removes "
            "those). With "
            "no check flag it runs as an stdio MCP server. Configuration is "
            "environment variables only: EVA_KNOWLEDGE_DIR (this server saves "
            "into its 'email' sub-folder), EVA_DOCUMENTS_DIR (it prints the "
            "day planner into the 'pdf' one), OUTLOOK_SEARCH_FOLDERS, "
            "OUTLOOK_MEETING_HOURS, OUTLOOK_BLACKLIST_FILE - see the "
            "CONFIGURATION section of this file's docstring."
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

    # Categories kept off the printed planner: the file's list plus the
    # environment's. Both add - neither can un-hide what the other hides.
    _HIDDEN_CATEGORIES = set(str(name).strip().lower()
                             for name in CALENDAR_HIDDEN_CATEGORIES
                             if str(name).strip())
    _HIDDEN_CATEGORIES |= parse_category_list(
        env("OUTLOOK_CALENDAR_HIDE_CATEGORIES"))

    # Colleagues printed on the back of the planner. The environment REPLACES
    # the file's list rather than adding to it: it is an ordered set of
    # columns, and two lists merged would push someone off the page.
    team_raw = env("OUTLOOK_CALENDAR_TEAM")
    if team_raw is None:
        team_raw = ";".join(str(entry) for entry in CALENDAR_TEAM)
    _CALENDAR_TEAM = parse_team(team_raw)
    if _CALENDAR_TEAM:
        log("Planner team view (back page): {0}.".format(
            ", ".join(label or spec for label, spec in _CALENDAR_TEAM)))

    # The team task sheet: direct reports (replacing the file's list, for the
    # same reason as above - it is an ordered set of columns), the Confluence
    # page that collates their tasks.
    tasks_raw = env("OUTLOOK_TASKS_TEAM")
    if tasks_raw is None:
        tasks_raw = ";".join(str(entry) for entry in TASKS_TEAM)
    _TASKS_TEAM = parse_team(tasks_raw, limit=TASKS_PEOPLE_MAX,
                             what="task sheet people (OUTLOOK_TASKS_TEAM)",
                             page="the task sheet")
    if _TASKS_TEAM:
        log("Task sheet columns: {0}.".format(", ".join(
            _tasks_column_label(label, spec) for label, spec in _TASKS_TEAM)))
    _TASKS_PAGE = env("OUTLOOK_TASKS_PAGE") or _one_line(TASKS_CONFLUENCE_PAGE)

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

    # How much of the sheet's height the planner prints on, e.g.
    # OUTLOOK_CALENDAR_PAGE_FILL="0.8" for a fifth left blank to fold up.
    page_fill = parse_page_fill(env("OUTLOOK_CALENDAR_PAGE_FILL"))
    if page_fill is not None:
        PLANNER_CONTENT_FRACTION = page_fill
    if PLANNER_CONTENT_FRACTION >= 1.0:
        log("Planner uses the whole sheet; nothing is left blank to fold up.")
    else:
        log("Planner uses the top {0:.0f}% of the sheet; the bottom {1:.0f}% "
            "prints blank to fold up.".format(
                PLANNER_CONTENT_FRACTION * 100,
                (1.0 - PLANNER_CONTENT_FRACTION) * 100))

    # Hours a meeting may be SUGGESTED in - separate from the planner's hours
    # above, because what you want to see on a printed day and when it is
    # acceptable to book somebody are not the same window.
    meeting_hours = env("OUTLOOK_MEETING_HOURS")
    if meeting_hours:
        match = re.fullmatch(r"\s*(\d{1,2})\s*[-:to]+\s*(\d{1,2})\s*", meeting_hours)
        if match and 0 <= int(match.group(1)) < int(match.group(2)) <= 24:
            _MEETING_START_HOUR = int(match.group(1))
            _MEETING_END_HOUR = int(match.group(2))
        else:
            log("WARNING: OUTLOOK_MEETING_HOURS='{0}' is not a range like '9-17'; "
                "keeping {1:02d}:00-{2:02d}:00.".format(
                    meeting_hours, _MEETING_START_HOUR, _MEETING_END_HOUR))
    log("Meeting suggestions within {0:02d}:00-{1:02d}:00.".format(
        _MEETING_START_HOUR, _MEETING_END_HOUR))

    # Ceiling on a directory scan. A silly value is worse than the default:
    # too low silently hides people, so a bad one is rejected, not clamped.
    scan_cap = env("OUTLOOK_GAL_SCAN_CAP")
    if scan_cap:
        try:
            parsed = int(scan_cap)
        except ValueError:
            parsed = 0
        if parsed >= 100:
            _GAL_SCAN_CAP = parsed
        else:
            log("WARNING: OUTLOOK_GAL_SCAN_CAP='{0}' is not a number of 100 or "
                "more; keeping {1}.".format(scan_cap, _GAL_SCAN_CAP))

    # The only writes to the mailbox this server can make.
    _ALLOW_DRAFTS = env_flag("OUTLOOK_ALLOW_DRAFTS", ALLOW_DRAFTS)
    apply_draft_setting()
    if _ALLOW_DRAFTS:
        log("Drafting ON: outlook_draft_meeting can SAVE an unsent meeting "
            "into the calendar and outlook_draft_email can open an unsent "
            "email in Outlook. Neither can ever send.")
    else:
        log("Drafting OFF (OUTLOOK_ALLOW_DRAFTS=false): the server is "
            "read-only on the mailbox and neither draft tool is offered.")

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
