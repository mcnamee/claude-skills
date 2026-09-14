---
name: meeting-scheduler
description: Schedule a meeting through the outlook MCP server - fuzzy-find people and rooms in the enterprise directory, check everyone's free/busy, and save an unsent meeting draft for the user to review and send, carrying across an agenda already written in the conversation. Use when the user wants to find a time with someone, book a meeting or a room, asks "when are Josh and I both free", "set up 30 minutes with the team next week", "find a suitable time with Josh Smith in Room R5-3-84", says "now find a time to discuss this" after drafting an agenda or a document, or asks for a meeting invite to be drafted.
---

# Scheduling a meeting (via the `outlook` MCP server)

Requires the `outlook.py` MCP server, v8.0.0 or later (Windows + classic
Outlook running and logged into an Exchange profile). If the three tools
below are missing, tell the user to wire it in and verify with
`python outlook.py --check`, which reports whether the directory and
free/busy are readable on that endpoint.

## The whole job in four steps

| Step | Tool | Never skip it because |
|---|---|---|
| 1. Turn names into addresses | `outlook_find_people` | The other tools take addresses, not names. "Josh Smith" is not an address |
| 2. Confirm who they meant | *ask the user* | Two Joshes is the normal case, not the edge case |
| 3. Find the times | `outlook_suggest_meeting_times` | Guessing a time from the user's own calendar ignores everyone else's |
| 4. Save the draft | `outlook_draft_meeting` | — |

Worked example — *"Find a suitable time with Josh Smith in Room R5-3-84"*:

1. `outlook_find_people` `{query: "Josh Smith", kind: "person"}` →
   `Smith, Joshua P` with `invite_as: joshua.smith@contoso.com`.
2. `outlook_find_people` `{query: "R5-3-84", kind: "room"}` →
   `Room R5-3-84 (12 seats)` with `invite_as: r5-3-84@contoso.com`.
   **Pass `kind: "room"`** — it searches the room lists first and is both
   faster and far more accurate than a mixed search.
3. One match each → say who and which room you're using, and carry on. More
   than one → ask before going further.
4. `outlook_suggest_meeting_times` `{attendees: ["joshua.smith@contoso.com"],
   rooms: ["r5-3-84@contoso.com"], duration_minutes: 30}`.
5. Put the options to the user as a short numbered list. **Stop there.**
6. They pick one → `outlook_draft_meeting` with that exact `start` and
   `duration_minutes`, plus a confirmed subject and, if there's an agenda in
   play, the `body` to match.

## Rules that matter

- **Never invent an address.** If `outlook_find_people` returns nothing, say
  so and ask how they spell it — do not guess `firstname.lastname@`. A wrong
  address either bounces or invites a stranger.
- **Ask when the directory is ambiguous.** More than one plausible match, or
  a top match below about 0.8, is a question for the user, not a coin toss.
  Show the job title and department: that is usually what tells two people
  with the same name apart.
- **Confirm the subject, don't invent one.** Where the user hasn't named it,
  take it from what the conversation is already about and put that to them —
  see [The subject and the agenda](#the-subject-and-the-agenda). Asking from
  scratch for something you're both looking at is a wasted turn; guessing
  silently is worse.
- **Nothing is sent.** The server cannot send: `outlook_draft_meeting` saves
  the meeting with its invitations unsent and opens it on screen. Always
  close by telling the user it is a draft and that **they** press Send. Never
  say you have "sent", "invited" or "booked" anyone.
- **Do not pre-read the calendar.** `outlook_suggest_meeting_times` already
  includes the user's own diary (`include_self`, default true), so a
  preceding `outlook_get_calendar` is a wasted round trip.

## Reading the suggestions

The reply ranks slots and annotates each one. Pass the annotations on — they
are what makes a choice informed:

| Annotation | What to tell the user |
|---|---|
| `everyone free` | Clean slot |
| `<name> is tentative` | They're pencilled in elsewhere; the slot is offered but may clash |
| `<name> is working elsewhere` | Fine for a call, worth flagging if the room matters |
| `<name> can't make it (optional)` | Say who would miss it, and let them decide |
| `Note: No free/busy published for ...` | **Say this out loud.** Their diary was never checked — external attendees especially. Do not present the slot as confirmed for them |

Judgement calls the tool makes, so you don't have to:

- A **tentative room** counts as taken; a tentative *person* does not.
- An **optional attendee never blocks** a slot, but slots they can make rank
  higher.
- Suggestions are already **spread across days** rather than being the same
  afternoon in 15-minute steps, so offer them as they come.

## The subject and the agenda

The usual shape of this request is two steps, and the second one leans on the
first:

> *"Draft an agenda for the platform migration review."*
> *"Now find a time with Josh Mann to discuss."*

"To discuss" means **that agenda**. Carry it across; don't make them type it
again, and don't start over with "what should the subject be?".

**The subject**, in order of preference:

1. What the user called it — their words win, always.
2. What the conversation is plainly about: the agenda you just drafted, the
   document you just wrote, the email thread you just read. Propose it.
3. Nothing to go on → ask. A meeting called "Catch-up" tells the invitees
   nothing.

**The agenda** goes in `body`:

- Attach what was actually written or agreed. **Never invent items** to pad it
  out — a fabricated agenda item is read as a commitment by everyone invited.
- Leave out your own scaffolding: working notes, options the user rejected,
  your commentary on their draft.
- **Write plain text.** Outlook renders no Markdown in a meeting body, so
  `**Agenda**` arrives with the asterisks showing. The server strips `**` and
  normalises `*` bullets to `-`, and deliberately does nothing more — headings
  and tables are on you.
- Nothing to attach is a perfectly good answer. A one-line body, or none, beats
  a padded one.

**Show both before you draft.** One short confirmation covering the subject,
the time, who's invited, the room and the body text — then draft on a yes:

> Subject: *Platform migration review*. Tuesday 16 Sep, 10:00–10:30, Room
> R5-3-84, with Joshua Smith. Agenda in the body: migration status, outstanding
> risks, go/no-go decision. Draft that?

That single check is what stops a wrong subject or a half-finished agenda
going out to a room full of people — it is cheaper than any of the fixes.

## When nothing fits

The reply lists the closest slots and names the one person in the way. Do not
widen the search silently and do not pick a time anyway. Put the options to
the user:

- search further ahead (`days`, or a later `start_date`);
- widen the hours (`earliest_hour` / `latest_hour` — default 09:00–17:00);
- make the blocker optional (`optional_attendees`);
- shorten the meeting;
- `include_weekends: true` if that is genuinely on the table.

## Drafting

```
outlook_draft_meeting {
  subject: "Data platform catch-up",
  start: "2026-09-16 10:00",        // exactly as the suggestion printed it
  duration_minutes: 30,
  attendees: ["joshua.smith@contoso.com"],
  rooms: ["r5-3-84@contoso.com"],   // becomes the Location automatically
  body: "Quick catch-up on the migration plan."
}
```

- **`start` is strict** — `YYYY-MM-DD HH:MM`, 24-hour. Copy it from the
  suggestion rather than retyping it; `16/09/2026` and `10am Tuesday` are
  both refused, on purpose.
- **Rooms go in `rooms`, not `location`.** That books them as a resource
  *and* fills in the Location. `location` alone is just text and reserves
  nothing.
- **If one attendee won't resolve, nothing is saved at all.** Fix that
  address with `outlook_find_people` and call again.
- Optional extras only if asked: `reminder_minutes` (default 15),
  `busy_status`, `categories`, `end` instead of a duration.
- `open_in_outlook` defaults to true and pops the draft on screen. Pass false
  only if the user asks you not to interrupt them.
- **`body` is where the agenda goes** — carry across one the user has just
  written rather than making them retype it, and show it to them first. What
  must never appear there is anything they *didn't* write: invented agenda
  items, your own working notes, options they rejected. Everyone invited
  reads it.

## Notes

- **Free/busy is what Exchange publishes**, the same source as Outlook's
  Scheduling Assistant. It shows *when* someone is busy, never with what, and
  a private appointment still reads as busy. Do not speculate about what is
  in someone's diary.
- Exchange publishes **30 days** of free/busy. A longer window is clamped and
  the reply says where it stopped.
- **On a very large directory the scan can stop early.** The reply says so.
  Shortening the search does *not* help — the scan runs in the directory's own
  order, not by relevance. Ask the user for the full name or the email address
  instead: Outlook then resolves it directly, without scanning at all.
- **Drafting can be switched off** on an endpoint (`OUTLOOK_ALLOW_DRAFTS=false`),
  in which case `outlook_draft_meeting` is not offered at all. Give the user
  the agreed time and let them create it in Outlook themselves.
- A subject or body caught by the **content blacklist** is refused rather than
  written into Outlook. Don't work around it.
- The server still cannot accept, decline, move, delete or send anything.
  Never promise to chase an RSVP or reschedule on someone's behalf.
