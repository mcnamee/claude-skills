# Exemplar emails

Emails **you actually sent** - the ones that sound like you.
[`/email-writer`](../SKILL.md) reads them to learn your voice, then writes new
emails in it.

An exemplar supplies **voice, never content**. No fact, figure, name, date or
sentence is carried across from one.

| | |
|---|---|
| **Formats** | `.md`, `.txt`, `.eml` (`.docx` and `.pdf` work, but see below) |
| **Committed to git?** | no (see [`.gitignore`](.gitignore)) |
| **Where it lands** | `%USERPROFILE%\.claude\skills\email-writer\exemplars\` |

## Naming

**Lead with the intent** - what the email was for. That is what the skill
matches against, because a thank-you and a status update are built differently
even when the same person writes both:

```
Encouragement - Team after go-live.md
Reporting - Monthly programme update.md
Advice - Vendor selection.md
Request - Data for the audit.md
Escalation - Vendor missed the deadline.md
Apology - Late response.md
Decline - Out of scope change.md
Logistics - Steering committee reschedule.md
Introduction - New delivery lead.md
```

Those nine are the intents the skill classifies against. Use the same words and
matching is exact; use something else and it will read the file and classify it
itself, which costs a tool call.

## What belongs here

- **A spread across intents**, not ten of the same kind. Two encouragement
  emails, two reporting emails and two advice emails teach the skill far more
  than six status updates, because what stays constant across different kinds
  is your voice and what changes is the intent.
- **A spread across relationships** too, if you write differently to your
  director than to your team. Note who each one went to in the index.
- Emails **as you sent them**, including the bits you would edit out on
  reflection. The typos and the shortcuts are the fingerprint.
- Anything **safe to keep in a skill folder**. Strip or fake names, client
  details, figures and personal information first. A redacted email teaches
  voice just as well.

Six to ten covers most of what gets asked for. Keep the signature block out of
them: the skill drafts the body, and your mail client adds the rest.

## Index

One line per file. `Who it went to` is what lets the skill match register as
well as intent.

| File | Intent | Who it went to |
|---|---|---|
| _(none yet - add yours here)_ | | |

## Formats

`.md`, `.txt` and `.eml` are read directly and always work. Saving an email as
Markdown, subject line on the first line and body underneath, is the cheapest
and the easiest to redact.

`.docx` and `.pdf` need the [`word`](../../../plugins/word) or
[`pdf-to-md`](../../../plugins/pdf-to-md) server, and both are confined to their
configured folders. This folder sits under `%USERPROFILE%\.claude\skills\`,
usually outside those, so save your exemplars as plain text and skip the
problem.
