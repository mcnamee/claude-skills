# email-writer

Draft an email **in your voice**, learned from your own sent mail.

It works out what the email is for, matches that intent to your exemplars in
[`exemplars/`](exemplars), reads your voice across all of them, drafts, runs
[`/unslop`](../unslop) so it doesn't read as machine-written, and returns the
subject line and body in the chat.

A standalone skill: no MCP server, no Python, no dependencies, so it works
anywhere including on an airgapped machine. It uses [`/unslop`](../unslop),
which should be installed alongside it.

## Install

Copy this folder into your Claude skills directory. From the root of this repo,
in **PowerShell**:

```powershell
$dest = "$env:USERPROFILE\.claude\skills"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Recurse -Force .\skills\email-writer $dest
Copy-Item -Recurse -Force .\skills\unslop $dest
```

The `exemplars\` folder travels with it, so put your emails in before you copy
and they land on the endpoint with the skill. For one project only, copy it to
`.claude\skills\email-writer\` inside that project instead. Run `/doctor` or
restart Claude Code if it doesn't show up.

See [`skills/README.md`](../README.md) for the general install notes.

## Use

```
/email-writer reply to Sarah about the audit data, we need it by Friday
/email-writer thank the delivery team for the go-live weekend
/email-writer               (with no argument: the thread or material above)
```

It won't ask you what kind of email it is - it reads the request and decides.
It asks only when the substance is missing: who it goes to, what the actual ask
is, a date it can't infer. One round, and it never blocks.

## Voice and intent are two problems

**Voice is constant.** Your greeting, your length, whether you sign off at all,
your habit of using hyphens for asides. It's the same whether you're thanking
someone or escalating.

**Intent varies.** An encouragement email and a reporting email are built
differently, and matching the wrong one gets you something that sounds like you
but does the wrong job.

So the skill classifies the intent first, reads two or three exemplars of that
kind for the moves, then reads one or two of a *different* kind as a control.
What stays constant across all of them is voice, and only voice is carried into
an email of a different kind. That second read is the step that stops a
thank-you note picking up the shape of a status report.

The nine intents it classifies against:

| | | |
|---|---|---|
| Encouragement | Reporting | Advice |
| Request | Escalation | Apology or correction |
| Decline | Logistics | Introduction |

Where two apply, it follows the dominant one's structure and borrows the
secondary one's opening move, and tells you it did.

## Exemplars

[`exemplars/`](exemplars) is the skill's own folder, and the point of it is that
it travels with the skill. Fill it on a machine where you have your sent mail,
copy the folder across, and the skill arrives on the endpoint already sounding
like you.

Name each file with the intent first:

```
Encouragement - Team after go-live.md
Reporting - Monthly programme update.md
Advice - Vendor selection.md
```

A spread across intents is worth more than volume - six emails covering three
kinds teach it far more than twenty status updates, because the contrast is what
separates your voice from the shape of the email. A spread across relationships
matters too, and where the exemplars conflict, the ones written to the same kind
of person win: most people write to their director and to their team more
differently than they write a report and a request.

Save them as `.md`, `.txt` or `.eml`, which read directly. Strip names, client
details and figures first - a redacted email teaches voice just as well. Nothing
in the folder is committed to git except its README.

**An exemplar supplies voice, never content.** No fact, figure, name, date or
sentence is carried across from one.

With an empty folder it says so and writes from whatever style preferences
you've stated, then offers to keep the approved email as your first exemplar.

## What you get back

The subject line and body in the chat, ready to copy, then one line naming the
intent it picked and the exemplars it followed. Flags only when there's
something to flag: a fact left in square brackets, an attachment referenced but
not seen, a commitment the material didn't clearly support.

Then you iterate. `shorter`, `warmer`, `firmer` are adjustments within your
voice, not licence to leave it, and each round returns the whole email again.

## With the other writing skills

[`/unslop`](../unslop) runs automatically as the last step, and the draft is
re-checked against your voice profile afterwards. `/unslop` is written for
writing in general, and its rules lose to a habit found in your exemplars - if
you consistently write a fragment or reuse a phrase it would flatten, it goes
back in.

[`/polish`](../polish) is **not** run by default. It rewrites into Australian
Public Service house style, which is right for a brief and wrong here: it would
overwrite the voice the exemplars just established. Ask for it explicitly if you
want formal correspondence.

For a formal paper going up to an executive rather than an email, that's
[`/brief-writer`](../brief-writer), which works the same way from its own
exemplars folder.
