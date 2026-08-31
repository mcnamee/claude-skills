# knowledge\email\

Emails you asked to keep, saved as Markdown. The `outlook` plugin writes a file
here **only when a request says to save the message** — reading mail to answer a
question saves nothing, so a thread skimmed on the way to an answer never
reaches the index.

| | |
|---|---|
| **Plugin setting** | `outlook` → knowledge-base folder (`--kb-dir` / `OUTLOOK_KB_DIR`) |
| **Default** | `C:\Eva\knowledge\email` |
| **Filenames** | `Email - <date> - <subject> (<id>).md` |
| **Overwritten** | yes, whenever the same message is saved again |

Written by the server only. Messages stopped by the compliance blacklist are
**never** written here.

## This folder is the sensitive one

Everything else in `knowledge\` holds material that was already a document.
This holds correspondence — names, salaries, disputes, whatever was in the
inbox — as plain text files that then get embedded and become quotable in
answers. Worth knowing:

- **Saying so is what writes.** "Summarise that email" leaves nothing here;
  "save that email to the knowledge base" writes one file. Nothing arrives as a
  side effect of reading.
- **The blacklist runs first.** Blocked messages produce no file whatever you
  ask for, so the terms file is still the control that matters most.
- **Deleting a file removes it from answers** only after the next `kb_index`
  run. Delete and reindex together.

`OUTLOOK_KB_AUTOSAVE=true` goes back to saving every message read (how the
plugin behaved before v5.0.0) — think twice before turning it on here of all
folders. To stop the plugin writing anything at all, set its knowledge-base
folder to `off`: it then touches no local file, and the rest of the suite is
unaffected.
