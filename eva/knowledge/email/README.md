# knowledge\email\

Emails, mirrored to Markdown. The `outlook` plugin writes a file here each time
it reads a message in full, so a thread you have already worked through stays
searchable afterwards.

| | |
|---|---|
| **Plugin setting** | `outlook` → knowledge-base folder (`--kb-dir` / `OUTLOOK_KB_DIR`) |
| **Default** | `C:\Eva\knowledge\email` |
| **Filenames** | `Email - <date> - <subject> (<id>).md` |
| **Overwritten** | yes, on re-read |

Written by the server only. Messages stopped by the compliance blacklist are
**never** written here.

## This folder is the sensitive one

Everything else in `knowledge\` mirrors material that was already a document.
This mirrors correspondence — names, salaries, disputes, whatever was in the
inbox — into plain text files that then get embedded and become quotable in
answers. Worth knowing:

- **Reading is what writes.** A file appears because a message was opened, not
  because anything was exported deliberately.
- **The blacklist runs first.** Blocked messages produce no file, so the terms
  file is the control that matters here, not this folder.
- **Deleting a file removes it from answers** only after the next `kb_index`
  run. Delete and reindex together.

If mirroring email is more than you want, clear the `outlook` plugin's
knowledge-base setting: the plugin then touches no local file at all, and the
rest of the suite is unaffected.
