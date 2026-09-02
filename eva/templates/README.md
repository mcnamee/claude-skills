# templates\

Blank, branded starting files — letterhead, report layout, contract
boilerplate, the deck shell with your title slide. A new document is **created
from** one of these, inheriting its styles, headers, footers, page setup,
layouts and boilerplate; the template itself is never touched.

One folder per plugin, matching [`..\documents`](../documents):

| Folder | Plugin → setting | Formats |
|---|---|---|
| [`word\`](word) | `word` → `--templates-dir` / `MSWORD_TEMPLATES_DIR` | `.docx` |
| [`powerpoint\`](powerpoint) | `powerpoint` → `--templates-dir` / `POWERPOINT_TEMPLATES_DIR` | `.pptx`, `.potx` |

Each folder carries its own README covering what makes a good template for that
format, how to name it and how to ask for it. Start there.

## Read-only, on purpose

A templates folder is a **read-only second root** for its plugin. Files in it
can be listed, opened and used as the base for something new — but **every save
whose target lands inside it is refused**, whether that is a save-as or a
save-in-place on a template opened by mistake. New documents always land in the
matching [`..\documents`](../documents) folder.

That is the whole reason this sits outside `documents\`: the folder stays a
clean set of blanks the assistant can start from and can never edit in place. A
plugin also **refuses to start** if its templates folder is, or contains, its
documents folder — that arrangement would refuse every save, so it fails loudly
rather than silently.

## Nothing here is indexed

The RAG index reads only [`..\knowledge`](../knowledge). A blank template has no
facts in it worth retrieving, and a half-filled one would put boilerplate into
the corpus with the same authority as a policy.

## Not the same thing as an exemplar

A template *is* the output's first draft. An **exemplar** is a finished document
read for guidance that never becomes the output — a board paper you would hand a
new starter and say *"write it like this"*.

Exemplars no longer live in this tree. They belong to the skill that reads them,
in its own `exemplars\` folder, so they travel with the skill when you copy it
to an endpoint:

| Skill | Exemplars folder |
|---|---|
| `/exemplar-writer` | `%USERPROFILE%\.claude\skills\exemplar-writer\exemplars\` |
| `/brief-writer` | `%USERPROFILE%\.claude\skills\brief-writer\exemplars\` |
| `/email-writer` | `%USERPROFILE%\.claude\skills\email-writer\exemplars\` |

A finished board paper goes to `exemplar-writer`'s folder; the empty branded
shell it was written in belongs here.

## Wiring it up

Both plugins default to their folder here, so a stock `C:\Eva` install needs
nothing set. To point one somewhere else, use the flag or the environment
variable above — flag beats environment variable beats the default — or pass
`off` to run with no templates at all.

Confirm it took: each server logs `templates folder (read-only) = ...` at
startup, and `msword_list_documents` / `powerpoint_list_presentations` with
`location: "templates"` lists exactly what is in the folder.

## Why `README.md` and not `CLAUDE.md`

`CLAUDE.md` files are auto-loaded instructions for an agent working *in this
repository*. These folders are consumed on your endpoint, by an assistant doing
your work — so what is written here is documentation for you, not instructions
for a coding agent.
