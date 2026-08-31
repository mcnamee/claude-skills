# output\

Everything the assistant **creates**. Kept separate from
[`..\documents`](../documents) so that generated files never mix with your
source library — you can always tell what Eva wrote from what you gave it.

| Folder | Plugin | Setting |
|---|---|---|
| [`word\`](word) | `word` | `--output-dir` / `MSWORD_OUTPUT_DIR` |

Only `word` writes here today. The folder is per-plugin for the same reason the
rest of the tree is: when another plugin gains a create-a-file tool, it gets its
own sub-folder rather than everything landing in one pile.

## Nothing else in the suite writes to disk

Worth knowing where files can appear, because it is a short list:

| Written by | Goes to |
|---|---|
| `word`, creating a document | `output\word\` |
| `word`, editing a document | the document itself, in `documents\word\` |
| `outlook`, `word` mirroring what they read, `confluence` saving a page you asked to keep | `..\knowledge\<source>\` |
| `pdf-to-md`, converting | `..\knowledge\pdf\` |
| `knowledge-base`, capturing a note | `..\knowledge\captures\` |
| `knowledge-base`, indexing | `..\index\` |
| `excel`, `jira` | nothing — read-only |
