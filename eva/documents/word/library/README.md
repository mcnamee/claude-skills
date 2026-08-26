# documents\word\library\

The documents you keep — policies, procedures, contracts, the reports you refer
back to. Stable material, as opposed to the working files in
[`..\inbox`](../inbox).

| | |
|---|---|
| **Read by** | the `word` plugin, as part of its documents folder |
| **Lifecycle** | long-lived |
| **Default path** | `C:\Eva\documents\word\library` |

No configuration of its own. Sub-folders are fine — the plugin searches
recursively, so `library\Policies\`, `library\Contracts\2026\` and so on cost
nothing and a bare filename still finds the file.

## Editable, and that is worth remembering

This is not a read-only archive. The `word` plugin can save over a document
here, in place. In practice that is what you want for a policy being revised,
but it means "the library" is not a safe copy of anything.

Two habits cover it:

- **Ask for tracked changes** on anything you might want to reject. The plugin
  writes real Word revision marks, so the document opens in Word with every
  edit reviewable.
- **Keep the authoritative copy elsewhere** — SharePoint, a document management
  system, wherever it already lives. This folder is a working library, not a
  system of record.

## Searchability

A document becomes searchable when it is opened, not when it is filed here —
opening mirrors its text to
[`..\..\..\knowledge\word`](../../../knowledge/word). To make a batch of
documents findable in one go, ask Eva to open them; each one gets mirrored on
the way past.
