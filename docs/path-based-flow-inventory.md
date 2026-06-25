# Path-based flow inventory

Status: 2026-06-24.

office-ai keeps local paths at the product boundary. MCP and web flows
should operate on `sessionId`, `documentId`, `commandId`, `artifactId`
and export records. The CLI may continue accepting `--file` and `--out`
for one-shot scripts, but those paths should wrap the same session
store and command services instead of becoming a separate product core.

## Canonical session-first CLI

| Command                                            | Path role                  | Core identity                              | Decision                                                           |
| -------------------------------------------------- | -------------------------- | ------------------------------------------ | ------------------------------------------------------------------ |
| `office-agent sessions create`                     | none                       | `sessionId`                                | Canonical CLI entry point for a workspace/session.                 |
| `office-agent sessions import --file`              | edge input                 | `sessionId`, `documentId`, audit log id    | Canonical CLI import wrapper; equivalent to MCP `import_document`. |
| `office-agent sessions list`                       | none                       | `sessionId`                                | Canonical inspection wrapper over the local data-dir.              |
| `office-agent sessions documents`                  | none                       | `documentId`                               | Canonical document inventory for CLI/web/MCP parity.               |
| `office-agent sessions projection --document-id`   | none                       | `documentId`                               | Canonical read wrapper; no source file path required.              |
| `office-agent sessions export --document-id --out` | edge output                | `documentId`, export history, audit log id | Canonical export wrapper; records command-basis diagnostics.       |
| `office-agent sessions inspect/migrate/cleanup`    | optional `--data-dir` only | data-dir schema version                    | Store maintenance, not document product flow.                      |

## Legacy path wrappers retained

These commands remain useful for quick shell usage and CI fixtures. They
should not be treated as the MCP/web product model.

| Command family                                                 | Current path flags                                   | Decision                                                                                                     |
| -------------------------------------------------------------- | ---------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `office-agent docx create --out`                               | output path only                                     | Keep as convenience; future session-first blank document should mirror MCP `create_document`.                |
| `office-agent docx inspect/read/search`                        | `--file`                                             | Keep as read-only convenience; canonical scripted product flow is `sessions import` + `sessions projection`. |
| `office-agent docx diff --before --after`                      | two input paths                                      | Keep as file comparator; not a mutable document workflow.                                                    |
| `office-agent docx write/style/comment/.../apply`              | `--file`, optional `--out`                           | Migrate gradually to command lifecycle wrappers that accept `--document-id` and produce command IDs.         |
| Catalogue-driven DOCX commands                                 | `--file`, optional `--out` via action-to-CLI adapter | Same decision as DOCX mutations: keep wrappers, move core mutations to session command lifecycle.            |
| `office-agent xlsx ...`                                        | `--file`, optional `--out` on mutation wrappers      | Keep as compatibility; session projection/import/export is now available for canonical cross-surface flows.  |
| `office-agent pptx ...`                                        | `--file`, optional `--out` on mutation wrappers      | Keep as compatibility; session projection/import/export is now available for canonical cross-surface flows.  |
| `office-agent pdf ...`                                         | input/output paths for read/edit/export helpers      | Keep as PDF edge tooling; session projection/import/export is now available for review/export parity.        |
| Top-level legacy shims `read/search/insert-text/comment/apply` | `-i/--input`, `-o/--output`                          | Retain only for backward compatibility; prefer `docx ...` or `sessions ...`.                                 |

`office-agent list-actions --surface cli` is the machine-readable list
of every catalogue action currently exposed as a CLI command. Those
actions inherit the family decision above: catalogue commands may remain
path-friendly wrappers, but any durable product mutation should be
traceable through session/document metadata and a command/audit ID.

## Migration rule

New MCP mutations must not take raw output paths as the primary model.
New web flows must never require a server-local path. New CLI features
should follow this sequence:

1. Import or create a session document.
2. Address reads/mutations by `documentId`.
3. Persist command diagnostics, pending review state and audit entries.
4. Export only through an explicit edge operation.

The old path commands can stay as adapters when they are useful, but the
shared behavior belongs in the session store, command bus, projection
service and import/export service.
