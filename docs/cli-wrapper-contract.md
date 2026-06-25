# CLI wrapper contract

Status: 2026-06-24.

The CLI is a convenience and automation surface over the same product
core used by MCP and the web editor. It must not become a separate
canonical document layer.

## Command categories

| Category         | Commands                                                                                  | Core service                                              |
| ---------------- | ----------------------------------------------------------------------------------------- | --------------------------------------------------------- |
| Import/read      | `sessions import`, `sessions projection`, legacy `docx/xlsx/pptx/pdf read/inspect/search` | Session store, projection service, format agents          |
| Command/mutation | `docx/xlsx/pptx/pdf` mutation commands, future `sessions command ...` wrappers            | Command bus, review policy, diagnostics                   |
| Export           | `sessions export`, legacy `--out` wrappers                                                | Export service, export history, command-basis diagnostics |
| Diagnostics      | `doctor`, `list-actions`, command diagnostics in JSON envelopes                           | Runtime doctor, action catalogue, command lifecycle       |
| Maintenance      | `sessions inspect/migrate/cleanup`                                                        | Data-dir schema, migration and temp-artifact cleanup      |

## Boundary rule

Paths are allowed at the shell boundary:

- `--file` imports a local file into a session document.
- `--out` exports a session document or legacy one-shot result.
- `--data-dir` selects the local store for tests or isolated workspaces.

The durable product state is addressed by `sessionId`, `documentId`,
audit/command IDs and export records. New persistent CLI flows should
import/create a document first, operate on `documentId`, then export via
an explicit edge command.

## Output modes

Session CLI commands support:

- human output by default for interactive use;
- `--json` for stable automation envelopes;
- `--pretty` for readable JSON;
- `--quiet` for commands where the caller only needs the exit code.

Human output is not a contract. JSON envelopes are schema-tagged
(`office-ai/...@1`) and should be safe for scripts.

## Error and exit-code contract

| Exit code | Meaning                                         |
| --------- | ----------------------------------------------- |
| `0`       | Success                                         |
| `1`       | Generic runtime error                           |
| `2`       | Missing or unknown document/file-style resource |
| `3`       | OOXML/PDF parse error                           |
| `4`       | Command handler error                           |
| `5`       | Serializer/export error                         |
| `64`      | Usage or boundary validation error              |

Errors go to stderr. Format-specific PDF commands already emit
structured JSON error envelopes. Session commands use typed `CliError`
codes today; the next wrapper slice should align every mutation command
with a structured JSON error envelope when `--json` is passed.

## Parity checks

The contract is covered by tests that:

- import, project and export DOCX/XLSX/PPTX/PDF through
  `office-agent sessions ...`;
- compare MCP `import_document` + `get_document_projection` with CLI
  `sessions import` + `sessions projection` for the same DOCX input;
- snapshot `office-agent sessions --help`;
- assert stable boundary exit codes for unsupported import extension and
  missing `documentId`.

Legacy path commands remain supported for one-shot shell workflows, but
new durable flows should prefer the session CLI.
