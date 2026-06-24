# Command lifecycle contract

The command lifecycle is the shared mutation contract for MCP, web and
CLI surfaces. It sits above the existing `CommandBus`: surfaces create a
`CommandEnvelope`, validate it against the current document revision,
preview it without side effects, then apply or queue it through the bus.

## Envelope

`@officeai/core` exports the lifecycle types from
`@officeai/core/commands`:

- `CommandEnvelope`
- `CommandTarget`
- `CommandSurface`
- `CommandPolicyMode`
- `CommandDiagnostic`
- `CommandLifecycleResult`

Every envelope carries:

- `id`: stable command id.
- `format`: `docx`, `xlsx`, `pptx` or `pdf`.
- `operation`: bus command type, for example `docx:insert-page-break`.
- `arguments`: payload passed to the handler.
- `target`: `sessionId`, `documentId`, expected `revision` and optional
  anchor details.
- `source`: `mcp`, `web`, `cli` or `internal`, plus optional actor id.
- `policy`: `dry_run`, `auto_apply` or `pending`, with a review flag.
- `createdAt`: timestamp.

## Stages

The explicit stages are:

- `created`
- `validated`
- `previewed`
- `applied`
- `queued`
- `reviewed`
- `exported`
- `failed`

The current helpers implement the first hard gate:

- `createCommandEnvelope(input)`
- `validateCommandEnvelope(envelope, snapshot)`
- `previewCommandEnvelope(envelope, snapshot, handler)`
- `applyCommandEnvelope(bus, envelope)`
- `hasBlockingDiagnostics(diagnostics)`

## Rules

- A stale target revision fails before preview or apply.
- A format mismatch fails before preview or apply.
- Preview runs the handler against the supplied snapshot and returns a
  diff, but does not mutate the bus, history, pending stack or files.
- `policy.mode: "dry_run"` cannot be applied. It must use preview.
- `policy.mode: "pending"` queues via the bus as an agent mutation.
- `policy.mode: "auto_apply"` dispatches through the bus and approves an
  agent-authored pending mutation immediately.
- Handler failures, including invalid anchors surfaced as `CommandError`,
  return structured diagnostics instead of silent partial writes.

## Diagnostics

Diagnostics use four levels:

- `info`
- `warning`
- `error`
- `destructive`

`error` and `destructive` are blocking. Lifecycle helpers currently emit
at least:

- `stale-revision`
- `format-mismatch`
- `dry-run-apply`
- handler `CommandError.code` values such as `invalid-anchor`
- `handler-threw` for non-`CommandError` exceptions

## Current integration

The lifecycle contract is typed and tested in core. The canonical MCP
surface now exposes `plan_command`, `preview_command`, `apply_command`,
`undo_command`, `list_pending_changes`, `approve_change` and
`reject_change` over the same `CommandEnvelope` shape for DOCX, XLSX,
PPTX and PDF.

Existing format-specific MCP and CLI tools remain compatibility
wrappers; new agent workflows should prefer the canonical lifecycle
tools. The remaining integration step is the web pending-changes panel
and session browser consuming the same envelope and review state.
