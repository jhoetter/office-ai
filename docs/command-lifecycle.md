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
- `resolveReviewPolicy`

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
  agent-authored pending mutation immediately only when the review
  policy allows it.
- `resolveReviewPolicy` is the safety gate before envelope creation:
  action-catalogue `requiresReview` entries and destructive operations
  such as delete/remove/reset/flatten/clear are forced to `pending`
  even if a caller requested `auto_apply`.
- Export never silently approves pending changes. If a document still
  has unreviewed pending mutations, `export_document` writes the
  artifact from the working state and emits an
  `unreviewed-pending-export` warning.
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
- `catalog-review-required`
- `destructive-command-review-required`
- `auto-apply-downgraded-to-pending`
- `unreviewed-pending-export`
- `change-review-undone`

## Semantic diffs

Command handlers continue to emit the compact core `DocumentDiff`
shape. MCP review surfaces additionally normalize that raw diff into
`office-ai/semantic-diff@1`:

- `summary`: short text, change count and aggregate review risk.
- `anchors`: format-aware anchors derived from changed paths or node IDs.
- `changes`: per-change kind, summary, risk, anchor and optional
  before/after values from diff metadata.
- `diagnostics`: diff-specific warnings such as
  `semantic-diff-fallback` or `semantic-diff-truncated`.
- `fallback`: true when the command changed revisions but no structured
  per-node diff was available.

`preview_command`, `apply_command`, `undo_command` and
`list_pending_changes` expose `semanticDiff` where a command diff is
available. Persisted pending-change and command-log metadata stores the
semantic diff so the web session inspector can render a safe summary
without exposing raw local paths or oversized internal diff payloads.

## Provenance and activity

Persisted command-log entries use `office-ai/audit-log-entry@1`. New
lifecycle entries record:

- operation, stage, status, source surface and optional actor id;
- command id where an MCP command envelope exists;
- provenance target: session id, document id, expected revision and
  optional anchor;
- a short argument summary capped for log safety;
- diagnostics and semantic diff summary when available;
- export metadata (`exportRef`) for export entries: exported timestamp,
  byte count and the command ids that formed the export basis.

The MCP `list_activity` tool returns a path-free activity feed over this
log. Browser session detail pages consume the same store and show source,
actor, diagnostics, diff summary, target revision and export basis
without exposing local source or export paths.

## Current integration

The lifecycle contract is typed and tested in core. The canonical MCP
surface now exposes `plan_command`, `preview_command`, `apply_command`,
`undo_command`, `list_pending_changes`, `approve_change` and
`reject_change` over the same `CommandEnvelope` shape for DOCX, XLSX,
PPTX and PDF.

Existing format-specific MCP and CLI tools remain compatibility
wrappers; new agent workflows should prefer the canonical lifecycle
tools. The web session browser consumes the same persisted pending
change metadata for human review. Its persisted review surface can
approve/reject pending metadata and undo a review decision back to
pending; live document replay still requires a live pending mutation
stack.
