# MCP command lifecycle tools

The canonical MCP mutation path is:

```text
plan_command -> preview_command -> apply_command -> export_document
```

The tools operate on the same canonical `documentId` returned by
`import_document` or `create_document`. Legacy format-specific mutation
tools remain available, but new agent workflows should prefer the
cross-format command lifecycle.

## Tools

### `plan_command`

Creates an `office-ai/command@1` envelope and validates it against the
current document snapshot.

Input can be direct:

```json
{
  "document_id": "doc_...",
  "operation": "xlsx:set-cell-value",
  "arguments": { "sheet": "Sheet1", "ref": "A1", "value": "Ready" },
  "target": {
    "anchor": { "kind": "range", "sheet": "Sheet1", "range": "A1:A1" }
  },
  "policy": { "mode": "auto_apply" }
}
```

or action-catalogue mapped:

```json
{
  "document_id": "doc_...",
  "action_id": "xlsx.set-cell",
  "arguments": { "sheet": "Sheet1", "ref": "A1", "value": "Ready" }
}
```

The returned `commandId` can be passed to `preview_command` and
`apply_command`.

### `preview_command`

Runs the matching command handler against the current snapshot and
returns diagnostics plus a structured diff. It does not mutate the
session.

### `apply_command`

Dispatches the planned command through the format agent. Policy controls
review state:

| Policy mode  | Behavior                                                                 |
| ------------ | ------------------------------------------------------------------------ |
| `dry_run`    | rejected by apply; use `preview_command` instead.                        |
| `pending`    | stages an agent-authored mutation for review.                            |
| `auto_apply` | applies and approves immediately when the command produces pending work. |

### `list_pending_changes`

Lists pending mutations for one canonical document or all canonical
documents. The response includes document metadata, mutation metadata and
the semantic diff.

### `approve_change` / `reject_change`

Approves or rejects a pending mutation by `mutation_id` for a canonical
document. Rejection can include a human-readable `reason` in the MCP
response; the underlying bus records the rejected mutation status.

### `undo_command`

Undoes the most recent approved mutation for one canonical document.

## Command envelope

`plan_command` returns:

```json
{
  "schema": "office-ai/command@1",
  "id": "command-id",
  "format": "docx",
  "operation": "docx:insert-text",
  "arguments": { "at": { "paragraph": 0 }, "text": "Draft " },
  "target": {
    "sessionId": "session_...",
    "documentId": "doc_...",
    "revision": 0,
    "anchor": { "kind": "paragraph", "index": 0 }
  },
  "source": { "surface": "mcp", "actorId": "agent-1" },
  "policy": { "mode": "pending", "requiresReview": true },
  "createdAt": 1782330000000
}
```

`revision` is checked before preview/apply. Stale commands fail before a
handler runs, so no half-valid file is written.

## Anchors

Anchors are advisory targeting metadata on the envelope. The command
payload still owns the format-specific edit details, but anchors let MCP
clients and the web review surface point at affected structures.

| Format | Anchor kind   | Required fields                         |
| ------ | ------------- | --------------------------------------- |
| DOCX   | `paragraph`   | `index`                                 |
| XLSX   | `range`       | `sheet`, `range`                        |
| PPTX   | `slide_shape` | `slideIndex`, optional `shapeId`        |
| PDF    | `page_region` | `page`, `rect: { x, y, width, height }` |

Invalid anchors are returned as `error` diagnostics and block
preview/apply.

## Diagnostics

Every command lifecycle response includes diagnostics. Levels are:

- `info`: successful validation, preview, apply, review or undo state.
- `warning`: reserved for non-blocking fidelity or targeting caveats.
- `error`: blocks preview/apply or reports command rejection.
- `destructive`: reserved for commands that require explicit destructive
  review.

## Current boundary

The lifecycle tools are backed by the in-process MCP session/document
registry. They are stable across tool calls in one server lifetime, but
restart-durable session storage is a separate workstream.
