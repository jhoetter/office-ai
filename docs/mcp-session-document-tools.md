# MCP session and document tools

office-ai now exposes a cross-format MCP entry point for DOCX, XLSX,
PPTX and PDF. These tools are the preferred API for agents; the older
format-specific tools remain compatibility wrappers and still accept the
new `documentId` as their matching legacy handle.

## Tool flow

1. `create_session`
   - creates an in-process working session.
   - returns `sessionId`, timestamps and next actions.
2. `import_document`
   - reads a `.docx`, `.xlsx`, `.pptx` or `.pdf` from disk.
   - infers format from extension unless `format` is passed.
   - returns a canonical `documentId`, diagnostics and a format summary.
3. `create_document`
   - creates a blank document for one core format.
   - returns a `documentId`; `export_document` requires `out_path` for
     newly created files.
4. `list_sessions` / `list_documents` / `get_document`
   - inspect current in-process sessions, document metadata,
     diagnostics and export history.
5. `get_document_projection`
   - reads `summary`, `markdown`, `json`, `text` or `page` projections.
   - supports format-specific windowing fields such as `sheet`,
     `range`, `slide`, `page`, `max_rows` and `max_cols`.
6. `plan_command` / `preview_command` / `apply_command`
   - mutates canonical documents through a shared command envelope,
     diagnostics, diff and review lifecycle.
   - see [`mcp-command-lifecycle-tools.md`](mcp-command-lifecycle-tools.md).
7. `export_document`
   - writes a real Office/PDF file.
   - defaults to the original import path when available; otherwise pass
     `out_path`.

## Envelope shape

Every canonical document response includes:

```json
{
  "schema": "office-ai/document@1",
  "documentId": "doc_...",
  "sessionId": "session_...",
  "format": "docx",
  "name": "report.docx",
  "status": "ready",
  "revision": 0,
  "diagnostics": [],
  "exportHistory": [],
  "createdAt": "2026-06-24T00:00:00.000Z",
  "updatedAt": "2026-06-24T00:00:00.000Z"
}
```

`documentId` is stable across tool calls in the same MCP server
lifetime and can be passed to legacy tools such as `docx_inspect`,
`xlsx_get_text`, `pptx_save` or `pdf_metadata` as `handle`.

## Example transcript

```text
create_session({ "title": "Q3 report build" })
→ { "sessionId": "session_..." }

import_document({
  "session_id": "session_...",
  "path": "/workspace/input/report.docx"
})
→ { "document": { "documentId": "doc_...", "format": "docx" } }

get_document_projection({
  "document_id": "doc_...",
  "projection": "markdown"
})
→ { "projection": "markdown", "content": "# ..." }

plan_command({
  "document_id": "doc_...",
  "operation": "docx:insert-text",
  "arguments": { "at": { "paragraph": 0 }, "text": "Draft " },
  "policy": { "mode": "auto_apply" }
})
→ { "commandId": "..." }

preview_command({ "command_id": "..." })
→ { "diff": { "changes": [ ... ] }, "diagnostics": [ ... ] }

apply_command({ "command_id": "..." })
→ { "stage": "applied", "mutation": { "status": "approved" } }

export_document({
  "document_id": "doc_...",
  "out_path": "/workspace/output/report.docx"
})
→ { "exported": { "path": "/workspace/output/report.docx", "bytes": 12345 } }
```

## Current boundary

This is the canonical MCP contract, but the backing registry is still
in-process. It gives agents stable IDs across tool calls; restart-safe
storage belongs to the local session-store/data-dir workstream.
