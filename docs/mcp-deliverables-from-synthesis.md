# MCP deliverables from synthesis payloads

`create_deliverable_from_synthesis` turns a neutral structured research
payload into canonical OfficeAI session documents. It is designed for a
Sonaloop handoff, but it has no Sonaloop runtime dependency: any MCP
client can send the same JSON.

## Tool

```text
create_deliverable_from_synthesis
```

Input:

```json
{
  "session_id": "session_...",
  "target_formats": ["docx", "xlsx"],
  "name": "care-team-synthesis",
  "actor_id": "sonaloop-demo",
  "payload": {
    "title": "Care team scheduling synthesis",
    "subtitle": "Generated from a research payload",
    "summary": "Care coordinators want fewer manual reconciliation steps.",
    "sections": [{ "id": "sec-1", "title": "Workflow pressure", "body": "..." }],
    "findings": [{ "id": "finding-1", "title": "Review burden", "text": "...", "evidence": ["council:..."] }],
    "quotes": [{ "id": "quote-1", "speaker": "Shift planner", "source": "council:...", "text": "..." }],
    "tables": [
      {
        "id": "table-1",
        "title": "Priority matrix",
        "columns": ["Need", "Impact"],
        "rows": [["Report", "High"]]
      }
    ],
    "assets": [
      { "id": "asset-1", "title": "Prototype screenshot", "kind": "image", "uri": "sonaloop://asset/..." }
    ]
  }
}
```

`target_formats` defaults to `["docx"]`. Supported formats are:

- `docx`: readable synthesis report.
- `xlsx`: structured companion workbook on a `Synthesis` sheet.

The tool returns:

- `documents`: canonical `office-ai/document@1` envelopes.
- `provenance`: row/paragraph references back to payload items.
- `diagnostics`: standard `{ level, code, message }` entries.
- `nextActions`: `get_document`, `get_document_projection`,
  `list_activity`, `export_document`.

## Review and export flow

```text
create_deliverable_from_synthesis
  -> get_document_projection
  -> list_activity
  -> export_document
  -> import_document
```

Generated commands are recorded in the local session activity log, so
the web session inspector can show the applied changes. Export produces
real `.docx` / `.xlsx` files, which can be reimported through the same
canonical session API.

## Provenance

The payload keeps source IDs outside OfficeAI. OfficeAI preserves those
IDs in the tool response:

```json
{
  "format": "docx",
  "paragraph": 5,
  "sourceKind": "finding",
  "sourceId": "finding-1"
}
```

For XLSX, provenance points at rows in the `Synthesis` sheet. Rich
inline citations and native document footnotes are tracked separately;
this tool establishes the handoff and reviewable artifact first.

## Local smoke

The MCP test `packages/agent/src/mcp.test.ts` covers the demo path:

- create DOCX and XLSX from one synthesis payload.
- read projections.
- verify activity contains generated Office commands.
- export both files.
- reimport both exported files.
