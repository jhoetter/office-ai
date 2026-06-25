# Asset handoff contract

OfficeAI exports are portable assets. `export_document` returns an
`office-ai/asset-handoff@1` envelope so another system can reference the
local file without sharing the OfficeAI session database.

## Export Asset

```json
{
  "schema": "office-ai/asset-handoff@1",
  "assetId": "asset_...",
  "role": "document-export",
  "status": "ready",
  "path": "/workspace/out/report.docx",
  "format": "docx",
  "mediaType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "bytes": 12345,
  "sha256": "...",
  "createdAt": "2026-06-25T00:00:00.000Z",
  "source": {
    "documentId": "doc_...",
    "sessionId": "session_...",
    "documentName": "report.docx",
    "revision": 3
  },
  "commandHistory": {
    "commandIds": ["cmd_..."]
  },
  "diagnostics": []
}
```

`exportHistory` on the canonical document also stores `path`, `bytes`,
`sha256` and `exportedAt`.

## Diagnostics Asset

Pass `diagnostics_out_path` to also write a JSON diagnostics payload:

```json
{
  "schema": "office-ai/export-diagnostics@1",
  "exportAsset": { "...": "..." },
  "diagnostics": []
}
```

The response includes a second `office-ai/asset-handoff@1` envelope with
`role: "export-diagnostics"` and `mediaType: "application/json"`.

## Failure Boundary

If export throws, `export_document` returns a tool error and no asset
envelope is marked ready. If export succeeds with warnings, the asset is
ready and the warnings remain attached in both the response and optional
diagnostics JSON.
