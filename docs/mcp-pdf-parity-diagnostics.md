# MCP PDF parity and diagnostics

PDF is a first-class canonical MCP format. New agent workflows should
open PDFs through `import_document` / `create_document`, then use
`documentId`-based PDF read tools plus the shared command lifecycle.
The older `pdf_load` handle tools and `in_path` / `out_path` mutation
helpers remain for backwards-compatible CLI parity.

## Canonical PDF tools

| Tool                       | Purpose                                                                 |
| -------------------------- | ----------------------------------------------------------------------- |
| `pdf_document_metadata`    | Metadata, page count, engine, PDF version, encryption and signatures.   |
| `pdf_document_page`        | One 1-based page with size, rotation, text-layer flag, text and fields. |
| `pdf_document_outline`     | Recursive outline tree for the canonical `documentId`.                  |
| `pdf_document_annotations` | Flat annotation list, optionally filtered by page.                      |
| `pdf_document_search`      | Text-layer search with hit ranges/rects suitable for anchors.           |
| `pdf_document_diagnostics` | Standalone diagnostic report, optionally with export policy.            |
| `pdf_plan_annotation`      | Plans `pdf:add-annotation`; use `preview_command` / `apply_command`.    |

All of these tools return the same outer pattern as OOXML MCP tools:

```json
{
  "schema": "office-ai/pdf-document-page@1",
  "document": { "schema": "office-ai/document@1", "documentId": "doc_..." },
  "projection": { "schema": "office-agent/pdf-read-page@1" },
  "diagnostics": [{ "level": "info", "code": "pdf-diagnostics-ok", "message": "..." }]
}
```

## Mutation flow

PDF mutations should use the shared command lifecycle:

```text
pdf_plan_annotation -> preview_command -> apply_command -> list_pending_changes/export_document
```

`pdf_plan_annotation` is a convenience wrapper over the canonical command
envelope:

```json
{
  "document_id": "doc_...",
  "kind": "highlight",
  "page": 1,
  "bbox": [72, 650, 180, 665],
  "contents": "Review this",
  "policy": { "mode": "pending" }
}
```

It returns `schema: "office-ai/command-plan@1"` and a `commandId`.
No `in_path` or `out_path` is required once the PDF is open in a
session. Pending mutations are visible through `list_pending_changes`
and the web review surface because they are the same command-bus
mutations used by DOCX, XLSX and PPTX.

## PDF anchors

PDF command targets accept:

| Kind          | Required fields                                                      |
| ------------- | -------------------------------------------------------------------- |
| `page`        | `page`                                                               |
| `page_region` | `page` plus `bbox: [x1,y1,x2,y2]` or `rect: { x, y, width, height }` |
| `text_span`   | `page`, `start`, `end` against the page text layer                   |
| `annotation`  | `annotation_id` or `annotationId`                                    |

Invalid anchors produce `error` diagnostics and block preview/apply.
Text-span anchors require a selectable text layer; scanned pages should
go through OCR first.

## Diagnostics

PDF diagnostics use the same `{ level, code, message }` shape as OOXML
command diagnostics. Current PDF-specific codes are:

| Code                         | Meaning                                                                   |
| ---------------------------- | ------------------------------------------------------------------------- |
| `pdf-diagnostics-ok`         | No PDF-specific caveat was detected.                                      |
| `pdf-encrypted`              | Encryption flags are present; read succeeded, but permissions may matter. |
| `pdf-signature-detected`     | Signature fields exist; OfficeAI does not validate signatures.            |
| `pdf-text-layer-missing`     | One or more pages have no selectable text layer.                          |
| `pdf-ocr-needed`             | The selected page or whole document likely needs OCR.                     |
| `pdf-unsupported-annotation` | Unknown annotation subtypes are projected but not semantically editable.  |
| `pdf-export-policy`          | Export path is incremental-compatible or requires full rewrite semantics. |

`export_document` appends PDF diagnostics and always includes
`pdf-export-policy` for PDFs. If pending changes are still unreviewed,
the normal cross-format `unreviewed-pending-export` warning is included
as well.

## Legacy boundary

The old `pdf_*` handle tools are still available:

- `pdf_metadata`, `pdf_read_page`, `pdf_outline`, `pdf_annotations`,
  `pdf_search`, `pdf_export_markdown` read from a `handle`.
- `pdf_rotate_pages`, `pdf_reorder_pages`, `pdf_delete_pages`,
  `pdf_extract_pages`, `pdf_set_metadata`, `pdf_add_watermark`,
  `pdf_add_page_numbers`, `pdf_fill_form`, `pdf_flatten_form`,
  `pdf_reset_form` and `pdf_merge` remain file-in/file-out helpers.

Do not add new PDF product workflows to the legacy path model. New MCP
and web work should expose a canonical `documentId` flow first, then
optionally keep CLI wrappers at the repository edge.
