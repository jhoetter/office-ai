# Web format parity

Status: 2026-06-25.

The web editor is the human product surface for the same session store,
document projections, review state and exports used by MCP. This matrix
states what is available now and what is intentionally surfaced as a
diagnostic/roadmap item instead of being implied as working.

| Format | Import | Read view | Editable structures | Review and diff | Export | Known limits                                                                                                           |
| ------ | ------ | --------- | ------------------- | --------------- | ------ | ---------------------------------------------------------------------------------------------------------------------- |
| DOCX   | Full   | Full      | Partial             | Partial         | Full   | Tables, styles, headers/footers and some structural commands are not all exposed as manual web controls yet.           |
| XLSX   | Full   | Full      | Partial             | Partial         | Full   | Charts, advanced tables and complete formula-management controls remain roadmap-visible.                               |
| PPTX   | Full   | Full      | Partial             | Partial         | Full   | Complex assets, layouts, charts and animation editing are not fully manual-editable in the web UI.                     |
| PDF    | Full   | Full      | Review-only         | Partial         | Full   | Annotation, highlight and text-layer editing are planned; the web product does not present them as available controls. |

## Product diagnostics

The web session browser and document detail view use the same matrix from
`apps/web/app/lib/sessions/format-parity.ts`.

Each format exposes a diagnostic code:

- `web-parity-docx-partial-edit`
- `web-parity-xlsx-partial-edit`
- `web-parity-pptx-partial-edit`
- `web-parity-pdf-review-only`

Runtime diagnostics still win when a document has concrete import,
command, review or export diagnostics. The parity diagnostic is shown as
the honest baseline when the document itself has no more specific
diagnostic.

## Follow-up tickets

- DOCX: expose table/style/header/footer controls in the session detail
  or format editor instead of requiring catalogue/CLI commands.
- XLSX: expose charts, table metadata and formula-management controls
  with command diagnostics.
- PPTX: expose assets, layout/theme editing, charts and animations as
  reviewable web controls.
- PDF: implement manual annotation/highlight workflows and export
  diagnostics for those annotations.

These are UI/product gaps, not hidden core promises. MCP and CLI may
already expose lower-level commands before the web editor gives humans a
complete manual control for the same structure.
