# Provenance citations

Synthesis deliverables can embed source traceability through
`citation_mode` on `create_deliverable_from_synthesis`.

Modes:

- `appendix` (default): visible citation appendix in DOCX, citation rows
  in XLSX, and provenance speaker notes in PPTX.
- `metadata`: keep citation diagnostics/provenance in the MCP response
  without adding visible DOCX appendix content.
- `none`: do not embed or diagnose citation content.

## Citation Sources

OfficeAI derives citations from the neutral synthesis payload:

- finding `evidence`
- quote `source` or `citation`
- asset `uri` or `path`
- table, section and asset ids for traceability

Missing explicit sources emit `citation-source-missing`. Local filesystem
paths are never written into generated customer-facing content; they are
replaced with `[local path omitted]` and reported with
`citation-local-path-omitted`.

## Format Mapping

- DOCX: visible `Citation appendix` paragraphs in `appendix` mode.
- PPTX: `pptx:set-slide-notes` stores provenance notes on generated
  slides.
- XLSX: the `Synthesis` sheet includes `citation` rows and source
  columns.
- PDF: the current PDF template is handoff-only; provenance should be
  carried as annotations on an existing PDF once that PDF is supplied.
