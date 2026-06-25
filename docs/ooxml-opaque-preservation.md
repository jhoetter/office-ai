# OOXML opaque preservation

Status: 2026-06-25.

DOCX, XLSX and PPTX packages contain many OPC parts that the editor does
not actively model. The product rule is: unknown parts and unrelated
relationships stay byte-for-byte unless a command explicitly owns and
dirties that exact part. If a command would need to mutate an opaque
structure, it must return a diagnostic instead of silently dropping or
rewriting the bytes.

## Contract

The machine-readable contract lives in
`packages/core/src/ooxml/preservation.ts` and is exported through
`@officeai/core/ooxml`.

| Format | Parsed examples                                                                       | Opaque examples                                                                                                            | Relationship rule                                                                             |
| ------ | ------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| DOCX   | `word/document.xml`, styles, numbering, comments, headers, footers, footnotes, charts | `customXml/**`, `word/embeddings/**`, media unless dirtied, themes, settings unless dirtied, unknown `word/**` extensions  | Dirty only the owning `.rels` part for relationship commands; keep unrelated ids and targets. |
| XLSX   | workbook, worksheets, shared strings, styles, comments, tables, drawings, charts      | `customXml/**`, embeddings, pivot caches/tables unless promoted, slicers, timelines, macros, unknown `xl/**` extensions    | Preserve unrelated workbook, worksheet, drawing and chart relationship entries.               |
| PPTX   | presentation, slides, layouts, masters, themes, notes, charts                         | `customXml/**`, embeddings, media unless dirtied, tags, view props, pres props unless dirtied, unknown `ppt/**` extensions | Preserve unrelated presentation, slide, layout, master, notes and chart relationship entries. |

## Diagnostics

The shared diagnostic codes are:

- `ooxml-opaque-part-preserved` (`info`) - confirms a preserved opaque part.
- `ooxml-opaque-preservation-risk` (`warning`) - a command is near an
  opaque structure and should not imply full editability.
- `ooxml-opaque-mutation-blocked` (`error`) - a command targets opaque
  content that cannot be safely rewritten.

These diagnostics use the same `level/code/message` shape as session and
command diagnostics, so MCP, CLI and Web can surface them consistently.

## Verification

`tests/roundtrip/ooxml-opaque-preservation.test.ts` injects unknown XML
parts, binary embedding parts and unknown relationship entries into real
or matrix-selected DOCX, XLSX and PPTX fixtures. For each format it
asserts:

- no-op roundtrip keeps the injected parts, `[Content_Types].xml` and
  owning `.rels` bytes stable;
- a known mutation dirties only the expected typed part and still keeps
  injected opaque parts plus relationship snapshots stable;
- diagnostic payloads remain snapshot-stable.

Run the focused gate with:

```bash
pnpm --filter @officeai/core test -- src/ooxml/preservation.test.ts
pnpm --filter @officeai/integration-tests test -- roundtrip/ooxml-opaque-preservation.test.ts
```
