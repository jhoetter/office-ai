# Current state inventory

Date: 2026-06-24

This report records what the repository currently proves through code,
tests and scripts. It is intentionally conservative: if a capability is
only described in historical docs but not visible in package code, tests
or action metadata, it is treated as unproven.

## Summary

office-ai already contains substantial document-engine work across DOCX,
XLSX, PPTX and PDF:

- headless format packages for parsing, serializing, commands and agent
  projections.
- a shared command bus and action catalogue.
- a Next.js web editor with per-format routes for DOCX, XLSX, PPTX and
  PDF.
- an `office-agent` package that exposes CLI commands and an MCP server.
- real fixture sets and broad unit/E2E coverage.

The main product gap is not format substance. The gap is product shape:
MCP is present, but path-loaded handles are still too close to the old
CLI-first model. Action callability and the command lifecycle are now
explicit contracts; MCP Plan/Apply tools and the web review surface still
need to adopt them end-to-end.

## Package map

| Package/app                                  | Product function                                                          | Current evidence                                                                                                               | Gaps                                                                                                                                          |
| -------------------------------------------- | ------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `packages/core`                              | Command bus, plugin registry, OOXML utilities, action types.              | `CommandLite`, mutation/diff types, action catalogue contracts, command lifecycle envelope/tests, architecture gate.           | Lifecycle helpers exist, but MCP/web surfaces do not yet route Plan/Preview/Apply through them uniformly.                                     |
| `packages/docx`                              | DOCX parser/model/serializer, commands, agent API, ProseMirror renderer.  | 60 command handlers, 90 catalogue entries, real-world fixtures, roundtrip tests, web editor route.                             | MCP/session tools still use format-specific path handles. Some command surfaces are intentionally hidden/no-surface.                          |
| `packages/xlsx`                              | XLSX parser/model/serializer, formula engine, commands, virtualized grid. | 65 command handlers, 102 catalogue entries reported by `pnpm actions`, synthetic fixtures, grid/editor E2E tests.              | Real-world XLSX fixture coverage is thinner than DOCX/PPTX. Some image/paste/fill commands have no public surface.                            |
| `packages/pptx`                              | PPTX parser/model/serializer, commands, agent API, SVG/HTML renderer.     | 60 command handlers, 80 catalogue entries, real and synthetic fixtures, web editor/present E2E tests.                          | Toolbar surface is narrower than palette/CLI; MCP action auto-binding is explicit but only applies where payload metadata is catalogue-owned. |
| `packages/pdf`                               | PDF document model, parser facade, command handlers, headless `PdfAgent`. | 14 command handlers, 47 catalogue entries, PDF tests, PDF viewer route.                                                        | PDF MCP parity is mixed: many CLI helpers are hand-written and not command-bus actions.                                                       |
| `packages/pdf-edit`                          | Page-level PDF operations.                                                | Rotate, reorder, delete, split, merge, extract, crop, watermark, page numbers, metadata tests.                                 | Operations are service helpers, not uniformly command-bus-backed.                                                                             |
| `packages/pdf-annotations`                   | Typed annotation model, writer, XFDF/FDF I/O.                             | Writer tests and annotation fixture tests.                                                                                     | Review/diff semantics for annotation changes need a shared command lifecycle.                                                                 |
| `packages/pdf-forms`                         | AcroForm list/fill/reset/flatten.                                         | Forms tests and CLI/MCP projections.                                                                                           | Form edits bypass the same command catalogue shape used by OOXML formats.                                                                     |
| `packages/pdf-engine`                        | PDF.js/PDFium read/render abstraction.                                    | Engine selection tests and PDF viewer integration.                                                                             | Fidelity fallback is optional; diagnostics need to surface chosen engine and limitations uniformly.                                           |
| `packages/pdf-ocr`                           | OCR text-layer bridge.                                                    | Text-layer tests and PDF action entry.                                                                                         | OCR runtime prereqs need `doctor` coverage.                                                                                                   |
| `packages/agent`                             | MCP server and CLI wrapper.                                               | `office-agent mcp`, canonical session/document MCP tools, format subcommands, action-to-CLI and action-to-MCP adapters, tests. | Session/document IDs are canonical in MCP but still in-process rather than restart-durable.                                                   |
| `packages/react-editors`                     | Embeddable React editor surfaces and blank-file builders.                 | Blank builder tests, bundle dry-run, component entry points.                                                                   | Embedding docs still need a generic adapter contract and removal of old host-specific assumptions.                                            |
| `packages/realtime` + `apps/realtime-server` | Yjs-backed realtime substrate.                                            | Command codec tests, identity tests, local server health endpoint.                                                             | Realtime identity/session state is not yet integrated with canonical document sessions.                                                       |
| `apps/web`                                   | Human web editor surface.                                                 | Routes for start page, DOCX, XLSX, PPTX, PDF; 52 E2E specs.                                                                    | No unified session browser, no cross-format pending-changes panel, no shared export history UI.                                               |

## Format and surface inventory

`pnpm actions` on 2026-06-24 reported:

| Format | Command handlers | Catalogue entries | Agent-callable | Web-callable | CLI-callable | UI-dispatched command types | Parity violations |
| ------ | ---------------: | ----------------: | -------------: | -----------: | -----------: | --------------------------: | ----------------: |
| DOCX   |               60 |                90 |             21 |           50 |           57 |                          52 |                 0 |
| XLSX   |               65 |               102 |             43 |           76 |           59 |                          65 |                 0 |
| PPTX   |               60 |                80 |              8 |           34 |           40 |                          44 |                 0 |
| PDF    |               14 |                47 |              0 |           30 |           21 |                          13 |                 0 |

Action catalogue UI surface counts from source inspection:

| Format | CLI | Palette | Toolbar | Context menu | No public surface |
| ------ | --: | ------: | ------: | -----------: | ----------------: |
| DOCX   |  57 |      52 |      24 |           11 |                10 |
| XLSX   |  59 |      85 |      21 |            7 |                 5 |
| PPTX   |  42 |      49 |       4 |            1 |                11 |
| PDF    |  21 |      34 |       0 |            0 |                 5 |

Interpretation:

- The command catalogue is already a strong cross-surface inventory.
- CLI is currently the most complete generated automation surface.
- Palette/web coverage is high for XLSX and decent for DOCX/PPTX, but
  the web product lacks one cross-format review/session frame.
- PDF has meaningful tooling, but many operations are helper-level CLI
  flows rather than command-bus mutations.

## MCP inventory

MCP exists in `packages/agent/src/mcp.ts` and starts through:

```bash
pnpm --filter @officeai/agent exec office-agent mcp
```

Current model:

- canonical session/document tools:
  `create_session`, `list_sessions`, `import_document`,
  `create_document`, `list_documents`, `get_document`,
  `get_document_projection`, `export_document`.
- format-specific load tools create in-process handles:
  `docx_load`, `xlsx_load`, `pptx_load`, `pdf_load`.
- format-specific read/projection tools expose text, JSON, pages,
  ranges, slides, metadata, outline, annotations, forms and search.
- generic apply tools exist per OOXML format, plus action-generated MCP
  mutation tools for catalogue entries that satisfy the adapter rules.
- save tools write back to the original path or an explicit output path.

Gaps against MCP-first:

- The canonical session/document registry is still process-lifetime,
  not a restart-durable local data-dir.
- Auto-generated MCP mutation tools now require explicit
  `agentCallable` metadata plus a catalogue-owned command schema; the
  remaining gap is a canonical session/document envelope around those
  actions.
- MCP responses do not yet share one envelope with document ID,
  diagnostics, export history and next-action hints across all formats.

## CLI inventory

The CLI is implemented in `packages/agent/src/cli.ts` with per-format
modules for XLSX, PPTX and PDF.

Current command groups:

- `office-agent docx ...`
- `office-agent xlsx ...`
- `office-agent pptx ...`
- `office-agent pdf ...`
- `office-agent mcp`
- `office-agent list-actions`

The CLI has two kinds of commands:

- hand-written commands for reads, custom output formats, apply/diff
  flows and PDF helper operations.
- generated mutation subcommands from action catalogue entries that
  set `cliCallable` and provide a catalogue-owned command schema.

Gaps:

- CLI remains a compatibility and power-user surface, but it is no
  longer the metadata source for MCP auto-binding.
- Some top-level DOCX commands remain as backward-compatible shims.
- CLI flows are path-first; session-first flows are not the default.

## Web editor inventory

Current web routes:

- `/` start page with create/open/sample entry points.
- `/editor` DOCX editor.
- `/xlsx-editor` XLSX editor.
- `/pptx-editor` PPTX editor.
- `/pdf-viewer` PDF viewer/editor surface.

Coverage evidence:

- 52 Playwright E2E specs under `apps/web/e2e`.
- Format-specific ribbons, toolbar tests, shortcuts, import/open tests,
  export tests, PDF viewer tests and visual-fixture tests.

Gaps:

- No session browser for recent/imported documents.
- No unified pending-changes panel for agent-authored commands.
- Export history is not a first-class cross-format UI object.
- Web does not yet consume a shared session/document service boundary.

## Fixtures and tests

Fixtures:

- Machine-readable fixture matrix: `fixtures/MATRIX.json`, enforced by
  `pnpm fixtures:check` / `make fixtures-check`.
- DOCX: 5 synthetic and 11 real-world fixtures.
- XLSX: 6 synthetic fixtures.
- PPTX: 11 synthetic and 3 real fixtures.
- PDF: 12 fixtures covering metadata, text, page size/rotation,
  annotations, forms, outline, signatures and larger documents.

Test file inventory:

| Area                  | Test/spec files |
| --------------------- | --------------: |
| `apps/web`            |              52 |
| `packages/docx`       |              44 |
| `packages/xlsx`       |              61 |
| `packages/pptx`       |              47 |
| `packages/agent`      |               5 |
| `packages/realtime`   |               3 |
| PDF packages combined |               9 |
| integration `tests/`  |              17 |

Baseline checks run for this inventory:

```bash
pnpm actions
pnpm fixtures:check
```

The action parity check passed with zero violations. The fixture matrix
check passed with 48 indexed fixtures and at least three simple and
three complex fixtures per core format.

## Biggest technical risks

1. MCP-first contract now exposes sessions/documents as the primary
   abstraction, but the backing store is still in-process.
2. MCP and web surfaces still need to adopt the shared command lifecycle
   end-to-end, including pending review and diagnostics.
3. PDF is not normalized into the same command/review/export lifecycle as
   OOXML formats.
4. Web editor parity is strong per format but weak at the cross-document
   product layer.
5. Several docs and comments still describe specific embedding or release
   environments instead of an optional adapter boundary.
6. Runtime prereqs for PDF rendering/OCR, LibreOffice roundtrip and
   Playwright are not discoverable through a single `doctor` command.

## Fastest wins

1. Add a local session store under a data-dir and move path-loaded handles
   behind `import_document`.
2. Add a web session browser and pending-changes panel backed by the same
   session state.
3. Normalize PDF diagnostics and export envelopes.
4. Replace environment-specific docs with generic adapter language and
   move old release notes into historical context.
5. Add one `doctor` command that reports runtime prereqs for PDF
   rendering/OCR, LibreOffice roundtrip and Playwright.
