# Release slices

Date: 2026-06-24

This roadmap turns the existing workspace into a shippable product
without splitting the work into separate "backend first, UI later"
tracks. Every release slice must prove the same product promise end to
end: MCP can drive the document engine, the web editor can inspect or
review the same state, and the result exports as a real OOXML/PDF file.

## Slice principles

- Each slice has a real file demo and ends with an export artifact.
- MCP is the canonical automation surface; CLI remains a wrapper and
  convenience surface.
- The web editor must expose the same session, projection, command and
  review state people need to trust agent-authored changes.
- DOCX, XLSX, PPTX and PDF stay real file formats. Internal projections
  are views, not replacement file formats.
- Optional local tools are allowed to be missing, but the product must
  report that clearly through `office-agent doctor`.

## R0: Product contract and baseline

| Field             | Definition                                                                                                                         |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| User goal         | A new contributor or evaluator understands what office-ai is, what it is not, and what is already proven.                          |
| Technical changes | Product contract, current-state inventory, host-neutral language, action parity inventory, fixture matrix and fast roundtrip gate. |
| Format coverage   | DOCX, XLSX, PPTX and PDF inventory with real/synthetic fixture policy.                                                             |
| Tests and gates   | `pnpm actions`, `pnpm fixtures:check`, generated surface scorecard, matrix roundtrip gate.                                         |
| Demo              | Read the README/product contract, run the inventory gates, and export at least one fixture through the fast roundtrip gate.        |
| Stop criteria     | The repository still reads like a historical integration, or no single doc states the product contract.                            |
| Deferred          | Deep install packaging, Sonaloop deliverable templates and full web editing parity.                                                |

Ticket allocation:

- `product-north-star-and-positioning`
- `current-state-inventory`
- `remove-host-specific-language`
- `fixture-matrix-real-documents`
- `roundtrip-validation-gates`
- `cross-surface-parity-scorecard`
- `roadmap-sequencing-and-release-slices`

## R1: Local MCP-first sessions

| Field             | Definition                                                                                                                                                     |
| ----------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| User goal         | An agent creates or imports a real document into a stable local session, reads projections, exports it, and a person can find that session in the web editor.  |
| Technical changes | Local data dir, session/document IDs, import/create/list/get/projection/export MCP tools, web sessions API/browser, forwarded dev profiles and runtime doctor. |
| Format coverage   | Create/import/export and read projections for DOCX, XLSX, PPTX and PDF.                                                                                        |
| Tests and gates   | Agent MCP/session tests, web API tests, web session-browser E2E, `make doctor`, `make dev-forwarded` and `make dev-forwarded-fugu` smoke.                      |
| Demo              | Start `make dev-forwarded`, import one DOCX/XLSX/PPTX/PDF through MCP, open the session browser, inspect projections and export the artifact.                  |
| Stop criteria     | Any core flow still requires unmanaged file paths as the primary identity, or web cannot inspect the session created through MCP.                              |
| Deferred          | Storage backends beyond local disk, long-term schema migrations, and CLI defaulting to session IDs.                                                            |

Ticket allocation:

- `session-store-data-dir`
- `mcp-session-document-tools`
- `document-projections-api`
- `web-session-browser`
- `web-import-create-export-flow`
- `make-dev-forwarded-profiles`
- `doctor-runtime-prereqs`
- `storage-adapter-boundary`
- `session-migrations-cleanup`

## R2: Command lifecycle and review

| Field             | Definition                                                                                                                                                |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| User goal         | An agent proposes a document mutation, the product previews the semantic impact, the person reviews it, and export reflects only approved changes.        |
| Technical changes | Shared command envelope, plan/preview/apply tools, review policy, pending changes, approve/reject, undo, command provenance and semantic diff vocabulary. |
| Format coverage   | At least one meaningful mutation per DOCX, XLSX, PPTX and PDF through the same lifecycle; richer command coverage follows the action catalogue.           |
| Tests and gates   | Core command lifecycle tests, agent MCP command tests, web pending-change API/E2E, generated action/scorecard checks and targeted roundtrip fixtures.     |
| Demo              | Use MCP to plan and hold a pending change, review it in web, approve or reject it, undo once, then export and reimport the file.                          |
| Stop criteria     | Agent mutations bypass review state, diffs are raw implementation details, or pending decisions cannot be audited.                                        |
| Deferred          | Perfect diff rendering for every format edge case and post-restart replay of every live pending mutation stack.                                           |

Ticket allocation:

- `command-lifecycle-contract`
- `mcp-command-plan-apply-tools`
- `review-policy-and-undo`
- `web-agent-pending-changes-panel`
- `semantic-diff-engine`
- `command-provenance-audit-log`

## R3: Web editor as full product surface

| Field             | Definition                                                                                                                                          |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| User goal         | A person can use the browser as the primary surface for ordinary import, create, edit, review and export work without dropping into the terminal.   |
| Technical changes | Cross-format session shell, file create/import/export flows, visible capability matrix, format-specific editing parity and consistent diagnostics.  |
| Format coverage   | DOCX, XLSX and PPTX expose edit/review/export flows; PDF exposes read/review/annotation/export with clearly marked limitations.                     |
| Tests and gates   | Web E2E for session browser, import/create/export, pending review, parity smoke per format and Playwright screenshots for forwarded profiles.       |
| Demo              | Open the web editor, create or import one file per format, perform one supported edit or review action, export each artifact and verify it reopens. |
| Stop criteria     | The web editor is still a format demo instead of the product shell, or unsupported actions silently disappear instead of showing capability state.  |
| Deferred          | Advanced desktop-office parity, collaborative realtime identity and remote/cloud storage.                                                           |

Ticket allocation:

- `web-format-editing-parity`
- `replace-path-based-flows`
- `cli-surface-wrapper-contract`
- `local-app-installation`

## R4: PDF and fidelity parity

| Field             | Definition                                                                                                                                                          |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| User goal         | PDF behaves like a first-class document type: agents can inspect and annotate it, people can review it, and export preserves untouched content as much as possible. |
| Technical changes | PDF command-bus parity, normalized diagnostics, annotation/export lifecycle, PDF sidecar fidelity checks and clearer opaque-preservation diagnostics for OOXML.     |
| Format coverage   | PDF read, page, metadata, form and annotation operations; OOXML opaque preservation remains gated across DOCX/XLSX/PPTX.                                            |
| Tests and gates   | PDF package tests, MCP PDF parity tests, annotation export/reimport fixtures, roundtrip gate, optional OCR diagnostics through doctor.                              |
| Demo              | Import a PDF, add an annotation or page operation through MCP, review diagnostics in web, export the PDF and reimport metadata/text/annotation projection.          |
| Stop criteria     | PDF remains a helper-only CLI path, annotations bypass review, or unsupported preservation cases fail without diagnostics.                                          |
| Deferred          | Full incremental PDF writer for every signature/encryption case and OCR as a required runtime.                                                                      |

Ticket allocation:

- `mcp-pdf-parity-and-diagnostics`
- `pdf-incremental-annotation-export`
- `ooxml-opaque-preservation`

## R5: Sonaloop deliverables and distribution

| Field             | Definition                                                                                                                                                 |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| User goal         | A Sonaloop research synthesis can become a real report, deck, spreadsheet or PDF deliverable through office-ai, while office-ai remains usable standalone. |
| Technical changes | Deliverable handoff contract, templates, provenance/citations, asset export metadata, installable local app path and examples package.                     |
| Format coverage   | DOCX report, PPTX executive deck, XLSX analysis workbook and PDF annotation/export package.                                                                |
| Tests and gates   | MCP deliverable smoke, template fixture roundtrip, provenance assertions, docs/examples smoke and install/doctor check.                                    |
| Demo              | Feed a synthesis payload into MCP, generate a DOCX report plus PPTX/XLSX companion artifacts, review them in web and export files with provenance.         |
| Stop criteria     | The demo depends on hidden Sonaloop internals, generated files lack provenance, or installation still requires knowing the monorepo topology.              |
| Deferred          | Hosted multi-user product packaging and direct cloud storage adapters.                                                                                     |

Ticket allocation:

- `mcp-deliverable-from-synthesis`
- `report-deck-spreadsheet-templates`
- `asset-handoff-contract`
- `provenance-citations-in-documents`
- `docs-and-examples-package`

## Early urgent ticket map

All urgent tickets are assigned to the first two proving layers:

| Ticket                               | Slice |
| ------------------------------------ | ----- |
| `product-north-star-and-positioning` | R0    |
| `current-state-inventory`            | R0    |
| `remove-host-specific-language`      | R0    |
| `fixture-matrix-real-documents`      | R0    |
| `roundtrip-validation-gates`         | R0    |
| `command-lifecycle-contract`         | R2    |
| `mcp-action-catalog-decouple-cli`    | R1    |
| `mcp-command-plan-apply-tools`       | R2    |
| `mcp-session-document-tools`         | R1    |
| `session-store-data-dir`             | R1    |
| `make-dev-forwarded-profiles`        | R1    |
| `web-agent-pending-changes-panel`    | R2    |

## Release review checklist

Run this checklist before calling a slice done:

1. The slice demo starts from a clean checkout and a real fixture file.
2. MCP can drive the relevant flow without inventing a private side path.
3. Web can inspect, review or edit the same session state.
4. Export emits a real DOCX, XLSX, PPTX or PDF artifact.
5. Reimport or projection verifies the exported artifact.
6. `office-agent doctor --json` explains any missing optional runtime.
7. The scorecard is current and no new cross-surface gap is undocumented.
