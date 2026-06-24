# `docs/`

This folder holds the **narrative** documentation: how the project
evolved, what shipped per phase, and what's deferred. The
**contracts** (specs the build is gated against) live in
[`spec/`](../spec).

## Where to start

| If you want to…                                    | Read this                                                                                                                                                                                                                                                                      |
| -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Understand the product contract                    | [`product-contract.md`](product-contract.md), then [`../README.md`](../README.md) for setup                                                                                                                                                                                    |
| Understand the current codebase state              | [`current-state-inventory.md`](current-state-inventory.md) — package map, surface coverage, MCP/CLI/Web gaps, fixtures and fast wins                                                                                                                                           |
| Understand the command lifecycle                   | [`command-lifecycle.md`](command-lifecycle.md) — command envelopes, revision checks, dry-run/preview/apply rules and diagnostics                                                                                                                                               |
| Use the canonical MCP session/document API         | [`mcp-session-document-tools.md`](mcp-session-document-tools.md) — create/import/list/project/export flow for DOCX/XLSX/PPTX/PDF over MCP                                                                                                                                      |
| Use MCP Plan/Preview/Apply mutations               | [`mcp-command-lifecycle-tools.md`](mcp-command-lifecycle-tools.md) — cross-format command envelopes, anchors, diagnostics, pending review and undo                                                                                                                             |
| Understand local session persistence               | [`session-store-data-dir.md`](session-store-data-dir.md) — data-dir resolution, file layout, restart hydration, corruption handling and privacy boundaries                                                                                                                     |
| Understand the fixture matrix                      | [`../fixtures/README.md`](../fixtures/README.md) — machine-readable DOCX/XLSX/PPTX/PDF fixture inventory, policy and validation gate                                                                                                                                           |
| Read the original build brief                      | [`../prompt.md`](../prompt.md)                                                                                                                                                                                                                                                 |
| See where the DOCX track is today                  | [`session-summary.md`](session-summary.md) — phase map (P0 → P3), test counts, what's deferred                                                                                                                                                                                 |
| Read what shipped on DOCX, in order, with caveats  | [`build-log/docx.md`](build-log/docx.md) — chronological per-batch log (P0 decisions → P3.7), each section: what shipped → decisions → caveats                                                                                                                                 |
| See where the XLSX track is today                  | [`session-summary-xlsx.md`](session-summary-xlsx.md) — phase map (P0 → P11), test counts, what's deferred                                                                                                                                                                      |
| Read what shipped on XLSX, in order, with caveats  | [`build-log/xlsx.md`](build-log/xlsx.md) — chronological per-phase log (P0 → P11), same shape as the DOCX log                                                                                                                                                                  |
| Understand how XLSX differs structurally from DOCX | [`architecture-xlsx-deltas.md`](architecture-xlsx-deltas.md) — cross-cutting deltas (renderer, formula engine, style table, selection, sizing, diff vocabulary, drag UX, test pyramid)                                                                                         |
| Understand how DOCX differs structurally from XLSX | [`architecture-docx-deltas.md`](architecture-docx-deltas.md) — cross-cutting deltas (PM bridge, two-clock loop, style cascade, opaque-blob classifier, comments lifecycle, tracked changes, header/footer parts, page chunker, real-world fixtures + LibreOffice CI, XSD gate) |
| Understand the cross-cutting CI / quality gates    | [`build-log/quality-gates.md`](build-log/quality-gates.md) (narrative) — and [`ci-pipeline.md`](ci-pipeline.md) for the operational cheat-sheet (jobs × make targets × skip semantics)                                                                                         |
| Find the typed contract for a DOCX feature         | [`../spec/docx/`](../spec/docx) — start with `ooxml-mapping.md`, then `document-model.md`, then the workstream specs                                                                                                                                                           |
| Find the typed contract for a XLSX feature         | [`../spec/xlsx/`](../spec/xlsx) — start with `analysis.md`, then `agent-commands.md` + `formula-engine.md`                                                                                                                                                                     |
| See research that fed into the roadmap             | [`../spec/docx/eigenpal-synthesis.md`](../spec/docx/eigenpal-synthesis.md) — code-level read of the eigenpal reference editor, source of P3 / P4 candidate items                                                                                                               |
| Skim historical roadmap docs                       | [`roadmap-docx-p1.md`](roadmap-docx-p1.md) — superseded by what shipped, kept for context                                                                                                                                                                                      |

## Conventions

- **Build log entries are append-only and dated.** Each batch (P*N.M*)
  gets one section: heading → date → commit → spec input → what shipped
  → decisions → caveats. Older entries are not rewritten when later
  phases supersede them; the cross-reference lives in the newer
  section's "decisions" or in `session-summary.md`.
- **`docs/` is narrative; `spec/` is contract.** If a statement gates
  what code must do, it belongs in `spec/`. If it explains what we did,
  why, or what's still pending, it belongs in `docs/`.
- **Caveats are first-class.** Every build-log section ends with an
  honest "what's deferred / known issues" block. Don't quietly drop
  these when the gap closes; cross-link to the section that closed it.
