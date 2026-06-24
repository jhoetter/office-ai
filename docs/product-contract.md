# office-ai product contract

office-ai is the MCP-first document engine and web editor for real
DOCX, XLSX, PPTX and PDF artifacts.

Agents use MCP to inspect, plan, apply, diff, review and export document
changes. People use the web editor to open the same files, make or
review changes, and export real Office/PDF files. The CLI remains useful
for terminals and scripts, but it is a wrapper around the same document
engine instead of the primary product contract.

## Canonical pitch

One sentence:

> office-ai is the MCP-first editor and artifact engine for real Office
> and PDF files, with a full web editor for people and the same safe
> document mutation layer for agents.

Short pitch:

> office-ai lets agents and people work on the same DOCX, XLSX, PPTX and
> PDF files without replacing the file format. MCP exposes the headless
> document engine to agents; the web editor exposes the same sessions,
> projections, review state and exports to people.

Long pitch:

> office-ai treats Office and PDF files as durable artifacts, not as
> chat output. It parses real files, preserves unknown OOXML/PDF
> structures where possible, applies changes through a command bus,
> surfaces reviewable diffs, and exports back to real document files.
> MCP is the primary automation interface. The web editor is the primary
> human interface. The CLI is a terminal wrapper for the same
> capabilities.

## Core formats

| Format | Role                                                                              |
| ------ | --------------------------------------------------------------------------------- |
| DOCX   | Reports, letters, contracts, comments, tracked changes, embedded tables/charts.   |
| XLSX   | Workbooks, formulas, sheets, formatting, tables, filters, charts, comments.       |
| PPTX   | Decks, slides, text, shapes, media, charts, speaker notes, transitions.           |
| PDF    | Read projections, metadata, page operations, annotations, forms, OCR text layers. |

## Core terms

| Term            | Meaning                                                                                                              |
| --------------- | -------------------------------------------------------------------------------------------------------------------- |
| Document        | A real imported or newly created DOCX, XLSX, PPTX or PDF artifact.                                                   |
| Session         | A stable working context that owns document handles, projections, diagnostics, pending mutations and export history. |
| Projection      | A read model for a document: markdown, text, JSON, page, range, slide, sheet, metadata or diagnostics.               |
| Command         | The only mutation path. A command has a type, payload, source and review lifecycle.                                  |
| Diff            | A semantic description of what changed between document snapshots.                                                   |
| Review          | Approval, rejection or pending state for agent-authored mutations before export.                                     |
| Export artifact | A real file emitted from a session: DOCX, XLSX, PPTX, PDF or a diagnostic sidecar.                                   |

## Surface contract

| Surface       | Product role      | Required behavior                                                                            |
| ------------- | ----------------- | -------------------------------------------------------------------------------------------- |
| MCP           | Primary agent API | Stable tools for sessions, documents, projections, commands, review, diagnostics and export. |
| Web editor    | Primary human API | Import/create files, edit per format, inspect agent changes, review diffs and export.        |
| CLI           | Terminal wrapper  | Script-friendly commands over the same service and command contracts.                        |
| React editors | Embedding surface | Optional editor components for hosts that want office-ai inside another app.                 |

## Non-goals

- office-ai is not a generic file-storage product.
- office-ai is not a document-agnostic chat surface.
- office-ai is not tied to a single embedding host or deployment
  environment.
- office-ai does not hide agent edits. Agent mutations must be commands
  with reviewable state.
- office-ai does not replace OOXML or PDF with a private canonical file
  format.

## Immediate product gaps

- Session and document IDs are now the canonical MCP entry point for
  create/import/project/export flows, but the backing registry is still
  in-process until the local session-store/data-dir workstream lands.
- Action metadata now separates `agentCallable`, `webCallable`, and
  `cliCallable`.
- MCP now exposes canonical Plan/Preview/Apply tools for structured
  command envelopes across DOCX, XLSX, PPTX and PDF, including
  diagnostics, pending review, approve/reject and undo. The backing
  registry is still in-process.
- The web editor has strong per-format editors, but no unified session
  browser or pending-changes panel.
- PDF has real packages and tools, but older docs under-described it
  compared with DOCX/XLSX/PPTX.
