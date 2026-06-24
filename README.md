# office-ai

office-ai is an **MCP-first document engine and web editor** for real
**DOCX, XLSX, PPTX and PDF** artifacts. Agents use MCP to inspect, plan,
apply, diff and export document changes; people use the web editor to
open the same files, review pending changes and save real Office/PDF
outputs.

The product contract is deliberately narrow:

- **OOXML and PDF files remain the source of truth.**
- **MCP is the primary agent API.** CLI commands are wrappers around the
  same headless document capabilities, not a separate product path.
- **The web editor is a complete human surface.** It should import,
  create, edit, review and export without requiring terminal access.
- **Every mutation must be visible as a command, diff, review decision
  and exportable artifact.**

See [`docs/product-contract.md`](docs/product-contract.md) for the
canonical terms and boundaries. The implementation sequence lives in
[`docs/release-slices.md`](docs/release-slices.md): each slice must
prove MCP, web review/editing, real file export and verification
together.

## Status

| Format | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ------ | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| DOCX   | active | Spec complete. Parser/serializer with opaque-blob preservation. Agent commands + tracked changes. ProseMirror renderer. CLI + MCP. **Footnotes (Fußnoten)** typed model + insert/edit/delete commands (F1).                                                                                                                                                                                                                                                                             |
| XLSX   | active | Spec complete. SheetJS-backed parser/serializer with opaque-blob preservation. **All 13 P0 commands** + column/row sizing. Sync formula engine (89 functions + GETPIVOTDATA / CUBE\* stubs). **Pivot tables** preserved end-to-end (Phase 1; render in grid is Phase 2). Excel-flavoured `/xlsx-editor`: open .xlsx from disk, multi-cell selection, type-to-edit, click-to-insert-ref, formula autocomplete, styling toolbar, merge / insert / delete, drag-resize headers. CLI + MCP. |
| PPTX   | active | Parser/serializer with opaque-blob preservation. Slide model (text shapes, tables, charts, **media — video/audio**, masters/layouts/themes). Headless agent commands incl. **shape geometry adjustments** (corner radius, etc.) and **animations with surgical timing-tail merge**. SVG renderer with HTML edit overlay. **Slide transitions** in Present mode. Shared text-formatting toolbar with B/I/U/S, font family/size/color/highlight pickers and real text-range selection.    |
| PDF    | active | Headless PDF model, PDF.js-backed read projections, pdf-lib-backed page edits, metadata, forms, annotations, OCR text-layer bridge, PDF viewer, CLI and MCP coverage. PDF is a first-class product format, not a side utility.                                                                                                                                                                                                                                                          |

## Stack

- **TypeScript everywhere.** Headless format engines, MCP server, CLI wrapper and Next.js web editor share one workspace.
- **Editing substrate**: ProseMirror (MIT)
- **OOXML container**: JSZip (MIT)
- **XML**: fast-xml-parser (MIT) with order/attribute preservation
- **MCP**: `@modelcontextprotocol/sdk`
- **CLI wrapper**: commander (MIT)
- **Frontend**: Next.js 15, React 19, Tailwind CSS 4
- **Design system**: `@officeai/design-tokens` + `@officeai/ui` (Notion-like aesthetic, light/dark)
- **Monorepo**: Turborepo + pnpm workspaces

## Layout

```
office-ai/
├── apps/web/                    # Next.js web editor for DOCX, XLSX, PPTX and PDF
├── packages/
│   ├── core/                    # command bus, plugin registry, OOXML utils, model abstractions
│   ├── docx/                    # DOCX parser, model, serializer, agent, ProseMirror renderer
│   ├── xlsx/                    # XLSX parser, model, serializer, agent, formula engine, virtualized grid
│   ├── pptx/                    # PPTX parser, model, serializer, agent, SVG renderer
│   ├── pdf/                     # PDF model, parser facade, command handlers, agent API
│   ├── pdf-edit/                # page-level PDF operations
│   ├── pdf-annotations/         # annotation model and writer
│   ├── pdf-forms/               # AcroForm helpers
│   ├── pdf-engine/              # PDF.js/PDFium read/render abstraction
│   ├── pdf-ocr/                 # OCR text-layer bridge
│   ├── text-formatting/         # shared run-level formatting contract (TextFormat, units, MIXED, providers)
│   ├── agent/                   # office-agent MCP server + CLI wrapper
│   ├── ui/                      # shared React primitives
│   └── design-tokens/           # brand colors, typography, spacing
├── spec/                        # the contract for the build
│   ├── shared/                  # format-agnostic spec
│   ├── docx/                    # DOCX spec
│   ├── xlsx/                    # XLSX spec
│   ├── pptx/                    # PPTX spec
│   ├── pdf/                     # PDF spec
│   └── agent/                   # agent surface spec
├── fixtures/{docx,xlsx,pptx,pdf}/ # real-world + synthetic test files
├── tests/
│   ├── roundtrip/{docx,xlsx,pptx,pdf}/ # parse/edit/export invariants
│   └── agent/{docx,pdf}/        # agent/MCP-adjacent smoke tests
└── docs/                        # product docs, inventory, build logs
```

## Quick start

```bash
make install         # install workspace deps
make doctor          # check local runtime prerequisites and repair hints
make dev             # Next.js editor host on :3100, realtime on :1234
make dev-forwarded   # tunnel-friendly ports :23003 and :21234
make dev-forwarded-fugu  # Fugu ports :63003 and :61234
make test            # roundtrip + agent tests
make cli             # build the office-agent CLI
make verify          # full quality gate (run before pushing)
```

Forwarded profiles need both the web and realtime ports tunnelled, for example:

```bash
ssh -L 23003:127.0.0.1:23003 -L 21234:127.0.0.1:21234 <host>
```

Canonical MCP sessions persist locally under `$XDG_DATA_HOME/office-ai`
or `~/.local/share/office-ai`. Set `OFFICEAI_DATA_DIR` to isolate one
workspace/test run; see
[`docs/session-store-data-dir.md`](docs/session-store-data-dir.md).

## Quality gates

`make verify` runs the same pipeline as CI (`.github/workflows/ci.yml`).
The gate is fail-fast and ordered cheapest → most expensive:

| Step           | Tool                             | What it catches                                                         |
| -------------- | -------------------------------- | ----------------------------------------------------------------------- |
| `format-check` | Prettier                         | Inconsistent formatting (run `make format` to fix)                      |
| `lint`         | ESLint (root flat config + Next) | Unused vars, deep `src/` imports, banned syntax, import boundaries      |
| `actions`      | Action parity scanner            | Missing or stale action catalogue entries across handlers/UI/CLI        |
| `scorecard`    | Surface scorecard check          | Stale generated MCP/Web/CLI parity matrix                               |
| `fixtures`     | Fixture matrix scanner           | Missing or unindexed real/synthetic fixture files                       |
| `architecture` | `scripts/check-architecture.mjs` | Forbidden cross-package deps in `package.json` (separation of concerns) |
| `typecheck`    | `tsc --noEmit` (per package)     | TypeScript errors                                                       |
| `test`         | Vitest                           | Behavioural regressions across every package                            |
| `build`        | Turbo + Next + tsc               | Build / integration regressions                                         |
| `roundtrip`    | Matrix roundtrip gate            | Import/project/export/reimport regressions across DOCX/XLSX/PPTX/PDF    |

The architecture check enforces this dep graph (see
[`scripts/check-architecture.mjs`](scripts/check-architecture.mjs) for
the source of truth):

```
core             ← leaf (no internal deps)
design-tokens    ← leaf
text-formatting  ← leaf
ui               → design-tokens, text-formatting
docx             → core, text-formatting
xlsx             → core, text-formatting
pptx             → core, text-formatting
agent            → core, docx, xlsx, pptx
web              → core, docx, xlsx, pptx, agent, ui, design-tokens, text-formatting
```

Headless packages (`core`, `docx`, `agent`, `design-tokens`) are
additionally banned from importing `react` / `react-dom` / `next` —
both at the manifest level (architecture check) and at the import
level (ESLint `no-restricted-imports`).

## MCP and CLI

After `make cli`:

```bash
# MCP server
pnpm --filter @officeai/agent exec office-agent mcp

# DOCX via CLI wrapper
pnpm --filter @officeai/agent exec office-agent docx read \
  --file fixtures/docx/real-world/01-styled-letter.docx --format markdown
pnpm --filter @officeai/agent exec office-agent docx write \
  --file fixtures/docx/real-world/01-styled-letter.docx \
  --at "paragraph:0" --text "Updated heading"
pnpm --filter @officeai/agent exec office-agent docx comment \
  --file fixtures/docx/real-world/01-styled-letter.docx \
  --range "paragraph:1" --text "Please review" --author "AI Agent"

# XLSX
pnpm --filter @officeai/agent exec office-agent xlsx inspect \
  --file fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx
pnpm --filter @officeai/agent exec office-agent xlsx read \
  --file fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx \
  --sheet Sheet1 --range A1:D10 --format markdown
pnpm --filter @officeai/agent exec office-agent xlsx set-formula \
  --file fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx \
  --sheet Sheet1 --cell B5 --formula "=SUM(B1:B4)" --out updated.xlsx

# PDF
pnpm --filter @officeai/agent exec office-agent pdf read-metadata fixtures/pdf/simple-text-1page.pdf
pnpm --filter @officeai/agent exec office-agent pdf read-page fixtures/pdf/simple-text-1page.pdf --page 1
```

The same document engine is exposed over MCP via `office-agent mcp`,
with `docx_*`, `xlsx_*`, `pptx_*`, `pdf_*` tools plus catalogue-driven
mutation tools where the action metadata is complete.

## Design principles

Pulled from [`prompt.md`](prompt.md) §Architecture Principles:

1. **Headless-first.** Core, parser, serializer, model, command bus run in Node with zero DOM.
2. **Commands are the only mutation path.** The bus is the invariant.
3. **OOXML/PDF are the source of truth.** Opaque blobs preserve everything we don't understand.
4. **Format-agnostic core.** `packages/core` knows nothing about DOCX, XLSX, PPTX or PDF internals.
5. **Fail loudly.** Import errors and roundtrip anomalies surface as structured errors.

## Reading order

1. [`docs/product-contract.md`](docs/product-contract.md) — the product contract
2. [`docs/current-state-inventory.md`](docs/current-state-inventory.md) — current implementation inventory
3. [`spec/shared/`](spec/shared) — what a "document" is in our system
4. [`spec/docx/`](spec/docx) — the DOCX contract (start with `ooxml-mapping.md`)
5. [`spec/xlsx/`](spec/xlsx) — the XLSX contract (start with `analysis.md`,
   then `agent-commands.md` + `formula-engine.md`)
6. [`spec/pptx/`](spec/pptx) and [`spec/pdf/`](spec/pdf) — PPTX/PDF contracts
