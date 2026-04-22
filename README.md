# officeAI

Browser-embeddable, AI-native document editors for **DOCX** and **XLSX**
(PPTX deferred) — built per [`prompt.md`](prompt.md). Every edit (human or
AI) flows through a single command bus, the editor runs **headless-first**,
and the OOXML file is the source of truth.

## Status

| Format | Status | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ------ | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| DOCX   | active | Spec complete. Parser/serializer with opaque-blob preservation. Agent commands + tracked changes. ProseMirror renderer. CLI + MCP. **Footnotes (Fußnoten)** typed model + insert/edit/delete commands (F1).                                                                                                                                                                                                                                                                             |
| XLSX   | active | Spec complete. SheetJS-backed parser/serializer with opaque-blob preservation. **All 13 P0 commands** + column/row sizing. Sync formula engine (89 functions + GETPIVOTDATA / CUBE\* stubs). **Pivot tables** preserved end-to-end (Phase 1; render in grid is Phase 2). Excel-flavoured `/xlsx-editor`: open .xlsx from disk, multi-cell selection, type-to-edit, click-to-insert-ref, formula autocomplete, styling toolbar, merge / insert / delete, drag-resize headers. CLI + MCP. |
| PPTX   | active | Parser/serializer with opaque-blob preservation. Slide model (text shapes, tables, charts, **media — video/audio**, masters/layouts/themes). Headless agent commands incl. **shape geometry adjustments** (corner radius, etc.) and **animations with surgical timing-tail merge**. SVG renderer with HTML edit overlay. **Slide transitions** in Present mode. Shared text-formatting toolbar with B/I/U/S, font family/size/color/highlight pickers and real text-range selection.    |

## Stack

- **TypeScript everywhere.** No backend; the agent runs as a Node CLI and an in-process API.
- **Editing substrate**: ProseMirror (MIT)
- **OOXML container**: JSZip (MIT)
- **XML**: fast-xml-parser (MIT) with order/attribute preservation
- **CLI**: commander (MIT)
- **Frontend**: Next.js 15, React 19, Tailwind CSS 4
- **Design system**: `@officeai/design-tokens` + `@officeai/ui` (Notion-like aesthetic, light/dark)
- **Monorepo**: Turborepo + pnpm workspaces

## Layout

```
office-ai/
├── apps/web/                    # Next.js host: DOCX editor + XLSX editor surfaces
├── packages/
│   ├── core/                    # command bus, plugin registry, OOXML utils, model abstractions
│   ├── docx/                    # DOCX parser, model, serializer, agent, ProseMirror renderer
│   ├── xlsx/                    # XLSX parser, model, serializer, agent, formula engine, virtualized grid
│   ├── pptx/                    # PPTX parser, model, serializer, agent, SVG renderer
│   ├── text-formatting/         # shared run-level formatting contract (TextFormat, units, MIXED, providers)
│   ├── agent/                   # office-agent CLI + MCP server (docx_* + xlsx_* tools)
│   ├── ui/                      # shared React primitives
│   └── design-tokens/           # brand colors, typography, spacing
├── spec/                        # the contract for the build
│   ├── shared/                  # format-agnostic spec
│   ├── docx/                    # DOCX spec
│   ├── xlsx/                    # XLSX spec
│   ├── pptx/                    # PPTX spec (skeleton)
│   └── agent/                   # CLI / programmatic agent spec
├── fixtures/{docx,xlsx}/        # real-world + synthetic test files
├── tests/
│   ├── roundtrip/{docx,xlsx}/   # parse → serialize → re-parse invariants
│   └── agent/{docx,xlsx}/       # CLI + agent-API smoke tests
└── docs/build-log/{docx,xlsx}.md  # decisions, deviations, deferrals
```

## Quick start

```bash
make install         # install workspace deps
make dev             # Next.js editor host on :3000
make test            # roundtrip + agent tests
make cli             # build the office-agent CLI
make verify          # full quality gate (run before pushing)
```

## Quality gates

`make verify` runs the same pipeline as CI (`.github/workflows/ci.yml`).
The gate is fail-fast and ordered cheapest → most expensive:

| Step           | Tool                             | What it catches                                                         |
| -------------- | -------------------------------- | ----------------------------------------------------------------------- |
| `format-check` | Prettier                         | Inconsistent formatting (run `make format` to fix)                      |
| `lint`         | ESLint (root flat config + Next) | Unused vars, deep `src/` imports, banned syntax, import boundaries      |
| `architecture` | `scripts/check-architecture.mjs` | Forbidden cross-package deps in `package.json` (separation of concerns) |
| `typecheck`    | `tsc --noEmit` (per package)     | TypeScript errors                                                       |
| `test`         | Vitest                           | Behavioural regressions across every package                            |
| `build`        | Turbo + Next + tsc               | Build / integration regressions                                         |

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

## CLI

After `make cli`:

```bash
# DOCX
pnpm --filter @officeai/agent exec office-agent docx read --file fixtures/docx/01-letter.docx --format markdown
pnpm --filter @officeai/agent exec office-agent docx write --file fixtures/docx/01-letter.docx \
  --at "paragraph:0" --text "Updated heading"
pnpm --filter @officeai/agent exec office-agent docx comment --file fixtures/docx/01-letter.docx \
  --range "paragraph:1" --text "Please review" --author "AI Agent"

# XLSX
pnpm --filter @officeai/agent exec office-agent xlsx inspect --file fixtures/xlsx/01-basic-grid.xlsx
pnpm --filter @officeai/agent exec office-agent xlsx read --file fixtures/xlsx/01-basic-grid.xlsx \
  --sheet Sheet1 --range A1:D10 --format markdown
pnpm --filter @officeai/agent exec office-agent xlsx set-formula --file fixtures/xlsx/01-basic-grid.xlsx \
  --sheet Sheet1 --cell B5 --formula "=SUM(B1:B4)" --out updated.xlsx
```

The same surface is exposed over MCP via `office-agent mcp`, with
`docx_*` and `xlsx_*` tool families that share one transport.

## Design principles

Pulled from [`prompt.md`](prompt.md) §Architecture Principles:

1. **Headless-first.** Core, parser, serializer, model, command bus run in Node with zero DOM.
2. **Commands are the only mutation path.** The bus is the invariant.
3. **OOXML is the source of truth.** Opaque blobs preserve everything we don't understand.
4. **Format-agnostic core.** `packages/core` knows nothing about DOCX, XLSX, or PPTX.
5. **Fail loudly.** Import errors and roundtrip anomalies surface as structured errors.

## Reading order

1. [`prompt.md`](prompt.md) — the brief
2. [`spec/shared/`](spec/shared) — what a "document" is in our system
3. [`spec/docx/`](spec/docx) — the DOCX contract (start with `ooxml-mapping.md`)
4. [`spec/xlsx/`](spec/xlsx) — the XLSX contract (start with `analysis.md`,
   then `agent-commands.md` + `formula-engine.md`)
5. [`docs/build-log/docx.md`](docs/build-log/docx.md) — what shipped, what was deferred, why (DOCX)
6. [`docs/build-log/xlsx.md`](docs/build-log/xlsx.md) — same, for XLSX
