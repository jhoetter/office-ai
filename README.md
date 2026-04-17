# officeAI

Browser-embeddable, AI-native document editors for **DOCX** (and, deferred,
XLSX/PPTX) — built per [`prompt.md`](prompt.md). Every edit (human or AI)
flows through a single command bus, the editor runs **headless-first**, and
the OOXML file is the source of truth.

## Status

| Format | Status      | Notes                                                |
| ------ | ----------- | ---------------------------------------------------- |
| DOCX   | active      | Spec complete. Parser/serializer with opaque-blob preservation. Six agent commands. ProseMirror renderer. CLI. |
| XLSX   | deferred    | Spec slot reserved; implementation in a follow-up.   |
| PPTX   | deferred    | Spec slot reserved; implementation in a follow-up.   |

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
├── apps/web/                    # Next.js host: the DOCX editor surface
├── packages/
│   ├── core/                    # command bus, plugin registry, OOXML utils, model abstractions
│   ├── docx/                    # DOCX parser, model, serializer, agent, ProseMirror renderer
│   ├── xlsx/                    # deferred
│   ├── pptx/                    # deferred
│   ├── agent/                   # office-agent CLI + programmatic API
│   ├── ui/                      # shared React primitives
│   └── design-tokens/           # brand colors, typography, spacing
├── spec/                        # the contract for the build
│   ├── shared/                  # format-agnostic spec
│   ├── docx/                    # DOCX spec
│   ├── xlsx/                    # XLSX spec (skeleton)
│   ├── pptx/                    # PPTX spec (skeleton)
│   └── agent/                   # CLI / programmatic agent spec
├── fixtures/docx/               # real-world (and synthetic) DOCX test files
├── tests/
│   ├── roundtrip/docx/          # parse → serialize → re-parse invariants
│   └── agent/docx/              # CLI + agent-API smoke tests
└── docs/build-log/docx.md       # decisions, deviations, deferrals
```

## Quick start

```bash
make install         # install workspace deps
make dev             # Next.js editor host on :3000
make test            # roundtrip + agent tests
make cli             # build the office-agent CLI
```

## CLI

After `make cli`:

```bash
pnpm --filter @officeai/agent exec office-agent docx read --file fixtures/docx/01-letter.docx --format markdown
pnpm --filter @officeai/agent exec office-agent docx write --file fixtures/docx/01-letter.docx \
  --at "paragraph:0" --text "Updated heading"
pnpm --filter @officeai/agent exec office-agent docx comment --file fixtures/docx/01-letter.docx \
  --range "paragraph:1" --text "Please review" --author "AI Agent"
```

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
4. [`docs/build-log/docx.md`](docs/build-log/docx.md) — what shipped, what was deferred, why
