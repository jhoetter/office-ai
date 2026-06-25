# ADR: office-ai is a Sonaloop app surface

Status: accepted

Date: 2026-06-25

## Context

office-ai started as a standalone MCP-first document engine and web
editor for DOCX, XLSX, PPTX and PDF. The next roadmap phase changes the
product framing: office-ai should feel like a Sonaloop app, not like a
foreign Office clone with Sonaloop branding applied later.

The constraint remains important: the document engines, command bus,
MCP tools, OOXML/PDF preservation and local app install path must stay
usable without `sonaloop-cloud`.

## Decision

office-ai is a **Sonaloop app surface for Office/PDF artifacts**.

The shared app language comes from `sonaloop-design`: tokens, app-shell
CSS, icons, command palette and semantic component classes. office-ai
keeps its own document engines and format-specific editor surfaces, but
the surrounding shell, review/inspector patterns, empty states,
dialogs, command entry and icon vocabulary converge on Sonaloop.

## Consequences

- `@officeai/core`, format packages and MCP code must not depend on
  `sonaloop-cloud`.
- `@officeai/ui` becomes a thin adapter layer for document-editor
  primitives and temporary compatibility, not a second design system.
- `@officeai/design-tokens` is a migration shim only. Token values must
  converge on `sonaloop-design` and eventually become re-exports or
  aliases.
- `lucide-react` is not a product dependency. Office-AI call sites use
  `@officeai/ui/sonaloop-icons`, backed by `sonaloop-design`; new
  Office-domain icon geometry belongs in `sonaloop-design`.
- CMD+K should use the `sonaloop-design` command palette with an
  Office-AI catalogue adapter.
- The app shell may be local/offline, but its structure should match
  Sonaloop app surfaces: topbar, workspace navigation, panels,
  inspector, review/export dialogs and semantic status treatments.
- Embedding remains native React, not iframe-based.

## Boundaries

The following stay independent:

- `@officeai/core`
- `@officeai/docx`, `@officeai/xlsx`, `@officeai/pptx`, `@officeai/pdf`
- `@officeai/agent` MCP and CLI runtime
- local filesystem session storage and install/doctor path

The following are allowed to consume the Sonaloop app layer:

- `apps/web`
- `@officeai/ui`
- `@officeai/react-editors` shell/chrome bundles, when they need shared
  visual language

## Follow-up tickets

- `design-system-gap-audit`
- `sonaloop-office-icon-taxonomy`
- `sonaloop-icon-set-completion`
- `office-ai-design-system-dependency-boundary`
- `sonaloop-design-tailwind-app-preset`
- `converge-design-tokens`
- `adopt-sonaloop-icon-layer`
- `adopt-sonaloop-command-palette`
- `office-ai-sonaloop-app-shell`
- `sonaloop-app-visual-parity-gate`
