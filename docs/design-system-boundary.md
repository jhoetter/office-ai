# Design system boundary

Office-AI is a Sonaloop app surface for Office/PDF artifacts. The shared design
source is `sonaloop-design`, not another local Office-AI design system.

## Allowed design sources

- `sonaloop-design`: canonical app tokens, app CSS preset, shared component
  classes, icon data and command palette primitives.
- `@officeai/design-tokens`: temporary migration shim only. It may remain while
  app/editor CSS aliases converge on `sonaloop-design/styles/tokens.css` and
  `sonaloop-design/app.css`.
- `lucide-react`: temporary migration dependency only. Existing imports are
  being replaced by Sonaloop `IconKey` rendering; new Office-domain icons belong
  in `sonaloop-design/icons.data.mjs`.

## Allowed consumers

Only app/UI shell packages may consume shared design dependencies:

- `@officeai/web`
- `@officeai/ui`
- `@officeai/react-editors`

Headless model/runtime packages (`core`, `docx`, `xlsx`, `pptx`, `pdf*`,
`agent`, `realtime*`) must not import `sonaloop-design`, React design
components, local token packages or Lucide icons.

## Gate

`pnpm architecture` enforces this boundary in `scripts/check-architecture.mjs`:

- no new local `@officeai/design-*` package may be introduced;
- `sonaloop-design` imports are limited to the app/UI shell packages above;
- `@officeai/design-tokens` and `lucide-react` remain limited to the migration
  packages above.

The long-term state is: Office-AI keeps document engines and format-specific UI
logic locally, while shared app language comes from `sonaloop-design`.
