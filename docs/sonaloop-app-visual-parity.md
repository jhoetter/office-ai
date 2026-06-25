# Sonaloop app visual parity gate

Office-AI is a Sonaloop app surface, but its document editors remain
format-specific. The visual parity gate checks that the surrounding app
shell, controls and review surfaces use the shared Sonaloop app language
while preserving DOCX/XLSX/PPTX/PDF editor affordances.

## How to run it

```bash
pnpm --filter @officeai/web build
pnpm visual:sonaloop
```

or:

```bash
make e2e-web-visual
```

The gate writes screenshots to:

```text
apps/web/test-results/sonaloop-visual-parity/
```

The `web-e2e` CI job runs the same spec as part of the normal Playwright
suite and uploads that directory with the Playwright report.

## Covered screens

- Home / local session browser.
- Session inspector with pending command review.
- DOCX, XLSX, PPTX and PDF editor shells opened from session bytes.
- CMD+K command palette.
- Export dialog.

## What fails the gate

The Playwright spec asserts the shared semantic layer, not only pixels:

- `sl-app-main`, `sl-app-topbar` and `sl-app-editor__toolbar` on editor
  surfaces.
- `sl-card`, `sl-table`, `sl-badge` and `sl-empty` on workspace surfaces.
- `sl-app-inspector*` on inspector/review panels.
- `sl-cmdk*` on CMD+K.
- `sl-modal*` on the export dialog.

Screenshots are review artifacts rather than committed binary baselines.
This keeps the gate low-churn while still making visual drift visible in
CI artifacts.

## Accepted format-specific differences

- DOCX keeps the page canvas, ruler/page decorations and text-editing
  controls.
- XLSX keeps the spreadsheet grid, formula bar, sheet tabs and cell
  selection chrome.
- PPTX keeps slide thumbnails, slide canvas and shape controls.
- PDF keeps review/navigation tooling and canvas rendering.

Those areas may look like their native document format. The surrounding
navigation, command, status, inspector, modal and empty-state language must
remain Sonaloop-consistent.
