# `@officeai/react-editors`

Embeddable office-ai editor surfaces and blank-file builders for React
hosts (e.g. [hof-os](https://github.com/jhoetter/hof-os)) that want
office-ai's DOCX / XLSX / PPTX / PDF editing UX without standing up the
`apps/web` Next.js shell.

## Phase 1 — what ships today

- `@officeai/react-editors/blanks` — re-exports the four `*Agent.empty()`
  builders (`makeBlankDocx`, `makeBlankXlsx`, `makeBlankPptx`,
  `makeBlankPdf`) so a host can mint zero-byte starter files **in the
  browser** and upload them via its own storage path. Plus the matching
  per-format subpaths (`./blanks/docx`, `./blanks/xlsx`, …) for tighter
  tree-shaking.
- `@officeai/react-editors/mime` — canonical MIME constants for the four
  formats (`DOCX_MIME`, `XLSX_MIME`, `PPTX_MIME`, `PDF_MIME`) plus
  default filenames (`Untitled.docx` etc).
- `@officeai/react-editors/contract` — TypeScript-only props contract
  the four editor components will satisfy in subsequent releases:
  `EmbeddedEditorProps` documents `initialBytes`, `onSave`, `onClose`,
  `locale`, `theme`. Phase 0 of the in-repo refactor wired those props
  into `apps/web`'s editor implementations; phases 1.5+ extract the
  components themselves so they can be imported from this package.

## Phase 1.5 — what's next

Per the
[plan](../../docs/embedding.md), the four editor `*.tsx` surfaces and
their colocated child components (toolbars, sidebars, dialogs, format
providers) move from `apps/web/app/{editor,xlsx-editor,pptx-editor,
pdf-viewer}/` into `packages/react-editors/src/`, exposed via subpath
exports:

- `@officeai/react-editors/docx` → `DocxEditor`
- `@officeai/react-editors/xlsx` → `XlsxEditor`
- `@officeai/react-editors/pptx` → `PptxEditor`
- `@officeai/react-editors/pdf` → `PdfEditor`

Each accepts the `EmbeddedEditorProps` defined under `./contract` today.

## Usage (phase 1)

```ts
import { makeBlankDocx, makeBlankXlsx, makeBlankPptx, makeBlankPdf } from "@officeai/react-editors/blanks";
import { DOCX_MIME, XLSX_MIME, PPTX_MIME, PDF_MIME } from "@officeai/react-editors/mime";

// Mint blank bytes in the browser, hand them to the host's existing
// upload pipeline (presigned S3 PUT, FileSystemAccess, …).
const bytes = await makeBlankDocx();
const file = new File([bytes], "Untitled.docx", { type: DOCX_MIME });
await uploadToS3(file);
```

## Distribution

This package is published as a tarball asset alongside `@officeai/agent`
on every push to `office-ai/main` — see the
[release pipeline doc](../../docs/release-pipeline.md). Hosts pin the
URL in their lockfile (e.g. `infra/officeai.lock.json` in hof-os).
