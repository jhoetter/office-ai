# Embedding office-ai editors in a host app

The `@officeai/react-editors` package is the supported way for
external React 19+ apps to embed the office-ai editing surfaces —
without standing up the `apps/web` Next.js shell and without
re-implementing OOXML / PDF I/O.

This doc is the contract — what the package exports today (Phase 1),
what's coming (Phase 1.5), and what assumptions hosts can rely on.

## Phase 1 (shipped)

The package is published as a self-contained tarball on every
office-ai GitHub Release alongside `officeai-agent-X.Y.Z.tgz` (see
[release-pipeline.md](./release-pipeline.md)).

Subpath exports:

| Import                                    | What you get                                                                                                                              | When to use                                                                                                                               |
| ----------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `@officeai/react-editors/blanks`          | `makeBlankDocx`, `makeBlankXlsx`, `makeBlankPptx`, `makeBlankPdf` (each `() => Promise<Uint8Array>`) and a `makeBlank(format)` dispatcher | "Create new" actions in a host file browser. All run entirely in the browser via `*Agent.empty().exportFile()` — no Node, no service hop. |
| `@officeai/react-editors/blanks/<format>` | Just one builder                                                                                                                          | When tree-shaking individual formats matters.                                                                                             |
| `@officeai/react-editors/mime`            | `DOCX_MIME`, `XLSX_MIME`, `PPTX_MIME`, `PDF_MIME`, `MIME_BY_FORMAT`, `detectFormatFromFilename`                                           | Wherever you need the canonical OOXML / PDF MIME strings (presigned PUT `Content-Type`, `<File>` constructor, server allowlists).         |
| `@officeai/react-editors/contract`        | Type-only `EmbeddedEditorProps`, `Locale`, `Theme`, `EmbeddedEditorOnSave`                                                                | For typing the future editor mounts; the editor _components_ themselves ship in Phase 1.5.                                                |

Example — Create-new + presigned object-storage PUT:

```ts
import { makeBlankXlsx } from "@officeai/react-editors/blanks/xlsx";
import { XLSX_MIME } from "@officeai/react-editors/mime";

const bytes = await makeBlankXlsx();
const file = new File([bytes], "Untitled.xlsx", { type: XLSX_MIME });
await uploadFile(file); // your existing presign+PUT pipeline
```

## Phase 1.5 (shipped)

`<DocxEditor />`, `<XlsxEditor />`, `<PptxEditor />`, `<PdfEditor />`
are now exported from `@officeai/react-editors/components/<format>`
and satisfy the `EmbeddedEditorProps` contract typed under `./contract`:

```ts
interface EmbeddedEditorProps {
  initialBytes?: Uint8Array; // host streams object bytes straight in
  initialFilename?: string; // shown in the editor titlebar
  onSave?: EmbeddedEditorOnSave; // (bytes, mime, filename) => Promise<void>
  onClose?: () => void;
  locale?: Locale; // "en" | "de"
  theme?: Theme; // "light" | "dark"
}
```

Example — open an `.xlsx`, edit, persist back to object storage:

```tsx
import { XlsxEditor } from "@officeai/react-editors/components/xlsx";

<XlsxEditor
  initialBytes={bytes /* Uint8Array from your presigned GET */}
  initialFilename="expenses.xlsx"
  onSave={async (bytes, mime, filename) => {
    const presign = await yourPresignUploadFn({ filename, content_type: mime });
    await fetch(presign.upload_url, {
      method: "PUT",
      headers: { "Content-Type": mime },
      body: bytes,
    });
  }}
  onClose={() => router.back()}
/>;
```

### Packaging notes

The components are **bundled by esbuild** from the `apps/web/app/...`
sources at release time. We took this route instead of physically
relocating the editor code because:

- Forking ~11 k LOC across four formats plus the shared
  `apps/web/app/lib/{shell,realtime,...}` subtree would duplicate,
  drift, and break `apps/web`'s own routes.
- The only Next-specific leak in the bundled graph is `next/link` in
  `lib/shell/EditorTopBar.tsx`. `build.mjs` aliases that to a
  plain-`<a>` shim under `src/shims/next-link.tsx`, so the bundle has
  zero Next runtime cost.
- The seven "format" workspace packages
  (`@officeai/{core,docx,xlsx,pptx,pdf,pdf-engine,pdf-annotations}`)
  ship pre-built `dist/*.js` and are externalized so the host
  bundler dedupes them with the headless `./blanks` entries.
- The five "shell" workspace packages
  (`@officeai/{ui,text-formatting,comments,realtime,design-tokens}`)
  ship raw `.ts` from `./src/`, which Vite refuses to transpile from
  `node_modules`. They're inlined into the bundle (~50 KB / editor)
  so hosts don't need to special-case them.

Heavy third-party libs (`pdfjs-dist`, `prosemirror-*`, `jszip`, `yjs`,
`y-websocket`) stay externalized for the same dedup
reason and ship as direct deps in `package.json` so pnpm + the
shameful-hoist step in `.github/workflows/auto-release.yml` puts them
at the consumer's deploy root (required for Vite's
`resolve.preserveSymlinks: true`).

## Versioning

The package version is bumped in lockstep with `@officeai/agent` by
[`scripts/bump-version.mjs`](../scripts/bump-version.mjs). Hosts pin
both via their own lockfile; there is no separate version axis for
editors vs CLI.

## What stays out

- **No iframe embedding** — the editors must run in the host's
  React tree to share the host's font, theme tokens, and event
  loop. iframes are not supported.
- **No standalone Next.js bundle** — `apps/web` is the developer
  reference shell, not an embedding target.
- **No Tailwind v3 backport** — the editors target Tailwind v4. Older
  hosts must upgrade.
