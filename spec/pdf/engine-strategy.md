# PDF — Engine Strategy

> Why PDF.js is the default, when PDFium-WASM is loaded, and how the
> abstraction lets us swap engines without a viewer rewrite.

## Decision

**Default engine: PDF.js (Apache 2.0).** Loaded eagerly behind
`packages/pdf-engine` for every document.

**Fidelity fallback: PDFium-WASM (BSD-3-Clause) via
`@embedpdf/pdfium`.** Loaded lazily, opt-in per document, or
auto-triggered by a `selectEngine()` heuristic.

**Strategic migration to PDFium-primary** is documented but **not
implemented**. The engine abstraction in
[`packages/pdf-engine/src/types.ts`](../../packages/pdf-engine/src/types.ts)
is engine-agnostic — the swap is internal.

The decision is also recorded as the as-built shape in
`PdfEngineKind = "pdfjs" | "pdfium"` and `selectEngine(hints)` (see
[`packages/pdf-engine/src/select-engine.ts`](../../packages/pdf-engine/src/select-engine.ts)).

## Why PDF.js as default

| Concern                       | PDF.js                          | PDFium-WASM                          |
| ----------------------------- | ------------------------------- | ------------------------------------ |
| License                       | Apache 2.0                      | Apache 2.0 wrapper / BSD-3 engine    |
| Bundle (gzipped)              | ≈350 KB                         | ≈3.5 MB lazy                         |
| First paint (50 MB doc)       | < 600 ms                        | > 1.5 s (WASM boot)                  |
| Worker isolation              | First-class, ships with library | Yes (lazy boot, opt-in)              |
| Text-layer extraction         | Mature `getTextContent()`       | Lower-level; we'd build the layer    |
| AcroForm widget layer         | Mature                          | Lower-level; we'd build widgets      |
| Struct-tree-layer (a11y)      | Mature                          | Available; we'd render               |
| Range-request streaming       | Yes                             | No                                   |
| Headless Node renderer        | `pdfjs-dist/legacy`             | Available; same WASM as browser      |
| Fidelity on edge fonts/colors | Approximate; substitutes        | Pixel-correct                        |

Two factors drive the default:

1. **First-paint matters more than fidelity for the median PDF.** A
   business letter, contract, invoice, scanned receipt, or LaTeX paper
   renders correctly under PDF.js. Loading 3.5 MB of WASM for these
   documents would lose us the "faster than Chrome" claim
   ([`performance.md`](./performance.md)).
2. **PDF.js gives us text-layer + AcroForm + struct-tree for free.**
   These are pre-built in PDF.js and would otherwise be hundreds of
   lines of replicated work in our own engine wrapper.

## Why PDFium-WASM as fallback (not primary)

The cases where PDF.js is approximately right are well-understood:

- Type3 fonts (PDF.js substitutes with a system font; PDFium renders).
- DeviceN / NChannel / Separation color spaces (PDF.js → RGB approx;
  PDFium honors tint transforms).
- Embedded ICC profiles for color-managed CMYK (PDF.js ignores; PDFium
  honors `/OutputIntent`).
- Custom CMaps (PDF.js sometimes glyph-substitutes; PDFium maps
  consistently via `/CIDSystemInfo`).
- Linearized PDFs with non-standard cross-reference streams in older
  PDF 1.5+ files (rare but real).

These are the exact signals fed into `selectEngine()`.

## Auto-fallback heuristics (`selectEngine(hints)`)

```typescript
export interface EngineSelectionHints {
  hasUncommonColorSpace?: boolean; // DeviceN/NChannel/Separation/Lab in /ColorSpace
  hasCustomCMap?: boolean;         // non-standard /CIDSystemInfo
  hasType3Fonts?: boolean;         // /Font/Subtype /Type3
  linearized?: boolean;            // informational; not a fallback trigger by itself
  inPdfiumAllowlist?: boolean;     // curated allowlist hit
  userPrefersFidelity?: boolean;   // explicit user opt-in
}

export const selectEngine = (hints: EngineSelectionHints = {}): PdfEngineKind => {
  if (hints.userPrefersFidelity === true) return "pdfium";
  if (hints.inPdfiumAllowlist === true) return "pdfium";
  if (hints.hasUncommonColorSpace === true) return "pdfium";
  if (hints.hasType3Fonts === true) return "pdfium";
  if (hints.hasCustomCMap === true) return "pdfium";
  return "pdfjs";
};
```

Order of precedence is fixed: explicit user preference wins over
allowlist; allowlist wins over heuristics. No probabilistic scoring;
the function is pure and trivially testable.

### How hints are gathered

The `pdfjs-dist` document is opened first (it has to be — PDF.js gets
us cheap probing). The hint extractor walks `/Resources/Font`,
`/Resources/ColorSpace`, and the trailer:

- `/Type /Font /Subtype /Type3` anywhere → `hasType3Fonts`.
- `/ColorSpace [/DeviceN ...]` or `[/Separation ...]` or `[/NChannel ...]`
  or `[/Lab ...]` → `hasUncommonColorSpace`.
- `/CIDSystemInfo` whose `/Registry` is not in `{Adobe}` and `/Ordering`
  is not in `{GB1, CNS1, Japan1, Japan2, Korea1, KR, UCS, Identity}` →
  `hasCustomCMap`.
- Allowlist (`prefers-pdfium.json` curated from
  `spec/pdf/engine-fidelity-audit.md`) → `inPdfiumAllowlist`.

If any hint fires, the document is **re-opened** in PDFium and the
PDF.js handle destroyed. If no hint fires, the PDF.js handle is kept
and used for read paths.

The user can always force PDFium via the page menu ("Use high-fidelity
rendering"), which sets `userPrefersFidelity: true` and re-opens.

## The engine abstraction

```typescript
export interface PdfEngine {
  readonly kind: PdfEngineKind;
  load(buffer: Uint8Array, opts?: PdfEngineLoadOptions): Promise<PdfEngineDocument>;
}

export interface PdfEngineDocument {
  readonly engine: PdfEngineKind;
  readonly numPages: number;
  getPage(pageNumber: number): Promise<PdfEnginePage>;
  getMetadata(): Promise<PdfEngineMetadata>;
  getOutline(): Promise<PdfEngineOutlineNode[] | null>;
  getAttachments(): Promise<ReadonlyArray<{ name: string; data: Uint8Array }>>;
  estimatedBytes(): number;
  destroy(): Promise<void>;
}
```

Both backends implement this. Consumers
(`packages/pdf/src/parser/parse.ts`,
`packages/pdf-ocr`, `packages/pdf-annotations` previews) depend on the
contract — never on `pdfjs-dist` or `@embedpdf/pdfium` directly. This
is enforced by `scripts/check-architecture.mjs`: only
`packages/pdf-engine/src/backends/*.ts` may import the underlying
engine packages.

## Bundle-size budget

| Layer                                | Budget (gzipped) |
| ------------------------------------ | ---------------- |
| `packages/pdf-engine` shell          | < 5 KB           |
| `pdfjs-dist` worker + main           | < 400 KB         |
| `@embedpdf/pdfium` (lazy)            | < 4 MB           |
| `@officeai/pdf` (model + parser)     | < 50 KB          |
| `@officeai/pdf-edit`                 | < 30 KB          |
| `@officeai/pdf-annotations`          | < 30 KB          |
| `@officeai/pdf-forms`                | < 20 KB          |
| `@officeai/pdf-ocr` (lazy)           | < 1 MB (tesseract)|

CI gate: `bundlesize` config in the web app fails the build on
regressions. PDFium and tesseract are excluded from the eager-load
budget — they must be lazy.

## Migration path (PDFium-primary, not this session)

If/when we move PDFium to primary:

1. **Keep the abstraction intact.** The interface has been engine-
   agnostic from day one — see `PdfEngineDocument`/`PdfEnginePage`.
2. **Re-implement the text-layer extractor on PDFium.** PDFium emits
   text fragments via `FPDFText_*`; our wrapper would build the same
   `PdfEngineTextItem[]` shape PDF.js currently emits.
3. **Re-implement the AcroForm widget layer on PDFium.** The
   `PdfEngineFormFieldLite` shape stays.
4. **Re-implement the struct-tree exporter on PDFium.** Same target
   shape (`PdfStructTree` in [`document-model.md`](./document-model.md)).
5. **Demote PDF.js to a lazy text-layer-only fallback** for documents
   that PDFium can't open (extremely rare in practice).

The migration is purely internal to `packages/pdf-engine`. Consumers
do not change. This is the whole point of the abstraction.

## Headless / Node story

Both backends ship a Node entry point:

- PDF.js → `pdfjs-dist/legacy/build/pdf.mjs` + `@napi-rs/canvas`. No
  DOM dependencies.
- PDFium → same WASM as browser; canvas backed by `@napi-rs/canvas`.

`@officeai/pdf-engine` chooses the backend identically in Node and
browser. The CLI ([`cli.md`](./cli.md)) calls `loadDocument()` and
`renderPage()` without caring which engine answers.

## Testing strategy

- **Render fidelity smoke** — 3 fixtures rendered under PDF.js +
  PDFium, asserted within < 1% pixel diff for non-CJK text-only pages
  and < 5% for graphics-heavy pages.
- **Engine selection unit tests** — every hint combination → expected
  backend; tested in
  [`packages/pdf-engine/src/select-engine.test.ts`](../../packages/pdf-engine/src/select-engine.test.ts).
- **Headless smoke** — `office-agent pdf render` exercised in CI on
  Node, asserts byte-identical PNGs across runs (deterministic
  rendering).
- **Bundle-size CI gate** — fails on regression past the budget table.
