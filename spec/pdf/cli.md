# PDF — CLI

> Every `office-agent pdf-*` subcommand the LLM is expected to call.
> JSON-by-default output; `--format text|md|json` switch where
> applicable. **The CLI is the AI integration surface for this
> product.** No LLM logic ships in-product.

Cross-references: agent shape in
[`packages/pdf/src/agent/agent.ts`](../../packages/pdf/src/agent/agent.ts);
typed command surface in [`agent-commands.md`](./agent-commands.md);
serializer in [`editing-pipeline.md`](./editing-pipeline.md).

## Conventions

- All commands live under `office-agent pdf …`.
- All write commands take `--file <input.pdf>` and `--out <output.pdf>`.
- All read commands take `--file <input.pdf>` and emit JSON to stdout
  unless `--format <text|md|json>` is overridden.
- Page indices are **1-indexed**.
- Page ranges accept `1,3,5`, `1-10`, `1-end`, or `all`.
- `--password <pw>` decrypts encrypted inputs in-memory; the password
  is never written.
- All JSON outputs carry a `format` field with a versioned tag (e.g.
  `pdf-metadata/v1`) so MCP clients can pin to a schema.
- Exit codes: `0` success, `1` user error (bad args / invalid input),
  `2` engine error (parse / render failure), `3` IO error (read /
  write).

## Read commands

### `office-agent pdf inspect`

```
office-agent pdf inspect --file f.pdf [--format json|text|md]
```

Aggregate read: pages + metadata + outline + annotations summary +
forms summary + signatures + fonts. The "go-to" first-call for an
agent.

```json
{
  "format": "pdf-inspect/v1",
  "file": "f.pdf",
  "pageCount": 24,
  "engine": "pdfjs",
  "pdfVersion": "1.7",
  "linearized": false,
  "encrypted": false,
  "metadata": { "title": "Report", "author": "Alice" },
  "pageSizes": [{ "pageNumber": 1, "width": 612, "height": 792 }, …],
  "outlineDepth": 3,
  "annotationCounts": { "highlight": 12, "note": 4, "ink": 0 },
  "formFieldCount": 8,
  "fontCount": 6,
  "attachmentCount": 0,
  "signatureCount": 0,
  "warnings": []
}
```

### `office-agent pdf metadata`

```
office-agent pdf metadata --file f.pdf
```

```json
{
  "format": "pdf-metadata/v1",
  "title": "Report",
  "author": "Alice",
  "subject": "Annual review",
  "keywords": "report,annual",
  "creator": "Microsoft Word",
  "producer": "Acrobat Distiller 2024",
  "creationDate": "2026-01-15T09:00:00Z",
  "modificationDate": "2026-04-01T12:00:00Z",
  "pdfVersion": "1.7",
  "linearized": false,
  "encryption": { "hasUserPassword": false, "hasOwnerPassword": false }
}
```

### `office-agent pdf outline`

```
office-agent pdf outline --file f.pdf
```

```json
{
  "format": "pdf-outline/v1",
  "items": [
    {
      "id": "n_…",
      "title": "Chapter 1",
      "pageNumber": 1,
      "children": [{ "id": "n_…", "title": "Section 1.1", "pageNumber": 3, "children": [] }]
    }
  ]
}
```

### `office-agent pdf list-pages`

```
office-agent pdf list-pages --file f.pdf
```

```json
{
  "format": "pdf-page-list/v1",
  "pages": [
    {
      "pageNumber": 1,
      "id": "n_…",
      "width": 612,
      "height": 792,
      "rotation": 0,
      "label": "i",
      "hasTextLayer": true,
      "hasAnnotations": false,
      "hasFormFields": false
    }
  ]
}
```

### `office-agent pdf list-annotations`

```
office-agent pdf list-annotations --file f.pdf [--page N] [--type highlight|note|ink|...]
```

```json
{
  "format": "pdf-annotations/v1",
  "annotations": [
    {
      "id": "n_…",
      "kind": "highlight",
      "subtype": "Highlight",
      "pageNumber": 5,
      "rect": [100, 200, 400, 220],
      "color": { "r": 1.0, "g": 1.0, "b": 0.0 },
      "contents": "Important clause.",
      "author": "Alice",
      "createdAt": "2026-04-19T10:00:00Z"
    }
  ]
}
```

### `office-agent pdf list-form-fields`

```
office-agent pdf list-form-fields --file f.pdf
```

```json
{
  "format": "pdf-form-fields/v1",
  "backing": "acroform",
  "fields": [
    {
      "id": "n_…",
      "name": "applicant.lastName",
      "type": "text",
      "value": "",
      "readOnly": false,
      "required": true,
      "maxLength": 64,
      "pageNumber": 1,
      "rect": [100, 600, 300, 620]
    }
  ]
}
```

### `office-agent pdf list-fonts`

```
office-agent pdf list-fonts --file f.pdf
```

```json
{
  "format": "pdf-fonts/v1",
  "fonts": [
    {
      "name": "Helvetica",
      "type": "Type1",
      "embedded": false,
      "subset": false,
      "encoding": "WinAnsiEncoding",
      "usedOnPages": [1, 2, 3]
    },
    {
      "name": "ABCDEF+TimesNewRoman",
      "type": "TrueType",
      "embedded": true,
      "subset": true,
      "encoding": "Identity-H",
      "usedOnPages": [4, 5]
    }
  ]
}
```

### `office-agent pdf list-attachments`

```
office-agent pdf list-attachments --file f.pdf [--out ./attachments/]
```

```json
{
  "format": "pdf-attachments/v1",
  "attachments": [{ "id": "n_…", "name": "appendix.xlsx", "bytes": 12345 }]
}
```

`--out` writes attachment bytes to the directory.

### `office-agent pdf list-signatures`

```
office-agent pdf list-signatures --file f.pdf
```

```json
{
  "format": "pdf-signatures/v1",
  "signatures": [
    {
      "fieldName": "Sig1",
      "pageNumber": 1,
      "rect": [100, 200, 300, 250],
      "signedAt": "2026-04-19T18:32:00Z",
      "signedBy": "John Hoetter",
      "valid": true,
      "coversWholeDocument": true
    }
  ]
}
```

### `office-agent pdf read`

```
office-agent pdf read --file f.pdf [--pages 1-5|all] [--bbox page=N,x=,y=,w=,h=] [--format markdown|json|text]
```

`--format text` (default for `--bbox`): plain text in reading order.
`--format markdown` (default for `--pages`): structured Markdown using
the struct tree (or heuristic) — headings, lists, paragraphs, page
breaks as `---`.
`--format json`: per-page JSON with text items.

### `office-agent pdf search`

```
office-agent pdf search --file f.pdf --query "..." [--regex] [--case] [--whole-word] [--pages 1-50] [--format json|text|md]
```

Detail in [`search.md`](./search.md).

### `office-agent pdf chunk`

```
office-agent pdf chunk --file f.pdf [--strategy outline|page|fixed-tokens=1000]
```

Splits the document into LLM-friendly chunks. Strategies:

- `outline` — one chunk per top-level outline entry (default if
  `/Outlines` is non-empty).
- `page` — one chunk per page.
- `fixed-tokens=N` — greedy fill to ~N tokens per chunk (estimated
  via `text.length / 4`).

```json
{
  "format": "pdf-chunks/v1",
  "strategy": "outline",
  "chunks": [{ "id": "ch_1", "title": "Chapter 1", "pages": [1, 2, 3, 4, 5], "tokens": 1234, "text": "…" }]
}
```

### `office-agent pdf render`

```
office-agent pdf render --file f.pdf --pages 1-3 --out ./out/ --format png|jpeg|webp --dpi 150
```

Rasterizes pages to disk via the headless engine. Output filenames
are `page-1.png`, `page-2.png`, … under `--out`.

```json
{
  "format": "pdf-render/v1",
  "renderedPages": [
    { "pageNumber": 1, "outPath": "./out/page-1.png", "width": 1275, "height": 1650, "bytes": 234567 }
  ]
}
```

### `office-agent pdf thumbnail`

```
office-agent pdf thumbnail --file f.pdf --page 1 --out cover.webp --width 400
```

Emits a single thumbnail at the given width (height computed from
the page aspect ratio).

### `office-agent pdf extract-images`

```
office-agent pdf extract-images --file f.pdf [--page N] --out ./images/
```

Extracts every embedded image (XObject of subtype `/Image`) at
original resolution. Output:

```json
{
  "format": "pdf-extracted-images/v1",
  "images": [
    { "pageNumber": 3, "outPath": "./images/p3-img1.png", "width": 800, "height": 600, "bytes": 45678 }
  ]
}
```

### `office-agent pdf extract-text-stream`

```
office-agent pdf extract-text-stream --file f.pdf --page N
```

Emits the raw content stream's text-show operators in JSON for
debugging / extreme-fidelity workflows.

## Write — page-level

### `office-agent pdf create`

```
office-agent pdf create --out blank.pdf [--size letter|a4|legal] [--orientation portrait|landscape]
```

Creates a single-page blank PDF.

### `office-agent pdf rotate`

```
office-agent pdf rotate --file f.pdf --pages 2,4,6 --angle 90 --out o.pdf
```

`--angle` must be a multiple of 90.

### `office-agent pdf reorder`

```
office-agent pdf reorder --file f.pdf --order "1,3,2,4-end" --out o.pdf
```

`--order` accepts the same range syntax as `--pages`. Must be a
permutation of `1..N`.

### `office-agent pdf insert-pages`

```
office-agent pdf insert-pages --file f.pdf --from other.pdf [--pages 1-3] --at N --out o.pdf
```

### `office-agent pdf delete-pages`

```
office-agent pdf delete-pages --file f.pdf --pages 5,9 --out o.pdf
```

### `office-agent pdf split`

```
office-agent pdf split --file f.pdf --by range|bookmark|size --out ./parts/ [--ranges "1-10,11-20"] [--max-bytes 5000000]
```

```json
{
  "format": "pdf-split/v1",
  "parts": [{ "outPath": "./parts/part-1.pdf", "pages": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], "bytes": 234567 }]
}
```

### `office-agent pdf merge`

```
office-agent pdf merge --files a.pdf b.pdf c.pdf --out merged.pdf
```

### `office-agent pdf extract-pages`

```
office-agent pdf extract-pages --file f.pdf --pages 1-10 --out ch1.pdf
```

### `office-agent pdf crop`

```
office-agent pdf crop --file f.pdf --margin "10,10,10,10" [--pages all] --out o.pdf
```

`--margin` is `left,top,right,bottom` in PDF user-units (1/72 in).

### `office-agent pdf watermark`

```
office-agent pdf watermark --file f.pdf [--text "DRAFT" | --image w.png] --opacity 0.2 [--position center|...] [--rotation 45] --out o.pdf
```

### `office-agent pdf add-page-numbers`

```
office-agent pdf add-page-numbers --file f.pdf --position bottom-center [--format "Page {n} of {total}"] [--start-at 1] --out o.pdf
```

### `office-agent pdf set-metadata`

```
office-agent pdf set-metadata --file f.pdf --json '{"title":"X","author":"Y"}' --out o.pdf
```

## Write — annotations / forms / OCR

### `office-agent pdf import-annotations`

```
office-agent pdf import-annotations --file f.pdf --annotations ./a.xfdf [--format xfdf|fdf|json] --out o.pdf
```

### `office-agent pdf export-annotations`

```
office-agent pdf export-annotations --file f.pdf [--format xfdf|fdf|json]
```

### `office-agent pdf fill-form`

```
office-agent pdf fill-form --file f.pdf --data ./values.json [--flatten] --out o.pdf
```

`values.json`:

```json
{ "applicant.firstName": "Alice", "subscribe": true, "color": "blue" }
```

### `office-agent pdf reset-form`

```
office-agent pdf reset-form --file f.pdf --out o.pdf
```

### `office-agent pdf flatten-form`

```
office-agent pdf flatten-form --file f.pdf --out o.pdf
```

### `office-agent pdf redact`

```
office-agent pdf redact --file f.pdf [--pattern "phone"] [--pattern "email"] [--rects ./rects.json] --out o.pdf --log redaction.json
```

Patterns are JS regex strings. `--rects` JSON is `[{ "page": N,
"rect": [x,y,w,h] }, …]`. The `--log` file records every redacted
region:

```json
{
  "format": "pdf-redaction-log/v1",
  "redactions": [
    { "pageNumber": 3, "rect": [100, 200, 300, 220], "matchedText": "555-1212", "pattern": "phone" }
  ]
}
```

### `office-agent pdf ocr`

```
office-agent pdf ocr --file scan.pdf --lang deu+eng --out searchable.pdf [--pages 1-3]
```

Adds an invisible-but-selectable text layer to scanned pages. Lazy-
loads `tesseract.js`.

## Apply / diff (existing convention)

### `office-agent pdf apply`

```
office-agent pdf apply --file f.pdf --commands ./batch.json --out o.pdf
```

`batch.json` is `[{ "type": "pdf:rotate-pages", "payload": { … } }, …]`.
Commands are dispatched in order; first failure aborts.

### `office-agent pdf diff`

```
office-agent pdf diff --before v1.pdf --after v2.pdf [--pretty]
```

```json
{
  "format": "pdf-diff/v1",
  "fromBytes": 234567,
  "toBytes": 234890,
  "pageCountDelta": 0,
  "rotationChanges": [{ "pageNumber": 3, "from": 0, "to": 90 }],
  "annotationCountDelta": { "highlight": 2, "note": 1 },
  "metadataChanges": ["title", "modificationDate"],
  "incrementalSaveUsed": true
}
```

`--pretty` prints a human-readable summary.

## MCP

Every command above is also exposed as an MCP tool via
`packages/agent/src/pdf-mcp.ts`. The tool names are
`pdf_inspect`, `pdf_metadata`, `pdf_outline`, … `pdf_apply`,
`pdf_diff`. MCP clients (Claude Desktop, Cursor, …) get the full
PDF surface as first-class tools.

## Why JSON-by-default

LLMs reliably consume JSON; `--format text|md` exists for the
human-in-the-loop "I want to read this" cases. The schema-versioned
`format` field on every JSON output lets MCP clients validate
responses with zod and pin against schema drift.
