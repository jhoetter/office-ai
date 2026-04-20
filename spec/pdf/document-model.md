# PDF — Document Model

> The typed projection over the original PDF byte buffer. The bytes
> are the source of truth; the model is what the editor operates on.

The shapes here mirror the as-built code in
[`packages/pdf/src/model/types.ts`](../../packages/pdf/src/model/types.ts).
Cross-references: shared core in
[`spec/shared/document-model.md`](../shared/document-model.md);
engine projections in
[`packages/pdf-engine/src/types.ts`](../../packages/pdf-engine/src/types.ts).

## Why a projection, not a full reconstruction

PDF is a binary object graph with cross-references, streams,
appearance dictionaries, and a forty-year history of producer quirks.
Reconstructing the full graph in TypeScript would put us on the path
PDF.js, pdf-lib, and PDFium have already walked — at the cost of
re-implementing PDF semantics in our own code. We refuse that path.

Instead:

- The `PdfAgent` retains the **original `Uint8Array`** for the life
  of the session. See
  [`PdfAgent.originalBytes()`](../../packages/pdf/src/agent/agent.ts).
- Reads project the byte buffer into a typed `PdfDocument` via the
  engine (`packages/pdf-engine`). The projection contains everything
  the editor exposes (pages, outline, annotations, form fields,
  attachments, comments, metadata).
- Writes flow through the `CommandBus` against the typed model. On
  `exportFile()`, the serializer produces a new `Uint8Array` by
  **incremental update** over the original — appending only the
  objects we changed, plus a fresh xref + trailer. See
  [`editing-pipeline.md`](./editing-pipeline.md).
- Untouched objects come back **byte-identical** because the original
  bytes are still there.

This is the same opaque-blob-preservation invariant as DOCX/XLSX/PPTX,
adapted for PDF's incremental-update primitive.

## Core types

```typescript
import type { DocumentSnapshot, NodeId } from "@officeai/core";

export type PdfRotation = 0 | 90 | 180 | 270;
export type PdfRect = readonly [number, number, number, number];

export interface PdfMetadata {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
  readonly creationDate?: string;
  readonly modificationDate?: string;
  readonly pdfVersion?: string;
  readonly linearized?: boolean;
  readonly encryption?: {
    readonly hasUserPassword: boolean;
    readonly hasOwnerPassword: boolean;
  };
}

export interface PdfPage {
  readonly id: NodeId;
  readonly pageNumber: number; // 1-indexed in current order
  readonly width: number; // PDF user-units (1/72 in)
  readonly height: number;
  readonly rotation: PdfRotation;
  readonly label?: string; // "iv", "A-12", …
  readonly text: string; // best-effort reading-order projection
  readonly hasTextLayer: boolean;
  readonly hasAnnotations: boolean;
  readonly hasFormFields: boolean;
}

export interface PdfOutlineNode {
  readonly id: NodeId;
  readonly title: string;
  readonly pageNumber?: number;
  readonly uri?: string;
  readonly children: ReadonlyArray<PdfOutlineNode>;
}

export type PdfAnnotationKind =
  | "highlight"
  | "underline"
  | "strikethrough"
  | "squiggly"
  | "note"
  | "free-text"
  | "ink"
  | "line"
  | "rectangle"
  | "ellipse"
  | "polygon"
  | "polyline"
  | "stamp"
  | "link"
  | "redaction"
  | "unknown";

export interface PdfAnnotation {
  readonly id: NodeId;
  readonly kind: PdfAnnotationKind;
  readonly subtype: string; // native PDF /Subtype
  readonly pageNumber: number;
  readonly rect: PdfRect; // PDF user-space
  readonly contents?: string;
  readonly author?: string;
  readonly color?: { r: number; g: number; b: number; a?: number };
  readonly url?: string; // for link
  readonly destPage?: number; // for goto-link
  readonly createdAt?: string;
  readonly nativeObjectNumber?: number; // for incremental save
}

export type PdfFormFieldType = "text" | "checkbox" | "radio" | "choice" | "button" | "signature" | "unknown";

export interface PdfFormField {
  readonly id: NodeId;
  readonly name: string; // /T fully-qualified name
  readonly type: PdfFormFieldType;
  readonly value?: string | boolean;
  readonly options?: ReadonlyArray<string>;
  readonly readOnly: boolean;
  readonly required: boolean;
  readonly maxLength?: number;
  readonly multiline?: boolean;
  readonly password?: boolean;
  readonly pageNumber: number;
  readonly rect: PdfRect;
}

export interface PdfAttachment {
  readonly id: NodeId;
  readonly name: string;
  readonly bytes: number;
}

export interface PdfComment {
  readonly id: NodeId;
  readonly author: string;
  readonly text: string;
  readonly resolved?: boolean;
  readonly parentId?: NodeId;
  readonly createdAt?: string;
  readonly pageNumber: number;
  readonly normalizedRect: PdfRect; // 0..1 normalized for stable anchor
}

export interface PdfDocument {
  readonly metadata: PdfMetadata;
  readonly pages: ReadonlyArray<PdfPage>;
  readonly outline: ReadonlyArray<PdfOutlineNode>;
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly formFields: ReadonlyArray<PdfFormField>;
  readonly attachments: ReadonlyArray<PdfAttachment>;
  readonly comments: ReadonlyArray<PdfComment>;
  readonly signatureCount: number;
  readonly engineKind: "pdfjs" | "pdfium";
}

export type PdfSnapshot = DocumentSnapshot<PdfDocument> & { readonly format: "pdf" };
```

## Identity rules

- **Every structural node carries a `NodeId`.** Minted by the parser
  (`opts.idMinter` falls through to `mintNodeId` from
  `@officeai/core`). Stable across mutations: an annotation that's
  edited keeps its `NodeId`. A page that's rotated keeps its
  `NodeId`.
- A node deleted by one mutation and re-inserted by a later mutation
  gets a **new** `NodeId`; ids are not recycled.
- For annotations and form fields we additionally store the **native
  PDF object number** (`nativeObjectNumber`) when known. Incremental
  save uses this to overwrite the original object in place rather
  than appending a duplicate.
- Snapshots are **immutable**. Mutations produce new snapshots; the
  old snapshot remains valid and addressable for diff purposes.

## `partHashes` for PDF

PDF is not a part-keyed container, but the byte-preservation invariant
needs a verifiable equivalent. We project the buffer into a synthetic
`partHashes` map keyed by **logical region**:

```
{
  "trailer:original": "sha256(original trailer bytes)",
  "xref:0":           "sha256(original xref section 0)",
  "page:1":           "sha256(/Page 1 dict + content streams)",
  "page:2":           "sha256(/Page 2 dict + content streams)",
  …
  "annots:page:1":    "sha256(concatenated /Annots dicts on page 1)",
  "form:field:Name1": "sha256(field dict + appearance)",
  "metadata:xmp":     "sha256(/Metadata stream bytes)",
  "outlines":         "sha256(/Outlines tree bytes)",
}
```

The hashes are computed once at parse time (cost: amortized one pass
over the buffer) and stored on the snapshot. After a `pdf:rotate-pages`
mutation, only the affected `page:N` keys change; everything else is
identical, which the audit-roundtrip pipeline confirms.

This is conceptually identical to OOXML's `partHashes` over zip
entries, just keyed differently.

## Comments anchored as `pdf-region`

Comments are not native PDF annotations — they live in the
`@officeai/comments` package and sync over Y.js. Their anchor is a
typed variant added to
[`packages/comments/src/types.ts`](../../packages/comments/src/types.ts):

```typescript
export interface PdfRegionAnchor {
  readonly kind: "pdf-region";
  readonly pageNumber: number; // 1-indexed
  readonly normalizedRect: PdfRect; // 0..1 in page user-space
  readonly nativeObjectNumber?: number; // optional native /Annot ref
}
```

**Why normalized rect:** stable across rotation and zoom. The viewer
denormalizes against the current page user-units to draw the comment
indicator. When the page is rotated by `pdf:set-page-rotation`, the
normalized rect is rotated in-place by the comment handler so the
indicator stays attached to the underlying content.

**Why optional native object number:** if a comment is mirrored as a
native sticky-note annotation (so it survives export to a non-Office-AI
viewer), we record the native object number so subsequent edits update
the same object. If the comment is Office-AI-internal only, the field
is absent.

## What's _not_ in the model

- The full PDF object graph. The model never carries a
  `PdfIndirectObject` or `PdfStream` type.
- Render bitmaps. Those live in the LRU raster cache in the renderer
  ([`rendering-pipeline.md`](./rendering-pipeline.md)), not the model.
- Font substitution decisions. Those are engine-internal.
- Compressed object streams (`/ObjStm`) layout. We round-trip them
  via incremental save without decoding.
- Resource dictionaries (`/Font`, `/XObject`, `/ColorSpace`,
  `/Pattern`, `/ExtGState`). These are byte-preserved.

## Lifecycle

1. **Parse.** `PdfAgent.fromBuffer(bytes)` → engine opens the
   document → projects each page into `PdfPage` (text + flags) →
   walks `/Outlines` → walks `/Annots` per page → walks `/AcroForm` →
   walks `/Names/EmbeddedFiles` → produces `PdfDocument` →
   constructs `PdfSnapshot` with revision 0.
2. **Mutate.** `agent.applyCommand(cmd)` → `CommandBus.dispatch` →
   handler produces new `PdfDocument` + bumps revision → snapshot
   replaced. Pending mutations from `source: "agent"` go to the
   pending queue first; human approves → moves to approved.
3. **Serialize.** `agent.exportFile()` → `serializePdf(snapshot,
originalBuffer)` → incremental update with only changed objects.
4. **Re-parse.** `PdfAgent.fromBuffer(exported)` produces a new
   snapshot at revision 0; the round-trip test asserts byte-equality
   on untouched regions.

## Headless invariant

`@officeai/pdf` (the model + parser + serializer + agent) does **not
transitively import** `react`, `react-dom`, `next`, or any DOM
global. Enforced by `scripts/check-architecture.mjs`. The agent is
fully usable in Node.js — that's what makes the CLI surface real.
