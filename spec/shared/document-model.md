# Shared Document Model

> The format-agnostic abstractions every editor in this monorepo speaks.
> DOCX, XLSX, and PPTX implementations all conform to these shapes.

## Core types

```typescript
/** A stable identifier minted by the parser/command bus, never reused. */
export type NodeId = string;

/** A document is a typed root node. */
export interface DocumentSnapshot<TRoot = unknown> {
  /** The format produced by the parser that built this snapshot. */
  readonly format: "docx" | "xlsx" | "pptx";
  /** Monotonically increasing per-document. Bumped by every applied mutation. */
  readonly revision: number;
  /** Format-specific root node (DocxDocument, XlsxWorkbook, PptxPresentation). */
  readonly root: TRoot;
  /**
   * Hash of every OOXML part that backs this snapshot, keyed by zip path.
   * Used by the serializer to detect "untouched" parts and re-emit them
   * byte-for-byte from the container cache.
   */
  readonly partHashes: Readonly<Record<string, string>>;
}

/**
 * Format-specific addressing. DOCX uses paragraph/run paths;
 * XLSX uses sheet+cell refs; PPTX uses slide+shape refs.
 * The shared core treats this as opaque.
 */
export type Position = unknown;
export type Selection = unknown;

/** A structured difference between two snapshots. */
export interface DocumentDiff {
  readonly format: "docx" | "xlsx" | "pptx";
  readonly fromRevision: number;
  readonly toRevision: number;
  readonly changes: readonly DiffChange[];
}

export type DiffChange =
  | { kind: "node-inserted"; nodeId: NodeId; path: ReadonlyArray<string | number>; summary: string }
  | { kind: "node-deleted"; nodeId: NodeId; path: ReadonlyArray<string | number>; summary: string }
  | { kind: "node-updated"; nodeId: NodeId; path: ReadonlyArray<string | number>; field: string; summary: string }
  | { kind: "node-moved";   nodeId: NodeId; from: ReadonlyArray<string | number>; to: ReadonlyArray<string | number>; summary: string };
```

## Identity rules

- Every structural node carries a `NodeId`. The parser mints them. They are
  stable across mutations (a paragraph that's edited keeps its id).
- A node deleted by one mutation and re-inserted by a later mutation gets a
  new id; ids are not recycled.
- Snapshots are **immutable**. Mutations produce new snapshots; the old
  snapshot remains valid and addressable for diff purposes.

## Per-format root shapes

These are defined fully in each format's `spec/{format}/document-model.md`.
At the shared level we only need to know:

- DOCX: `DocxDocument` is a sequence of block-level nodes (`Paragraph`, `Table`, `OpaqueBlock`).
- XLSX: `XlsxWorkbook` is a list of `Sheet`s plus shared resources (styles, strings, defined names).
- PPTX: `PptxPresentation` is a list of `Slide`s plus master/layout references.

## What lives in the shared core

- `DocumentSnapshot`, `DocumentDiff`, `DiffChange`, `NodeId` — types only.
- `mintNodeId()` — uuid-v4-backed factory.
- `hashPart(buffer)` — SHA-256 over a zip part's bytes.
- `freezeSnapshot(snapshot)` — `Object.freeze` walker for defensive immutability in dev/test.

What does **not** live in the core: any format-specific node type, any
parser, any renderer.
