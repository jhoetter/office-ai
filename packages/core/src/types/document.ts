/**
 * Format-agnostic document abstractions. See spec/shared/document-model.md.
 */

export type NodeId = string;

export type DocumentFormat = "docx" | "xlsx" | "pptx" | "pdf";

export interface DocumentSnapshot<TRoot = unknown> {
  readonly format: DocumentFormat;
  readonly revision: number;
  readonly root: TRoot;
  readonly partHashes: Readonly<Record<string, string>>;
}

export type DiffPath = ReadonlyArray<string | number>;

/**
 * Optional structured payload attached to any diff change. Used for
 * format-specific extras that don't fit the generic
 * `{kind, nodeId, path, summary}` shape — e.g. `xlsx:set-cell-formula`
 * surfaces `{ cycle: ["Sheet1!0:0", ...] }` for circular references,
 * and `xlsx:set-cell-formula` carries before/after value snapshots
 * for inverse-handler reconstruction. Readers that don't recognise a
 * `meta` key MUST ignore it.
 */
export type DiffMeta = Readonly<Record<string, unknown>>;

export type DiffChange =
  | {
      kind: "node-inserted";
      nodeId: NodeId;
      path: DiffPath;
      summary: string;
      meta?: DiffMeta;
    }
  | {
      kind: "node-deleted";
      nodeId: NodeId;
      path: DiffPath;
      summary: string;
      meta?: DiffMeta;
    }
  | {
      kind: "node-updated";
      nodeId: NodeId;
      path: DiffPath;
      field: string;
      summary: string;
      meta?: DiffMeta;
    }
  | {
      kind: "node-moved";
      nodeId: NodeId;
      from: DiffPath;
      to: DiffPath;
      summary: string;
      meta?: DiffMeta;
    }
  | {
      /**
       * A new OPC part was added to the package (e.g. a media binary
       * inserted alongside its rel + content-type registration). `path`
       * carries the part path as a single-segment array — e.g.
       * `["word/media/image3.png"]` — so consumers that already key off
       * `change.path[0]` get the part path without a special case.
       *
       * Added in P1.3 / W8 (image insertion). Older changes do not emit
       * this kind; downstream readers that rely on the union should
       * handle it via an exhaustive switch.
       */
      kind: "part-added";
      path: DiffPath;
      summary: string;
      meta?: DiffMeta;
    };

export interface DocumentDiff {
  readonly format: DocumentFormat;
  readonly fromRevision: number;
  readonly toRevision: number;
  readonly changes: readonly DiffChange[];
}

export const emptyDiff = (
  format: DocumentFormat,
  fromRevision: number,
  toRevision: number
): DocumentDiff => ({
  format,
  fromRevision,
  toRevision,
  changes: [],
});
