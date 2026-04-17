/**
 * Format-agnostic document abstractions. See spec/shared/document-model.md.
 */

export type NodeId = string;

export type DocumentFormat = "docx" | "xlsx" | "pptx";

export interface DocumentSnapshot<TRoot = unknown> {
  readonly format: DocumentFormat;
  readonly revision: number;
  readonly root: TRoot;
  readonly partHashes: Readonly<Record<string, string>>;
}

export type DiffPath = ReadonlyArray<string | number>;

export type DiffChange =
  | {
      kind: "node-inserted";
      nodeId: NodeId;
      path: DiffPath;
      summary: string;
    }
  | {
      kind: "node-deleted";
      nodeId: NodeId;
      path: DiffPath;
      summary: string;
    }
  | {
      kind: "node-updated";
      nodeId: NodeId;
      path: DiffPath;
      field: string;
      summary: string;
    }
  | {
      kind: "node-moved";
      nodeId: NodeId;
      from: DiffPath;
      to: DiffPath;
      summary: string;
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
  toRevision: number,
): DocumentDiff => ({
  format,
  fromRevision,
  toRevision,
  changes: [],
});
