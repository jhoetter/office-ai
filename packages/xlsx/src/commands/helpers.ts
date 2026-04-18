import type { DocumentDiff, DiffChange } from "@officeai/core";
import type { Sheet, XlsxDirtyFlags, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";

/**
 * Build a new snapshot from `prev` with a mutated workbook. Increments
 * `revision` and merges in the supplied dirty flags. Mirrors
 * `evolveSnapshot` from `@officeai/docx`.
 */
export function evolveSnapshot(
  prev: XlsxSnapshot,
  next: XlsxWorkbook,
  dirtyPatch: PartialDirtyFlags
): XlsxSnapshot {
  return {
    format: "xlsx",
    revision: prev.revision + 1,
    root: next,
    partHashes: prev.partHashes,
    container: prev.container,
    dirty: mergeDirty(prev.dirty, dirtyPatch),
  };
}

export type PartialDirtyFlags = Partial<{
  workbook: boolean;
  sharedStrings: boolean;
  styles: boolean;
  contentTypes: boolean;
  rels: boolean;
  /** Sheet part paths to add to the dirty set. */
  sheets: ReadonlyArray<string>;
  comments: ReadonlyArray<string>;
  threadedComments: ReadonlyArray<string>;
  sheetRels: ReadonlyArray<string>;
  /** Sheet part paths the serializer must drop entirely (delete-sheet). */
  removedSheetParts: ReadonlyArray<string>;
  /** Sheet part paths whose drawing part needs re-emitting. */
  drawings: ReadonlyArray<string>;
  /** Media part paths to (re)emit. */
  media: ReadonlyArray<string>;
  /** Media part paths the serializer must drop entirely. */
  removedMediaParts: ReadonlyArray<string>;
}>;

export function mergeDirty(prev: XlsxDirtyFlags, patch: PartialDirtyFlags): XlsxDirtyFlags {
  const sheets = new Set(prev.sheets);
  for (const p of patch.sheets ?? []) sheets.add(p);
  const comments = new Set(prev.comments);
  for (const p of patch.comments ?? []) comments.add(p);
  const threadedComments = new Set(prev.threadedComments);
  for (const p of patch.threadedComments ?? []) threadedComments.add(p);
  const sheetRels = new Set(prev.sheetRels);
  for (const p of patch.sheetRels ?? []) sheetRels.add(p);
  const removedSheetParts = new Set(prev.removedSheetParts);
  for (const p of patch.removedSheetParts ?? []) removedSheetParts.add(p);
  const drawings = new Set(prev.drawings);
  for (const p of patch.drawings ?? []) drawings.add(p);
  const media = new Set(prev.media);
  for (const p of patch.media ?? []) media.add(p);
  const removedMediaParts = new Set(prev.removedMediaParts);
  for (const p of patch.removedMediaParts ?? []) removedMediaParts.add(p);
  return {
    workbook: patch.workbook ?? prev.workbook,
    sharedStrings: patch.sharedStrings ?? prev.sharedStrings,
    styles: patch.styles ?? prev.styles,
    contentTypes: patch.contentTypes ?? prev.contentTypes,
    rels: patch.rels ?? prev.rels,
    sheets,
    comments,
    threadedComments,
    sheetRels,
    removedSheetParts,
    drawings,
    media,
    removedMediaParts,
  };
}

/** Locate a sheet by name. Case-sensitive (matches Excel's lookup). */
export function findSheet(workbook: XlsxWorkbook, name: string): Sheet | undefined {
  return workbook.sheets.find((s) => s.name === name);
}

export function replaceSheet(workbook: XlsxWorkbook, next: Sheet): XlsxWorkbook {
  const sheets = workbook.sheets.slice();
  sheets[next.index] = next;
  return { ...workbook, sheets };
}

export function buildDiff(
  prevRevision: number,
  nextRevision: number,
  changes: ReadonlyArray<DiffChange>
): DocumentDiff {
  return {
    format: "xlsx",
    fromRevision: prevRevision,
    toRevision: nextRevision,
    changes: [...changes],
  };
}
