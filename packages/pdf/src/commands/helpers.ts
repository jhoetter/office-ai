import { CommandError, type DiffChange, type DocumentDiff } from "@officeai/core";
import type { PdfDocument, PdfSnapshot } from "../model/types.js";

export const evolvePdf = (snapshot: PdfSnapshot, root: PdfDocument): PdfSnapshot => ({
  ...snapshot,
  revision: snapshot.revision + 1,
  root,
});

export const buildDiff = (
  fromRevision: number,
  toRevision: number,
  ...changes: DiffChange[]
): DocumentDiff => ({
  format: "pdf",
  fromRevision,
  toRevision,
  changes,
});

export const makeError = (code: string, message: string): CommandError => new CommandError(code, message);

export const validatePages = (snapshot: PdfSnapshot, pages: ReadonlyArray<number>, label: string): void => {
  const total = snapshot.root.pages.length;
  for (const p of pages) {
    if (!Number.isInteger(p) || p < 1 || p > total) {
      throw makeError("invalid-page", `${label}: page ${p} out of range (1..${total})`);
    }
  }
};
