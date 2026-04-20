import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import type { PdfDocument, PdfPage, PdfSnapshot } from "../model/types.js";
import {
  deletePagesHandler,
  reorderPagesHandler,
  rotatePagesHandler,
  setPageRotationHandler,
} from "./page-commands.js";

const makePage = (n: number): PdfPage => ({
  id: `page-${n}`,
  pageNumber: n,
  sourceIndex: n - 1,
  width: 612,
  height: 792,
  rotation: 0,
  text: `page ${n} content`,
  structured: { pageWidth: 612, pageHeight: 792, blocks: [], columnCount: 1 },
  hasTextLayer: true,
  hasAnnotations: false,
  hasFormFields: false,
});

const makeSnapshot = (pageCount: number): PdfSnapshot => {
  const pages: PdfPage[] = Array.from({ length: pageCount }, (_, i) => makePage(i + 1));
  const root: PdfDocument = {
    metadata: {},
    pages,
    outline: [],
    annotations: [],
    formFields: [],
    attachments: [],
    comments: [],
    signatureCount: 0,
    engineKind: "pdfjs",
  };
  return { format: "pdf", revision: 0, root, partHashes: { "pdf-bytes": "deadbeef" } };
};

const ctx = { mintNodeId: deterministicIdMinter("t"), now: () => 1700000000000 };

describe("rotatePagesHandler", () => {
  it("adds rotation modulo 360 to the requested pages only", () => {
    const snap = makeSnapshot(3);
    const { next, diff } = rotatePagesHandler.apply(snap, { pages: [1, 3], delta: 90 }, ctx);
    expect(next.root.pages[0].rotation).toBe(90);
    expect(next.root.pages[1].rotation).toBe(0);
    expect(next.root.pages[2].rotation).toBe(90);
    expect(diff.changes.length).toBe(2);
    expect(next.revision).toBe(1);
  });

  it("rejects deltas that are not multiples of 90", () => {
    const snap = makeSnapshot(2);
    expect(() => rotatePagesHandler.apply(snap, { pages: [1], delta: 45 as unknown as 90 }, ctx)).toThrow(
      /multiple of 90/
    );
  });

  it("rejects pages that are out of range", () => {
    const snap = makeSnapshot(2);
    expect(() => rotatePagesHandler.apply(snap, { pages: [3], delta: 90 }, ctx)).toThrow(/out of range/);
  });
});

describe("setPageRotationHandler", () => {
  it("sets a single page rotation", () => {
    const snap = makeSnapshot(2);
    const { next } = setPageRotationHandler.apply(snap, { pageNumber: 2, rotation: 270 }, ctx);
    expect(next.root.pages[1].rotation).toBe(270);
    expect(next.root.pages[0].rotation).toBe(0);
  });
});

describe("reorderPagesHandler", () => {
  it("reorders pages and renumbers them", () => {
    const snap = makeSnapshot(3);
    const { next } = reorderPagesHandler.apply(snap, { order: [3, 1, 2] }, ctx);
    expect(next.root.pages.map((p) => p.text)).toEqual([
      "page 3 content",
      "page 1 content",
      "page 2 content",
    ]);
    expect(next.root.pages.map((p) => p.pageNumber)).toEqual([1, 2, 3]);
  });

  it("rejects partial permutations", () => {
    const snap = makeSnapshot(3);
    expect(() => reorderPagesHandler.apply(snap, { order: [1, 2] }, ctx)).toThrow();
    expect(() => reorderPagesHandler.apply(snap, { order: [1, 1, 2] }, ctx)).toThrow();
    expect(() => reorderPagesHandler.apply(snap, { order: [1, 2, 4] }, ctx)).toThrow();
  });
});

describe("deletePagesHandler", () => {
  it("removes the requested pages and renumbers the survivors", () => {
    const snap = makeSnapshot(4);
    const { next, diff } = deletePagesHandler.apply(snap, { pages: [2, 4] }, ctx);
    expect(next.root.pages.map((p) => p.text)).toEqual(["page 1 content", "page 3 content"]);
    expect(next.root.pages.map((p) => p.pageNumber)).toEqual([1, 2]);
    expect(diff.changes.length).toBe(2);
  });

  it("refuses to delete every page", () => {
    const snap = makeSnapshot(2);
    expect(() => deletePagesHandler.apply(snap, { pages: [1, 2] }, ctx)).toThrow(/every page/);
  });
});
