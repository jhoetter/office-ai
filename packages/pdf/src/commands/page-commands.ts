import type { CommandHandler } from "@officeai/core";
import type { PdfDocument, PdfPage, PdfRotation, PdfSnapshot } from "../model/types.js";
import { buildDiff, evolvePdf, makeError, validatePages } from "./helpers.js";
import type {
  DeletePagesPayload,
  ReorderPagesPayload,
  RotatePagesPayload,
  SetPageRotationPayload,
} from "./payloads.js";

const normalizeRotation = (raw: number): PdfRotation => {
  const r = ((raw % 360) + 360) % 360;
  if (r === 0 || r === 90 || r === 180 || r === 270) return r as PdfRotation;
  return 0;
};

const renumberPages = (pages: ReadonlyArray<PdfPage>): PdfPage[] =>
  pages.map((p, i) => ({ ...p, pageNumber: i + 1 }));

export const rotatePagesHandler: CommandHandler<RotatePagesPayload, PdfSnapshot> = {
  type: "pdf:rotate-pages",
  apply(snapshot, payload) {
    if (![90, 180, 270, -90, -180, -270].includes(payload.delta)) {
      throw makeError("invalid-rotation", `delta ${payload.delta} is not a multiple of 90`);
    }
    validatePages(snapshot, payload.pages, "pdf:rotate-pages");

    const targetSet = new Set(payload.pages);
    const nextPages: PdfPage[] = snapshot.root.pages.map((p) =>
      targetSet.has(p.pageNumber)
        ? { ...p, rotation: normalizeRotation(p.rotation + payload.delta) }
        : p,
    );
    const root: PdfDocument = { ...snapshot.root, pages: nextPages };
    const next = evolvePdf(snapshot, root);

    const changes = payload.pages.map((pageNumber) => {
      const original = snapshot.root.pages.find((p) => p.pageNumber === pageNumber)!;
      return {
        kind: "node-updated" as const,
        nodeId: original.id,
        path: ["pages", pageNumber - 1, "rotation"],
        field: "rotation",
        summary: `page ${pageNumber} rotation`,
      };
    });
    return { next, diff: buildDiff(snapshot.revision, next.revision, ...changes) };
  },
};

export const setPageRotationHandler: CommandHandler<SetPageRotationPayload, PdfSnapshot> = {
  type: "pdf:set-page-rotation",
  apply(snapshot, payload) {
    validatePages(snapshot, [payload.pageNumber], "pdf:set-page-rotation");
    const idx = payload.pageNumber - 1;
    const original = snapshot.root.pages[idx];
    const nextPages = [...snapshot.root.pages];
    nextPages[idx] = { ...original, rotation: payload.rotation };
    const root: PdfDocument = { ...snapshot.root, pages: nextPages };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: original.id,
        path: ["pages", idx, "rotation"],
        field: "rotation",
        summary: `page ${payload.pageNumber} rotation`,
      }),
    };
  },
};

export const reorderPagesHandler: CommandHandler<ReorderPagesPayload, PdfSnapshot> = {
  type: "pdf:reorder-pages",
  apply(snapshot, payload) {
    const len = snapshot.root.pages.length;
    if (payload.order.length !== len) {
      throw makeError("invalid-order", `order length ${payload.order.length} != page count ${len}`);
    }
    const seen = new Set<number>();
    for (const n of payload.order) {
      if (!Number.isInteger(n) || n < 1 || n > len) {
        throw makeError("invalid-order", `order entry ${n} out of range (1..${len})`);
      }
      if (seen.has(n)) throw makeError("invalid-order", `order entry ${n} duplicated`);
      seen.add(n);
    }
    const reordered = renumberPages(payload.order.map((p) => snapshot.root.pages[p - 1]));
    const root: PdfDocument = { ...snapshot.root, pages: reordered };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: "pages",
        path: ["pages"],
        field: "order",
        summary: "page order",
      }),
    };
  },
};

export const deletePagesHandler: CommandHandler<DeletePagesPayload, PdfSnapshot> = {
  type: "pdf:delete-pages",
  apply(snapshot, payload) {
    if (payload.pages.length === 0) {
      throw makeError("invalid-pages", "pdf:delete-pages requires at least one page");
    }
    if (payload.pages.length >= snapshot.root.pages.length) {
      throw makeError("would-empty-document", "cannot delete every page");
    }
    validatePages(snapshot, payload.pages, "pdf:delete-pages");
    const drop = new Set(payload.pages);
    const removed = snapshot.root.pages.filter((p) => drop.has(p.pageNumber));
    const remaining = renumberPages(snapshot.root.pages.filter((p) => !drop.has(p.pageNumber)));
    const root: PdfDocument = { ...snapshot.root, pages: remaining };
    const next = evolvePdf(snapshot, root);
    const changes = removed.map((p) => ({
      kind: "node-deleted" as const,
      nodeId: p.id,
      path: ["pages", p.pageNumber - 1],
      summary: `page ${p.pageNumber}`,
    }));
    return { next, diff: buildDiff(snapshot.revision, next.revision, ...changes) };
  },
};
