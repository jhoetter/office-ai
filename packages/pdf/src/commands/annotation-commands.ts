import type { CommandHandler } from "@officeai/core";
import type {
  PdfAnnotation,
  PdfAnnotationKind,
  PdfDocument,
  PdfRect,
  PdfSnapshot,
} from "../model/types.js";
import { buildDiff, evolvePdf, makeError } from "./helpers.js";
import type {
  AddAnnotationPayload,
  RemoveAnnotationPayload,
  UpdateAnnotationPayload,
} from "./payloads.js";

const KIND_TO_SUBTYPE: Record<PdfAnnotationKind, string> = {
  highlight: "Highlight",
  underline: "Underline",
  strikethrough: "StrikeOut",
  squiggly: "Squiggly",
  note: "Text",
  "free-text": "FreeText",
  ink: "Ink",
  line: "Line",
  rectangle: "Square",
  ellipse: "Circle",
  polygon: "Polygon",
  polyline: "PolyLine",
  stamp: "Stamp",
  link: "Link",
  redaction: "Redact",
  unknown: "Unknown",
};

const cloneAnnotations = (snapshot: PdfSnapshot): PdfAnnotation[] => [...snapshot.root.annotations];

const findIndex = (annotations: ReadonlyArray<PdfAnnotation>, id: string): number => {
  const idx = annotations.findIndex((a) => a.id === id);
  if (idx < 0) throw makeError("unknown-target", `annotation ${id} not found`);
  return idx;
};

const toRect = (rect: PdfRect): PdfRect => [rect[0], rect[1], rect[2], rect[3]] as PdfRect;

export const addAnnotationHandler: CommandHandler<AddAnnotationPayload, PdfSnapshot> = {
  type: "pdf:add-annotation",
  apply(snapshot, payload, ctx) {
    if (payload.pageNumber < 1 || payload.pageNumber > snapshot.root.pages.length) {
      throw makeError(
        "invalid-payload",
        `pdf:add-annotation page ${payload.pageNumber} out of range (1..${snapshot.root.pages.length})`,
      );
    }
    const id = payload.id ?? ctx.mintNodeId();
    const subtype = KIND_TO_SUBTYPE[payload.kind] ?? "Unknown";
    const annot: PdfAnnotation = {
      id,
      kind: payload.kind,
      subtype,
      pageNumber: payload.pageNumber,
      rect: toRect(payload.rect),
      source: "session",
      createdAt: new Date(ctx.now()).toISOString(),
      ...(payload.contents !== undefined ? { contents: payload.contents } : {}),
      ...(payload.author !== undefined ? { author: payload.author } : {}),
      ...(payload.color !== undefined ? { color: payload.color } : {}),
      ...(payload.quadRects !== undefined
        ? { quadRects: payload.quadRects.map(toRect) }
        : {}),
    };
    const annotations = [...snapshot.root.annotations, annot];
    const pages = snapshot.root.pages.map((p) =>
      p.pageNumber === payload.pageNumber ? { ...p, hasAnnotations: true } : p,
    );
    const root: PdfDocument = { ...snapshot.root, annotations, pages };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: id,
        path: ["annotations", annotations.length - 1],
        summary: `annotation:${payload.kind}`,
      }),
    };
  },
};

export const updateAnnotationHandler: CommandHandler<UpdateAnnotationPayload, PdfSnapshot> = {
  type: "pdf:update-annotation",
  apply(snapshot, payload) {
    const annotations = cloneAnnotations(snapshot);
    const idx = findIndex(annotations, payload.annotationId);
    const current = annotations[idx];
    annotations[idx] = {
      ...current,
      ...(payload.contents !== undefined ? { contents: payload.contents } : {}),
      ...(payload.color !== undefined ? { color: payload.color } : {}),
      ...(payload.rect !== undefined ? { rect: toRect(payload.rect) } : {}),
    };
    const root: PdfDocument = { ...snapshot.root, annotations };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: current.id,
        path: ["annotations", idx],
        field: "contents",
        summary: "annotation",
      }),
    };
  },
};

export const removeAnnotationHandler: CommandHandler<RemoveAnnotationPayload, PdfSnapshot> = {
  type: "pdf:remove-annotation",
  apply(snapshot, payload) {
    const annotations = cloneAnnotations(snapshot);
    const idx = findIndex(annotations, payload.annotationId);
    const [removed] = annotations.splice(idx, 1);
    if (removed.source === "loaded") {
      // Loaded annotations live in the source PDF byte stream and the
      // serializer cannot strip them yet (it would require rewriting
      // page /Annots arrays in pdf-lib). Keeping the projection
      // consistent across save/reload is more important than
      // pretending we removed it.
      throw makeError(
        "unsupported",
        "Removing annotations that already exist in the source PDF isn't supported yet — only session-added annotations can be deleted.",
      );
    }
    const stillOnPage = annotations.some((a) => a.pageNumber === removed.pageNumber);
    const pages = stillOnPage
      ? snapshot.root.pages
      : snapshot.root.pages.map((p) =>
          p.pageNumber === removed.pageNumber ? { ...p, hasAnnotations: false } : p,
        );
    const root: PdfDocument = { ...snapshot.root, annotations, pages };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: removed.id,
        path: ["annotations", idx],
        summary: "annotation",
      }),
    };
  },
};
