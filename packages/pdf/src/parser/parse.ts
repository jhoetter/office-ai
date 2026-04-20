import { defaultIdMinter, sha256Hex, type DocumentFormat } from "@officeai/core";
import {
  loadDocument,
  type PdfEngineDocument,
  type PdfEngineKind,
  type PdfEngineLoadOptions,
  type EngineSelectionHints,
} from "@officeai/pdf-engine";
import { selectEngine } from "@officeai/pdf-engine";
import type {
  PdfAnnotation,
  PdfAnnotationKind,
  PdfDocument,
  PdfFormField,
  PdfMetadata,
  PdfOutlineNode,
  PdfPage,
  PdfRect,
  PdfRotation,
  PdfSnapshot,
} from "../model/types.js";
import { PdfParseError } from "./errors.js";

export interface PdfParseOptions extends PdfEngineLoadOptions {
  /** Engine selection hints (overridden by `forceEngine`). */
  hints?: EngineSelectionHints;
  /** Optional id minter (deterministic for tests). */
  idMinter?: () => string;
}

const FORMAT: DocumentFormat = "pdf";

const ANNOT_KIND_BY_SUBTYPE: Record<string, PdfAnnotationKind> = {
  Highlight: "highlight",
  Underline: "underline",
  StrikeOut: "strikethrough",
  Squiggly: "squiggly",
  Text: "note",
  FreeText: "free-text",
  Ink: "ink",
  Line: "line",
  Square: "rectangle",
  Circle: "ellipse",
  Polygon: "polygon",
  PolyLine: "polyline",
  Stamp: "stamp",
  Link: "link",
  Redact: "redaction",
};

export const parsePdf = async (
  input: ArrayBuffer | Uint8Array,
  opts: PdfParseOptions = {},
): Promise<PdfSnapshot> => {
  const buffer = input instanceof Uint8Array ? input : new Uint8Array(input);
  const mint = opts.idMinter ?? defaultIdMinter;

  const forceEngine: PdfEngineKind | undefined = opts.forceEngine ?? selectEngineFromHints(opts.hints);

  let doc: PdfEngineDocument;
  try {
    doc = await loadDocument(buffer, {
      ...(opts.password !== undefined ? { password: opts.password } : {}),
      ...(forceEngine !== undefined ? { forceEngine } : {}),
    });
  } catch (err) {
    throw new PdfParseError(
      err instanceof Error ? `Failed to parse PDF: ${err.message}` : "Failed to parse PDF",
      err,
    );
  }

  try {
    const metadata: PdfMetadata = await doc.getMetadata();
    const pages: PdfPage[] = [];
    const annotations: PdfAnnotation[] = [];
    const formFields: PdfFormField[] = [];

    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const text = await page.getTextContent();
      const annots = await page.getAnnotations();
      const fields = await page.getFormFields();

      pages.push({
        id: mint(),
        pageNumber: i,
        width: page.info.width,
        height: page.info.height,
        rotation: page.info.rotation as PdfRotation,
        ...(page.info.label !== undefined ? { label: page.info.label } : {}),
        text: text.plain,
        hasTextLayer: text.items.length > 0,
        hasAnnotations: annots.length > 0,
        hasFormFields: fields.length > 0,
      });

      for (const a of annots) {
        const kind: PdfAnnotationKind = ANNOT_KIND_BY_SUBTYPE[a.subtype] ?? "unknown";
        const annot: PdfAnnotation = {
          id: mint(),
          kind,
          subtype: a.subtype,
          pageNumber: i,
          rect: [a.rect[0], a.rect[1], a.rect[2], a.rect[3]] as PdfRect,
          ...(a.contents !== undefined ? { contents: a.contents } : {}),
          ...(a.author !== undefined ? { author: a.author } : {}),
          ...(a.url !== undefined ? { url: a.url } : {}),
          ...(a.destPage !== undefined ? { destPage: a.destPage } : {}),
        };
        annotations.push(annot);
      }

      for (const f of fields) {
        const field: PdfFormField = {
          id: mint(),
          name: f.name,
          type: f.type,
          ...(f.value !== undefined ? { value: f.value } : {}),
          ...(f.options !== undefined ? { options: f.options } : {}),
          readOnly: f.readOnly,
          required: f.required,
          ...(f.maxLength !== undefined ? { maxLength: f.maxLength } : {}),
          ...(f.multiline !== undefined ? { multiline: f.multiline } : {}),
          ...(f.password !== undefined ? { password: f.password } : {}),
          pageNumber: f.pageNumber,
          rect: [f.rect[0], f.rect[1], f.rect[2], f.rect[3]] as PdfRect,
        };
        formFields.push(field);
      }

      page.destroy();
    }

    const rawOutline = await doc.getOutline();
    const outline: PdfOutlineNode[] = rawOutline ? mapOutline(rawOutline, mint) : [];

    const rawAttachments = await doc.getAttachments();
    const attachments = rawAttachments.map((a) => ({
      id: mint(),
      name: a.name,
      bytes: a.data.byteLength,
    }));

    const signatureCount = formFields.filter((f) => f.type === "signature").length;

    const document: PdfDocument = {
      metadata,
      pages,
      outline,
      annotations,
      formFields,
      attachments,
      comments: [],
      signatureCount,
      engineKind: doc.engine,
    };

    const partHashes: Record<string, string> = {
      "pdf-bytes": sha256Hex(buffer),
    };

    return {
      format: "pdf",
      revision: 0,
      root: document,
      partHashes,
    } satisfies PdfSnapshot;
  } finally {
    await doc.destroy();
  }
};

const mapOutline = (
  nodes: ReadonlyArray<{
    title: string;
    pageNumber?: number;
    uri?: string;
    children: ReadonlyArray<{ title: string; pageNumber?: number; uri?: string; children: ReadonlyArray<unknown> }>;
  }>,
  mint: () => string,
): PdfOutlineNode[] =>
  nodes.map((n) => ({
    id: mint(),
    title: n.title,
    ...(n.pageNumber !== undefined ? { pageNumber: n.pageNumber } : {}),
    ...(n.uri !== undefined ? { uri: n.uri } : {}),
    children: mapOutline(n.children as never, mint),
  }));

const selectEngineFromHints = (hints?: EngineSelectionHints): PdfEngineKind | undefined => {
  if (!hints) return undefined;
  return selectEngine(hints);
};

void FORMAT;
