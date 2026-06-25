import { randomUUID } from "node:crypto";
import { createLocalSessionStore, type LocalSessionStore } from "@officeai/agent/session-store";
import type { ProjectionSource } from "@officeai/agent/projections";
import { DocxAgent } from "@officeai/docx";
import { PdfAgent } from "@officeai/pdf";
import { PptxAgent } from "@officeai/pptx";
import { XlsxAgent } from "@officeai/xlsx";
import type { WebOfficeFormat } from "./web-sessions";

export interface PreparedDocumentBytes {
  readonly bytes: Uint8Array;
  readonly revision: number;
}

export function inferFormatFromName(name: string): WebOfficeFormat | null {
  const lower = name.toLowerCase();
  if (lower.endsWith(".docx")) return "docx";
  if (lower.endsWith(".xlsx")) return "xlsx";
  if (lower.endsWith(".pptx")) return "pptx";
  if (lower.endsWith(".pdf")) return "pdf";
  if (lower.endsWith(".eml") || lower.endsWith(".msg")) return "email";
  if (/\.(png|jpe?g|webp|gif|svg|bmp|tiff?|heic|heif)$/i.test(lower)) return "image";
  return null;
}

export function ensureExtension(name: string | undefined, format: WebOfficeFormat): string {
  const trimmed = name?.trim();
  const fallback =
    format === "email" ? "untitled.eml" : format === "image" ? "untitled.png" : `untitled.${format}`;
  const base = trimmed && trimmed.length > 0 ? trimmed : fallback;
  if (format === "email" && /\.(eml|msg)$/i.test(base)) return base;
  if (format === "image" && /\.(png|jpe?g|webp|gif|svg|bmp|tiff?|heic|heif)$/i.test(base)) return base;
  return base.toLowerCase().endsWith(`.${format}`) ? base : `${base}.${defaultExtensionForFormat(format)}`;
}

export function mimeForFormat(format: WebOfficeFormat): string {
  switch (format) {
    case "docx":
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    case "xlsx":
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    case "pptx":
      return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    case "pdf":
      return "application/pdf";
    case "email":
      return "message/rfc822";
    case "image":
      return "application/octet-stream";
  }
}

export function contentDisposition(filename: string): string {
  const ascii = filename.replace(/[^\x20-\x7E]/g, "_").replace(/["\\]/g, "_");
  return `attachment; filename="${ascii || "office-ai-export"}"`;
}

export async function prepareImportedBytes(
  format: WebOfficeFormat,
  bytes: Uint8Array
): Promise<PreparedDocumentBytes> {
  switch (format) {
    case "docx":
      return { bytes, revision: (await DocxAgent.fromBuffer(bytes)).getSnapshot().revision };
    case "xlsx":
      return { bytes, revision: (await XlsxAgent.fromBuffer(bytes)).getSnapshot().revision };
    case "pptx":
      return { bytes, revision: (await PptxAgent.fromBuffer(bytes)).getSnapshot().revision };
    case "pdf":
      return { bytes, revision: (await PdfAgent.fromBuffer(bytes)).getSnapshot().revision };
    case "email":
    case "image":
      if (bytes.byteLength === 0) throw new Error(`${format} import received an empty file.`);
      return { bytes, revision: 1 };
  }
}

export async function createBlankDocumentBytes(format: WebOfficeFormat): Promise<PreparedDocumentBytes> {
  switch (format) {
    case "docx": {
      const agent = await DocxAgent.empty();
      const bytes = asUint8Array(await agent.exportFile());
      return { bytes, revision: agent.getSnapshot().revision };
    }
    case "xlsx": {
      const agent = await XlsxAgent.empty();
      const bytes = asUint8Array(await agent.exportFile());
      return { bytes, revision: agent.getSnapshot().revision };
    }
    case "pptx": {
      const agent = await PptxAgent.empty();
      const bytes = asUint8Array(await agent.exportFile());
      return { bytes, revision: agent.getSnapshot().revision };
    }
    case "pdf": {
      const agent = await PdfAgent.empty();
      const bytes = asUint8Array(await agent.exportFile());
      return { bytes, revision: agent.getSnapshot().revision };
    }
    case "email":
    case "image":
      throw new Error(`Blank ${format} documents are not supported.`);
  }
}

export async function projectionSourceFromBytes(
  format: WebOfficeFormat,
  bytes: Uint8Array
): Promise<ProjectionSource> {
  switch (format) {
    case "docx":
      return { format, agent: await DocxAgent.fromBuffer(bytes) };
    case "xlsx":
      return { format, agent: await XlsxAgent.fromBuffer(bytes) };
    case "pptx":
      return { format, agent: await PptxAgent.fromBuffer(bytes) };
    case "pdf":
      return { format, agent: await PdfAgent.fromBuffer(bytes) };
    case "email":
    case "image":
      throw new Error(`${format} documents do not expose document projections.`);
  }
}

export async function sessionForNewDocument(opts: {
  readonly store?: LocalSessionStore;
  readonly sessionId?: string;
  readonly documentId: string;
  readonly title?: string;
  readonly now: string;
}): Promise<{
  readonly id: string;
  readonly title: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly documentIds: ReadonlyArray<string>;
}> {
  const store = opts.store ?? createLocalSessionStore();
  if (opts.sessionId) {
    const existing = await store.getSession(opts.sessionId);
    return {
      id: existing.id,
      title: existing.title,
      createdAt: existing.createdAt,
      updatedAt: opts.now,
      documentIds: [...new Set([...existing.documentIds, opts.documentId])],
    };
  }
  return {
    id: `session_${randomUUID()}`,
    title: opts.title ?? "Web workspace",
    createdAt: opts.now,
    updatedAt: opts.now,
    documentIds: [opts.documentId],
  };
}

function asUint8Array(bytes: ArrayBuffer | Uint8Array): Uint8Array {
  return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
}

function defaultExtensionForFormat(format: WebOfficeFormat): string {
  switch (format) {
    case "docx":
    case "xlsx":
    case "pptx":
    case "pdf":
      return format;
    case "email":
      return "eml";
    case "image":
      return "png";
  }
}
