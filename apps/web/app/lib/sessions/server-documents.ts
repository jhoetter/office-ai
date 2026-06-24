import { randomUUID } from "node:crypto";
import { createLocalSessionStore, type LocalSessionStore } from "@officeai/agent/session-store";
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
  return null;
}

export function ensureExtension(name: string | undefined, format: WebOfficeFormat): string {
  const trimmed = name?.trim();
  const base = trimmed && trimmed.length > 0 ? trimmed : `untitled.${format}`;
  return base.toLowerCase().endsWith(`.${format}`) ? base : `${base}.${format}`;
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
