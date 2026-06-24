import type {
  StoredCommandExportRef,
  StoredCommandLogEntry,
  StoredCommandProvenance,
  StoredDocumentRecord,
  StoredExportRecord,
  StoredPendingChange,
  StoredSessionRecord,
} from "@officeai/agent/session-store";

export type WebOfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";

export interface WebSessionEntry {
  readonly sessionId: string;
  readonly title: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly documentCount: number;
}

export interface WebDiagnosticEntry {
  readonly level: string;
  readonly code: string;
  readonly message: string;
}

export interface WebDocumentEntry {
  readonly documentId: string;
  readonly sessionId: string;
  readonly format: WebOfficeFormat;
  readonly name: string;
  readonly status: "ready" | "error";
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly revision: number;
  readonly diagnostics: ReadonlyArray<WebDiagnosticEntry>;
  readonly exportCount: number;
  readonly pendingChangeCount: number;
  readonly commandLogCount: number;
  readonly artifacts: {
    readonly hasOriginal: boolean;
    readonly hasWorking: boolean;
  };
}

export interface WebExportEntry {
  readonly bytes: number;
  readonly exportedAt: string;
}

export interface WebPendingChangeEntry {
  readonly id: string;
  readonly operation: string;
  readonly status: string;
  readonly source: string;
  readonly actorId?: string;
  readonly timestamp?: number;
  readonly hasDiff: boolean;
  readonly diffSummary?: string;
  readonly rejection?: {
    readonly code: string;
    readonly message: string;
  };
}

export interface WebCommandLogEntry {
  readonly id: string;
  readonly commandId?: string;
  readonly operation: string;
  readonly status: string;
  readonly stage: string;
  readonly source: string;
  readonly actorId?: string;
  readonly recordedAt: string;
  readonly hasDiff: boolean;
  readonly diagnostics: ReadonlyArray<WebDiagnosticEntry>;
  readonly provenance?: StoredCommandProvenance;
  readonly exportRef?: StoredCommandExportRef;
}

export interface WebDocumentDetailEntry extends WebDocumentEntry {
  readonly exports: ReadonlyArray<WebExportEntry>;
  readonly pendingChanges: ReadonlyArray<WebPendingChangeEntry>;
  readonly commandLog: ReadonlyArray<WebCommandLogEntry>;
}

export interface WebSessionsPayload {
  readonly schema: "office-ai/web-sessions@1";
  readonly sessions: ReadonlyArray<WebSessionEntry>;
  readonly documents: ReadonlyArray<WebDocumentEntry>;
}

export interface WebDocumentPayload {
  readonly schema: "office-ai/web-document@1";
  readonly session: WebSessionEntry;
  readonly document: WebDocumentDetailEntry;
}

export function documentsForSession(
  payload: WebSessionsPayload,
  sessionId: string
): ReadonlyArray<WebDocumentEntry> {
  return payload.documents.filter((document) => document.sessionId === sessionId);
}

export function sessionBrowserCounts(payload: WebSessionsPayload): {
  readonly sessions: number;
  readonly documents: number;
  readonly pending: number;
  readonly diagnostics: number;
} {
  return {
    sessions: payload.sessions.length,
    documents: payload.documents.length,
    pending: payload.documents.reduce((sum, document) => sum + document.pendingChangeCount, 0),
    diagnostics: payload.documents.reduce((sum, document) => sum + document.diagnostics.length, 0),
  };
}

export function toWebSessionEntry(session: StoredSessionRecord): WebSessionEntry {
  return {
    sessionId: session.id,
    title: session.title,
    createdAt: session.createdAt,
    updatedAt: session.updatedAt,
    documentCount: session.documentIds.length,
  };
}

export function toWebDocumentEntry(document: StoredDocumentRecord): WebDocumentEntry {
  return {
    documentId: document.id,
    sessionId: document.sessionId,
    format: document.format,
    name: document.name,
    status: document.status,
    createdAt: document.createdAt,
    updatedAt: document.updatedAt,
    revision: document.revision,
    diagnostics: document.diagnostics,
    exportCount: document.exportHistory.length,
    pendingChangeCount: document.pendingChanges.filter((change) => change.status === "pending").length,
    commandLogCount: document.commandLog.length,
    artifacts: {
      hasOriginal: Boolean(document.artifacts.originalPath),
      hasWorking: Boolean(document.artifacts.workingPath),
    },
  };
}

export function toWebDocumentDetailEntry(document: StoredDocumentRecord): WebDocumentDetailEntry {
  return {
    ...toWebDocumentEntry(document),
    exports: document.exportHistory.map(toWebExportEntry),
    pendingChanges: document.pendingChanges.map(toWebPendingChangeEntry),
    commandLog: document.commandLog.map(toWebCommandLogEntry),
  };
}

function toWebExportEntry(record: StoredExportRecord): WebExportEntry {
  return {
    bytes: record.bytes,
    exportedAt: record.exportedAt,
  };
}

function toWebPendingChangeEntry(change: StoredPendingChange): WebPendingChangeEntry {
  return {
    id: change.id,
    operation: change.operation,
    status: change.status,
    source: change.source,
    ...(change.actorId ? { actorId: change.actorId } : {}),
    ...(change.timestamp !== undefined ? { timestamp: change.timestamp } : {}),
    hasDiff: change.diff !== undefined,
    ...(change.diff !== undefined ? { diffSummary: diffSummary(change.diff) } : {}),
    ...(change.rejection ? { rejection: change.rejection } : {}),
  };
}

function toWebCommandLogEntry(entry: StoredCommandLogEntry): WebCommandLogEntry {
  return {
    id: entry.id,
    ...(entry.commandId ? { commandId: entry.commandId } : {}),
    operation: entry.operation,
    status: entry.status,
    stage: entry.stage,
    source: entry.source,
    ...(entry.actorId ? { actorId: entry.actorId } : {}),
    recordedAt: entry.recordedAt,
    hasDiff: entry.diff !== undefined,
    diagnostics: entry.diagnostics ?? [],
    ...(entry.provenance ? { provenance: entry.provenance } : {}),
    ...(entry.exportRef ? { exportRef: entry.exportRef } : {}),
  };
}

function diffSummary(diff: unknown): string {
  if (
    diff &&
    typeof diff === "object" &&
    (diff as { schema?: unknown }).schema === "office-ai/semantic-diff@1" &&
    typeof (diff as { summary?: { text?: unknown } }).summary?.text === "string"
  ) {
    return (diff as { summary: { text: string } }).summary.text;
  }
  return "Structured diff available";
}
