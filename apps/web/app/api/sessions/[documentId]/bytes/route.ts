/**
 * GET/PUT /api/sessions/:documentId/bytes
 *
 * Editor-facing raw working artifact endpoint. Metadata endpoints stay
 * path-free; this route returns only the persisted document bytes plus
 * revision headers needed for optimistic save-back.
 */

import { randomUUID } from "node:crypto";
import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";
import {
  contentDisposition,
  ensureExtension,
  mimeForFormat,
  prepareImportedBytes,
} from "@/lib/sessions/server-documents";
import { toWebDocumentDetailEntry } from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(
  _request: Request,
  context: { readonly params: Promise<{ readonly documentId: string }> }
): Promise<NextResponse> {
  try {
    const { documentId } = await context.params;
    const store = createLocalSessionStore();
    const document = await store.getDocument(documentId);
    if (!document) return documentNotFound(documentId);

    const filename = ensureExtension(document.name, document.format);
    const bytes = await store.readWorkingBytes(document);
    return new NextResponse(Buffer.from(bytes), {
      status: 200,
      headers: {
        "content-type": mimeForFormat(document.format),
        "content-disposition": contentDisposition(filename),
        "cache-control": "no-store",
        etag: revisionEtag(document.id, document.revision),
        "x-officeai-document-id": document.id,
        "x-officeai-session-id": document.sessionId,
        "x-officeai-format": document.format,
        "x-officeai-filename": encodeURIComponent(filename),
        "x-officeai-revision": String(document.revision),
      },
    });
  } catch (err) {
    return sessionError(err, "session-bytes-error");
  }
}

export async function PUT(
  request: Request,
  context: { readonly params: Promise<{ readonly documentId: string }> }
): Promise<NextResponse> {
  try {
    const { documentId } = await context.params;
    const store = createLocalSessionStore();
    const document = await store.getDocument(documentId);
    if (!document) return documentNotFound(documentId);

    const lock = readRevisionLock(request);
    if (!revisionLockMatches(lock, document.id, document.revision)) {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "stale-revision",
          message: `Document ${document.id} is at revision ${document.revision}; reload before saving.`,
          documentId: document.id,
          currentRevision: document.revision,
        },
        { status: 409, headers: { etag: revisionEtag(document.id, document.revision) } }
      );
    }

    const bytes = new Uint8Array(await request.arrayBuffer());
    if (bytes.byteLength === 0) {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "empty-document-bytes",
          message: "Session save received an empty document body.",
        },
        { status: 400 }
      );
    }
    await prepareImportedBytes(document.format, bytes);

    const now = new Date().toISOString();
    const filename = ensureExtension(
      decodeHeaderFilename(request.headers.get("x-officeai-filename")) ?? document.name,
      document.format
    );
    const nextRevision = document.revision + 1;
    const diagnostics = [
      {
        level: "info" as const,
        code: "web-editor-save",
        message: `Saved ${filename} from the web editor at revision ${nextRevision}.`,
      },
    ];
    const updated = await store.putDocument(
      {
        id: document.id,
        sessionId: document.sessionId,
        format: document.format,
        name: filename,
        status: "ready",
        ...(document.sourcePath ? { sourcePath: document.sourcePath } : {}),
        createdAt: document.createdAt,
        updatedAt: now,
        revision: nextRevision,
        diagnostics: [...document.diagnostics, ...diagnostics],
        exportHistory: document.exportHistory,
        pendingChanges: document.pendingChanges,
        commandLog: [
          ...document.commandLog,
          {
            schema: "office-ai/audit-log-entry@1",
            schemaVersion: 1,
            id: `log_${randomUUID()}`,
            operation: "save_document",
            status: "applied",
            stage: "saved",
            source: "web",
            recordedAt: now,
            diagnostics,
            provenance: {
              surface: "web",
              sessionId: document.sessionId,
              documentId: document.id,
              targetRevision: document.revision,
              argumentsSummary: `${bytes.byteLength} bytes`,
            },
          },
        ],
      },
      { workingBytes: bytes }
    );
    const session = await store.getSession(document.sessionId);
    await store.putSession({
      id: session.id,
      title: session.title,
      createdAt: session.createdAt,
      updatedAt: now,
      documentIds: session.documentIds,
    });

    return NextResponse.json(
      {
        schema: "office-ai/session-bytes-save@1",
        document: toWebDocumentDetailEntry(updated),
        etag: revisionEtag(updated.id, updated.revision),
      },
      { status: 200, headers: { etag: revisionEtag(updated.id, updated.revision) } }
    );
  } catch (err) {
    return sessionError(err, "session-save-error");
  }
}

function documentNotFound(documentId: string): NextResponse {
  return NextResponse.json(
    {
      schema: "office-ai/web-sessions-error@1",
      code: "document-not-found",
      message: `Document ${documentId} was not found in the local workspace.`,
    },
    { status: 404 }
  );
}

function sessionError(err: unknown, code: string): NextResponse {
  const message = err instanceof Error ? err.message : String(err);
  if (err instanceof SessionStoreCorruptError) {
    return NextResponse.json(
      {
        schema: "office-ai/web-sessions-error@1",
        code: "corrupt-session-store",
        message,
      },
      { status: 409 }
    );
  }
  return NextResponse.json(
    {
      schema: "office-ai/web-sessions-error@1",
      code,
      message,
    },
    { status: 500 }
  );
}

function revisionEtag(documentId: string, revision: number): string {
  return `"officeai:${documentId}:${revision}"`;
}

function readRevisionLock(request: Request): string | number | null {
  const ifMatch = request.headers.get("if-match");
  if (ifMatch) return ifMatch;
  const rawRevision = request.headers.get("x-officeai-base-revision");
  if (!rawRevision) return null;
  const parsed = Number.parseInt(rawRevision, 10);
  return Number.isInteger(parsed) ? parsed : null;
}

function revisionLockMatches(lock: string | number | null, documentId: string, revision: number): boolean {
  if (lock === null) return false;
  if (typeof lock === "number") return lock === revision;
  return lock === revisionEtag(documentId, revision) || lock === "*";
}

function decodeHeaderFilename(value: string | null): string | null {
  if (!value) return null;
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}
