/**
 * POST /api/sessions/create
 *
 * Create a blank DOCX/XLSX/PPTX/PDF document in the same local data-dir used
 * by MCP. This is the persisted web counterpart to the transient editor
 * "new file" entry points.
 */

import { randomUUID } from "node:crypto";
import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";
import {
  createBlankDocumentBytes,
  ensureExtension,
  sessionForNewDocument,
} from "@/lib/sessions/server-documents";
import {
  toWebDocumentDetailEntry,
  toWebSessionEntry,
  type WebOfficeFormat,
} from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request): Promise<NextResponse> {
  try {
    const rawBody = await request.json().catch(() => ({}));
    const body =
      rawBody && typeof rawBody === "object" && !Array.isArray(rawBody)
        ? (rawBody as Record<string, unknown>)
        : {};
    const format = isSupportedFormat(body.format) ? body.format : null;
    if (!format) {
      return badRequest("unsupported-format", "format must be one of docx, xlsx, pptx or pdf.");
    }

    const now = new Date().toISOString();
    const store = createLocalSessionStore();
    const documentId = `doc_${randomUUID()}`;
    const name = ensureExtension(typeof body.name === "string" ? body.name : undefined, format);
    const session = await sessionForNewDocument({
      store,
      sessionId: typeof body.session_id === "string" ? body.session_id : undefined,
      documentId,
      title: typeof body.title === "string" ? body.title : "Web creates",
      now,
    });
    const prepared = await createBlankDocumentBytes(format);
    const diagnostics = [
      { level: "info" as const, code: "created", message: `Created blank ${format} document.` },
    ];

    await store.putSession(session);
    const document = await store.putDocument(
      {
        id: documentId,
        sessionId: session.id,
        format,
        name,
        status: "ready",
        createdAt: now,
        updatedAt: now,
        revision: prepared.revision,
        diagnostics,
        exportHistory: [],
        pendingChanges: [],
        commandLog: [
          {
            schema: "office-ai/audit-log-entry@1",
            schemaVersion: 1,
            id: `log_${randomUUID()}`,
            operation: "create_document",
            status: "applied",
            stage: "created",
            source: "web",
            recordedAt: now,
            diagnostics,
            provenance: {
              surface: "web",
              sessionId: session.id,
              documentId,
              targetRevision: prepared.revision,
            },
          },
        ],
      },
      { workingBytes: prepared.bytes }
    );
    const persistedSession = await store.getSession(session.id);

    return NextResponse.json(
      {
        schema: "office-ai/web-create@1",
        session: toWebSessionEntry(persistedSession),
        document: toWebDocumentDetailEntry(document),
      },
      { status: 201 }
    );
  } catch (err) {
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
        code: "session-create-error",
        message,
      },
      { status: 500 }
    );
  }
}

function badRequest(code: string, message: string): NextResponse {
  return NextResponse.json(
    {
      schema: "office-ai/web-sessions-error@1",
      code,
      message,
    },
    { status: 400 }
  );
}

function isSupportedFormat(value: unknown): value is WebOfficeFormat {
  return value === "docx" || value === "xlsx" || value === "pptx" || value === "pdf";
}
