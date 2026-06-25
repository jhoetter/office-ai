/**
 * POST /api/sessions/import
 *
 * Browser upload entry point for the same local OfficeAI data-dir used by MCP.
 * The response mirrors the web session detail payload and intentionally omits
 * server-local paths.
 */

import { randomUUID } from "node:crypto";
import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";
import {
  inferFormatFromName,
  prepareImportedBytes,
  sessionForNewDocument,
} from "@/lib/sessions/server-documents";
import { toWebDocumentDetailEntry, toWebSessionEntry } from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request): Promise<NextResponse> {
  try {
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File) || file.size === 0) {
      return badRequest("missing-file", "Upload a supported office, PDF, email or image file.");
    }

    const format = inferFormatFromName(file.name);
    if (!format) {
      return badRequest("unsupported-format", `Unsupported file type for ${file.name || "upload"}.`);
    }

    const prepared = await prepareImportedBytes(format, new Uint8Array(await file.arrayBuffer()));
    const now = new Date().toISOString();
    const store = createLocalSessionStore();
    const documentId = `doc_${randomUUID()}`;
    const titleField = stringFormField(form, "title");
    const sessionId = stringFormField(form, "session_id");
    const session = await sessionForNewDocument({
      store,
      sessionId,
      documentId,
      title: titleField ?? "Web imports",
      now,
    });
    const diagnostics = [
      { level: "info" as const, code: "imported", message: `Imported ${file.name} as ${format}.` },
    ];

    await store.putSession(session);
    const document = await store.putDocument(
      {
        id: documentId,
        sessionId: session.id,
        format,
        name: file.name || `untitled.${format}`,
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
            operation: "import_document",
            status: "applied",
            stage: "imported",
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
      { originalBytes: prepared.bytes, workingBytes: prepared.bytes }
    );
    const persistedSession = await store.getSession(session.id);

    return NextResponse.json(
      {
        schema: "office-ai/web-import@1",
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
        code: "session-import-error",
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

function stringFormField(form: FormData, key: string): string | undefined {
  const value = form.get(key);
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}
