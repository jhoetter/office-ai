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
import { DocxAgent } from "@officeai/docx";
import { PdfAgent } from "@officeai/pdf";
import { PptxAgent } from "@officeai/pptx";
import { XlsxAgent } from "@officeai/xlsx";
import {
  toWebDocumentDetailEntry,
  toWebSessionEntry,
  type WebOfficeFormat,
} from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request): Promise<NextResponse> {
  try {
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File) || file.size === 0) {
      return badRequest("missing-file", "Upload a DOCX, XLSX, PPTX or PDF file.");
    }

    const format = inferFormatFromName(file.name);
    if (!format) {
      return badRequest("unsupported-format", `Unsupported file type for ${file.name || "upload"}.`);
    }

    const bytes = new Uint8Array(await file.arrayBuffer());
    const revision = await revisionFromUploadedBytes(format, bytes);
    const now = new Date().toISOString();
    const store = createLocalSessionStore();
    const documentId = `doc_${randomUUID()}`;
    const titleField = stringFormField(form, "title");
    const sessionId = stringFormField(form, "session_id");
    const session = sessionId
      ? await appendToExistingSession(store, sessionId, documentId, now)
      : {
          id: `session_${randomUUID()}`,
          title: titleField ?? "Web imports",
          createdAt: now,
          updatedAt: now,
          documentIds: [documentId],
        };
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
        revision,
        diagnostics,
        exportHistory: [],
        pendingChanges: [],
        commandLog: [
          {
            id: `log_${randomUUID()}`,
            operation: "import_document",
            status: "applied",
            stage: "imported",
            source: "web",
            recordedAt: now,
            diagnostics,
          },
        ],
      },
      { originalBytes: bytes, workingBytes: bytes }
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

function inferFormatFromName(name: string): WebOfficeFormat | null {
  const lower = name.toLowerCase();
  if (lower.endsWith(".docx")) return "docx";
  if (lower.endsWith(".xlsx")) return "xlsx";
  if (lower.endsWith(".pptx")) return "pptx";
  if (lower.endsWith(".pdf")) return "pdf";
  return null;
}

async function revisionFromUploadedBytes(format: WebOfficeFormat, bytes: Uint8Array): Promise<number> {
  switch (format) {
    case "docx":
      return (await DocxAgent.fromBuffer(bytes)).getSnapshot().revision;
    case "xlsx":
      return (await XlsxAgent.fromBuffer(bytes)).getSnapshot().revision;
    case "pptx":
      return (await PptxAgent.fromBuffer(bytes)).getSnapshot().revision;
    case "pdf":
      return (await PdfAgent.fromBuffer(bytes)).getSnapshot().revision;
  }
}

async function appendToExistingSession(
  store: ReturnType<typeof createLocalSessionStore>,
  sessionId: string,
  documentId: string,
  updatedAt: string
): Promise<{
  readonly id: string;
  readonly title: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly documentIds: ReadonlyArray<string>;
}> {
  const existing = await store.getSession(sessionId);
  return {
    id: existing.id,
    title: existing.title,
    createdAt: existing.createdAt,
    updatedAt,
    documentIds: [...new Set([...existing.documentIds, documentId])],
  };
}
