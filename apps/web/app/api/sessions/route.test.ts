import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createLocalSessionStore } from "@officeai/agent/session-store";
import { GET } from "./route";

let previousOfficeAiDataDir: string | undefined;
let dataDir: string;

beforeEach(() => {
  previousOfficeAiDataDir = process.env.OFFICEAI_DATA_DIR;
  dataDir = mkdtempSync(join(tmpdir(), "officeai-web-data-"));
  process.env.OFFICEAI_DATA_DIR = dataDir;
});

afterEach(() => {
  if (previousOfficeAiDataDir === undefined) {
    delete process.env.OFFICEAI_DATA_DIR;
  } else {
    process.env.OFFICEAI_DATA_DIR = previousOfficeAiDataDir;
  }
  rmSync(dataDir, { recursive: true, force: true });
});

describe("GET /api/sessions", () => {
  it("returns path-free session and document metadata from the local data-dir", async () => {
    const store = createLocalSessionStore();
    await store.putSession({
      id: "session_1",
      title: "Web visible",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:02:00.000Z",
      documentIds: ["doc_1"],
    });
    await store.putDocument(
      {
        id: "doc_1",
        sessionId: "session_1",
        format: "pdf",
        name: "statement.pdf",
        status: "ready",
        sourcePath: "/very/local/statement.pdf",
        createdAt: "2026-06-24T10:00:00.000Z",
        updatedAt: "2026-06-24T10:02:00.000Z",
        revision: 4,
        diagnostics: [{ level: "info", code: "imported", message: "Imported statement.pdf as pdf." }],
        exportHistory: [
          { path: "/very/local/export.pdf", bytes: 42, exportedAt: "2026-06-24T10:02:00.000Z" },
        ],
        pendingChanges: [
          {
            id: "mut_1",
            operation: "pdf:rotate-pages",
            status: "pending",
            source: "agent",
          },
        ],
        commandLog: [
          {
            id: "log_1",
            operation: "pdf:rotate-pages",
            status: "pending",
            stage: "queued",
            source: "agent",
            recordedAt: "2026-06-24T10:01:00.000Z",
          },
        ],
      },
      { originalBytes: Buffer.from("%PDF-original"), workingBytes: Buffer.from("%PDF-working") }
    );

    const response = await GET();
    expect(response.status).toBe(200);
    const payload = (await response.json()) as {
      schema: string;
      sessions: Array<{ sessionId: string; documentCount: number }>;
      documents: Array<{
        documentId: string;
        exportCount: number;
        pendingChangeCount: number;
        commandLogCount: number;
        artifacts: { hasOriginal: boolean; hasWorking: boolean };
      }>;
    };

    expect(payload.schema).toBe("office-ai/web-sessions@1");
    expect(payload.sessions).toEqual([
      {
        sessionId: "session_1",
        title: "Web visible",
        createdAt: "2026-06-24T10:00:00.000Z",
        updatedAt: "2026-06-24T10:02:00.000Z",
        documentCount: 1,
      },
    ]);
    expect(payload.documents[0]).toMatchObject({
      documentId: "doc_1",
      exportCount: 1,
      pendingChangeCount: 1,
      commandLogCount: 1,
      artifacts: { hasOriginal: true, hasWorking: true },
    });
    expect(JSON.stringify(payload)).not.toContain("/very/local");
    expect(JSON.stringify(payload)).not.toContain(dataDir);
  });

  it("reports corrupt store metadata clearly", async () => {
    const store = createLocalSessionStore();
    await store.putSession({
      id: "session_1",
      title: "Corrupt",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:00:00.000Z",
      documentIds: [],
    });
    writeFileSync(join(dataDir, "sessions", "session_1", "session.json"), "{not json");

    const response = await GET();
    const payload = (await response.json()) as { code: string; message: string };
    expect(response.status).toBe(409);
    expect(payload.code).toBe("corrupt-session-store");
    expect(payload.message).toContain("Corrupt metadata file");
  });
});
