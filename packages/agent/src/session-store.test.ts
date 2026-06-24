import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { LocalSessionStore, SessionStoreCorruptError, resolveOfficeAiDataDir } from "./session-store.js";

const tempDirs: string[] = [];

function tempDataDir(): string {
  const dir = mkdtempSync(join(tmpdir(), "officeai-store-"));
  tempDirs.push(dir);
  return dir;
}

afterEach(() => {
  for (const dir of tempDirs.splice(0)) {
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("LocalSessionStore", () => {
  it("writes and reads versioned sessions, documents and artifacts", async () => {
    const store = new LocalSessionStore({ dataDir: tempDataDir() });
    await store.putSession({
      id: "session_1",
      title: "Contract review",
      createdAt: "2026-06-24T09:00:00.000Z",
      updatedAt: "2026-06-24T09:00:00.000Z",
      documentIds: ["doc_1"],
    });
    await store.putDocument(
      {
        id: "doc_1",
        sessionId: "session_1",
        format: "docx",
        name: "contract.docx",
        status: "ready",
        createdAt: "2026-06-24T09:00:00.000Z",
        updatedAt: "2026-06-24T09:01:00.000Z",
        revision: 3,
        diagnostics: [{ level: "info", code: "imported", message: "Imported contract.docx as docx." }],
        exportHistory: [],
        pendingChanges: [],
        commandLog: [],
      },
      {
        originalBytes: Buffer.from("original"),
        workingBytes: Buffer.from("working"),
      }
    );

    await expect(readFile(join(store.dataDir, "VERSION.json"), "utf8")).resolves.toContain(
      "office-ai/data-dir@1"
    );
    const sessions = await store.listSessions();
    const documents = await store.listDocuments("session_1");

    expect(sessions).toHaveLength(1);
    expect(sessions[0]).toMatchObject({ schema: "office-ai/session-record@1", id: "session_1" });
    expect(documents).toHaveLength(1);
    expect(documents[0]).toMatchObject({
      schema: "office-ai/document-record@1",
      id: "doc_1",
      artifacts: {
        originalPath: expect.stringContaining("original.docx"),
        workingPath: expect.stringContaining("working.docx"),
      },
    });
    await expect(readFile(documents[0]!.artifacts.workingPath!, "utf8")).resolves.toBe("working");
  });

  it("throws a clear corrupt-store error for invalid metadata", async () => {
    const store = new LocalSessionStore({ dataDir: tempDataDir() });
    await store.putSession({
      id: "session_1",
      title: "Broken",
      createdAt: "2026-06-24T09:00:00.000Z",
      updatedAt: "2026-06-24T09:00:00.000Z",
      documentIds: [],
    });
    writeFileSync(join(store.dataDir, "sessions", "session_1", "session.json"), "{not json");

    await expect(store.listSessions()).rejects.toBeInstanceOf(SessionStoreCorruptError);
  });

  it("isolates stores by OFFICEAI_DATA_DIR", async () => {
    const prev = process.env.OFFICEAI_DATA_DIR;
    const first = tempDataDir();
    const second = tempDataDir();
    try {
      process.env.OFFICEAI_DATA_DIR = first;
      expect(resolveOfficeAiDataDir()).toBe(first);
      const a = new LocalSessionStore();
      await a.putSession({
        id: "session_a",
        title: "A",
        createdAt: "2026-06-24T09:00:00.000Z",
        updatedAt: "2026-06-24T09:00:00.000Z",
        documentIds: [],
      });

      process.env.OFFICEAI_DATA_DIR = second;
      const b = new LocalSessionStore();
      expect(await b.listSessions()).toHaveLength(0);
      expect(existsSync(join(first, "sessions", "session_a", "session.json"))).toBe(true);
    } finally {
      if (prev === undefined) {
        delete process.env.OFFICEAI_DATA_DIR;
      } else {
        process.env.OFFICEAI_DATA_DIR = prev;
      }
    }
  });
});
