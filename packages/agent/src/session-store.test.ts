import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import {
  LocalSessionStore,
  SessionStoreCorruptError,
  SessionStoreStorageError,
  resolveOfficeAiDataDir,
  type SessionStorageAdapter,
  type SessionStorageRemoveOptions,
} from "./session-store.js";

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

  it("can run against a test storage adapter", async () => {
    const storage = new MemorySessionStorageAdapter();
    const store = new LocalSessionStore({ storage });
    await store.putSession({
      id: "session_mem",
      title: "In memory",
      createdAt: "2026-06-24T09:00:00.000Z",
      updatedAt: "2026-06-24T09:00:00.000Z",
      documentIds: ["doc_mem"],
    });
    const stored = await store.putDocument(
      {
        id: "doc_mem",
        sessionId: "session_mem",
        format: "pdf",
        name: "memory.pdf",
        status: "ready",
        createdAt: "2026-06-24T09:00:00.000Z",
        updatedAt: "2026-06-24T09:01:00.000Z",
        revision: 1,
        diagnostics: [],
        exportHistory: [],
        pendingChanges: [],
        commandLog: [],
      },
      {
        workingBytes: Buffer.from("memory bytes"),
      }
    );

    expect(store.dataDir).toBe("memory://office-ai");
    expect(stored.artifacts.workingPath).toBe(
      "memory://office-ai/sessions/session_mem/documents/doc_mem/artifacts/working.pdf"
    );
    await expect(store.listSessions()).resolves.toHaveLength(1);
    expect(Buffer.from(await store.readWorkingBytes(stored)).toString("utf8")).toBe("memory bytes");
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

  it("migrates legacy session metadata after writing a backup", async () => {
    const store = new LocalSessionStore({ dataDir: tempDataDir() });
    const sessionDir = join(store.dataDir, "sessions", "session_legacy");
    mkdirSync(sessionDir, { recursive: true });
    writeFileSync(
      join(sessionDir, "session.json"),
      JSON.stringify(
        {
          id: "session_legacy",
          title: "Legacy",
          createdAt: "2026-06-24T09:00:00.000Z",
          updatedAt: "2026-06-24T09:00:00.000Z",
          documentIds: [],
        },
        null,
        2
      )
    );

    const before = await store.inspectDataDir();
    expect(before.needsMigration).toBe(true);
    expect(before.diagnostics).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: "session-migration-required" })])
    );

    const result = await store.migrateDataDir();
    expect(result.migrations).toHaveLength(1);
    expect(existsSync(result.migrations[0]!.backupPath)).toBe(true);
    const sessions = await store.listSessions();
    expect(sessions[0]).toMatchObject({
      schema: "office-ai/session-record@1",
      schemaVersion: 1,
      version: 1,
      id: "session_legacy",
    });
    await expect(store.inspectDataDir()).resolves.toMatchObject({ needsMigration: false });
  });

  it("cleans temporary store files without deleting original artifacts", async () => {
    const store = new LocalSessionStore({ dataDir: tempDataDir() });
    await store.putSession({
      id: "session_1",
      title: "Cleanup",
      createdAt: "2026-06-24T09:00:00.000Z",
      updatedAt: "2026-06-24T09:00:00.000Z",
      documentIds: ["doc_1"],
    });
    const stored = await store.putDocument(
      {
        id: "doc_1",
        sessionId: "session_1",
        format: "pdf",
        name: "cleanup.pdf",
        status: "ready",
        createdAt: "2026-06-24T09:00:00.000Z",
        updatedAt: "2026-06-24T09:01:00.000Z",
        revision: 1,
        diagnostics: [],
        exportHistory: [],
        pendingChanges: [],
        commandLog: [],
      },
      {
        originalBytes: Buffer.from("original"),
        workingBytes: Buffer.from("working"),
      }
    );
    const tempPath = join(store.documentDir("session_1", "doc_1"), "artifacts", ".working.pdf.1.tmp");
    writeFileSync(tempPath, "temp");

    const result = await store.cleanupTemporaryArtifacts();
    expect(result.removed).toContain(tempPath);
    expect(existsSync(tempPath)).toBe(false);
    expect(existsSync(stored.artifacts.originalPath!)).toBe(true);
    expect(existsSync(stored.artifacts.workingPath!)).toBe(true);
  });

  it("surfaces storage adapter failures as structured diagnostics", async () => {
    const fileRoot = join(tempDataDir(), "not-a-directory");
    writeFileSync(fileRoot, "not a dir");
    const store = new LocalSessionStore({ dataDir: fileRoot });

    await expect(store.listSessions()).rejects.toMatchObject({
      name: "SessionStoreStorageError",
      diagnostic: {
        level: "error",
        code: "storage-ensure-dir",
      },
    });
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

class MemorySessionStorageAdapter implements SessionStorageAdapter {
  readonly kind = "memory";
  readonly root = "memory://office-ai";
  readonly capabilities = {
    atomicWrite: true,
    localPaths: false,
    locks: "none" as const,
    watch: false,
  };

  private readonly files = new Map<string, Uint8Array>();
  private readonly dirs = new Set<string>([this.root]);

  join(...segments: ReadonlyArray<string>): string {
    const [first = this.root, ...rest] = segments;
    return [first.replace(/\/+$/g, ""), ...rest.map((segment) => segment.replace(/^\/+|\/+$/g, ""))]
      .filter(Boolean)
      .join("/");
  }

  async ensureDir(path: string): Promise<void> {
    this.ensureParents(path);
    this.dirs.add(path);
  }

  async list(path: string): Promise<ReadonlyArray<string>> {
    const prefix = `${path.replace(/\/+$/g, "")}/`;
    const entries = new Set<string>();
    for (const candidate of [...this.dirs, ...this.files.keys()]) {
      if (!candidate.startsWith(prefix)) continue;
      const entry = candidate.slice(prefix.length).split("/")[0];
      if (entry) entries.add(entry);
    }
    return [...entries];
  }

  async exists(path: string): Promise<boolean> {
    return this.dirs.has(path) || this.files.has(path);
  }

  async readBytes(path: string): Promise<Uint8Array> {
    const bytes = this.files.get(path);
    if (!bytes) throw Object.assign(new Error(`No such file: ${path}`), { code: "ENOENT" });
    return bytes;
  }

  async writeBytesAtomic(path: string, bytes: Uint8Array | Buffer): Promise<void> {
    this.ensureParents(path);
    this.files.set(path, new Uint8Array(bytes));
  }

  async copyFromLocalFile(): Promise<void> {
    throw new SessionStoreStorageError("copy", this.root, new Error("copyFromLocalFile unsupported"));
  }

  async remove(path: string, opts: SessionStorageRemoveOptions = {}): Promise<void> {
    this.files.delete(path);
    this.dirs.delete(path);
    if (opts.recursive) {
      const prefix = `${path.replace(/\/+$/g, "")}/`;
      for (const key of [...this.files.keys()]) {
        if (key.startsWith(prefix)) this.files.delete(key);
      }
      for (const key of [...this.dirs]) {
        if (key.startsWith(prefix)) this.dirs.delete(key);
      }
    }
  }

  private ensureParents(path: string): void {
    const parts = path.split("/");
    for (let i = 3; i < parts.length; i += 1) {
      this.dirs.add(parts.slice(0, i).join("/"));
    }
  }
}
