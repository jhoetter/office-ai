import { randomUUID } from "node:crypto";
import { existsSync } from "node:fs";
import { copyFile, mkdir, readFile, readdir, rename, rm, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";

export type StoredOfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";
export type StoredDiagnosticLevel = "info" | "warning" | "error" | "destructive";

export interface StoredDiagnostic {
  readonly level: StoredDiagnosticLevel;
  readonly code: string;
  readonly message: string;
}

export interface StoredExportRecord {
  readonly path: string;
  readonly bytes: number;
  readonly exportedAt: string;
}

export interface StoredPendingChange {
  readonly id: string;
  readonly operation: string;
  readonly status: string;
  readonly source: string;
  readonly actorId?: string;
  readonly timestamp?: number;
  readonly diff?: unknown;
  readonly rejection?: { readonly code: string; readonly message: string };
}

export interface StoredCommandLogEntry {
  readonly id: string;
  readonly commandId?: string;
  readonly operation: string;
  readonly status: string;
  readonly stage: string;
  readonly source: string;
  readonly actorId?: string;
  readonly recordedAt: string;
  readonly diff?: unknown;
  readonly diagnostics?: ReadonlyArray<StoredDiagnostic>;
}

export interface StoredSessionRecord {
  readonly schema: "office-ai/session-record@1";
  readonly version: 1;
  readonly id: string;
  readonly title: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly documentIds: ReadonlyArray<string>;
  readonly lease?: {
    readonly pid: number;
    readonly host: string;
    readonly updatedAt: string;
  };
}

export interface StoredDocumentRecord {
  readonly schema: "office-ai/document-record@1";
  readonly version: 1;
  readonly id: string;
  readonly sessionId: string;
  readonly format: StoredOfficeFormat;
  readonly name: string;
  readonly status: "ready" | "error";
  readonly sourcePath?: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly revision: number;
  readonly diagnostics: ReadonlyArray<StoredDiagnostic>;
  readonly exportHistory: ReadonlyArray<StoredExportRecord>;
  readonly artifacts: {
    readonly originalPath?: string;
    readonly workingPath?: string;
  };
  readonly pendingChanges: ReadonlyArray<StoredPendingChange>;
  readonly commandLog: ReadonlyArray<StoredCommandLogEntry>;
}

export interface LocalSessionStoreOptions {
  readonly dataDir?: string;
}

export class SessionStoreCorruptError extends Error {
  readonly path: string;
  constructor(path: string, message: string, cause?: unknown) {
    super(`${message}: ${path}`);
    this.name = "SessionStoreCorruptError";
    this.path = path;
    if (cause !== undefined) (this as unknown as { cause: unknown }).cause = cause;
  }
}

export class LocalSessionStore {
  readonly dataDir: string;

  constructor(opts: LocalSessionStoreOptions = {}) {
    this.dataDir = resolveOfficeAiDataDir(opts.dataDir);
  }

  async init(): Promise<void> {
    await mkdir(this.sessionsDir(), { recursive: true });
    await atomicWriteJson(join(this.dataDir, "VERSION.json"), {
      schema: "office-ai/data-dir@1",
      version: 1,
      createdOrUpdatedAt: new Date().toISOString(),
    });
  }

  sessionDir(sessionId: string): string {
    return join(this.sessionsDir(), safeSegment(sessionId));
  }

  documentDir(sessionId: string, documentId: string): string {
    return join(this.sessionDir(sessionId), "documents", safeSegment(documentId));
  }

  artifactPath(
    sessionId: string,
    documentId: string,
    kind: "original" | "working",
    format: StoredOfficeFormat
  ): string {
    return join(this.documentDir(sessionId, documentId), "artifacts", `${kind}.${format}`);
  }

  async putSession(record: Omit<StoredSessionRecord, "schema" | "version" | "lease">): Promise<void> {
    await this.init();
    const full: StoredSessionRecord = {
      schema: "office-ai/session-record@1",
      version: 1,
      ...record,
      lease: {
        pid: process.pid,
        host: "local",
        updatedAt: new Date().toISOString(),
      },
    };
    await atomicWriteJson(this.sessionJsonPath(record.id), full);
  }

  async putDocument(
    record: Omit<StoredDocumentRecord, "schema" | "version" | "artifacts">,
    opts: {
      readonly originalBytes?: Uint8Array | Buffer;
      readonly originalSourcePath?: string;
      readonly workingBytes?: Uint8Array | Buffer;
    } = {}
  ): Promise<StoredDocumentRecord> {
    await this.init();
    await mkdir(join(this.documentDir(record.sessionId, record.id), "artifacts"), { recursive: true });

    const existing = await this.getDocument(record.id).catch((err: unknown) => {
      if (err instanceof SessionStoreCorruptError) throw err;
      return null;
    });
    const originalPath =
      existing?.artifacts.originalPath ??
      (opts.originalBytes || opts.originalSourcePath
        ? this.artifactPath(record.sessionId, record.id, "original", record.format)
        : undefined);
    const workingPath =
      existing?.artifacts.workingPath ??
      (opts.workingBytes
        ? this.artifactPath(record.sessionId, record.id, "working", record.format)
        : undefined);

    if (opts.originalBytes && originalPath) {
      await writeFileAtomic(originalPath, opts.originalBytes);
    } else if (opts.originalSourcePath && originalPath && !existsSync(originalPath)) {
      await mkdir(dirname(originalPath), { recursive: true });
      await copyFile(opts.originalSourcePath, originalPath);
    }
    if (opts.workingBytes && workingPath) {
      await writeFileAtomic(workingPath, opts.workingBytes);
    }

    const full: StoredDocumentRecord = {
      schema: "office-ai/document-record@1",
      version: 1,
      ...record,
      artifacts: {
        ...(originalPath ? { originalPath } : {}),
        ...(workingPath ? { workingPath } : {}),
      },
    };
    await atomicWriteJson(this.documentJsonPath(record.sessionId, record.id), full);
    return full;
  }

  async listSessions(): Promise<ReadonlyArray<StoredSessionRecord>> {
    await this.init();
    let entries: string[];
    try {
      entries = await readdir(this.sessionsDir());
    } catch {
      return [];
    }
    const out: StoredSessionRecord[] = [];
    for (const id of entries) {
      const path = this.sessionJsonPath(id);
      if (!existsSync(path)) continue;
      out.push(await readStoredJson<StoredSessionRecord>(path, "office-ai/session-record@1"));
    }
    out.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    return out;
  }

  async listDocuments(sessionId?: string): Promise<ReadonlyArray<StoredDocumentRecord>> {
    const sessions = sessionId ? [await this.getSession(sessionId)] : await this.listSessions();
    const out: StoredDocumentRecord[] = [];
    for (const session of sessions) {
      for (const documentId of session.documentIds) {
        const doc = await this.getDocument(documentId);
        if (doc) out.push(doc);
      }
    }
    out.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    return out;
  }

  async getSession(sessionId: string): Promise<StoredSessionRecord> {
    return readStoredJson<StoredSessionRecord>(this.sessionJsonPath(sessionId), "office-ai/session-record@1");
  }

  async getDocument(documentId: string): Promise<StoredDocumentRecord | null> {
    const sessions = await this.listSessions();
    for (const session of sessions) {
      if (!session.documentIds.includes(documentId)) continue;
      try {
        return await readStoredJson<StoredDocumentRecord>(
          this.documentJsonPath(session.id, documentId),
          "office-ai/document-record@1"
        );
      } catch (err) {
        if (isEnoent(err)) return null;
        throw err;
      }
    }
    return null;
  }

  async readWorkingBytes(record: StoredDocumentRecord): Promise<Uint8Array> {
    const path = record.artifacts.workingPath ?? record.artifacts.originalPath;
    if (!path) throw new Error(`Document ${record.id} has no persisted artifact.`);
    return new Uint8Array(await readFile(path));
  }

  async clear(): Promise<void> {
    await rm(this.dataDir, { recursive: true, force: true });
  }

  private sessionsDir(): string {
    return join(this.dataDir, "sessions");
  }

  private sessionJsonPath(sessionId: string): string {
    return join(this.sessionDir(sessionId), "session.json");
  }

  private documentJsonPath(sessionId: string, documentId: string): string {
    return join(this.documentDir(sessionId, documentId), "document.json");
  }
}

export function createLocalSessionStore(opts: LocalSessionStoreOptions = {}): LocalSessionStore {
  return new LocalSessionStore(opts);
}

export function resolveOfficeAiDataDir(explicit?: string): string {
  if (explicit && explicit.trim().length > 0) return resolve(explicit);
  const env = process.env.OFFICEAI_DATA_DIR;
  if (env && env.trim().length > 0) return resolve(env);
  const xdg = process.env.XDG_DATA_HOME;
  if (xdg && xdg.trim().length > 0) return resolve(xdg, "office-ai");
  return resolve(homedir(), ".local", "share", "office-ai");
}

async function readStoredJson<T extends { readonly schema?: string }>(
  path: string,
  schema: string
): Promise<T> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(await readFile(path, "utf8"));
  } catch (err) {
    if (isEnoent(err)) throw err;
    throw new SessionStoreCorruptError(path, "Corrupt metadata file", err);
  }
  if (!parsed || typeof parsed !== "object" || (parsed as { schema?: unknown }).schema !== schema) {
    throw new SessionStoreCorruptError(path, `Metadata schema mismatch, expected ${schema}`);
  }
  return parsed as T;
}

function isEnoent(err: unknown): boolean {
  return (err as NodeJS.ErrnoException)?.code === "ENOENT";
}

async function atomicWriteJson(path: string, value: unknown): Promise<void> {
  await writeFileAtomic(path, Buffer.from(`${JSON.stringify(value, null, 2)}\n`, "utf8"));
}

async function writeFileAtomic(path: string, bytes: Uint8Array | Buffer): Promise<void> {
  await mkdir(dirname(path), { recursive: true });
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
  await writeFile(tmp, bytes);
  await rename(tmp, path);
}

function safeSegment(value: string): string {
  const cleaned = value.replace(/[^a-zA-Z0-9_.-]/g, "_");
  if (cleaned.length === 0 || cleaned === "." || cleaned === "..") {
    throw new Error(`Invalid data-dir segment "${value}".`);
  }
  return cleaned;
}
