import { randomUUID } from "node:crypto";
import { copyFile, mkdir, readFile, readdir, rename, rm, stat, writeFile } from "node:fs/promises";
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

export type SessionStorageOperation = "ensure-dir" | "list" | "exists" | "read" | "write" | "copy" | "remove";

export interface SessionStorageCapabilities {
  readonly atomicWrite: boolean;
  readonly localPaths: boolean;
  readonly locks: "advisory" | "none";
  readonly watch: boolean;
}

export interface SessionStorageRemoveOptions {
  readonly recursive?: boolean;
  readonly force?: boolean;
}

export interface SessionStorageAdapter {
  readonly kind: string;
  readonly root: string;
  readonly capabilities: SessionStorageCapabilities;
  join(...segments: ReadonlyArray<string>): string;
  ensureDir(path: string): Promise<void>;
  list(path: string): Promise<ReadonlyArray<string>>;
  exists(path: string): Promise<boolean>;
  readBytes(path: string): Promise<Uint8Array>;
  writeBytesAtomic(path: string, bytes: Uint8Array | Buffer): Promise<void>;
  copyFromLocalFile(sourcePath: string, targetPath: string): Promise<void>;
  remove(path: string, opts?: SessionStorageRemoveOptions): Promise<void>;
}

export interface LocalSessionStoreOptions {
  readonly dataDir?: string;
  readonly storage?: SessionStorageAdapter;
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

export class SessionStoreStorageError extends Error {
  readonly operation: SessionStorageOperation;
  readonly path: string;
  readonly causeCode?: string;
  readonly diagnostic: StoredDiagnostic;

  constructor(operation: SessionStorageOperation, path: string, cause: unknown) {
    const message = cause instanceof Error ? cause.message : String(cause);
    super(`Storage ${operation} failed for ${path}: ${message}`);
    this.name = "SessionStoreStorageError";
    this.operation = operation;
    this.path = path;
    this.causeCode = (cause as NodeJS.ErrnoException | undefined)?.code;
    this.diagnostic = {
      level: "error",
      code: `storage-${operation}`,
      message: `Storage ${operation} failed for ${path}: ${message}`,
    };
    if (cause !== undefined) (this as unknown as { cause: unknown }).cause = cause;
  }
}

export class LocalFilesystemSessionStorageAdapter implements SessionStorageAdapter {
  readonly kind = "local-filesystem";
  readonly root: string;
  readonly capabilities: SessionStorageCapabilities = {
    atomicWrite: true,
    localPaths: true,
    locks: "advisory",
    watch: false,
  };

  constructor(root: string) {
    this.root = resolve(root);
  }

  join(...segments: ReadonlyArray<string>): string {
    return join(...segments);
  }

  async ensureDir(path: string): Promise<void> {
    await this.wrap("ensure-dir", path, () => mkdir(path, { recursive: true }));
  }

  async list(path: string): Promise<ReadonlyArray<string>> {
    try {
      return await readdir(path);
    } catch (err) {
      if (isEnoent(err)) return [];
      throw new SessionStoreStorageError("list", path, err);
    }
  }

  async exists(path: string): Promise<boolean> {
    try {
      await stat(path);
      return true;
    } catch (err) {
      if (isEnoent(err)) return false;
      throw new SessionStoreStorageError("exists", path, err);
    }
  }

  async readBytes(path: string): Promise<Uint8Array> {
    return this.wrap("read", path, async () => new Uint8Array(await readFile(path)));
  }

  async writeBytesAtomic(path: string, bytes: Uint8Array | Buffer): Promise<void> {
    await this.wrap("write", path, async () => {
      await mkdir(dirname(path), { recursive: true });
      const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
      await writeFile(tmp, bytes);
      await rename(tmp, path);
    });
  }

  async copyFromLocalFile(sourcePath: string, targetPath: string): Promise<void> {
    await this.wrap("copy", targetPath, async () => {
      await mkdir(dirname(targetPath), { recursive: true });
      await copyFile(sourcePath, targetPath);
    });
  }

  async remove(path: string, opts: SessionStorageRemoveOptions = {}): Promise<void> {
    await this.wrap("remove", path, () =>
      rm(path, { recursive: opts.recursive ?? false, force: opts.force ?? false })
    );
  }

  private async wrap<T>(operation: SessionStorageOperation, path: string, fn: () => Promise<T>): Promise<T> {
    try {
      return await fn();
    } catch (err) {
      throw new SessionStoreStorageError(operation, path, err);
    }
  }
}

export class LocalSessionStore {
  readonly dataDir: string;
  readonly storage: SessionStorageAdapter;

  constructor(opts: LocalSessionStoreOptions = {}) {
    this.storage = opts.storage ?? createLocalFilesystemSessionStorageAdapter(opts.dataDir);
    this.dataDir = this.storage.root;
  }

  async init(): Promise<void> {
    await this.storage.ensureDir(this.sessionsDir());
    await atomicWriteJson(this.storage, this.storage.join(this.dataDir, "VERSION.json"), {
      schema: "office-ai/data-dir@1",
      version: 1,
      createdOrUpdatedAt: new Date().toISOString(),
    });
  }

  sessionDir(sessionId: string): string {
    return this.storage.join(this.sessionsDir(), safeSegment(sessionId));
  }

  documentDir(sessionId: string, documentId: string): string {
    return this.storage.join(this.sessionDir(sessionId), "documents", safeSegment(documentId));
  }

  artifactPath(
    sessionId: string,
    documentId: string,
    kind: "original" | "working",
    format: StoredOfficeFormat
  ): string {
    return this.storage.join(this.documentDir(sessionId, documentId), "artifacts", `${kind}.${format}`);
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
    await atomicWriteJson(this.storage, this.sessionJsonPath(record.id), full);
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
    await this.storage.ensureDir(
      this.storage.join(this.documentDir(record.sessionId, record.id), "artifacts")
    );

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
      await this.storage.writeBytesAtomic(originalPath, opts.originalBytes);
    } else if (opts.originalSourcePath && originalPath && !(await this.storage.exists(originalPath))) {
      await this.storage.copyFromLocalFile(opts.originalSourcePath, originalPath);
    }
    if (opts.workingBytes && workingPath) {
      await this.storage.writeBytesAtomic(workingPath, opts.workingBytes);
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
    await atomicWriteJson(this.storage, this.documentJsonPath(record.sessionId, record.id), full);
    return full;
  }

  async listSessions(): Promise<ReadonlyArray<StoredSessionRecord>> {
    await this.init();
    const entries = await this.storage.list(this.sessionsDir());
    const out: StoredSessionRecord[] = [];
    for (const id of entries) {
      const path = this.sessionJsonPath(id);
      if (!(await this.storage.exists(path))) continue;
      out.push(await readStoredJson<StoredSessionRecord>(this.storage, path, "office-ai/session-record@1"));
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
    return readStoredJson<StoredSessionRecord>(
      this.storage,
      this.sessionJsonPath(sessionId),
      "office-ai/session-record@1"
    );
  }

  async getDocument(documentId: string): Promise<StoredDocumentRecord | null> {
    const sessions = await this.listSessions();
    for (const session of sessions) {
      if (!session.documentIds.includes(documentId)) continue;
      try {
        return await readStoredJson<StoredDocumentRecord>(
          this.storage,
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
    return this.storage.readBytes(path);
  }

  async clear(): Promise<void> {
    await this.storage.remove(this.dataDir, { recursive: true, force: true });
  }

  private sessionsDir(): string {
    return this.storage.join(this.dataDir, "sessions");
  }

  private sessionJsonPath(sessionId: string): string {
    return this.storage.join(this.sessionDir(sessionId), "session.json");
  }

  private documentJsonPath(sessionId: string, documentId: string): string {
    return this.storage.join(this.documentDir(sessionId, documentId), "document.json");
  }
}

export function createLocalSessionStore(opts: LocalSessionStoreOptions = {}): LocalSessionStore {
  return new LocalSessionStore(opts);
}

export function createLocalFilesystemSessionStorageAdapter(
  dataDir?: string
): LocalFilesystemSessionStorageAdapter {
  return new LocalFilesystemSessionStorageAdapter(resolveOfficeAiDataDir(dataDir));
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
  storage: SessionStorageAdapter,
  path: string,
  schema: string
): Promise<T> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(await storage.readBytes(path)).toString("utf8"));
  } catch (err) {
    if (isEnoent(err)) throw err;
    if (err instanceof SessionStoreStorageError) throw err;
    throw new SessionStoreCorruptError(path, "Corrupt metadata file", err);
  }
  if (!parsed || typeof parsed !== "object" || (parsed as { schema?: unknown }).schema !== schema) {
    throw new SessionStoreCorruptError(path, `Metadata schema mismatch, expected ${schema}`);
  }
  return parsed as T;
}

function isEnoent(err: unknown): boolean {
  return (
    (err as NodeJS.ErrnoException)?.code === "ENOENT" ||
    (err instanceof SessionStoreStorageError && err.causeCode === "ENOENT")
  );
}

async function atomicWriteJson(storage: SessionStorageAdapter, path: string, value: unknown): Promise<void> {
  await storage.writeBytesAtomic(path, Buffer.from(`${JSON.stringify(value, null, 2)}\n`, "utf8"));
}

function safeSegment(value: string): string {
  const cleaned = value.replace(/[^a-zA-Z0-9_.-]/g, "_");
  if (cleaned.length === 0 || cleaned === "." || cleaned === "..") {
    throw new Error(`Invalid data-dir segment "${value}".`);
  }
  return cleaned;
}
