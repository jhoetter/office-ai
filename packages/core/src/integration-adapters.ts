import type { DocumentFormat } from "./types/document.js";

export type OfficeAiDiagnosticLevel = "info" | "warning" | "error" | "destructive";

export interface OfficeAiAdapterDiagnostic {
  readonly level: OfficeAiDiagnosticLevel;
  readonly code: string;
  readonly message: string;
}

export interface OfficeAiStorageCapabilities {
  readonly atomicWrite: boolean;
  readonly localPaths: boolean;
  readonly locks: "advisory" | "none";
  readonly watch: boolean;
}

export interface OfficeAiStorageRemoveOptions {
  readonly recursive?: boolean;
  readonly force?: boolean;
}

export interface OfficeAiStorageAdapter {
  readonly kind: string;
  readonly root: string;
  readonly capabilities: OfficeAiStorageCapabilities;
  join(...segments: ReadonlyArray<string>): string;
  ensureDir(path: string): Promise<void>;
  list(path: string): Promise<ReadonlyArray<string>>;
  exists(path: string): Promise<boolean>;
  readBytes(path: string): Promise<Uint8Array>;
  writeBytesAtomic(path: string, bytes: Uint8Array): Promise<void>;
  copyFromLocalFile(sourcePath: string, targetPath: string): Promise<void>;
  remove(path: string, opts?: OfficeAiStorageRemoveOptions): Promise<void>;
}

export interface OfficeAiActor {
  readonly id: string;
  readonly displayName?: string;
  readonly email?: string;
  readonly roles?: ReadonlyArray<string>;
}

export interface OfficeAiIdentityAdapter {
  readonly kind: string;
  getCurrentActor(): Promise<OfficeAiActor | null>;
}

export type OfficeAiAssetFormat = DocumentFormat | "json" | "markdown" | "html" | "binary";

export interface OfficeAiAssetSource {
  readonly kind: "bytes" | "local-path" | "url" | "generated";
  readonly label?: string;
  readonly localPath?: string;
  readonly url?: string;
}

export interface OfficeAiAssetRef {
  readonly schema: "office-ai/asset-ref@1";
  readonly id: string;
  readonly format: OfficeAiAssetFormat;
  readonly mediaType: string;
  readonly bytes?: number;
  readonly sha256?: string;
  readonly source?: OfficeAiAssetSource;
  readonly diagnostics: ReadonlyArray<OfficeAiAdapterDiagnostic>;
}

export interface OfficeAiAssetImportRequest {
  readonly id?: string;
  readonly format: OfficeAiAssetFormat;
  readonly mediaType?: string;
  readonly bytes?: Uint8Array;
  readonly source?: OfficeAiAssetSource;
}

export interface OfficeAiAssetExportRequest {
  readonly assetId: string;
  readonly requestedFormat?: OfficeAiAssetFormat;
}

export interface OfficeAiAssetAdapter {
  readonly kind: string;
  importAsset(request: OfficeAiAssetImportRequest): Promise<OfficeAiAssetRef>;
  exportAsset(request: OfficeAiAssetExportRequest): Promise<OfficeAiAssetRef>;
  readAssetBytes(assetId: string): Promise<Uint8Array>;
}

export interface OfficeAiEvent {
  readonly schema: "office-ai/event@1";
  readonly type: string;
  readonly occurredAt: string;
  readonly actor?: OfficeAiActor;
  readonly payload?: Readonly<Record<string, unknown>>;
}

export interface OfficeAiEventHookAdapter {
  readonly kind: string;
  emit(event: OfficeAiEvent): Promise<void>;
}

export interface OfficeAiUiEmbeddingRequest {
  readonly format?: DocumentFormat;
  readonly documentId?: string;
  readonly sessionId?: string;
  readonly locale?: string;
}

export interface OfficeAiUiEmbeddingDescriptor {
  readonly mode: "web-url" | "react-package" | "none";
  readonly url?: string;
  readonly packageName?: string;
  readonly entrypoint?: string;
  readonly diagnostics: ReadonlyArray<OfficeAiAdapterDiagnostic>;
}

export interface OfficeAiUiEmbeddingAdapter {
  readonly kind: string;
  getEmbedding(request?: OfficeAiUiEmbeddingRequest): Promise<OfficeAiUiEmbeddingDescriptor>;
}

export interface OfficeAiMcpHostDescriptor {
  readonly transport: "stdio" | "http" | "hosted";
  readonly command?: string;
  readonly args?: ReadonlyArray<string>;
  readonly url?: string;
  readonly diagnostics: ReadonlyArray<OfficeAiAdapterDiagnostic>;
}

export interface OfficeAiMcpHostAdapter {
  readonly kind: string;
  describeHost(): Promise<OfficeAiMcpHostDescriptor>;
}

export interface OfficeAiIntegrationAdapters {
  readonly storage?: OfficeAiStorageAdapter;
  readonly identity?: OfficeAiIdentityAdapter;
  readonly assets?: OfficeAiAssetAdapter;
  readonly events?: OfficeAiEventHookAdapter;
  readonly ui?: OfficeAiUiEmbeddingAdapter;
  readonly mcp?: OfficeAiMcpHostAdapter;
}

export interface LocalOfficeAiIntegrationAdapterOptions {
  readonly root?: string;
  readonly actor?: OfficeAiActor | null;
  readonly webBaseUrl?: string;
  readonly mcpCommand?: string;
  readonly mcpArgs?: ReadonlyArray<string>;
}

export interface LocalOfficeAiIntegrationAdapters extends OfficeAiIntegrationAdapters {
  readonly storage: MemoryOfficeAiStorageAdapter;
  readonly identity: StaticOfficeAiIdentityAdapter;
  readonly assets: MemoryOfficeAiAssetAdapter;
  readonly events: MemoryOfficeAiEventHookAdapter;
  readonly ui: LocalOfficeAiUiEmbeddingAdapter;
  readonly mcp: LocalOfficeAiMcpHostAdapter;
}

export class MemoryOfficeAiStorageAdapter implements OfficeAiStorageAdapter {
  readonly kind = "memory";
  readonly root: string;
  readonly capabilities: OfficeAiStorageCapabilities = {
    atomicWrite: true,
    localPaths: false,
    locks: "none",
    watch: false,
  };

  private readonly dirs = new Set<string>(["/"]);
  private readonly files = new Map<string, Uint8Array>();

  constructor(root = "memory://office-ai") {
    this.root = root;
  }

  join(...segments: ReadonlyArray<string>): string {
    return normalizeStoragePath(segments.join("/"));
  }

  async ensureDir(path: string): Promise<void> {
    this.ensureDirSync(normalizeStoragePath(path));
  }

  async list(path: string): Promise<ReadonlyArray<string>> {
    const normalized = normalizeStoragePath(path);
    if (!this.dirs.has(normalized)) return [];
    const prefix = normalized === "/" ? "/" : `${normalized}/`;
    const children = new Set<string>();
    for (const dir of this.dirs) {
      collectChild(prefix, dir, children);
    }
    for (const file of this.files.keys()) {
      collectChild(prefix, file, children);
    }
    return [...children].sort();
  }

  async exists(path: string): Promise<boolean> {
    const normalized = normalizeStoragePath(path);
    return this.dirs.has(normalized) || this.files.has(normalized);
  }

  async readBytes(path: string): Promise<Uint8Array> {
    const normalized = normalizeStoragePath(path);
    const found = this.files.get(normalized);
    if (!found) throw createNotFoundError(normalized);
    return copyBytes(found);
  }

  async writeBytesAtomic(path: string, bytes: Uint8Array): Promise<void> {
    const normalized = normalizeStoragePath(path);
    this.ensureDirSync(parentStoragePath(normalized));
    this.files.set(normalized, copyBytes(bytes));
  }

  async copyFromLocalFile(): Promise<void> {
    throw createUnsupportedError("copy-from-local-file");
  }

  async remove(path: string, opts: OfficeAiStorageRemoveOptions = {}): Promise<void> {
    const normalized = normalizeStoragePath(path);
    if (this.files.delete(normalized)) return;
    if (!this.dirs.has(normalized)) {
      if (opts.force) return;
      throw createNotFoundError(normalized);
    }

    const prefix = normalized === "/" ? "/" : `${normalized}/`;
    const childFiles = [...this.files.keys()].filter((file) => file.startsWith(prefix));
    const childDirs = [...this.dirs].filter((dir) => dir !== normalized && dir.startsWith(prefix));
    if (!opts.recursive && (childFiles.length > 0 || childDirs.length > 0)) {
      throw new Error(`Storage directory is not empty: ${normalized}`);
    }
    for (const file of childFiles) this.files.delete(file);
    for (const dir of childDirs) this.dirs.delete(dir);
    if (normalized !== "/") this.dirs.delete(normalized);
  }

  private ensureDirSync(path: string): void {
    const normalized = normalizeStoragePath(path);
    if (this.dirs.has(normalized)) return;
    const segments = normalized.split("/").filter(Boolean);
    let cursor = "";
    for (const segment of segments) {
      cursor = `${cursor}/${segment}`;
      this.dirs.add(cursor);
    }
  }
}

export class StaticOfficeAiIdentityAdapter implements OfficeAiIdentityAdapter {
  readonly kind = "static";
  private readonly actor: OfficeAiActor | null;

  constructor(actor: OfficeAiActor | null = null) {
    this.actor = actor;
  }

  async getCurrentActor(): Promise<OfficeAiActor | null> {
    return this.actor;
  }
}

export class MemoryOfficeAiEventHookAdapter implements OfficeAiEventHookAdapter {
  readonly kind = "memory";
  readonly events: OfficeAiEvent[] = [];

  async emit(event: OfficeAiEvent): Promise<void> {
    this.events.push({
      ...event,
      payload: event.payload ? { ...event.payload } : undefined,
    });
  }
}

export class MemoryOfficeAiAssetAdapter implements OfficeAiAssetAdapter {
  readonly kind = "memory";
  private readonly refs = new Map<string, OfficeAiAssetRef>();
  private readonly bytes = new Map<string, Uint8Array>();
  private nextId = 1;

  async importAsset(request: OfficeAiAssetImportRequest): Promise<OfficeAiAssetRef> {
    if (!request.bytes) throw createUnsupportedError("asset-import-without-bytes");
    const id = request.id ?? `asset-${this.nextId++}`;
    const ref: OfficeAiAssetRef = {
      schema: "office-ai/asset-ref@1",
      id,
      format: request.format,
      mediaType: request.mediaType ?? mediaTypeForAssetFormat(request.format),
      bytes: request.bytes.byteLength,
      source: request.source,
      diagnostics: [],
    };
    this.refs.set(id, ref);
    this.bytes.set(id, copyBytes(request.bytes));
    return ref;
  }

  async exportAsset(request: OfficeAiAssetExportRequest): Promise<OfficeAiAssetRef> {
    const ref = this.refs.get(request.assetId);
    if (!ref) throw createNotFoundError(request.assetId);
    if (!request.requestedFormat || request.requestedFormat === ref.format) return ref;
    return {
      ...ref,
      format: request.requestedFormat,
      diagnostics: [
        ...ref.diagnostics,
        {
          level: "warning",
          code: "asset-format-not-transcoded",
          message: "Reference adapter returns the stored bytes without transcoding.",
        },
      ],
    };
  }

  async readAssetBytes(assetId: string): Promise<Uint8Array> {
    const found = this.bytes.get(assetId);
    if (!found) throw createNotFoundError(assetId);
    return copyBytes(found);
  }
}

export class LocalOfficeAiUiEmbeddingAdapter implements OfficeAiUiEmbeddingAdapter {
  readonly kind = "local-web";
  private readonly webBaseUrl?: string;

  constructor(webBaseUrl?: string) {
    this.webBaseUrl = webBaseUrl;
  }

  async getEmbedding(request: OfficeAiUiEmbeddingRequest = {}): Promise<OfficeAiUiEmbeddingDescriptor> {
    if (!this.webBaseUrl) {
      return {
        mode: "react-package",
        packageName: "@officeai/react-editors",
        entrypoint: request.format ? `./components/${request.format}` : "./components",
        diagnostics: [],
      };
    }

    const url = new URL(this.webBaseUrl);
    if (request.documentId) url.pathname = `/sessions/${encodeURIComponent(request.documentId)}`;
    if (request.sessionId) url.searchParams.set("session", request.sessionId);
    if (request.locale) url.searchParams.set("locale", request.locale);
    return { mode: "web-url", url: url.toString(), diagnostics: [] };
  }
}

export class LocalOfficeAiMcpHostAdapter implements OfficeAiMcpHostAdapter {
  readonly kind = "local-stdio";
  private readonly command: string;
  private readonly args: ReadonlyArray<string>;

  constructor(command = "office-agent", args: ReadonlyArray<string> = ["mcp"]) {
    this.command = command;
    this.args = [...args];
  }

  async describeHost(): Promise<OfficeAiMcpHostDescriptor> {
    return {
      transport: "stdio",
      command: this.command,
      args: this.args,
      diagnostics: [],
    };
  }
}

export function createLocalIntegrationAdapters(
  opts: LocalOfficeAiIntegrationAdapterOptions = {}
): LocalOfficeAiIntegrationAdapters {
  return {
    storage: new MemoryOfficeAiStorageAdapter(opts.root),
    identity: new StaticOfficeAiIdentityAdapter(opts.actor ?? null),
    assets: new MemoryOfficeAiAssetAdapter(),
    events: new MemoryOfficeAiEventHookAdapter(),
    ui: new LocalOfficeAiUiEmbeddingAdapter(opts.webBaseUrl),
    mcp: new LocalOfficeAiMcpHostAdapter(opts.mcpCommand, opts.mcpArgs),
  };
}

export function createOfficeAiEvent(
  type: string,
  opts: {
    readonly payload?: Readonly<Record<string, unknown>>;
    readonly actor?: OfficeAiActor;
    readonly occurredAt?: string;
    readonly now?: () => string;
  } = {}
): OfficeAiEvent {
  return {
    schema: "office-ai/event@1",
    type,
    occurredAt: opts.occurredAt ?? opts.now?.() ?? new Date().toISOString(),
    actor: opts.actor,
    payload: opts.payload,
  };
}

function copyBytes(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes);
}

function normalizeStoragePath(path: string): string {
  const cleaned = path.replace(/\\/g, "/").split("/").filter(Boolean).join("/");
  return cleaned.length === 0 ? "/" : `/${cleaned}`;
}

function parentStoragePath(path: string): string {
  const normalized = normalizeStoragePath(path);
  if (normalized === "/") return "/";
  const idx = normalized.lastIndexOf("/");
  return idx <= 0 ? "/" : normalized.slice(0, idx);
}

function collectChild(prefix: string, candidate: string, out: Set<string>): void {
  if (candidate === "/" || !candidate.startsWith(prefix)) return;
  const rest = candidate.slice(prefix.length);
  if (!rest || rest.includes("/")) return;
  out.add(rest);
}

function createNotFoundError(id: string): Error & { code: "ENOENT" } {
  const err = new Error(`No such office-ai integration resource: ${id}`) as Error & { code: "ENOENT" };
  err.code = "ENOENT";
  return err;
}

function createUnsupportedError(code: string): Error & { code: string } {
  const err = new Error(`Unsupported office-ai integration adapter operation: ${code}`) as Error & {
    code: string;
  };
  err.code = code;
  return err;
}

function mediaTypeForAssetFormat(format: OfficeAiAssetFormat): string {
  switch (format) {
    case "docx":
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    case "xlsx":
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    case "pptx":
      return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    case "pdf":
      return "application/pdf";
    case "json":
      return "application/json";
    case "markdown":
      return "text/markdown";
    case "html":
      return "text/html";
    case "binary":
      return "application/octet-stream";
  }
}
