import JSZip from "jszip";
import { sha256Hex } from "../util/hash.js";

export interface OoxmlPart {
  readonly path: string;
  readonly bytes: Uint8Array;
  readonly dirty: boolean;
}

interface MutablePart {
  path: string;
  bytes: Uint8Array;
  dirty: boolean;
}

export class OoxmlContainerError extends Error {
  readonly code: string;
  readonly partPath?: string;
  constructor(code: string, message: string, opts?: { partPath?: string; cause?: unknown }) {
    super(message);
    this.name = "OoxmlContainerError";
    this.code = code;
    this.partPath = opts?.partPath;
    if (opts?.cause !== undefined) (this as unknown as { cause: unknown }).cause = opts.cause;
  }
}

const TEXT_DECODER = new TextDecoder("utf-8", { fatal: false });
const TEXT_ENCODER = new TextEncoder();

/**
 * The OOXML zip container. Loads every part eagerly into memory so we can
 * guarantee byte-for-byte preservation of untouched parts on serialize.
 *
 * See spec/shared/ooxml-utils.md.
 */
export class OoxmlContainer {
  private readonly _parts = new Map<string, MutablePart>();
  private readonly _order: string[] = [];

  static async load(buffer: ArrayBuffer | Uint8Array): Promise<OoxmlContainer> {
    const zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    const container = new OoxmlContainer();
    const entries = Object.keys(zip.files).sort();
    for (const path of entries) {
      const file = zip.files[path];
      if (file.dir) continue;
      const bytes = await file.async("uint8array");
      container._parts.set(path, { path, bytes, dirty: false });
      container._order.push(path);
    }
    return container;
  }

  /** Map of all parts, in load order. */
  get parts(): ReadonlyMap<string, OoxmlPart> {
    return this._parts;
  }

  has(path: string): boolean {
    return this._parts.has(path);
  }

  /** Read a part as UTF-8 text. Throws if the part is missing. */
  readText(path: string): string {
    const part = this._parts.get(path);
    if (!part) {
      throw new OoxmlContainerError("missing-part", `OOXML part not found: ${path}`, { partPath: path });
    }
    let bytes = part.bytes;
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
      bytes = bytes.subarray(3);
    }
    return TEXT_DECODER.decode(bytes);
  }

  readBytes(path: string): Uint8Array {
    const part = this._parts.get(path);
    if (!part) {
      throw new OoxmlContainerError("missing-part", `OOXML part not found: ${path}`, { partPath: path });
    }
    return part.bytes;
  }

  /** Replace a part's contents. Marks dirty. */
  writeText(path: string, content: string): void {
    const bytes = TEXT_ENCODER.encode(content);
    this.writeBytes(path, bytes);
  }

  writeBytes(path: string, content: Uint8Array): void {
    const existing = this._parts.get(path);
    if (existing) {
      existing.bytes = content;
      existing.dirty = true;
    } else {
      this._parts.set(path, { path, bytes: content, dirty: true });
      this._order.push(path);
    }
  }

  /** Add a brand new part. Throws if it already exists. */
  addPart(path: string, content: Uint8Array): void {
    if (this._parts.has(path)) {
      throw new OoxmlContainerError("part-exists", `Cannot addPart; already exists: ${path}`, {
        partPath: path,
      });
    }
    this._parts.set(path, { path, bytes: content, dirty: true });
    this._order.push(path);
  }

  removePart(path: string): void {
    if (!this._parts.has(path)) return;
    this._parts.delete(path);
    const idx = this._order.indexOf(path);
    if (idx >= 0) this._order.splice(idx, 1);
  }

  /** Whether a specific part has been modified since load. */
  isDirty(path: string): boolean {
    return this._parts.get(path)?.dirty ?? false;
  }

  hash(path: string): string {
    return sha256Hex(this.readBytes(path));
  }

  allHashes(): Record<string, string> {
    const out: Record<string, string> = {};
    for (const [path, part] of this._parts) {
      out[path] = sha256Hex(part.bytes);
    }
    return out;
  }

  /**
   * Re-emit the container as a zip ArrayBuffer.
   * - For untouched parts: the original bytes are written.
   * - For dirty parts: the new bytes are written.
   *
   * We use DEFLATE level 6 (JSZip default) for new/dirty parts; untouched
   * parts are re-compressed from their bytes — note that this means the
   * raw zip-level bytes of an untouched part may differ from input even
   * though the part's content bytes are byte-identical. The container's
   * byte-preservation guarantee is at the **part-content** level, not the
   * **zip-archive** level (which is fine because we hash part contents,
   * not the archive).
   */
  async serialize(): Promise<ArrayBuffer> {
    const out = new JSZip();
    for (const path of this._order) {
      const part = this._parts.get(path);
      if (!part) continue;
      out.file(path, part.bytes, {
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });
    }
    const arrayBuffer = await out.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    return arrayBuffer;
  }

  /** Make a shallow clone — parts are shared by reference; new parts/edits don't leak back. */
  clone(): OoxmlContainer {
    const c = new OoxmlContainer();
    for (const path of this._order) {
      const part = this._parts.get(path);
      if (!part) continue;
      c._parts.set(path, { path, bytes: part.bytes, dirty: part.dirty });
      c._order.push(path);
    }
    return c;
  }
}
