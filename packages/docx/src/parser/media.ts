import { ooxml, sha256Hex } from "@officeai/core";
import type { EmbeddedBinaryPart, MediaPart } from "../model/types.js";

/**
 * Walk every part under `word/media/` (case-insensitively, since the OOXML
 * spec is case-insensitive on part names) and lift it into a typed
 * `MediaPart`. The byte slice is shared with the container's part cache —
 * we don't copy because `MediaPart.bytes` is treated as immutable.
 *
 * `mimeType` is resolved from `[Content_Types].xml`: prefer an explicit
 * `<Override PartName>` for that exact part, then fall back to the
 * `<Default Extension>` for the part's extension, and finally to a small
 * built-in extension table for the very common image types so a fixture
 * with a missing default still gets a sensible mime.
 */
export function parseMediaParts(container: ooxml.OoxmlContainer): Map<string, MediaPart> {
  const out = new Map<string, MediaPart>();
  const ct = ooxml.ContentTypes.load(container);

  for (const partPath of container.parts.keys()) {
    if (!isMediaPart(partPath)) continue;
    const bytes = container.readBytes(partPath);
    const mimeType = resolveMimeType(partPath, ct);
    const digest = sha256Hex(bytes);
    out.set(partPath, { partPath, mimeType, bytes, digest });
  }
  return out;
}

/**
 * Built-in fallback table for the image types Word recognises by default.
 * Used only when neither an Override nor a Default in
 * `[Content_Types].xml` claims the extension; this matches Word's own
 * behaviour, which auto-fills the Defaults block on save.
 */
const FALLBACK_MIME_BY_EXT: Readonly<Record<string, string>> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  webp: "image/webp",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
};

export function resolveMimeType(partPath: string, ct: ooxml.ContentTypes): string {
  const partName = `/${partPath}`;
  const override = ct.overrides.find((o) => o.partName === partName);
  if (override) return override.contentType;
  const ext = extensionOf(partPath).toLowerCase();
  if (ext.length > 0) {
    const def = ct.defaults.find((d) => d.extension.toLowerCase() === ext);
    if (def) return def.contentType;
    if (FALLBACK_MIME_BY_EXT[ext]) return FALLBACK_MIME_BY_EXT[ext];
  }
  return "application/octet-stream";
}

export function isMediaPart(partPath: string): boolean {
  return partPath.toLowerCase().startsWith("word/media/");
}

/**
 * Walk every part under `word/embeddings/` and lift it into a typed
 * `EmbeddedBinaryPart`. Used for OLE-Excel `.xlsx` packages and any
 * other binary embedded inside the doc package. Mirrors
 * `parseMediaParts` but for the embeddings directory.
 */
export function parseEmbeddingParts(container: ooxml.OoxmlContainer): Map<string, EmbeddedBinaryPart> {
  const out = new Map<string, EmbeddedBinaryPart>();
  const ct = ooxml.ContentTypes.load(container);
  for (const partPath of container.parts.keys()) {
    if (!isEmbeddingPart(partPath)) continue;
    const bytes = container.readBytes(partPath);
    const contentType = resolveMimeType(partPath, ct);
    out.set(partPath, { partPath, contentType, bytes });
  }
  return out;
}

export function isEmbeddingPart(partPath: string): boolean {
  return partPath.toLowerCase().startsWith("word/embeddings/");
}

export function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf(".");
  if (dot < 0) return "";
  return partPath.slice(dot + 1);
}
