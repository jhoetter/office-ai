import { CommandError, sha256Hex, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  HeaderFooterPart,
  InlineImageDrawing,
  InlineNode,
  MediaPart,
  Paragraph,
  Relationship,
  Run,
  RunChild,
} from "../model/types.js";
import { buildDiff, buildDiffMulti, evolveSnapshot } from "./helpers.js";
import { mintMediaPath, mintDocPrId } from "./insert-image.js";
import { mergeHeaderFooterDirty } from "./set-header-text.js";
import type { InsertHeaderFooterImagePayload } from "./payloads.js";

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const SUPPORTED_MIME_BY_EXT: Readonly<Record<string, string>> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/jpg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
  "image/webp": "webp",
  "image/svg+xml": "svg",
};

/**
 * Insert an inline image into a header or footer part. Mirrors the
 * body-side `docx:insert-image` handler but writes into the part's
 * own relationships file (`word/_rels/{partFileName}.rels`) instead
 * of `word/document.xml.rels`. The image bytes are de-duplicated
 * against existing media by SHA-256 so re-inserting the same logo
 * across multiple sections only mints one `word/media/imageN.{ext}`
 * part.
 *
 * Steps:
 * 1. Validate payload (positive dimensions, supported MIME, byte buffer).
 * 2. SHA-256 the bytes; reuse an existing media part if one matches.
 * 3. Reuse or mint an `image` relationship inside the targeted H/F
 *    part's relationships file. New file at `word/header1.xml` →
 *    rels keyed by `"word/header1.xml"`.
 * 4. Mint a fresh `<wp:docPr id>` unique across the whole document.
 * 5. Build the typed `InlineImageDrawing` (no `raw` so the
 *    serializer regenerates the subtree from the typed model).
 * 6. Splice the image leaf into the targeted paragraph as a fresh
 *    run (or append a new paragraph holding just the image when
 *    `paragraphIndex` is omitted).
 */
export const insertHeaderFooterImageHandler: CommandHandler<InsertHeaderFooterImagePayload, DocxSnapshot> = {
  type: "docx:insert-header-footer-image",
  apply(snapshot, payload, ctx) {
    const validated = validatePayload(payload);
    const partIdx = snapshot.root.headersAndFooters.findIndex((p) => p.partPath === payload.partPath);
    if (partIdx < 0) {
      throw new CommandError("unknown-target", `no header/footer part with path "${payload.partPath}"`);
    }
    const part = snapshot.root.headersAndFooters[partIdx];

    const digest = sha256Hex(validated.bytes);
    const ext = extensionForMime(validated.mimeType);

    const existingMedia = findMediaByDigest(snapshot.root.media, digest);
    let mediaPart: MediaPart;
    let mediaWasAdded = false;
    if (existingMedia) {
      mediaPart = existingMedia;
    } else {
      const newPath = mintMediaPath(snapshot.root.media, ext);
      mediaPart = { partPath: newPath, mimeType: validated.mimeType, bytes: validated.bytes, digest };
      mediaWasAdded = true;
    }

    // Resolve relationship id within the H/F part's own rels file.
    const partRelsKey = part.partPath;
    const partRels = snapshot.root.relationships.get(partRelsKey) ?? [];
    // Targets are stored relative to the part's containing folder;
    // both `word/header1.xml` and `word/media/image1.png` live under
    // `word/`, so the relative target is `media/imageN.{ext}`.
    const relTarget = mediaPart.partPath.startsWith("word/")
      ? mediaPart.partPath.slice("word/".length)
      : mediaPart.partPath;
    let relId: string;
    let relsChanged = false;
    let nextRels: ReadonlyArray<Relationship> = partRels;
    const existingRel = partRels.find((r) => r.type === IMAGE_REL_TYPE && r.target === relTarget);
    if (existingRel) {
      relId = existingRel.id;
    } else {
      relId = mintRelId(partRels);
      const newRel: Relationship = { id: relId, type: IMAGE_REL_TYPE, target: relTarget };
      nextRels = [...partRels, newRel];
      relsChanged = true;
    }

    const docPrId = mintDocPrId(snapshot.root);

    const cx = pxToEmu(validated.width);
    const cy = pxToEmu(validated.height);
    const leaf: InlineImageDrawing = {
      kind: "drawing",
      subkind: "inline-image",
      id: ctx.mintNodeId(),
      relId,
      cx,
      cy,
      docPrId,
      name: payload.name ?? `Picture ${docPrId}`,
      ...(payload.altText !== undefined ? { descr: payload.altText } : {}),
      properties: { distT: 0, distB: 0, distL: 0, distR: 0 },
    };

    const newBody = spliceImageIntoPart(part.body, payload, leaf, ctx.mintNodeId);

    const updatedPart: HeaderFooterPart = { ...part, body: newBody };
    const newParts = snapshot.root.headersAndFooters.slice();
    newParts[partIdx] = updatedPart;

    const newMedia = mediaWasAdded
      ? new Map(snapshot.root.media).set(mediaPart.partPath, mediaPart)
      : snapshot.root.media;
    const newRelationshipsMap = relsChanged
      ? new Map(snapshot.root.relationships).set(partRelsKey, nextRels)
      : snapshot.root.relationships;

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      headersAndFooters: newParts,
      media: newMedia,
      relationships: newRelationshipsMap,
    };

    const nextMediaDirty: ReadonlySet<string> = mediaWasAdded
      ? withAddition(snapshot.dirty.media, mediaPart.partPath)
      : snapshot.dirty.media;
    const nextRelsDirty: ReadonlySet<string> = relsChanged
      ? withAddition(snapshot.dirty.relationships, partRelsKey)
      : snapshot.dirty.relationships;
    const nextContentTypesDirty = mediaWasAdded || snapshot.dirty.contentTypes;
    const nextHfDirty = mergeHeaderFooterDirty(snapshot.dirty, part.partPath);

    const next = evolveSnapshot(snapshot, nextDoc, {
      headersAndFooters: nextHfDirty,
      media: nextMediaDirty,
      relationships: nextRelsDirty,
      contentTypes: nextContentTypesDirty,
    });

    const insertedChange = {
      kind: "node-inserted" as const,
      nodeId: leaf.id,
      path: ["headersAndFooters", partIdx, "image"],
      summary: `+image (${cx}×${cy} EMU, rel ${relId} in ${part.kind} ${part.partPath})`,
    };
    if (mediaWasAdded) {
      return {
        next,
        diff: buildDiffMulti(snapshot.revision, next.revision, [
          insertedChange,
          { kind: "part-added", path: [mediaPart.partPath], summary: `+media ${mediaPart.partPath}` },
        ]),
      };
    }
    return { next, diff: buildDiff(snapshot.revision, next.revision, insertedChange) };
  },
};

function findMediaByDigest(media: ReadonlyMap<string, MediaPart>, digest: string): MediaPart | undefined {
  for (const part of media.values()) {
    if (part.digest === digest) return part;
  }
  return undefined;
}

function mintRelId(rels: ReadonlyArray<Relationship>): string {
  const taken = new Set(rels.map((r) => r.id));
  let i = rels.length + 1;
  while (taken.has(`rId${i}`)) i++;
  return `rId${i}`;
}

function pxToEmu(px: number): number {
  return Math.round(px * 9525);
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}

interface ValidatedPayload {
  bytes: Uint8Array;
  mimeType: string;
  width: number;
  height: number;
}

function validatePayload(payload: InsertHeaderFooterImagePayload): ValidatedPayload {
  if (!payload.mimeType || typeof payload.mimeType !== "string") {
    throw new CommandError("invalid-payload", "mimeType is required");
  }
  const lc = payload.mimeType.toLowerCase();
  if (!(lc in SUPPORTED_MIME_BY_EXT)) {
    throw new CommandError(
      "invalid-payload",
      `unsupported image MIME type: ${payload.mimeType} (supported: ${Object.keys(SUPPORTED_MIME_BY_EXT).join(", ")})`
    );
  }
  if (!Number.isFinite(payload.width) || payload.width <= 0) {
    throw new CommandError("invalid-payload", `width must be a positive number (got ${payload.width})`);
  }
  if (!Number.isFinite(payload.height) || payload.height <= 0) {
    throw new CommandError("invalid-payload", `height must be a positive number (got ${payload.height})`);
  }
  const bytes = toUint8Array(payload.data);
  if (bytes.byteLength === 0) {
    throw new CommandError("invalid-payload", "image data is empty");
  }
  return { bytes, mimeType: lc, width: payload.width, height: payload.height };
}

function toUint8Array(data: Uint8Array | ArrayBuffer): Uint8Array {
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data);
}

function extensionForMime(mime: string): string {
  return SUPPORTED_MIME_BY_EXT[mime] ?? "bin";
}

/**
 * Splice the image leaf into the part's body. When the caller
 * supplies a `paragraphIndex` we splice a fresh image-only run
 * inside that paragraph (preserving any text already there);
 * otherwise we append a new paragraph at the end of the body and
 * place the image there.
 */
function spliceImageIntoPart(
  body: ReadonlyArray<BlockNode>,
  payload: InsertHeaderFooterImagePayload,
  leaf: InlineImageDrawing,
  mintNodeId: IdMinter
): BlockNode[] {
  const next = body.slice();
  const imageRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [leaf as RunChild],
  };
  if (typeof payload.paragraphIndex === "number") {
    const block = next[payload.paragraphIndex];
    if (!block || block.kind !== "paragraph") {
      throw new CommandError(
        "unknown-target",
        `block at index ${payload.paragraphIndex} in ${payload.partPath} is not a paragraph`
      );
    }
    const para = block;
    const inserted: InlineNode[] = [...para.children, imageRun];
    const updated: Paragraph = { ...para, children: inserted };
    next[payload.paragraphIndex] = updated;
    return next;
  }
  const para: Paragraph = {
    kind: "paragraph",
    id: mintNodeId(),
    properties: {},
    children: [imageRun],
  };
  next.push(para);
  return next;
}
