import { CommandError, sha256Hex, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  DrawingLeaf,
  InlineImageDrawing,
  InlineNode,
  MediaPart,
  Paragraph,
  Relationship,
  Run,
  RunChild,
} from "../model/types.js";
import { buildDiff, buildDiffMulti, evolveSnapshot } from "./helpers.js";
import type { InsertImagePayload } from "./payloads.js";

/**
 * Insert an inline image at the given position.
 *
 * Steps:
 * 1. Validate the payload (positive dimensions, supported MIME, byte buffer).
 * 2. SHA-256 the bytes and try to reuse an existing media part with the
 *    same digest. If found, also try to reuse the existing
 *    `image`-relationship rId pointing to it.
 * 3. Otherwise mint a fresh `word/media/image{N}.{ext}` part path and a
 *    fresh `rId{N}` in `word/document.xml.rels`.
 * 4. Mint a unique `<wp:docPr id>` (max existing in body + 1).
 * 5. Convert the pixel-sized payload to EMUs (9525 EMU = 1 px @ 96 DPI).
 * 6. Build a typed `InlineImageDrawing` (no `raw`, so the serializer
 *    regenerates the subtree from the typed model).
 * 7. Splice the image into the targeted paragraph as a new run.
 * 8. Update `media`, `relationships`, dirty flags, and emit a two-change
 *    diff: a `node-inserted` for the drawing leaf plus a `part-added`
 *    for the media part (the latter only when we actually created a new
 *    media part).
 */
export const insertImageHandler: CommandHandler<InsertImagePayload, DocxSnapshot> = {
  type: "docx:insert-image",
  apply(snapshot, payload, ctx) {
    const validated = validatePayload(payload);
    const bytes = validated.bytes;

    const bodyLen = snapshot.root.body.length;
    if (
      !Number.isInteger(payload.at.paragraph) ||
      payload.at.paragraph < 0 ||
      payload.at.paragraph >= bodyLen
    ) {
      throw new CommandError(
        "invalid-position",
        `paragraph index ${payload.at.paragraph} out of range [0, ${bodyLen})`
      );
    }
    const target = snapshot.root.body[payload.at.paragraph];
    if (target.kind !== "paragraph") {
      throw new CommandError(
        "not-paragraph",
        `block at body index ${payload.at.paragraph} is not a paragraph (kind=${target.kind})`
      );
    }

    const digest = sha256Hex(bytes);
    const ext = extensionForMime(validated.mimeType);

    // Step 1: media de-dup. Find an existing media part with the same SHA-256.
    const existingMedia = findMediaByDigest(snapshot.root.media, digest);
    let mediaPart: MediaPart;
    let mediaWasAdded = false;
    if (existingMedia) {
      mediaPart = existingMedia;
    } else {
      const partPath = mintMediaPath(snapshot.root.media, ext);
      mediaPart = { partPath, mimeType: validated.mimeType, bytes, digest };
      mediaWasAdded = true;
    }

    // Step 2: relationship reuse / minting against word/document.xml.rels.
    const docRelsKey = "word/document.xml";
    const docRels = snapshot.root.relationships.get(docRelsKey) ?? [];
    const relTarget = relTargetFor(mediaPart.partPath);
    let relId: string;
    let nextRels: ReadonlyArray<Relationship> = docRels;
    let relsChanged = false;
    const existingRel = docRels.find((r) => r.type === IMAGE_REL_TYPE && r.target === relTarget);
    if (existingRel) {
      relId = existingRel.id;
    } else {
      relId = mintRelId(docRels);
      const newRel: Relationship = { id: relId, type: IMAGE_REL_TYPE, target: relTarget };
      nextRels = [...docRels, newRel];
      relsChanged = true;
    }

    // Step 3: docPr id uniqueness across the entire body.
    const docPrId = mintDocPrId(snapshot.root);

    // Step 4: build the typed leaf.
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
      properties: {
        distT: 0,
        distB: 0,
        distL: 0,
        distR: 0,
      },
    };

    // Step 5: splice into the paragraph.
    const updatedParagraph = insertDrawingIntoParagraph(
      target,
      payload.at.run,
      payload.at.offset ?? 0,
      leaf,
      ctx.mintNodeId
    );
    const newBody = snapshot.root.body.slice();
    newBody[payload.at.paragraph] = updatedParagraph;

    // Step 6: assemble the new document.
    const newMedia = mediaWasAdded
      ? new Map(snapshot.root.media).set(mediaPart.partPath, mediaPart)
      : snapshot.root.media;
    const newRelationshipsMap = relsChanged
      ? new Map(snapshot.root.relationships).set(docRelsKey, nextRels)
      : snapshot.root.relationships;

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      body: newBody,
      media: newMedia,
      relationships: newRelationshipsMap,
    };

    // Step 7: dirty flags. We always dirty body. We dirty media + rels +
    // contentTypes only when we actually added new package parts.
    const nextMediaDirty: ReadonlySet<string> = mediaWasAdded
      ? withAddition(snapshot.dirty.media, mediaPart.partPath)
      : snapshot.dirty.media;
    const nextRelsDirty: ReadonlySet<string> = relsChanged
      ? withAddition(snapshot.dirty.relationships, docRelsKey)
      : snapshot.dirty.relationships;
    const nextContentTypesDirty = mediaWasAdded || snapshot.dirty.contentTypes;

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      media: nextMediaDirty,
      relationships: nextRelsDirty,
      contentTypes: nextContentTypesDirty,
    });

    // Step 8: diff. Two-change diff when a new media part was added,
    // single-change diff when we de-duplicated against an existing part.
    const insertedChange = {
      kind: "node-inserted" as const,
      nodeId: leaf.id,
      path: ["body", payload.at.paragraph, "image"],
      summary: `+image (${cx}×${cy} EMU, rel ${relId})`,
    };
    if (mediaWasAdded) {
      return {
        next,
        diff: buildDiffMulti(snapshot.revision, next.revision, [
          insertedChange,
          {
            kind: "part-added",
            path: [mediaPart.partPath],
            summary: `+media ${mediaPart.partPath}`,
          },
        ]),
      };
    }
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, insertedChange),
    };
  },
};

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

interface ValidatedPayload {
  bytes: Uint8Array;
  mimeType: string;
  width: number;
  height: number;
}

function validatePayload(payload: InsertImagePayload): ValidatedPayload {
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

function findMediaByDigest(media: ReadonlyMap<string, MediaPart>, digest: string): MediaPart | undefined {
  for (const part of media.values()) {
    if (part.digest === digest) return part;
  }
  return undefined;
}

/**
 * Pick the next free `word/media/image{N}.{ext}` part path. We start at 1
 * and walk up until we find a path no existing media part claims, mirroring
 * Word's own naming convention so a saved-then-reopened file stays
 * readable to humans inspecting the package.
 */
export function mintMediaPath(media: ReadonlyMap<string, MediaPart>, ext: string): string {
  const taken = new Set(media.keys());
  let i = 1;
  while (taken.has(`word/media/image${i}.${ext}`)) i++;
  return `word/media/image${i}.${ext}`;
}

function relTargetFor(partPath: string): string {
  // `word/media/imageN.ext` is referenced from `word/document.xml` as
  // `media/imageN.ext` (relative to `word/`).
  return partPath.startsWith("word/") ? partPath.slice("word/".length) : partPath;
}

function mintRelId(rels: ReadonlyArray<Relationship>): string {
  const taken = new Set(rels.map((r) => r.id));
  let i = rels.length + 1;
  while (taken.has(`rId${i}`)) i++;
  return `rId${i}`;
}

/**
 * Walk the entire document body collecting every `<wp:docPr id>` we
 * already use, and return one greater than the max. `docPrId` must be
 * unique across the whole document (not just one paragraph) per the
 * OOXML spec; Word will silently re-number on save when this collides
 * but we'd rather be deterministic.
 */
export function mintDocPrId(doc: DocxDocument): number {
  let max = 0;
  walkInlineImages(doc, (leaf) => {
    if (leaf.docPrId > max) max = leaf.docPrId;
  });
  return max + 1;
}

function walkInlineImages(doc: DocxDocument, visit: (leaf: InlineImageDrawing) => void): void {
  for (const block of doc.body) {
    walkBlockForImages(block, visit);
  }
}

function walkBlockForImages(block: BlockNode, visit: (leaf: InlineImageDrawing) => void): void {
  if (block.kind === "paragraph") {
    for (const inline of block.children) walkInlineForImages(inline, visit);
  } else if (block.kind === "table") {
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const b of cell.body) {
          walkBlockForImages(b, visit);
        }
      }
    }
  }
}

function walkInlineForImages(node: InlineNode, visit: (leaf: InlineImageDrawing) => void): void {
  if (node.kind === "run") {
    for (const c of node.children) {
      if (c.kind === "drawing" && c.subkind === "inline-image") visit(c);
    }
  } else if (node.kind === "hyperlink") {
    for (const r of node.children) walkInlineForImages(r, visit);
  } else if (node.kind === "revision") {
    for (const ic of node.children) walkInlineForImages(ic, visit);
  }
}

function pxToEmu(px: number): number {
  return Math.round(px * 9525);
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}

/**
 * Splice a new run carrying the image leaf into a paragraph.
 *
 * - When `runIndex` is undefined: prepend the image as the first inline
 *   of the paragraph (matches `docx:insert-text` no-target semantics).
 * - When `runIndex` is set but the targeted child is not a run (e.g. a
 *   hyperlink or comment marker) we insert the image run *before* it,
 *   preserving the surrounding inlines.
 * - When `runIndex` is set and points at a run, we split the run's text
 *   children at `offset` so the image lands between the two halves.
 */
function insertDrawingIntoParagraph(
  p: Paragraph,
  runIndex: number | undefined,
  offset: number,
  leaf: DrawingLeaf,
  mintNodeId: IdMinter
): Paragraph {
  const imageRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [leaf as RunChild],
  };

  if (runIndex === undefined) {
    return { ...p, children: [imageRun, ...p.children] };
  }
  if (runIndex < 0 || runIndex >= p.children.length) {
    return { ...p, children: [...p.children, imageRun] };
  }
  const target = p.children[runIndex];
  if (target.kind !== "run") {
    const next = p.children.slice();
    next.splice(runIndex, 0, imageRun);
    return { ...p, children: next };
  }

  const { before, after } = splitRunAtOffset(target, offset, mintNodeId);
  const next: InlineNode[] = [];
  for (let i = 0; i < p.children.length; i++) {
    if (i === runIndex) {
      if (before) next.push(before);
      next.push(imageRun);
      if (after) next.push(after);
    } else {
      next.push(p.children[i]);
    }
  }
  return { ...p, children: next };
}

interface SplitRun {
  before: Run | null;
  after: Run | null;
}

function splitRunAtOffset(run: Run, offset: number, mintNodeId: IdMinter): SplitRun {
  const beforeChildren: RunChild[] = [];
  const afterChildren: RunChild[] = [];
  let consumed = 0;
  let placed = false;
  for (const c of run.children) {
    if (placed) {
      afterChildren.push(c);
      continue;
    }
    if (c.kind !== "text") {
      beforeChildren.push(c);
      continue;
    }
    const len = c.text.length;
    if (offset >= consumed + len) {
      beforeChildren.push(c);
      consumed += len;
      continue;
    }
    const local = Math.max(0, offset - consumed);
    if (local > 0) beforeChildren.push({ ...c, text: c.text.slice(0, local) });
    if (local < len) {
      afterChildren.push({ ...c, id: mintNodeId(), text: c.text.slice(local) });
    }
    placed = true;
    consumed += len;
  }
  const before: Run | null =
    beforeChildren.length > 0 ? { ...run, id: mintNodeId(), children: beforeChildren } : null;
  const after: Run | null =
    afterChildren.length > 0 ? { ...run, id: mintNodeId(), children: afterChildren } : null;
  return { before, after };
}
