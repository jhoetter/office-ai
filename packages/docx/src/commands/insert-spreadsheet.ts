import { CommandError, sha256Hex, type CommandHandler, type IdMinter } from "@officeai/core";
import { gridToPng } from "@officeai/xlsx";
import type {
  DocxDocument,
  DocxSnapshot,
  EmbeddedBinaryPart,
  EmbeddedSpreadsheet,
  InlineNode,
  MediaPart,
  Paragraph,
  Relationship,
  Run,
  RunChild,
} from "../model/types.js";
import { buildDiffMulti, evolveSnapshot } from "./helpers.js";
import { mintMediaPath } from "./insert-image.js";
import type { InsertSpreadsheetPayload, UpdateSpreadsheetPayload } from "./payloads.js";

/**
 * Insert an OLE-Excel spreadsheet at the targeted paragraph position.
 *
 * Pipeline (mirrors `docx:insert-image` + `docx:insert-chart`):
 *   1. Validate payload (non-empty grid).
 *   2. Build the embedded `.xlsx` bytes via {@link buildEmbeddedXlsx}.
 *   3. Render a PNG preview of the grid via {@link gridToPng}.
 *   4. Mint `word/embeddings/oleObjectN.xlsx` and
 *      `word/media/imageN.png` part paths.
 *   5. Add two rels in `word/document.xml.rels`: `oleObject` →
 *      embedding, `image` → preview PNG.
 *   6. Build a typed {@link EmbeddedSpreadsheet} leaf wired with both
 *      rel ids; the run-child serializer emits the `<w:object>`
 *      envelope, the embeddings serializer materialises the xlsx
 *      bytes + content-type override, and the media serializer
 *      writes the preview PNG.
 *   7. Splice the leaf into the targeted paragraph (run-aware splice,
 *      same shape as `insert-image`).
 *   8. Mark dirty: body, embeddings, media, relationships,
 *      contentTypes.
 */
export const insertSpreadsheetHandler: CommandHandler<InsertSpreadsheetPayload, DocxSnapshot> = {
  type: "docx:insert-spreadsheet",
  apply(snapshot, payload, ctx) {
    validatePayload(payload);

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

    const sheetName = payload.sheetName ?? "Sheet1";
    const preview = gridToPng(payload.data);

    const embeddingPath = mintEmbeddingPath(snapshot.root.embeddings);
    const previewPath = mintMediaPath(snapshot.root.media, "png");

    const docRelsKey = "word/document.xml";
    const docRels = snapshot.root.relationships.get(docRelsKey) ?? [];

    const embedRelId = mintRelId(docRels);
    const embedRel: Relationship = {
      id: embedRelId,
      type: OLE_REL_TYPE,
      target: relativeTargetFromDocument(embeddingPath),
    };
    const imageRelId = mintRelId([...docRels, embedRel]);
    const imageRel: Relationship = {
      id: imageRelId,
      type: IMAGE_REL_TYPE,
      target: relativeTargetFromDocument(previewPath),
    };
    const nextRels: ReadonlyArray<Relationship> = [...docRels, embedRel, imageRel];

    const embedded: EmbeddedSpreadsheet = {
      kind: "embedded-spreadsheet",
      id: ctx.mintNodeId(),
      oleRelId: embedRelId,
      embeddingPartPath: embeddingPath,
      progId: "Excel.Sheet.12",
      embeddingKind: "xlsx",
      previewImageRelId: imageRelId,
      previewImagePartPath: previewPath,
      oleObjectAttrs: {
        Type: "Embed",
        DrawAspect: "Content",
      },
    };

    const updatedParagraph = insertLeafIntoParagraph(
      target,
      payload.at.run,
      payload.at.offset ?? 0,
      embedded,
      ctx.mintNodeId
    );
    const newBody = snapshot.root.body.slice();
    newBody[payload.at.paragraph] = updatedParagraph;

    const newEmbeddings = new Map(snapshot.root.embeddings);
    const embeddingPart: EmbeddedBinaryPart = {
      partPath: embeddingPath,
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      pendingGrid: payload.data,
      pendingSheetName: sheetName,
    };
    newEmbeddings.set(embeddingPath, embeddingPart);

    const newMedia = new Map(snapshot.root.media);
    const previewPart: MediaPart = {
      partPath: previewPath,
      mimeType: "image/png",
      bytes: preview.bytes,
      digest: sha256Hex(preview.bytes),
    };
    newMedia.set(previewPath, previewPart);

    const newRelationships = new Map(snapshot.root.relationships);
    newRelationships.set(docRelsKey, nextRels);

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      body: newBody,
      embeddings: newEmbeddings,
      media: newMedia,
      relationships: newRelationships,
    };

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      embeddings: withAddition(snapshot.dirty.embeddings, embeddingPath),
      media: withAddition(snapshot.dirty.media, previewPath),
      relationships: withAddition(snapshot.dirty.relationships, docRelsKey),
      contentTypes: true,
    });

    return {
      next,
      diff: buildDiffMulti(snapshot.revision, next.revision, [
        {
          kind: "node-inserted",
          nodeId: embedded.id,
          path: ["body", payload.at.paragraph, "spreadsheet"],
          summary: `+spreadsheet (${payload.data.length}×${payload.data.reduce((m, r) => Math.max(m, r.length), 0)})`,
        },
        {
          kind: "part-added",
          path: [embeddingPath],
          summary: `+embedding ${embeddingPath}`,
        },
        {
          kind: "part-added",
          path: [previewPath],
          summary: `+media ${previewPath}`,
        },
      ]),
    };
  },
};

/**
 * Replace the bytes of an existing embedded spreadsheet (used by the
 * editor's "double-click → edit → save" round-trip flow). Optionally
 * regenerates the cached PNG preview from a fresh grid snapshot.
 */
export const updateSpreadsheetHandler: CommandHandler<UpdateSpreadsheetPayload, DocxSnapshot> = {
  type: "docx:update-spreadsheet",
  apply(snapshot, payload) {
    if (!payload.embeddingPartPath || typeof payload.embeddingPartPath !== "string") {
      throw new CommandError("invalid-payload", "embeddingPartPath is required");
    }
    const existing = snapshot.root.embeddings.get(payload.embeddingPartPath);
    if (!existing) {
      throw new CommandError("missing-embedding", `no embedded part at ${payload.embeddingPartPath}`);
    }
    if (!(payload.bytes instanceof Uint8Array) || payload.bytes.byteLength === 0) {
      throw new CommandError("invalid-payload", "bytes must be a non-empty Uint8Array");
    }

    const newEmbeddings = new Map(snapshot.root.embeddings);
    newEmbeddings.set(payload.embeddingPartPath, {
      ...existing,
      bytes: payload.bytes,
    });

    let newMedia = snapshot.root.media;
    let mediaDirty = snapshot.dirty.media;
    let previewPath: string | undefined;
    if (payload.previewGrid) {
      const leafPreviewPath = findPreviewPathForEmbedding(snapshot, payload.embeddingPartPath);
      if (leafPreviewPath) {
        const png = gridToPng(payload.previewGrid);
        const prev = snapshot.root.media.get(leafPreviewPath);
        const next = new Map(snapshot.root.media);
        next.set(leafPreviewPath, {
          partPath: leafPreviewPath,
          mimeType: prev?.mimeType ?? "image/png",
          bytes: png.bytes,
          digest: sha256Hex(png.bytes),
        });
        newMedia = next;
        mediaDirty = withAddition(mediaDirty, leafPreviewPath);
        previewPath = leafPreviewPath;
      }
    }

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      embeddings: newEmbeddings,
      media: newMedia,
    };
    const next = evolveSnapshot(snapshot, nextDoc, {
      embeddings: withAddition(snapshot.dirty.embeddings, payload.embeddingPartPath),
      media: mediaDirty,
    });

    const path = ["embeddings", payload.embeddingPartPath];
    return {
      next,
      diff: buildDiffMulti(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: payload.embeddingPartPath,
          path,
          field: "bytes",
          summary: `~spreadsheet ${payload.embeddingPartPath}${previewPath ? " (preview)" : ""}`,
        },
      ]),
    };
  },
};

const OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function validatePayload(p: InsertSpreadsheetPayload): void {
  if (!Array.isArray(p.data) || p.data.length === 0) {
    throw new CommandError("invalid-payload", "data must contain at least one row");
  }
  for (let i = 0; i < p.data.length; i++) {
    if (!Array.isArray(p.data[i])) {
      throw new CommandError("invalid-payload", `data[${i}] must be an array`);
    }
  }
}

function mintEmbeddingPath(map: ReadonlyMap<string, EmbeddedBinaryPart>): string {
  let n = 1;
  while (map.has(`word/embeddings/oleObject${n}.xlsx`)) n++;
  return `word/embeddings/oleObject${n}.xlsx`;
}

function relativeTargetFromDocument(partPath: string): string {
  return partPath.startsWith("word/") ? partPath.slice("word/".length) : partPath;
}

function mintRelId(rels: ReadonlyArray<Relationship>): string {
  const taken = new Set(rels.map((r) => r.id));
  let i = rels.length + 1;
  while (taken.has(`rId${i}`)) i++;
  return `rId${i}`;
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}

function findPreviewPathForEmbedding(snapshot: DocxSnapshot, embeddingPath: string): string | undefined {
  for (const block of snapshot.root.body) {
    if (block.kind !== "paragraph") continue;
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      for (const child of inline.children) {
        if (
          child.kind === "embedded-spreadsheet" &&
          child.embeddingPartPath === embeddingPath &&
          child.previewImagePartPath
        ) {
          return child.previewImagePartPath;
        }
      }
    }
  }
  return undefined;
}

function insertLeafIntoParagraph(
  p: Paragraph,
  runIndex: number | undefined,
  offset: number,
  leaf: EmbeddedSpreadsheet,
  mintNodeId: IdMinter
): Paragraph {
  const leafRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [leaf as RunChild],
  };
  if (runIndex === undefined) {
    return { ...p, children: [leafRun, ...p.children] };
  }
  if (runIndex < 0 || runIndex >= p.children.length) {
    return { ...p, children: [...p.children, leafRun] };
  }
  const target = p.children[runIndex];
  if (target.kind !== "run") {
    const next = p.children.slice();
    next.splice(runIndex, 0, leafRun);
    return { ...p, children: next };
  }
  const { before, after } = splitRunAtOffset(target, offset, mintNodeId);
  const next: InlineNode[] = [];
  for (let i = 0; i < p.children.length; i++) {
    if (i === runIndex) {
      if (before) next.push(before);
      next.push(leafRun);
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
