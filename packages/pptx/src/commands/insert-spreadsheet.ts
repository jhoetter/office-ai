import { type CommandHandler, sha256Hex } from "@officeai/core";
import { gridToPng } from "@officeai/xlsx";
import type {
  ContentTypesSnap,
  EmbeddedBinaryPart,
  MediaPart,
  OleSpreadsheetShape,
  OpaqueXml,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
} from "../model/types.js";
import { evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { PptxInsertSpreadsheetPayload, PptxUpdateSpreadsheetPayload } from "./payloads.js";

const REL_TYPE_OLE_OBJECT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const OLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/presentationml/2006/ole";
const CT_SPREADSHEETML_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * Insert an OLE-Excel spreadsheet onto a slide. Authors a typed
 * {@link OleSpreadsheetShape} that the serializer emits as a
 * `<p:graphicFrame>` carrying a `<p:oleObj progId="Excel.Sheet.12">`,
 * plus an embedded `.xlsx` package and a PNG preview image so Office
 * has something to render before the user double-clicks to activate
 * Excel for editing.
 *
 * The handler is synchronous: the embedded workbook bytes are
 * deferred to the serializer (via `EmbeddedBinaryPart.pendingGrid`),
 * which calls `buildEmbeddedXlsx` lazily. The preview PNG, by
 * contrast, is small and pure-JS so it's materialised eagerly.
 */
export const insertSpreadsheetHandler: CommandHandler<PptxInsertSpreadsheetPayload, PptxSnapshot> = {
  type: "pptx:insert-spreadsheet",
  apply(snapshot, payload, ctx) {
    if (!Array.isArray(payload.data) || payload.data.length === 0) {
      throw makeError("invalid-payload", "data must contain at least one row");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const slideRelsPath = relsPathFor(slide.partPath);

    const sheetName = payload.sheetName ?? "Sheet1";
    const cx = Math.round(payload.cx ?? 2_743_200);
    const cy = Math.round(payload.cy ?? 1_828_800);

    // Render the static preview image. Pure JS, tiny — fine to do
    // synchronously in the command handler.
    const previewBytes = gridToPng(payload.data).bytes;
    const previewSha = sha256Hex(previewBytes);

    // Dedup the preview against existing media (re-inserting the same
    // spreadsheet shouldn't churn out duplicate `image{N}.png` parts).
    let mediaPartPath = "";
    for (const m of snapshot.root.media.values()) {
      if (m.sha256 === previewSha) {
        mediaPartPath = m.partPath;
        break;
      }
    }
    let nextMediaPartIndex = snapshot.root.idGen.nextMediaPartIndex;
    let newMedia: ReadonlyMap<string, MediaPart> = snapshot.root.media;
    let mediaWasNew = false;
    if (!mediaPartPath) {
      mediaPartPath = `ppt/media/image${nextMediaPartIndex}.png`;
      const mediaPart: MediaPart = {
        partPath: mediaPartPath,
        bytes: previewBytes,
        sha256: previewSha,
        contentType: "image/png",
      };
      const m = new Map(snapshot.root.media);
      m.set(mediaPartPath, mediaPart);
      newMedia = m;
      nextMediaPartIndex += 1;
      mediaWasNew = true;
    }

    // Mint a fresh `ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`
    // path. The serializer materialises the actual bytes from
    // `pendingGrid` on first flush.
    const embeddingPath = mintEmbeddingPath(snapshot.root.embeddings);

    // Mint two slide rels: one for the OLE object package, one for
    // the preview image.
    const slideRels = snapshot.relationships.get(slideRelsPath);
    const existingEntries = slideRels?.entries ?? [];
    const oleRelId = nextRelId(existingEntries.map((e) => e.id));
    const imageRelId = nextRelId([...existingEntries.map((e) => e.id), oleRelId]);
    const newSlideEntries = [
      ...existingEntries,
      {
        id: oleRelId,
        type: REL_TYPE_OLE_OBJECT,
        target: relativeFromRels(slideRelsPath, embeddingPath),
      },
      {
        id: imageRelId,
        type: REL_TYPE_IMAGE,
        target: relativeFromRels(slideRelsPath, mediaPartPath),
      },
    ];
    const newRelsMap = new Map(snapshot.relationships);
    newRelsMap.set(slideRelsPath, { relsPath: slideRelsPath, entries: newSlideEntries });

    // Register `image/png` as a default content type when this is the
    // first PNG we've added to the deck (mirrors `insertImageHandler`).
    let newContentTypes: ContentTypesSnap = snapshot.contentTypes;
    let contentTypesDirty = false;
    if (mediaWasNew) {
      const hasPng = snapshot.contentTypes.defaults.some((d) => d.extension.toLowerCase() === "png");
      if (!hasPng) {
        newContentTypes = {
          ...snapshot.contentTypes,
          defaults: [...snapshot.contentTypes.defaults, { extension: "png", contentType: "image/png" }],
        };
        contentTypesDirty = true;
      }
    }

    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const previewPicId = cNvPrId + 1;
    const ole: OleSpreadsheetShape = {
      kind: "ole-spreadsheet",
      id: ctx.mintNodeId(),
      cNvPrId,
      name: payload.name ?? `Embedded Spreadsheet ${cNvPrId}`,
      position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
      size: { cxEmu: cx, cyEmu: cy },
      oleRelId,
      embeddingPartPath: embeddingPath,
      progId: "Excel.Sheet.12",
      embeddingKind: "xlsx",
      previewMediaRelId: imageRelId,
      previewMediaPartPath: mediaPartPath,
      oleObjAttrs: {
        showAsIcon: "0",
        imgW: String(cx),
        imgH: String(cy),
      },
      oleObjChildrenRaw: buildOleObjChildren({
        previewMediaRelId: imageRelId,
        previewPicId,
        previewName: `${payload.name ?? "Embedded Spreadsheet"} Preview`,
        position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
        size: { cxEmu: cx, cyEmu: cy },
      }),
      nvGraphicFramePrTail: [],
      graphicDataUri: OLE_GRAPHIC_DATA_URI,
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, ole] };
    const newSlides = [...snapshot.root.slides];
    newSlides[sIdx] = newSlide;

    const newEmbeddings = new Map(snapshot.root.embeddings);
    const embeddingPart: EmbeddedBinaryPart = {
      partPath: embeddingPath,
      contentType: CT_SPREADSHEETML_SHEET,
      pendingGrid: payload.data,
      pendingSheetName: sheetName,
    };
    newEmbeddings.set(embeddingPath, embeddingPart);

    const root: PptxPresentation = {
      ...snapshot.root,
      slides: newSlides,
      media: newMedia,
      embeddings: newEmbeddings,
      idGen: { ...snapshot.root.idGen, nextMediaPartIndex },
    };

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        slides: [slide.partPath],
        media: mediaWasNew ? [mediaPartPath] : [],
        embeddings: [embeddingPath],
        relationships: [slideRelsPath],
        contentTypes: contentTypesDirty,
      },
      {
        relationships: newRelsMap as ReadonlyMap<string, RelationshipsSnap>,
        contentTypes: newContentTypes,
      }
    );

    const cols = payload.data.reduce((m, r) => Math.max(m, r.length), 0);
    return {
      next,
      diff: {
        format: "pptx",
        fromRevision: snapshot.revision,
        toRevision: next.revision,
        changes: [
          {
            kind: "node-inserted",
            nodeId: ole.id,
            path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1] as ReadonlyArray<string | number>,
            summary: `+spreadsheet (${payload.data.length}×${cols})`,
          },
          {
            kind: "part-added",
            path: [embeddingPath],
            summary: `+embedding ${embeddingPath}`,
          },
        ],
      },
    };
  },
};

/**
 * Replace the bytes of an existing PPTX OLE-Excel embed. Used by the
 * editor's double-click → edit-in-XlsxAgent → save loop, mirroring
 * `docx:update-spreadsheet`.
 */
export const updateSpreadsheetHandler: CommandHandler<PptxUpdateSpreadsheetPayload, PptxSnapshot> = {
  type: "pptx:update-spreadsheet",
  apply(snapshot, payload) {
    const existing = snapshot.root.embeddings.get(payload.embeddingPartPath);
    if (!existing) {
      throw makeError("missing-embedding", `no embedded part at ${payload.embeddingPartPath}`);
    }
    if (!(payload.bytes instanceof Uint8Array) || payload.bytes.byteLength === 0) {
      throw makeError("invalid-payload", "bytes must be a non-empty Uint8Array");
    }

    const newEmbeddings = new Map(snapshot.root.embeddings);
    newEmbeddings.set(payload.embeddingPartPath, {
      partPath: existing.partPath,
      contentType: existing.contentType,
      bytes: payload.bytes,
    });

    let newMedia = snapshot.root.media;
    let mediaDirty: string[] = [];
    if (payload.previewGrid) {
      const previewPath = findPreviewPathForEmbedding(snapshot, payload.embeddingPartPath);
      if (previewPath) {
        const png = gridToPng(payload.previewGrid).bytes;
        const next = new Map(snapshot.root.media);
        const prev = snapshot.root.media.get(previewPath);
        next.set(previewPath, {
          partPath: previewPath,
          bytes: png,
          sha256: sha256Hex(png),
          contentType: prev?.contentType ?? "image/png",
        });
        newMedia = next;
        mediaDirty = [previewPath];
      }
    }

    const root: PptxPresentation = {
      ...snapshot.root,
      embeddings: newEmbeddings,
      media: newMedia,
    };
    const next = evolveSnapshot(snapshot, root, {
      embeddings: [payload.embeddingPartPath],
      media: mediaDirty,
    });
    return {
      next,
      diff: {
        format: "pptx",
        fromRevision: snapshot.revision,
        toRevision: next.revision,
        changes: [
          {
            kind: "node-updated",
            nodeId: payload.embeddingPartPath,
            path: ["embeddings", payload.embeddingPartPath],
            field: "bytes",
            summary: `~spreadsheet ${payload.embeddingPartPath}`,
          },
        ],
      },
    };
  },
};

// ─── helpers ──────────────────────────────────────────────────────────────

function mintEmbeddingPath(map: ReadonlyMap<string, EmbeddedBinaryPart>): string {
  let n = 1;
  while (map.has(`ppt/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`)) n++;
  return `ppt/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
}

function findPreviewPathForEmbedding(snapshot: PptxSnapshot, embeddingPath: string): string | undefined {
  for (const slide of snapshot.root.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "ole-spreadsheet" && shape.embeddingPartPath === embeddingPath) {
        return shape.previewMediaPartPath;
      }
    }
  }
  return undefined;
}

function nextRelId(existing: ReadonlyArray<string>): string {
  let max = 0;
  for (const id of existing) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `rId${max + 1}`;
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  const dir = slash >= 0 ? partPath.slice(0, slash) : "";
  const file = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return `${dir}${dir ? "/" : ""}_rels/${file}.rels`;
}

/**
 * Compute a relationship `Target` attribute relative to the rels
 * file's owning part dir. e.g. for `ppt/slides/_rels/slide1.xml.rels`
 * a target of `ppt/embeddings/foo.xlsx` becomes `../embeddings/foo.xlsx`.
 */
function relativeFromRels(relsPath: string, targetAbsPath: string): string {
  const ownerDir = ownerDirOfRels(relsPath).split("/").filter(Boolean);
  const target = targetAbsPath.split("/").filter(Boolean);
  let i = 0;
  while (i < ownerDir.length && i < target.length - 1 && ownerDir[i] === target[i]) {
    i++;
  }
  const ups = ownerDir.length - i;
  const rest = target.slice(i);
  const parts: string[] = [];
  for (let k = 0; k < ups; k++) parts.push("..");
  for (const r of rest) parts.push(r);
  return parts.join("/");
}

function ownerDirOfRels(relsPath: string): string {
  const idx = relsPath.lastIndexOf("/_rels/");
  if (idx < 0) return "";
  return relsPath.slice(0, idx);
}

interface OleChildOpts {
  readonly previewMediaRelId: string;
  readonly previewPicId: number;
  readonly previewName: string;
  readonly position: { readonly xEmu: number; readonly yEmu: number };
  readonly size: { readonly cxEmu: number; readonly cyEmu: number };
}

/**
 * Synthesise the `<p:oleObj>` child subtree PowerPoint expects for an
 * embedded Excel object: a `<p:embed/>` followed by a `<p:pic>` that
 * carries the rendered preview image. The pic is what Office paints
 * before the user activates the embed for editing.
 */
function buildOleObjChildren(opts: OleChildOpts): ReadonlyArray<OpaqueXml> {
  const embed: OpaqueXml = { tag: "p:embed", attrs: {}, rawAttrs: {}, subtree: [] };
  const pic: OpaqueXml = {
    tag: "p:pic",
    attrs: {},
    rawAttrs: {},
    subtree: [
      {
        "p:nvPicPr": [
          {
            "p:cNvPr": [],
            ":@": {
              "@_id": String(opts.previewPicId),
              "@_name": opts.previewName,
            },
          },
          { "p:cNvPicPr": [{ "a:picLocks": [], ":@": { "@_noChangeAspect": "1" } }] },
          { "p:nvPr": [] },
        ],
      },
      {
        "p:blipFill": [
          { "a:blip": [], ":@": { "@_r:embed": opts.previewMediaRelId } },
          { "a:stretch": [{ "a:fillRect": [] }] },
        ],
      },
      {
        "p:spPr": [
          {
            "a:xfrm": [
              {
                "a:off": [],
                ":@": {
                  "@_x": String(opts.position.xEmu),
                  "@_y": String(opts.position.yEmu),
                },
              },
              {
                "a:ext": [],
                ":@": {
                  "@_cx": String(opts.size.cxEmu),
                  "@_cy": String(opts.size.cyEmu),
                },
              },
            ],
          },
          {
            "a:prstGeom": [{ "a:avLst": [] }],
            ":@": { "@_prst": "rect" },
          },
        ],
      },
    ],
  };
  return [embed, pic];
}
