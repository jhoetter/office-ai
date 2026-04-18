import { type CommandHandler, sha256Hex } from "@officeai/core";
import type {
  ContentTypesSnap,
  MediaPart,
  Picture,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
} from "../model/types.js";
import { evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { InsertImagePayload } from "./payloads.js";

const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const SUPPORTED: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
  "image/webp": "webp",
  "image/svg+xml": "svg",
};

export const insertImageHandler: CommandHandler<InsertImagePayload, PptxSnapshot> = {
  type: "pptx:insert-image",
  apply(snapshot, payload, ctx) {
    if (!payload.mimeType || !(payload.mimeType in SUPPORTED)) {
      throw makeError("invalid-payload", `unsupported MIME type: ${payload.mimeType}`);
    }
    if (payload.width <= 0 || payload.height <= 0) {
      throw makeError("invalid-payload", "width and height must be > 0");
    }
    const dataView = toUint8(payload.data);
    if (dataView.byteLength === 0) {
      throw makeError("invalid-payload", "image data is empty");
    }
    const ext = SUPPORTED[payload.mimeType];

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const slideRelsPath = relsPathFor(slide.partPath);

    // Dedup by SHA-256.
    const sha = sha256Hex(dataView);
    let mediaPartPath = "";
    let mediaWasNew = false;
    for (const m of snapshot.root.media.values()) {
      if (m.sha256 === sha) {
        mediaPartPath = m.partPath;
        break;
      }
    }
    let newMedia: ReadonlyMap<string, MediaPart> = snapshot.root.media;
    let nextMediaPartIndex = snapshot.root.idGen.nextMediaPartIndex;
    if (!mediaPartPath) {
      mediaPartPath = `ppt/media/image${nextMediaPartIndex}.${ext}`;
      const part: MediaPart = {
        partPath: mediaPartPath,
        bytes: dataView,
        sha256: sha,
        contentType: payload.mimeType,
      };
      const m = new Map(snapshot.root.media);
      m.set(mediaPartPath, part);
      newMedia = m;
      nextMediaPartIndex += 1;
      mediaWasNew = true;
    }

    // Find or mint slide rel for this media path.
    const slideRels = snapshot.relationships.get(slideRelsPath);
    const existingEntries = slideRels?.entries ?? [];
    let mediaRelId = "";
    for (const r of existingEntries) {
      if (r.type !== REL_TYPE_IMAGE) continue;
      const abs = resolveRelTarget(slideRelsPath, r.target);
      if (abs === mediaPartPath) {
        mediaRelId = r.id;
        break;
      }
    }
    let newRelationships: ReadonlyMap<string, RelationshipsSnap> = snapshot.relationships;
    let dirtyRelsPaths: string[] = [];
    if (!mediaRelId) {
      mediaRelId = nextRelId(existingEntries.map((e) => e.id));
      const newSlideEntries = [
        ...existingEntries,
        {
          id: mediaRelId,
          type: REL_TYPE_IMAGE,
          target: relativeFromRels(slideRelsPath, mediaPartPath),
        },
      ];
      const m = new Map(snapshot.relationships);
      m.set(slideRelsPath, { relsPath: slideRelsPath, entries: newSlideEntries });
      newRelationships = m;
      dirtyRelsPaths = [slideRelsPath];
    }

    // Update content types: add Default for the extension if not present.
    let newContentTypes: ContentTypesSnap = snapshot.contentTypes;
    let contentTypesDirty = false;
    if (mediaWasNew) {
      const hasDefault = snapshot.contentTypes.defaults.some(
        (d) => d.extension.toLowerCase() === ext.toLowerCase()
      );
      if (!hasDefault) {
        newContentTypes = {
          ...snapshot.contentTypes,
          defaults: [...snapshot.contentTypes.defaults, { extension: ext, contentType: payload.mimeType }],
        };
        contentTypesDirty = true;
      }
    }

    // Mint cNvPrId in slide scope.
    const cNvPrId = maxCNvPrId(slide.shapes) + 1;

    const pic: Picture = {
      kind: "pic",
      id: ctx.mintNodeId(),
      cNvPrId,
      name: payload.name ?? `Picture ${cNvPrId}`,
      position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
      size: { cxEmu: Math.round(payload.width), cyEmu: Math.round(payload.height) },
      mediaRelId,
      mediaPartPath,
      nvPicPrTail: defaultPicNvTail(payload.altText),
      blipFillTail: defaultBlipFillTail(),
      spPrTail: defaultPicSpPrTail(),
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, pic] };
    const newSlides = [...snapshot.root.slides];
    newSlides[sIdx] = newSlide;

    const root: PptxPresentation = {
      ...snapshot.root,
      slides: newSlides,
      media: newMedia,
      idGen: { ...snapshot.root.idGen, nextMediaPartIndex },
    };

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        slides: [slide.partPath],
        media: mediaWasNew ? [mediaPartPath] : [],
        relationships: dirtyRelsPaths,
        contentTypes: contentTypesDirty,
      },
      {
        relationships: newRelationships,
        contentTypes: newContentTypes,
      }
    );

    const changes = [
      {
        kind: "node-inserted" as const,
        nodeId: pic.id,
        path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1] as ReadonlyArray<string | number>,
        summary: "picture",
      },
    ];
    if (mediaWasNew) {
      changes.push({
        kind: "node-inserted" as const,
        nodeId: pic.id,
        path: [mediaPartPath] as ReadonlyArray<string | number>,
        summary: `media ${payload.mimeType}`,
      });
    }
    return {
      next,
      diff: {
        format: "pptx",
        fromRevision: snapshot.revision,
        toRevision: next.revision,
        changes: mediaWasNew
          ? [
              changes[0],
              {
                kind: "part-added",
                path: [mediaPartPath],
                summary: `media ${payload.mimeType}`,
              },
            ]
          : [changes[0]],
      },
    };
  },
};

function toUint8(data: Uint8Array | ArrayBuffer): Uint8Array {
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data);
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

function resolveRelTarget(relsPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  // Owner is the rels file's "owner" — i.e. the part it describes.
  // For ppt/slides/_rels/slide1.xml.rels owner part is ppt/slides/slide1.xml,
  // so its directory is ppt/slides.
  const ownerDir = ownerDirOfRels(relsPath);
  const segs = (ownerDir ? `${ownerDir}/${target}` : target).split("/");
  const out: string[] = [];
  for (const s of segs) {
    if (!s || s === ".") continue;
    if (s === "..") {
      out.pop();
      continue;
    }
    out.push(s);
  }
  return out.join("/");
}

function relativeFromRels(relsPath: string, targetPath: string): string {
  const ownerDir = ownerDirOfRels(relsPath);
  const ownerSegs = ownerDir.split("/").filter((s) => s.length > 0);
  const targetSegs = targetPath.split("/");
  let i = 0;
  while (i < ownerSegs.length && i < targetSegs.length - 1 && ownerSegs[i] === targetSegs[i]) {
    i++;
  }
  const ups = ownerSegs.length - i;
  const downs = targetSegs.slice(i);
  return [...Array(ups).fill(".."), ...downs].join("/");
}

function ownerDirOfRels(relsPath: string): string {
  const m = /^(.*?)_rels\/[^/]+\.rels$/.exec(relsPath);
  if (!m) return "";
  return (m[1] ?? "").replace(/\/$/, "");
}

function defaultPicNvTail(altText: string | undefined): import("../model/types.js").OpaqueXml[] {
  return [
    {
      tag: "p:cNvPr",
      attrs: { id: "0", name: "" },
      rawAttrs: {
        "@_id": "0",
        "@_name": "",
        ...(altText ? { "@_descr": altText } : {}),
      },
      subtree: [],
    },
    {
      tag: "p:cNvPicPr",
      attrs: {},
      rawAttrs: {},
      subtree: [{ "a:picLocks": [], ":@": { "@_noChangeAspect": "1" } }],
    },
    { tag: "p:nvPr", attrs: {}, rawAttrs: {}, subtree: [] },
  ];
}

function defaultBlipFillTail(): import("../model/types.js").OpaqueXml[] {
  return [
    { tag: "a:blip", attrs: {}, rawAttrs: {}, subtree: [] },
    {
      tag: "a:stretch",
      attrs: {},
      rawAttrs: {},
      subtree: [{ "a:fillRect": [] }],
    },
  ];
}

function defaultPicSpPrTail(): import("../model/types.js").OpaqueXml[] {
  return [
    {
      tag: "a:prstGeom",
      attrs: { prst: "rect" },
      rawAttrs: { "@_prst": "rect" },
      subtree: [{ "a:avLst": [] }],
    },
  ];
}
