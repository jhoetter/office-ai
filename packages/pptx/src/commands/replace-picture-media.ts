import { type CommandHandler, sha256Hex } from "@officeai/core";
import type {
  ContentTypesSnap,
  MediaPart,
  OpaqueXml,
  Picture,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
} from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
} from "./helpers.js";
import type { ReplacePictureMediaPayload } from "./payloads.js";

/**
 * D9 — replace the bitmap behind an existing `Picture` while
 * preserving everything else (position, size, alt-text, spPrTail
 * styling). Mirrors `insert-image`'s media-mint dance so we don't
 * duplicate the dedup/relationship/contentType bookkeeping logic;
 * the only structural difference is that we mutate an existing
 * `Picture.mediaRelId` instead of inserting a new shape.
 */

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

export const replacePictureMediaHandler: CommandHandler<ReplacePictureMediaPayload, PptxSnapshot> = {
  type: "pptx:replace-picture-media",
  apply(snapshot, payload) {
    if (!payload.mimeType || !(payload.mimeType in SUPPORTED)) {
      throw makeError("invalid-payload", `unsupported MIME type: ${payload.mimeType}`);
    }
    const data = toUint8(payload.data);
    if (data.byteLength === 0) {
      throw makeError("invalid-payload", "image data is empty");
    }
    const ext = SUPPORTED[payload.mimeType];

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const found = findShapeInSlide(slide, payload.shapeId);
    if (found.shape.kind !== "pic") {
      throw makeError(
        "invalid-target",
        `shape ${payload.shapeId} is not a picture (kind=${found.shape.kind})`
      );
    }
    const pic = found.shape;

    const slideRelsPath = relsPathFor(slide.partPath);

    const sha = sha256Hex(data);
    let mediaPartPath = "";
    for (const m of snapshot.root.media.values()) {
      if (m.sha256 === sha) {
        mediaPartPath = m.partPath;
        break;
      }
    }
    let newMedia: ReadonlyMap<string, MediaPart> = snapshot.root.media;
    let nextMediaPartIndex = snapshot.root.idGen.nextMediaPartIndex;
    let mediaWasNew = false;
    if (!mediaPartPath) {
      mediaPartPath = `ppt/media/image${nextMediaPartIndex}.${ext}`;
      const part: MediaPart = {
        partPath: mediaPartPath,
        bytes: data,
        sha256: sha,
        contentType: payload.mimeType,
      };
      const m = new Map(snapshot.root.media);
      m.set(mediaPartPath, part);
      newMedia = m;
      nextMediaPartIndex += 1;
      mediaWasNew = true;
    }

    // Locate (or mint) the slide-rels entry pointing at the new media.
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

    // Update the alt-text in nvPicPrTail if the caller provided one.
    // Preserving every other entry verbatim keeps round-trip clean
    // for shape locks, hyperlink-on-pic, and similar opaque siblings.
    const nextNvPicPrTail =
      payload.altText !== undefined ? withAltText(pic.nvPicPrTail, payload.altText) : pic.nvPicPrTail;

    const updated: Picture = {
      ...pic,
      mediaRelId,
      mediaPartPath,
      nvPicPrTail: nextNvPicPrTail,
    };

    const nextShapes = replaceShape(slide.shapes, found.path, updated);
    const newSlide: Slide = { ...slide, shapes: nextShapes };
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

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: pic.id,
        path: ["slides", sIdx, "shapes", ...found.path],
        field: "media",
        summary: `replace-picture-media:${payload.mimeType}`,
      }),
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

/**
 * Update or inject the `@_descr` attribute on the picture's
 * `<p:cNvPr>` opaque entry. The DOCX/XLSX/PPTX schemas all use the
 * same `descr` attribute for accessible alt-text, and PowerPoint
 * happily reads back what we emit.
 */
function withAltText(nvPicPrTail: ReadonlyArray<OpaqueXml>, altText: string): ReadonlyArray<OpaqueXml> {
  return nvPicPrTail.map((entry) => {
    if (entry.tag !== "p:cNvPr") return entry;
    return {
      ...entry,
      attrs: { ...entry.attrs, descr: altText },
      rawAttrs: { ...entry.rawAttrs, "@_descr": altText },
    };
  });
}
