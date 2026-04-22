import { type CommandHandler, type DiffChange, sha256Hex } from "@officeai/core";
import type {
  ContentTypesSnap,
  MediaPart,
  MediaShape,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
} from "../model/types.js";
import { buildMediaPicRaw } from "../serializer/media.js";
import { evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { InsertMediaPayload } from "./payloads.js";

const REL_TYPE_VIDEO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
const REL_TYPE_AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const VIDEO_MIME: Record<string, string> = {
  "video/mp4": "mp4",
  "video/webm": "webm",
  "video/quicktime": "mov",
};

const AUDIO_MIME: Record<string, string> = {
  "audio/mpeg": "mp3",
  "audio/mp4": "m4a",
  "audio/wav": "wav",
  "audio/ogg": "ogg",
};

// 1×1 transparent PNG. Used as the default poster image so the
// `<p:blipFill>` reference always resolves; the renderer paints its
// own typed placeholder over it during edit mode.
const POSTER_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

/**
 * Phase-1 insert-media handler. Mirrors the structure of
 * `insertImageHandler` (SHA-256-dedup, slide-rel minting, content-types
 * upkeep) but lays down a `MediaShape` whose `<p:pic>` raw blob
 * carries the `<a:videoFile>` / `<a:audioFile>` reference. Always
 * registers a transparent 1×1 PNG as the poster so the `<p:blipFill>`
 * resolves cleanly when the file is opened in PowerPoint; the editor's
 * own renderer paints a media-typed placeholder over the empty image
 * during edit mode.
 */
export const insertMediaHandler: CommandHandler<InsertMediaPayload, PptxSnapshot> = {
  type: "pptx:insert-media",
  apply(snapshot, payload, ctx) {
    const supported = payload.mediaType === "video" ? VIDEO_MIME : AUDIO_MIME;
    if (!payload.contentType || !(payload.contentType in supported)) {
      throw makeError(
        "invalid-payload",
        `unsupported MIME type for ${payload.mediaType}: ${payload.contentType}`
      );
    }
    if (!payload.position || !payload.size) {
      throw makeError("invalid-payload", "position and size are required");
    }
    if (payload.size.cxEmu <= 0 || payload.size.cyEmu <= 0) {
      throw makeError("invalid-payload", "size.cxEmu and size.cyEmu must be > 0");
    }
    const dataView = toUint8(payload.bytes);
    if (dataView.byteLength === 0) {
      throw makeError("invalid-payload", "media bytes are empty");
    }
    const ext = supported[payload.contentType];
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const slideRelsPath = relsPathFor(slide.partPath);

    // ── Media binary part (SHA-256 dedup, mints `ppt/media/mediaN.<ext>`).
    const sha = sha256Hex(dataView);
    let mediaPath = "";
    let mediaWasNew = false;
    for (const m of snapshot.root.media.values()) {
      if (m.sha256 === sha) {
        mediaPath = m.partPath;
        break;
      }
    }
    let nextMediaPartIndex = snapshot.root.idGen.nextMediaPartIndex;
    let newMedia: ReadonlyMap<string, MediaPart> = snapshot.root.media;
    if (!mediaPath) {
      mediaPath = mintMediaPath(snapshot.root.media, ext);
      const part: MediaPart = {
        partPath: mediaPath,
        bytes: dataView,
        sha256: sha,
        contentType: payload.contentType,
      };
      const m = new Map(snapshot.root.media);
      m.set(mediaPath, part);
      newMedia = m;
      mediaWasNew = true;
    }

    // ── Poster image (transparent 1×1 PNG). Reused across every
    // media insert via SHA-256 dedup.
    const posterSha = sha256Hex(POSTER_PNG);
    let posterPath = "";
    let posterWasNew = false;
    for (const m of newMedia.values()) {
      if (m.sha256 === posterSha) {
        posterPath = m.partPath;
        break;
      }
    }
    if (!posterPath) {
      posterPath = `ppt/media/image${nextMediaPartIndex}.png`;
      nextMediaPartIndex += 1;
      const part: MediaPart = {
        partPath: posterPath,
        bytes: POSTER_PNG,
        sha256: posterSha,
        contentType: "image/png",
      };
      const m = new Map(newMedia);
      m.set(posterPath, part);
      newMedia = m;
      posterWasNew = true;
    }

    // ── Slide rels: one entry per (mediaPath, posterPath), reused
    // when an existing rel already targets the same part.
    const slideRels = snapshot.relationships.get(slideRelsPath);
    const existingEntries = slideRels?.entries ?? [];
    const mediaRelType = payload.mediaType === "video" ? REL_TYPE_VIDEO : REL_TYPE_AUDIO;

    let mediaRelId = findRelId(existingEntries, mediaRelType, slideRelsPath, mediaPath);
    let posterRelId = findRelId(existingEntries, REL_TYPE_IMAGE, slideRelsPath, posterPath);

    const newEntries = [...existingEntries];
    const taken = new Set(existingEntries.map((e) => e.id));
    let relsDirty = false;
    if (!mediaRelId) {
      mediaRelId = nextRelId(taken);
      taken.add(mediaRelId);
      newEntries.push({
        id: mediaRelId,
        type: mediaRelType,
        target: relativeFromRels(slideRelsPath, mediaPath),
      });
      relsDirty = true;
    }
    if (!posterRelId) {
      posterRelId = nextRelId(taken);
      taken.add(posterRelId);
      newEntries.push({
        id: posterRelId,
        type: REL_TYPE_IMAGE,
        target: relativeFromRels(slideRelsPath, posterPath),
      });
      relsDirty = true;
    }
    let newRelationships: ReadonlyMap<string, RelationshipsSnap> = snapshot.relationships;
    if (relsDirty) {
      const m = new Map(snapshot.relationships);
      m.set(slideRelsPath, { relsPath: slideRelsPath, entries: newEntries });
      newRelationships = m;
    }

    // ── Content types: ensure Default entries for each new extension.
    let newContentTypes: ContentTypesSnap = snapshot.contentTypes;
    let contentTypesDirty = false;
    if (mediaWasNew) {
      const updated = ensureDefault(newContentTypes, ext, payload.contentType);
      if (updated !== newContentTypes) {
        newContentTypes = updated;
        contentTypesDirty = true;
      }
    }
    if (posterWasNew) {
      const updated = ensureDefault(newContentTypes, "png", "image/png");
      if (updated !== newContentTypes) {
        newContentTypes = updated;
        contentTypesDirty = true;
      }
    }

    // ── Mint typed shape.
    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const name = payload.name ?? `${payload.mediaType === "video" ? "Video" : "Audio"} ${cNvPrId}`;
    const position = {
      xEmu: Math.round(payload.position.xEmu),
      yEmu: Math.round(payload.position.yEmu),
    };
    const size = {
      cxEmu: Math.round(payload.size.cxEmu),
      cyEmu: Math.round(payload.size.cyEmu),
    };
    const media: MediaShape = {
      kind: "media",
      id: ctx.mintNodeId(),
      cNvPrId,
      name,
      position,
      size,
      mediaType: payload.mediaType,
      mediaRelId,
      mediaPath,
      posterRelId,
      posterPath,
      raw: buildMediaPicRaw({
        cNvPrId,
        name,
        mediaType: payload.mediaType,
        mediaContentType: payload.contentType,
        mediaRelId,
        posterRelId,
        position,
        size,
      }),
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, media] };
    const newSlides = [...snapshot.root.slides];
    newSlides[sIdx] = newSlide;

    const root: PptxPresentation = {
      ...snapshot.root,
      slides: newSlides,
      media: newMedia,
      idGen: { ...snapshot.root.idGen, nextMediaPartIndex },
    };

    const dirtyMedia: string[] = [];
    if (mediaWasNew) dirtyMedia.push(mediaPath);
    if (posterWasNew) dirtyMedia.push(posterPath);

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        slides: [slide.partPath],
        media: dirtyMedia,
        relationships: relsDirty ? [slideRelsPath] : [],
        contentTypes: contentTypesDirty,
      },
      {
        relationships: newRelationships,
        contentTypes: newContentTypes,
      }
    );

    const changes: DiffChange[] = [
      {
        kind: "node-inserted",
        nodeId: media.id,
        path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1],
        summary: payload.mediaType,
      },
    ];
    if (mediaWasNew) {
      changes.push({ kind: "part-added", path: [mediaPath], summary: `media ${payload.contentType}` });
    }
    return {
      next,
      diff: {
        format: "pptx",
        fromRevision: snapshot.revision,
        toRevision: next.revision,
        changes,
      },
    };
  },
};

function toUint8(data: Uint8Array | ArrayBuffer): Uint8Array {
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data);
}

function mintMediaPath(media: ReadonlyMap<string, MediaPart>, ext: string): string {
  let n = 1;
  while (media.has(`ppt/media/media${n}.${ext}`)) n++;
  return `ppt/media/media${n}.${ext}`;
}

function ensureDefault(ct: ContentTypesSnap, extension: string, contentType: string): ContentTypesSnap {
  const lower = extension.toLowerCase();
  if (ct.defaults.some((d) => d.extension.toLowerCase() === lower)) return ct;
  return { ...ct, defaults: [...ct.defaults, { extension, contentType }] };
}

function findRelId(
  entries: ReadonlyArray<{ id: string; type: string; target: string }>,
  type: string,
  relsPath: string,
  absPath: string
): string {
  for (const r of entries) {
    if (r.type !== type) continue;
    if (resolveRelTarget(relsPath, r.target) === absPath) return r.id;
  }
  return "";
}

function nextRelId(taken: ReadonlySet<string>): string {
  let max = 0;
  for (const id of taken) {
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
