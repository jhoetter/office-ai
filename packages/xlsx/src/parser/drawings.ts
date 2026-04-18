import { ooxml, sha256Hex, type IdMinter } from "@officeai/core";
import {
  contentTypeForExtension,
  emuToPx,
  type ImageBlob,
  type ImageContentType,
  type SheetImage,
} from "../model/index.js";
import { XlsxParseError } from "./errors.js";
import { resolveTargetPath } from "./parse.js";

/**
 * Relationship type for sheet → drawing parts.
 */
export const DRAWING_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

/**
 * Relationship type for drawing → media (raster image).
 */
export const IMAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * Result of resolving a sheet's `xdr:wsDr` payload.
 */
export interface ResolvedDrawings {
  /** `xl/drawings/drawingN.xml` part path, when one is referenced. */
  readonly drawingPartPath?: string;
  /** Sheet-level images in their declared (z-order) sequence. */
  readonly images: ReadonlyArray<SheetImage>;
  /** Newly-discovered media blobs to splice into `XlsxWorkbook.images`. */
  readonly mediaBlobs: ReadonlyArray<ImageBlob>;
}

/**
 * Walk the rels for `sheetPartPath`, find the (at most one)
 * `xdr:wsDr` payload, and decode every embedded image into a
 * {@link SheetImage} + the corresponding {@link ImageBlob}.
 *
 * Anchors we don't fully understand (chart frames, vector shapes)
 * are skipped silently — they remain in `opaqueParts` and round-trip
 * unchanged because we don't claim ownership of the drawing part
 * unless every anchor was modeled.
 *
 * Returns `images: []` when the sheet has no drawings rel. The caller
 * is responsible for adding `drawingPartPath` (and the media paths)
 * to the modeled-paths set so they aren't double-stored as opaque.
 */
export function resolveDrawings(
  container: ooxml.OoxmlContainer,
  sheetPartPath: string,
  mintNodeId: IdMinter,
  existingMedia: ReadonlyMap<string, ImageBlob>
): ResolvedDrawings {
  const sheetRels = ooxml.RelationshipGraph.loadFor(container, sheetPartPath);
  const drawingRel = sheetRels.relationships.find((r) => r.type === DRAWING_REL_TYPE);
  if (!drawingRel) return { images: [], mediaBlobs: [] };

  const drawingPartPath = resolveTargetPath(sheetPartPath, drawingRel.target);
  if (!container.has(drawingPartPath)) return { images: [], mediaBlobs: [] };

  const drawingRels = ooxml.RelationshipGraph.loadFor(container, drawingPartPath);
  const xml = container.readText(drawingPartPath);
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new XlsxParseError("invalid-xml", `Failed to parse drawing part`, {
      partPath: drawingPartPath,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) return { drawingPartPath, images: [], mediaBlobs: [] };

  const root = (tree as unknown[])
    .map((n) => ooxml.asElement(n))
    .find((el): el is ooxml.XmlElement => el !== null && el.tag === "xdr:wsDr");
  if (!root) return { drawingPartPath, images: [], mediaBlobs: [] };

  const images: SheetImage[] = [];
  const mediaBlobs: ImageBlob[] = [];
  const seenMedia = new Map<string, ImageBlob>();

  for (const child of root.children) {
    const el = ooxml.asElement(child);
    if (!el) continue;

    let editAs: SheetImage["anchor"]["editAs"] = "oneCell";
    if (el.tag === "xdr:twoCellAnchor") {
      editAs = (el.attrs.editAs as SheetImage["anchor"]["editAs"]) ?? "twoCell";
    } else if (el.tag === "xdr:oneCellAnchor") {
      editAs = "oneCell";
    } else if (el.tag === "xdr:absoluteAnchor") {
      editAs = "absolute";
    } else {
      continue;
    }

    const decoded = decodeAnchor(el, editAs);
    if (!decoded) continue;

    const pic = findDescendant(el.children, "xdr:pic");
    if (!pic) continue;
    const blipFill = ooxml.findChild(pic.children, "xdr:blipFill");
    const blip = blipFill ? ooxml.findChild(blipFill.children, "a:blip") : null;
    const embedId = blip?.attrs["r:embed"];
    if (!embedId) continue;

    const mediaRel = drawingRels.byId(embedId);
    if (!mediaRel || mediaRel.type !== IMAGE_REL_TYPE) continue;
    const mediaPath = resolveTargetPath(drawingPartPath, mediaRel.target);
    if (!container.has(mediaPath)) continue;

    const blob = ensureMediaBlob(container, mediaPath, existingMedia, seenMedia);
    if (!blob) continue;
    if (!existingMedia.has(mediaPath) && !mediaBlobs.some((b) => b.partPath === mediaPath)) {
      mediaBlobs.push(blob);
    }

    const nvPicPr = ooxml.findChild(pic.children, "xdr:nvPicPr");
    const cNvPr = nvPicPr ? ooxml.findChild(nvPicPr.children, "xdr:cNvPr") : null;
    const name = cNvPr?.attrs.name;
    const altText = cNvPr?.attrs.descr;

    images.push({
      id: mintNodeId(),
      anchor: { ...decoded, editAs },
      mediaRef: mediaPath,
      ...(name ? { name } : {}),
      ...(altText ? { altText } : {}),
    });
  }

  return { drawingPartPath, images, mediaBlobs };
}

/** Decode the from/ext payload of any of the three xdr anchor flavours into pixels. */
function decodeAnchor(
  anchor: ooxml.XmlElement,
  _editAs: SheetImage["anchor"]["editAs"]
): Omit<SheetImage["anchor"], "editAs"> | null {
  const from = ooxml.findChild(anchor.children, "xdr:from");
  if (!from) return null;
  const fromCol = readIntChild(from, "xdr:col");
  const fromRow = readIntChild(from, "xdr:row");
  const fromColOffEmu = readIntChild(from, "xdr:colOff") ?? 0;
  const fromRowOffEmu = readIntChild(from, "xdr:rowOff") ?? 0;
  if (fromCol === undefined || fromRow === undefined) return null;

  let widthPx: number;
  let heightPx: number;

  if (anchor.tag === "xdr:twoCellAnchor") {
    const to = ooxml.findChild(anchor.children, "xdr:to");
    if (!to) return null;
    const toCol = readIntChild(to, "xdr:col");
    const toRow = readIntChild(to, "xdr:row");
    if (toCol === undefined || toRow === undefined) return null;
    // Width/height live on the picture's `<a:ext cx cy>` (inside
    // `xdr:pic > xdr:spPr > a:xfrm`). Cell anchoring is preserved
    // exactly via from-row/col + offset; the (toRow, toCol) endpoint
    // is rederived at serialize time from from + size.
    const ext = findDescendant(anchor.children, "a:ext");
    widthPx = ext ? emuToPx(Number(ext.attrs.cx ?? 0)) : 0;
    heightPx = ext ? emuToPx(Number(ext.attrs.cy ?? 0)) : 0;
    if (widthPx === 0 || heightPx === 0) {
      // Last-ditch approximation when Excel omits the ext (rare).
      widthPx = Math.max(0, (toCol - fromCol) * 64) || 96;
      heightPx = Math.max(0, (toRow - fromRow) * 20) || 96;
    }
  } else {
    const ext = ooxml.findChild(anchor.children, "xdr:ext");
    widthPx = ext ? emuToPx(Number(ext.attrs.cx ?? 0)) : 96;
    heightPx = ext ? emuToPx(Number(ext.attrs.cy ?? 0)) : 96;
  }

  return {
    fromRow,
    fromCol,
    fromOffsetXPx: emuToPx(fromColOffEmu),
    fromOffsetYPx: emuToPx(fromRowOffEmu),
    widthPx,
    heightPx,
  };
}

function readIntChild(parent: ooxml.XmlElement, tag: string): number | undefined {
  const el = ooxml.findChild(parent.children, tag);
  if (!el) return undefined;
  const text = ooxml.getTextContent(el.entry).trim();
  if (!text) return undefined;
  const n = Number(text);
  return Number.isFinite(n) ? n : undefined;
}

function findDescendant(siblings: ReadonlyArray<unknown>, tag: string): ooxml.XmlElement | null {
  for (const s of siblings) {
    const el = ooxml.asElement(s);
    if (!el) continue;
    if (el.tag === tag) return el;
    const inner = findDescendant(el.children, tag);
    if (inner) return inner;
  }
  return null;
}

function ensureMediaBlob(
  container: ooxml.OoxmlContainer,
  partPath: string,
  existingMedia: ReadonlyMap<string, ImageBlob>,
  seen: Map<string, ImageBlob>
): ImageBlob | null {
  const cached = existingMedia.get(partPath) ?? seen.get(partPath);
  if (cached) return cached;
  const bytes = container.readBytes(partPath);
  const ext = partPath.slice(partPath.lastIndexOf(".") + 1);
  const contentType: ImageContentType | undefined = contentTypeForExtension(ext);
  if (!contentType) return null;
  const blob: ImageBlob = {
    partPath,
    bytes,
    contentType,
    hash: sha256Hex(bytes),
  };
  seen.set(partPath, blob);
  return blob;
}
