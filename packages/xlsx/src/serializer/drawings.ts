import { ooxml } from "@officeai/core";
import { pxToEmu, type ImageBlob, type SheetImage } from "../model/index.js";
import { DRAWING_REL_TYPE, IMAGE_REL_TYPE } from "../parser/drawings.js";

export const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";

const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELS_OFFICE_DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Serialize a sheet's `images` array into an `xdr:wsDr` XML payload.
 *
 * Anchor flavour is fixed at `<xdr:twoCellAnchor editAs="oneCell">` —
 * v1 always authors that combination. The `to` endpoint is a
 * conservative approximation (`from + 1` cell) since Excel recomputes
 * the anchored geometry from the picture's `<a:ext cx cy>` payload at
 * render time when `editAs="oneCell"`. Existing files keep whatever
 * `to` they had on disk because we only re-author drawings on edit.
 *
 * `embedRidByMediaPath` maps each unique media part path to its rels
 * id WITHIN the drawing rels graph, so two image instances sharing
 * the same media part collapse to a single `<a:blip r:embed>`.
 */
export function serializeDrawingPart(
  images: ReadonlyArray<SheetImage>,
  embedRidByMediaPath: ReadonlyMap<string, string>
): string {
  const parts: string[] = [];
  let nextPicId = 1;
  for (const img of images) {
    const rid = embedRidByMediaPath.get(img.mediaRef);
    if (!rid) continue;
    const id = nextPicId++;
    const name = escapeXmlAttr(img.name ?? `Picture ${id}`);
    const altText = img.altText ? ` descr="${escapeXmlAttr(img.altText)}"` : "";

    const fromOff = {
      cx: pxToEmu(img.anchor.fromOffsetXPx),
      cy: pxToEmu(img.anchor.fromOffsetYPx),
    };
    const ext = {
      cx: pxToEmu(img.anchor.widthPx),
      cy: pxToEmu(img.anchor.heightPx),
    };

    // For oneCell-style positioning the (toRow, toCol) endpoint is
    // recomputed by Excel from `from` + the picture's `a:ext`. We
    // emit a placeholder `to` one cell down/right of `from`; Excel
    // overwrites it on the next save anyway.
    const toRow = img.anchor.fromRow + 1;
    const toCol = img.anchor.fromCol + 1;

    parts.push(
      `<xdr:twoCellAnchor editAs="${escapeXmlAttr(img.anchor.editAs)}">` +
        `<xdr:from>` +
        `<xdr:col>${img.anchor.fromCol}</xdr:col>` +
        `<xdr:colOff>${fromOff.cx}</xdr:colOff>` +
        `<xdr:row>${img.anchor.fromRow}</xdr:row>` +
        `<xdr:rowOff>${fromOff.cy}</xdr:rowOff>` +
        `</xdr:from>` +
        `<xdr:to>` +
        `<xdr:col>${toCol}</xdr:col>` +
        `<xdr:colOff>0</xdr:colOff>` +
        `<xdr:row>${toRow}</xdr:row>` +
        `<xdr:rowOff>0</xdr:rowOff>` +
        `</xdr:to>` +
        `<xdr:pic>` +
        `<xdr:nvPicPr>` +
        `<xdr:cNvPr id="${id}" name="${name}"${altText}/>` +
        `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>` +
        `</xdr:nvPicPr>` +
        `<xdr:blipFill>` +
        `<a:blip xmlns:r="${RELS_OFFICE_DOC_NS}" r:embed="${escapeXmlAttr(rid)}"/>` +
        `<a:stretch><a:fillRect/></a:stretch>` +
        `</xdr:blipFill>` +
        `<xdr:spPr>` +
        `<a:xfrm>` +
        `<a:off x="0" y="0"/>` +
        `<a:ext cx="${ext.cx}" cy="${ext.cy}"/>` +
        `</a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</xdr:spPr>` +
        `</xdr:pic>` +
        `<xdr:clientData/>` +
        `</xdr:twoCellAnchor>`
    );
  }

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING_NS}" xmlns:a="${DRAWINGML_NS}">` +
    parts.join("") +
    `</xdr:wsDr>`
  );
}

/**
 * Build the `xl/drawings/_rels/drawingN.xml.rels` graph for a freshly
 * authored drawing part. Each unique media path collapses to exactly
 * one rel, named `rId1`, `rId2`, …, returned as a path → rId map.
 */
export function buildDrawingRels(
  drawingPartPath: string,
  images: ReadonlyArray<SheetImage>,
  workbookImages: ReadonlyMap<string, ImageBlob>
): { graph: ooxml.RelationshipGraph; embedRidByMediaPath: Map<string, string> } {
  const relsPath = ooxml.RelationshipGraph.relsPathFor(drawingPartPath);
  const graph = new ooxml.RelationshipGraph(relsPath, []);
  const embedRidByMediaPath = new Map<string, string>();
  for (const img of images) {
    if (embedRidByMediaPath.has(img.mediaRef)) continue;
    if (!workbookImages.has(img.mediaRef)) continue;
    const target = relsRelativeTarget(drawingPartPath, img.mediaRef);
    const r = graph.add({ type: IMAGE_REL_TYPE, target });
    embedRidByMediaPath.set(img.mediaRef, r.id);
  }
  return { graph, embedRidByMediaPath };
}

/**
 * Find an unused `xl/drawings/drawingN.xml` slot in the container.
 */
export function mintDrawingPartPath(container: ooxml.OoxmlContainer): string {
  let i = 1;
  while (container.has(`xl/drawings/drawing${i}.xml`)) i++;
  return `xl/drawings/drawing${i}.xml`;
}

/**
 * Inject (or replace) the `<drawing r:id="..."/>` reference in a
 * worksheet XML string. Excel requires the element to live inside
 * `<worksheet>` (after `<mergeCells>` / `<autoFilter>` is fine). When
 * `rid` is `null` any pre-existing `<drawing>` is removed.
 */
export function injectDrawingRef(xml: string, rid: string | null): string {
  const next = xml.replace(/<drawing\b[^/>]*\/>/g, "");
  if (!rid) return next;
  const block = `<drawing r:id="${escapeXmlAttr(rid)}"/>`;
  const closeIdx = next.lastIndexOf("</worksheet>");
  if (closeIdx === -1) return next + block;
  return next.slice(0, closeIdx) + block + next.slice(closeIdx);
}

/**
 * Add/remove the drawing relationship on a sheet's rels graph. Other
 * relationship types (comments, hyperlinks, …) are preserved verbatim.
 *
 * Returns the rId Excel will see in the worksheet's `<drawing>` ref.
 */
export function upsertSheetDrawingRel(
  graph: ooxml.RelationshipGraph,
  sheetPartPath: string,
  drawingPartPath: string | null
): string | null {
  for (const r of [...graph.relationships]) {
    if (r.type === DRAWING_REL_TYPE) graph.remove(r.id);
  }
  if (!drawingPartPath) return null;
  const target = relsRelativeTarget(sheetPartPath, drawingPartPath);
  const r = graph.add({ type: DRAWING_REL_TYPE, target });
  return r.id;
}

function relsRelativeTarget(ownerPartPath: string, targetPath: string): string {
  const ownerDir = ownerPartPath.includes("/") ? ownerPartPath.slice(0, ownerPartPath.lastIndexOf("/")) : "";
  const ownerSegs = ownerDir ? ownerDir.split("/") : [];
  const targetSegs = targetPath.split("/");
  let common = 0;
  while (
    common < ownerSegs.length &&
    common < targetSegs.length &&
    ownerSegs[common] === targetSegs[common]
  ) {
    common++;
  }
  const ups = ownerSegs.length - common;
  const rest = targetSegs.slice(common).join("/");
  return ups > 0 ? `${"../".repeat(ups)}${rest}` : rest;
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
