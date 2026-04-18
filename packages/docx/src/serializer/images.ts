import type { InlineImageDrawing, OpaqueXml } from "../model/types.js";
import { opaqueToEntry } from "../parser/xml-helpers.js";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/**
 * Serialize an `InlineImageDrawing` back to a fast-xml-parser preserveOrder
 * entry suitable for embedding inside a `<w:r>`.
 *
 * Byte-preservation contract: callers (the run-child serializer) MUST
 * check `leaf.raw` BEFORE calling this function. If `raw` is set, they
 * re-emit the cached subtree via `opaqueToEntry(leaf.raw)` directly. This
 * function only runs for leaves whose typed model has been touched and
 * must be regenerated (currently: leaves produced by
 * `docx:insert-image`).
 *
 * The emitted shape mirrors what Word writes: a `<w:drawing>` containing a
 * `<wp:inline>` whose graphic data uri is the drawingml/picture URI, with
 * a single `<pic:pic>` referencing the embedded media via `r:embed`.
 *
 * Namespaces — `xmlns:wp` is declared locally on `<wp:inline>` (and
 * `xmlns:a` / `xmlns:pic` on `<a:graphic>` / `<pic:pic>` respectively)
 * because the document root attrs (`documentRootAttrs`) capture only the
 * namespaces that were declared by the original author. Synthetic docs
 * created in-app declare just `xmlns:w` / `xmlns:r`, so without a local
 * declaration here Word rejects the file with "namespace prefix wp is not
 * defined" and treats the package as corrupted.
 */
export function serializeInlineImageDrawing(leaf: InlineImageDrawing): unknown {
  const inlineChildren: unknown[] = [];

  inlineChildren.push({
    "wp:extent": [],
    ":@": { "@_cx": String(leaf.cx), "@_cy": String(leaf.cy) },
  });

  const eff = leaf.properties?.effectExtent;
  if (eff) {
    inlineChildren.push({
      "wp:effectExtent": [],
      ":@": {
        "@_l": String(eff.l),
        "@_t": String(eff.t),
        "@_r": String(eff.r),
        "@_b": String(eff.b),
      },
    });
  }

  const docPrAttrs: Record<string, string> = {
    "@_id": String(leaf.docPrId),
    "@_name": leaf.name,
  };
  if (leaf.descr !== undefined) docPrAttrs["@_descr"] = leaf.descr;
  if (leaf.properties?.title !== undefined) docPrAttrs["@_title"] = leaf.properties.title;
  inlineChildren.push({ "wp:docPr": [], ":@": docPrAttrs });

  const blip: Record<string, unknown> = {
    "a:blip": [],
    ":@": { "@_r:embed": leaf.relId },
  };
  const stretch: Record<string, unknown> = {
    "a:stretch": [{ "a:fillRect": [] }],
  };
  const blipFill: Record<string, unknown> = {
    "pic:blipFill": [blip, stretch],
  };

  const xfrm: Record<string, unknown> = {
    "a:xfrm": [
      { "a:off": [], ":@": { "@_x": "0", "@_y": "0" } },
      { "a:ext": [], ":@": { "@_cx": String(leaf.cx), "@_cy": String(leaf.cy) } },
    ],
  };
  const prstGeom: Record<string, unknown> = {
    "a:prstGeom": [{ "a:avLst": [] }],
    ":@": { "@_prst": "rect" },
  };
  const spPr: Record<string, unknown> = {
    "pic:spPr": [xfrm, prstGeom],
  };

  const nvPicPr: Record<string, unknown> = {
    "pic:nvPicPr": [
      {
        "pic:cNvPr": [],
        ":@": { "@_id": "0", "@_name": leaf.name, ...(leaf.descr ? { "@_descr": leaf.descr } : {}) },
      },
      { "pic:cNvPicPr": [] },
    ],
  };

  const pic: Record<string, unknown> = {
    "pic:pic": [nvPicPr, blipFill, spPr],
    ":@": { "@_xmlns:pic": PIC_NS },
  };

  const graphicData: Record<string, unknown> = {
    "a:graphicData": [pic],
    ":@": { "@_uri": PIC_NS },
  };
  const graphic: Record<string, unknown> = {
    "a:graphic": [graphicData],
    ":@": { "@_xmlns:a": A_NS },
  };
  inlineChildren.push(graphic);

  if (leaf.properties?.opaqueProps) {
    for (const op of leaf.properties.opaqueProps) {
      inlineChildren.push(opaqueToEntry(op));
    }
  }

  // `xmlns:wp` MUST be declared at or above `<wp:inline>`; we declare it
  // locally because the document root attrs may not include it (see
  // function-level comment above).
  const inlineAttrs: Record<string, string> = { "@_xmlns:wp": WP_NS };
  if (leaf.properties?.distT !== undefined) inlineAttrs["@_distT"] = String(leaf.properties.distT);
  if (leaf.properties?.distB !== undefined) inlineAttrs["@_distB"] = String(leaf.properties.distB);
  if (leaf.properties?.distL !== undefined) inlineAttrs["@_distL"] = String(leaf.properties.distL);
  if (leaf.properties?.distR !== undefined) inlineAttrs["@_distR"] = String(leaf.properties.distR);

  const wpInline: Record<string, unknown> = { "wp:inline": inlineChildren, ":@": inlineAttrs };

  return { "w:drawing": [wpInline] };
}

/**
 * Lightweight metadata used by the renderer when there's no cached `raw`
 * available (e.g. an image leaf produced by `docx:insert-image`). Lives
 * here rather than in the renderer so the renderer doesn't have to know
 * about EMU layout details.
 */
export function inlineImageRendererStub(leaf: InlineImageDrawing): unknown {
  return {
    kind: "inline-image",
    relId: leaf.relId,
    cx: leaf.cx,
    cy: leaf.cy,
    docPrId: leaf.docPrId,
    name: leaf.name,
    ...(leaf.descr !== undefined ? { descr: leaf.descr } : {}),
  };
}

/** Re-export used by serialize.ts to avoid a circular import path. */
export type { OpaqueXml };
