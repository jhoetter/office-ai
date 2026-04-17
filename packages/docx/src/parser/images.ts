import { ooxml, type IdMinter } from "@officeai/core";
import type {
  DrawingLeaf,
  InlineImageDrawing,
  InlineImageProperties,
  OpaqueDrawing,
} from "../model/types.js";
import { attrOf, captureOpaque, elementEntries, findElementEntry } from "./xml-helpers.js";

/**
 * Try to recognise a `<w:drawing>` entry as a typed inline image (a
 * `<wp:inline>` containing a `pic:pic` whose `a:blip` carries `r:embed`).
 * If recognition succeeds we return an `InlineImageDrawing` with the
 * relationship id, EMU dimensions, and `<wp:docPr>` metadata extracted;
 * the entire `<w:drawing>` subtree is also cached in `raw` so the
 * serializer can re-emit byte-identical bytes when no typed field has
 * changed.
 *
 * Anything we don't recognise (anchored floats, charts, shapes,
 * SmartArt, embedded objects, …) falls through to the catch-all
 * `OpaqueDrawing` branch so it round-trips verbatim. Typed promotion is
 * intentionally one drawing class at a time — adding `subkind`s in
 * future workstreams must not require changing the un-promoted parts.
 */
export function parseDrawing(entry: Record<string, unknown>, mintNodeId: IdMinter): DrawingLeaf {
  const inline = parseInlineImage(entry, mintNodeId);
  if (inline) return inline;
  return parseOpaqueDrawing(entry, mintNodeId);
}

function parseOpaqueDrawing(entry: Record<string, unknown>, mintNodeId: IdMinter): OpaqueDrawing {
  return { kind: "drawing", subkind: "opaque", id: mintNodeId(), raw: captureOpaque(entry) };
}

function parseInlineImage(entry: Record<string, unknown>, mintNodeId: IdMinter): InlineImageDrawing | null {
  const drawingChildren = (entry["w:drawing"] as unknown[] | undefined) ?? [];
  const wpInline = findElementEntry(drawingChildren, "wp:inline");
  if (!wpInline) return null;
  const inlineChildren = (wpInline["wp:inline"] as unknown[] | undefined) ?? [];

  const extent = findElementEntry(inlineChildren, "wp:extent");
  if (!extent) return null;
  const cx = numAttr(extent, "cx");
  const cy = numAttr(extent, "cy");
  if (cx === null || cy === null || cx <= 0 || cy <= 0) return null;

  const docPrEntry = findElementEntry(inlineChildren, "wp:docPr");
  if (!docPrEntry) return null;
  const docPrId = numAttr(docPrEntry, "id");
  if (docPrId === null) return null;
  const name = attrOf(docPrEntry, "name") ?? "";
  const descr = attrOf(docPrEntry, "descr");
  const title = attrOf(docPrEntry, "title");

  const graphic = findElementEntry(inlineChildren, "a:graphic");
  if (!graphic) return null;
  const graphicChildren = (graphic["a:graphic"] as unknown[] | undefined) ?? [];
  const graphicData = findElementEntry(graphicChildren, "a:graphicData");
  if (!graphicData) return null;
  const uri = attrOf(graphicData, "uri");
  if (uri !== "http://schemas.openxmlformats.org/drawingml/2006/picture") return null;

  const graphicDataChildren = (graphicData["a:graphicData"] as unknown[] | undefined) ?? [];
  const pic = findElementEntry(graphicDataChildren, "pic:pic");
  if (!pic) return null;
  const blipFill = findElementEntry((pic["pic:pic"] as unknown[] | undefined) ?? [], "pic:blipFill");
  if (!blipFill) return null;
  const blip = findElementEntry((blipFill["pic:blipFill"] as unknown[] | undefined) ?? [], "a:blip");
  if (!blip) return null;
  const relId = attrOf(blip, "r:embed");
  if (!relId) return null;

  const properties: InlineImageProperties = {};
  const distT = numAttr(wpInline, "distT");
  const distB = numAttr(wpInline, "distB");
  const distL = numAttr(wpInline, "distL");
  const distR = numAttr(wpInline, "distR");
  if (distT !== null) (properties as { distT?: number }).distT = distT;
  if (distB !== null) (properties as { distB?: number }).distB = distB;
  if (distL !== null) (properties as { distL?: number }).distL = distL;
  if (distR !== null) (properties as { distR?: number }).distR = distR;
  if (title !== undefined) (properties as { title?: string }).title = title;
  const effectExtent = findElementEntry(inlineChildren, "wp:effectExtent");
  if (effectExtent) {
    const t = numAttr(effectExtent, "t") ?? 0;
    const r = numAttr(effectExtent, "r") ?? 0;
    const b = numAttr(effectExtent, "b") ?? 0;
    const l = numAttr(effectExtent, "l") ?? 0;
    (properties as { effectExtent?: { t: number; r: number; b: number; l: number } }).effectExtent = {
      t,
      r,
      b,
      l,
    };
  }

  // Capture every wp:inline child we don't typed-model, in document order,
  // so the byte-equality guarantee survives even when commands rebuild
  // the leaf from typed fields. We skip entries whose tag we DID model.
  const opaqueProps: ReturnType<typeof captureOpaque>[] = [];
  const HANDLED_TAGS = new Set(["wp:extent", "wp:effectExtent", "wp:docPr", "a:graphic"]);
  for (const c of elementEntries(inlineChildren)) {
    const tag = ooxml.getTag(c);
    if (HANDLED_TAGS.has(tag)) continue;
    opaqueProps.push(captureOpaque(c));
  }
  if (opaqueProps.length > 0) {
    (properties as { opaqueProps?: ReadonlyArray<ReturnType<typeof captureOpaque>> }).opaqueProps =
      opaqueProps;
  }

  const out: InlineImageDrawing = {
    kind: "drawing",
    subkind: "inline-image",
    id: mintNodeId(),
    relId,
    cx,
    cy,
    docPrId,
    name,
    ...(descr !== undefined ? { descr } : {}),
    ...(Object.keys(properties).length > 0 ? { properties } : {}),
    raw: captureOpaque(entry),
  };
  return out;
}

function numAttr(entry: Record<string, unknown>, attr: string): number | null {
  const v = attrOf(entry, attr);
  if (v === undefined) return null;
  const n = Number(v);
  if (!Number.isFinite(n)) return null;
  return n;
}
