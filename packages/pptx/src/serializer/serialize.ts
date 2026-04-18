import { ooxml } from "@officeai/core";
import type {
  GroupShape,
  OpaqueShape,
  OpaqueXml,
  Picture,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Shape,
  Slide,
  TextBody,
  TextParagraph,
  TextRun,
  TextShape,
} from "../model/types.js";
import { ATTR_KEY, ATTR_PREFIX, opaqueToEntry } from "../parser/xml-helpers.js";
import { PptxSerializeError } from "./errors.js";

const PRESENTATION_PART = "ppt/presentation.xml";
const PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export async function serializePptx(snapshot: PptxSnapshot): Promise<ArrayBuffer> {
  const container = snapshot.container.clone();

  // 1) Drop removed parts (and their rels).
  for (const partPath of snapshot.removedParts) {
    container.removePart(partPath);
    const relsPath = ooxml.RelationshipGraph.relsPathFor(partPath);
    if (container.has(relsPath)) container.removePart(relsPath);
  }

  // 2) Rewrite dirty slide parts.
  for (const slide of snapshot.root.slides) {
    if (!snapshot.dirty.slides.has(slide.partPath)) continue;
    try {
      const xml = serializeSlideXml(slide);
      container.writeText(slide.partPath, xml);
    } catch (err) {
      throw new PptxSerializeError("slide-failed", `Failed to serialize ${slide.partPath}`, {
        partPath: slide.partPath,
        cause: err,
      });
    }
  }

  // 3) Rewrite dirty media parts (binary bytes).
  for (const partPath of snapshot.dirty.media) {
    const media = snapshot.root.media.get(partPath);
    if (!media) continue;
    if (container.has(partPath)) {
      container.writeBytes(partPath, media.bytes);
    } else {
      container.addPart(partPath, media.bytes);
    }
  }

  // 4) Rewrite dirty rels parts.
  for (const relsPath of snapshot.dirty.relationships) {
    const snap = snapshot.relationships.get(relsPath);
    if (!snap) {
      // Snapshot expects this rels part to exist; if missing, skip silently.
      continue;
    }
    const xml = serializeRelsXml(snap);
    container.writeText(relsPath, xml);
  }

  // 5) Rewrite presentation.xml if dirty (slide order, sldIdLst).
  if (snapshot.dirty.presentation) {
    try {
      const xml = serializePresentationXml(snapshot.root);
      container.writeText(PRESENTATION_PART, xml);
    } catch (err) {
      throw new PptxSerializeError(
        "presentation-failed",
        `Failed to serialize ${PRESENTATION_PART}`,
        { partPath: PRESENTATION_PART, cause: err }
      );
    }
  }

  // 6) Rewrite content types if dirty.
  if (snapshot.dirty.contentTypes) {
    const ct = new ooxml.ContentTypes(
      snapshot.contentTypes.defaults.map((d) => ({
        extension: d.extension,
        contentType: d.contentType,
      })),
      snapshot.contentTypes.overrides.map((o) => ({
        partName: o.partName,
        contentType: o.contentType,
      }))
    );
    ct.writeBack(container);
  }

  return container.serialize();
}

// ─── Slide serialization ──────────────────────────────────────────────────

function serializeSlideXml(slide: Slide): string {
  const spTreeChildren: unknown[] = [];
  for (const head of slide.spTreeHead) spTreeChildren.push(opaqueToEntry(head));
  for (const shape of slide.shapes) spTreeChildren.push(shapeToEntry(shape));
  const spTree = makeEntry("p:spTree", spTreeChildren);

  const cSldChildren: unknown[] = [];
  for (const head of slide.cSldHead) cSldChildren.push(opaqueToEntry(head));
  cSldChildren.push(spTree);
  const cSld = makeEntry("p:cSld", cSldChildren, slide.cSldAttrs);

  const sldChildren: unknown[] = [cSld];
  for (const tail of slide.slideOpaqueTail) sldChildren.push(opaqueToEntry(tail));
  const sld = makeEntry("p:sld", sldChildren, slide.slideRootAttrs);

  return ooxml.serializeXml([sld]);
}

function shapeToEntry(shape: Shape): Record<string, unknown> {
  switch (shape.kind) {
    case "text":
      return textShapeToEntry(shape);
    case "pic":
      return pictureToEntry(shape);
    case "group":
      return groupShapeToEntry(shape);
    case "opaque":
      return opaqueShapeToEntry(shape);
  }
}

function textShapeToEntry(shape: TextShape): Record<string, unknown> {
  const nvSpPrChildren: unknown[] = [];
  // We rebuild p:cNvPr from model id+name; everything else captured opaquely
  // (p:cNvSpPr, p:nvPr including ph) is passed through verbatim.
  let emittedCNvPr = false;
  for (const o of shape.nvSpPrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvSpPrChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else {
      nvSpPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvSpPrChildren.unshift(
      makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name })
    );
  }

  const nvSpPr = makeEntry("p:nvSpPr", nvSpPrChildren);

  const spPrChildren: unknown[] = [];
  let emittedXfrm = false;
  for (const o of shape.spPrTail) {
    if (o.tag === "a:xfrm") {
      spPrChildren.push(buildXfrm(shape.position, shape.size, o));
      emittedXfrm = true;
    } else {
      spPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedXfrm && (shape.position || shape.size)) {
    spPrChildren.unshift(buildXfrm(shape.position, shape.size, undefined));
  }
  const spPr = makeEntry("p:spPr", spPrChildren);

  const children: unknown[] = [nvSpPr, spPr];
  if (shape.styleRaw) children.push(opaqueToEntry(shape.styleRaw));
  children.push(textBodyToEntry(shape.txBody));

  return makeEntry("p:sp", children);
}

function pictureToEntry(shape: Picture): Record<string, unknown> {
  const nvPicPrChildren: unknown[] = [];
  let emittedCNvPr = false;
  for (const o of shape.nvPicPrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvPicPrChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else {
      nvPicPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvPicPrChildren.unshift(
      makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name })
    );
  }
  const nvPicPr = makeEntry("p:nvPicPr", nvPicPrChildren);

  const blipFillChildren: unknown[] = [];
  let emittedBlip = false;
  for (const o of shape.blipFillTail) {
    if (o.tag === "a:blip" && !emittedBlip) {
      blipFillChildren.push(rebuildBlip(shape.mediaRelId, o));
      emittedBlip = true;
    } else {
      blipFillChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedBlip) {
    blipFillChildren.unshift(makeEntry("a:blip", [], { "r:embed": shape.mediaRelId }));
  }
  const blipFill = makeEntry("p:blipFill", blipFillChildren);

  const spPrChildren: unknown[] = [];
  let emittedXfrm = false;
  for (const o of shape.spPrTail) {
    if (o.tag === "a:xfrm") {
      spPrChildren.push(buildXfrm(shape.position, shape.size, o));
      emittedXfrm = true;
    } else {
      spPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedXfrm && (shape.position || shape.size)) {
    spPrChildren.unshift(buildXfrm(shape.position, shape.size, undefined));
  }
  const spPr = makeEntry("p:spPr", spPrChildren);

  const children: unknown[] = [nvPicPr, blipFill, spPr];
  if (shape.styleRaw) children.push(opaqueToEntry(shape.styleRaw));
  return makeEntry("p:pic", children);
}

function groupShapeToEntry(shape: GroupShape): Record<string, unknown> {
  const nvChildren: unknown[] = [];
  let emittedCNvPr = false;
  for (const o of shape.nvGrpSpPrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else {
      nvChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvChildren.unshift(
      makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name })
    );
  }
  const nvGrpSpPr = makeEntry("p:nvGrpSpPr", nvChildren);

  const grpSpPrChildren: unknown[] = [];
  let emittedXfrm = false;
  for (const o of shape.grpSpPrTail) {
    if (o.tag === "a:xfrm") {
      grpSpPrChildren.push(buildGroupXfrm(shape, o));
      emittedXfrm = true;
    } else {
      grpSpPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedXfrm && (shape.position || shape.size || shape.chOffExtRaw.length > 0)) {
    grpSpPrChildren.unshift(buildGroupXfrm(shape, undefined));
  }
  const grpSpPr = makeEntry("p:grpSpPr", grpSpPrChildren);

  const children: unknown[] = [nvGrpSpPr, grpSpPr];
  for (const c of shape.children) children.push(shapeToEntry(c));
  return makeEntry("p:grpSp", children);
}

function opaqueShapeToEntry(shape: OpaqueShape): Record<string, unknown> {
  return opaqueToEntry(shape.raw);
}

function rebuildCNvPr(id: number, name: string, captured: OpaqueXml): Record<string, unknown> {
  // Preserve any sub-children of the original p:cNvPr (e.g. <a:hlinkClick>).
  const attrs: Record<string, string> = { ...captured.attrs };
  attrs.id = String(id);
  attrs.name = name;
  const rawAttrs = makeRawAttrs(attrs);
  const entry: Record<string, unknown> = { "p:cNvPr": captured.subtree };
  if (Object.keys(rawAttrs).length > 0) entry[ATTR_KEY] = rawAttrs;
  return entry;
}

function rebuildBlip(relId: string, captured: OpaqueXml): Record<string, unknown> {
  const attrs: Record<string, string> = { ...captured.attrs };
  attrs["r:embed"] = relId;
  const rawAttrs = makeRawAttrs(attrs);
  const entry: Record<string, unknown> = { "a:blip": captured.subtree };
  if (Object.keys(rawAttrs).length > 0) entry[ATTR_KEY] = rawAttrs;
  return entry;
}

function buildXfrm(
  position: { xEmu: number; yEmu: number } | undefined,
  size: { cxEmu: number; cyEmu: number } | undefined,
  captured: OpaqueXml | undefined
): Record<string, unknown> {
  // Preserve attrs on a:xfrm itself (e.g. flipH, rot).
  const xfrmAttrs = captured ? captured.attrs : {};
  const subChildren: unknown[] = [];
  // Re-emit a:off and a:ext from model when present; else fall back to captured subtree.
  if (position) {
    subChildren.push(makeEntry("a:off", [], { x: String(position.xEmu), y: String(position.yEmu) }));
  }
  if (size) {
    subChildren.push(makeEntry("a:ext", [], { cx: String(size.cxEmu), cy: String(size.cyEmu) }));
  }
  if (!position && !size && captured) {
    return opaqueToEntry(captured);
  }
  const entry: Record<string, unknown> = { "a:xfrm": subChildren };
  if (Object.keys(xfrmAttrs).length > 0) entry[ATTR_KEY] = makeRawAttrs(xfrmAttrs);
  return entry;
}

function buildGroupXfrm(
  shape: GroupShape,
  captured: OpaqueXml | undefined
): Record<string, unknown> {
  const xfrmAttrs = captured ? captured.attrs : {};
  const subChildren: unknown[] = [];
  if (shape.position) {
    subChildren.push(
      makeEntry("a:off", [], { x: String(shape.position.xEmu), y: String(shape.position.yEmu) })
    );
  }
  if (shape.size) {
    subChildren.push(
      makeEntry("a:ext", [], { cx: String(shape.size.cxEmu), cy: String(shape.size.cyEmu) })
    );
  }
  for (const o of shape.chOffExtRaw) subChildren.push(opaqueToEntry(o));
  const entry: Record<string, unknown> = { "a:xfrm": subChildren };
  if (Object.keys(xfrmAttrs).length > 0) entry[ATTR_KEY] = makeRawAttrs(xfrmAttrs);
  return entry;
}

// ─── Text body serialization ──────────────────────────────────────────────

function textBodyToEntry(body: TextBody): Record<string, unknown> {
  const children: unknown[] = [];
  if (body.bodyPrRaw) children.push(opaqueToEntry(body.bodyPrRaw));
  else children.push(makeEntry("a:bodyPr", []));
  if (body.lstStyleRaw) children.push(opaqueToEntry(body.lstStyleRaw));
  else children.push(makeEntry("a:lstStyle", []));
  for (const p of body.paragraphs) children.push(paragraphToEntry(p));
  return makeEntry("p:txBody", children);
}

function paragraphToEntry(p: TextParagraph): Record<string, unknown> {
  const children: unknown[] = [];
  if (p.properties.opaqueAttrs || p.properties.opaqueChildren) {
    const pPrChildren: unknown[] = [];
    for (const c of p.properties.opaqueChildren ?? []) pPrChildren.push(opaqueToEntry(c));
    const attrs = p.properties.opaqueAttrs ?? {};
    if (Object.keys(attrs).length > 0 || pPrChildren.length > 0) {
      const entry: Record<string, unknown> = { "a:pPr": pPrChildren };
      if (Object.keys(attrs).length > 0) entry[ATTR_KEY] = makeRawAttrs(attrs);
      children.push(entry);
    }
  } else if (
    p.properties.alignment !== undefined ||
    p.properties.level !== undefined
  ) {
    const attrs: Record<string, string> = {};
    if (p.properties.level !== undefined) attrs.lvl = String(p.properties.level);
    if (p.properties.alignment) {
      const map: Record<NonNullable<typeof p.properties.alignment>, string> = {
        left: "l",
        center: "ctr",
        right: "r",
        justify: "just",
      };
      attrs.algn = map[p.properties.alignment];
    }
    const entry: Record<string, unknown> = { "a:pPr": [] };
    if (Object.keys(attrs).length > 0) entry[ATTR_KEY] = makeRawAttrs(attrs);
    children.push(entry);
  }
  for (const r of p.runs) {
    if (r.isLineBreak) children.push(brToEntry(r));
    else children.push(runToEntry(r));
  }
  if (p.endParaRPrRaw) children.push(opaqueToEntry(p.endParaRPrRaw));
  return makeEntry("a:p", children);
}

function runToEntry(r: TextRun): Record<string, unknown> {
  const rPrChildren: unknown[] = [];
  for (const c of r.properties.opaqueChildren ?? []) rPrChildren.push(opaqueToEntry(c));
  const rPrAttrs = mergeRunAttrs(r);
  const rPrEntry: Record<string, unknown> = { "a:rPr": rPrChildren };
  if (Object.keys(rPrAttrs).length > 0) rPrEntry[ATTR_KEY] = makeRawAttrs(rPrAttrs);

  const tEntry: Record<string, unknown> = { "a:t": [{ "#text": r.text }] };
  // Add xml:space=preserve when text has leading/trailing whitespace.
  if (r.text !== r.text.trim()) {
    tEntry[ATTR_KEY] = { [`${ATTR_PREFIX}xml:space`]: "preserve" };
  }
  return makeEntry("a:r", [rPrEntry, tEntry]);
}

function brToEntry(r: TextRun): Record<string, unknown> {
  const rPrChildren: unknown[] = [];
  for (const c of r.properties.opaqueChildren ?? []) rPrChildren.push(opaqueToEntry(c));
  const rPrAttrs = mergeRunAttrs(r);
  const rPrEntry: Record<string, unknown> = { "a:rPr": rPrChildren };
  if (Object.keys(rPrAttrs).length > 0) rPrEntry[ATTR_KEY] = makeRawAttrs(rPrAttrs);
  return makeEntry("a:br", [rPrEntry]);
}

function mergeRunAttrs(r: TextRun): Record<string, string> {
  const out: Record<string, string> = { ...(r.properties.opaqueAttrs ?? {}) };
  if (r.properties.bold !== undefined) out.b = r.properties.bold ? "1" : "0";
  if (r.properties.italic !== undefined) out.i = r.properties.italic ? "1" : "0";
  if (r.properties.underline !== undefined) {
    if (r.properties.underline === false) out.u = "none";
    else if (r.properties.underline === true) out.u = "sng";
    else out.u = String(r.properties.underline);
  }
  if (r.properties.strike !== undefined) {
    out.strike = r.properties.strike ? "sngStrike" : "noStrike";
  }
  if (r.properties.fontSizeHundredths !== undefined) {
    out.sz = String(r.properties.fontSizeHundredths);
  }
  return out;
}

// ─── Presentation serialization ───────────────────────────────────────────

function serializePresentationXml(root: PptxPresentation): string {
  // We rebuild from the captured presentationOpaqueTail, replacing
  // <p:sldIdLst> in-place with one that matches root.slides order.
  const newChildren: unknown[] = [];
  for (const o of root.presentationOpaqueTail) {
    if (o.tag === "p:sldIdLst") {
      newChildren.push(buildSldIdLst(root));
    } else {
      newChildren.push(opaqueToEntry(o));
    }
  }
  // If the original presentation had no sldIdLst, insert one before sldSz.
  const hasSldIdLst = root.presentationOpaqueTail.some((o) => o.tag === "p:sldIdLst");
  if (!hasSldIdLst) {
    const idx = newChildren.findIndex((n) => {
      if (!n || typeof n !== "object" || Array.isArray(n)) return false;
      const obj = n as Record<string, unknown>;
      const keys = Object.keys(obj).filter((k) => k !== ATTR_KEY);
      return keys[0] === "p:sldSz";
    });
    if (idx >= 0) newChildren.splice(idx, 0, buildSldIdLst(root));
    else newChildren.unshift(buildSldIdLst(root));
  }

  const presEntry: Record<string, unknown> = { "p:presentation": newChildren };
  if (Object.keys(root.presentationRootAttrs).length > 0) {
    presEntry[ATTR_KEY] = makeRawAttrs(root.presentationRootAttrs);
  }
  return ooxml.serializeXml([presEntry]);
}

function buildSldIdLst(root: PptxPresentation): Record<string, unknown> {
  const children: unknown[] = root.slides.map((s) =>
    makeEntry("p:sldId", [], { id: String(s.slideId), "r:id": s.relId })
  );
  const entry: Record<string, unknown> = { "p:sldIdLst": children };
  if (Object.keys(root.sldIdLstAttrs).length > 0) {
    entry[ATTR_KEY] = makeRawAttrs(root.sldIdLstAttrs);
  }
  return entry;
}

// ─── Rels serialization ───────────────────────────────────────────────────

function serializeRelsXml(snap: RelationshipsSnap): string {
  const children = snap.entries.map((r) =>
    makeEntry("Relationship", [], {
      Id: r.id,
      Type: r.type,
      Target: r.target,
      ...(r.targetMode ? { TargetMode: r.targetMode } : {}),
    })
  );
  const tree = [makeEntry("Relationships", children, { xmlns: RELS_NS })];
  return ooxml.serializeXml(tree);
}

// ─── Generic helpers ──────────────────────────────────────────────────────

function makeEntry(
  tag: string,
  children: ReadonlyArray<unknown>,
  attrs?: Record<string, string>
): Record<string, unknown> {
  const entry: Record<string, unknown> = { [tag]: children };
  if (attrs && Object.keys(attrs).length > 0) {
    entry[ATTR_KEY] = makeRawAttrs(attrs);
  }
  return entry;
}

function makeRawAttrs(attrs: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    out[`${ATTR_PREFIX}${k}`] = v;
  }
  return out;
}
