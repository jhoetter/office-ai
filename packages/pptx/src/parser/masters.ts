/**
 * F1 master-editing — typed parsers for slide masters and themes.
 *
 * The current renderer/serializer only need a verbatim blob for these
 * parts (so untouched files round-trip byte-identical), but the master
 * view UI and theme editor planned in
 * `spec/pptx/master-editing.md` need addressable identity: each
 * master's numeric `id` from the parent presentation's
 * `<p:sldMasterIdLst>`, the theme part it points at, and the layouts
 * it owns. This module promotes those parts from `OpaquePart` to
 * typed `SlideMaster` / `Theme` while preserving the raw OOXML in
 * `raw` for byte-faithful round-trip.
 *
 * Layouts already had a typed parser in `parse.ts`
 * (`parseSlideLayoutFromXml`) — here we just enrich the result with
 * the OOXML `type` attribute, the parent master's part path, the
 * `layoutId` from the master's `<p:sldLayoutIdLst>`, and the verbatim
 * `_rels` text. Everything else (placeholders, kind, name, raw) is
 * unchanged.
 */

import { ooxml } from "@officeai/core";
import type { OpaqueXml, SlideLayout, SlideMaster, Theme } from "../model/types.js";
import { PptxParseError } from "./errors.js";
import { REL_TYPE_SLIDE_LAYOUT, REL_TYPE_THEME, resolveTarget } from "./parse.js";
import { attrOf, elementEntries, findElementEntry, readRawAttrs, readRootAttrs } from "./xml-helpers.js";

/**
 * Parse one `ppt/slideMasters/slideMasterN.xml` part into a typed
 * `SlideMaster`. Walks the master's rels to (a) discover the layouts
 * it owns and (b) resolve the linked theme part. The `masterId` is
 * looked up against the parent presentation's `<p:sldMasterIdLst>`
 * via the `masterIdByPath` map the caller assembles upfront.
 */
export function parseSlideMaster(
  container: ooxml.OoxmlContainer,
  partPath: string,
  masterIdByPath: ReadonlyMap<string, number>,
  layoutsByPath: ReadonlyMap<string, SlideLayout>
): SlideMaster {
  const raw = readRawPart(container, partPath, "p:sldMaster");
  const rels = ooxml.RelationshipGraph.loadFor(container, partPath);

  // Layouts the master lists in its rels, in rel order.
  const layoutPartPaths: string[] = [];
  let themePartPath: string | undefined;
  // `<p:sldLayoutIdLst>` lives inside the master XML and pairs each
  // layout's rels id with its numeric `id`. We need that to populate
  // SlideLayout.layoutId for the layouts owned by this master.
  const layoutIdByRelId = readLayoutIdByRelId(raw);
  const layoutIdByPath = new Map<string, number>();
  for (const r of rels.relationships) {
    if (r.type === REL_TYPE_SLIDE_LAYOUT) {
      const target = resolveTarget(partPath, r.target);
      layoutPartPaths.push(target);
      const id = layoutIdByRelId.get(r.id);
      if (id !== undefined) layoutIdByPath.set(target, id);
    } else if (r.type === REL_TYPE_THEME) {
      themePartPath = resolveTarget(partPath, r.target);
    }
  }

  const layouts: SlideLayout[] = [];
  for (const lp of layoutPartPaths) {
    const base = layoutsByPath.get(lp);
    if (!base) continue;
    const layoutType = extractLayoutType(base);
    const layoutRelsXml = loadRelsXml(container, lp);
    const layoutId = layoutIdByPath.get(lp);
    layouts.push({
      ...base,
      ...(layoutId !== undefined ? { layoutId } : {}),
      masterPartPath: partPath,
      ...(layoutType !== undefined ? { type: layoutType } : {}),
      ...(layoutRelsXml !== undefined ? { relsXml: layoutRelsXml } : {}),
    });
  }

  const relsXml = loadRelsXml(container, partPath);

  return {
    partPath,
    masterId: masterIdByPath.get(partPath) ?? 0,
    layouts,
    ...(themePartPath ? { themePartPath } : {}),
    ...(relsXml !== undefined ? { relsXml } : {}),
    raw,
  };
}

/**
 * Re-derive an enriched `SlideLayout` from the typed layout already in
 * the snapshot. Used when the layout's parent master couldn't be
 * resolved from rels (defensive; shouldn't happen on valid PPTX) so
 * the public `layouts` map still carries the F1 fields.
 */
export function enrichLayout(container: ooxml.OoxmlContainer, base: SlideLayout): SlideLayout {
  const relsXml = loadRelsXml(container, base.partPath);
  const type = extractLayoutType(base);
  return {
    ...base,
    ...(type !== undefined ? { type } : {}),
    ...(relsXml !== undefined ? { relsXml } : {}),
  };
}

/**
 * Parse one `ppt/theme/themeN.xml` into a typed `Theme`. The full
 * color scheme parsing lives in `parser/theme.ts` and resolves into
 * `PptxPresentation.themeDefault`; here we only lift the part
 * identity (`partPath`, `name`) — `raw` carries the rest.
 */
export function parseTheme(container: ooxml.OoxmlContainer, partPath: string): Theme {
  const raw = readRawPart(container, partPath, "a:theme");
  const name = raw.attrs["name"];
  return {
    partPath,
    ...(name ? { name } : {}),
    raw,
  };
}

/**
 * Walk the presentation's `<p:sldMasterIdLst>` and turn its
 * `(rId, id)` pairs into a `partPath → masterId` map by resolving
 * each rId against the presentation's rels.
 */
export function buildMasterIdByPath(
  container: ooxml.OoxmlContainer,
  presentationPart: string,
  presentationXml: string
): Map<string, number> {
  const out = new Map<string, number>();
  let tree: unknown;
  try {
    tree = ooxml.parseXml(presentationXml);
  } catch {
    return out;
  }
  if (!Array.isArray(tree)) return out;
  const presEntry = findElementEntry(tree as unknown[], "p:presentation");
  if (!presEntry) return out;
  const presChildren = (presEntry["p:presentation"] as unknown[] | undefined) ?? [];
  const sldMasterIdLst = findElementEntry(presChildren, "p:sldMasterIdLst");
  if (!sldMasterIdLst) return out;

  const rels = ooxml.RelationshipGraph.loadFor(container, presentationPart);
  const targetByRelId = new Map<string, string>();
  for (const r of rels.relationships) targetByRelId.set(r.id, r.target);

  for (const c of elementEntries((sldMasterIdLst["p:sldMasterIdLst"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(c) !== "p:sldMasterId") continue;
    const idStr = attrOf(c, "id");
    const rid = attrOf(c, "r:id");
    if (!idStr || !rid) continue;
    const target = targetByRelId.get(rid);
    if (!target) continue;
    const partPath = resolveTarget(presentationPart, target);
    const id = Number(idStr);
    if (!Number.isNaN(id)) out.set(partPath, id);
  }
  return out;
}

// ─── Internals ────────────────────────────────────────────────────────────

function readRawPart(container: ooxml.OoxmlContainer, partPath: string, rootTag: string): OpaqueXml {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(partPath));
  } catch (err) {
    throw new PptxParseError("invalid-xml", `Failed to parse ${partPath}`, {
      partPath,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) {
    throw new PptxParseError("invalid-xml", `Expected XML tree at ${partPath}`, { partPath });
  }
  const r = findElementEntry(tree as unknown[], rootTag);
  if (!r) {
    throw new PptxParseError("invalid-xml", `Missing <${rootTag}> in ${partPath}`, { partPath });
  }
  return {
    tag: rootTag,
    attrs: readRootAttrs(r),
    rawAttrs: readRawAttrs(r),
    subtree: (r[rootTag] as unknown[] | undefined) ?? [],
  };
}

/**
 * Walk the master's `<p:sldLayoutIdLst>` and return a `relId → numeric id`
 * map. Each `<p:sldLayoutId id="N" r:id="rIdM"/>` row pairs the two ids
 * we need to thread the numeric `layoutId` onto each `SlideLayout`.
 */
function readLayoutIdByRelId(raw: OpaqueXml): Map<string, number> {
  const out = new Map<string, number>();
  const lst = findElementEntry(raw.subtree, "p:sldLayoutIdLst");
  if (!lst) return out;
  for (const c of elementEntries((lst["p:sldLayoutIdLst"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(c) !== "p:sldLayoutId") continue;
    const idStr = attrOf(c, "id");
    const rid = attrOf(c, "r:id");
    if (!idStr || !rid) continue;
    const id = Number(idStr);
    if (!Number.isNaN(id)) out.set(rid, id);
  }
  return out;
}

/**
 * Pull `<p:sldLayout type="...">` straight off the captured raw blob.
 * The parser already promoted this to `SlideLayout.kind` via the
 * classifier, but the OOXML token itself is what the master view picker
 * shows in tooltips and what round-trips most cleanly to PowerPoint.
 */
function extractLayoutType(layout: SlideLayout): string | undefined {
  const t = layout.raw.attrs["type"] ?? layout.raw.rawAttrs["@_type"];
  return t && t.length > 0 ? t : undefined;
}

function loadRelsXml(container: ooxml.OoxmlContainer, partPath: string): string | undefined {
  const relsPath = ooxml.RelationshipGraph.relsPathFor(partPath);
  if (!container.has(relsPath)) return undefined;
  try {
    return container.readText(relsPath);
  } catch {
    return undefined;
  }
}
