import { ooxml } from "@officeai/core";
import { buildChartGrid, buildEmbeddedXlsx } from "@officeai/xlsx";
import type {
  ChartPart,
  ChartShape,
  ConnectorDashStyle,
  ConnectorEndpoint,
  ConnectorShape,
  ConnectorSide,
  GroupShape,
  OleSpreadsheetShape,
  OpaqueShape,
  OpaqueXml,
  Picture,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Shape,
  ShapeAnimation,
  Slide,
  SlideTransition,
  TableShape,
  TextBody,
  TextParagraph,
  TextRun,
  TextShape,
} from "../model/types.js";
import { connectorXfrm, resolveEndpoint } from "../model/connector-geometry.js";
import { ATTR_KEY, ATTR_PREFIX, opaqueToEntry } from "../parser/xml-helpers.js";
import { findPreset, subtypeFor, type EmitHelpers } from "../animation/presets.js";
import { mediaShapeToEntry } from "./media.js";
import { PptxSerializeError } from "./errors.js";

const PRESENTATION_PART = "ppt/presentation.xml";
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

  // 3b) Rewrite dirty embedded binary parts (OLE-Excel `.xlsx`
  // packages, …) and register the matching content-type override so
  // Office recognises the part. Newly-authored parts carry a
  // `pendingGrid` instead of `bytes`; we materialise those via
  // `buildEmbeddedXlsx` here so command handlers stay synchronous.
  // Owning relationships are produced by `pptx:insert-spreadsheet`
  // and round-trip through the existing rels-write pass.
  if (snapshot.dirty.embeddings.size > 0) {
    const embedContentTypes = ooxml.ContentTypes.load(container);
    let touchedEmbedCt = false;
    for (const partPath of snapshot.dirty.embeddings) {
      const part = snapshot.root.embeddings.get(partPath);
      if (!part) continue;
      let bytes = part.bytes;
      if (!bytes && part.pendingGrid) {
        const built = await buildEmbeddedXlsx(part.pendingGrid, {
          sheetName: part.pendingSheetName ?? "Sheet1",
        });
        bytes = built.bytes;
      }
      if (!bytes) continue;
      if (container.has(partPath)) container.writeBytes(partPath, bytes);
      else container.addPart(partPath, bytes);
      const overrideName = partPath.startsWith("/") ? partPath : `/${partPath}`;
      if (!embedContentTypes.hasOverride(overrideName)) {
        embedContentTypes.addOverride(overrideName, part.contentType);
        touchedEmbedCt = true;
      }
    }
    if (touchedEmbedCt) embedContentTypes.writeBack(container);
  }

  // 3a) Rewrite dirty layout parts. We always emit the layout's verbatim
  // `raw` blob — the typed fields (`kind`, `name`, `placeholders`) are
  // derived, so there's nothing to serialise from the typed model that
  // isn't already captured in `raw`.
  for (const layoutPath of snapshot.dirty.layouts) {
    const layout = snapshot.root.layouts.get(layoutPath);
    if (!layout) continue;
    try {
      const xml = serializeLayoutXml(layout.raw);
      if (container.has(layoutPath)) container.writeText(layoutPath, xml);
      else container.addPart(layoutPath, new TextEncoder().encode(xml));
    } catch (err) {
      throw new PptxSerializeError("layout-failed", `Failed to serialize ${layoutPath}`, {
        partPath: layoutPath,
        cause: err,
      });
    }
  }

  // 3a1a) Rewrite dirty master parts. Slide masters are typed as
  // `OpaquePart` (the typed model never edits master XML directly —
  // future master/theme commands will mutate the raw blob and set
  // this dirty flag). Same emit shape as layouts: the captured raw
  // tree is the single source of truth.
  for (const masterPath of snapshot.dirty.masters) {
    const master = snapshot.root.masters.get(masterPath);
    if (!master) continue;
    try {
      const xml = serializeLayoutXml(master.raw);
      if (container.has(masterPath)) container.writeText(masterPath, xml);
      else container.addPart(masterPath, new TextEncoder().encode(xml));
    } catch (err) {
      throw new PptxSerializeError("layout-failed", `Failed to serialize ${masterPath}`, {
        partPath: masterPath,
        cause: err,
      });
    }
  }

  // 3a1b) Rewrite dirty theme parts (`ppt/theme/themeN.xml`). Same
  // shape as masters — the typed model treats them as opaque, but
  // a future theme-color command can mutate `raw` and set this flag
  // to flush the change back to disk.
  for (const themePath of snapshot.dirty.theme) {
    const theme = snapshot.root.theme.get(themePath);
    if (!theme) continue;
    try {
      const xml = serializeLayoutXml(theme.raw);
      if (container.has(themePath)) container.writeText(themePath, xml);
      else container.addPart(themePath, new TextEncoder().encode(xml));
    } catch (err) {
      throw new PptxSerializeError("layout-failed", `Failed to serialize ${themePath}`, {
        partPath: themePath,
        cause: err,
      });
    }
  }

  // 3a2) Rewrite dirty notes parts. The typed `body` lives inside the
  // verbatim `raw` blob (set-slide-notes rebuilds raw from body), so we
  // simply emit raw — no extra reconciliation needed.
  for (const notesPath of snapshot.dirty.notesSlides) {
    const notes = snapshot.root.notesSlides.get(notesPath);
    if (!notes) continue;
    try {
      const xml = serializeLayoutXml(notes.raw);
      if (container.has(notesPath)) container.writeText(notesPath, xml);
      else container.addPart(notesPath, new TextEncoder().encode(xml));
    } catch (err) {
      throw new PptxSerializeError("notes-failed", `Failed to serialize ${notesPath}`, {
        partPath: notesPath,
        cause: err,
      });
    }
  }

  // 3a3) Rewrite dirty per-slide comments parts.
  for (const partPath of snapshot.dirty.comments) {
    const part = snapshot.root.commentsByPart.get(partPath);
    if (!part) continue;
    try {
      const xml = serializeCommentsXml(part);
      if (container.has(partPath)) container.writeText(partPath, xml);
      else container.addPart(partPath, new TextEncoder().encode(xml));
    } catch (err) {
      throw new PptxSerializeError("comments-failed", `Failed to serialize ${partPath}`, {
        partPath,
        cause: err,
      });
    }
  }

  // 3a4) Rewrite the deck-wide commentAuthors part if dirty.
  if (snapshot.dirty.commentAuthors && snapshot.root.commentAuthors) {
    const part = snapshot.root.commentAuthors;
    try {
      const xml = serializeCommentAuthorsXml(part);
      if (container.has(part.partPath)) container.writeText(part.partPath, xml);
      else container.addPart(part.partPath, new TextEncoder().encode(xml));
    } catch (err) {
      throw new PptxSerializeError("comment-authors-failed", `Failed to serialize ${part.partPath}`, {
        partPath: part.partPath,
        cause: err,
      });
    }
  }

  // 3b) Rewrite dirty chart parts (F3) + refresh embedded xlsx.
  // We materialise the live ContentTypes once so chart serialisation can
  // register `<Override>` entries for newly authored embedded packages
  // without races against the dirty-flag block below.
  const liveContentTypes = new ooxml.ContentTypes(
    snapshot.contentTypes.defaults.map((d) => ({
      extension: d.extension,
      contentType: d.contentType,
    })),
    snapshot.contentTypes.overrides.map((o) => ({
      partName: o.partName,
      contentType: o.contentType,
    }))
  );
  let contentTypesTouchedByCharts = false;
  for (const chartPath of snapshot.dirty.charts) {
    const part = snapshot.root.charts.get(chartPath);
    if (!part) continue;
    try {
      const result = await serializeChartWithEmbedding(container, liveContentTypes, part);
      if (result.contentTypesChanged) contentTypesTouchedByCharts = true;
    } catch (err) {
      throw new PptxSerializeError("chart-failed", `Failed to serialize ${chartPath}`, {
        partPath: chartPath,
        cause: err,
      });
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
      throw new PptxSerializeError("presentation-failed", `Failed to serialize ${PRESENTATION_PART}`, {
        partPath: PRESENTATION_PART,
        cause: err,
      });
    }
  }

  // 6) Rewrite content types if dirty.
  if (snapshot.dirty.contentTypes || contentTypesTouchedByCharts) {
    liveContentTypes.writeBack(container);
  }

  return container.serialize();
}

/**
 * Serialise one chart part. When the chart references (or implicitly
 * needs) an embedded xlsx package — which is what Word/PowerPoint
 * authors for `Insert > Chart` and what enables "Edit Data" round-trip
 * — we (re)build the workbook bytes from the typed model, register
 * the part + content-type override + `package` relationship from the
 * chart, and emit `<c:externalData>` pointing at it.
 */
async function serializeChartWithEmbedding(
  container: ooxml.OoxmlContainer,
  contentTypes: ooxml.ContentTypes,
  part: ChartPart
): Promise<{ contentTypesChanged: boolean }> {
  const sheetName = part.embeddingSheetName ?? "Sheet1";
  const chartGrid = buildChartGrid(
    [...part.categories],
    part.series.map((s) => ({ values: [...s.values], ...(s.name !== undefined ? { name: s.name } : {}) })),
    { sheetName }
  );

  const embeddingPath = part.embeddingPartPath ?? mintChartEmbeddingPath(container, part.partPath);

  const built = await buildEmbeddedXlsx(chartGrid.grid, { sheetName });

  const chartRels = ooxml.RelationshipGraph.loadFor(container, part.partPath);
  const relTarget = ooxml.relativeTarget(part.partPath, embeddingPath);
  const added = ooxml.addEmbeddedPart({
    container,
    contentTypes,
    ownerRels: chartRels,
    partPath: embeddingPath,
    bytes: built.bytes,
    contentType: ooxml.CT_SPREADSHEETML_SHEET,
    relTarget,
    relType: ooxml.REL_TYPE_PACKAGE,
    ...(part.embeddingRelId ? { relId: part.embeddingRelId } : {}),
  });
  chartRels.writeBack(container);

  const xml = serializeChartPartXml(part, {
    embeddingRelId: added.relId,
    categoryRef: chartGrid.categoryRef,
    valueRefs: chartGrid.valueRefs,
    nameRefs: chartGrid.nameRefs,
  });
  if (container.has(part.partPath)) container.writeText(part.partPath, xml);
  else container.addPart(part.partPath, new TextEncoder().encode(xml));

  return { contentTypesChanged: true };
}

/**
 * Pick the next free `ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`
 * path. Mirrors PowerPoint's own naming so an opened-then-saved file
 * stays human-grokable when inspecting the package.
 */
function mintChartEmbeddingPath(container: ooxml.OoxmlContainer, chartPartPath: string): string {
  const root = chartPartPath.startsWith("ppt/") ? "ppt" : (chartPartPath.split("/")[0] ?? "ppt");
  let n = 1;
  while (container.has(`${root}/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`)) n++;
  return `${root}/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
}

// ─── Slide serialization ──────────────────────────────────────────────────

function serializeSlideXml(slide: Slide): string {
  // Resolve every shape (groups included) by its `cNvPrId` so connectors
  // can find their anchored targets when emitting `<a:xfrm>`/start/end
  // points without re-walking the tree per connector.
  const shapesByCNvPrId = new Map<number, Shape>();
  collectShapesByCNvPrId(slide.shapes, shapesByCNvPrId);
  const spTreeChildren: unknown[] = [];
  for (const head of slide.spTreeHead) spTreeChildren.push(opaqueToEntry(head));
  for (const shape of slide.shapes) spTreeChildren.push(shapeToEntry(shape, shapesByCNvPrId));
  const spTree = makeEntry("p:spTree", spTreeChildren);

  const cSldChildren: unknown[] = [];
  for (const head of slide.cSldHead) cSldChildren.push(opaqueToEntry(head));
  cSldChildren.push(spTree);
  const cSld = makeEntry("p:cSld", cSldChildren, slide.cSldAttrs);

  const sldChildren: unknown[] = [cSld];
  for (const tail of slide.slideOpaqueTail) sldChildren.push(opaqueToEntry(tail));
  // F4: re-emit the typed transition (verbatim if untouched, rebuilt if
  // dirty) and the raw timing tail. Both live AFTER any other slide tail
  // children to mirror PowerPoint's element order: cSld → clrMapOvr →
  // transition → timing → extLst.
  if (slide.transition) sldChildren.push(transitionToEntry(slide.transition));
  // F4 surgical-merge:
  //  - With a captured `timingTailRaw` AND typed animations: walk the
  //    tail and splice in rebuilt `<p:par>` for edited animations
  //    while preserving every other tail node verbatim.
  //  - With a captured `timingTailRaw` but no typed animations: the
  //    user removed every typed animation; emit only the unmodelled
  //    pieces (sound effects, exit/emphasis siblings) by stripping
  //    typed-animation pars from the tail.
  //  - With no captured tail but typed animations: synthesise a
  //    minimal timing tree (new from scratch).
  //  - With neither: omit `<p:timing>`.
  if (slide.timingTailRaw) {
    sldChildren.push(mergeTimingFromAnimations(slide.timingTailRaw, slide.animations));
  } else if (slide.animations.length > 0) {
    sldChildren.push(timingFromAnimations(slide.animations));
  }
  const sld = makeEntry("p:sld", sldChildren, slide.slideRootAttrs);

  return ooxml.serializeXml([sld]);
}

function transitionToEntry(t: SlideTransition): Record<string, unknown> {
  if (t.raw) return opaqueToEntry(t.raw);
  const inner: unknown[] = [];
  if (t.kind !== "none" && t.kind !== "unsupported") {
    inner.push(makeEntry(`p:${t.kind}`, []));
  }
  return makeEntry("p:transition", inner, t.speed ? { spd: t.speed } : {});
}

// ─── F4 v2: Build a fresh <p:timing> tree from typed shape animations ───
//
// Mirrors the structure PowerPoint emits for the main timing sequence:
//
//   <p:timing>
//     <p:tnLst>
//       <p:par>
//         <p:cTn id=1 dur="indefinite" restart="never" nodeType="tmRoot">
//           <p:childTnLst>
//             <p:seq concurrent="1" nextAc="seek">
//               <p:cTn id=2 dur="indefinite" nodeType="mainSeq">
//                 <p:childTnLst>
//                   <p:par>          ← per typed animation
//                     <p:cTn id=N presetID=X presetClass="…" nodeType="clickEffect|withEffect|afterEffect">
//                       <p:childTnLst>
//                         …per-preset behaviour bodies (set / anim / animEffect / animRot / animScale / animMotion)…
//                       </p:childTnLst>
//                     </p:cTn>
//                   </p:par>
//                 </p:childTnLst>
//               </p:cTn>
//             </p:seq>
//           </p:childTnLst>
//         </p:cTn>
//       </p:par>
//     </p:tnLst>
//   </p:timing>
//
// The per-preset emitters live in `animation/presets.ts`; this function
// only owns the envelope (mainSeq / per-animation `<p:par>` / `<p:cTn>`
// attributes / trigger node types). When an animation carries a
// captured `raw` blob we re-emit it verbatim so byte-perfect
// round-trip survives unrelated edits.
function timingFromAnimations(animations: ReadonlyArray<ShapeAnimation>): Record<string, unknown> {
  let cTnId = 3; // 1 = tmRoot, 2 = mainSeq, ≥3 = per-animation cTn ids
  const animPars: unknown[] = [];
  for (const a of animations) {
    if (a.raw) {
      animPars.push(opaqueToEntry(a.raw));
      continue;
    }
    const spec = findPreset(a.category, a.preset);
    const animCTnId = cTnId++;
    const helpers = makeEmitHelpers(() => cTnId++);
    const body = spec
      ? spec.emitBody(a, helpers)
      : helpers.childTnLst([
          helpers.setAttr(a.targetCNvPrId, "style.visibility", "visible", a.durationMs ?? 1),
        ]);
    const presetClass = spec?.presetClass ?? "entr";
    const presetId = spec?.presetId ?? 1;
    const subtype = spec ? subtypeFor(spec, a.direction) : 0;
    const animCTnAttrs: Record<string, string> = {
      id: String(animCTnId),
      presetID: String(presetId),
      presetClass,
      presetSubtype: String(subtype),
      fill: "hold",
      nodeType: triggerNodeType(a.trigger ?? "onClick"),
    };
    if (a.durationMs !== undefined) animCTnAttrs.dur = String(a.durationMs);
    if (a.delayMs !== undefined) animCTnAttrs.delay = String(a.delayMs);
    animPars.push(makeEntry("p:par", [makeEntry("p:cTn", [body], animCTnAttrs)]));
  }
  const seq = makeEntry(
    "p:seq",
    [
      makeEntry("p:cTn", [makeEntry("p:childTnLst", animPars)], {
        id: "2",
        dur: "indefinite",
        nodeType: "mainSeq",
      }),
    ],
    { concurrent: "1", nextAc: "seek" }
  );
  const tmRoot = makeEntry("p:par", [
    makeEntry("p:cTn", [makeEntry("p:childTnLst", [seq])], {
      id: "1",
      dur: "indefinite",
      restart: "never",
      nodeType: "tmRoot",
    }),
  ]);
  return makeEntry("p:timing", [makeEntry("p:tnLst", [tmRoot])]);
}

function triggerNodeType(trigger: ShapeAnimation["trigger"]): string {
  switch (trigger) {
    case "withPrevious":
      return "withEffect";
    case "afterPrevious":
      return "afterEffect";
    case "onClick":
      return "clickEffect";
    default: {
      const _exhaustive: never = trigger;
      void _exhaustive;
      return "clickEffect";
    }
  }
}

// ─── F4 v3: surgical merge of typed animations into captured tail ───
//
// The parser captures the entire `<p:timing>` blob as
// `slide.timingTailRaw`. Each typed animation also captures its own
// `<p:par>` carrier as `a.raw`. When the user runs an animation
// command we now keep the tail (we used to drop it — that lost
// every unmodelled effect on the slide). The serializer reconciles:
//
// - For each typed animation that **was edited** (`a.raw` is
//   undefined): rebuild a fresh `<p:par>` carrier via
//   `buildAnimationPar` and splice it where the original lived.
// - For each typed animation that **was preserved** (`a.raw` is
//   defined): emit `a.raw` verbatim — same effect as the previous
//   "happy path" but now also when other animations on the slide
//   were edited.
// - For each `<p:par>` carrier in the tail that doesn't correspond
//   to any typed animation (an unmodelled exit / emphasis effect
//   we kept opaque): leave it untouched.
// - For typed animations that have no matching carrier in the tail
//   (newly added): append them at the end of the mainSeq's
//   `<p:childTnLst>`.
//
// The walk identifies a "typed-animation carrier" as a `<p:par>`
// whose first `<p:cTn>` declares `presetClass` + `presetID` — the
// same predicate the parser uses. This keeps the merge trivially
// symmetric with parsing.
function mergeTimingFromAnimations(
  tail: OpaqueXml,
  animations: ReadonlyArray<ShapeAnimation>
): Record<string, unknown> {
  // Track which typed animations have already been emitted (by
  // identity) so they appear exactly once: either spliced into a
  // `<p:childTnLst>` that already had a matching carrier, or
  // appended into the deepest mainSeq if no carrier matched.
  const consumed = new Set<ShapeAnimation>();

  // Walk the tail, looking for `<p:childTnLst>` containers whose
  // children include at least one typed-animation `<p:par>` carrier.
  // For those containers, re-emit the carriers in `slide.animations`
  // order (which honours user-driven reordering) while preserving
  // every non-carrier sibling at its original position.
  const rewrite = (nodes: ReadonlyArray<unknown>): unknown[] => {
    const out: unknown[] = [];
    for (const node of nodes) {
      if (!node || typeof node !== "object" || Array.isArray(node)) {
        out.push(node);
        continue;
      }
      const e = node as Record<string, unknown>;
      const tag = entryTag(e);
      if (!tag) {
        out.push(node);
        continue;
      }

      if (tag === "p:childTnLst") {
        const children = (e[tag] as unknown[] | undefined) ?? [];
        const rewrittenChildren = rewriteChildTnLstChildren(children, animations, consumed, rewrite);
        out.push(rebuildEntry(e, tag, rewrittenChildren));
        continue;
      }

      const children = e[tag] as unknown[] | undefined;
      if (Array.isArray(children)) {
        const nextChildren = rewrite(children);
        out.push(rebuildEntry(e, tag, nextChildren));
      } else {
        out.push(node);
      }
    }
    return out;
  };

  const rewritten = rewrite(tail.subtree);

  // Append any typed animations that had no matching carrier (newly
  // added since parse). We splice them into the deepest reachable
  // mainSeq `<p:childTnLst>`. If we can't find one (synthetic tail),
  // fall back to appending a fresh mainSeq tree at the top level —
  // this mirrors `timingFromAnimations`.
  const newPars: unknown[] = [];
  for (const a of animations) {
    if (consumed.has(a)) continue;
    newPars.push(a.raw ? opaqueToEntry(a.raw) : buildAnimationPar(a));
  }
  let withAppended: unknown[];
  if (newPars.length > 0) {
    const appended = appendIntoMainSeq(rewritten, newPars);
    withAppended = appended.ok ? appended.subtree : [...rewritten, ...newPars];
  } else {
    withAppended = rewritten;
  }

  // Re-wrap as `<p:timing>` with the original attrs.
  const attrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(tail.rawAttrs)) attrs[k] = v;
  const entry: Record<string, unknown> = { [tail.tag]: withAppended };
  if (Object.keys(attrs).length > 0) entry[ATTR_KEY] = attrs;
  return entry;
}

/**
 * Rewrite the children of a `<p:childTnLst>` from a captured timing
 * tail.
 *
 * The container may hold a mix of typed-animation `<p:par>` carriers
 * (matched against `slide.animations` by `(spid, presetClass,
 * presetId)`) and arbitrary unmodelled siblings (sound effects,
 * complex sub-sequences, etc.). The contract:
 *
 *   - Non-carrier siblings keep their original positions verbatim.
 *   - The slot where the first carrier appeared in tail order
 *     becomes the anchor for ALL typed animations matching this
 *     container, emitted in `slide.animations` order. This makes
 *     reorder commands actually re-order on disk while preserving
 *     non-carrier siblings.
 *   - Carriers whose typed counterpart was deleted are dropped.
 *   - Typed animations not represented in this container are
 *     deferred to the caller (which may append them into the
 *     deepest mainSeq via `appendIntoMainSeq`).
 *
 * `consumed` tracks identity of typed animations emitted into any
 * container so the caller's "append leftover" pass doesn't double-
 * emit them.
 */
function rewriteChildTnLstChildren(
  children: ReadonlyArray<unknown>,
  animations: ReadonlyArray<ShapeAnimation>,
  consumed: Set<ShapeAnimation>,
  rewriteOther: (nodes: ReadonlyArray<unknown>) => unknown[]
): unknown[] {
  // First pass — find which carriers live in this list and which
  // typed animations they correspond to (by appearance order). The
  // carrier→typed-animation association is stable across reorders
  // because we re-derive it from the carrier key.
  type CarrierSlot = { kind: "carrier"; key: string };
  type SiblingSlot = { kind: "sibling"; node: unknown };
  const slots: Array<CarrierSlot | SiblingSlot> = [];
  const carrierKeyOrder: string[] = [];
  for (const c of children) {
    if (c && typeof c === "object" && !Array.isArray(c)) {
      const e = c as Record<string, unknown>;
      const tag = entryTag(e);
      if (tag === "p:par") {
        const carrier = identifyCarrier(e);
        if (carrier) {
          const key = `${carrier.spid}|${carrier.presetClass}|${carrier.presetId}`;
          slots.push({ kind: "carrier", key });
          if (!carrierKeyOrder.includes(key)) carrierKeyOrder.push(key);
          continue;
        }
      }
    }
    slots.push({ kind: "sibling", node: c });
  }
  // No carriers in this list → recurse normally so nested mainSeq
  // levels still get visited.
  if (carrierKeyOrder.length === 0) {
    return rewriteOther(children);
  }
  // Bucket typed animations by carrier key so we can decide which
  // ones belong to THIS childTnLst (vs a nested one). Only typed
  // animations whose carrier key is also represented in this list
  // are emitted here; the rest fall through to outer containers /
  // the leftover-append pass.
  const keysInThisList = new Set(carrierKeyOrder);
  const typedHere: ShapeAnimation[] = [];
  for (const a of animations) {
    if (consumed.has(a)) continue;
    const spec = findPreset(a.category, a.preset);
    if (!spec) continue;
    const key = `${a.targetCNvPrId}|${spec.presetClass}|${spec.presetId}`;
    if (!keysInThisList.has(key)) continue;
    typedHere.push(a);
    consumed.add(a);
  }
  // Emit: anchor the typed-animation block at the position of the
  // first carrier slot. Subsequent carrier slots collapse (the
  // typed animations are already laid out at the anchor). Non-
  // carrier siblings keep their positions.
  const out: unknown[] = [];
  let anchored = false;
  for (const slot of slots) {
    if (slot.kind === "sibling") {
      out.push(slot.node);
      continue;
    }
    if (anchored) continue;
    for (const a of typedHere) {
      out.push(a.raw ? opaqueToEntry(a.raw) : buildAnimationPar(a));
    }
    anchored = true;
  }
  return out;
}

function entryTag(entry: Record<string, unknown>): string | undefined {
  for (const k of Object.keys(entry)) {
    if (k === ATTR_KEY) continue;
    return k;
  }
  return undefined;
}

function rebuildEntry(e: Record<string, unknown>, tag: string, children: unknown[]): Record<string, unknown> {
  const out: Record<string, unknown> = { [tag]: children };
  const a = e[ATTR_KEY];
  if (a !== undefined) out[ATTR_KEY] = a;
  return out;
}

function identifyCarrier(
  par: Record<string, unknown>
): { spid: number; presetClass: string; presetId: number } | null {
  const tag = entryTag(par);
  if (!tag || tag !== "p:par") return null;
  const children = par[tag] as unknown[] | undefined;
  if (!Array.isArray(children)) return null;
  for (const c of children) {
    if (!c || typeof c !== "object") continue;
    const ce = c as Record<string, unknown>;
    if (entryTag(ce) !== "p:cTn") continue;
    const attrs = ce[ATTR_KEY] as Record<string, string> | undefined;
    const presetClass = attrs?.[`${ATTR_PREFIX}presetClass`];
    const presetIdStr = attrs?.[`${ATTR_PREFIX}presetID`];
    if (!presetClass || !presetIdStr) return null;
    const presetId = Number(presetIdStr);
    const spid = findSpTgtSpidEntry(ce);
    if (spid === null) return null;
    return { spid, presetClass, presetId };
  }
  return null;
}

function findSpTgtSpidEntry(node: unknown): number | null {
  if (!node || typeof node !== "object") return null;
  const e = node as Record<string, unknown>;
  const tag = entryTag(e);
  if (!tag) return null;
  if (tag === "p:spTgt") {
    const attrs = e[ATTR_KEY] as Record<string, string> | undefined;
    const v = attrs?.[`${ATTR_PREFIX}spid`];
    if (v && /^\d+$/.test(v)) return Number(v);
    return null;
  }
  const children = e[tag] as unknown[] | undefined;
  if (Array.isArray(children)) {
    for (const c of children) {
      const r = findSpTgtSpidEntry(c);
      if (r !== null) return r;
    }
  }
  return null;
}

/**
 * Append additional `<p:par>` carriers into the deepest mainSeq's
 * `<p:childTnLst>`. Returns `{ ok: false }` if no mainSeq was
 * found, in which case the caller falls back to appending at the
 * top level.
 */
function appendIntoMainSeq(
  subtree: ReadonlyArray<unknown>,
  newPars: ReadonlyArray<unknown>
): { ok: true; subtree: unknown[] } | { ok: false } {
  let found = false;
  const visit = (nodes: ReadonlyArray<unknown>): unknown[] => {
    return nodes.map((node) => {
      if (!node || typeof node !== "object" || Array.isArray(node)) return node;
      const e = node as Record<string, unknown>;
      const tag = entryTag(e);
      if (!tag) return node;
      if (tag === "p:cTn") {
        const attrs = e[ATTR_KEY] as Record<string, string> | undefined;
        const nodeType = attrs?.[`${ATTR_PREFIX}nodeType`];
        if (nodeType === "mainSeq") {
          // Inject into this cTn's `<p:childTnLst>` (last one wins
          // if multiple — there usually is only one).
          const children = (e[tag] as unknown[] | undefined) ?? [];
          const nextChildren: unknown[] = [];
          let appended = false;
          for (const c of children) {
            if (!c || typeof c !== "object") {
              nextChildren.push(c);
              continue;
            }
            const ce = c as Record<string, unknown>;
            if (entryTag(ce) === "p:childTnLst" && !appended) {
              const lst = (ce["p:childTnLst"] as unknown[] | undefined) ?? [];
              nextChildren.push({ "p:childTnLst": [...lst, ...newPars] });
              appended = true;
            } else {
              nextChildren.push(c);
            }
          }
          if (!appended) {
            nextChildren.push({ "p:childTnLst": [...newPars] });
          }
          found = true;
          return rebuildEntry(e, tag, nextChildren);
        }
      }
      const children = e[tag] as unknown[] | undefined;
      if (Array.isArray(children)) {
        return rebuildEntry(e, tag, visit(children));
      }
      return node;
    });
  };
  const out = visit(subtree);
  return found ? { ok: true, subtree: out } : { ok: false };
}

/**
 * Build a single typed animation as a `<p:par>` carrier — the same
 * shape `timingFromAnimations` would build, but for one animation.
 * Used by `mergeTimingFromAnimations` to splice a rebuilt carrier
 * back into the captured tail without rebuilding the whole tree.
 */
function buildAnimationPar(a: ShapeAnimation): Record<string, unknown> {
  let cTnId = 1_000_000; // out of range of typical mainSeq ids; merge is best-effort on uniqueness here
  const helpers = makeEmitHelpers(() => cTnId++);
  const spec = findPreset(a.category, a.preset);
  const body = spec
    ? spec.emitBody(a, helpers)
    : helpers.childTnLst([
        helpers.setAttr(a.targetCNvPrId, "style.visibility", "visible", a.durationMs ?? 1),
      ]);
  const presetClass = spec?.presetClass ?? "entr";
  const presetId = spec?.presetId ?? 1;
  const subtype = spec ? subtypeFor(spec, a.direction) : 0;
  const animCTnAttrs: Record<string, string> = {
    id: String(cTnId++),
    presetID: String(presetId),
    presetClass,
    presetSubtype: String(subtype),
    fill: "hold",
    nodeType: triggerNodeType(a.trigger ?? "onClick"),
  };
  if (a.durationMs !== undefined) animCTnAttrs.dur = String(a.durationMs);
  if (a.delayMs !== undefined) animCTnAttrs.delay = String(a.delayMs);
  return makeEntry("p:par", [makeEntry("p:cTn", [body], animCTnAttrs)]);
}

/**
 * Build an `EmitHelpers` bag for a single animation. The shared `id`
 * counter is passed in so behaviour-level `<p:cTn>` ids stay unique
 * across helpers within the same animation.
 */
function makeEmitHelpers(nextId: () => number): EmitHelpers {
  const cBhvr = (
    spid: number,
    durMs: number,
    attrName: string | null,
    extra?: Record<string, string>
  ): unknown => {
    const attrs: Record<string, string> = {
      id: String(nextId()),
      dur: String(Math.max(1, durMs)),
      fill: "hold",
      ...(extra ?? {}),
    };
    const children: unknown[] = [makeEntry("p:cTn", [], attrs)];
    children.push(makeEntry("p:tgtEl", [makeEntry("p:spTgt", [], { spid: String(spid) })]));
    if (attrName) {
      children.push(makeEntry("p:attrNameLst", [makeEntry("p:attrName", [{ "#text": attrName }])]));
    }
    return makeEntry("p:cBhvr", children);
  };

  return {
    setAttr: (spid, attrName, value, durMs) =>
      makeEntry("p:set", [
        cBhvr(spid, durMs, attrName),
        makeEntry("p:to", [makeEntry("p:strVal", [], { val: value })]),
      ]),
    anim: (spid, attrName, durMs, fromTo, extra) => {
      const tavLst = makeEntry("p:tavLst", [
        ...(fromTo.from !== undefined
          ? [
              makeEntry("p:tav", [makeEntry("p:val", [makeEntry("p:strVal", [], { val: fromTo.from })])], {
                tm: "0",
              }),
            ]
          : []),
        makeEntry("p:tav", [makeEntry("p:val", [makeEntry("p:strVal", [], { val: fromTo.to })])], {
          tm: "100000",
        }),
      ]);
      const animAttrs: Record<string, string> = { calcmode: "lin", valueType: "num" };
      if (extra) Object.assign(animAttrs, extra);
      return makeEntry("p:anim", [cBhvr(spid, durMs, attrName), tavLst], animAttrs);
    },
    animEffect: (spid, durMs, transition, filter, extra) => {
      const attrs: Record<string, string> = { transition, filter };
      if (extra) Object.assign(attrs, extra);
      return makeEntry("p:animEffect", [cBhvr(spid, durMs, null)], attrs);
    },
    animRot: (spid, durMs, fromDeg, toDeg, extra) => {
      const attrs: Record<string, string> = {
        by: String(Math.round((toDeg - fromDeg) * 60000)),
        from: String(Math.round(fromDeg * 60000)),
        to: String(Math.round(toDeg * 60000)),
      };
      if (extra) Object.assign(attrs, extra);
      return makeEntry("p:animRot", [cBhvr(spid, durMs, null)], attrs);
    },
    animScale: (spid, durMs, from, to, extra) => {
      const attrs: Record<string, string> = {};
      if (extra) Object.assign(attrs, extra);
      return makeEntry(
        "p:animScale",
        [
          cBhvr(spid, durMs, null),
          makeEntry("p:from", [], { x: String(from.x * 1000), y: String(from.y * 1000) }),
          makeEntry("p:to", [], { x: String(to.x * 1000), y: String(to.y * 1000) }),
        ],
        attrs
      );
    },
    animMotion: (spid, durMs, path, extra) => {
      const attrs: Record<string, string> = {
        origin: "layout",
        path,
        pathEditMode: "relative",
      };
      if (extra) Object.assign(attrs, extra);
      return makeEntry("p:animMotion", [cBhvr(spid, durMs, null)], attrs);
    },
    childTnLst: (children) => makeEntry("p:childTnLst", children),
  };
}

function shapeToEntry(shape: Shape, shapesByCNvPrId: ReadonlyMap<number, Shape>): Record<string, unknown> {
  switch (shape.kind) {
    case "text":
      return textShapeToEntry(shape);
    case "pic":
      return pictureToEntry(shape);
    case "media":
      return mediaShapeToEntry(shape);
    case "group":
      return groupShapeToEntry(shape, shapesByCNvPrId);
    case "table":
      return tableShapeToEntry(shape);
    case "chart":
      return chartShapeToEntry(shape);
    case "ole-spreadsheet":
      return oleSpreadsheetShapeToEntry(shape);
    case "connector":
      return connectorToEntry(shape, shapesByCNvPrId);
    case "opaque":
      return opaqueShapeToEntry(shape);
  }
}

function collectShapesByCNvPrId(shapes: ReadonlyArray<Shape>, out: Map<number, Shape>): void {
  for (const s of shapes) {
    if (s.cNvPrId > 0) out.set(s.cNvPrId, s);
    if (s.kind === "group") collectShapesByCNvPrId(s.children, out);
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
    nvSpPrChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
  }

  const nvSpPr = makeEntry("p:nvSpPr", nvSpPrChildren);

  const spPrChildren: unknown[] = [];
  let emittedXfrm = false;
  for (const o of shape.spPrTail) {
    if (o.tag === "a:xfrm") {
      spPrChildren.push(buildXfrm(shape.position, shape.size, o, shape.rotation));
      emittedXfrm = true;
    } else {
      spPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedXfrm && (shape.position || shape.size || shape.rotation !== undefined)) {
    spPrChildren.unshift(buildXfrm(shape.position, shape.size, undefined, shape.rotation));
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
    nvPicPrChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
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
      spPrChildren.push(buildXfrm(shape.position, shape.size, o, shape.rotation));
      emittedXfrm = true;
    } else {
      spPrChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedXfrm && (shape.position || shape.size || shape.rotation !== undefined)) {
    spPrChildren.unshift(buildXfrm(shape.position, shape.size, undefined, shape.rotation));
  }
  const spPr = makeEntry("p:spPr", spPrChildren);

  const children: unknown[] = [nvPicPr, blipFill, spPr];
  if (shape.styleRaw) children.push(opaqueToEntry(shape.styleRaw));
  return makeEntry("p:pic", children);
}

function groupShapeToEntry(
  shape: GroupShape,
  shapesByCNvPrId: ReadonlyMap<number, Shape>
): Record<string, unknown> {
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
    nvChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
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
  for (const c of shape.children) children.push(shapeToEntry(c, shapesByCNvPrId));
  return makeEntry("p:grpSp", children);
}

function opaqueShapeToEntry(shape: OpaqueShape): Record<string, unknown> {
  return opaqueToEntry(shape.raw);
}

// ─── Connector serialization ──────────────────────────────────────────────

const CONNECTOR_PRST_BY_TYPE: Readonly<Record<ConnectorShape["connectorType"], string>> = {
  straight: "line",
  elbow: "bentConnector3",
  curved: "curvedConnector3",
  // `unsupported` should never reach the serializer for a freshly authored
  // connector — when round-tripping a parsed connector we still need a
  // legal preset, so fall back to a straight line which renders as a
  // single segment.
  unsupported: "line",
};

const SIDE_TO_CXN_IDX: Readonly<Record<ConnectorSide, string>> = {
  n: "0",
  e: "1",
  s: "2",
  w: "3",
  center: "0",
};

function connectorToEntry(
  shape: ConnectorShape,
  shapesByCNvPrId: ReadonlyMap<number, Shape>
): Record<string, unknown> {
  // ── p:nvCxnSpPr ─────────────────────────────────────────────────────
  // We rebuild p:cNvPr from the model and rebuild p:cNvCxnSpPr from the
  // typed start/end so anchor edits round-trip without us having to
  // reach into opaque XML. p:nvPr (and any captured tail children we
  // don't recognise) pass through verbatim.
  const nvChildren: unknown[] = [];
  let emittedCNvPr = false;
  let emittedCNvCxnSpPr = false;
  for (const o of shape.nvCxnSpPrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else if (o.tag === "p:cNvCxnSpPr" && !emittedCNvCxnSpPr) {
      nvChildren.push(rebuildCNvCxnSpPr(shape, o));
      emittedCNvCxnSpPr = true;
    } else {
      nvChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
  }
  if (!emittedCNvCxnSpPr) {
    // Insert directly after p:cNvPr; PowerPoint mandates the order
    // p:cNvPr → p:cNvCxnSpPr → p:nvPr inside p:nvCxnSpPr.
    nvChildren.splice(1, 0, rebuildCNvCxnSpPr(shape, undefined));
  }
  if (!nvChildren.some((n) => isTag(n, "p:nvPr"))) {
    nvChildren.push(makeEntry("p:nvPr", []));
  }
  const nvCxnSpPr = makeEntry("p:nvCxnSpPr", nvChildren);

  // ── p:spPr (xfrm + prstGeom + tail) ─────────────────────────────────
  const startPt = resolveEndpoint(shape.start, shapesByCNvPrId);
  const endPt = resolveEndpoint(shape.end, shapesByCNvPrId);
  // Fallback when an anchored target was deleted: re-use whatever the
  // shape's stored bounding box says so we don't crash.
  const fallbackStart = {
    x: shape.position?.xEmu ?? 0,
    y: shape.position?.yEmu ?? 0,
  };
  const fallbackEnd = {
    x: (shape.position?.xEmu ?? 0) + (shape.size?.cxEmu ?? 0),
    y: (shape.position?.yEmu ?? 0) + (shape.size?.cyEmu ?? 0),
  };
  const xfrmInfo = connectorXfrm(startPt ?? fallbackStart, endPt ?? fallbackEnd);
  const xfrmAttrs: Record<string, string> = {};
  if (xfrmInfo.flipH) xfrmAttrs.flipH = "1";
  if (xfrmInfo.flipV) xfrmAttrs.flipV = "1";
  const xfrmEntry: Record<string, unknown> = {
    "a:xfrm": [
      makeEntry("a:off", [], { x: String(xfrmInfo.box.x), y: String(xfrmInfo.box.y) }),
      makeEntry("a:ext", [], { cx: String(xfrmInfo.box.cx), cy: String(xfrmInfo.box.cy) }),
    ],
  };
  if (Object.keys(xfrmAttrs).length > 0) xfrmEntry[ATTR_KEY] = makeRawAttrs(xfrmAttrs);

  const prstGeom = makeEntry("a:prstGeom", [makeEntry("a:avLst", [])], {
    prst: CONNECTOR_PRST_BY_TYPE[shape.connectorType],
  });

  const spPrChildren: unknown[] = [xfrmEntry, prstGeom];
  if (shape.stroke || shape.headEnd || shape.tailEnd) {
    spPrChildren.push(buildConnectorLn(shape));
  }
  for (const o of shape.spPrTail) {
    // The parser captured everything inside <p:spPr> (including the
    // a:xfrm/a:prstGeom/a:ln we now rebuild typed), so skip those tags
    // when re-emitting the tail to avoid duplicates.
    if (o.tag === "a:xfrm" || o.tag === "a:prstGeom") continue;
    if (o.tag === "a:ln" && (shape.stroke || shape.headEnd || shape.tailEnd)) continue;
    spPrChildren.push(opaqueToEntry(o));
  }
  const spPr = makeEntry("p:spPr", spPrChildren);

  return makeEntry("p:cxnSp", [nvCxnSpPr, spPr]);
}

function isTag(node: unknown, tag: string): boolean {
  if (!node || typeof node !== "object" || Array.isArray(node)) return false;
  return ooxml.getTag(node as Record<string, unknown>) === tag;
}

function rebuildCNvCxnSpPr(shape: ConnectorShape, captured: OpaqueXml | undefined): Record<string, unknown> {
  const children: unknown[] = [];
  if (shape.start.kind === "anchored") {
    children.push(buildCxnEndpointEntry("a:stCxn", shape.start));
  }
  if (shape.end.kind === "anchored") {
    children.push(buildCxnEndpointEntry("a:endCxn", shape.end));
  }
  // Pass through anything the parser stashed inside p:cNvCxnSpPr that
  // isn't a stCxn/endCxn (e.g. a:extLst). When `captured` is `undefined`
  // we emit a minimal element — that's the fresh-add case.
  if (captured) {
    for (const sub of captured.subtree) {
      if (!sub || typeof sub !== "object" || Array.isArray(sub)) continue;
      const tag = ooxml.getTag(sub as Record<string, unknown>);
      if (tag === "a:stCxn" || tag === "a:endCxn") continue;
      children.push(sub);
    }
  }
  const entry: Record<string, unknown> = { "p:cNvCxnSpPr": children };
  if (captured && Object.keys(captured.rawAttrs).length > 0) {
    entry[ATTR_KEY] = { ...captured.rawAttrs };
  }
  return entry;
}

function buildCxnEndpointEntry(
  tag: "a:stCxn" | "a:endCxn",
  ep: Extract<ConnectorEndpoint, { kind: "anchored" }>
): Record<string, unknown> {
  return makeEntry(tag, [], {
    id: String(ep.targetCNvPrId),
    idx: SIDE_TO_CXN_IDX[ep.side],
  });
}

function buildConnectorLn(shape: ConnectorShape): Record<string, unknown> {
  const lnAttrs: Record<string, string> = {};
  if (shape.stroke && shape.stroke.widthEmu > 0) {
    lnAttrs.w = String(shape.stroke.widthEmu);
  }
  const lnChildren: unknown[] = [];
  if (shape.stroke) {
    // Prefer theme color references over literal hex when both are
    // present (theme refs are how PowerPoint expresses "follow the
    // theme palette" — collapsing to the literal RGB would freeze
    // the rendered color and break theme switches).
    if (shape.stroke.colorTheme) {
      lnChildren.push(
        makeEntry("a:solidFill", [makeEntry("a:schemeClr", [], { val: shape.stroke.colorTheme })])
      );
    } else {
      lnChildren.push(makeEntry("a:solidFill", [makeEntry("a:srgbClr", [], { val: shape.stroke.color })]));
    }
  }
  if (shape.stroke?.dash && shape.stroke.dash !== "solid") {
    // PowerPoint canonicalises the dash via `<a:prstDash>` after the
    // fill but before the head/tail end markers — we follow that order
    // so re-saved files diff cleanly against authentic PowerPoint
    // output.
    lnChildren.push(makeEntry("a:prstDash", [], { val: prstDashValue(shape.stroke.dash) }));
  }
  if (shape.headEnd) {
    lnChildren.push(makeEntry("a:headEnd", [], { type: shape.headEnd }));
  }
  if (shape.tailEnd) {
    lnChildren.push(makeEntry("a:tailEnd", [], { type: shape.tailEnd }));
  }
  return makeEntry("a:ln", lnChildren, lnAttrs);
}

function prstDashValue(dash: Exclude<ConnectorDashStyle, "solid">): string {
  // OOXML `ST_PresetLineDashVal` round-trip table. New canonical
  // tokens map to themselves; the legacy short aliases are folded
  // back to their canonical token so we don't emit non-spec values.
  switch (dash) {
    case "dot":
    case "dash":
    case "lgDash":
    case "dashDot":
    case "lgDashDot":
    case "lgDashDotDot":
    case "sysDash":
    case "sysDot":
    case "sysDashDot":
    case "sysDashDotDot":
      return dash;
    case "dashed":
      return "dash";
    case "dotted":
      return "dot";
    case "longDash":
      return "lgDash";
  }
}

function tableShapeToEntry(shape: TableShape): Record<string, unknown> {
  const nvChildren: unknown[] = [];
  let emittedCNvPr = false;
  for (const o of shape.nvGraphicFramePrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else {
      nvChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
  }
  const nvGraphicFramePr = makeEntry("p:nvGraphicFramePr", nvChildren);

  const xfrmChildren: unknown[] = [];
  if (shape.position) {
    xfrmChildren.push(
      makeEntry("a:off", [], {
        x: String(shape.position.xEmu),
        y: String(shape.position.yEmu),
      })
    );
  }
  if (shape.size) {
    xfrmChildren.push(
      makeEntry("a:ext", [], {
        cx: String(shape.size.cxEmu),
        cy: String(shape.size.cyEmu),
      })
    );
  }
  const xfrm = makeEntry("p:xfrm", xfrmChildren);

  // <a:tbl>
  const tblChildren: unknown[] = [];
  if (shape.tblPrRaw) tblChildren.push(opaqueToEntry(shape.tblPrRaw));
  else tblChildren.push(makeEntry("a:tblPr", []));

  const gridChildren: unknown[] = shape.columnWidths.map((w) => makeEntry("a:gridCol", [], { w: String(w) }));
  tblChildren.push(makeEntry("a:tblGrid", gridChildren));

  for (const row of shape.rows) {
    const trChildrenOut: unknown[] = [];
    for (const cell of row.cells) {
      const tcChildrenOut: unknown[] = [];
      tcChildrenOut.push(textBodyToEntryWith("a:txBody", cell.txBody));
      if (cell.tcPrRaw) tcChildrenOut.push(opaqueToEntry(cell.tcPrRaw));
      const tcEntry: Record<string, unknown> = { "a:tc": tcChildrenOut };
      if (Object.keys(cell.tcAttrs).length > 0) {
        tcEntry[ATTR_KEY] = makeRawAttrs(cell.tcAttrs);
      }
      trChildrenOut.push(tcEntry);
    }
    const trAttrs: Record<string, string> = { ...row.trAttrs, h: String(row.height) };
    const trEntry: Record<string, unknown> = { "a:tr": trChildrenOut };
    trEntry[ATTR_KEY] = makeRawAttrs(trAttrs);
    tblChildren.push(trEntry);
  }
  const tbl = makeEntry("a:tbl", tblChildren);

  const graphicData = makeEntry("a:graphicData", [tbl], { uri: shape.graphicDataUri });
  const graphic = makeEntry("a:graphic", [graphicData]);

  return makeEntry("p:graphicFrame", [nvGraphicFramePr, xfrm, graphic]);
}

// ─── Chart parts (F3) ─────────────────────────────────────────────────────

/**
 * Serialize a `ChartPart` back to XML. The strategy is dirty-flag-driven:
 *  - Unmodified parts are never re-serialized; the container's original
 *    bytes are preserved, guaranteeing byte-roundtrip for charts.
 *  - When a chart is dirtied (after a chart command) we rebuild the
 *    chart XML by replacing only the typed children inside `<c:chart>`
 *    and `<c:plotArea>`, splicing them back into the verbatim
 *    `<c:chartSpace>` subtree to keep axes, legend, embedded xlsx, and
 *    other unmodeled bits intact.
 */
/**
 * Serialise a `SlideLayout`'s opaque blob back into a `<p:sldLayout>`
 * XML document. We don't introspect the placeholder shapes when
 * writing — the parser captured every child verbatim, and built-in
 * layouts arrive with their own pre-authored XML.
 */
function serializeCommentsXml(part: import("../model/types.js").PptxCommentsPart): string {
  const cmEntries: unknown[] = [];
  for (const c of part.comments) {
    const attrs: Record<string, string> = {
      "@_authorId": String(c.authorId),
      "@_idx": String(c.idx),
    };
    if (c.createdAt) attrs["@_dt"] = c.createdAt;
    const cmChildren: unknown[] = [
      {
        "p:pos": [],
        ":@": {
          "@_x": String(Math.round(c.xEmu / 127)),
          "@_y": String(Math.round(c.yEmu / 127)),
        },
      },
      { "p:text": [{ "#text": c.text }] },
    ];
    const ext: unknown[] = [];
    if (c.parentId) {
      ext.push({
        "p:ext": [],
        ":@": { "@_uri": "officeai:parent", "@_id": c.parentId },
      });
    }
    if (c.resolved !== undefined) {
      ext.push({
        "p:ext": [],
        ":@": { "@_uri": "officeai:resolved", "@_value": c.resolved ? "1" : "0" },
      });
    }
    if (c.shapeId) {
      ext.push({
        "p:ext": [],
        ":@": { "@_uri": "officeai:shapeAnchor", "@_id": c.shapeId },
      });
    }
    if (ext.length > 0) cmChildren.push({ "p:extLst": ext });
    cmEntries.push({ "p:cm": cmChildren, ":@": attrs });
  }
  const root: Record<string, unknown> = {
    "p:cmLst": cmEntries,
    ":@": {
      "@_xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@_xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "@_xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    },
  };
  return ooxml.serializeXml([root]);
}

function serializeCommentAuthorsXml(part: import("../model/types.js").PptxCommentAuthorsPart): string {
  const authors: unknown[] = part.authors.map((a) => {
    const attrs: Record<string, string> = {
      "@_id": String(a.id),
      "@_name": a.name,
    };
    if (a.initials) attrs["@_initials"] = a.initials;
    if (a.lastIdx !== undefined) attrs["@_lastIdx"] = String(a.lastIdx);
    if (a.clrIdx !== undefined) attrs["@_clrIdx"] = String(a.clrIdx);
    return { "p:cmAuthor": [], ":@": attrs };
  });
  const root: Record<string, unknown> = {
    "p:cmAuthorLst": authors,
    ":@": {
      "@_xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "@_xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    },
  };
  return ooxml.serializeXml([root]);
}

function serializeLayoutXml(raw: import("../model/types.js").OpaqueXml): string {
  const root: Record<string, unknown> = { [raw.tag]: raw.subtree };
  if (Object.keys(raw.rawAttrs).length > 0) {
    root[ATTR_KEY] = { ...raw.rawAttrs };
  }
  return ooxml.serializeXml([root]);
}

interface ChartSerializeContext {
  readonly embeddingRelId: string;
  readonly categoryRef: string;
  readonly valueRefs: ReadonlyArray<string>;
  readonly nameRefs: ReadonlyArray<string>;
}

function serializeChartPartXml(part: ChartPart, ctx?: ChartSerializeContext): string {
  const subtree = part.chartSpaceRaw.subtree as unknown[];
  const out: unknown[] = [];
  let externalDataEmitted = false;
  for (const node of subtree) {
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      out.push(node);
      continue;
    }
    const obj = node as Record<string, unknown>;
    const tag = ooxml.getTag(obj);
    if (tag === "c:chart") {
      out.push(rebuildChartElement(part, obj, ctx));
    } else if (tag === "c:externalData" && ctx) {
      out.push(rebuildExternalData(ctx.embeddingRelId));
      externalDataEmitted = true;
    } else {
      out.push(obj);
    }
  }
  if (ctx && !externalDataEmitted) {
    out.push(rebuildExternalData(ctx.embeddingRelId));
  }
  const chartSpace: Record<string, unknown> = { "c:chartSpace": out };
  if (Object.keys(part.chartSpaceRaw.rawAttrs).length > 0) {
    chartSpace[ATTR_KEY] = { ...part.chartSpaceRaw.rawAttrs };
  } else {
    // Synthetic chart parts may have no captured attrs; emit canonical
    // namespace declarations so the XML validates against the schema
    // referenced by Office.
    chartSpace[ATTR_KEY] = makeRawAttrs({
      "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    });
  }
  return ooxml.serializeXml([chartSpace]);
}

function rebuildExternalData(relId: string): Record<string, unknown> {
  const auto = makeEntry("c:autoUpdate", [], { val: "0" });
  const entry: Record<string, unknown> = { "c:externalData": [auto] };
  entry[ATTR_KEY] = makeRawAttrs({ "r:id": relId });
  return entry;
}

function rebuildChartElement(
  part: ChartPart,
  chart: Record<string, unknown>,
  ctx?: ChartSerializeContext
): Record<string, unknown> {
  const chartChildren = (chart["c:chart"] as unknown[] | undefined) ?? [];
  const out: unknown[] = [];
  let titleEmitted = false;
  let plotAreaEmitted = false;
  for (const node of chartChildren) {
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      out.push(node);
      continue;
    }
    const obj = node as Record<string, unknown>;
    const tag = ooxml.getTag(obj);
    if (tag === "c:title" && !titleEmitted) {
      titleEmitted = true;
      if (part.title !== undefined) out.push(rebuildChartTitle(part.title));
      // If title is undefined we drop the element.
      continue;
    }
    if (tag === "c:plotArea") {
      out.push(rebuildPlotArea(part, ctx));
      plotAreaEmitted = true;
      continue;
    }
    out.push(obj);
  }
  if (!titleEmitted && part.title !== undefined) {
    out.unshift(rebuildChartTitle(part.title));
  }
  if (!plotAreaEmitted) {
    // Synthetic chart with no captured plotArea — emit one from typed data.
    out.push(rebuildPlotArea(part, ctx));
  }
  const result: Record<string, unknown> = { "c:chart": out };
  const attrs = (chart[ATTR_KEY] as Record<string, unknown> | undefined) ?? undefined;
  if (attrs) result[ATTR_KEY] = { ...attrs };
  return result;
}

function rebuildChartTitle(title: string): Record<string, unknown> {
  const aT = makeEntry("a:t", [{ "#text": title }]);
  const aR = makeEntry("a:r", [aT]);
  const aP = makeEntry("a:p", [aR]);
  const cRich = makeEntry("c:rich", [aP]);
  const cTx = makeEntry("c:tx", [cRich]);
  const cOverlay = makeEntry("c:overlay", [], { val: "0" });
  return makeEntry("c:title", [cTx, cOverlay]);
}

function rebuildPlotArea(part: ChartPart, ctx?: ChartSerializeContext): Record<string, unknown> {
  const out: unknown[] = [];
  out.push(makeEntry("c:layout", []));
  out.push(rebuildChartTypeElement(part, ctx));
  for (const tail of part.plotAreaTailRaw) out.push(opaqueToEntry(tail));
  return makeEntry("c:plotArea", out);
}

function rebuildChartTypeElement(part: ChartPart, ctx?: ChartSerializeContext): Record<string, unknown> {
  const tag = chartTypeTag(part.chartType);
  const children: unknown[] = [];
  if (part.chartType === "bar") {
    children.push(makeEntry("c:barDir", [], { val: "col" }));
    children.push(makeEntry("c:grouping", [], { val: "clustered" }));
  }
  for (let i = 0; i < part.series.length; i++) {
    const s = part.series[i]!;
    children.push(rebuildSeries(part, s, ctx, i));
  }
  return makeEntry(tag, children);
}

function chartTypeTag(t: ChartPart["chartType"]): string {
  switch (t) {
    case "bar":
      return "c:barChart";
    case "line":
      return "c:lineChart";
    case "pie":
      return "c:pieChart";
    case "area":
      return "c:areaChart";
    case "unsupported":
      return "c:barChart";
  }
}

function rebuildSeries(
  part: ChartPart,
  s: ChartPart["series"][number],
  ctx: ChartSerializeContext | undefined,
  seriesIndex: number
): Record<string, unknown> {
  const children: unknown[] = [];
  children.push(makeEntry("c:idx", [], { val: String(s.idx) }));
  children.push(makeEntry("c:order", [], { val: String(s.idx) }));
  if (s.name !== undefined) {
    const nameRef = ctx?.nameRefs[seriesIndex];
    if (nameRef) {
      children.push(rebuildSeriesNameRef(s.name, nameRef));
    } else {
      children.push(makeEntry("c:tx", [makeEntry("c:v", [{ "#text": s.name }])]));
    }
  }
  if (part.categories.length > 0) {
    children.push(rebuildCategoryRef(part.categories, ctx?.categoryRef));
  }
  children.push(rebuildValueRef(s.values, ctx?.valueRefs[seriesIndex]));
  return makeEntry("c:ser", children);
}

function rebuildSeriesNameRef(name: string, ref: string): Record<string, unknown> {
  const cache = makeEntry("c:strCache", [
    makeEntry("c:ptCount", [], { val: "1" }),
    makeEntry("c:pt", [makeEntry("c:v", [{ "#text": name }])], { idx: "0" }),
  ]);
  return makeEntry("c:tx", [makeEntry("c:strRef", [makeEntry("c:f", [{ "#text": ref }]), cache])]);
}

function rebuildCategoryRef(
  categories: ReadonlyArray<string>,
  ref: string | undefined
): Record<string, unknown> {
  const ptCount = makeEntry("c:ptCount", [], { val: String(categories.length) });
  const pts: unknown[] = [ptCount];
  for (let i = 0; i < categories.length; i++) {
    pts.push(makeEntry("c:pt", [makeEntry("c:v", [{ "#text": categories[i] }])], { idx: String(i) }));
  }
  // Use c:strRef + c:strCache (standard PowerPoint shape) so reparse
  // round-trips cleanly. When an embedded XLSX is available the `<c:f>`
  // reference points at the real workbook range; otherwise a synthetic
  // Sheet1 reference is used as a fallback.
  const formula = ref ?? "Sheet1!$A$2:$A$" + (categories.length + 1);
  const cache = makeEntry("c:strCache", pts);
  const refNode = makeEntry("c:strRef", [makeEntry("c:f", [{ "#text": formula }]), cache]);
  return makeEntry("c:cat", [refNode]);
}

function rebuildValueRef(values: ReadonlyArray<number>, ref: string | undefined): Record<string, unknown> {
  const ptCount = makeEntry("c:ptCount", [], { val: String(values.length) });
  const pts: unknown[] = [ptCount];
  for (let i = 0; i < values.length; i++) {
    pts.push(
      makeEntry("c:pt", [makeEntry("c:v", [{ "#text": String(values[i] ?? 0) }])], { idx: String(i) })
    );
  }
  const formula = ref ?? "Sheet1!$B$2:$B$" + (values.length + 1);
  const cache = makeEntry("c:numCache", [makeEntry("c:formatCode", [{ "#text": "General" }]), ...pts]);
  const refNode = makeEntry("c:numRef", [makeEntry("c:f", [{ "#text": formula }]), cache]);
  return makeEntry("c:val", [refNode]);
}

function chartShapeToEntry(shape: ChartShape): Record<string, unknown> {
  const nvChildren: unknown[] = [];
  let emittedCNvPr = false;
  for (const o of shape.nvGraphicFramePrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else {
      nvChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
  }
  const nvGraphicFramePr = makeEntry("p:nvGraphicFramePr", nvChildren);

  const xfrmChildren: unknown[] = [];
  if (shape.position) {
    xfrmChildren.push(
      makeEntry("a:off", [], {
        x: String(shape.position.xEmu),
        y: String(shape.position.yEmu),
      })
    );
  }
  if (shape.size) {
    xfrmChildren.push(
      makeEntry("a:ext", [], {
        cx: String(shape.size.cxEmu),
        cy: String(shape.size.cyEmu),
      })
    );
  }
  const xfrm = makeEntry("p:xfrm", xfrmChildren);

  const chartEntry: Record<string, unknown> = {
    "c:chart": [],
  };
  chartEntry[ATTR_KEY] = makeRawAttrs({
    "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "r:id": shape.chartRelId,
  });

  const graphicData = makeEntry("a:graphicData", [chartEntry], { uri: shape.graphicDataUri });
  const graphic = makeEntry("a:graphic", [graphicData]);

  return makeEntry("p:graphicFrame", [nvGraphicFramePr, xfrm, graphic]);
}

/**
 * Re-emit `<p:graphicFrame>` for an OLE spreadsheet shape. The
 * `<p:oleObj>` payload (preview `<p:pic>`, `<p:embed>` flags, follow
 * content) was captured opaquely at parse time so untouched files
 * round-trip byte-identically; we only patch the rebuilt cNvPr id/name
 * and re-emit the frame's xfrm from the typed model.
 */
function oleSpreadsheetShapeToEntry(shape: OleSpreadsheetShape): Record<string, unknown> {
  const nvChildren: unknown[] = [];
  let emittedCNvPr = false;
  for (const o of shape.nvGraphicFramePrTail) {
    if (o.tag === "p:cNvPr" && !emittedCNvPr) {
      nvChildren.push(rebuildCNvPr(shape.cNvPrId, shape.name, o));
      emittedCNvPr = true;
    } else {
      nvChildren.push(opaqueToEntry(o));
    }
  }
  if (!emittedCNvPr) {
    nvChildren.unshift(makeEntry("p:cNvPr", [], { id: String(shape.cNvPrId), name: shape.name }));
  }
  const nvGraphicFramePr = makeEntry("p:nvGraphicFramePr", nvChildren);

  const xfrmChildren: unknown[] = [];
  if (shape.position) {
    xfrmChildren.push(
      makeEntry("a:off", [], {
        x: String(shape.position.xEmu),
        y: String(shape.position.yEmu),
      })
    );
  }
  if (shape.size) {
    xfrmChildren.push(
      makeEntry("a:ext", [], {
        cx: String(shape.size.cxEmu),
        cy: String(shape.size.cyEmu),
      })
    );
  }
  const xfrm = makeEntry("p:xfrm", xfrmChildren);

  // Patch r:id (the rel pointing at the embedded xlsx) into the
  // captured oleObj attribute bag. Everything else (spid, name, imgW,
  // imgH, …) is preserved as-is.
  const oleAttrs: Record<string, string> = { ...shape.oleObjAttrs };
  oleAttrs["r:id"] = shape.oleRelId;
  oleAttrs["progId"] = shape.progId;
  const oleObjEntry: Record<string, unknown> = {
    "p:oleObj": shape.oleObjChildrenRaw.map((o) => opaqueToEntry(o)),
  };
  oleObjEntry[ATTR_KEY] = makeRawAttrs(oleAttrs);

  const graphicData = makeEntry("a:graphicData", [oleObjEntry], { uri: shape.graphicDataUri });
  const graphic = makeEntry("a:graphic", [graphicData]);
  return makeEntry("p:graphicFrame", [nvGraphicFramePr, xfrm, graphic]);
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
  captured: OpaqueXml | undefined,
  rotation?: number
): Record<string, unknown> {
  // Preserve attrs on a:xfrm itself (e.g. flipH, rot).
  const xfrmAttrs: Record<string, string> = captured ? { ...captured.attrs } : {};
  // The model stores rotation in degrees; OOXML wants 60000ths of a
  // degree. `undefined`/`0` clear the attr so an unrotated shape
  // serialises identically to PowerPoint's own (no `rot=`), and a
  // previously-set rotation can be removed without leaving stale
  // attrs from the captured `<a:xfrm>` behind.
  if (rotation !== undefined && Number.isFinite(rotation)) {
    const normalised = ((rotation % 360) + 360) % 360;
    if (normalised === 0) {
      delete xfrmAttrs.rot;
    } else {
      xfrmAttrs.rot = String(Math.round(normalised * 60000));
    }
  } else {
    delete xfrmAttrs.rot;
  }
  const subChildren: unknown[] = [];
  // Re-emit a:off and a:ext from model when present; else fall back to captured subtree.
  if (position) {
    subChildren.push(makeEntry("a:off", [], { x: String(position.xEmu), y: String(position.yEmu) }));
  }
  if (size) {
    subChildren.push(makeEntry("a:ext", [], { cx: String(size.cxEmu), cy: String(size.cyEmu) }));
  }
  if (!position && !size && captured && Object.keys(xfrmAttrs).length === 0) {
    return opaqueToEntry(captured);
  }
  const entry: Record<string, unknown> = { "a:xfrm": subChildren };
  if (Object.keys(xfrmAttrs).length > 0) entry[ATTR_KEY] = makeRawAttrs(xfrmAttrs);
  return entry;
}

function buildGroupXfrm(shape: GroupShape, captured: OpaqueXml | undefined): Record<string, unknown> {
  const xfrmAttrs = captured ? captured.attrs : {};
  const subChildren: unknown[] = [];
  if (shape.position) {
    subChildren.push(
      makeEntry("a:off", [], { x: String(shape.position.xEmu), y: String(shape.position.yEmu) })
    );
  }
  if (shape.size) {
    subChildren.push(makeEntry("a:ext", [], { cx: String(shape.size.cxEmu), cy: String(shape.size.cyEmu) }));
  }
  for (const o of shape.chOffExtRaw) subChildren.push(opaqueToEntry(o));
  const entry: Record<string, unknown> = { "a:xfrm": subChildren };
  if (Object.keys(xfrmAttrs).length > 0) entry[ATTR_KEY] = makeRawAttrs(xfrmAttrs);
  return entry;
}

// ─── Text body serialization ──────────────────────────────────────────────

function textBodyToEntry(body: TextBody): Record<string, unknown> {
  return textBodyToEntryWith("p:txBody", body);
}

function textBodyToEntryWith(tag: "p:txBody" | "a:txBody", body: TextBody): Record<string, unknown> {
  const children: unknown[] = [];
  if (body.bodyPrRaw) children.push(opaqueToEntry(body.bodyPrRaw));
  else children.push(makeEntry("a:bodyPr", []));
  if (body.lstStyleRaw) children.push(opaqueToEntry(body.lstStyleRaw));
  else children.push(makeEntry("a:lstStyle", []));
  for (const p of body.paragraphs) children.push(paragraphToEntry(p));
  return makeEntry(tag, children);
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
  } else if (p.properties.alignment !== undefined || p.properties.level !== undefined) {
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
  const rPrChildren = buildRPrChildren(r);
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
  const rPrChildren = buildRPrChildren(r);
  const rPrAttrs = mergeRunAttrs(r);
  const rPrEntry: Record<string, unknown> = { "a:rPr": rPrChildren };
  if (Object.keys(rPrAttrs).length > 0) rPrEntry[ATTR_KEY] = makeRawAttrs(rPrAttrs);
  return makeEntry("a:br", [rPrEntry]);
}

/**
 * Build a:rPr children: re-emit a:solidFill (color) and a:latin (font) from
 * model properties, drop any matching captured children, and pass everything
 * else through verbatim. This ensures format-text changes are reflected.
 */
function buildRPrChildren(r: TextRun): unknown[] {
  const out: unknown[] = [];
  const captured = r.properties.opaqueChildren ?? [];
  const wantsSolidFill = r.properties.color !== undefined;
  const wantsHighlight = r.properties.highlight !== undefined;
  const dropHighlight = r.properties.highlight === "";

  // Effective typeface per script. The literal slot wins over the
  // theme ref at the same level (mirrors PowerPoint's resolution
  // rule). When neither typed slot is set we fall through to the
  // captured opaque element so source presentations that only
  // declare e.g. `a:cs` survive unchanged.
  const latinTypeface = r.properties.fontFamily ?? r.properties.fontFamilyLatinTheme ?? undefined;
  const eaTypeface = r.properties.fontFamilyEastAsia ?? r.properties.fontFamilyEastAsiaTheme ?? undefined;
  const csTypeface =
    r.properties.fontFamilyComplexScript ?? r.properties.fontFamilyComplexScriptTheme ?? undefined;
  const symTypeface = r.properties.fontFamilySymbol ?? undefined;
  const wantsLatin = latinTypeface !== undefined;
  const wantsEa = eaTypeface !== undefined;
  const wantsCs = csTypeface !== undefined;
  const wantsSym = symTypeface !== undefined;

  // Emit fill first, then highlight, then other children, then the
  // typeface elements, matching OOXML's preferred a:rPr child order
  // (a:ln, a:solidFill, a:highlight, a:effectLst, …,
  // a:latin, a:ea, a:cs, a:sym, …).
  if (wantsSolidFill) {
    out.push(makeEntry("a:solidFill", [makeEntry("a:srgbClr", [], { val: String(r.properties.color) })]));
  }
  if (wantsHighlight && !dropHighlight) {
    out.push(makeEntry("a:highlight", [makeEntry("a:srgbClr", [], { val: String(r.properties.highlight) })]));
  }
  for (const c of captured) {
    if (c.tag === "a:solidFill" && wantsSolidFill) continue;
    if (c.tag === "a:latin" && wantsLatin) continue;
    if (c.tag === "a:ea" && wantsEa) continue;
    if (c.tag === "a:cs" && wantsCs) continue;
    if (c.tag === "a:sym" && wantsSym) continue;
    if (c.tag === "a:highlight" && wantsHighlight) continue;
    out.push(opaqueToEntry(c));
  }
  if (wantsLatin) out.push(makeEntry("a:latin", [], { typeface: String(latinTypeface) }));
  if (wantsEa) out.push(makeEntry("a:ea", [], { typeface: String(eaTypeface) }));
  if (wantsCs) out.push(makeEntry("a:cs", [], { typeface: String(csTypeface) }));
  if (wantsSym) out.push(makeEntry("a:sym", [], { typeface: String(symTypeface) }));
  return out;
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
  // <p:sldIdLst> with one that matches root.slides order and <p:sldSz>
  // with one matching root.slideSize (which may have been updated by
  // pptx:set-slide-size).
  const newChildren: unknown[] = [];
  for (const o of root.presentationOpaqueTail) {
    if (o.tag === "p:sldIdLst") {
      newChildren.push(buildSldIdLst(root));
    } else if (o.tag === "p:sldSz") {
      newChildren.push(buildSldSz(root));
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

function buildSldSz(root: PptxPresentation): Record<string, unknown> {
  const attrs: Record<string, string> = {
    cx: String(root.slideSize.cxEmu),
    cy: String(root.slideSize.cyEmu),
  };
  if (root.slideSize.type) attrs.type = root.slideSize.type;
  return makeEntry("p:sldSz", [], attrs);
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
