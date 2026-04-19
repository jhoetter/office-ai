import { defaultIdMinter, ooxml, sha256Hex, type IdMinter } from "@officeai/core";
import { DEFAULT_THEME, type ThemeColorScheme } from "../renderer/layout/color.js";
import { parseThemeColorScheme } from "./theme.js";
import type {
  ChartPart,
  ChartSeries,
  ChartShape,
  ChartType,
  ConnectorEndShape,
  ConnectorEndpoint,
  ConnectorShape,
  ConnectorSide,
  ConnectorStroke,
  ConnectorType,
  ContentTypesSnap,
  EntranceAnimation,
  EntranceEffect,
  GroupShape,
  MediaPart,
  OpaquePart,
  OpaqueShape,
  OpaqueXml,
  Picture,
  NotesSlide,
  PlaceholderSpec,
  PptxComment,
  PptxCommentAuthor,
  PptxCommentAuthorsPart,
  PptxCommentsPart,
  Position,
  PptxIdGen,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Shape,
  Size,
  Slide,
  SlideLayout,
  SlideSize,
  SlideTransition,
  TableCell,
  TableRow,
  TableShape,
  TextBody,
  TextParagraph,
  TextParagraphProperties,
  TextRun,
  TextRunProperties,
  TextShape,
  TransitionKind,
  TransitionSpeed,
} from "../model/types.js";
import { emptyDirty } from "../model/types.js";
import { PptxParseError } from "./errors.js";
import {
  attrOf,
  captureOpaque,
  elementEntries,
  findElementEntry,
  readRawAttrs,
  readRootAttrs,
  readText,
  rootEntry,
} from "./xml-helpers.js";

export interface ParseOptions {
  readonly idMinter?: IdMinter;
}

const PRESENTATION_PART = "ppt/presentation.xml";
const REL_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_TYPE_SLIDE_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_TYPE_SLIDE_LAYOUT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_TYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_TYPE_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_TYPE_COMMENT_AUTHORS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

export async function parsePptx(
  input: ArrayBuffer | Uint8Array,
  opts: ParseOptions = {}
): Promise<PptxSnapshot> {
  let container: ooxml.OoxmlContainer;
  try {
    container = await ooxml.OoxmlContainer.load(input);
  } catch (err) {
    throw new PptxParseError("invalid-zip", "Failed to read PPTX as a zip archive", { cause: err });
  }
  if (!container.has(PRESENTATION_PART)) {
    throw new PptxParseError("missing-main-part", `Missing required part: ${PRESENTATION_PART}`, {
      partPath: PRESENTATION_PART,
    });
  }

  const mintNodeId: IdMinter = opts.idMinter ?? defaultIdMinter;

  // ── presentation.xml ────────────────────────────────────────────────
  const presentationXml = container.readText(PRESENTATION_PART);
  let presentationTree: unknown;
  try {
    presentationTree = ooxml.parseXml(presentationXml);
  } catch (err) {
    throw new PptxParseError("invalid-xml", `Failed to parse ${PRESENTATION_PART}`, {
      partPath: PRESENTATION_PART,
      cause: err,
    });
  }
  let presentationEntry: Record<string, unknown>;
  try {
    presentationEntry = rootEntry(presentationTree, "p:presentation");
  } catch (err) {
    throw new PptxParseError("missing-root", "Missing <p:presentation> root", {
      partPath: PRESENTATION_PART,
      cause: err,
    });
  }

  const presentationRootAttrs = readRootAttrs(presentationEntry);
  const presentationChildren = (presentationEntry["p:presentation"] as unknown[] | undefined) ?? [];

  // Slide id list — drives slide order, slideId, and rId.
  const sldIdLst = findElementEntry(presentationChildren, "p:sldIdLst");
  const sldIdLstAttrs = sldIdLst ? readRootAttrs(sldIdLst) : {};
  const sldIdEntries: Array<{ slideId: number; relId: string }> = [];
  if (sldIdLst) {
    for (const c of elementEntries((sldIdLst["p:sldIdLst"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(c) !== "p:sldId") continue;
      const idStr = attrOf(c, "id") ?? "0";
      const rid = attrOf(c, "r:id") ?? "";
      sldIdEntries.push({ slideId: Number(idStr), relId: rid });
    }
  }

  // Slide size.
  const slideSize = readSlideSize(presentationChildren) ?? {
    cxEmu: 9144000,
    cyEmu: 6858000,
  };
  const notesSize = readNotesSize(presentationChildren);

  // Anything in the presentation we don't model is captured as opaque tail
  // (we re-emit the whole tree from these spans on dirty serialize).
  const presentationOpaqueTail: OpaqueXml[] = [];
  for (const c of elementEntries(presentationChildren)) {
    presentationOpaqueTail.push(captureOpaque(c));
  }

  // ── presentation rels → resolve slide part paths ────────────────────
  const presRels = ooxml.RelationshipGraph.loadFor(container, PRESENTATION_PART);
  const slideRelById = new Map<string, ooxml.Relationship>();
  for (const r of presRels.relationships) slideRelById.set(r.id, r);

  // ── Discover all slides, masters, layouts, themes, notes ────────────
  const masterPaths = collectPartsByType(presRels, REL_TYPE_SLIDE_MASTER);
  const layoutPaths = new Set<string>();
  const themePaths = new Set<string>();
  for (const m of masterPaths) {
    const masterRels = ooxml.RelationshipGraph.loadFor(container, m);
    for (const r of masterRels.relationships) {
      if (r.type === REL_TYPE_SLIDE_LAYOUT) {
        layoutPaths.add(resolveTarget(m, r.target));
      } else if (r.type === REL_TYPE_THEME) {
        themePaths.add(resolveTarget(m, r.target));
      }
    }
  }

  // Slides in <p:sldIdLst> order.
  const slides: Slide[] = [];
  let nextSlidePartIndex = 1;
  const notesSlidePaths = new Set<string>();

  for (const entry of sldIdEntries) {
    const rel = slideRelById.get(entry.relId);
    if (!rel) {
      throw new PptxParseError(
        "missing-slide-rel",
        `Slide rel ${entry.relId} not found in presentation rels`,
        { partPath: PRESENTATION_PART }
      );
    }
    const partPath = resolveTarget(PRESENTATION_PART, rel.target);
    if (!container.has(partPath)) {
      throw new PptxParseError("missing-slide-part", `Slide part missing: ${partPath}`, {
        partPath,
      });
    }
    const slide = parseSlide(container, partPath, entry.slideId, entry.relId, mintNodeId);
    slides.push(slide);
    const m = /slide(\d+)\.xml$/i.exec(partPath);
    if (m) {
      const n = Number(m[1]);
      if (!Number.isNaN(n) && n >= nextSlidePartIndex) nextSlidePartIndex = n + 1;
    }
    if (slide.notesSlidePartPath) notesSlidePaths.add(slide.notesSlidePartPath);
  }

  // ── Opaque master/layout/theme/notes/media maps ─────────────────────
  const masters = new Map<string, OpaquePart>();
  for (const p of masterPaths) {
    masters.set(p, opaquePartFor(container, p, "p:sldMaster"));
  }
  const layouts = new Map<string, SlideLayout>();
  for (const p of layoutPaths) {
    layouts.set(p, parseSlideLayout(container, p));
  }
  const theme = new Map<string, OpaquePart>();
  let themeDefault: ThemeColorScheme = DEFAULT_THEME;
  let themeDefaultFromPath: string | null = null;
  for (const p of themePaths) {
    theme.set(p, opaquePartFor(container, p, "a:theme"));
    // First theme part wins; usually `theme1.xml`. Subsequent themes are
    // preserved verbatim but not used to drive `<a:schemeClr>` resolution.
    if (themeDefaultFromPath === null) {
      try {
        themeDefault = parseThemeColorScheme(container.readText(p));
        themeDefaultFromPath = p;
      } catch {
        // Keep DEFAULT_THEME on any parse hiccup.
      }
    }
  }
  const notesSlides = new Map<string, NotesSlide>();
  for (const p of notesSlidePaths) {
    if (container.has(p)) notesSlides.set(p, parseNotesSlide(container, p, mintNodeId));
  }

  // ── Comments ────────────────────────────────────────────────────────
  const commentsByPart = new Map<string, PptxCommentsPart>();
  for (const slide of slides) {
    if (!slide.commentsPartPath || !container.has(slide.commentsPartPath)) continue;
    if (commentsByPart.has(slide.commentsPartPath)) continue;
    commentsByPart.set(
      slide.commentsPartPath,
      parseCommentsPart(container, slide.commentsPartPath, mintNodeId)
    );
  }
  let commentAuthors: PptxCommentAuthorsPart | null = null;
  // commentAuthors is referenced from the presentation rels.
  for (const r of presRels.relationships) {
    if (r.type !== REL_TYPE_COMMENT_AUTHORS) continue;
    const p = resolveTarget(PRESENTATION_PART, r.target);
    if (!container.has(p)) continue;
    commentAuthors = parseCommentAuthorsPart(container, p);
    break;
  }

  // Media: every binary under ppt/media/.
  const media = new Map<string, MediaPart>();
  let nextMediaPartIndex = 1;
  for (const partPath of container.parts.keys()) {
    if (!partPath.startsWith("ppt/media/")) continue;
    const bytes = container.readBytes(partPath);
    const sha = sha256Hex(bytes);
    const ext = partPath.slice(partPath.lastIndexOf(".") + 1).toLowerCase();
    const contentType = mediaContentType(ext);
    media.set(partPath, { partPath, bytes, sha256: sha, contentType });
    const im = /media\/image(\d+)\./i.exec(partPath);
    if (im) {
      const n = Number(im[1]);
      if (!Number.isNaN(n) && n >= nextMediaPartIndex) nextMediaPartIndex = n + 1;
    }
  }

  // Compute next slideId (≥ 256, monotonically increasing).
  let nextSlideId = 256;
  for (const e of sldIdEntries) {
    if (e.slideId >= nextSlideId) nextSlideId = e.slideId + 1;
  }

  const idGen: PptxIdGen = {
    nextSlideId,
    nextSlidePartIndex,
    nextMediaPartIndex,
  };

  // Collect all relationships parts as snapshots (used by serializer to
  // detect stale rels and rewrite when needed).
  const relationships = new Map<string, RelationshipsSnap>();
  for (const partPath of container.parts.keys()) {
    if (!partPath.endsWith(".rels")) continue;
    const graph = ooxml.RelationshipGraph.loadFor(container, relsPathToOwnerPath(partPath) ?? partPath);
    relationships.set(partPath, {
      relsPath: partPath,
      entries: graph.relationships.map((r) => ({
        id: r.id,
        type: r.type,
        target: r.target,
        ...(r.targetMode ? { targetMode: r.targetMode } : {}),
      })),
    });
  }

  // Content types snapshot.
  const ct = ooxml.ContentTypes.load(container);
  const contentTypes: ContentTypesSnap = {
    defaults: ct.defaults.map((d) => ({ extension: d.extension, contentType: d.contentType })),
    overrides: ct.overrides.map((o) => ({ partName: o.partName, contentType: o.contentType })),
  };

  // ── Charts (F3): typed parts referenced by ChartShape graphic frames ─
  const charts = new Map<string, ChartPart>();
  const chartPartPaths = new Set<string>();
  for (const slide of slides) {
    walkShapesForCharts(slide.shapes, (path) => chartPartPaths.add(path));
  }
  for (const chartPath of chartPartPaths) {
    if (!container.has(chartPath)) continue;
    const chartXml = container.readText(chartPath);
    const ctOverride = ct.overrides.find((o) => o.partName === `/${chartPath}`)?.contentType;
    const ctype = ctOverride ?? "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    try {
      charts.set(chartPath, parseChartPart(chartPath, chartXml, ctype, mintNodeId));
    } catch {
      // Skip malformed chart parts; they remain as opaque container bytes.
    }
  }

  const root: PptxPresentation = {
    id: mintNodeId(),
    slides,
    slideSize,
    ...(notesSize ? { notesSize } : {}),
    masters,
    layouts,
    theme,
    themeDefault,
    notesSlides,
    commentsByPart,
    commentAuthors,
    media,
    charts,
    presentationRootAttrs,
    presentationOpaqueTail,
    sldIdLstAttrs,
    idGen,
  };

  const partHashes: Record<string, string> = {};
  for (const path of container.parts.keys()) {
    partHashes[path] = sha256Hex(container.readBytes(path));
  }

  return {
    format: "pptx",
    revision: 0,
    root,
    partHashes,
    container,
    dirty: emptyDirty(),
    removedParts: new Set<string>(),
    relationships,
    contentTypes,
  };
}

// ─── Slide parser ─────────────────────────────────────────────────────────

function parseSlide(
  container: ooxml.OoxmlContainer,
  partPath: string,
  slideId: number,
  relId: string,
  mintNodeId: IdMinter
): Slide {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(partPath));
  } catch (err) {
    throw new PptxParseError("invalid-xml", `Failed to parse slide XML: ${partPath}`, {
      partPath,
      cause: err,
    });
  }
  const sld = rootEntry(tree, "p:sld");
  const slideRootAttrs = readRootAttrs(sld);

  const sldChildren = (sld["p:sld"] as unknown[] | undefined) ?? [];
  const cSld = findElementEntry(sldChildren, "p:cSld");
  if (!cSld) {
    throw new PptxParseError("missing-csld", `Missing <p:cSld> in slide: ${partPath}`, {
      partPath,
    });
  }
  const cSldAttrs = readRootAttrs(cSld);
  const cSldChildren = (cSld["p:cSld"] as unknown[] | undefined) ?? [];
  const spTree = findElementEntry(cSldChildren, "p:spTree");
  if (!spTree) {
    throw new PptxParseError("missing-sptree", `Missing <p:spTree> in slide: ${partPath}`, {
      partPath,
    });
  }

  // Capture cSld head children that come BEFORE p:spTree (e.g. <p:bg>).
  const cSldHead: OpaqueXml[] = [];
  for (const c of elementEntries(cSldChildren)) {
    const tag = ooxml.getTag(c);
    if (tag === "p:spTree") break;
    cSldHead.push(captureOpaque(c));
  }
  // Slide rels → layout, notes slide. Loaded BEFORE shape parsing so picture
  // rels can be resolved to absolute media part paths in one pass.
  const rels = ooxml.RelationshipGraph.loadFor(container, partPath);
  const slideRelTargets = new Map<string, string>();
  let layoutPartPath: string | undefined;
  let notesSlidePartPath: string | undefined;
  let commentsPartPath: string | undefined;
  for (const r of rels.relationships) {
    slideRelTargets.set(r.id, r.target);
    if (r.type === REL_TYPE_SLIDE_LAYOUT) {
      layoutPartPath = resolveTarget(partPath, r.target);
    } else if (r.type === REL_TYPE_NOTES_SLIDE) {
      notesSlidePartPath = resolveTarget(partPath, r.target);
    } else if (r.type === REL_TYPE_COMMENTS) {
      commentsPartPath = resolveTarget(partPath, r.target);
    }
  }

  // spTree: head children (nvGrpSpPr, grpSpPr) preserved opaquely;
  // shapes are parsed into the model.
  const spTreeChildren = (spTree["p:spTree"] as unknown[] | undefined) ?? [];
  const spTreeHead: OpaqueXml[] = [];
  const shapes: Shape[] = [];
  for (const c of elementEntries(spTreeChildren)) {
    const tag = ooxml.getTag(c);
    if (tag === "p:nvGrpSpPr" || tag === "p:grpSpPr") {
      spTreeHead.push(captureOpaque(c));
      continue;
    }
    shapes.push(parseShape(c, mintNodeId, partPath, slideRelTargets));
  }

  // Anything in <p:sld> after <p:cSld> we capture as opaque tail
  // (e.g. <p:clrMapOvr>, <p:transition>, <p:timing>). F4 promotes
  // <p:transition> and the typed pieces of <p:timing> into typed
  // model fields. Done AFTER shape parsing so adding/removing typed
  // animations doesn't shift shape NodeIds (they would otherwise depend
  // on whether <p:timing> was present at parse time).
  const slideOpaqueTail: OpaqueXml[] = [];
  let transition: SlideTransition | undefined;
  let timingTailRaw: OpaqueXml | undefined;
  const animations: EntranceAnimation[] = [];
  let pastCsld = false;
  for (const c of elementEntries(sldChildren)) {
    if (!pastCsld) {
      if (ooxml.getTag(c) === "p:cSld") pastCsld = true;
      continue;
    }
    const tag = ooxml.getTag(c);
    if (tag === "p:transition") {
      transition = parseSlideTransition(c, mintNodeId);
      continue;
    }
    if (tag === "p:timing") {
      const parsedTiming = parseSlideTiming(c, mintNodeId);
      animations.push(...parsedTiming.animations);
      if (parsedTiming.tail) timingTailRaw = parsedTiming.tail;
      continue;
    }
    slideOpaqueTail.push(captureOpaque(c));
  }

  return {
    id: mintNodeId(),
    partPath,
    slideId,
    relId,
    ...(layoutPartPath ? { layoutPartPath } : {}),
    ...(notesSlidePartPath ? { notesSlidePartPath } : {}),
    ...(commentsPartPath ? { commentsPartPath } : {}),
    shapes,
    ...(transition ? { transition } : {}),
    animations,
    ...(timingTailRaw ? { timingTailRaw } : {}),
    slideOpaqueTail,
    slideRootAttrs,
    cSldAttrs,
    spTreeHead,
    cSldHead,
  };
}

// ─── Shape parser ─────────────────────────────────────────────────────────

function parseShape(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  partPath: string,
  slideRelTargets: ReadonlyMap<string, string>
): Shape {
  const tag = ooxml.getTag(entry);
  switch (tag) {
    case "p:sp":
      return parseSp(entry, mintNodeId);
    case "p:pic":
      return parsePic(entry, mintNodeId, partPath, slideRelTargets);
    case "p:cxnSp":
      return parseCxnSp(entry, mintNodeId);
    case "p:grpSp":
      return parseGrpSp(entry, mintNodeId, partPath, slideRelTargets);
    case "p:graphicFrame": {
      const table = parseGraphicFrameTable(entry, mintNodeId);
      if (table) return table;
      const chart = parseGraphicFrameChart(entry, mintNodeId, partPath, slideRelTargets);
      if (chart) return chart;
      return parseOpaqueShape(entry, mintNodeId);
    }
    default:
      return parseOpaqueShape(entry, mintNodeId);
  }
}

const TABLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const CHART_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/**
 * Parse `<p:graphicFrame>` ⇒ `ChartShape` if its `<a:graphicData @uri>`
 * is the chart URI. Returns `null` otherwise so the caller can fall
 * back to `OpaqueShape` for SmartArt and friends.
 */
function parseGraphicFrameChart(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  slidePartPath: string,
  slideRelTargets: ReadonlyMap<string, string>
): ChartShape | null {
  const children = (entry["p:graphicFrame"] as unknown[] | undefined) ?? [];
  const nvGFPr = findElementEntry(children, "p:nvGraphicFramePr");
  const xfrm = findElementEntry(children, "p:xfrm");
  const graphic = findElementEntry(children, "a:graphic");
  if (!graphic) return null;
  const graphicData = findElementEntry(
    (graphic["a:graphic"] as unknown[] | undefined) ?? [],
    "a:graphicData"
  );
  if (!graphicData) return null;
  const uri = attrOf(graphicData, "uri") ?? "";
  if (uri !== CHART_GRAPHIC_DATA_URI) return null;
  const chart = findElementEntry((graphicData["a:graphicData"] as unknown[] | undefined) ?? [], "c:chart");
  if (!chart) return null;
  const chartRelId = attrOf(chart, "r:id") ?? "";
  if (!chartRelId) return null;
  const targetRel = slideRelTargets.get(chartRelId);
  const chartPartPath = targetRel ? resolveTarget(slidePartPath, targetRel) : "";

  let cNvPrId = 0;
  let name = "";
  const nvGFPrTail: OpaqueXml[] = [];
  if (nvGFPr) {
    for (const c of elementEntries((nvGFPr["p:nvGraphicFramePr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "p:cNvPr") {
        cNvPrId = Number(attrOf(c, "id") ?? "0");
        name = attrOf(c, "name") ?? "";
      }
      nvGFPrTail.push(captureOpaque(c));
    }
  }

  let position: { xEmu: number; yEmu: number } | undefined;
  let size: { cxEmu: number; cyEmu: number } | undefined;
  if (xfrm) {
    const xfrmChildren = (xfrm["p:xfrm"] as unknown[] | undefined) ?? [];
    const off = findElementEntry(xfrmChildren, "a:off");
    const ext = findElementEntry(xfrmChildren, "a:ext");
    if (off) {
      position = {
        xEmu: Number(attrOf(off, "x") ?? "0"),
        yEmu: Number(attrOf(off, "y") ?? "0"),
      };
    }
    if (ext) {
      size = {
        cxEmu: Number(attrOf(ext, "cx") ?? "0"),
        cyEmu: Number(attrOf(ext, "cy") ?? "0"),
      };
    }
  }

  return {
    kind: "chart",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    chartRelId,
    chartPartPath,
    nvGraphicFramePrTail: nvGFPrTail,
    graphicDataUri: uri,
  };
}

/**
 * Parse `<p:graphicFrame>` ⇒ `TableShape` if and only if its
 * `<a:graphicData @uri>` is the table URI. Returns `null` otherwise so
 * the caller can fall back to `OpaqueShape` for charts / SmartArt.
 */
function parseGraphicFrameTable(entry: Record<string, unknown>, mintNodeId: IdMinter): TableShape | null {
  const children = (entry["p:graphicFrame"] as unknown[] | undefined) ?? [];
  const nvGFPr = findElementEntry(children, "p:nvGraphicFramePr");
  const xfrm = findElementEntry(children, "p:xfrm");
  const graphic = findElementEntry(children, "a:graphic");
  if (!graphic) return null;
  const graphicData = findElementEntry(
    (graphic["a:graphic"] as unknown[] | undefined) ?? [],
    "a:graphicData"
  );
  if (!graphicData) return null;
  const uri = attrOf(graphicData, "uri") ?? "";
  if (uri !== TABLE_GRAPHIC_DATA_URI) return null;
  const tbl = findElementEntry((graphicData["a:graphicData"] as unknown[] | undefined) ?? [], "a:tbl");
  if (!tbl) return null;

  let cNvPrId = 0;
  let name = "";
  const nvGFPrTail: OpaqueXml[] = [];
  if (nvGFPr) {
    for (const c of elementEntries((nvGFPr["p:nvGraphicFramePr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "p:cNvPr") {
        cNvPrId = Number(attrOf(c, "id") ?? "0");
        name = attrOf(c, "name") ?? "";
      }
      nvGFPrTail.push(captureOpaque(c));
    }
  }

  let position: { xEmu: number; yEmu: number } | undefined;
  let size: { cxEmu: number; cyEmu: number } | undefined;
  if (xfrm) {
    const xfrmChildren = (xfrm["p:xfrm"] as unknown[] | undefined) ?? [];
    const off = findElementEntry(xfrmChildren, "a:off");
    const ext = findElementEntry(xfrmChildren, "a:ext");
    if (off) {
      position = {
        xEmu: Number(attrOf(off, "x") ?? "0"),
        yEmu: Number(attrOf(off, "y") ?? "0"),
      };
    }
    if (ext) {
      size = {
        cxEmu: Number(attrOf(ext, "cx") ?? "0"),
        cyEmu: Number(attrOf(ext, "cy") ?? "0"),
      };
    }
  }

  const tblChildren = (tbl["a:tbl"] as unknown[] | undefined) ?? [];
  const tblPr = findElementEntry(tblChildren, "a:tblPr");
  const tblGrid = findElementEntry(tblChildren, "a:tblGrid");

  const columnWidths: number[] = [];
  if (tblGrid) {
    for (const c of elementEntries((tblGrid["a:tblGrid"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(c) !== "a:gridCol") continue;
      columnWidths.push(Number(attrOf(c, "w") ?? "0"));
    }
  }

  const rows: TableRow[] = [];
  for (const tr of elementEntries(tblChildren)) {
    if (ooxml.getTag(tr) !== "a:tr") continue;
    const trAttrs = readRootAttrs(tr);
    const heightStr = trAttrs["h"];
    const height = heightStr !== undefined ? Number(heightStr) : 0;
    const trAttrsRest: Record<string, string> = { ...trAttrs };
    delete trAttrsRest["h"];

    const cells: TableCell[] = [];
    for (const tc of elementEntries((tr["a:tr"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(tc) !== "a:tc") continue;
      const tcAttrs = readRootAttrs(tc);
      const tcChildren = (tc["a:tc"] as unknown[] | undefined) ?? [];
      const txBodyEntry = findElementEntry(tcChildren, "a:txBody");
      const tcPr = findElementEntry(tcChildren, "a:tcPr");
      const txBody: TextBody = txBodyEntry
        ? parseTextBodyChildren((txBodyEntry["a:txBody"] as unknown[] | undefined) ?? [], mintNodeId)
        : emptyTextBody();
      cells.push({
        id: mintNodeId(),
        txBody,
        ...(tcPr ? { tcPrRaw: captureOpaque(tcPr) } : {}),
        tcAttrs,
      });
    }

    rows.push({
      id: mintNodeId(),
      height,
      cells,
      trAttrs: trAttrsRest,
    });
  }

  return {
    kind: "table",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    columnWidths,
    rows,
    ...(tblPr ? { tblPrRaw: captureOpaque(tblPr) } : {}),
    nvGraphicFramePrTail: nvGFPrTail,
    graphicDataUri: uri,
  };
}

function parseSp(entry: Record<string, unknown>, mintNodeId: IdMinter): TextShape {
  const children = (entry["p:sp"] as unknown[] | undefined) ?? [];
  const nvSpPr = findElementEntry(children, "p:nvSpPr");
  const spPr = findElementEntry(children, "p:spPr");
  const txBodyEntry = findElementEntry(children, "p:txBody");
  const styleEntry = findElementEntry(children, "p:style");

  const { cNvPrId, name, nvTail } = readNvSpPr(nvSpPr);
  const { position, size, spPrTail } = readSpPr(spPr);
  const placeholder = readPlaceholder(nvSpPr);

  const txBody = txBodyEntry ? parseTextBody(txBodyEntry, mintNodeId) : emptyTextBody();

  return {
    kind: "text",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    ...(placeholder ? { placeholder } : {}),
    txBody,
    nvSpPrTail: nvTail,
    spPrTail,
    ...(styleEntry ? { styleRaw: captureOpaque(styleEntry) } : {}),
  };
}

/**
 * Parse `<p:cxnSp>` ⇒ `ConnectorShape`. Endpoint anchoring is recovered
 * from `<a:stCxn>` / `<a:endCxn>` inside `<p:spPr>`; when those are
 * absent we fall back to the bounding-box corners (with `<a:xfrm>`'s
 * `flipH`/`flipV` flags consulted to pick which corner). The connector
 * type is read from `<a:prstGeom @prst>` (line / bentConnector* /
 * curvedConnector*); anything else lands as `unsupported`.
 */
function parseCxnSp(entry: Record<string, unknown>, mintNodeId: IdMinter): ConnectorShape {
  const children = (entry["p:cxnSp"] as unknown[] | undefined) ?? [];
  const nvCxnSpPr = findElementEntry(children, "p:nvCxnSpPr");
  const spPr = findElementEntry(children, "p:spPr");

  let cNvPrId = 0;
  let name = "";
  const nvCxnSpPrTail: OpaqueXml[] = [];
  let stCxnSpid: number | null = null;
  let stCxnIdx: string | null = null;
  let endCxnSpid: number | null = null;
  let endCxnIdx: string | null = null;
  if (nvCxnSpPr) {
    for (const c of elementEntries((nvCxnSpPr["p:nvCxnSpPr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "p:cNvPr") {
        cNvPrId = Number(attrOf(c, "id") ?? "0");
        name = attrOf(c, "name") ?? "";
      }
      if (tag === "p:cNvCxnSpPr") {
        for (const cx of elementEntries((c["p:cNvCxnSpPr"] as unknown[] | undefined) ?? [])) {
          const cxTag = ooxml.getTag(cx);
          if (cxTag === "a:stCxn") {
            const id = attrOf(cx, "id");
            const idx = attrOf(cx, "idx");
            if (id && /^-?\d+$/.test(id)) stCxnSpid = Number(id);
            if (idx) stCxnIdx = idx;
          } else if (cxTag === "a:endCxn") {
            const id = attrOf(cx, "id");
            const idx = attrOf(cx, "idx");
            if (id && /^-?\d+$/.test(id)) endCxnSpid = Number(id);
            if (idx) endCxnIdx = idx;
          }
        }
      }
      nvCxnSpPrTail.push(captureOpaque(c));
    }
  }

  let position: { xEmu: number; yEmu: number } | undefined;
  let size: { cxEmu: number; cyEmu: number } | undefined;
  let flipH = false;
  let flipV = false;
  let connectorType: ConnectorType = "unsupported";
  let stroke: ConnectorStroke | undefined;
  let headEnd: ConnectorEndShape | undefined;
  let tailEnd: ConnectorEndShape | undefined;
  const spPrTail: OpaqueXml[] = [];
  if (spPr) {
    for (const c of elementEntries((spPr["p:spPr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "a:xfrm") {
        const xfrmAttrs = readRootAttrs(c);
        flipH = xfrmAttrs.flipH === "1" || xfrmAttrs.flipH === "true";
        flipV = xfrmAttrs.flipV === "1" || xfrmAttrs.flipV === "true";
        const xfrmChildren = (c["a:xfrm"] as unknown[] | undefined) ?? [];
        const off = findElementEntry(xfrmChildren, "a:off");
        const ext = findElementEntry(xfrmChildren, "a:ext");
        if (off) {
          position = {
            xEmu: Number(attrOf(off, "x") ?? "0"),
            yEmu: Number(attrOf(off, "y") ?? "0"),
          };
        }
        if (ext) {
          size = {
            cxEmu: Number(attrOf(ext, "cx") ?? "0"),
            cyEmu: Number(attrOf(ext, "cy") ?? "0"),
          };
        }
        continue;
      }
      if (tag === "a:prstGeom") {
        const prst = attrOf(c, "prst") ?? "";
        connectorType = mapPrstToConnectorType(prst);
        continue;
      }
      if (tag === "a:ln") {
        const lnAttrs = readRootAttrs(c);
        const widthAttr = lnAttrs.w;
        let widthEmu = 0;
        if (widthAttr && /^\d+$/.test(widthAttr)) widthEmu = Number(widthAttr);
        let color: string | undefined;
        let dash: "solid" | "dashed" | "dotted" | "longDash" | "dashDot" | undefined;
        for (const ln of elementEntries((c["a:ln"] as unknown[] | undefined) ?? [])) {
          const lnTag = ooxml.getTag(ln);
          if (lnTag === "a:solidFill") {
            const srgb = findElementEntry((ln["a:solidFill"] as unknown[] | undefined) ?? [], "a:srgbClr");
            if (srgb) {
              const v = attrOf(srgb, "val");
              if (v) color = v;
            }
          } else if (lnTag === "a:headEnd") {
            const t = attrOf(ln, "type");
            if (t) headEnd = mapEndShape(t);
          } else if (lnTag === "a:tailEnd") {
            const t = attrOf(ln, "type");
            if (t) tailEnd = mapEndShape(t);
          } else if (lnTag === "a:prstDash") {
            const v = attrOf(ln, "val");
            if (v) dash = mapPrstDash(v);
          }
        }
        if (color !== undefined || widthEmu > 0 || dash !== undefined) {
          stroke = { color: color ?? "000000", widthEmu, ...(dash !== undefined ? { dash } : {}) };
        }
      }
      spPrTail.push(captureOpaque(c));
    }
  }

  // Resolve endpoints. PowerPoint encodes start/end in two places:
  //   1) `<a:stCxn>` / `<a:endCxn>` carry an id + side index (anchored).
  //   2) `<a:xfrm>` carries the bounding box; combined with `flipH`/
  //      `flipV` this gives the actual start/end corners for free
  //      endpoints (no anchor).
  const start = resolveCxnEndpoint(stCxnSpid, stCxnIdx, "start", position, size, flipH, flipV);
  const end = resolveCxnEndpoint(endCxnSpid, endCxnIdx, "end", position, size, flipH, flipV);

  return {
    kind: "connector",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    connectorType,
    start,
    end,
    ...(stroke ? { stroke } : {}),
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
    nvCxnSpPrTail,
    spPrTail,
  };
}

function mapPrstToConnectorType(prst: string): ConnectorType {
  if (prst === "line" || prst === "straightConnector1") return "straight";
  if (prst.startsWith("bentConnector")) return "elbow";
  if (prst.startsWith("curvedConnector")) return "curved";
  return "unsupported";
}

function mapPrstDash(
  v: string
): "solid" | "dashed" | "dotted" | "longDash" | "dashDot" {
  // PowerPoint exposes ~10 dash presets; we collapse them to the five
  // the editor exposes. Mapping mirrors `serialize.ts` so a round-trip
  // doesn't drift between presets. Unknown presets fall through to
  // solid, which is the no-op default.
  const k = v.toLowerCase();
  if (k === "solid") return "solid";
  if (k === "dot" || k === "sysdot") return "dotted";
  if (k === "lgdash" || k === "syslgdash" || k === "lgdashdot" || k === "lgdashdotdot") {
    return "longDash";
  }
  if (k === "dashdot" || k === "sysdashdot" || k === "sysdashdotdot") return "dashDot";
  // dash, sysdash, dashlongdash, etc.
  return "dashed";
}

function mapEndShape(t: string): ConnectorEndShape {
  switch (t) {
    case "none":
      return "none";
    case "triangle":
      return "triangle";
    case "oval":
      return "oval";
    case "arrow":
    case "stealth":
    case "diamond":
      return "arrow";
    default:
      return "arrow";
  }
}

function resolveCxnEndpoint(
  cxnSpid: number | null,
  cxnIdx: string | null,
  which: "start" | "end",
  position: { xEmu: number; yEmu: number } | undefined,
  size: { cxEmu: number; cyEmu: number } | undefined,
  flipH: boolean,
  flipV: boolean
): ConnectorEndpoint {
  if (cxnSpid !== null) {
    return {
      kind: "anchored",
      targetCNvPrId: cxnSpid,
      side: connectorIdxToSide(cxnIdx),
    };
  }
  // Free endpoints — derive from bounding box corners. PowerPoint stores
  // a connector's two endpoints implicitly: when `flipH`/`flipV` are
  // false, start = top-left and end = bottom-right; flips swap which
  // corner each endpoint takes.
  const x = position?.xEmu ?? 0;
  const y = position?.yEmu ?? 0;
  const cx = size?.cxEmu ?? 0;
  const cy = size?.cyEmu ?? 0;
  const startX = flipH ? x + cx : x;
  const startY = flipV ? y + cy : y;
  const endX = flipH ? x : x + cx;
  const endY = flipV ? y : y + cy;
  if (which === "start") return { kind: "free", xEmu: startX, yEmu: startY };
  return { kind: "free", xEmu: endX, yEmu: endY };
}

/**
 * `<a:stCxn @idx>` is shape-kind-specific (rectangles use 0..3, etc.).
 * For our anchor model we collapse to the five named sides — 0/1/2/3
 * are the four cardinal sides for rectangles in PowerPoint's standard
 * connection-site ordering, anything else falls back to "center".
 */
function connectorIdxToSide(idx: string | null): ConnectorSide {
  switch (idx) {
    case "0":
      return "n";
    case "1":
      return "e";
    case "2":
      return "s";
    case "3":
      return "w";
    default:
      return "center";
  }
}

function parsePic(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  partPath: string,
  slideRelTargets: ReadonlyMap<string, string>
): Picture {
  const children = (entry["p:pic"] as unknown[] | undefined) ?? [];
  const nvPicPr = findElementEntry(children, "p:nvPicPr");
  const blipFill = findElementEntry(children, "p:blipFill");
  const spPr = findElementEntry(children, "p:spPr");
  const styleEntry = findElementEntry(children, "p:style");

  let cNvPrId = 0;
  let name = "";
  const nvPicPrTail: OpaqueXml[] = [];
  if (nvPicPr) {
    for (const c of elementEntries((nvPicPr["p:nvPicPr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "p:cNvPr") {
        cNvPrId = Number(attrOf(c, "id") ?? "0");
        name = attrOf(c, "name") ?? "";
      }
      nvPicPrTail.push(captureOpaque(c));
    }
  }

  // r:embed
  let mediaRelId = "";
  const blipFillTail: OpaqueXml[] = [];
  if (blipFill) {
    for (const c of elementEntries((blipFill["p:blipFill"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "a:blip") {
        const e = attrOf(c, "r:embed");
        if (e) mediaRelId = e;
      }
      blipFillTail.push(captureOpaque(c));
    }
  }

  const { position, size, spPrTail } = readSpPr(spPr);

  const targetRel = mediaRelId ? slideRelTargets.get(mediaRelId) : undefined;
  const mediaPartPath = targetRel ? resolveTarget(partPath, targetRel) : "";

  return {
    kind: "pic",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    mediaRelId,
    mediaPartPath,
    nvPicPrTail,
    blipFillTail,
    spPrTail,
    ...(styleEntry ? { styleRaw: captureOpaque(styleEntry) } : {}),
  };
}

function parseGrpSp(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter,
  partPath: string,
  slideRelTargets: ReadonlyMap<string, string>
): GroupShape {
  const children = (entry["p:grpSp"] as unknown[] | undefined) ?? [];
  const nvGrpSpPr = findElementEntry(children, "p:nvGrpSpPr");
  const grpSpPr = findElementEntry(children, "p:grpSpPr");

  let cNvPrId = 0;
  let name = "";
  const nvGrpSpPrTail: OpaqueXml[] = [];
  if (nvGrpSpPr) {
    for (const c of elementEntries((nvGrpSpPr["p:nvGrpSpPr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "p:cNvPr") {
        cNvPrId = Number(attrOf(c, "id") ?? "0");
        name = attrOf(c, "name") ?? "";
      }
      nvGrpSpPrTail.push(captureOpaque(c));
    }
  }

  let position: { xEmu: number; yEmu: number } | undefined;
  let size: { cxEmu: number; cyEmu: number } | undefined;
  const chOffExtRaw: OpaqueXml[] = [];
  const grpSpPrTail: OpaqueXml[] = [];
  if (grpSpPr) {
    for (const c of elementEntries((grpSpPr["p:grpSpPr"] as unknown[] | undefined) ?? [])) {
      const tag = ooxml.getTag(c);
      if (tag === "a:xfrm") {
        const xfrmChildren = (c["a:xfrm"] as unknown[] | undefined) ?? [];
        const off = findElementEntry(xfrmChildren, "a:off");
        const ext = findElementEntry(xfrmChildren, "a:ext");
        if (off) {
          const x = Number(attrOf(off, "x") ?? "0");
          const y = Number(attrOf(off, "y") ?? "0");
          position = { xEmu: x, yEmu: y };
        }
        if (ext) {
          const cx = Number(attrOf(ext, "cx") ?? "0");
          const cy = Number(attrOf(ext, "cy") ?? "0");
          size = { cxEmu: cx, cyEmu: cy };
        }
        const chOff = findElementEntry(xfrmChildren, "a:chOff");
        const chExt = findElementEntry(xfrmChildren, "a:chExt");
        if (chOff) chOffExtRaw.push(captureOpaque(chOff));
        if (chExt) chOffExtRaw.push(captureOpaque(chExt));
      }
      grpSpPrTail.push(captureOpaque(c));
    }
  }

  const groupChildren: Shape[] = [];
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "p:nvGrpSpPr" || tag === "p:grpSpPr") continue;
    groupChildren.push(parseShape(c, mintNodeId, partPath, slideRelTargets));
  }

  return {
    kind: "group",
    id: mintNodeId(),
    cNvPrId,
    name,
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    chOffExtRaw,
    children: groupChildren,
    grpSpPrTail,
    nvGrpSpPrTail,
  };
}

function parseOpaqueShape(entry: Record<string, unknown>, mintNodeId: IdMinter): OpaqueShape {
  const tag = ooxml.getTag(entry);
  return {
    kind: "opaque",
    id: mintNodeId(),
    cNvPrId: 0,
    name: "",
    tag,
    raw: captureOpaque(entry),
  };
}

// ─── nvSpPr / spPr helpers ────────────────────────────────────────────────

function readNvSpPr(nvSpPr: Record<string, unknown> | null): {
  cNvPrId: number;
  name: string;
  nvTail: OpaqueXml[];
} {
  if (!nvSpPr) return { cNvPrId: 0, name: "", nvTail: [] };
  let cNvPrId = 0;
  let name = "";
  const nvTail: OpaqueXml[] = [];
  for (const c of elementEntries((nvSpPr["p:nvSpPr"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(c);
    if (tag === "p:cNvPr") {
      cNvPrId = Number(attrOf(c, "id") ?? "0");
      name = attrOf(c, "name") ?? "";
    }
    nvTail.push(captureOpaque(c));
  }
  return { cNvPrId, name, nvTail };
}

function readPlaceholder(
  nvSpPr: Record<string, unknown> | null
): { type?: string; idx?: number } | undefined {
  if (!nvSpPr) return undefined;
  for (const c of elementEntries((nvSpPr["p:nvSpPr"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(c) !== "p:nvPr") continue;
    for (const c2 of elementEntries((c["p:nvPr"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(c2) !== "p:ph") continue;
      const t = attrOf(c2, "type");
      const idxStr = attrOf(c2, "idx");
      const out: { type?: string; idx?: number } = {};
      if (t) out.type = t;
      if (idxStr !== undefined) out.idx = Number(idxStr);
      return Object.keys(out).length > 0 ? out : undefined;
    }
  }
  return undefined;
}

function readSpPr(spPr: Record<string, unknown> | null): {
  position?: { xEmu: number; yEmu: number };
  size?: { cxEmu: number; cyEmu: number };
  spPrTail: OpaqueXml[];
} {
  if (!spPr) return { spPrTail: [] };
  let position: { xEmu: number; yEmu: number } | undefined;
  let size: { cxEmu: number; cyEmu: number } | undefined;
  const spPrTail: OpaqueXml[] = [];
  for (const c of elementEntries((spPr["p:spPr"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(c);
    if (tag === "a:xfrm") {
      const xfrmChildren = (c["a:xfrm"] as unknown[] | undefined) ?? [];
      const off = findElementEntry(xfrmChildren, "a:off");
      const ext = findElementEntry(xfrmChildren, "a:ext");
      if (off) {
        const x = Number(attrOf(off, "x") ?? "0");
        const y = Number(attrOf(off, "y") ?? "0");
        position = { xEmu: x, yEmu: y };
      }
      if (ext) {
        const cx = Number(attrOf(ext, "cx") ?? "0");
        const cy = Number(attrOf(ext, "cy") ?? "0");
        size = { cxEmu: cx, cyEmu: cy };
      }
    }
    spPrTail.push(captureOpaque(c));
  }
  return {
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
    spPrTail,
  };
}

// ─── Text body ────────────────────────────────────────────────────────────

function emptyTextBody(): TextBody {
  return { paragraphs: [] };
}

function parseTextBody(entry: Record<string, unknown>, mintNodeId: IdMinter): TextBody {
  const children = (entry["p:txBody"] as unknown[] | undefined) ?? [];
  return parseTextBodyChildren(children, mintNodeId);
}

/**
 * Parse a text-body's children. Used both for `<p:txBody>` (inside `<p:sp>`)
 * and `<a:txBody>` (inside table `<a:tc>`); the children themselves are
 * always in the `a:` namespace so the wrapper tag doesn't matter.
 */
function parseTextBodyChildren(children: ReadonlyArray<unknown>, mintNodeId: IdMinter): TextBody {
  const bodyPr = findElementEntry(children, "a:bodyPr");
  const lstStyle = findElementEntry(children, "a:lstStyle");
  const paragraphs: TextParagraph[] = [];
  for (const c of elementEntries(children)) {
    if (ooxml.getTag(c) === "a:p") paragraphs.push(parseParagraph(c, mintNodeId));
  }
  return {
    ...(bodyPr ? { bodyPrRaw: captureOpaque(bodyPr) } : {}),
    ...(lstStyle ? { lstStyleRaw: captureOpaque(lstStyle) } : {}),
    paragraphs,
  };
}

function parseParagraph(entry: Record<string, unknown>, mintNodeId: IdMinter): TextParagraph {
  const children = (entry["a:p"] as unknown[] | undefined) ?? [];
  const pPr = findElementEntry(children, "a:pPr");
  const properties: TextParagraphProperties = pPr ? parseParagraphProperties(pPr) : {};
  const runs: TextRun[] = [];
  let endParaRPrRaw: OpaqueXml | undefined;
  for (const c of elementEntries(children)) {
    const tag = ooxml.getTag(c);
    if (tag === "a:pPr") continue;
    if (tag === "a:r") {
      runs.push(parseRun(c, mintNodeId));
    } else if (tag === "a:br") {
      const rPr = findElementEntry((c["a:br"] as unknown[] | undefined) ?? [], "a:rPr");
      runs.push({
        id: mintNodeId(),
        properties: rPr ? parseRunProperties(rPr) : {},
        text: "",
        isLineBreak: true,
      });
    } else if (tag === "a:fld") {
      // Fields: capture text-equivalent and treat properties opaquely.
      const txt = readText(c);
      runs.push({
        id: mintNodeId(),
        properties: {
          opaqueAttrs: readRootAttrs(c),
        },
        text: txt,
      });
    } else if (tag === "a:endParaRPr") {
      endParaRPrRaw = captureOpaque(c);
    }
  }
  return {
    id: mintNodeId(),
    properties,
    runs,
    ...(endParaRPrRaw ? { endParaRPrRaw } : {}),
  };
}

function parseParagraphProperties(entry: Record<string, unknown>): TextParagraphProperties {
  const attrs = readRootAttrs(entry);
  const props: { -readonly [K in keyof TextParagraphProperties]: TextParagraphProperties[K] } = {};
  if (attrs.lvl !== undefined) props.level = Number(attrs.lvl);
  const algn = attrs.algn;
  if (algn) {
    const map: Record<string, NonNullable<TextParagraphProperties["alignment"]>> = {
      l: "left",
      ctr: "center",
      r: "right",
      just: "justify",
    };
    if (map[algn]) props.alignment = map[algn];
  }
  // Preserve all attrs opaquely so re-emit is byte-faithful (we only USE a
  // subset, but re-emit all of them to avoid losing things like `marL`).
  if (Object.keys(attrs).length > 0) props.opaqueAttrs = attrs;
  const opaqueChildren: OpaqueXml[] = [];
  for (const c of elementEntries((entry["a:pPr"] as unknown[] | undefined) ?? [])) {
    opaqueChildren.push(captureOpaque(c));
  }
  if (opaqueChildren.length > 0) props.opaqueChildren = opaqueChildren;
  return props;
}

function parseRun(entry: Record<string, unknown>, mintNodeId: IdMinter): TextRun {
  const children = (entry["a:r"] as unknown[] | undefined) ?? [];
  const rPr = findElementEntry(children, "a:rPr");
  const properties = rPr ? parseRunProperties(rPr) : {};
  const tEntry = findElementEntry(children, "a:t");
  const text = tEntry ? readText(tEntry) : "";
  return { id: mintNodeId(), properties, text };
}

function parseRunProperties(entry: Record<string, unknown>): TextRunProperties {
  const attrs = readRootAttrs(entry);
  const props: { -readonly [K in keyof TextRunProperties]: TextRunProperties[K] } = {};
  if (attrs.b === "1" || attrs.b === "true") props.bold = true;
  if (attrs.b === "0" || attrs.b === "false") props.bold = false;
  if (attrs.i === "1" || attrs.i === "true") props.italic = true;
  if (attrs.i === "0" || attrs.i === "false") props.italic = false;
  if (attrs.u !== undefined) {
    if (attrs.u === "none") props.underline = false;
    else if (attrs.u === "sng") props.underline = true;
    else props.underline = attrs.u;
  }
  if (attrs.strike !== undefined) {
    props.strike = attrs.strike !== "noStrike";
  }
  if (attrs.sz !== undefined) props.fontSizeHundredths = Number(attrs.sz);
  if (Object.keys(attrs).length > 0) props.opaqueAttrs = attrs;

  // Inspect children for color (a:solidFill > a:srgbClr) and font.
  const opaqueChildren: OpaqueXml[] = [];
  for (const c of elementEntries((entry["a:rPr"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(c);
    if (tag === "a:solidFill") {
      const fill = (c["a:solidFill"] as unknown[] | undefined) ?? [];
      const srgb = findElementEntry(fill, "a:srgbClr");
      if (srgb) {
        const v = attrOf(srgb, "val");
        if (v) props.color = v;
      }
    } else if (tag === "a:latin") {
      const v = attrOf(c, "typeface");
      if (v) props.fontFamily = v;
    } else if (tag === "a:highlight") {
      const hl = (c["a:highlight"] as unknown[] | undefined) ?? [];
      const srgb = findElementEntry(hl, "a:srgbClr");
      if (srgb) {
        const v = attrOf(srgb, "val");
        if (v) props.highlight = v;
      }
    }
    opaqueChildren.push(captureOpaque(c));
  }
  if (opaqueChildren.length > 0) props.opaqueChildren = opaqueChildren;
  return props;
}

// ─── Misc helpers ─────────────────────────────────────────────────────────

function readSlideSize(presChildren: ReadonlyArray<unknown>): SlideSize | undefined {
  const e = findElementEntry(presChildren, "p:sldSz");
  if (!e) return undefined;
  const cx = Number(attrOf(e, "cx") ?? "0");
  const cy = Number(attrOf(e, "cy") ?? "0");
  const t = attrOf(e, "type");
  return { cxEmu: cx, cyEmu: cy, ...(t ? { type: t } : {}) };
}

function readNotesSize(presChildren: ReadonlyArray<unknown>): SlideSize | undefined {
  const e = findElementEntry(presChildren, "p:notesSz");
  if (!e) return undefined;
  const cx = Number(attrOf(e, "cx") ?? "0");
  const cy = Number(attrOf(e, "cy") ?? "0");
  return { cxEmu: cx, cyEmu: cy };
}

function collectPartsByType(graph: ooxml.RelationshipGraph, type: string): string[] {
  const out: string[] = [];
  const sourcePart = relsPathToOwnerPath(graph.relsPath);
  for (const r of graph.relationships) {
    if (r.type !== type) continue;
    out.push(resolveTarget(sourcePart ?? "", r.target));
  }
  return out;
}

/**
 * Resolve an OPC `Target` (relative path from the rels' owner part) to a
 * package-absolute path. Per OPC: if Target starts with `/`, it's absolute.
 */
export function resolveTarget(ownerPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  // Owner directory is the directory containing the OWNER part (not the rels file).
  const slash = ownerPath.lastIndexOf("/");
  const ownerDir = slash >= 0 ? ownerPath.slice(0, slash) : "";
  const segments = (ownerDir ? `${ownerDir}/${target}` : target).split("/");
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join("/");
}

function relsPathToOwnerPath(relsPath: string): string | null {
  // word/_rels/document.xml.rels  →  word/document.xml
  // _rels/.rels                   →  ""
  const m = /^(.*\/)?_rels\/(.+)\.rels$/.exec(relsPath);
  if (!m) return null;
  const dir = m[1] ?? "";
  const file = m[2] ?? "";
  if (dir === "" && file === "") return "";
  return `${dir}${file}`;
}

function mediaContentType(ext: string): string {
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "svg":
      return "image/svg+xml";
    case "tiff":
    case "tif":
      return "image/tiff";
    case "webp":
      return "image/webp";
    default:
      return "application/octet-stream";
  }
}

/**
 * Parse a `<p:sldLayout>` part into a typed `SlideLayout`. We extract
 * (a) the layout name from `<p:cSld name>`, (b) every `<p:ph>` declared
 * inside `<p:sld>`/`<p:spTree>` (to know which placeholders new slides
 * should clone), and (c) classify the layout into one of the standard
 * 11 PowerPoint layouts so the picker UI can show consistent labels.
 *
 * The verbatim `<p:sldLayout>` blob is preserved in `raw` so unedited
 * layouts round-trip byte-for-byte; the typed fields are derived.
 */
function parseSlideLayout(container: ooxml.OoxmlContainer, partPath: string): SlideLayout {
  return parseSlideLayoutFromXml(partPath, container.readText(partPath));
}

/**
 * String-input variant of `parseSlideLayout`. Used by the layout-cloning
 * helper to materialise a `SlideLayout` from a built-in XML template
 * without a backing container part.
 */
export function parseSlideLayoutFromXml(partPath: string, xml: string): SlideLayout {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(xml);
  } catch (err) {
    throw new PptxParseError("invalid-xml", `Failed to parse layout ${partPath}`, {
      partPath,
      cause: err,
    });
  }
  if (!Array.isArray(tree)) {
    throw new PptxParseError("invalid-xml", `Expected XML tree at ${partPath}`, { partPath });
  }
  const r = findElementEntry(tree as unknown[], "p:sldLayout");
  if (!r) {
    throw new PptxParseError("invalid-xml", `Missing <p:sldLayout> in ${partPath}`, {
      partPath,
    });
  }
  const raw: OpaqueXml = {
    tag: "p:sldLayout",
    attrs: readRootAttrs(r),
    rawAttrs: readRawAttrs(r),
    subtree: (r["p:sldLayout"] as unknown[] | undefined) ?? [],
  };
  const placeholders = collectLayoutPlaceholders(raw);
  const name = readLayoutName(raw) ?? defaultLayoutName(placeholders);
  const kind = classifyLayoutKind(raw, placeholders);
  return { partPath, kind, name, placeholders, raw };
}

function collectLayoutPlaceholders(raw: OpaqueXml): PlaceholderSpec[] {
  const cSld = findElementEntry(raw.subtree, "p:cSld");
  if (!cSld) return [];
  const spTree = findElementEntry((cSld["p:cSld"] as unknown[] | undefined) ?? [], "p:spTree");
  if (!spTree) return [];
  const out: PlaceholderSpec[] = [];
  for (const child of elementEntries((spTree["p:spTree"] as unknown[] | undefined) ?? [])) {
    const tag = ooxml.getTag(child);
    if (tag !== "p:sp") continue;
    const ph = readLayoutPlaceholder(child);
    if (ph) out.push(ph);
  }
  return out;
}

function readLayoutPlaceholder(spEntry: Record<string, unknown>): PlaceholderSpec | null {
  const spChildren = (spEntry["p:sp"] as unknown[] | undefined) ?? [];
  const nvSpPr = findElementEntry(spChildren, "p:nvSpPr");
  if (!nvSpPr) return null;
  let phEntry: Record<string, unknown> | null = null;
  for (const c of elementEntries((nvSpPr["p:nvSpPr"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(c) !== "p:nvPr") continue;
    for (const c2 of elementEntries((c["p:nvPr"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(c2) === "p:ph") {
        phEntry = c2;
        break;
      }
    }
    if (phEntry) break;
  }
  if (!phEntry) return null;
  const type = attrOf(phEntry, "type") ?? "body";
  const idxStr = attrOf(phEntry, "idx");
  const sz = attrOf(phEntry, "sz");
  const idx = idxStr !== undefined && /^\d+$/.test(idxStr) ? Number(idxStr) : 0;

  let position: Position | undefined;
  let size: Size | undefined;
  const spPr = findElementEntry(spChildren, "p:spPr");
  if (spPr) {
    for (const c of elementEntries((spPr["p:spPr"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(c) !== "a:xfrm") continue;
      const xfrmChildren = (c["a:xfrm"] as unknown[] | undefined) ?? [];
      const off = findElementEntry(xfrmChildren, "a:off");
      const ext = findElementEntry(xfrmChildren, "a:ext");
      if (off) {
        position = {
          xEmu: Number(attrOf(off, "x") ?? "0"),
          yEmu: Number(attrOf(off, "y") ?? "0"),
        };
      }
      if (ext) {
        size = {
          cxEmu: Number(attrOf(ext, "cx") ?? "0"),
          cyEmu: Number(attrOf(ext, "cy") ?? "0"),
        };
      }
    }
  }
  const out: PlaceholderSpec = {
    type,
    idx,
    ...(sz ? { sz } : {}),
    ...(position ? { position } : {}),
    ...(size ? { size } : {}),
  };
  return out;
}

function readLayoutName(raw: OpaqueXml): string | null {
  const cSld = findElementEntry(raw.subtree, "p:cSld");
  if (!cSld) return null;
  const attrs = readRootAttrs(cSld);
  return attrs.name ?? null;
}

function defaultLayoutName(placeholders: ReadonlyArray<PlaceholderSpec>): string {
  if (placeholders.length === 0) return "Blank";
  if (placeholders.length === 1) return "Title Only";
  return "Layout";
}

/**
 * Classify a layout from its placeholder set + name. PowerPoint stores
 * the layout's official type in `<p:sldLayout type="...">` (attribute on
 * the root element), so we honour that when present and fall back to a
 * heuristic over the placeholder shape otherwise.
 *
 * The `type` attribute → `LayoutKind` mapping mirrors the values the
 * PowerPoint UI emits: `title`, `obj` (titleAndContent), `secHead`,
 * `twoObj` (twoContent), `twoTxTwoObj` (comparison), `titleOnly`, `blank`,
 * `objTx` (contentWithCaption), `picTx` (pictureWithCaption), `tx`
 * (titleSlide). Unknowns surface as `unknown` so the picker can still
 * round-trip them.
 */
function classifyLayoutKind(
  raw: OpaqueXml,
  placeholders: ReadonlyArray<PlaceholderSpec>
): import("../model/types.js").LayoutKind {
  const t = raw.attrs.type ?? raw.rawAttrs["@_type"];
  switch (t) {
    case "title":
      return "title";
    case "obj":
      return "titleAndContent";
    case "secHead":
      return "sectionHeader";
    case "twoObj":
      return "twoContent";
    case "twoTxTwoObj":
    case "fourObj":
      return "comparison";
    case "titleOnly":
      return "titleOnly";
    case "blank":
      return "blank";
    case "objTx":
      return "contentWithCaption";
    case "picTx":
      return "pictureWithCaption";
    case "tx":
      return "titleSlide";
    default:
      break;
  }
  // Heuristic fallback when no `type` attribute is set.
  if (placeholders.length === 0) return "blank";
  const types = new Set(placeholders.map((p) => p.type));
  if (types.size === 1 && types.has("title")) return "titleOnly";
  if (placeholders.length >= 4) return "comparison";
  if (placeholders.length === 3) return "twoContent";
  return "titleAndContent";
}

/**
 * Parse a `<p:notes>` part into a typed `NotesSlide`. We pull out the
 * `<p:txBody>` of the body placeholder so the speaker-notes panel can
 * read/write it as structured text; everything else (slide image
 * placeholder, header/footer placeholders, formatting) lives verbatim
 * in `raw` for byte-faithful round-trip when nothing has changed.
 */
function parseNotesSlide(
  container: ooxml.OoxmlContainer,
  partPath: string,
  mintNodeId: IdMinter
): NotesSlide {
  const opaque = opaquePartFor(container, partPath, "p:notes");
  const body = extractNotesBody(opaque.raw, mintNodeId) ?? { paragraphs: [] };
  return { partPath, body, raw: opaque.raw };
}

function extractNotesBody(raw: OpaqueXml, mintNodeId: IdMinter): TextBody | null {
  const cSld = findElementEntry(raw.subtree, "p:cSld");
  if (!cSld) return null;
  const spTree = findElementEntry((cSld["p:cSld"] as unknown[] | undefined) ?? [], "p:spTree");
  if (!spTree) return null;
  for (const child of elementEntries((spTree["p:spTree"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(child) !== "p:sp") continue;
    const spChildren = (child["p:sp"] as unknown[] | undefined) ?? [];
    const nvSpPr = findElementEntry(spChildren, "p:nvSpPr");
    if (!nvSpPr) continue;
    let phType: string | null = null;
    for (const c of elementEntries((nvSpPr["p:nvSpPr"] as unknown[] | undefined) ?? [])) {
      if (ooxml.getTag(c) !== "p:nvPr") continue;
      for (const c2 of elementEntries((c["p:nvPr"] as unknown[] | undefined) ?? [])) {
        if (ooxml.getTag(c2) !== "p:ph") continue;
        phType = attrOf(c2, "type") ?? "body";
      }
    }
    if (phType !== "body") continue;
    const txBody = findElementEntry(spChildren, "p:txBody");
    if (txBody) return parseTextBody(txBody, mintNodeId);
  }
  return null;
}

/**
 * Parse `ppt/comments/commentN.xml`. PowerPoint stores per-slide
 * comments in a `<p:cmLst>` document; each `<p:cm>` carries
 * `authorId`, `idx` (per-author monotonic), `dt` (creation time),
 * a child `<p:pos x="…" y="…"/>` pin location and a `<p:text>` body.
 *
 * We synthesise stable comment ids from `${authorId}:${idx}` so they
 * survive a round-trip without depending on PowerPoint minting global
 * UUIDs (which the legacy format doesn't). Replies are encoded as
 * extLst `<p:ext uri="parent">…</p:ext>` siblings — non-standard, but
 * faithful to the way PowerPoint represents them in the modern
 * "modernComments" schema we're shimming over the legacy one.
 */
function parseCommentsPart(
  container: ooxml.OoxmlContainer,
  partPath: string,
  _mintNodeId: IdMinter
): PptxCommentsPart {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(partPath));
  } catch (err) {
    throw new PptxParseError("invalid-xml", `Failed to parse ${partPath}`, { partPath, cause: err });
  }
  if (!Array.isArray(tree)) return { partPath, comments: [] };
  const cmLst = findElementEntry(tree as unknown[], "p:cmLst");
  if (!cmLst) return { partPath, comments: [] };
  const out: PptxComment[] = [];
  for (const child of elementEntries((cmLst["p:cmLst"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(child) !== "p:cm") continue;
    const cm = child;
    const authorId = Number(attrOf(cm, "authorId") ?? "0");
    const idx = Number(attrOf(cm, "idx") ?? "1");
    const createdAt = attrOf(cm, "dt") ?? undefined;
    const cmChildren = (cm["p:cm"] as unknown[] | undefined) ?? [];
    let xEmu = 0;
    let yEmu = 0;
    let text = "";
    let parentId: string | undefined;
    let resolved: boolean | undefined;
    let shapeId: string | undefined;
    for (const ch of elementEntries(cmChildren)) {
      const t = ooxml.getTag(ch);
      if (t === "p:pos") {
        // PowerPoint stores comment pin coordinates in 1/100 of a point;
        // convert to EMU (1 pt = 12700 EMU, so 1/100 pt = 127 EMU).
        const xRaw = Number(attrOf(ch, "x") ?? "0");
        const yRaw = Number(attrOf(ch, "y") ?? "0");
        xEmu = Math.round(xRaw * 127);
        yEmu = Math.round(yRaw * 127);
      } else if (t === "p:text") {
        const inner = (ch["p:text"] as unknown[] | undefined) ?? [];
        for (const tn of inner) {
          if (tn && typeof tn === "object" && !Array.isArray(tn)) {
            const obj = tn as Record<string, unknown>;
            const v = obj["#text"];
            if (typeof v === "string") text += v;
          }
        }
      } else if (t === "p:extLst") {
        // Walk extensions for our parent-ref/resolved flag.
        for (const ext of elementEntries((ch["p:extLst"] as unknown[] | undefined) ?? [])) {
          if (ooxml.getTag(ext) !== "p:ext") continue;
          const uri = attrOf(ext, "uri");
          if (uri === "officeai:parent") {
            const pid = attrOf(ext, "id");
            if (pid) parentId = pid;
          } else if (uri === "officeai:resolved") {
            resolved = attrOf(ext, "value") === "1";
          } else if (uri === "officeai:shapeAnchor") {
            const sid = attrOf(ext, "id");
            if (sid) shapeId = sid;
          }
        }
      }
    }
    out.push({
      id: `${authorId}:${idx}`,
      authorId,
      idx,
      ...(createdAt ? { createdAt } : {}),
      xEmu,
      yEmu,
      text,
      ...(parentId ? { parentId } : {}),
      ...(resolved !== undefined ? { resolved } : {}),
      ...(shapeId ? { shapeId } : {}),
    });
  }
  return { partPath, comments: out };
}

/**
 * Parse `ppt/commentAuthors.xml`. Authors are identified by a numeric
 * `id`; PowerPoint additionally tracks `lastIdx` (highest comment idx
 * minted by this author) and `clrIdx` (palette slot). We round-trip
 * everything but only `name` and `id` matter to the UI.
 */
function parseCommentAuthorsPart(container: ooxml.OoxmlContainer, partPath: string): PptxCommentAuthorsPart {
  let tree: unknown;
  try {
    tree = ooxml.parseXml(container.readText(partPath));
  } catch (err) {
    throw new PptxParseError("invalid-xml", `Failed to parse ${partPath}`, { partPath, cause: err });
  }
  if (!Array.isArray(tree)) return { partPath, authors: [] };
  const root = findElementEntry(tree as unknown[], "p:cmAuthorLst");
  if (!root) return { partPath, authors: [] };
  const authors: PptxCommentAuthor[] = [];
  for (const child of elementEntries((root["p:cmAuthorLst"] as unknown[] | undefined) ?? [])) {
    if (ooxml.getTag(child) !== "p:cmAuthor") continue;
    const id = Number(attrOf(child, "id") ?? "0");
    const name = attrOf(child, "name") ?? "";
    const initials = attrOf(child, "initials") ?? undefined;
    const lastIdx = attrOf(child, "lastIdx");
    const clrIdx = attrOf(child, "clrIdx");
    authors.push({
      id,
      name,
      ...(initials ? { initials } : {}),
      ...(lastIdx ? { lastIdx: Number(lastIdx) } : {}),
      ...(clrIdx ? { clrIdx: Number(clrIdx) } : {}),
    });
  }
  return { partPath, authors };
}

function opaquePartFor(container: ooxml.OoxmlContainer, partPath: string, rootTag: string): OpaquePart {
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
    throw new PptxParseError("invalid-xml", `Missing <${rootTag}> in ${partPath}`, {
      partPath,
    });
  }
  return {
    partPath,
    raw: {
      tag: r ? rootTag : "",
      attrs: r ? readRootAttrs(r) : {},
      rawAttrs: r ? readRawAttrs(r) : {},
      subtree: r ? ((r[rootTag] as unknown[] | undefined) ?? []) : [],
    },
  };
}

// Resolve picture rel → media path AFTER all slides parsed, so callers
// can do this if needed. We expose this helper for the agent layer.
export function resolvePictureMediaPath(
  container: ooxml.OoxmlContainer,
  slidePartPath: string,
  relId: string
): string | null {
  const graph = ooxml.RelationshipGraph.loadFor(container, slidePartPath);
  const r = graph.byId(relId);
  if (!r || r.type !== REL_TYPE_IMAGE) return null;
  return resolveTarget(slidePartPath, r.target);
}

export {
  REL_TYPE_CHART,
  REL_TYPE_IMAGE,
  REL_TYPE_NOTES_SLIDE,
  REL_TYPE_SLIDE,
  REL_TYPE_SLIDE_LAYOUT,
  REL_TYPE_SLIDE_MASTER,
  REL_TYPE_THEME,
};

function walkShapesForCharts(shapes: ReadonlyArray<Shape>, visit: (chartPartPath: string) => void): void {
  for (const sh of shapes) {
    if (sh.kind === "chart" && sh.chartPartPath) visit(sh.chartPartPath);
    if (sh.kind === "group") walkShapesForCharts(sh.children, visit);
  }
}

const CHART_TYPE_TAGS: ReadonlyArray<{ tag: string; type: ChartType }> = [
  { tag: "c:barChart", type: "bar" },
  { tag: "c:bar3DChart", type: "bar" },
  { tag: "c:lineChart", type: "line" },
  { tag: "c:line3DChart", type: "line" },
  { tag: "c:pieChart", type: "pie" },
  { tag: "c:pie3DChart", type: "pie" },
  { tag: "c:doughnutChart", type: "pie" },
  { tag: "c:areaChart", type: "area" },
  { tag: "c:area3DChart", type: "area" },
];

/**
 * Parse `ppt/charts/chart{N}.xml` ⇒ `ChartPart` with typed title, type,
 * categories, and series. Everything we don't model is preserved verbatim
 * so the serializer can rebuild the part byte-for-byte unless explicitly
 * dirtied.
 */
function parseChartPart(partPath: string, xml: string, contentType: string, mintNodeId: IdMinter): ChartPart {
  const tree = ooxml.parseXml(xml) as unknown[];
  const chartSpace = findElementEntry(tree, "c:chartSpace");
  if (!chartSpace) {
    throw new PptxParseError("invalid-xml", `Missing <c:chartSpace> in ${partPath}`, {
      partPath,
    });
  }
  const chartSpaceChildren = (chartSpace["c:chartSpace"] as unknown[] | undefined) ?? [];
  const chart = findElementEntry(chartSpaceChildren, "c:chart");
  const chartChildren = chart ? ((chart["c:chart"] as unknown[] | undefined) ?? []) : [];
  const title = readChartTitle(chartChildren);
  const plotArea = findElementEntry(chartChildren, "c:plotArea");
  const plotAreaChildren = plotArea ? ((plotArea["c:plotArea"] as unknown[] | undefined) ?? []) : [];

  let chartType: ChartType = "unsupported";
  let chartTypeEntry: Record<string, unknown> | null = null;
  for (const cand of CHART_TYPE_TAGS) {
    const e = findElementEntry(plotAreaChildren, cand.tag);
    if (e) {
      chartType = cand.type;
      chartTypeEntry = e;
      break;
    }
  }

  const plotAreaTailRaw: OpaqueXml[] = [];
  for (const child of elementEntries(plotAreaChildren)) {
    const tag = ooxml.getTag(child);
    if (chartTypeEntry && tag === ooxml.getTag(chartTypeEntry)) continue;
    plotAreaTailRaw.push(captureOpaque(child));
  }

  const series: ChartSeries[] = [];
  const seriesRaw = new Map<number, OpaqueXml>();
  let categories: ReadonlyArray<string> = [];
  if (chartTypeEntry) {
    const ctChildren = (chartTypeEntry[ooxml.getTag(chartTypeEntry)] as unknown[] | undefined) ?? [];
    for (const sEntry of elementEntries(ctChildren)) {
      if (ooxml.getTag(sEntry) !== "c:ser") continue;
      const sChildren = (sEntry["c:ser"] as unknown[] | undefined) ?? [];
      const idxEntry = findElementEntry(sChildren, "c:idx");
      const idx = idxEntry ? Number(attrOf(idxEntry, "val") ?? "0") : series.length;
      const txEntry = findElementEntry(sChildren, "c:tx");
      const name = txEntry ? readSeriesText(txEntry) : undefined;
      const valEntry = findElementEntry(sChildren, "c:val");
      const values = valEntry ? readNumericCache(valEntry, "c:val") : [];
      const catEntry = findElementEntry(sChildren, "c:cat");
      if (categories.length === 0 && catEntry) {
        categories = readStringCache(catEntry, "c:cat");
      }
      series.push({
        id: mintNodeId(),
        idx,
        ...(name !== undefined ? { name } : {}),
        values,
      });
      seriesRaw.set(idx, captureOpaque(sEntry));
    }
  }

  return {
    partPath,
    contentType,
    chartType,
    ...(title !== undefined ? { title } : {}),
    categories,
    series,
    chartSpaceRaw: captureOpaque(chartSpace),
    plotAreaTailRaw,
    seriesRaw,
  };
}

function readChartTitle(chartChildren: ReadonlyArray<unknown>): string | undefined {
  const titleEntry = findElementEntry(chartChildren, "c:title");
  if (!titleEntry) return undefined;
  const tx = findElementEntry((titleEntry["c:title"] as unknown[] | undefined) ?? [], "c:tx");
  if (!tx) return undefined;
  const rich = findElementEntry((tx["c:tx"] as unknown[] | undefined) ?? [], "c:rich");
  if (!rich) {
    const strRef = findElementEntry((tx["c:tx"] as unknown[] | undefined) ?? [], "c:strRef");
    if (strRef) {
      const f = findElementEntry((strRef["c:strRef"] as unknown[] | undefined) ?? [], "c:f");
      if (f) return readText(f).trim() || undefined;
    }
    return undefined;
  }
  const richChildren = (rich["c:rich"] as unknown[] | undefined) ?? [];
  const out: string[] = [];
  for (const p of elementEntries(richChildren)) {
    if (ooxml.getTag(p) !== "a:p") continue;
    const pChildren = (p["a:p"] as unknown[] | undefined) ?? [];
    for (const r of elementEntries(pChildren)) {
      if (ooxml.getTag(r) !== "a:r") continue;
      const rChildren = (r["a:r"] as unknown[] | undefined) ?? [];
      const tEntry = findElementEntry(rChildren, "a:t");
      if (tEntry) out.push(readText(tEntry));
    }
  }
  const txt = out.join("").trim();
  return txt.length > 0 ? txt : undefined;
}

function readSeriesText(txEntry: Record<string, unknown>): string | undefined {
  const txChildren = (txEntry["c:tx"] as unknown[] | undefined) ?? [];
  const strRef = findElementEntry(txChildren, "c:strRef");
  if (strRef) {
    const cache = findElementEntry((strRef["c:strRef"] as unknown[] | undefined) ?? [], "c:strCache");
    if (cache) {
      const pt = findElementEntry((cache["c:strCache"] as unknown[] | undefined) ?? [], "c:pt");
      if (pt) {
        const v = findElementEntry((pt["c:pt"] as unknown[] | undefined) ?? [], "c:v");
        if (v) return readText(v);
      }
    }
  }
  const v = findElementEntry(txChildren, "c:v");
  if (v) return readText(v);
  return undefined;
}

function readNumericCache(entry: Record<string, unknown>, parentTag: string): number[] {
  const children = (entry[parentTag] as unknown[] | undefined) ?? [];
  const ref = findElementEntry(children, "c:numRef");
  const lit = findElementEntry(children, "c:numLit");
  const cacheParent = ref ?? lit;
  if (!cacheParent) return [];
  const cacheChildren = (cacheParent[ooxml.getTag(cacheParent)] as unknown[] | undefined) ?? [];
  const cache = findElementEntry(cacheChildren, "c:numCache") ?? cacheParent;
  const cacheNodeChildren = (cache[ooxml.getTag(cache)] as unknown[] | undefined) ?? cacheChildren;
  const out: number[] = [];
  for (const pt of elementEntries(cacheNodeChildren)) {
    if (ooxml.getTag(pt) !== "c:pt") continue;
    const idx = Number(attrOf(pt, "idx") ?? `${out.length}`);
    const v = findElementEntry((pt["c:pt"] as unknown[] | undefined) ?? [], "c:v");
    if (v) out[idx] = Number(readText(v));
  }
  return out;
}

function readStringCache(entry: Record<string, unknown>, parentTag: string): string[] {
  const children = (entry[parentTag] as unknown[] | undefined) ?? [];
  const ref = findElementEntry(children, "c:strRef");
  const lit = findElementEntry(children, "c:strLit");
  const multi = findElementEntry(children, "c:multiLvlStrRef");
  const cacheParent = ref ?? lit ?? multi;
  if (!cacheParent) return [];
  const cacheChildren = (cacheParent[ooxml.getTag(cacheParent)] as unknown[] | undefined) ?? [];
  const cache =
    findElementEntry(cacheChildren, "c:strCache") ??
    findElementEntry(cacheChildren, "c:multiLvlStrCache") ??
    cacheParent;
  const cacheNodeChildren = (cache[ooxml.getTag(cache)] as unknown[] | undefined) ?? cacheChildren;
  const out: string[] = [];
  for (const node of elementEntries(cacheNodeChildren)) {
    const tag = ooxml.getTag(node);
    if (tag === "c:pt") {
      const idx = Number(attrOf(node, "idx") ?? `${out.length}`);
      const v = findElementEntry((node["c:pt"] as unknown[] | undefined) ?? [], "c:v");
      if (v) out[idx] = readText(v);
    } else if (tag === "c:lvl") {
      const lvlChildren = (node["c:lvl"] as unknown[] | undefined) ?? [];
      for (const pt of elementEntries(lvlChildren)) {
        if (ooxml.getTag(pt) !== "c:pt") continue;
        const idx = Number(attrOf(pt, "idx") ?? `${out.length}`);
        const v = findElementEntry((pt["c:pt"] as unknown[] | undefined) ?? [], "c:v");
        if (v && out[idx] === undefined) out[idx] = readText(v);
      }
    }
  }
  return out;
}

// ── F4: animations & transitions parsers ─────────────────────────────────

const TRANSITION_KIND_TAGS: ReadonlyArray<{ tag: string; kind: TransitionKind }> = [
  { tag: "p:fade", kind: "fade" },
  { tag: "p:push", kind: "push" },
  { tag: "p:wipe", kind: "wipe" },
  { tag: "p:split", kind: "split" },
  { tag: "p:cut", kind: "cut" },
];

const ENTRANCE_PRESET_CLASS = "entr";
const ENTRANCE_PRESET_IDS: ReadonlyMap<number, EntranceEffect> = new Map([
  [1, "appear"],
  [2, "fly-in"],
  [3, "fade"],
  [10, "wipe"],
]);

function parseSlideTransition(entry: Record<string, unknown>, mintNodeId: IdMinter): SlideTransition {
  const children = (entry["p:transition"] as unknown[] | undefined) ?? [];
  const speedAttr = attrOf(entry, "spd");
  const speed: TransitionSpeed | undefined =
    speedAttr === "slow" || speedAttr === "med" || speedAttr === "fast" ? speedAttr : undefined;
  let kind: TransitionKind = "unsupported";
  for (const cand of TRANSITION_KIND_TAGS) {
    if (findElementEntry(children, cand.tag)) {
      kind = cand.kind;
      break;
    }
  }
  return {
    id: mintNodeId(),
    kind,
    ...(speed ? { speed } : {}),
    raw: captureOpaque(entry),
  };
}

/**
 * Parse `<p:timing>` and promote a flat list of typed entrance animations.
 * Anything we can't model is preserved as the raw `<p:timing>` blob — the
 * serializer re-emits the raw verbatim when no commands have touched the
 * typed list.
 *
 * The walk is intentionally tolerant: PowerPoint nests timing nodes deeply
 * (`p:tnLst` → `p:par` → multiple `p:childTnLst` → … → `p:set` / `p:anim`),
 * so we recursively look for `p:cTn @presetClass="entr"` carriers and
 * resolve their `<p:spTgt @spid>` in any descendant.
 */
function parseSlideTiming(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter
): { animations: EntranceAnimation[]; tail: OpaqueXml | undefined } {
  const animations: EntranceAnimation[] = [];
  const visit = (node: unknown): void => {
    if (!node || typeof node !== "object") return;
    const e = node as Record<string, unknown>;
    const tag = ooxml.getTag(e);
    if (tag === "p:cTn") {
      const presetClass = attrOf(e, "presetClass");
      const presetId = Number(attrOf(e, "presetID") ?? "0");
      if (presetClass === ENTRANCE_PRESET_CLASS && ENTRANCE_PRESET_IDS.has(presetId)) {
        const spid = findSpTgtSpid(e);
        if (spid !== null) {
          const dur = attrOf(e, "dur");
          const effect = ENTRANCE_PRESET_IDS.get(presetId);
          if (effect) {
            animations.push({
              id: mintNodeId(),
              targetCNvPrId: spid,
              effect,
              ...(dur && /^\d+$/.test(dur) ? { durationMs: Number(dur) } : {}),
              order: animations.length,
            });
          }
        }
      }
    }
    const children = e[tag] as unknown[] | undefined;
    if (Array.isArray(children)) {
      for (const c of children) visit(c);
    }
  };
  visit(entry);
  return { animations, tail: captureOpaque(entry) };
}

/**
 * Find the first `<p:spTgt @spid>` value anywhere under the given timing
 * node. This is intentionally a recursive scan because `p:spTgt` can sit
 * many levels deep inside `p:tgtEl` → `p:cBhvr` → … wrappers.
 */
function findSpTgtSpid(node: unknown): number | null {
  if (!node || typeof node !== "object") return null;
  const e = node as Record<string, unknown>;
  const tag = ooxml.getTag(e);
  if (tag === "p:spTgt") {
    const spid = attrOf(e, "spid");
    if (spid && /^-?\d+$/.test(spid)) return Number(spid);
  }
  const children = e[tag] as unknown[] | undefined;
  if (Array.isArray(children)) {
    for (const c of children) {
      const found = findSpTgtSpid(c);
      if (found !== null) return found;
    }
  }
  return null;
}
