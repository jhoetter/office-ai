import { defaultIdMinter, ooxml, sha256Hex, type IdMinter } from "@officeai/core";
import { DEFAULT_THEME, type ThemeColorScheme } from "../renderer/layout/color.js";
import { parseThemeColorScheme } from "./theme.js";
import type {
  ContentTypesSnap,
  GroupShape,
  MediaPart,
  OpaquePart,
  OpaqueShape,
  OpaqueXml,
  Picture,
  PptxIdGen,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Shape,
  Slide,
  SlideSize,
  TableCell,
  TableRow,
  TableShape,
  TextBody,
  TextParagraph,
  TextParagraphProperties,
  TextRun,
  TextRunProperties,
  TextShape,
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
const REL_TYPE_NOTES_SLIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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
  const presentationChildren =
    (presentationEntry["p:presentation"] as unknown[] | undefined) ?? [];

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
  const layouts = new Map<string, OpaquePart>();
  for (const p of layoutPaths) {
    layouts.set(p, opaquePartFor(container, p, "p:sldLayout"));
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
  const notesSlides = new Map<string, OpaquePart>();
  for (const p of notesSlidePaths) {
    if (container.has(p)) notesSlides.set(p, opaquePartFor(container, p, "p:notes"));
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
    const graph = ooxml.RelationshipGraph.loadFor(
      container,
      relsPathToOwnerPath(partPath) ?? partPath
    );
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
    media,
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
  // Anything in <p:sld> after <p:cSld> we capture as opaque tail
  // (e.g. <p:clrMapOvr>, <p:transition>, <p:timing>).
  const slideOpaqueTail: OpaqueXml[] = [];
  let pastCsld = false;
  for (const c of elementEntries(sldChildren)) {
    if (!pastCsld) {
      if (ooxml.getTag(c) === "p:cSld") pastCsld = true;
      continue;
    }
    slideOpaqueTail.push(captureOpaque(c));
  }

  // Slide rels → layout, notes slide. Loaded BEFORE shape parsing so picture
  // rels can be resolved to absolute media part paths in one pass.
  const rels = ooxml.RelationshipGraph.loadFor(container, partPath);
  const slideRelTargets = new Map<string, string>();
  let layoutPartPath: string | undefined;
  let notesSlidePartPath: string | undefined;
  for (const r of rels.relationships) {
    slideRelTargets.set(r.id, r.target);
    if (r.type === REL_TYPE_SLIDE_LAYOUT) {
      layoutPartPath = resolveTarget(partPath, r.target);
    } else if (r.type === REL_TYPE_NOTES_SLIDE) {
      notesSlidePartPath = resolveTarget(partPath, r.target);
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

  return {
    id: mintNodeId(),
    partPath,
    slideId,
    relId,
    ...(layoutPartPath ? { layoutPartPath } : {}),
    ...(notesSlidePartPath ? { notesSlidePartPath } : {}),
    shapes,
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
    case "p:grpSp":
      return parseGrpSp(entry, mintNodeId, partPath, slideRelTargets);
    case "p:graphicFrame": {
      const typed = parseGraphicFrameTable(entry, mintNodeId);
      if (typed) return typed;
      return parseOpaqueShape(entry, mintNodeId);
    }
    default:
      return parseOpaqueShape(entry, mintNodeId);
  }
}

const TABLE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

/**
 * Parse `<p:graphicFrame>` ⇒ `TableShape` if and only if its
 * `<a:graphicData @uri>` is the table URI. Returns `null` otherwise so
 * the caller can fall back to `OpaqueShape` for charts / SmartArt.
 */
function parseGraphicFrameTable(
  entry: Record<string, unknown>,
  mintNodeId: IdMinter
): TableShape | null {
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
  const tbl = findElementEntry(
    (graphicData["a:graphicData"] as unknown[] | undefined) ?? [],
    "a:tbl"
  );
  if (!tbl) return null;

  let cNvPrId = 0;
  let name = "";
  const nvGFPrTail: OpaqueXml[] = [];
  if (nvGFPr) {
    for (const c of elementEntries(
      (nvGFPr["p:nvGraphicFramePr"] as unknown[] | undefined) ?? []
    )) {
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
    for (const c of elementEntries(
      (tblGrid["a:tblGrid"] as unknown[] | undefined) ?? []
    )) {
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
        ? parseTextBodyChildren(
            (txBodyEntry["a:txBody"] as unknown[] | undefined) ?? [],
            mintNodeId
          )
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
function parseTextBodyChildren(
  children: ReadonlyArray<unknown>,
  mintNodeId: IdMinter
): TextBody {
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

function collectPartsByType(
  graph: ooxml.RelationshipGraph,
  type: string
): string[] {
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

function opaquePartFor(
  container: ooxml.OoxmlContainer,
  partPath: string,
  rootTag: string
): OpaquePart {
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
  REL_TYPE_IMAGE,
  REL_TYPE_NOTES_SLIDE,
  REL_TYPE_SLIDE,
  REL_TYPE_SLIDE_LAYOUT,
  REL_TYPE_SLIDE_MASTER,
  REL_TYPE_THEME,
};
