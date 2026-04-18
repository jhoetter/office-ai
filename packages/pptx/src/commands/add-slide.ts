import type { CommandHandler } from "@officeai/core";
import type { OpaquePart, PptxPresentation, PptxSnapshot, Slide } from "../model/types.js";
import { buildDiff, evolveSnapshot, makeError } from "./helpers.js";
import { resolveTarget } from "../parser/parse.js";
import type { AddSlidePayload } from "./payloads.js";

const PRES_RELS_PATH = "ppt/_rels/presentation.xml.rels";
const SLIDE_RELS_PREFIX = "ppt/slides/_rels/";
const REL_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_TYPE_LAYOUT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

export const addSlideHandler: CommandHandler<AddSlidePayload, PptxSnapshot> = {
  type: "pptx:add-slide",
  apply(snapshot, payload) {
    const slides = snapshot.root.slides;
    const at = payload.at ?? slides.length;
    if (at < 0 || at > slides.length) {
      throw makeError("invalid-position", `at ${at} out of range (0..${slides.length})`);
    }

    let layout: OpaquePart | undefined;
    if (payload.layoutPartPath) {
      layout = snapshot.root.layouts.get(payload.layoutPartPath);
      if (!layout) {
        throw makeError(
          "unknown-target",
          `layout part ${payload.layoutPartPath} not found in presentation.layouts`
        );
      }
    }

    const slideId = snapshot.root.idGen.nextSlideId;
    const partIndex = snapshot.root.idGen.nextSlidePartIndex;
    const partPath = `ppt/slides/slide${partIndex}.xml`;
    const slideRelsPath = `${SLIDE_RELS_PREFIX}slide${partIndex}.xml.rels`;

    const presRels = snapshot.relationships.get(PRES_RELS_PATH);
    if (!presRels) {
      throw makeError("missing-rels", `presentation rels not found: ${PRES_RELS_PATH}`);
    }
    const relId = nextRelId(presRels.entries.map((e) => e.id));

    const slide: Slide = {
      id: `slide-new-${partIndex}`,
      partPath,
      slideId,
      relId,
      ...(payload.layoutPartPath ? { layoutPartPath: payload.layoutPartPath } : {}),
      shapes: [],
      animations: [],
      slideOpaqueTail: [defaultClrMapOvr()],
      slideRootAttrs: {
        "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      },
      cSldAttrs: {},
      spTreeHead: [defaultNvGrpSpPr(), defaultGrpSpPr()],
      cSldHead: [],
    };

    // Update relationships snapshot for presentation.xml.rels.
    const newPresEntries = [
      ...presRels.entries,
      {
        id: relId,
        type: REL_TYPE_SLIDE,
        target: relativeFrom(PRES_RELS_PATH, partPath),
      },
    ];
    const newRelationships = new Map(snapshot.relationships);
    newRelationships.set(PRES_RELS_PATH, { relsPath: PRES_RELS_PATH, entries: newPresEntries });

    // Slide rels (only if layout specified).
    const dirtyRelsPaths: string[] = [PRES_RELS_PATH];
    if (payload.layoutPartPath) {
      const slideRelsEntries = [
        {
          id: "rId1",
          type: REL_TYPE_LAYOUT,
          target: relativeFrom(slideRelsPath, payload.layoutPartPath),
        },
      ];
      newRelationships.set(slideRelsPath, { relsPath: slideRelsPath, entries: slideRelsEntries });
      dirtyRelsPaths.push(slideRelsPath);
    }

    // Content types: add Override for the new slide.
    const overrideExists = snapshot.contentTypes.overrides.some(
      (o) => o.partName === `/${partPath}`
    );
    const newContentTypes = overrideExists
      ? snapshot.contentTypes
      : {
          ...snapshot.contentTypes,
          overrides: [
            ...snapshot.contentTypes.overrides,
            { partName: `/${partPath}`, contentType: SLIDE_CONTENT_TYPE },
          ],
        };

    // Insert slide and bump idGen.
    const newSlides = [...slides];
    newSlides.splice(at, 0, slide);

    const root: PptxPresentation = {
      ...snapshot.root,
      slides: newSlides,
      idGen: {
        ...snapshot.root.idGen,
        nextSlideId: slideId + 1,
        nextSlidePartIndex: partIndex + 1,
      },
    };

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        presentation: true,
        contentTypes: !overrideExists,
        slides: [partPath],
        relationships: dirtyRelsPaths,
      },
      {
        relationships: newRelationships,
        contentTypes: newContentTypes,
      }
    );

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: slide.id,
        path: ["slides", at],
        summary: "slide",
      }),
    };
  },
};

function defaultClrMapOvr(): import("../model/types.js").OpaqueXml {
  return {
    tag: "p:clrMapOvr",
    attrs: {},
    rawAttrs: {},
    subtree: [{ "a:masterClrMapping": [] }],
  };
}

function nextRelId(existing: ReadonlyArray<string>): string {
  let max = 0;
  for (const id of existing) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `rId${max + 1}`;
}

/**
 * Compute an OPC `Target` (relative path) from the rels file path to the
 * absolute target part path. Mirrors the inverse of resolveTarget.
 */
function relativeFrom(relsPath: string, targetPath: string): string {
  const ownerDir = relsPathOwnerDir(relsPath);
  const targetSegments = targetPath.split("/");
  const ownerSegments = ownerDir.split("/").filter((s) => s.length > 0);

  // Find common prefix.
  let i = 0;
  while (
    i < ownerSegments.length &&
    i < targetSegments.length - 1 &&
    ownerSegments[i] === targetSegments[i]
  ) {
    i++;
  }
  const ups = ownerSegments.length - i;
  const downs = targetSegments.slice(i);
  return [...Array(ups).fill(".."), ...downs].join("/");
}

function relsPathOwnerDir(relsPath: string): string {
  // ppt/_rels/presentation.xml.rels  →  ppt
  // ppt/slides/_rels/slide1.xml.rels →  ppt/slides
  const m = /^(.*?)_rels\/[^/]+\.rels$/.exec(relsPath);
  if (!m) return "";
  return (m[1] ?? "").replace(/\/$/, "");
}

function defaultNvGrpSpPr(): import("../model/types.js").OpaqueXml {
  return {
    tag: "p:nvGrpSpPr",
    attrs: {},
    rawAttrs: {},
    subtree: [
      { "p:cNvPr": [], ":@": { "@_id": "1", "@_name": "" } },
      { "p:cNvGrpSpPr": [] },
      { "p:nvPr": [] },
    ],
  };
}

function defaultGrpSpPr(): import("../model/types.js").OpaqueXml {
  return {
    tag: "p:grpSpPr",
    attrs: {},
    rawAttrs: {},
    subtree: [
      {
        "a:xfrm": [
          { "a:off": [], ":@": { "@_x": "0", "@_y": "0" } },
          { "a:ext": [], ":@": { "@_cx": "0", "@_cy": "0" } },
          { "a:chOff": [], ":@": { "@_x": "0", "@_y": "0" } },
          { "a:chExt": [], ":@": { "@_cx": "0", "@_cy": "0" } },
        ],
      },
    ],
  };
}

// Re-export for completeness; tests import these.
export { resolveTarget };
