import type { CommandHandler } from "@officeai/core";
import type {
  GroupShape,
  OpaqueShape,
  Picture,
  PptxPresentation,
  PptxSnapshot,
  Shape,
  Slide,
  TextShape,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, maxCNvPrId } from "./helpers.js";
import type { DuplicateSlidePayload } from "./payloads.js";

const PRES_RELS_PATH = "ppt/_rels/presentation.xml.rels";
const REL_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

export const duplicateSlideHandler: CommandHandler<DuplicateSlidePayload, PptxSnapshot> = {
  type: "pptx:duplicate-slide",
  apply(snapshot, payload, ctx) {
    const { slide: src, index } = findSlide(snapshot, payload.slideIndex);

    const partIndex = snapshot.root.idGen.nextSlidePartIndex;
    const slideId = snapshot.root.idGen.nextSlideId;
    const partPath = `ppt/slides/slide${partIndex}.xml`;
    const slideRelsPath = `ppt/slides/_rels/slide${partIndex}.xml.rels`;
    const srcRelsPath = `ppt/slides/_rels/${baseName(src.partPath)}.rels`;

    // Mint new rId on presentation.xml.rels.
    const presRels = snapshot.relationships.get(PRES_RELS_PATH);
    if (!presRels) {
      throw new Error("missing presentation rels");
    }
    const relId = nextRelId(presRels.entries.map((e) => e.id));

    // Re-mint cNvPrIds in the cloned shapes (start above the source's max).
    const baseId = maxCNvPrId(src.shapes);
    let counter = baseId;
    const newShapes = src.shapes.map((s) => deepCloneShape(s, () => ++counter, ctx.mintNodeId));

    const cloned: Slide = {
      id: ctx.mintNodeId(),
      partPath,
      slideId,
      relId,
      ...(src.layoutPartPath ? { layoutPartPath: src.layoutPartPath } : {}),
      shapes: newShapes,
      slideOpaqueTail: src.slideOpaqueTail.map(cloneOpaque),
      slideRootAttrs: { ...src.slideRootAttrs },
      cSldAttrs: { ...src.cSldAttrs },
      spTreeHead: src.spTreeHead.map(cloneOpaque),
      cSldHead: src.cSldHead.map(cloneOpaque),
    };

    const newSlides = [...snapshot.root.slides];
    newSlides.splice(index + 1, 0, cloned);

    // Copy slide rels file from source to new path.
    const srcSlideRels = snapshot.relationships.get(srcRelsPath);
    const newRelationships = new Map(snapshot.relationships);
    if (srcSlideRels) {
      newRelationships.set(slideRelsPath, {
        relsPath: slideRelsPath,
        entries: srcSlideRels.entries.map((e) => ({ ...e })),
      });
    }
    const newPresEntries = [
      ...presRels.entries,
      {
        id: relId,
        type: REL_TYPE_SLIDE,
        target: `slides/slide${partIndex}.xml`,
      },
    ];
    newRelationships.set(PRES_RELS_PATH, {
      relsPath: PRES_RELS_PATH,
      entries: newPresEntries,
    });

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

    const root: PptxPresentation = {
      ...snapshot.root,
      slides: newSlides,
      idGen: {
        ...snapshot.root.idGen,
        nextSlideId: slideId + 1,
        nextSlidePartIndex: partIndex + 1,
      },
    };

    const dirtyRels = [PRES_RELS_PATH];
    if (srcSlideRels) dirtyRels.push(slideRelsPath);

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        presentation: true,
        contentTypes: !overrideExists,
        slides: [partPath],
        relationships: dirtyRels,
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
        nodeId: cloned.id,
        path: ["slides", index + 1],
        summary: "slide (duplicate)",
      }),
    };
  },
};

function deepCloneShape(
  s: Shape,
  nextCNvPrId: () => number,
  mintNodeId: () => string
): Shape {
  const id = mintNodeId();
  const cNvPrId = nextCNvPrId();
  switch (s.kind) {
    case "text": {
      const c: TextShape = {
        ...s,
        id,
        cNvPrId,
        nvSpPrTail: s.nvSpPrTail.map(cloneOpaque),
        spPrTail: s.spPrTail.map(cloneOpaque),
        ...(s.styleRaw ? { styleRaw: cloneOpaque(s.styleRaw) } : {}),
        txBody: {
          ...s.txBody,
          paragraphs: s.txBody.paragraphs.map((p) => ({
            ...p,
            id: mintNodeId(),
            properties: { ...p.properties },
            runs: p.runs.map((r) => ({
              ...r,
              id: mintNodeId(),
              properties: { ...r.properties },
            })),
            ...(p.endParaRPrRaw ? { endParaRPrRaw: cloneOpaque(p.endParaRPrRaw) } : {}),
          })),
          ...(s.txBody.bodyPrRaw ? { bodyPrRaw: cloneOpaque(s.txBody.bodyPrRaw) } : {}),
          ...(s.txBody.lstStyleRaw ? { lstStyleRaw: cloneOpaque(s.txBody.lstStyleRaw) } : {}),
        },
      };
      return c;
    }
    case "pic": {
      const c: Picture = {
        ...s,
        id,
        cNvPrId,
        nvPicPrTail: s.nvPicPrTail.map(cloneOpaque),
        blipFillTail: s.blipFillTail.map(cloneOpaque),
        spPrTail: s.spPrTail.map(cloneOpaque),
        ...(s.styleRaw ? { styleRaw: cloneOpaque(s.styleRaw) } : {}),
      };
      return c;
    }
    case "group": {
      const c: GroupShape = {
        ...s,
        id,
        cNvPrId,
        nvGrpSpPrTail: s.nvGrpSpPrTail.map(cloneOpaque),
        grpSpPrTail: s.grpSpPrTail.map(cloneOpaque),
        chOffExtRaw: s.chOffExtRaw.map(cloneOpaque),
        children: s.children.map((c2) => deepCloneShape(c2, nextCNvPrId, mintNodeId)),
      };
      return c;
    }
    case "table": {
      const c: Shape = {
        ...s,
        id,
        cNvPrId,
        nvGraphicFramePrTail: s.nvGraphicFramePrTail.map(cloneOpaque),
        ...(s.tblPrRaw ? { tblPrRaw: cloneOpaque(s.tblPrRaw) } : {}),
        columnWidths: [...s.columnWidths],
        rows: s.rows.map((r) => ({
          ...r,
          id: mintNodeId(),
          trAttrs: { ...r.trAttrs },
          cells: r.cells.map((cell) => ({
            ...cell,
            id: mintNodeId(),
            tcAttrs: { ...cell.tcAttrs },
            ...(cell.tcPrRaw ? { tcPrRaw: cloneOpaque(cell.tcPrRaw) } : {}),
            txBody: {
              ...cell.txBody,
              ...(cell.txBody.bodyPrRaw
                ? { bodyPrRaw: cloneOpaque(cell.txBody.bodyPrRaw) }
                : {}),
              ...(cell.txBody.lstStyleRaw
                ? { lstStyleRaw: cloneOpaque(cell.txBody.lstStyleRaw) }
                : {}),
              paragraphs: cell.txBody.paragraphs.map((p) => ({
                ...p,
                id: mintNodeId(),
                properties: { ...p.properties },
                runs: p.runs.map((r2) => ({
                  ...r2,
                  id: mintNodeId(),
                  properties: { ...r2.properties },
                })),
                ...(p.endParaRPrRaw
                  ? { endParaRPrRaw: cloneOpaque(p.endParaRPrRaw) }
                  : {}),
              })),
            },
          })),
        })),
      };
      return c;
    }
    case "chart": {
      const c: Shape = {
        ...s,
        id,
        cNvPrId,
        nvGraphicFramePrTail: s.nvGraphicFramePrTail.map(cloneOpaque),
      };
      return c;
    }
    case "opaque": {
      const c: OpaqueShape = { ...s, id, cNvPrId, raw: cloneOpaque(s.raw) };
      return c;
    }
  }
}

function cloneOpaque<T extends { subtree: ReadonlyArray<unknown>; rawAttrs: Record<string, string> }>(
  o: T
): T {
  return { ...o, rawAttrs: { ...o.rawAttrs }, subtree: deepClone(o.subtree) as typeof o.subtree };
}

function deepClone<T>(v: T): T {
  if (v === null || typeof v !== "object") return v;
  if (Array.isArray(v)) return v.map(deepClone) as unknown as T;
  const out: Record<string, unknown> = {};
  for (const [k, val] of Object.entries(v as Record<string, unknown>)) {
    out[k] = deepClone(val);
  }
  return out as T;
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

function baseName(p: string): string {
  const slash = p.lastIndexOf("/");
  return slash >= 0 ? p.slice(slash + 1) : p;
}
