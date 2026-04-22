import type { CommandHandler, NodeId } from "@officeai/core";
import type { PptxSnapshot, Shape, TextParagraph, TextRun } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError, maxCNvPrId, withSlide } from "./helpers.js";
import type { PasteShapesPayload } from "./payloads.js";

const DEFAULT_OFFSET_EMU = 228_600; // ¼ inch — matches PowerPoint's paste nudge.

/**
 * Re-stamp a list of shapes onto a slide. The companion to
 * `pptx:duplicate-shape`, but accepting an externally-supplied shape
 * array (typically deserialised from the clipboard) rather than a
 * shape id resolved against the current snapshot.
 *
 * Phase-1 scope (matches `spec/shared/clipboard.md` §"PPTX shapes"):
 *   • Supported kinds: text, table (inline data), connector, group
 *     (when every descendant is a supported kind), and the simple
 *     prst geometry shapes (which present as `text` shapes with an
 *     empty/short txBody — they're modelled the same way).
 *   • Refused kinds: pic, chart, ole-spreadsheet, media, opaque.
 *     These reference container-side parts (media bytes, embedded
 *     workbooks, etc.) that paste cannot recreate without copying
 *     those parts; same-deck paste of these shapes should go through
 *     `pptx:duplicate-shape` instead, which keeps the existing rels.
 *
 * Connector endpoints are passed through verbatim. Anchored
 * connectors keep referring to their target shape's `cNvPrId`, which
 * means cross-deck pastes will silently render as orphan elbows
 * pointing at coordinates with no matching shape — that's the same
 * behaviour PowerPoint shows for the same scenario, and is preferred
 * over rewriting them to free endpoints (which would surprise users
 * who paste a connector + its targets together).
 */
export const pasteShapesHandler: CommandHandler<PasteShapesPayload, PptxSnapshot> = {
  type: "pptx:paste-shapes",
  apply(snapshot, payload, ctx) {
    if (payload.shapes.length === 0) {
      throw makeError("not-applicable", "no shapes to paste");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);

    for (const s of payload.shapes) {
      assertSupported(s);
    }

    const dx = Number.isFinite(payload.dxEmu) ? Math.round(payload.dxEmu!) : DEFAULT_OFFSET_EMU;
    const dy = Number.isFinite(payload.dyEmu) ? Math.round(payload.dyEmu!) : DEFAULT_OFFSET_EMU;

    let nextCNvPrId = maxCNvPrId(slide.shapes) + 1;
    const clones: Shape[] = [];
    for (const src of payload.shapes) {
      const cnv = nextCNvPrId++;
      clones.push(remintShape(src, dx, dy, cnv, ctx));
    }

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: [...s.shapes, ...clones],
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: clones[0].id,
        path: ["slides", sIdx, "shapes", slide.shapes.length],
        summary: `paste ${clones.length} shape${clones.length === 1 ? "" : "s"}`,
      }),
    };
  },
};

interface MintCtx {
  readonly mintNodeId: () => NodeId;
}

function assertSupported(shape: Shape): void {
  switch (shape.kind) {
    case "text":
    case "table":
    case "connector":
      return;
    case "group":
      for (const child of shape.children) assertSupported(child);
      return;
    case "pic":
    case "chart":
    case "ole-spreadsheet":
    case "media":
      throw makeError(
        "not-applicable",
        `paste of ${shape.kind} shapes across decks is not supported yet ` +
          `(reference parts cannot be copied through the clipboard); ` +
          `use Cmd+D to duplicate within the same deck instead`
      );
    case "opaque":
      throw makeError("not-applicable", "cannot paste opaque shapes");
  }
}

function remintShape(shape: Shape, dx: number, dy: number, newCNvPrId: number, ctx: MintCtx): Shape {
  const next: Shape = { ...shape, id: ctx.mintNodeId(), cNvPrId: newCNvPrId };
  const offset = nudgePosition(shape, dx, dy);
  switch (next.kind) {
    case "text":
      return { ...next, ...offset, txBody: cloneTextBody(next.txBody, ctx) } as Shape;
    case "table":
      return { ...next, ...offset } as Shape;
    case "connector": {
      const start =
        next.start.kind === "free"
          ? { ...next.start, xEmu: next.start.xEmu + dx, yEmu: next.start.yEmu + dy }
          : next.start;
      const end =
        next.end.kind === "free"
          ? { ...next.end, xEmu: next.end.xEmu + dx, yEmu: next.end.yEmu + dy }
          : next.end;
      return { ...next, ...offset, start, end } as Shape;
    }
    case "group":
      return {
        ...next,
        ...offset,
        children: next.children.map((c) => remintShape(c, 0, 0, c.cNvPrId, ctx)),
      } as Shape;
    case "pic":
    case "chart":
    case "ole-spreadsheet":
    case "media":
    case "opaque":
      // Refused above; reachable only if assertSupported regresses.
      throw makeError("not-applicable", `cannot paste ${next.kind}`);
  }
}

function nudgePosition(shape: Shape, dx: number, dy: number): Pick<Shape, "position"> {
  if (!shape.position) return {};
  return { position: { xEmu: shape.position.xEmu + dx, yEmu: shape.position.yEmu + dy } };
}

function cloneTextBody(
  body: import("../model/types.js").TextBody,
  ctx: MintCtx
): import("../model/types.js").TextBody {
  const paragraphs: TextParagraph[] = body.paragraphs.map((p) => ({
    ...p,
    id: ctx.mintNodeId(),
    runs: p.runs.map(
      (r): TextRun => ({
        ...r,
        id: ctx.mintNodeId(),
      })
    ),
  }));
  return { ...body, paragraphs };
}
