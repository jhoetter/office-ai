import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxSnapshot,
  InlineImageDrawing,
  Paragraph,
  Run,
  RunChild,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetImagePropertiesPayload } from "./payloads.js";

const EMU_PER_PIXEL = 9525;

/**
 * B6 — Update an inline image's display dimensions and accessibility
 * metadata.
 *
 * Targets the {@link InlineImageDrawing} whose `id` matches
 * `imageId`; the leaf is replaced in place inside its owning
 * paragraph's run, and the new leaf drops the `raw` cache so the
 * serializer regenerates the `<w:drawing>` subtree from the typed
 * model. Untouched drawings still round-trip byte-identical because
 * only the targeted leaf gets touched.
 *
 * `widthPx` / `heightPx` accept CSS pixels at 96 DPI for ergonomics
 * (the editor surface speaks pixels everywhere), and are converted to
 * EMU on the way in. Pass `null` for any field to leave it
 * unchanged.
 */
export const setImagePropertiesHandler: CommandHandler<SetImagePropertiesPayload, DocxSnapshot> = {
  type: "docx:set-image-properties",
  apply(snapshot, payload) {
    const { imageId, widthPx, heightPx, altText, name } = payload;
    if (!imageId) {
      throw new CommandError("invalid-payload", "imageId must be a non-empty string");
    }
    if (widthPx === undefined && heightPx === undefined && altText === undefined && name === undefined) {
      throw new CommandError("invalid-payload", "set-image-properties requires at least one field");
    }

    const located = findImage(snapshot, imageId);
    if (!located) {
      throw new CommandError("unknown-target", `no inline image with id "${imageId}"`);
    }

    const current = located.image;
    const nextCx =
      widthPx !== undefined && widthPx !== null
        ? Math.max(1, Math.round(widthPx * EMU_PER_PIXEL))
        : current.cx;
    const nextCy =
      heightPx !== undefined && heightPx !== null
        ? Math.max(1, Math.round(heightPx * EMU_PER_PIXEL))
        : current.cy;
    const nextDescr = altText === undefined ? current.descr : (altText ?? undefined);
    const nextName = name === undefined ? current.name : name;

    if (
      nextCx === current.cx &&
      nextCy === current.cy &&
      (nextDescr ?? "") === (current.descr ?? "") &&
      nextName === current.name
    ) {
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: imageId,
          path: ["body", located.paragraphIndex, "runs", located.runIndex, "children", located.childIndex],
          field: "imageProperties",
          summary: "no-op (nothing to update)",
        }),
      };
    }

    const updated: InlineImageDrawing = {
      ...current,
      cx: nextCx,
      cy: nextCy,
      descr: nextDescr,
      name: nextName,
      raw: undefined,
    };

    const newBody = replaceImage(snapshot.root.body, located, updated);
    const next = evolveSnapshot(snapshot, { ...snapshot.root, body: newBody }, { body: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: imageId,
        path: ["body", located.paragraphIndex, "runs", located.runIndex, "children", located.childIndex],
        field: "imageProperties",
        summary: `cx=${nextCx} cy=${nextCy}${nextDescr !== undefined ? ` descr="${nextDescr}"` : ""}`,
      }),
    };
  },
};

interface LocatedImage {
  readonly image: InlineImageDrawing;
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly childIndex: number;
}

function findImage(snapshot: DocxSnapshot, imageId: string): LocatedImage | null {
  const body = snapshot.root.body;
  for (let p = 0; p < body.length; p++) {
    const block = body[p];
    if (block.kind !== "paragraph") continue;
    const located = findImageInParagraph(block, imageId, p);
    if (located) return located;
  }
  return null;
}

function findImageInParagraph(
  paragraph: Paragraph,
  imageId: string,
  paragraphIndex: number
): LocatedImage | null {
  const children = paragraph.children;
  for (let r = 0; r < children.length; r++) {
    const node = children[r];
    if (node.kind !== "run") continue;
    const run: Run = node;
    for (let c = 0; c < run.children.length; c++) {
      const child: RunChild = run.children[c];
      if (child.kind === "drawing" && child.subkind === "inline-image" && child.id === imageId) {
        return { image: child, paragraphIndex, runIndex: r, childIndex: c };
      }
    }
  }
  return null;
}

function replaceImage(
  body: ReadonlyArray<BlockNode>,
  located: LocatedImage,
  next: InlineImageDrawing
): BlockNode[] {
  const out: BlockNode[] = body.slice();
  const block = body[located.paragraphIndex];
  if (block.kind !== "paragraph") return out;
  const newChildren = block.children.slice();
  const oldRunNode = newChildren[located.runIndex];
  if (oldRunNode.kind !== "run") return out;
  const newRunChildren = oldRunNode.children.slice();
  newRunChildren[located.childIndex] = next;
  const newRun: Run = { ...oldRunNode, children: newRunChildren };
  newChildren[located.runIndex] = newRun;
  const newParagraph: Paragraph = { ...block, children: newChildren };
  out[located.paragraphIndex] = newParagraph;
  return out;
}
