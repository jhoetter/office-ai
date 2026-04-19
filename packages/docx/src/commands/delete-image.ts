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
import type { DeleteImagePayload } from "./payloads.js";

/**
 * B6 — Delete an inline image leaf by id.
 *
 * Removes the matching {@link InlineImageDrawing} from its owning
 * run's children. If that leaves the run completely empty *and* the
 * run was the only inline node in the paragraph, we leave behind an
 * empty paragraph (callers that want the paragraph removed too can
 * follow up with a paragraph-level delete) — this mirrors what Word
 * does when you select-and-delete a picture: the surrounding
 * formatting and paragraph structure stay intact.
 *
 * The owning run keeps its `properties` so any preserved text
 * formatting (font, size, colour) survives intact for the next
 * keystroke. Other runs in the paragraph are left untouched, so
 * surrounding text never re-flows in surprising ways.
 *
 * Media + relationships are intentionally *not* garbage-collected
 * here; an undo would have to rebuild them, and OOXML round-trip
 * tolerates orphan media just fine. A separate "Compact media"
 * pass can prune later.
 */
export const deleteImageHandler: CommandHandler<DeleteImagePayload, DocxSnapshot> = {
  type: "docx:delete-image",
  apply(snapshot, payload) {
    const { imageId } = payload;
    if (!imageId) {
      throw new CommandError("invalid-payload", "imageId must be a non-empty string");
    }

    const located = findImage(snapshot, imageId);
    if (!located) {
      throw new CommandError("unknown-target", `no inline image with id "${imageId}"`);
    }

    const newBody = removeImage(snapshot.root.body, located);
    const next = evolveSnapshot(snapshot, { ...snapshot.root, body: newBody }, { body: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: imageId,
        path: ["body", located.paragraphIndex, "runs", located.runIndex, "children", located.childIndex],
        summary: `−image (${located.image.cx}×${located.image.cy} EMU)`,
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

function removeImage(body: ReadonlyArray<BlockNode>, located: LocatedImage): BlockNode[] {
  const out: BlockNode[] = body.slice();
  const block = body[located.paragraphIndex];
  if (block.kind !== "paragraph") return out;
  const newChildren = block.children.slice();
  const oldRunNode = newChildren[located.runIndex];
  if (oldRunNode.kind !== "run") return out;
  const newRunChildren = oldRunNode.children.slice();
  newRunChildren.splice(located.childIndex, 1);
  const newRun: Run = { ...oldRunNode, children: newRunChildren };
  newChildren[located.runIndex] = newRun;
  const newParagraph: Paragraph = { ...block, children: newChildren };
  out[located.paragraphIndex] = newParagraph;
  return out;
}
