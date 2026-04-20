import type { CommandHandler } from "@officeai/core";
import type { PdfDocument, PdfOutlineNode, PdfSnapshot } from "../model/types.js";
import { buildDiff, evolvePdf, makeError, validatePages } from "./helpers.js";
import type { AddBookmarkPayload } from "./payloads.js";

const insertChild = (
  outline: ReadonlyArray<PdfOutlineNode>,
  parentId: string,
  child: PdfOutlineNode
): { next: PdfOutlineNode[]; inserted: boolean } => {
  const out: PdfOutlineNode[] = [];
  let inserted = false;
  for (const node of outline) {
    if (node.id === parentId) {
      out.push({ ...node, children: [...node.children, child] });
      inserted = true;
      continue;
    }
    if (node.children.length > 0) {
      const sub = insertChild(node.children, parentId, child);
      if (sub.inserted) {
        inserted = true;
        out.push({ ...node, children: sub.next });
        continue;
      }
    }
    out.push(node);
  }
  return { next: out, inserted };
};

export const addBookmarkHandler: CommandHandler<AddBookmarkPayload, PdfSnapshot> = {
  type: "pdf:add-bookmark",
  apply(snapshot, payload, ctx) {
    validatePages(snapshot, [payload.pageNumber], "pdf:add-bookmark");
    const node: PdfOutlineNode = {
      id: ctx.mintNodeId(),
      title: payload.title,
      pageNumber: payload.pageNumber,
      children: [],
    };
    let outline: ReadonlyArray<PdfOutlineNode>;
    let path: ReadonlyArray<string | number>;
    if (payload.parentId) {
      const result = insertChild(snapshot.root.outline, payload.parentId, node);
      if (!result.inserted) {
        throw makeError("unknown-target", `bookmark parent ${payload.parentId} not found`);
      }
      outline = result.next;
      path = ["outline", payload.parentId, "children", node.id];
    } else {
      outline = [...snapshot.root.outline, node];
      path = ["outline", outline.length - 1];
    }
    const root: PdfDocument = { ...snapshot.root, outline };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: node.id,
        path,
        summary: "bookmark",
      }),
    };
  },
};
