import type { CommandHandler } from "@officeai/core";
import type { BlockNode, DocxSnapshot, InlineNode } from "../model/types.js";
import { resolveTrackedChange } from "./accept-change.js";
import { buildDiff } from "./helpers.js";
import type { AcceptAllChangesPayload } from "./payloads.js";

/**
 * B8 — Accept every tracked change in the document in a single
 * transaction. We collect the unique `revisionId`s present in the
 * body + every header/footer body, then fold them one at a time
 * through {@link resolveTrackedChange}. Folding sequentially keeps
 * the existing tree-rewriting code as the single source of truth
 * for what "accept" means (insertions stay, deletions go), while the
 * outer loop guarantees an "accept all" button can never leave the
 * snapshot half-resolved.
 *
 * The diff is summarised as a single `node-deleted` change against
 * the body — a faithful per-revision diff would require N entries
 * for N changes, which the agent log already captures via the
 * intermediate `applyCommand` calls. The summary string carries the
 * count so the toast row can read "Accepted N changes".
 */
export const acceptAllChangesHandler: CommandHandler<AcceptAllChangesPayload, DocxSnapshot> = {
  type: "docx:accept-all-changes",
  apply(snapshot) {
    return resolveAll(snapshot, "accept");
  },
};

export function resolveAll(
  snapshot: DocxSnapshot,
  resolution: "accept" | "reject"
): { next: DocxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const ids = collectRevisionIds(snapshot);
  if (ids.length === 0) {
    return {
      next: snapshot,
      diff: buildDiff(snapshot.revision, snapshot.revision, {
        kind: "node-updated",
        nodeId: "tracked-changes",
        path: ["body"],
        field: "trackedChanges",
        summary: "no tracked changes to resolve",
      }),
    };
  }
  let current = snapshot;
  for (const id of ids) {
    const result = resolveTrackedChange(current, id, resolution);
    current = result.next;
  }
  const verb = resolution === "accept" ? "Accepted" : "Rejected";
  return {
    next: current,
    diff: buildDiff(snapshot.revision, current.revision, {
      kind: "node-updated",
      nodeId: "tracked-changes",
      path: ["body"],
      field: "trackedChanges",
      summary: `${verb} ${ids.length} change${ids.length === 1 ? "" : "s"}`,
    }),
  };
}

function collectRevisionIds(snapshot: DocxSnapshot): string[] {
  const seen = new Set<string>();
  const order: string[] = [];
  for (const block of snapshot.root.body) walkBlock(block, seen, order);
  for (const part of snapshot.root.headersAndFooters) {
    for (const block of part.body) walkBlock(block, seen, order);
  }
  return order;
}

function walkBlock(block: BlockNode, seen: Set<string>, order: string[]): void {
  if (block.kind !== "paragraph") return;
  for (const child of block.children) walkInline(child, seen, order);
}

function walkInline(node: InlineNode, seen: Set<string>, order: string[]): void {
  if (node.kind !== "revision") return;
  if (!seen.has(node.revisionId)) {
    seen.add(node.revisionId);
    order.push(node.revisionId);
  }
  for (const c of node.children) walkInline(c, seen, order);
}
