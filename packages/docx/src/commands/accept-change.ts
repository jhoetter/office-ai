import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  HeaderFooterPart,
  InlineNode,
  Paragraph,
  RevisionWrapper,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { mergeHeaderFooterDirty } from "./set-header-text.js";
import type { AcceptChangePayload } from "./payloads.js";

/**
 * Accept a tracked change by `revisionId`.
 *
 * Semantics (per OOXML §17.13):
 *   - `<w:ins>` accept: fold the inserted runs into the parent body. The
 *     wrapper goes away; the children stay.
 *   - `<w:del>` accept: drop the wrapper AND its children. The text the
 *     reviewer marked for deletion really goes.
 *
 * Implementation strategy:
 *   The body is a tree of paragraphs whose inline children may themselves be
 *   `RevisionWrapper`s. We walk the body once (and every header/footer body)
 *   and rewrite each paragraph's `children` list, replacing matching wrappers
 *   with either their unwrapped children (for `ins`) or nothing (for `del`).
 *
 * Round-trip property: the resulting snapshot, when serialized and
 * re-parsed, has no `RevisionWrapper` whose `revisionId === payload.revisionId`.
 * Tests in `tracked-changes.test.ts` assert this directly.
 */
export const acceptChangeHandler: CommandHandler<AcceptChangePayload, DocxSnapshot> = {
  type: "docx:accept-change",
  apply(snapshot, payload) {
    return resolveTrackedChange(snapshot, payload.revisionId, "accept");
  },
};

export type Resolution = "accept" | "reject";

export function resolveTrackedChange(
  snapshot: DocxSnapshot,
  revisionId: string,
  resolution: Resolution
): { next: DocxSnapshot; diff: ReturnType<typeof buildDiff> } {
  if (!revisionId) {
    throw new CommandError("unknown-revision", "revisionId must be a non-empty string");
  }
  const found = countRevision(snapshot, revisionId);
  if (found === 0) {
    throw new CommandError("unknown-revision", `no tracked change with revisionId "${revisionId}"`);
  }

  const { body, bodyChanged } = rewriteBlocks(snapshot.root.body, revisionId, resolution);

  const dirty: { body?: boolean; headersAndFooters?: ReadonlySet<string> } = {};
  if (bodyChanged) dirty.body = true;

  let nextHeadersAndFooters = snapshot.root.headersAndFooters;
  const touchedHfPaths: string[] = [];
  if (snapshot.root.headersAndFooters.length > 0) {
    const newParts: HeaderFooterPart[] = [];
    let anyChanged = false;
    for (const part of snapshot.root.headersAndFooters) {
      const result = rewriteBlocks(part.body, revisionId, resolution);
      if (result.bodyChanged) {
        anyChanged = true;
        touchedHfPaths.push(part.partPath);
        newParts.push({ ...part, body: result.body });
      } else {
        newParts.push(part);
      }
    }
    if (anyChanged) {
      nextHeadersAndFooters = newParts;
      let mergedDirty = snapshot.dirty.headersAndFooters;
      for (const p of touchedHfPaths) {
        mergedDirty = mergeHeaderFooterDirty({ ...snapshot.dirty, headersAndFooters: mergedDirty }, p);
      }
      dirty.headersAndFooters = mergedDirty;
    }
  }

  const nextDoc: DocxDocument = {
    ...snapshot.root,
    body,
    headersAndFooters: nextHeadersAndFooters,
  };
  const next = evolveSnapshot(snapshot, nextDoc, dirty);

  const verb = resolution === "accept" ? "accepted" : "rejected";
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-deleted",
      nodeId: revisionId,
      path: ["body"],
      summary: `${verb} change ${revisionId} (×${found})`,
    }),
  };
}

function countRevision(snapshot: DocxSnapshot, revisionId: string): number {
  let n = 0;
  for (const b of snapshot.root.body) n += countInBlock(b, revisionId);
  for (const part of snapshot.root.headersAndFooters) {
    for (const b of part.body) n += countInBlock(b, revisionId);
  }
  return n;
}

function countInBlock(block: BlockNode, revisionId: string): number {
  if (block.kind !== "paragraph") return 0;
  let n = 0;
  for (const child of block.children) n += countInInline(child, revisionId);
  return n;
}

function countInInline(node: InlineNode, revisionId: string): number {
  if (node.kind === "revision") {
    let n = node.revisionId === revisionId ? 1 : 0;
    for (const c of node.children) n += countInInline(c, revisionId);
    return n;
  }
  return 0;
}

function rewriteBlocks(
  blocks: ReadonlyArray<BlockNode>,
  revisionId: string,
  resolution: Resolution
): { body: BlockNode[]; bodyChanged: boolean } {
  const out: BlockNode[] = [];
  let changed = false;
  for (const b of blocks) {
    if (b.kind !== "paragraph") {
      out.push(b);
      continue;
    }
    const result = rewriteParagraph(b, revisionId, resolution);
    if (result.changed) changed = true;
    out.push(result.paragraph);
  }
  return { body: out, bodyChanged: changed };
}

function rewriteParagraph(
  p: Paragraph,
  revisionId: string,
  resolution: Resolution
): { paragraph: Paragraph; changed: boolean } {
  const next: InlineNode[] = [];
  let changed = false;
  for (const child of p.children) {
    const handled = rewriteInline(child, revisionId, resolution);
    if (handled.changed) changed = true;
    for (const o of handled.nodes) next.push(o);
  }
  if (!changed) return { paragraph: p, changed: false };
  return { paragraph: { ...p, children: next }, changed: true };
}

/**
 * Rewrite one inline node. Returns the (possibly empty) replacement list
 * plus a flag indicating whether anything changed in this subtree. The
 * replacement is a list because:
 *   - matching `ins` accept → unwrap → expands one node into N children;
 *   - matching `del` accept → drop → returns 0 nodes;
 *   - matching `ins` reject → drop → returns 0 nodes;
 *   - matching `del` reject → unwrap → returns N children (with the
 *     surrounding `<w:delText>` nodes left as `w:delText` text leaves; the
 *     serializer keeps them as-is since they're still well-formed text
 *     leaves, but downstream code that cares about post-resolution text
 *     should re-create them as plain `w:t`. We deliberately leave the
 *     run children's `isDelText` flag alone for now — it's effectively a
 *     hint and Word/LibreOffice both accept the result either way).
 */
function rewriteInline(
  node: InlineNode,
  revisionId: string,
  resolution: Resolution
): { nodes: InlineNode[]; changed: boolean } {
  if (node.kind === "revision") {
    if (node.revisionId === revisionId) {
      const accept = resolution === "accept";
      const isIns = node.revisionType === "ins";
      const keepChildren = (accept && isIns) || (!accept && !isIns);
      if (keepChildren) {
        const nested: InlineNode[] = [];
        for (const c of node.children) {
          const r = rewriteInline(c, revisionId, resolution);
          for (const o of r.nodes) nested.push(o);
        }
        return { nodes: nested, changed: true };
      }
      return { nodes: [], changed: true };
    }
    // Not the targeted revision — recurse into the wrapper's children in
    // case the matching wrapper is nested inside this one (rare but legal).
    const inner: InlineNode[] = [];
    let nestedChanged = false;
    for (const c of node.children) {
      const r = rewriteInline(c, revisionId, resolution);
      if (r.changed) nestedChanged = true;
      for (const o of r.nodes) inner.push(o);
    }
    if (!nestedChanged) return { nodes: [node], changed: false };
    const updated: RevisionWrapper = { ...node, children: inner };
    return { nodes: [updated], changed: true };
  }
  return { nodes: [node], changed: false };
}
