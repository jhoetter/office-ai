import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  DocxDirtyFlags,
  DocxSnapshot,
  HeaderFooterPart,
  Paragraph,
  Run,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, textLeaf } from "./helpers.js";
import type { SetHeaderTextPayload } from "./payloads.js";

/**
 * Replace the text content of one paragraph inside a header part. Idempotent
 * (re-running with the same text bumps the revision but produces an
 * equivalent snapshot). Header / footer parts share their entire shape
 * (paragraphs of runs); the only thing this handler keys on is `kind:
 * "header"`. The footer counterpart in `set-footer-text.ts` keys on
 * `kind: "footer"` but uses the same internal builder via
 * `applySetTextToHeaderFooter`.
 */
export const setHeaderTextHandler: CommandHandler<SetHeaderTextPayload, DocxSnapshot> = {
  type: "docx:set-header-text",
  apply(snapshot, payload, ctx) {
    return applySetTextToHeaderFooter(snapshot, "header", payload, ctx.mintNodeId);
  },
};

/**
 * Shared implementation for `set-header-text` and `set-footer-text`. Lives
 * here (rather than in `helpers.ts`) because helpers.ts is owned by an
 * earlier workstream; W4 keeps its mutations local to its own files.
 */
export function applySetTextToHeaderFooter(
  snapshot: DocxSnapshot,
  kind: "header" | "footer",
  payload: { partId: string; paragraphIndex: number; text: string },
  mintNodeId: IdMinter
): { next: DocxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const { partId, paragraphIndex, text } = payload;
  const partIdx = snapshot.root.headersAndFooters.findIndex((p) => p.id === partId && p.kind === kind);
  if (partIdx < 0) {
    throw new CommandError(
      "unknown-target",
      `no ${kind} part with id "${partId}" (looked across ${snapshot.root.headersAndFooters.length} header/footer parts)`
    );
  }
  const part = snapshot.root.headersAndFooters[partIdx];
  if (paragraphIndex < 0 || paragraphIndex >= part.body.length) {
    throw new CommandError(
      "unknown-target",
      `paragraph index ${paragraphIndex} out of range for ${kind} part "${partId}" (has ${part.body.length} blocks)`
    );
  }
  const block = part.body[paragraphIndex];
  if (block.kind !== "paragraph") {
    throw new CommandError(
      "unknown-target",
      `block at index ${paragraphIndex} in ${kind} part "${partId}" is not a paragraph (kind=${block.kind})`
    );
  }

  const replacedParagraph = replaceParagraphText(block, text, mintNodeId);
  const newBody: BlockNode[] = part.body.slice();
  newBody[paragraphIndex] = replacedParagraph;
  const updatedPart: HeaderFooterPart = { ...part, body: newBody };
  const newParts = snapshot.root.headersAndFooters.slice();
  newParts[partIdx] = updatedPart;

  const dirty = mergeHeaderFooterDirty(snapshot.dirty, part.partPath);
  const next = evolveSnapshot(
    snapshot,
    { ...snapshot.root, headersAndFooters: newParts },
    { headersAndFooters: dirty }
  );

  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-updated",
      nodeId: block.id,
      path: ["headersAndFooters", partIdx, "body", paragraphIndex],
      field: "text",
      summary: `${kind} ${partId} p${paragraphIndex} := ${JSON.stringify(text)}`,
    }),
  };
}

/**
 * Build a paragraph that consists of a single run containing the given
 * literal text. Preserves the original paragraph's `id`, `properties`, and
 * (best-effort) the first run's `properties` so heading styles / italics on
 * the existing header content survive a text rewrite.
 */
function replaceParagraphText(p: Paragraph, text: string, mintNodeId: IdMinter): Paragraph {
  const firstRun = p.children.find((c): c is Run => c.kind === "run");
  const runProps = firstRun?.properties ?? {};
  const newRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: runProps,
    children: text === "" ? [] : [textLeaf(mintNodeId, text)],
  };
  return { ...p, children: [newRun] };
}

/**
 * Add `partPath` to the existing dirty set without mutating the prior set.
 * The set is treated as immutable across snapshots so older snapshots in
 * mutation history keep their own dirty view.
 */
export function mergeHeaderFooterDirty(prev: DocxDirtyFlags, partPath: string): ReadonlySet<string> {
  const next = new Set<string>(prev.headersAndFooters);
  next.add(partPath);
  return next;
}
