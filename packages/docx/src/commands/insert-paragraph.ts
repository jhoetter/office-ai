import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  DocxSnapshot,
  Hyperlink,
  InlineNode,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, insertBlock, paragraphPlainText, replaceBlock } from "./helpers.js";
import type { InsertParagraphPayload } from "./payloads.js";

/**
 * Insert (or split into) a paragraph at `at`.
 *
 * Semantics match what an editor's Enter key produces:
 *
 *   - If `at.offset` is 0 or undefined → insert an empty paragraph
 *     **before** `at.paragraph` (cursor sits at the original paragraph).
 *   - If `at.offset >= text length` → insert an empty paragraph
 *     **after** `at.paragraph`.
 *   - Otherwise → split `at.paragraph` at the offset; the original
 *     paragraph keeps the text before the offset, a new paragraph
 *     receives the text after.
 *
 * Style of the new paragraph:
 *   - explicit `style` payload wins,
 *   - else inherit the source paragraph's `styleId` only when splitting
 *     in the middle (so an empty leading/trailing paragraph is plain).
 */
export const insertParagraphHandler: CommandHandler<InsertParagraphPayload, DocxSnapshot> = {
  type: "docx:insert-paragraph",
  apply(snapshot, payload, ctx) {
    const { at, style } = payload;
    if (at.paragraph < 0 || at.paragraph >= snapshot.root.body.length) {
      throw new CommandError("invalid-position", `paragraph index ${at.paragraph} out of range`);
    }
    const sourceBlock = snapshot.root.body[at.paragraph];
    if (sourceBlock.kind !== "paragraph") {
      throw new CommandError("not-paragraph", `block at ${at.paragraph} is not a paragraph`);
    }
    const sourceText = paragraphPlainText(sourceBlock);
    const offset = at.offset ?? 0;

    let nextDoc;
    let insertedAt: number;
    let inheritedStyle: string | undefined;

    if (offset <= 0) {
      const newP = makeEmptyParagraph(ctx.mintNodeId, style);
      nextDoc = insertBlock(snapshot.root, at.paragraph, newP);
      insertedAt = at.paragraph;
    } else if (offset >= sourceText.length) {
      const newP = makeEmptyParagraph(ctx.mintNodeId, style);
      nextDoc = insertBlock(snapshot.root, at.paragraph + 1, newP);
      insertedAt = at.paragraph + 1;
    } else {
      const { left, right } = splitParagraph(sourceBlock, offset, ctx.mintNodeId);
      inheritedStyle = sourceBlock.properties.styleId;
      const styleForRight = style ?? inheritedStyle;
      const rightWithStyle: Paragraph = styleForRight
        ? { ...right, properties: { ...right.properties, styleId: styleForRight } }
        : right;
      const docWithLeft = replaceBlock(snapshot.root, at.paragraph, left);
      nextDoc = insertBlock(docWithLeft, at.paragraph + 1, rightWithStyle);
      insertedAt = at.paragraph + 1;
    }

    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    const insertedBlock = next.root.body[insertedAt];
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: insertedBlock.id,
        path: ["body", insertedAt],
        summary: `+paragraph${style || inheritedStyle ? ` (style=${style ?? inheritedStyle})` : ""}`,
      }),
    };
  },
};

function makeEmptyParagraph(mintNodeId: () => string, style: string | undefined): Paragraph {
  return {
    kind: "paragraph",
    id: mintNodeId(),
    properties: style ? { styleId: style } : {},
    children: [
      {
        kind: "run",
        id: mintNodeId(),
        properties: {},
        children: [],
      },
    ],
  };
}

/**
 * Split a paragraph at character offset. Walks the inline children in
 * document order, accumulating until offset is reached, then partitions.
 *
 * Limitations: we currently flatten hyperlinks/revisions on the boundary
 * by keeping them with whichever side they straddle into; full structural
 * splitting of those wrappers is tracked in docs/build-log/docx.md.
 */
function splitParagraph(
  p: Paragraph,
  offset: number,
  mintNodeId: () => string
): { left: Paragraph; right: Paragraph } {
  const leftChildren: InlineNode[] = [];
  const rightChildren: InlineNode[] = [];
  let consumed = 0;
  let placed = false;

  for (const child of p.children) {
    if (placed) {
      rightChildren.push(child);
      continue;
    }
    const len = inlineLength(child);
    if (consumed + len <= offset) {
      leftChildren.push(child);
      consumed += len;
      if (consumed === offset) {
        placed = true;
      }
      continue;
    }
    if (child.kind === "run") {
      const localOffset = offset - consumed;
      const { left, right } = splitRun(child, localOffset, mintNodeId);
      if (left) leftChildren.push(left);
      if (right) rightChildren.push(right);
      placed = true;
      continue;
    }
    if (child.kind === "hyperlink") {
      const localOffset = offset - consumed;
      const { left, right } = splitHyperlink(child, localOffset, mintNodeId);
      if (left) leftChildren.push(left);
      if (right) rightChildren.push(right);
      placed = true;
      continue;
    }
    leftChildren.push(child);
    consumed += len;
  }

  const ensureRun = (children: InlineNode[]): InlineNode[] => {
    if (children.some((c) => c.kind === "run")) return children;
    return [...children, { kind: "run", id: mintNodeId(), properties: {}, children: [] }];
  };

  const left: Paragraph = { ...p, children: ensureRun(leftChildren) };
  const right: Paragraph = { ...p, id: mintNodeId(), children: ensureRun(rightChildren) };
  return { left, right };
}

function splitRun(
  run: Run,
  offset: number,
  mintNodeId: () => string
): { left: Run | null; right: Run | null } {
  const leftChildren: RunChild[] = [];
  const rightChildren: RunChild[] = [];
  let consumed = 0;
  let placed = false;
  for (const child of run.children) {
    if (placed || child.kind !== "text") {
      if (placed) rightChildren.push(child);
      else leftChildren.push(child);
      continue;
    }
    const len = child.text.length;
    if (consumed + len <= offset) {
      leftChildren.push(child);
      consumed += len;
      if (consumed === offset) placed = true;
      continue;
    }
    const localOffset = offset - consumed;
    const before: TextLeaf = { ...child, text: child.text.slice(0, localOffset) };
    const after: TextLeaf = {
      ...child,
      id: mintNodeId(),
      text: child.text.slice(localOffset),
    };
    if (before.text.length > 0) leftChildren.push(before);
    if (after.text.length > 0) rightChildren.push(after);
    placed = true;
  }
  const left: Run | null = leftChildren.length > 0 ? { ...run, children: leftChildren } : null;
  const right: Run | null =
    rightChildren.length > 0 ? { ...run, id: mintNodeId(), children: rightChildren } : null;
  return { left, right };
}

function splitHyperlink(
  link: Hyperlink,
  offset: number,
  mintNodeId: () => string
): { left: Hyperlink | null; right: Hyperlink | null } {
  const leftRuns: Run[] = [];
  const rightRuns: Run[] = [];
  let consumed = 0;
  let placed = false;
  for (const r of link.children) {
    const len = runLength(r);
    if (placed) {
      rightRuns.push(r);
      continue;
    }
    if (consumed + len <= offset) {
      leftRuns.push(r);
      consumed += len;
      if (consumed === offset) placed = true;
      continue;
    }
    const localOffset = offset - consumed;
    const split = splitRun(r, localOffset, mintNodeId);
    if (split.left) leftRuns.push(split.left);
    if (split.right) rightRuns.push(split.right);
    placed = true;
  }
  const left = leftRuns.length > 0 ? { ...link, children: leftRuns } : null;
  const right = rightRuns.length > 0 ? { ...link, id: mintNodeId(), children: rightRuns } : null;
  return { left, right };
}

function inlineLength(node: InlineNode): number {
  if (node.kind === "run") return runLength(node);
  if (node.kind === "hyperlink") {
    let l = 0;
    for (const r of node.children) l += runLength(r);
    return l;
  }
  if (node.kind === "revision") {
    let l = 0;
    for (const ic of node.children) {
      if (ic.kind === "run") l += runLength(ic);
    }
    return l;
  }
  return 0;
}

function runLength(run: Run): number {
  let l = 0;
  for (const c of run.children) {
    if (c.kind === "text") l += c.text.length;
  }
  return l;
}
