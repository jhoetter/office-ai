import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxSnapshot,
  InlineNode,
  Paragraph,
  RevisionWrapper,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, textLeaf, withParagraph } from "./helpers.js";
import type { InsertTextTrackedPayload } from "./payloads.js";

/**
 * Tracked-changes ("Suggesting" mode) variant of `docx:insert-text`.
 *
 * Instead of splicing the new text directly into the targeted run,
 * we wrap a fresh single-leaf run in a `<w:ins>` revision wrapper
 * and splice that wrapper into the paragraph at the requested
 * position. The wrapper carries `author`, `date`, and `revisionId`
 * so the tracked-changes ribbon, the inline hover overlay, and the
 * `accept-change` / `reject-change` handlers can all operate on it
 * with no further plumbing.
 *
 * Position semantics match `docx:insert-text`:
 *   - `at.run === undefined` → insert at the start of the paragraph
 *     (prepend the wrapper),
 *   - `at.run` set + `at.offset === 0` → insert before that run,
 *   - `at.run` set + `at.offset === runLength` → insert after that
 *     run,
 *   - `at.run` set + `0 < at.offset < runLength` → split the run at
 *     `offset` and insert the wrapper between the two halves. The
 *     run's properties carry over to both halves so formatting is
 *     preserved.
 *
 * Round-trip property: serializing the resulting snapshot produces a
 * `<w:ins w:id="…" w:author="…" w:date="…"><w:r><w:t>…</w:t></w:r></w:ins>`
 * element identical in shape to what Word writes. Re-parsing yields
 * the same `RevisionWrapper`, and `accept-change` folds it back into
 * the body.
 */
export const insertTextTrackedHandler: CommandHandler<InsertTextTrackedPayload, DocxSnapshot> = {
  type: "docx:insert-text-tracked",
  apply(snapshot, payload, ctx) {
    const { at, text, author } = payload;
    if (!text) {
      return {
        next: { ...snapshot, revision: snapshot.revision + 1 },
        diff: buildDiff(snapshot.revision, snapshot.revision + 1, {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["body"],
          field: "noop",
          summary: "no-op insert-text-tracked",
        }),
      };
    }
    if (!author) {
      throw new CommandError("invalid-payload", "insert-text-tracked requires a non-empty author");
    }
    if (at.paragraph < 0 || at.paragraph >= snapshot.root.body.length) {
      throw new CommandError("invalid-position", `paragraph index ${at.paragraph} out of range`);
    }
    const targetBlock = snapshot.root.body[at.paragraph];
    if (targetBlock.kind !== "paragraph") {
      throw new CommandError("not-paragraph", `block at ${at.paragraph} is not a paragraph`);
    }

    const date = payload.date ?? new Date().toISOString();
    // Word coalesces consecutive same-author insertions in Suggesting
    // mode into a single revision (one balloon, one accept/reject).
    // When the caller pins an explicit `revisionId` we honour that
    // intent — they're explicitly addressing a distinct wrapper —
    // otherwise we look for an adjacent same-author `<w:ins>` to
    // extend. The PM funnel never supplies a `revisionId`, so normal
    // typing always coalesces.
    const canCoalesce = payload.revisionId === undefined;
    const revisionId = payload.revisionId ?? mintRevisionId(snapshot);

    const nextDoc = withParagraph(snapshot.root, at.paragraph, (p) =>
      insertRevisionIntoParagraph(
        p,
        at.run,
        at.offset ?? 0,
        text,
        author,
        date,
        revisionId,
        ctx.mintNodeId,
        canCoalesce
      )
    );

    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: targetBlock.id,
        path: ["body", at.paragraph],
        field: "tracked-insert",
        summary: `+ins ${JSON.stringify(text)}`,
      }),
    };
  },
};

/**
 * Mint a fresh `w:id` for a new revision wrapper. We pick the
 * smallest positive integer that does not collide with any existing
 * revision in the document body OR in any header / footer part. The
 * value is returned as a string because OOXML `w:id` is a decimal
 * string and downstream code (parser, accept-change) compares
 * strings.
 */
function mintRevisionId(snapshot: DocxSnapshot): string {
  const used = new Set<number>();
  for (const b of snapshot.root.body) visitBlockForRevisions(b, used);
  for (const part of snapshot.root.headersAndFooters) {
    for (const b of part.body) visitBlockForRevisions(b, used);
  }
  let id = 1;
  while (used.has(id)) id++;
  return String(id);
}

function visitBlockForRevisions(block: BlockNode, used: Set<number>): void {
  if (block.kind !== "paragraph") return;
  for (const child of block.children) collectFromInline(child, used);
}

function collectFromInline(node: InlineNode, used: Set<number>): void {
  if (node.kind === "revision") {
    const n = Number(node.revisionId);
    if (Number.isFinite(n) && n > 0 && Math.floor(n) === n) used.add(n);
    for (const c of node.children) collectFromInline(c, used);
  }
}

function insertRevisionIntoParagraph(
  p: Paragraph,
  runIndex: number | undefined,
  offset: number,
  text: string,
  author: string,
  date: string,
  revisionId: string,
  mintNodeId: () => string,
  canCoalesce: boolean
): Paragraph {
  if (runIndex === undefined) {
    // Paragraph-global offset mode. Walk children counting visible
    // text characters (treating revision wrappers, hyperlinks, and
    // other inline containers as opaque text spans of their inner
    // length) and splice the wrapper at the right boundary,
    // splitting a run when the offset lands inside one. This is
    // what callers from the PM funnel use because a paragraph's
    // run/wrapper structure is a model concept that PM positions
    // don't expose.
    return spliceWrapperAtParagraphOffset(p, offset, text, author, date, revisionId, mintNodeId, canCoalesce);
  }

  const wrapper = makeInsRevision(text, author, date, revisionId, mintNodeId);
  const target = p.children[runIndex];
  if (!target || target.kind !== "run") {
    // Non-run target (existing revision wrapper, hyperlink, opaque
    // inline, etc.). Fall back to splicing the wrapper in front of
    // the targeted child without trying to split it; this matches
    // the conservative behaviour of `insert-text` for the same case.
    if (canCoalesce) {
      const merged = tryCoalesceAtBoundary(p, runIndex, text, author, mintNodeId);
      if (merged) return merged;
    }
    const children = [...p.children];
    children.splice(runIndex, 0, wrapper);
    return { ...p, children };
  }

  const runLength = runTextLength(target);
  if (offset <= 0) {
    if (canCoalesce) {
      const merged = tryCoalesceAtBoundary(p, runIndex, text, author, mintNodeId);
      if (merged) return merged;
    }
    const children = [...p.children];
    children.splice(runIndex, 0, wrapper);
    return { ...p, children };
  }
  if (offset >= runLength) {
    if (canCoalesce) {
      const merged = tryCoalesceAtBoundary(p, runIndex + 1, text, author, mintNodeId);
      if (merged) return merged;
    }
    const children = [...p.children];
    children.splice(runIndex + 1, 0, wrapper);
    return { ...p, children };
  }

  // Split the run at `offset` so the wrapper lands between the two
  // halves. We split at the text-leaf boundary and rebuild two runs
  // that share the original run's properties.
  const [before, after] = splitRunAt(target, offset, mintNodeId);
  const children = [...p.children];
  children.splice(runIndex, 1, before, wrapper, after);
  return { ...p, children };
}

/**
 * Attempt to extend a neighbouring same-author `<w:ins>` wrapper at
 * the splice boundary `boundaryIndex` (the index where a fresh
 * wrapper would otherwise be inserted into `p.children`).
 *
 * Boundary semantics: a splice at index `i` lands between
 * `children[i-1]` and `children[i]`. We prefer extending the
 * predecessor (append) to match the typing flow where the cursor
 * advances past the previously inserted character; if that fails we
 * try the successor (prepend) so typing immediately in front of an
 * existing tracked insert still merges.
 *
 * Returns `null` when neither neighbour is a same-author `<w:ins>`
 * wrapper, in which case the caller falls back to creating a fresh
 * wrapper as before.
 */
function tryCoalesceAtBoundary(
  p: Paragraph,
  boundaryIndex: number,
  text: string,
  author: string,
  mintNodeId: () => string
): Paragraph | null {
  const prev = boundaryIndex > 0 ? p.children[boundaryIndex - 1] : undefined;
  if (prev && isCoalescableInsRevision(prev, author)) {
    const extended = appendTextToInsRevision(prev, text, mintNodeId);
    const children = [...p.children];
    children[boundaryIndex - 1] = extended;
    return { ...p, children };
  }
  const next = boundaryIndex < p.children.length ? p.children[boundaryIndex] : undefined;
  if (next && isCoalescableInsRevision(next, author)) {
    const extended = prependTextToInsRevision(next, text, mintNodeId);
    const children = [...p.children];
    children[boundaryIndex] = extended;
    return { ...p, children };
  }
  return null;
}

function isCoalescableInsRevision(node: InlineNode, author: string): node is RevisionWrapper {
  return node.kind === "revision" && node.revisionType === "ins" && node.author === author;
}

/**
 * Append `text` to the end of an `<w:ins>` wrapper. We extend the
 * last text leaf of the last inner run when one exists so the
 * serialized output stays one `<w:t>` per contiguous run; otherwise
 * we add a fresh text leaf (or a fresh inner run when the wrapper
 * is empty).
 */
function appendTextToInsRevision(
  wrapper: RevisionWrapper,
  text: string,
  mintNodeId: () => string
): RevisionWrapper {
  const children = [...wrapper.children];
  for (let i = children.length - 1; i >= 0; i--) {
    const child = children[i];
    if (child.kind !== "run") continue;
    const runChildren = [...child.children];
    for (let j = runChildren.length - 1; j >= 0; j--) {
      const leaf = runChildren[j];
      if (leaf.kind === "text") {
        const merged: TextLeaf = {
          ...leaf,
          text: leaf.text + text,
          xmlSpacePreserve: /^\s|\s$/.test(leaf.text + text) || (leaf.text + text).length === 0,
        };
        runChildren[j] = merged;
        children[i] = { ...child, children: runChildren };
        return { ...wrapper, children };
      }
    }
    runChildren.push(textLeaf(mintNodeId, text));
    children[i] = { ...child, children: runChildren };
    return { ...wrapper, children };
  }
  const innerRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [textLeaf(mintNodeId, text)],
  };
  children.push(innerRun);
  return { ...wrapper, children };
}

/**
 * Prepend `text` to the start of an `<w:ins>` wrapper. Mirrors
 * {@link appendTextToInsRevision} for the leading edge.
 */
function prependTextToInsRevision(
  wrapper: RevisionWrapper,
  text: string,
  mintNodeId: () => string
): RevisionWrapper {
  const children = [...wrapper.children];
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (child.kind !== "run") continue;
    const runChildren = [...child.children];
    for (let j = 0; j < runChildren.length; j++) {
      const leaf = runChildren[j];
      if (leaf.kind === "text") {
        const merged: TextLeaf = {
          ...leaf,
          text: text + leaf.text,
          xmlSpacePreserve: /^\s|\s$/.test(text + leaf.text) || (text + leaf.text).length === 0,
        };
        runChildren[j] = merged;
        children[i] = { ...child, children: runChildren };
        return { ...wrapper, children };
      }
    }
    runChildren.unshift(textLeaf(mintNodeId, text));
    children[i] = { ...child, children: runChildren };
    return { ...wrapper, children };
  }
  const innerRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [textLeaf(mintNodeId, text)],
  };
  children.unshift(innerRun);
  return { ...wrapper, children };
}

/**
 * Splice a revision wrapper into a paragraph at a paragraph-global
 * text offset (counting characters across all inline children,
 * including ones inside existing revision wrappers and hyperlinks).
 *
 *   - If the offset lands on a child boundary, the wrapper is
 *     spliced between children with no splitting.
 *   - If the offset lands inside a run, the run is split at that
 *     point (inheriting run properties on both halves) and the
 *     wrapper goes between the halves.
 *   - If the offset lands inside a non-run container (existing
 *     revision wrapper, hyperlink, opaque inline), we splice the
 *     wrapper in front of that container — splitting opaque
 *     containers is out of scope and would change author
 *     attribution semantics.
 */
function spliceWrapperAtParagraphOffset(
  p: Paragraph,
  offset: number,
  text: string,
  author: string,
  date: string,
  revisionId: string,
  mintNodeId: () => string,
  canCoalesce: boolean
): Paragraph {
  const makeWrapper = (): RevisionWrapper => makeInsRevision(text, author, date, revisionId, mintNodeId);

  if (offset <= 0 || p.children.length === 0) {
    if (canCoalesce) {
      const merged = tryCoalesceAtBoundary(p, 0, text, author, mintNodeId);
      if (merged) return merged;
    }
    return { ...p, children: [makeWrapper(), ...p.children] };
  }
  let consumed = 0;
  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    const childLen = inlineTextLength(child);
    const childEnd = consumed + childLen;
    if (offset === consumed) {
      if (canCoalesce) {
        const merged = tryCoalesceAtBoundary(p, i, text, author, mintNodeId);
        if (merged) return merged;
      }
      const children = [...p.children];
      children.splice(i, 0, makeWrapper());
      return { ...p, children };
    }
    if (offset > consumed && offset < childEnd) {
      if (child.kind === "run") {
        const [before, after] = splitRunAt(child, offset - consumed, mintNodeId);
        const children = [...p.children];
        children.splice(i, 1, before, makeWrapper(), after);
        return { ...p, children };
      }
      // Opaque-ish container: don't split, splice in front of it.
      // Coalescing applies if the container immediately preceding
      // the splice point is a same-author `<w:ins>` we can extend.
      if (canCoalesce) {
        const merged = tryCoalesceAtBoundary(p, i, text, author, mintNodeId);
        if (merged) return merged;
      }
      const children = [...p.children];
      children.splice(i, 0, makeWrapper());
      return { ...p, children };
    }
    consumed = childEnd;
  }
  if (canCoalesce) {
    const merged = tryCoalesceAtBoundary(p, p.children.length, text, author, mintNodeId);
    if (merged) return merged;
  }
  return { ...p, children: [...p.children, makeWrapper()] };
}

function inlineTextLength(node: InlineNode): number {
  if (node.kind === "run") return runTextLength(node);
  if (node.kind === "revision") {
    let n = 0;
    for (const c of node.children) n += inlineTextLength(c);
    return n;
  }
  if (node.kind === "hyperlink") {
    let n = 0;
    for (const c of node.children) n += inlineTextLength(c);
    return n;
  }
  return 0;
}

function makeInsRevision(
  text: string,
  author: string,
  date: string,
  revisionId: string,
  mintNodeId: () => string
): RevisionWrapper {
  const innerRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [textLeaf(mintNodeId, text)],
  };
  return {
    kind: "revision",
    id: mintNodeId(),
    revisionType: "ins",
    author,
    date,
    revisionId,
    children: [innerRun],
  };
}

function runTextLength(r: Run): number {
  let n = 0;
  for (const c of r.children) if (c.kind === "text") n += c.text.length;
  return n;
}

/**
 * Split a run at character offset `offset` (1 <= offset <= runLength
 * - 1; callers handle the boundary cases). Both halves keep the
 * original run's `properties` so visual formatting is preserved
 * around the inserted revision wrapper.
 */
function splitRunAt(run: Run, offset: number, mintNodeId: () => string): [Run, Run] {
  let consumed = 0;
  const beforeChildren: RunChild[] = [];
  const afterChildren: RunChild[] = [];
  for (const child of run.children) {
    if (child.kind !== "text") {
      // Non-text run children (breaks, fields, etc.) bind to whichever
      // half their position falls in. They contribute zero to the
      // text-offset count, so anything before the placed split goes
      // into `before`, the rest into `after`.
      if (consumed < offset) beforeChildren.push(child);
      else afterChildren.push(child);
      continue;
    }
    const len = child.text.length;
    const start = consumed;
    const end = consumed + len;
    if (end <= offset) {
      beforeChildren.push(child);
    } else if (start >= offset) {
      afterChildren.push(child);
    } else {
      const cut = offset - start;
      const head: TextLeaf = { ...child, text: child.text.slice(0, cut) };
      const tail: TextLeaf = { ...child, id: mintNodeId(), text: child.text.slice(cut) };
      beforeChildren.push(head);
      afterChildren.push(tail);
    }
    consumed = end;
  }
  const before: Run = { ...run, children: beforeChildren };
  const after: Run = { ...run, id: mintNodeId(), children: afterChildren };
  return [before, after];
}
