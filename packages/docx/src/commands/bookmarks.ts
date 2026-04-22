import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  DocxDocument,
  DocxSnapshot,
  InlineNode,
  OpaqueInline,
  OpaqueXml,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";
import type { DeleteBookmarkPayload, InsertBookmarkPayload } from "./payloads.js";

/**
 * Word's "Bookmark" feature. Inserts a `<w:bookmarkStart>` /
 * `<w:bookmarkEnd>` pair anchored at flat-text byte offsets inside a
 * single body paragraph, the way Word's `Insert > Bookmark` dialog
 * does. The pair is carried as two {@link OpaqueInline} nodes so the
 * serializer round-trips them via the existing opaque-XML path —
 * Word's parser sees them as first-class metadata elements
 * (`opaque-classification.ts` already classifies them as `metadata`)
 * and our serializer's `opaque-inline` branch re-emits them
 * byte-equivalent.
 *
 * Cross-paragraph bookmarks are out of scope for v1: Word allows them
 * but they require split anchors that survive paragraph splits, which
 * is significant additional complexity. The handler explicitly rejects
 * paragraphIds that don't resolve and offsets that exceed the
 * paragraph's flat length.
 *
 * Idempotency: calling `insert-bookmark` with a `name` that already
 * exists in this paragraph deletes the old anchors first so the
 * "rename / move" flow is a single round-trip.
 */
export const insertBookmarkHandler: CommandHandler<InsertBookmarkPayload, DocxSnapshot> = {
  type: "docx:insert-bookmark",
  apply(snapshot, payload, ctx) {
    const name = (payload.name ?? "").trim();
    if (name.length === 0) {
      throw new CommandError("invalid-payload", "name must be a non-empty string");
    }
    if (!/^[A-Za-z_][\w]*$/.test(name)) {
      throw new CommandError(
        "invalid-payload",
        `bookmark name "${name}" must match Word's identifier rules ([A-Za-z_][\\w]*)`
      );
    }
    const { paragraphId, startOffset, endOffset } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (!Number.isInteger(startOffset) || startOffset < 0) {
      throw new CommandError("invalid-payload", `startOffset must be a non-negative integer (got ${startOffset})`);
    }
    if (!Number.isInteger(endOffset) || endOffset < startOffset) {
      throw new CommandError(
        "invalid-payload",
        `endOffset (${endOffset}) must be an integer >= startOffset (${startOffset})`
      );
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }
    const paragraph = located.paragraph;
    const flatLen = paragraphFlatLength(paragraph);
    if (startOffset > flatLen) {
      throw new CommandError(
        "invalid-payload",
        `startOffset (${startOffset}) exceeds paragraph flat length (${flatLen})`
      );
    }
    if (endOffset > flatLen) {
      throw new CommandError(
        "invalid-payload",
        `endOffset (${endOffset}) exceeds paragraph flat length (${flatLen})`
      );
    }

    const id = mintBookmarkId(snapshot.root);
    const stripped = stripBookmarkByName(paragraph, name);
    const withStart = insertOpaqueInlineAt(stripped, startOffset, makeBookmarkStart(id, name, ctx.mintNodeId));
    // Re-clamp end after a potential start insertion: the start anchor
    // doesn't consume flat-text bytes, so offsets stay stable, but we
    // still walk by flat offset and the start anchor occupies a slot
    // that splitRunAt skips.
    const withEnd = insertOpaqueInlineAt(withStart, endOffset, makeBookmarkEnd(id, ctx.mintNodeId));

    const nextDoc = located.replace(withEnd);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: paragraph.id,
        path: [...located.path],
        field: "bookmark",
        summary: `+bookmark "${name}" id=${id} ${startOffset}..${endOffset}`,
      }),
    };
  },
};

export const deleteBookmarkHandler: CommandHandler<DeleteBookmarkPayload, DocxSnapshot> = {
  type: "docx:delete-bookmark",
  apply(snapshot, payload) {
    const name = (payload.name ?? "").trim();
    if (name.length === 0) {
      throw new CommandError("invalid-payload", "name must be a non-empty string");
    }
    const found = findBookmarkByName(snapshot.root, name);
    if (!found) {
      // Idempotent no-op so callers can wire it to a "remove if exists"
      // affordance without pre-checking.
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: name,
          path: ["bookmarks"],
          field: "noop",
          summary: `bookmark "${name}" not present`,
        }),
      };
    }
    const stripped = stripBookmarkByName(found.paragraph, name);
    const nextDoc = found.replace(stripped);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: found.paragraph.id,
        path: [...found.path],
        field: "bookmark",
        summary: `-bookmark "${name}"`,
      }),
    };
  },
};

/**
 * Surface every named bookmark currently anchored in the body. Used
 * by the cross-reference dialog to populate its picker.
 */
export interface BookmarkAnchor {
  readonly name: string;
  readonly id: string;
  readonly paragraphId: string;
  readonly startOffset: number;
  readonly endOffset: number;
}

export function listBookmarks(doc: DocxDocument): ReadonlyArray<BookmarkAnchor> {
  const out: BookmarkAnchor[] = [];
  for (const block of doc.body) {
    if (block.kind !== "paragraph") continue;
    let consumed = 0;
    let openStart: { id: string; name: string; offset: number } | null = null;
    for (const child of block.children) {
      if (child.kind === "opaque-inline") {
        const tag = child.raw.tag;
        if (tag === "w:bookmarkStart") {
          const id = child.raw.attrs["w:id"] ?? "";
          const name = child.raw.attrs["w:name"] ?? "";
          if (name && id) openStart = { id, name, offset: consumed };
        } else if (tag === "w:bookmarkEnd") {
          const id = child.raw.attrs["w:id"] ?? "";
          if (openStart && openStart.id === id) {
            out.push({
              name: openStart.name,
              id,
              paragraphId: block.id,
              startOffset: openStart.offset,
              endOffset: consumed,
            });
            openStart = null;
          }
        }
        continue;
      }
      if (child.kind === "run") consumed += runFlatLength(child);
    }
  }
  return out;
}

interface BookmarkLocation {
  readonly paragraph: Paragraph;
  readonly path: ReadonlyArray<string | number>;
  readonly replace: (next: Paragraph) => DocxDocument;
}

function findBookmarkByName(doc: DocxDocument, name: string): BookmarkLocation | null {
  for (const block of doc.body) {
    if (block.kind !== "paragraph") continue;
    if (paragraphHasBookmark(block, name)) {
      const located = locateParagraph(doc, block.id);
      if (located) return { paragraph: located.paragraph, path: located.path, replace: located.replace };
    }
  }
  return null;
}

function paragraphHasBookmark(p: Paragraph, name: string): boolean {
  for (const child of p.children) {
    if (
      child.kind === "opaque-inline" &&
      child.raw.tag === "w:bookmarkStart" &&
      (child.raw.attrs["w:name"] ?? "") === name
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Strip every `<w:bookmarkStart>` / `<w:bookmarkEnd>` pair carrying the
 * given name from a paragraph. Pairs are matched by `w:id` so a stray
 * end anchor without a matching start (malformed input) is left alone.
 */
function stripBookmarkByName(p: Paragraph, name: string): Paragraph {
  const removeIds = new Set<string>();
  for (const child of p.children) {
    if (
      child.kind === "opaque-inline" &&
      child.raw.tag === "w:bookmarkStart" &&
      (child.raw.attrs["w:name"] ?? "") === name
    ) {
      const id = child.raw.attrs["w:id"];
      if (id) removeIds.add(id);
    }
  }
  if (removeIds.size === 0) return p;
  const newChildren: InlineNode[] = [];
  for (const child of p.children) {
    if (child.kind === "opaque-inline") {
      const tag = child.raw.tag;
      const id = child.raw.attrs["w:id"];
      if ((tag === "w:bookmarkStart" || tag === "w:bookmarkEnd") && id && removeIds.has(id)) continue;
    }
    newChildren.push(child);
  }
  return { ...p, children: newChildren };
}

function makeBookmarkStart(id: string, name: string, mint: IdMinter): OpaqueInline {
  return {
    kind: "opaque-inline",
    id: mint(),
    raw: makeOpaqueXml("w:bookmarkStart", { "w:id": id, "w:name": name }),
  };
}

function makeBookmarkEnd(id: string, mint: IdMinter): OpaqueInline {
  return {
    kind: "opaque-inline",
    id: mint(),
    raw: makeOpaqueXml("w:bookmarkEnd", { "w:id": id }),
  };
}

function makeOpaqueXml(tag: string, attrs: Record<string, string>): OpaqueXml {
  const rawAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) rawAttrs[`@_${k}`] = v;
  return { tag, attrs, subtree: [], rawAttrs };
}

/**
 * Insert an `OpaqueInline` carrier at the given flat-text offset in
 * the paragraph. The carrier itself contributes zero flat-text bytes,
 * so consecutive insertions at the same offset stack in insertion
 * order — that's exactly what we want for `<w:bookmarkStart>` followed
 * by `<w:bookmarkEnd>` at a zero-length anchor.
 */
function insertOpaqueInlineAt(p: Paragraph, offset: number, anchor: OpaqueInline): Paragraph {
  const flatLength = paragraphFlatLength(p);
  const clampedOffset = Math.max(0, Math.min(offset, flatLength));
  let consumed = 0;
  const newChildren: InlineNode[] = [];
  let placed = false;

  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    if (placed || child.kind !== "run") {
      newChildren.push(child);
      continue;
    }
    const runLength = runFlatLength(child);
    if (clampedOffset >= consumed && clampedOffset <= consumed + runLength) {
      const localOffset = clampedOffset - consumed;
      const { before, after } = splitRunAt(child, localOffset);
      if (before) newChildren.push(before);
      newChildren.push(anchor);
      if (after) newChildren.push(after);
      placed = true;
    } else {
      newChildren.push(child);
    }
    consumed += runLength;
  }
  if (!placed) {
    newChildren.push(anchor);
  }
  return { ...p, children: newChildren };
}

function paragraphFlatLength(p: Paragraph): number {
  let n = 0;
  for (const c of p.children) {
    if (c.kind === "run") n += runFlatLength(c);
  }
  return n;
}

function runFlatLength(r: Run): number {
  let n = 0;
  for (const c of r.children) {
    if (c.kind === "text") n += c.text.length;
  }
  return n;
}

interface SplitRun {
  readonly before: Run | null;
  readonly after: Run | null;
}

function splitRunAt(run: Run, localOffset: number): SplitRun {
  if (localOffset <= 0) return { before: null, after: run };
  const total = runFlatLength(run);
  if (localOffset >= total) return { before: run, after: null };

  const beforeChildren: RunChild[] = [];
  const afterChildren: RunChild[] = [];
  let consumed = 0;
  let split = false;
  for (const child of run.children) {
    if (split) {
      afterChildren.push(child);
      continue;
    }
    if (child.kind !== "text") {
      beforeChildren.push(child);
      continue;
    }
    const len = child.text.length;
    if (localOffset >= consumed && localOffset <= consumed + len) {
      const inLeaf = localOffset - consumed;
      const beforeText = child.text.slice(0, inLeaf);
      const afterText = child.text.slice(inLeaf);
      if (beforeText.length > 0) {
        const beforeLeaf: TextLeaf = { ...child, text: beforeText };
        beforeChildren.push(beforeLeaf);
      }
      if (afterText.length > 0) {
        const afterLeaf: TextLeaf = { ...child, text: afterText };
        afterChildren.push(afterLeaf);
      }
      split = true;
    } else {
      beforeChildren.push(child);
    }
    consumed += len;
  }
  const before: Run | null =
    beforeChildren.length > 0
      ? { kind: "run", id: run.id, properties: run.properties, children: beforeChildren }
      : null;
  const after: Run | null =
    afterChildren.length > 0
      ? { kind: "run", id: `${run.id}-bm-tail`, properties: run.properties, children: afterChildren }
      : null;
  return { before, after };
}

function mintBookmarkId(doc: DocxDocument): string {
  let max = -1;
  for (const block of doc.body) {
    if (block.kind !== "paragraph") continue;
    for (const child of block.children) {
      if (
        child.kind === "opaque-inline" &&
        (child.raw.tag === "w:bookmarkStart" || child.raw.tag === "w:bookmarkEnd")
      ) {
        const raw = child.raw.attrs["w:id"];
        const n = raw ? Number.parseInt(raw, 10) : Number.NaN;
        if (Number.isFinite(n) && n > max) max = n;
      }
    }
  }
  return String(max + 1);
}
