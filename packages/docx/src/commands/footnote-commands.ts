import { CommandError, type CommandHandler, type IdMinter, type NodeId } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxSnapshot,
  Footnote,
  FootnoteReferenceLeaf,
  FootnotesPart,
  HeaderFooterPart,
  InlineNode,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";

/* ── Payloads (re-exported through ./payloads.js) ────────────────────────── */

export interface InsertFootnotePayload {
  /** Stable id of the body paragraph (or any header/footer paragraph). */
  paragraphId: NodeId;
  /** Byte offset in the paragraph's flat text. Clamped to [0, length]. */
  offset: number;
  /**
   * Body of the new footnote. Defaults to a single empty paragraph
   * styled `FootnoteText` (the canonical Word default).
   */
  body?: ReadonlyArray<BlockNode>;
}

export interface SetFootnoteBodyPayload {
  footnoteId: number;
  body: ReadonlyArray<BlockNode>;
}

export interface DeleteFootnotePayload {
  footnoteId: number;
}

/* ── Shared helpers ──────────────────────────────────────────────────────── */

const FOOTNOTE_TEXT_STYLE = "FootnoteText";
const FOOTNOTE_REF_STYLE = "FootnoteReference";

/**
 * Allocate the next OOXML id for a new footnote. The reserved values
 * `-1` (separator) and `0` (continuationSeparator) are skipped; the
 * first authored footnote in a fresh document is `1`. We mint
 * `max(existing) + 1` so ids stay monotonic — Word does the same.
 */
function nextFootnoteId(part: FootnotesPart | undefined): number {
  if (!part || part.footnotes.length === 0) return 1;
  let max = 0;
  for (const fn of part.footnotes) {
    if (fn.id > max) max = fn.id;
  }
  return Math.max(max + 1, 1);
}

/**
 * Default body for a freshly inserted footnote: one empty paragraph
 * styled `FootnoteText` with a leading reference run carrying the
 * footnote glyph (`<w:footnoteRef/>` in Word — opaque-XML for now)
 * followed by an empty text run the user can type into.
 *
 * F1 keeps the body lean: a single empty paragraph with the right
 * style. The leading `<w:footnoteRef/>` is part of Word's default
 * skeleton but is rendered automatically by Word from the style; we
 * do not need to synthesize it here for the round-trip invariants in
 * the spec to hold.
 */
function defaultFootnoteBody(mintNodeId: IdMinter): BlockNode[] {
  const para: Paragraph = {
    kind: "paragraph",
    id: mintNodeId(),
    properties: { styleId: FOOTNOTE_TEXT_STYLE },
    children: [
      {
        kind: "run",
        id: mintNodeId(),
        properties: {},
        children: [],
      },
    ],
  };
  return [para];
}

interface ParagraphLocation {
  readonly kind: "body" | "header-footer";
  readonly bodyIndex?: number;
  readonly partIdx?: number;
  readonly paragraphIdx: number;
  readonly paragraph: Paragraph;
}

function findParagraph(snapshot: DocxSnapshot, paragraphId: NodeId): ParagraphLocation | null {
  for (let i = 0; i < snapshot.root.body.length; i++) {
    const block = snapshot.root.body[i];
    if (block.kind === "paragraph" && block.id === paragraphId) {
      return { kind: "body", bodyIndex: i, paragraphIdx: i, paragraph: block };
    }
  }
  for (let i = 0; i < snapshot.root.headersAndFooters.length; i++) {
    const part = snapshot.root.headersAndFooters[i];
    for (let j = 0; j < part.body.length; j++) {
      const block = part.body[j];
      if (block.kind === "paragraph" && block.id === paragraphId) {
        return { kind: "header-footer", partIdx: i, paragraphIdx: j, paragraph: block };
      }
    }
  }
  return null;
}

/**
 * Replace one paragraph in a `HeaderFooterPart` and return the
 * resulting parts array plus the merged dirty set. Mirrors the
 * pattern used by `insert-page-number` but inlined here so the
 * footnote commands don't reach across files for a one-liner.
 */
function replaceHeaderFooterParagraph(
  parts: ReadonlyArray<HeaderFooterPart>,
  partIdx: number,
  paragraphIdx: number,
  next: Paragraph
): HeaderFooterPart[] {
  const part = parts[partIdx];
  const newBody: BlockNode[] = part.body.slice();
  newBody[paragraphIdx] = next;
  const updated: HeaderFooterPart = { ...part, body: newBody };
  const out = parts.slice();
  out[partIdx] = updated;
  return out;
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

function splitRunAt(run: Run, localOffset: number, mintNodeId: IdMinter): SplitRun {
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
      const localInLeaf = localOffset - consumed;
      const beforeText = child.text.slice(0, localInLeaf);
      const afterText = child.text.slice(localInLeaf);
      if (beforeText.length > 0) {
        const beforeLeaf: TextLeaf = { ...child, text: beforeText };
        beforeChildren.push(beforeLeaf);
      }
      if (afterText.length > 0) {
        const afterLeaf: TextLeaf = { ...child, id: mintNodeId(), text: afterText };
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
      ? { kind: "run", id: mintNodeId(), properties: run.properties, children: afterChildren }
      : null;
  return { before, after };
}

/**
 * Splice a fresh `FootnoteReferenceLeaf` (wrapped in a zero-text
 * reference run styled `FootnoteReference`) at `offset` inside `p`.
 * Mirrors the run-splitting rules used by `insert-page-number`.
 */
function insertFootnoteRefIntoParagraph(
  p: Paragraph,
  offset: number,
  footnoteId: number,
  mintNodeId: IdMinter
): Paragraph {
  const refLeaf: FootnoteReferenceLeaf = {
    kind: "footnote-ref",
    id: mintNodeId(),
    footnoteId,
  };
  const refRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [refLeaf],
  };
  // The Word convention is to style the reference run with
  // `FootnoteReference`. We could surface that as a typed `rStyle`
  // field, but the existing model captures `<w:rStyle>` through the
  // style cascade only — synthesising one here would require touching
  // the styles part as well. F1 pragmatically emits the run with no
  // explicit style and lets the reader's style cascade carry the
  // superscript formatting from the run's character style if present.
  // Documented as a deferred polish item in docs/build-log/docx.md.
  void FOOTNOTE_REF_STYLE;

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
    const length = runFlatLength(child);
    if (clampedOffset >= consumed && clampedOffset <= consumed + length) {
      const localOffset = clampedOffset - consumed;
      const { before, after } = splitRunAt(child, localOffset, mintNodeId);
      if (before) newChildren.push(before);
      newChildren.push(refRun);
      if (after) newChildren.push(after);
      placed = true;
    } else {
      newChildren.push(child);
    }
    consumed += length;
  }

  if (!placed) {
    newChildren.push(refRun);
  }

  return { ...p, children: newChildren };
}

/**
 * Walk every run in a paragraph and strip any `FootnoteReferenceLeaf`
 * whose `footnoteId` matches `targetId`. Returns the same paragraph
 * (referentially equal) when no leaf was matched, so callers can
 * detect "did we change anything?" cheaply.
 */
function stripFootnoteReferencesFromParagraph(p: Paragraph, targetId: number): Paragraph {
  let changed = false;
  const newChildren: InlineNode[] = [];
  for (const child of p.children) {
    if (child.kind !== "run") {
      newChildren.push(child);
      continue;
    }
    const filteredRunChildren: RunChild[] = [];
    let runChanged = false;
    for (const rc of child.children) {
      if (rc.kind === "footnote-ref" && rc.footnoteId === targetId) {
        runChanged = true;
        continue;
      }
      filteredRunChildren.push(rc);
    }
    if (runChanged) {
      changed = true;
      newChildren.push({ ...child, children: filteredRunChildren });
    } else {
      newChildren.push(child);
    }
  }
  if (!changed) return p;
  return { ...p, children: newChildren };
}

/**
 * Recursively strip footnote references from every paragraph reachable
 * from a block (including paragraphs nested inside tables and opaque
 * blocks with typed `children`). Returns the same block (referentially
 * equal) when nothing changed.
 */
function stripFootnoteReferencesFromBlock(block: BlockNode, targetId: number): BlockNode {
  switch (block.kind) {
    case "paragraph":
      return stripFootnoteReferencesFromParagraph(block, targetId);
    case "table": {
      let changed = false;
      const rows = block.rows.map((row) => {
        let rowChanged = false;
        const cells = row.cells.map((cell) => {
          let cellChanged = false;
          const body = cell.body.map((b) => {
            const next = stripFootnoteReferencesFromBlock(b, targetId);
            if (next !== b) {
              cellChanged = true;
              changed = true;
            }
            return next;
          });
          if (!cellChanged) return cell;
          rowChanged = true;
          return { ...cell, body };
        });
        if (!rowChanged) return row;
        return { ...row, cells };
      });
      if (!changed) return block;
      // Mutating a table forces the per-table dirty marker — drop
      // `raw` so the serializer regenerates it from the typed model.
      return { ...block, rows, raw: undefined };
    }
    case "opaque-block": {
      if (!block.children || block.children.length === 0) return block;
      let changed = false;
      const children = block.children.map((c) => {
        const next = stripFootnoteReferencesFromBlock(c, targetId);
        if (next !== c) changed = true;
        return next;
      });
      if (!changed) return block;
      return { ...block, children, subtreeDirty: true };
    }
    case "section-break":
    case "wrapper-marker":
      return block;
    default: {
      const _exhaustive: never = block;
      void _exhaustive;
      return block;
    }
  }
}

/* ── docx:insert-footnote ────────────────────────────────────────────────── */

export const insertFootnoteHandler: CommandHandler<InsertFootnotePayload, DocxSnapshot> = {
  type: "docx:insert-footnote",
  apply(snapshot, payload, ctx) {
    if (!payload.paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (!Number.isInteger(payload.offset) || payload.offset < 0) {
      throw new CommandError(
        "invalid-payload",
        `offset must be a non-negative integer (got ${payload.offset})`
      );
    }

    const located = findParagraph(snapshot, payload.paragraphId);
    if (!located) {
      throw new CommandError(
        "unknown-target",
        `no body or header/footer paragraph with id "${payload.paragraphId}"`
      );
    }

    const id = nextFootnoteId(snapshot.root.footnotesPart);
    const body = payload.body && payload.body.length > 0 ? payload.body : defaultFootnoteBody(ctx.mintNodeId);
    const newFootnote: Footnote = {
      id,
      type: "normal",
      body,
      // No `raw` — the serializer regenerates from typed fields.
    };

    const prev = snapshot.root.footnotesPart;
    const nextPart: FootnotesPart = prev
      ? { ...prev, footnotes: [...prev.footnotes, newFootnote] }
      : { footnotes: [newFootnote], rootAttrs: {} };

    const updatedParagraph = insertFootnoteRefIntoParagraph(
      located.paragraph,
      payload.offset,
      id,
      ctx.mintNodeId
    );

    let nextDoc: DocxDocument;
    let dirtyHeadersAndFooters = snapshot.dirty.headersAndFooters;
    let dirtyBody = snapshot.dirty.body;

    if (located.kind === "body" && located.bodyIndex !== undefined) {
      const newBody = snapshot.root.body.slice();
      newBody[located.bodyIndex] = updatedParagraph;
      nextDoc = { ...snapshot.root, body: newBody, footnotesPart: nextPart };
      dirtyBody = true;
    } else if (located.kind === "header-footer" && located.partIdx !== undefined) {
      const newParts = replaceHeaderFooterParagraph(
        snapshot.root.headersAndFooters,
        located.partIdx,
        located.paragraphIdx,
        updatedParagraph
      );
      nextDoc = { ...snapshot.root, headersAndFooters: newParts, footnotesPart: nextPart };
      const partPath = newParts[located.partIdx].partPath;
      dirtyHeadersAndFooters = withAddition(snapshot.dirty.headersAndFooters, partPath);
    } else {
      // Defensive — `findParagraph` only returns the two cases above.
      throw new CommandError("unknown-target", "could not place footnote reference");
    }

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: dirtyBody,
      headersAndFooters: dirtyHeadersAndFooters,
      footnotes: true,
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: String(id),
        path: ["footnotesPart", "footnotes", nextPart.footnotes.length - 1],
        summary: `+footnote id=${id} ref@${located.paragraph.id}:${payload.offset}`,
      }),
    };
  },
};

/* ── docx:set-footnote-body ──────────────────────────────────────────────── */

export const setFootnoteBodyHandler: CommandHandler<SetFootnoteBodyPayload, DocxSnapshot> = {
  type: "docx:set-footnote-body",
  apply(snapshot, payload) {
    if (!Number.isInteger(payload.footnoteId)) {
      throw new CommandError("invalid-payload", `footnoteId must be an integer (got ${payload.footnoteId})`);
    }
    if (!Array.isArray(payload.body) || payload.body.length === 0) {
      throw new CommandError("invalid-payload", "body must be a non-empty Block[]");
    }
    const part = snapshot.root.footnotesPart;
    if (!part) {
      throw new CommandError("unknown-target", `no footnotesPart on document`);
    }
    const idx = part.footnotes.findIndex((f) => f.id === payload.footnoteId);
    if (idx < 0) {
      throw new CommandError("unknown-target", `no footnote with id ${payload.footnoteId}`);
    }
    const existing = part.footnotes[idx];
    const replaced: Footnote = {
      id: existing.id,
      type: existing.type,
      body: payload.body,
      // raw deliberately dropped — the serializer regenerates this
      // footnote's bytes while siblings re-emit verbatim from their
      // own `raw`.
    };
    const newFootnotes = part.footnotes.slice();
    newFootnotes[idx] = replaced;
    const nextPart: FootnotesPart = { ...part, footnotes: newFootnotes };
    const nextDoc: DocxDocument = { ...snapshot.root, footnotesPart: nextPart };
    const next = evolveSnapshot(snapshot, nextDoc, { footnotes: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: String(existing.id),
        path: ["footnotesPart", "footnotes", idx],
        field: "body",
        summary: `replaced body of footnote ${existing.id}`,
      }),
    };
  },
};

/* ── docx:delete-footnote ────────────────────────────────────────────────── */

export const deleteFootnoteHandler: CommandHandler<DeleteFootnotePayload, DocxSnapshot> = {
  type: "docx:delete-footnote",
  apply(snapshot, payload) {
    if (!Number.isInteger(payload.footnoteId)) {
      throw new CommandError("invalid-payload", `footnoteId must be an integer (got ${payload.footnoteId})`);
    }
    const part = snapshot.root.footnotesPart;
    if (!part) {
      throw new CommandError("unknown-target", `no footnotesPart on document`);
    }
    const idx = part.footnotes.findIndex((f) => f.id === payload.footnoteId);
    if (idx < 0) {
      throw new CommandError("unknown-target", `no footnote with id ${payload.footnoteId}`);
    }

    const removed = part.footnotes[idx];
    const newFootnotes = part.footnotes.slice();
    newFootnotes.splice(idx, 1);
    const nextPart: FootnotesPart = { ...part, footnotes: newFootnotes };

    // Strip body references.
    let bodyChanged = false;
    const newBody = snapshot.root.body.map((b) => {
      const next = stripFootnoteReferencesFromBlock(b, removed.id);
      if (next !== b) bodyChanged = true;
      return next;
    });

    // Strip header/footer references and accumulate dirty part paths.
    let dirtyHeadersAndFooters = snapshot.dirty.headersAndFooters;
    const newParts = snapshot.root.headersAndFooters.map((part) => {
      let partChanged = false;
      const newPartBody = part.body.map((b) => {
        const nextBlock = stripFootnoteReferencesFromBlock(b, removed.id);
        if (nextBlock !== b) partChanged = true;
        return nextBlock;
      });
      if (!partChanged) return part;
      dirtyHeadersAndFooters = withAddition(dirtyHeadersAndFooters, part.partPath);
      return { ...part, body: newPartBody };
    });

    const nextDoc: DocxDocument = {
      ...snapshot.root,
      body: newBody,
      headersAndFooters: newParts,
      footnotesPart: nextPart,
    };
    const next = evolveSnapshot(snapshot, nextDoc, {
      footnotes: true,
      ...(bodyChanged ? { body: true } : {}),
      headersAndFooters: dirtyHeadersAndFooters,
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: String(removed.id),
        path: ["footnotesPart", "footnotes", idx],
        summary: `-footnote id=${removed.id}`,
      }),
    };
  },
};

function withAddition(set: ReadonlySet<string>, path: string): ReadonlySet<string> {
  if (set.has(path)) return set;
  const next = new Set(set);
  next.add(path);
  return next;
}
