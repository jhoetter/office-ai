import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  DocxDocument,
  DocxSnapshot,
  Hyperlink,
  InlineNode,
  Paragraph,
  Relationship,
  Run,
  RunChild,
} from "../model/types.js";
import { buildDiffMulti, evolveSnapshot } from "./helpers.js";
import { locateParagraph } from "./set-paragraph-list.js";
import type { InsertHyperlinkPayload } from "./payloads.js";

const HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const DOC_RELS_KEY = "word/document.xml";

/**
 * Wrap a flat-text range inside a paragraph into a typed `Hyperlink`.
 * Either `url` (external) or `anchor` (internal bookmark) is required;
 * passing both, or neither, is `invalid-payload`. Run formatting inside
 * the wrapped span is preserved verbatim — Word's "Hyperlink" character
 * style is intentionally NOT auto-applied; that's a UI concern handled
 * separately.
 *
 * Steps:
 * 1. Validate payload (XOR on url/anchor, well-formed url, range bounds).
 * 2. Locate paragraph; reject `unknown-target` on miss.
 * 3. Reject `invalid-position` if the range straddles an existing
 *    hyperlink or any non-run inline (we only wrap pure-run spans this
 *    round; nested hyperlinks are not legal in OOXML).
 * 4. Split runs at the flat-text boundaries so the captured span is a
 *    contiguous run sequence.
 * 5. If `url` is set, mint (or de-dup) an `external` relationship in
 *    `word/_rels/document.xml.rels` and set `dirty.relationships`.
 * 6. Wrap the captured runs in a `Hyperlink` node and splice it back
 *    into the paragraph at the original position.
 *
 * Diff: emits a `node-inserted` for the hyperlink wrapper and a
 * `node-updated` for the host paragraph (the children array changed
 * shape). When a brand-new rel was minted, the diff stays single — the
 * rel mutation is reflected in `dirty.relationships`, not in a diff
 * change kind (we don't have a `rel-added` change today and the brief
 * doesn't ask for one).
 */
export const insertHyperlinkHandler: CommandHandler<InsertHyperlinkPayload, DocxSnapshot> = {
  type: "docx:insert-hyperlink",
  apply(snapshot, payload, ctx) {
    const { paragraphId, range, url, anchor } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    const hasUrl = typeof url === "string" && url.length > 0;
    const hasAnchor = typeof anchor === "string" && anchor.length > 0;
    if (hasUrl === hasAnchor) {
      throw new CommandError(
        "invalid-payload",
        "exactly one of `url` or `anchor` must be provided (mutually exclusive)"
      );
    }
    if (hasUrl) {
      try {
        new URL(url as string);
      } catch {
        throw new CommandError("invalid-payload", `url is not well-formed: ${url}`);
      }
    }
    if (!range || !Number.isInteger(range.start) || !Number.isInteger(range.end)) {
      throw new CommandError("invalid-payload", "range.start and range.end must be integers");
    }
    if (range.start < 0 || range.end <= range.start) {
      throw new CommandError(
        "invalid-payload",
        `range.start must be < range.end (got ${range.start}..${range.end})`
      );
    }

    const located = locateParagraph(snapshot.root, paragraphId);
    if (!located) {
      throw new CommandError("unknown-target", `no paragraph with id "${paragraphId}"`);
    }

    const totalLen = paragraphFlatTextLength(located.paragraph);
    if (range.end > totalLen) {
      throw new CommandError(
        "invalid-position",
        `range.end ${range.end} exceeds paragraph text length ${totalLen}`
      );
    }
    if (overlapsExistingHyperlink(located.paragraph, range.start, range.end)) {
      throw new CommandError(
        "invalid-position",
        "range straddles or overlaps an existing hyperlink (nested hyperlinks are not supported)"
      );
    }

    const wrapped = wrapRangeInHyperlink(located.paragraph, range.start, range.end, ctx.mintNodeId);
    if (!wrapped) {
      throw new CommandError(
        "invalid-position",
        "range crosses a non-run inline (comment marker, revision wrapper, or other) — not supported in this round"
      );
    }

    let relId: string | undefined;
    let nextRelationships = snapshot.root.relationships;
    let nextRelsDirty = snapshot.dirty.relationships;
    if (hasUrl) {
      const minted = mintOrReuseHyperlinkRel(snapshot.root.relationships, url as string);
      relId = minted.id;
      if (minted.added) {
        nextRelationships = new Map(snapshot.root.relationships).set(DOC_RELS_KEY, minted.nextRels);
        nextRelsDirty = withAddition(snapshot.dirty.relationships, DOC_RELS_KEY);
      }
    }

    const hyperlink: Hyperlink = {
      kind: "hyperlink",
      id: ctx.mintNodeId(),
      ...(relId ? { relationshipId: relId } : {}),
      ...(hasAnchor ? { anchor: anchor as string } : {}),
      children: wrapped.captured,
    };
    const newInlines = [...wrapped.before, hyperlink, ...wrapped.after];
    const newParagraph: Paragraph = { ...located.paragraph, children: newInlines };
    let nextDoc: DocxDocument = located.replace(newParagraph);
    if (nextRelationships !== snapshot.root.relationships) {
      nextDoc = { ...nextDoc, relationships: nextRelationships };
    }

    const next = evolveSnapshot(snapshot, nextDoc, {
      body: true,
      relationships: nextRelsDirty,
    });

    const summary = hasUrl ? `+hyperlink → ${url}` : `+hyperlink → #${anchor}`;
    return {
      next,
      diff: buildDiffMulti(snapshot.revision, next.revision, [
        {
          kind: "node-inserted",
          nodeId: hyperlink.id,
          path: located.path,
          summary,
        },
        {
          kind: "node-updated",
          nodeId: located.paragraph.id,
          path: located.path,
          field: "children",
          summary: `wrap ${range.end - range.start} chars in hyperlink`,
        },
      ]),
    };
  },
};

interface MintedRel {
  readonly id: string;
  readonly added: boolean;
  readonly nextRels: ReadonlyArray<Relationship>;
}

function mintOrReuseHyperlinkRel(
  relationships: ReadonlyMap<string, ReadonlyArray<Relationship>>,
  url: string
): MintedRel {
  const docRels = relationships.get(DOC_RELS_KEY) ?? [];
  const existing = docRels.find(
    (r) => r.type === HYPERLINK_REL_TYPE && r.target === url && r.targetMode === "External"
  );
  if (existing) {
    return { id: existing.id, added: false, nextRels: docRels };
  }
  const id = mintRelId(docRels);
  const next: Relationship = {
    id,
    type: HYPERLINK_REL_TYPE,
    target: url,
    targetMode: "External",
  };
  return { id, added: true, nextRels: [...docRels, next] };
}

function mintRelId(rels: ReadonlyArray<Relationship>): string {
  const taken = new Set(rels.map((r) => r.id));
  let i = rels.length + 1;
  while (taken.has(`rId${i}`)) i++;
  return `rId${i}`;
}

function withAddition(prev: ReadonlySet<string>, member: string): ReadonlySet<string> {
  const next = new Set(prev);
  next.add(member);
  return next;
}

function paragraphFlatTextLength(p: Paragraph): number {
  let n = 0;
  for (const inline of p.children) {
    n += inlineTextLength(inline);
  }
  return n;
}

function inlineTextLength(node: InlineNode): number {
  if (node.kind === "run") return runTextLength(node);
  if (node.kind === "hyperlink") {
    let n = 0;
    for (const r of node.children) n += runTextLength(r);
    return n;
  }
  if (node.kind === "revision") {
    let n = 0;
    for (const c of node.children) n += inlineTextLength(c);
    return n;
  }
  return 0;
}

function runTextLength(r: Run): number {
  let n = 0;
  for (const c of r.children) if (c.kind === "text") n += c.text.length;
  return n;
}

function overlapsExistingHyperlink(p: Paragraph, start: number, end: number): boolean {
  let cursor = 0;
  for (const inline of p.children) {
    const len = inlineTextLength(inline);
    const lo = cursor;
    const hi = cursor + len;
    cursor = hi;
    if (inline.kind !== "hyperlink") continue;
    if (hi <= start || lo >= end) continue;
    return true;
  }
  return false;
}

interface Wrapped {
  readonly before: ReadonlyArray<InlineNode>;
  readonly captured: ReadonlyArray<Run>;
  readonly after: ReadonlyArray<InlineNode>;
}

/**
 * Walk the paragraph's inlines, splitting any run that straddles the
 * range boundaries. Returns the inlines that come strictly before, the
 * runs to wrap, and the inlines that come strictly after. Returns
 * `null` if the range crosses a non-run inline (comment markers,
 * revision wrappers, opaque inlines) — those would change the
 * paragraph's structure if blindly absorbed into a hyperlink and we
 * keep them out of scope this round.
 */
function wrapRangeInHyperlink(
  p: Paragraph,
  start: number,
  end: number,
  mintNodeId: IdMinter
): Wrapped | null {
  const before: InlineNode[] = [];
  const captured: Run[] = [];
  const after: InlineNode[] = [];
  let cursor = 0;
  for (const inline of p.children) {
    const len = inlineTextLength(inline);
    const lo = cursor;
    const hi = cursor + len;
    cursor = hi;
    if (hi <= start) {
      before.push(inline);
      continue;
    }
    if (lo >= end) {
      after.push(inline);
      continue;
    }
    if (inline.kind !== "run") {
      // Non-run inline straddles or sits inside the range. Existing
      // hyperlinks were already rejected upstream; everything else
      // (comment markers, revisions, opaque inlines) is out of scope.
      return null;
    }
    const split = splitRunAtRange(inline, lo, start, end, mintNodeId);
    if (split.left) before.push(split.left);
    if (split.middle) captured.push(split.middle);
    if (split.right) after.push(split.right);
  }
  if (captured.length === 0) return null;
  return { before, captured, after };
}

interface RunSplit {
  readonly left: Run | null;
  readonly middle: Run | null;
  readonly right: Run | null;
}

function splitRunAtRange(
  run: Run,
  runStart: number,
  rangeStart: number,
  rangeEnd: number,
  mintNodeId: IdMinter
): RunSplit {
  const text = collectText(run);
  const runEnd = runStart + text.length;
  const lo = Math.max(rangeStart - runStart, 0);
  const hi = Math.min(rangeEnd - runStart, runEnd - runStart);
  const left = lo > 0 ? makeTextRun(run, text.slice(0, lo), mintNodeId) : null;
  const middle = hi > lo ? makeTextRun(run, text.slice(lo, hi), mintNodeId) : null;
  const right = hi < text.length ? makeTextRun(run, text.slice(hi), mintNodeId) : null;
  return { left, middle, right };
}

function collectText(r: Run): string {
  let out = "";
  for (const c of r.children) if (c.kind === "text") out += c.text;
  return out;
}

function makeTextRun(template: Run, text: string, mintNodeId: IdMinter): Run {
  const child: RunChild = {
    kind: "text",
    id: mintNodeId(),
    text,
    xmlSpacePreserve: /^\s|\s$/.test(text),
  };
  return {
    kind: "run",
    id: mintNodeId(),
    properties: { ...template.properties },
    children: [child],
  };
}
