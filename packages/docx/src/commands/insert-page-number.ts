import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type {
  BlockNode,
  DocxSnapshot,
  HeaderFooterPart,
  PageNumberFieldLeaf,
  Paragraph,
  Run,
  RunChild,
  TextLeaf,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { mergeHeaderFooterDirty } from "./set-header-text.js";
import type { InsertPageNumberPayload } from "./payloads.js";

/**
 * P3.4 / W16 — insert a typed `PageNumberFieldLeaf` into a header or
 * footer paragraph at a flat-text byte offset.
 *
 * The new field is wrapped in its own fresh {@link Run} (separate from
 * any surrounding text run) so the serializer can lift the run into a
 * `<w:fldSimple>` element without disturbing neighbouring runs. Run
 * properties are inherited from the run that contained the split
 * point so the surrounding font/size carries over.
 */
export const insertPageNumberHandler: CommandHandler<InsertPageNumberPayload, DocxSnapshot> = {
  type: "docx:insert-page-number",
  apply(snapshot, payload, ctx) {
    const { paragraphId, offset } = payload;
    if (!paragraphId) {
      throw new CommandError("invalid-payload", "paragraphId is required");
    }
    if (!Number.isInteger(offset) || offset < 0) {
      throw new CommandError("invalid-payload", `offset must be a non-negative integer (got ${offset})`);
    }
    const field = payload.field ?? "PAGE";
    if (field !== "PAGE" && field !== "NUMPAGES") {
      throw new CommandError("invalid-payload", `field must be "PAGE" or "NUMPAGES" (got ${field})`);
    }

    const located = locateParagraphInHeaderFooter(snapshot, paragraphId);
    if (!located) {
      throw new CommandError(
        "unknown-target",
        `no header/footer paragraph with id "${paragraphId}" — page numbers are only supported inside header/footer parts in P3.4`
      );
    }

    const { partIdx, part, paragraph, paragraphIdx } = located;
    const updatedParagraph = insertFieldIntoParagraph(paragraph, offset, field, ctx.mintNodeId);
    const newBody: BlockNode[] = part.body.slice();
    newBody[paragraphIdx] = updatedParagraph;
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
        nodeId: paragraph.id,
        path: ["headersAndFooters", partIdx, "body", paragraphIdx],
        field: "page-number",
        summary: `inserted ${field} field at offset ${offset} in ${part.kind} ${part.partPath}`,
      }),
    };
  },
};

interface LocatedHFParagraph {
  readonly partIdx: number;
  readonly part: HeaderFooterPart;
  readonly paragraph: Paragraph;
  readonly paragraphIdx: number;
}

function locateParagraphInHeaderFooter(snapshot: DocxSnapshot, paragraphId: string): LocatedHFParagraph | null {
  for (let i = 0; i < snapshot.root.headersAndFooters.length; i++) {
    const part = snapshot.root.headersAndFooters[i];
    for (let j = 0; j < part.body.length; j++) {
      const block = part.body[j];
      if (block.kind === "paragraph" && block.id === paragraphId) {
        return { partIdx: i, part, paragraph: block, paragraphIdx: j };
      }
    }
  }
  return null;
}

/**
 * Build the new paragraph children array with the field run spliced
 * in at the requested flat-text offset. Splits the targeting run on
 * a text-leaf boundary; any non-text run children (drawings, tabs,
 * page-breaks) before the split point keep their original placement
 * relative to the split.
 */
function insertFieldIntoParagraph(
  p: Paragraph,
  offset: number,
  field: "PAGE" | "NUMPAGES",
  mintNodeId: IdMinter
): Paragraph {
  const fieldLeaf: PageNumberFieldLeaf = {
    kind: "page-number-field",
    id: mintNodeId(),
    field,
    instr: ` ${field} \\* MERGEFORMAT `,
  };

  const flatLength = paragraphFlatLength(p);
  const clampedOffset = Math.max(0, Math.min(offset, flatLength));

  let consumed = 0;
  const newChildren = [];
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
      const { before, after } = splitRunAt(child, localOffset, mintNodeId);
      const fieldRun: Run = {
        kind: "run",
        id: mintNodeId(),
        properties: child.properties,
        children: [fieldLeaf],
      };
      if (before) newChildren.push(before);
      newChildren.push(fieldRun);
      if (after) newChildren.push(after);
      placed = true;
    } else {
      newChildren.push(child);
    }
    consumed += runLength;
  }

  if (!placed) {
    const fieldRun: Run = {
      kind: "run",
      id: mintNodeId(),
      properties: {},
      children: [fieldLeaf],
    };
    newChildren.push(fieldRun);
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

function splitRunAt(run: Run, localOffset: number, mintNodeId: IdMinter): SplitRun {
  if (localOffset <= 0) {
    const isEmpty = runFlatLength(run) === 0;
    return { before: isEmpty ? null : null, after: run };
  }
  const total = runFlatLength(run);
  if (localOffset >= total) {
    return { before: run, after: null };
  }
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
