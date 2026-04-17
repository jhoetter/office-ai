import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import type {
  BlockNode,
  DocxDocument,
  DocxPosition,
  DocxSnapshot,
  Paragraph,
  Run,
  RunProperties,
} from "../model/types.js";
import { buildDiff, buildDiffMulti, evolveSnapshot, replaceBlock } from "./helpers.js";
import type { FormatRangePayload, TextFormat } from "./payloads.js";

/**
 * `docx:format-range` handler.
 *
 * Single-paragraph ranges split the affected runs at the boundaries and
 * apply the format to runs that fall fully within. Multi-paragraph
 * ranges (`start.paragraph !== end.paragraph`) walk the paragraph span:
 * the start paragraph is formatted from the start boundary to its end,
 * intermediate paragraphs are formatted entirely, and the end paragraph
 * is formatted from its beginning to the end boundary. Non-paragraph
 * blocks inside the span are skipped.
 */
export const formatRangeHandler: CommandHandler<FormatRangePayload, DocxSnapshot> = {
  type: "docx:format-range",
  apply(snapshot, payload, ctx) {
    const { range, format } = payload;
    const [startPos, endPos] = orderPositions(range.start, range.end);

    if (startPos.paragraph === endPos.paragraph) {
      return applySingleParagraph(snapshot, startPos, endPos, format, ctx.mintNodeId);
    }
    return applyMultiParagraph(snapshot, startPos, endPos, format, ctx.mintNodeId);
  },
};

function orderPositions(a: DocxPosition, b: DocxPosition): [DocxPosition, DocxPosition] {
  if (a.paragraph < b.paragraph) return [a, b];
  if (a.paragraph > b.paragraph) return [b, a];
  // Same paragraph — fall back to offsets only when a run is given on both
  // sides; otherwise leave order alone (the handler normalizes via min/max).
  return [a, b];
}

function applySingleParagraph(
  snapshot: DocxSnapshot,
  startPos: DocxPosition,
  endPos: DocxPosition,
  format: TextFormat,
  mintNodeId: () => string
) {
  const idx = startPos.paragraph;
  const block = snapshot.root.body[idx];
  if (!block || block.kind !== "paragraph") {
    throw new CommandError("invalid-position", `paragraph index ${idx} is not a paragraph`);
  }
  const lo = paragraphTextOffset(block, startPos.run, startPos.offset ?? 0);
  const hi = paragraphTextOffset(block, endPos.run, endPos.offset ?? 0);
  const a = Math.min(lo, hi);
  const b = Math.max(lo, hi);
  if (a === b) {
    return {
      next: { ...snapshot, revision: snapshot.revision + 1 },
      diff: buildDiff(snapshot.revision, snapshot.revision + 1, {
        kind: "node-updated",
        nodeId: block.id,
        path: ["body", idx],
        field: "noop",
        summary: "no-op format-range",
      }),
    };
  }
  const updated = formatWithinParagraph(block, a, b, format, mintNodeId);
  const nextDoc = replaceBlock(snapshot.root, idx, updated);
  const next = evolveSnapshot(snapshot, nextDoc, { body: true });
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-updated",
      nodeId: block.id,
      path: ["body", idx],
      field: "format",
      summary: summarizeFormat(format),
    }),
  };
}

function applyMultiParagraph(
  snapshot: DocxSnapshot,
  startPos: DocxPosition,
  endPos: DocxPosition,
  format: TextFormat,
  mintNodeId: () => string
) {
  const startIdx = startPos.paragraph;
  const endIdx = endPos.paragraph;
  if (startIdx < 0 || endIdx >= snapshot.root.body.length) {
    throw new CommandError("invalid-position", `paragraph indices ${startIdx}..${endIdx} out of range`);
  }

  let doc: DocxDocument = snapshot.root;
  const changes: DiffChange[] = [];
  const formatSummary = summarizeFormat(format);

  for (let i = startIdx; i <= endIdx; i++) {
    const block: BlockNode = doc.body[i];
    if (!block || block.kind !== "paragraph") {
      // Skip non-paragraph blocks (tables / opaque blocks / section breaks).
      continue;
    }
    const totalLen = paragraphTextLength(block);
    let lo: number;
    let hi: number;
    if (i === startIdx) {
      lo = paragraphTextOffset(block, startPos.run, startPos.offset ?? 0);
      hi = totalLen;
    } else if (i === endIdx) {
      lo = 0;
      hi = paragraphTextOffset(block, endPos.run, endPos.offset ?? 0);
    } else {
      lo = 0;
      hi = totalLen;
    }
    if (lo >= hi) continue;
    const updated = formatWithinParagraph(block, lo, hi, format, mintNodeId);
    doc = replaceBlock(doc, i, updated);
    changes.push({
      kind: "node-updated",
      nodeId: block.id,
      path: ["body", i],
      field: "format",
      summary: formatSummary,
    });
  }

  const next = evolveSnapshot(snapshot, doc, { body: true });
  if (changes.length === 0) {
    // The selection landed entirely on non-paragraph blocks. We still bump
    // the revision so the bus emits a mutation, but flag it as a no-op.
    return {
      next,
      diff: buildDiffMulti(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["body"],
          field: "noop",
          summary: "no-op multi-paragraph format-range",
        },
      ]),
    };
  }
  return {
    next,
    diff: buildDiffMulti(snapshot.revision, next.revision, changes),
  };
}

function summarizeFormat(f: TextFormat): string {
  const parts: string[] = [];
  if (f.bold !== undefined) parts.push(`bold=${f.bold}`);
  if (f.italic !== undefined) parts.push(`italic=${f.italic}`);
  if (f.underline !== undefined) parts.push(`underline=${f.underline}`);
  if (f.strike !== undefined) parts.push(`strike=${f.strike}`);
  if (f.color) parts.push(`color=${f.color}`);
  if (f.fontSize !== undefined) parts.push(`size=${f.fontSize}`);
  if (f.fontFamily) parts.push(`font=${f.fontFamily}`);
  if (f.highlight) parts.push(`highlight=${f.highlight}`);
  return parts.join(", ");
}

/**
 * Convert a `(runIndex, localOffset)` pair into a paragraph-wide character
 * offset. When `runIndex` is undefined we treat `localOffset` as already
 * paragraph-wide and clamp it to the paragraph's text length so callers
 * (notably the multi-paragraph driver) can pass paragraph-wide offsets
 * directly.
 */
function paragraphTextOffset(p: Paragraph, runIndex: number | undefined, localOffset: number): number {
  if (runIndex === undefined) {
    return Math.max(0, Math.min(localOffset, paragraphTextLength(p)));
  }
  let offset = 0;
  let i = 0;
  for (const child of p.children) {
    if (child.kind !== "run") {
      i++;
      continue;
    }
    if (i === runIndex) {
      let consumed = 0;
      for (const c of child.children) {
        if (c.kind !== "text") continue;
        if (localOffset <= consumed + c.text.length) {
          return offset + (localOffset - consumed);
        }
        offset += c.text.length;
        consumed += c.text.length;
      }
      return offset + Math.max(0, localOffset - consumed);
    }
    for (const c of child.children) {
      if (c.kind === "text") offset += c.text.length;
    }
    i++;
  }
  return offset;
}

function paragraphTextLength(p: Paragraph): number {
  let n = 0;
  for (const child of p.children) {
    if (child.kind !== "run") continue;
    for (const c of child.children) {
      if (c.kind === "text") n += c.text.length;
    }
  }
  return n;
}

function formatWithinParagraph(
  p: Paragraph,
  a: number,
  b: number,
  format: TextFormat,
  mintNodeId: () => string
): Paragraph {
  let cursor = 0;
  const out: Paragraph["children"][number][] = [];
  for (const child of p.children) {
    if (child.kind !== "run") {
      out.push(child);
      continue;
    }
    const start = cursor;
    const end = cursor + runTextLength(child);
    if (b <= start || a >= end) {
      out.push(child);
    } else if (a <= start && b >= end) {
      out.push(applyFormatToRun(child, format));
    } else {
      const split = splitRun(child, Math.max(0, a - start), Math.min(end - start, b - start), mintNodeId);
      if (split.left) out.push(split.left);
      out.push(applyFormatToRun(split.middle, format));
      if (split.right) out.push(split.right);
    }
    cursor = end;
  }
  return { ...p, children: out };
}

function runTextLength(r: Run): number {
  let n = 0;
  for (const c of r.children) if (c.kind === "text") n += c.text.length;
  return n;
}

function applyFormatToRun(r: Run, f: TextFormat): Run {
  type Mut = { -readonly [K in keyof RunProperties]: RunProperties[K] };
  const props: Mut = { ...r.properties };
  if (f.bold !== undefined) props.bold = f.bold;
  if (f.italic !== undefined) props.italic = f.italic;
  if (f.underline !== undefined) props.underline = f.underline;
  if (f.strike !== undefined) props.strike = f.strike;
  if (f.fontFamily !== undefined) props.fontFamily = f.fontFamily;
  if (f.fontSize !== undefined) props.fontSize = f.fontSize;
  if (f.color !== undefined) props.color = f.color;
  if (f.highlight !== undefined) props.highlight = f.highlight;
  return { ...r, properties: props };
}

interface SplitResult {
  left: Run | null;
  middle: Run;
  right: Run | null;
}

function splitRun(r: Run, lo: number, hi: number, mintNodeId: () => string): SplitResult {
  const text = collectText(r);
  const before = text.slice(0, lo);
  const middle = text.slice(lo, hi);
  const after = text.slice(hi);
  const left: Run | null = before.length > 0 ? makeTextRun(r, before, mintNodeId) : null;
  const middleRun: Run = makeTextRun(r, middle, mintNodeId);
  const right: Run | null = after.length > 0 ? makeTextRun(r, after, mintNodeId) : null;
  return { left, middle: middleRun, right };
}

function collectText(r: Run): string {
  let out = "";
  for (const c of r.children) if (c.kind === "text") out += c.text;
  return out;
}

function makeTextRun(template: Run, text: string, mintNodeId: () => string): Run {
  return {
    kind: "run",
    id: mintNodeId(),
    properties: { ...template.properties },
    children: [
      {
        kind: "text",
        id: mintNodeId(),
        text,
        xmlSpacePreserve: /^\s|\s$/.test(text),
      },
    ],
  };
}
