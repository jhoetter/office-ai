import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Paragraph, Run, RunProperties } from "../model/types.js";
import { buildDiff, evolveSnapshot, withParagraph } from "./helpers.js";
import type { FormatRangePayload, TextFormat } from "./payloads.js";

/**
 * Single-paragraph format-range. Splits the affected runs at range boundaries
 * and applies the format to runs that fall fully within. Multi-paragraph
 * formatting is P1.
 */
export const formatRangeHandler: CommandHandler<FormatRangePayload, DocxSnapshot> = {
  type: "docx:format-range",
  apply(snapshot, payload, ctx) {
    const { range, format } = payload;
    if (range.start.paragraph !== range.end.paragraph) {
      throw new CommandError(
        "multi-paragraph-format",
        "Multi-paragraph format-range is P1; apply per paragraph."
      );
    }
    const idx = range.start.paragraph;
    const block = snapshot.root.body[idx];
    if (!block || block.kind !== "paragraph") {
      throw new CommandError("invalid-position", `paragraph index ${idx} is not a paragraph`);
    }
    const lo = paragraphTextOffset(block, range.start.run, range.start.offset ?? 0);
    const hi = paragraphTextOffset(block, range.end.run, range.end.offset ?? 0);
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

    const nextDoc = withParagraph(snapshot.root, idx, (p) =>
      formatWithinParagraph(p, a, b, format, ctx.mintNodeId)
    );
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
  },
};

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

function paragraphTextOffset(p: Paragraph, runIndex: number | undefined, localOffset: number): number {
  let offset = 0;
  let i = 0;
  for (const child of p.children) {
    if (child.kind !== "run") {
      i++;
      continue;
    }
    if (runIndex !== undefined && i === runIndex) {
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

function formatWithinParagraph(
  p: Paragraph,
  a: number,
  b: number,
  format: TextFormat,
  mintNodeId: () => string
): Paragraph {
  let cursor = 0;
  const out = [];
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
      // Entirely within range — apply directly.
      out.push(applyFormatToRun(child, format));
    } else {
      // Partial overlap — split the run at boundaries.
      const split = splitRun(child, Math.max(0, a - start), Math.min(end - start, b - start), mintNodeId);
      out.push(split.left ?? null, applyFormatToRun(split.middle, format), split.right ?? null);
    }
    cursor = end;
  }
  return { ...p, children: out.filter((c): c is NonNullable<typeof c> => c !== null) };
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
