import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, TextParagraph, TextRun, TextRunProperties, TextShape } from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  isTextShape,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { FormatTextPayload, TextFormatPayload } from "./payloads.js";

export const formatTextHandler: CommandHandler<FormatTextPayload, PptxSnapshot> = {
  type: "pptx:format-text",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTextShape(shape)) {
      throw makeError("not-applicable", `shape is not a text shape`);
    }
    const { range, format } = payload;
    if (
      !Number.isInteger(range.paragraph) ||
      range.paragraph < 0 ||
      range.paragraph >= shape.txBody.paragraphs.length
    ) {
      throw makeError("invalid-payload", `paragraph index out of range`);
    }
    const para = shape.txBody.paragraphs[range.paragraph];
    const flatLen = para.runs.reduce((acc, r) => acc + (r.isLineBreak ? 0 : r.text.length), 0);
    if (range.start < 0 || range.end < range.start || range.end > flatLen) {
      throw makeError("invalid-payload", `range out of bounds`);
    }
    if (range.start === range.end) {
      // No-op but valid; emit a no-change diff.
      return {
        next: { ...snapshot, revision: snapshot.revision + 1 },
        diff: buildDiff(snapshot.revision, snapshot.revision + 1, {
          kind: "node-updated",
          nodeId: shape.id,
          path: ["slides", sIdx, "shapes", ...path, "txBody", "paragraphs", range.paragraph],
          field: "format",
          summary: "no-op format-text",
        }),
      };
    }

    const newRuns = applyFormatToRuns(para.runs, range.start, range.end, format, ctx.mintNodeId);
    const coalesced = coalesce(newRuns);

    const updatedPara: TextParagraph = { ...para, runs: coalesced };
    const updatedShape: TextShape = {
      ...shape,
      txBody: {
        ...shape.txBody,
        paragraphs: shape.txBody.paragraphs.map((p, i) => (i === range.paragraph ? updatedPara : p)),
      },
    };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, updatedShape),
    }));

    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path, "txBody", "paragraphs", range.paragraph],
        field: "format",
        summary: summarize(format),
      }),
    };
  },
};

function applyFormatToRuns(
  runs: ReadonlyArray<TextRun>,
  start: number,
  end: number,
  format: TextFormatPayload,
  mint: () => string
): TextRun[] {
  const out: TextRun[] = [];
  let pos = 0;
  for (const r of runs) {
    if (r.isLineBreak) {
      out.push(r);
      continue;
    }
    const len = r.text.length;
    const runStart = pos;
    const runEnd = pos + len;
    pos = runEnd;
    if (runEnd <= start || runStart >= end) {
      out.push(r);
      continue;
    }

    const localStart = Math.max(0, start - runStart);
    const localEnd = Math.min(len, end - runStart);

    if (localStart > 0) {
      out.push({ ...r, id: mint(), text: r.text.slice(0, localStart) });
    }
    out.push({
      ...r,
      id: mint(),
      text: r.text.slice(localStart, localEnd),
      properties: mergeProps(r.properties, format),
    });
    if (localEnd < len) {
      out.push({ ...r, id: mint(), text: r.text.slice(localEnd) });
    }
  }
  return out;
}

function mergeProps(base: TextRunProperties, fmt: TextFormatPayload): TextRunProperties {
  const out: { -readonly [K in keyof TextRunProperties]: TextRunProperties[K] } = { ...base };
  if (fmt.bold !== undefined) out.bold = fmt.bold;
  if (fmt.italic !== undefined) out.italic = fmt.italic;
  if (fmt.underline !== undefined) out.underline = fmt.underline;
  if (fmt.strike !== undefined) out.strike = fmt.strike;
  if (fmt.fontFamily !== undefined) out.fontFamily = fmt.fontFamily;
  if (fmt.fontSizeHundredths !== undefined) out.fontSizeHundredths = fmt.fontSizeHundredths;
  if (fmt.color !== undefined) out.color = fmt.color;
  if (fmt.highlight !== undefined) out.highlight = fmt.highlight;
  return out;
}

function coalesce(runs: ReadonlyArray<TextRun>): TextRun[] {
  const out: TextRun[] = [];
  for (const r of runs) {
    const last = out[out.length - 1];
    if (last && !r.isLineBreak && !last.isLineBreak && propsEqual(last.properties, r.properties)) {
      out[out.length - 1] = { ...last, text: last.text + r.text };
    } else {
      out.push(r);
    }
  }
  return out;
}

function propsEqual(a: TextRunProperties, b: TextRunProperties): boolean {
  // Compare the JSON-serialised view; properties are simple.
  const stripIds = (p: TextRunProperties) => ({
    ...p,
    opaqueAttrs: p.opaqueAttrs,
    opaqueChildren: p.opaqueChildren,
  });
  return JSON.stringify(stripIds(a)) === JSON.stringify(stripIds(b));
}

function summarize(f: TextFormatPayload): string {
  const parts: string[] = [];
  for (const k of Object.keys(f) as Array<keyof TextFormatPayload>) {
    parts.push(`${k}=${String(f[k])}`);
  }
  return parts.join(",");
}
