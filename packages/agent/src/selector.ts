/**
 * Selector grammar for the office-agent CLI.
 *
 *   paragraph:N
 *   paragraph:N/run:M
 *   paragraph:N/text:OFFSET
 *   paragraph:N/text:A..B
 *   paragraph:N..paragraph:M
 *
 * See spec/agent/cli.md.
 */

import type { DocxPosition, DocxSelection } from "@officeai/docx";

export interface ParagraphSelector {
  kind: "paragraph";
  position: DocxPosition;
}

export interface ParagraphRangeSelector {
  kind: "range";
  range: DocxSelection;
}

export type Selector = ParagraphSelector | ParagraphRangeSelector;

export class SelectorError extends Error {
  readonly code = "invalid-selector";
}

export function parseSelector(input: string): Selector {
  const trimmed = input.trim();
  if (!trimmed) throw new SelectorError("Empty selector");

  if (trimmed.includes("..")) {
    return parseRangeSelector(trimmed);
  }

  return { kind: "paragraph", position: parsePosition(trimmed) };
}

function parsePosition(input: string): DocxPosition {
  const segments = input.split("/");
  let paragraph = -1;
  let run: number | undefined;
  let offset: number | undefined;
  for (const seg of segments) {
    const [kind, raw] = seg.split(":", 2);
    if (raw === undefined) throw new SelectorError(`Selector segment missing value: ${seg}`);
    const n = Number(raw);
    if (!Number.isInteger(n) || n < 0) {
      throw new SelectorError(`Selector segment must be a non-negative integer: ${seg}`);
    }
    switch (kind) {
      case "paragraph":
        paragraph = n;
        break;
      case "run":
        run = n;
        break;
      case "text":
        offset = n;
        break;
      default:
        throw new SelectorError(`Unknown selector segment kind: ${kind}`);
    }
  }
  if (paragraph < 0) throw new SelectorError(`Selector must include "paragraph:N"; got "${input}"`);
  const out: DocxPosition = { paragraph };
  if (run !== undefined) (out as { run?: number }).run = run;
  if (offset !== undefined) (out as { offset?: number }).offset = offset;
  return out;
}

function parseRangeSelector(input: string): ParagraphRangeSelector {
  // Forms supported:
  //   "paragraph:N..paragraph:M"
  //   "paragraph:N/text:A..B"
  if (/\/text:\d+\.\.\d+$/.test(input)) {
    const [head, tail] = input.split("/text:");
    const [aStr, bStr] = tail.split("..", 2);
    const head1 = parsePosition(head);
    const a = Number(aStr);
    const b = Number(bStr);
    if (!Number.isInteger(a) || !Number.isInteger(b)) {
      throw new SelectorError(`text range must be two integers: ${input}`);
    }
    return {
      kind: "range",
      range: {
        start: {
          paragraph: head1.paragraph,
          ...(head1.run !== undefined ? { run: head1.run } : { run: 0 }),
          offset: a,
        },
        end: {
          paragraph: head1.paragraph,
          ...(head1.run !== undefined ? { run: head1.run } : { run: 0 }),
          offset: b,
        },
      },
    };
  }
  const [a, b] = input.split("..", 2);
  const start = parsePosition(a);
  const end = parsePosition(b);
  return { kind: "range", range: { start, end } };
}
