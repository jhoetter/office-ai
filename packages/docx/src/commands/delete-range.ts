import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Paragraph, Run, RunChild, TextLeaf } from "../model/types.js";
import { buildDiff, evolveSnapshot, withParagraph } from "./helpers.js";
import type { DeleteRangePayload } from "./payloads.js";

/**
 * Minimal delete-range: supports deletion within a single paragraph (most
 * common case from the agent). Multi-paragraph deletion is a P1 enhancement
 * — see docs/build-log/docx.md.
 */
export const deleteRangeHandler: CommandHandler<DeleteRangePayload, DocxSnapshot> = {
  type: "docx:delete-range",
  apply(snapshot, payload) {
    const { range } = payload;
    if (range.start.paragraph !== range.end.paragraph) {
      throw new CommandError(
        "multi-paragraph-delete",
        "Multi-paragraph delete is P1; collapse to a single paragraph first.",
      );
    }
    const idx = range.start.paragraph;
    if (idx < 0 || idx >= snapshot.root.body.length) {
      throw new CommandError("invalid-position", `paragraph index ${idx} out of range`);
    }
    const block = snapshot.root.body[idx];
    if (block.kind !== "paragraph") {
      throw new CommandError("not-paragraph", `block at ${idx} is not a paragraph`);
    }
    const startOff = paragraphTextOffset(block, range.start.run, range.start.offset ?? 0);
    const endOff = paragraphTextOffset(block, range.end.run, range.end.offset ?? 0);
    const lo = Math.min(startOff, endOff);
    const hi = Math.max(startOff, endOff);
    if (lo === hi) {
      return {
        next: { ...snapshot, revision: snapshot.revision + 1 },
        diff: buildDiff(snapshot.revision, snapshot.revision + 1, {
          kind: "node-updated",
          nodeId: block.id,
          path: ["body", idx],
          field: "noop",
          summary: "no-op delete-range",
        }),
      };
    }

    const nextDoc = withParagraph(snapshot.root, idx, (p) => deleteWithinParagraph(p, lo, hi));
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: block.id,
        path: ["body", idx],
        field: "text",
        summary: `−${hi - lo} chars`,
      }),
    };
  },
};

/** Convert a (runIndex, localOffset) into a paragraph-wide text offset. */
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
  if (runIndex === undefined) return localOffset;
  return offset;
}

function deleteWithinParagraph(p: Paragraph, lo: number, hi: number): Paragraph {
  const newChildren = p.children.map((c) => {
    if (c.kind !== "run") return c;
    return deleteWithinRun(c, lo, hi)[0];
  });
  // Walk paragraph-wide and adjust ranges as we move along.
  let cursor = 0;
  const finalChildren = p.children.map((c) => {
    if (c.kind !== "run") return c;
    const before = cursor;
    const after = before + runTextLength(c);
    if (hi <= before || lo >= after) {
      cursor = after;
      return c;
    }
    const localLo = Math.max(0, lo - before);
    const localHi = Math.min(after - before, hi - before);
    const [updated] = deleteWithinRun(c, localLo, localHi);
    cursor = after;
    return updated;
  });
  void newChildren;
  return { ...p, children: finalChildren };
}

function runTextLength(r: Run): number {
  let n = 0;
  for (const c of r.children) if (c.kind === "text") n += c.text.length;
  return n;
}

function deleteWithinRun(r: Run, lo: number, hi: number): [Run, number] {
  let consumed = 0;
  const out: RunChild[] = [];
  for (const child of r.children) {
    if (child.kind !== "text") {
      out.push(child);
      continue;
    }
    const len = child.text.length;
    const start = consumed;
    const end = consumed + len;
    if (hi <= start || lo >= end) {
      out.push(child);
    } else {
      const lLo = Math.max(0, lo - start);
      const lHi = Math.min(len, hi - start);
      const before = child.text.slice(0, lLo);
      const after = child.text.slice(lHi);
      const merged = before + after;
      const updated: TextLeaf = { ...child, text: merged };
      out.push(updated);
    }
    consumed = end;
  }
  return [{ ...r, children: out }, hi - lo];
}
