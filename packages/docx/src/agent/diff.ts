/**
 * Structural diff between two DOCX snapshots.
 *
 * Produces a {@link DocumentDiff} (the format-agnostic core type), suitable
 * for piping into the same review surface that consumes per-command
 * mutation diffs. The implementation deliberately stays at the body
 * (paragraph) + comments + relationships granularity — finer-grained
 * change attribution is the CommandBus's job (one mutation per command).
 *
 * Spec contract: `spec/shared/agent-api.md` §getDiff and the
 * `DocumentAgent.getDiff` row in `prompt.md`.
 */

import type { DiffChange, DocumentDiff } from "@officeai/core";
import type { BlockNode, DocxComment, DocxSnapshot, Paragraph } from "../model/types.js";
import { paragraphPlainText } from "../commands/helpers.js";

interface ProjectedParagraph {
  readonly id: string;
  readonly index: number;
  readonly styleId?: string;
  readonly text: string;
}

/**
 * Walk the body and produce a flat paragraph projection. Non-paragraph
 * blocks (tables, opaque blocks, …) are intentionally skipped — diffing
 * them at this level would either hide changes inside their typed
 * subtrees or raise too many false positives. Cross-paragraph diffs are
 * what reviewers care about; richer per-table diffs come from the
 * mutation diffs emitted by the typed table commands.
 */
function projectParagraphs(snap: DocxSnapshot): ProjectedParagraph[] {
  const out: ProjectedParagraph[] = [];
  const body = snap.root.body;
  for (let i = 0; i < body.length; i++) {
    const block: BlockNode = body[i];
    if (block.kind !== "paragraph") continue;
    const p = block as Paragraph;
    out.push({
      id: p.id,
      index: i,
      ...(p.properties.styleId ? { styleId: p.properties.styleId } : {}),
      text: paragraphPlainText(p),
    });
  }
  return out;
}

/** Match paragraphs across snapshots by stable id, falling back to index. */
function matchParagraphs(
  before: ProjectedParagraph[],
  after: ProjectedParagraph[]
): {
  matched: Array<{ before: ProjectedParagraph; after: ProjectedParagraph }>;
  inserted: ProjectedParagraph[];
  removed: ProjectedParagraph[];
} {
  const beforeById = new Map(before.map((p) => [p.id, p]));
  const afterById = new Map(after.map((p) => [p.id, p]));
  const matched: Array<{ before: ProjectedParagraph; after: ProjectedParagraph }> = [];
  const inserted: ProjectedParagraph[] = [];
  const removed: ProjectedParagraph[] = [];

  for (const a of after) {
    const b = beforeById.get(a.id);
    if (b) {
      matched.push({ before: b, after: a });
    } else {
      inserted.push(a);
    }
  }
  for (const b of before) {
    if (!afterById.has(b.id)) removed.push(b);
  }
  return { matched, inserted, removed };
}

function commentText(c: DocxComment): string {
  return c.body
    .map((b) => (b.kind === "paragraph" ? paragraphPlainText(b as Paragraph) : ""))
    .join("\n")
    .trim();
}

/** Produce a structural diff between two DOCX snapshots. */
export function diffDocxSnapshots(before: DocxSnapshot, after: DocxSnapshot): DocumentDiff {
  const beforeP = projectParagraphs(before);
  const afterP = projectParagraphs(after);
  const { matched, inserted, removed } = matchParagraphs(beforeP, afterP);

  const changes: DiffChange[] = [];

  for (const m of matched) {
    if (m.before.text !== m.after.text) {
      changes.push({
        kind: "node-updated",
        nodeId: m.after.id,
        path: ["body", m.after.index],
        field: "text",
        summary: `paragraph[${m.after.index}] text: "${m.before.text.slice(0, 40)}" → "${m.after.text.slice(0, 40)}"`,
      });
    }
    if ((m.before.styleId ?? "") !== (m.after.styleId ?? "")) {
      changes.push({
        kind: "node-updated",
        nodeId: m.after.id,
        path: ["body", m.after.index],
        field: "styleId",
        summary: `paragraph[${m.after.index}] style: ${m.before.styleId ?? "(none)"} → ${m.after.styleId ?? "(none)"}`,
      });
    }
    if (m.before.index !== m.after.index) {
      changes.push({
        kind: "node-moved",
        nodeId: m.after.id,
        from: ["body", m.before.index],
        to: ["body", m.after.index],
        summary: `paragraph moved [${m.before.index}] → [${m.after.index}]`,
      });
    }
  }

  for (const a of inserted) {
    changes.push({
      kind: "node-inserted",
      nodeId: a.id,
      path: ["body", a.index],
      summary: `+paragraph[${a.index}] ${a.text.slice(0, 60)}`,
    });
  }
  for (const r of removed) {
    changes.push({
      kind: "node-deleted",
      nodeId: r.id,
      path: ["body", r.index],
      summary: `-paragraph[${r.index}] ${r.text.slice(0, 60)}`,
    });
  }

  // Comments
  const beforeComments = new Map(before.root.comments.map((c) => [c.id, c]));
  const afterComments = new Map(after.root.comments.map((c) => [c.id, c]));
  for (const [id, c] of afterComments) {
    const prev = beforeComments.get(id);
    if (!prev) {
      changes.push({
        kind: "node-inserted",
        nodeId: id,
        path: ["comments", id],
        summary: `+comment ${id} by ${c.author}: ${commentText(c).slice(0, 40)}`,
      });
      continue;
    }
    if ((prev.resolved === true) !== (c.resolved === true)) {
      changes.push({
        kind: "node-updated",
        nodeId: id,
        path: ["comments", id],
        field: "resolved",
        summary: c.resolved ? `comment ${id} resolved` : `comment ${id} reopened`,
      });
    }
    if (commentText(prev) !== commentText(c)) {
      changes.push({
        kind: "node-updated",
        nodeId: id,
        path: ["comments", id],
        field: "text",
        summary: `comment ${id} text changed`,
      });
    }
  }
  for (const [id, c] of beforeComments) {
    if (!afterComments.has(id)) {
      changes.push({
        kind: "node-deleted",
        nodeId: id,
        path: ["comments", id],
        summary: `-comment ${id} by ${c.author}`,
      });
    }
  }

  // OPC parts (e.g. media binaries added by docx:insert-image)
  const beforeParts = new Set(before.container.parts.keys());
  const afterParts = new Set(after.container.parts.keys());
  for (const p of afterParts) {
    if (!beforeParts.has(p)) {
      changes.push({
        kind: "part-added",
        path: [p],
        summary: `+part ${p}`,
      });
    }
  }
  for (const p of beforeParts) {
    if (!afterParts.has(p)) {
      changes.push({
        kind: "node-deleted",
        nodeId: p,
        path: ["parts", p],
        summary: `-part ${p}`,
      });
    }
  }

  return {
    format: "docx",
    fromRevision: before.revision,
    toRevision: after.revision,
    changes,
  };
}
