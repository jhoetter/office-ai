import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, InlineNode, Paragraph, Run, RunChild, TextLeaf } from "../model/types.js";
import { buildDiff, emptyRun, evolveSnapshot, textLeaf, withParagraph } from "./helpers.js";
import type { InsertTextPayload } from "./payloads.js";

export const insertTextHandler: CommandHandler<InsertTextPayload, DocxSnapshot> = {
  type: "docx:insert-text",
  apply(snapshot, payload, ctx) {
    const { at, text } = payload;
    if (!text) {
      // Nothing to do; emit an empty diff but still bump revision.
      return {
        next: { ...snapshot, revision: snapshot.revision + 1 },
        diff: buildDiff(snapshot.revision, snapshot.revision + 1, {
          kind: "node-updated",
          nodeId: snapshot.root.id,
          path: ["body"],
          field: "noop",
          summary: "no-op insert-text",
        }),
      };
    }
    if (at.paragraph < 0 || at.paragraph >= snapshot.root.body.length) {
      throw new CommandError("invalid-position", `paragraph index ${at.paragraph} out of range`);
    }
    const targetBlock = snapshot.root.body[at.paragraph];
    if (targetBlock.kind !== "paragraph") {
      throw new CommandError("not-paragraph", `block at ${at.paragraph} is not a paragraph`);
    }

    const nextDoc = withParagraph(snapshot.root, at.paragraph, (p) =>
      insertTextIntoParagraph(p, at.run, at.offset ?? 0, text, ctx.mintNodeId)
    );

    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: targetBlock.id,
        path: ["body", at.paragraph],
        field: "text",
        summary: `+${JSON.stringify(text)}`,
      }),
    };
  },
};

function insertTextIntoParagraph(
  p: Paragraph,
  runIndex: number | undefined,
  offset: number,
  text: string,
  mintNodeId: () => string
): Paragraph {
  if (runIndex === undefined) {
    // Paragraph-global offset: walk children and pick the right run
    // to splice into. If the offset lands inside an opaque container
    // (existing revision wrapper, hyperlink, etc.), splice a fresh
    // run in front of it (we don't crack open opaque containers from
    // the plain-edit path). If it lands at a child boundary, we
    // splice a new run there. If it lands inside a run, we extend
    // that run's text leaves.
    return spliceTextAtParagraphOffset(p, offset, text, mintNodeId);
  }

  const target = p.children[runIndex];
  if (!target || target.kind !== "run") {
    // Fall back: prepend.
    const newRun: Run = {
      kind: "run",
      id: mintNodeId(),
      properties: {},
      children: [textLeaf(mintNodeId, text)],
    };
    const children = [...p.children];
    children.splice(runIndex, 0, newRun);
    return { ...p, children };
  }

  const updatedRun = insertTextIntoRun(target, offset, text, mintNodeId);
  const children = [...p.children];
  children[runIndex] = updatedRun;
  return { ...p, children };
}

function spliceTextAtParagraphOffset(
  p: Paragraph,
  offset: number,
  text: string,
  mintNodeId: () => string
): Paragraph {
  if (p.children.length === 0 || offset <= 0) {
    const newRun: Run = {
      kind: "run",
      id: mintNodeId(),
      properties: {},
      children: [textLeaf(mintNodeId, text)],
    };
    return { ...p, children: [newRun, ...p.children] };
  }
  let consumed = 0;
  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    const childLen = inlineTextLength(child);
    const childEnd = consumed + childLen;
    if (offset === consumed) {
      const newRun: Run = {
        kind: "run",
        id: mintNodeId(),
        properties: {},
        children: [textLeaf(mintNodeId, text)],
      };
      const children = [...p.children];
      children.splice(i, 0, newRun);
      return { ...p, children };
    }
    if (offset > consumed && offset <= childEnd) {
      if (child.kind === "run") {
        const updated = insertTextIntoRun(child, offset - consumed, text, mintNodeId);
        const children = [...p.children];
        children[i] = updated;
        return { ...p, children };
      }
      // Opaque container — splice in front (or after if at the end of it).
      const newRun: Run = {
        kind: "run",
        id: mintNodeId(),
        properties: {},
        children: [textLeaf(mintNodeId, text)],
      };
      const children = [...p.children];
      const insertAt = offset === childEnd ? i + 1 : i;
      children.splice(insertAt, 0, newRun);
      return { ...p, children };
    }
    consumed = childEnd;
  }
  // Beyond all children → append.
  const newRun: Run = {
    kind: "run",
    id: mintNodeId(),
    properties: {},
    children: [textLeaf(mintNodeId, text)],
  };
  return { ...p, children: [...p.children, newRun] };
}

function inlineTextLength(node: InlineNode): number {
  if (node.kind === "run") {
    let n = 0;
    for (const c of node.children) if (c.kind === "text") n += c.text.length;
    return n;
  }
  if (node.kind === "revision") {
    let n = 0;
    for (const c of node.children) n += inlineTextLength(c);
    return n;
  }
  if (node.kind === "hyperlink") {
    let n = 0;
    for (const c of node.children) n += inlineTextLength(c);
    return n;
  }
  return 0;
}

function insertTextIntoRun(run: Run, offset: number, text: string, mintNodeId: () => string): Run {
  let consumed = 0;
  const newChildren: RunChild[] = [];
  let placed = false;

  for (const child of run.children) {
    if (placed || child.kind !== "text") {
      newChildren.push(child);
      continue;
    }
    const len = child.text.length;
    if (offset >= consumed && offset <= consumed + len) {
      const localOffset = offset - consumed;
      const before = child.text.slice(0, localOffset);
      const after = child.text.slice(localOffset);
      if (before.length === 0) {
        newChildren.push(textLeaf(mintNodeId, text));
        newChildren.push({ ...child, text: after });
      } else if (after.length === 0) {
        newChildren.push({ ...child, text: before });
        newChildren.push(textLeaf(mintNodeId, text));
      } else {
        const beforeLeaf: TextLeaf = { ...child, text: before };
        const middleLeaf = textLeaf(mintNodeId, text);
        const afterLeaf: TextLeaf = { ...child, id: mintNodeId(), text: after };
        newChildren.push(beforeLeaf, middleLeaf, afterLeaf);
      }
      placed = true;
    } else {
      newChildren.push(child);
    }
    consumed += len;
  }

  if (!placed) {
    if (run.children.length === 0) {
      newChildren.push(textLeaf(mintNodeId, text));
    } else {
      newChildren.push(textLeaf(mintNodeId, text));
    }
  }
  // If the run was empty before and we want to use emptyRun helper for typing
  void emptyRun;
  return { ...run, children: newChildren };
}
