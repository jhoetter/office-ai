import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Paragraph, Run } from "../model/types.js";
import { buildDiff, evolveSnapshot, insertBlock } from "./helpers.js";
import type { InsertParagraphPayload } from "./payloads.js";

export const insertParagraphHandler: CommandHandler<InsertParagraphPayload, DocxSnapshot> = {
  type: "docx:insert-paragraph",
  apply(snapshot, payload, ctx) {
    const { at, style } = payload;
    if (at.paragraph < 0 || at.paragraph > snapshot.root.body.length) {
      throw new CommandError("invalid-position", `paragraph index ${at.paragraph} out of range`);
    }
    const newRun: Run = {
      kind: "run",
      id: ctx.mintNodeId(),
      properties: {},
      children: [],
    };
    const newP: Paragraph = {
      kind: "paragraph",
      id: ctx.mintNodeId(),
      properties: style ? { styleId: style } : {},
      children: [newRun],
    };
    const nextDoc = insertBlock(snapshot.root, at.paragraph, newP);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newP.id,
        path: ["body", at.paragraph],
        summary: `+paragraph${style ? ` (style=${style})` : ""}`,
      }),
    };
  },
};
