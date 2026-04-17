import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, withParagraph } from "./helpers.js";
import type { SetParagraphStylePayload } from "./payloads.js";

export const setParagraphStyleHandler: CommandHandler<SetParagraphStylePayload, DocxSnapshot> = {
  type: "docx:set-paragraph-style",
  apply(snapshot, payload) {
    const { at, style } = payload;
    if (at.paragraph < 0 || at.paragraph >= snapshot.root.body.length) {
      throw new CommandError("invalid-position", `paragraph index ${at.paragraph} out of range`);
    }
    const block = snapshot.root.body[at.paragraph];
    if (block.kind !== "paragraph") {
      throw new CommandError("not-paragraph", `block at ${at.paragraph} is not a paragraph`);
    }
    const previous = block.properties.styleId;
    const nextDoc = withParagraph(snapshot.root, at.paragraph, (p) => ({
      ...p,
      properties: { ...p.properties, styleId: style },
    }));
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: block.id,
        path: ["body", at.paragraph, "properties"],
        field: "styleId",
        summary: `${previous ?? "(none)"} → ${style}`,
      }),
    };
  },
};
