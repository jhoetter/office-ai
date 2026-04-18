import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, makeError } from "./helpers.js";
import type { MoveSlidePayload } from "./payloads.js";

export const moveSlideHandler: CommandHandler<MoveSlidePayload, PptxSnapshot> = {
  type: "pptx:move-slide",
  apply(snapshot, payload) {
    const len = snapshot.root.slides.length;
    if (payload.from < 0 || payload.from >= len) {
      throw makeError("invalid-position", `from ${payload.from} out of range`);
    }
    if (payload.to < 0 || payload.to >= len) {
      throw makeError("invalid-position", `to ${payload.to} out of range`);
    }

    const newSlides = [...snapshot.root.slides];
    const [moved] = newSlides.splice(payload.from, 1);
    newSlides.splice(payload.to, 0, moved);

    const root = { ...snapshot.root, slides: newSlides };
    const next = evolveSnapshot(snapshot, root, { presentation: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-moved",
        nodeId: moved.id,
        from: ["slides", payload.from],
        to: ["slides", payload.to],
        summary: "slide",
      }),
    };
  },
};
