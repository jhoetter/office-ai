import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, Slide } from "../model/types.js";
import { normaliseFillSpec, spliceSlideBackground, type FillSpec } from "../model/fill.js";
import { buildDiff, evolveSnapshot, findSlide, makeError } from "./helpers.js";
import type { SetSlideBackgroundPayload } from "./payloads.js";

/**
 * Replace (or clear) the slide-level `<p:bg>` block. Mirrors PowerPoint's
 * "Format Background" → "Apply to current slide". Any FillSpec is
 * accepted (solid colour, gradient, pattern, picture); pass `null` to
 * remove the background entirely so the slide inherits its layout/master
 * background instead.
 *
 * The handler mutates the slide's `cSldHead` opaque tree directly (the
 * parser captures every pre-`<p:spTree>` `<p:cSld>` child verbatim,
 * which is where `<p:bg>` lives), then marks the slide part dirty so
 * the serializer round-trips the change.
 */
export const setSlideBackgroundHandler: CommandHandler<SetSlideBackgroundPayload, PptxSnapshot> = {
  type: "pptx:set-slide-background",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);

    const spec = coerceFill(payload.fill);
    const nextSlide: Slide = { ...slide, cSldHead: spliceSlideBackground(slide.cSldHead, spec) };

    const root: PptxSnapshot["root"] = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? nextSlide : s)),
    };
    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: slide.id,
        path: ["slides", sIdx, "background"],
        field: "background",
        summary: summarise(spec),
      }),
    };
  },
};

function coerceFill(input: FillSpec | null): FillSpec | null {
  if (input === null) return null;
  try {
    return normaliseFillSpec(input);
  } catch (err) {
    throw makeError("invalid-payload", err instanceof Error ? err.message : String(err));
  }
}

function summarise(spec: FillSpec | null): string {
  if (spec === null) return "(reset)";
  switch (spec.type) {
    case "none":
      return "(none)";
    case "solid":
      return `#${spec.color}`;
    case "gradient":
      return `gradient(${spec.kind})`;
    case "pattern":
      return `pattern(${spec.preset})`;
    case "picture":
      return `picture(${spec.embedRelId})`;
    default: {
      const _exhaustive: never = spec;
      void _exhaustive;
      return "(unknown)";
    }
  }
}
