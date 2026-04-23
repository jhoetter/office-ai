import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, Slide } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide } from "./helpers.js";
import type { SetSlideHiddenPayload } from "./payloads.js";

/**
 * Hide / unhide a slide. Mirrors PowerPoint's "Hide Slide" toggle on
 * the Slide Show ribbon: writes `show="0"` (or removes it) on
 * `<p:sld>`. Hidden slides remain part of the deck and the editor
 * still renders them in the slide tray, but presentation mode skips
 * them.
 *
 * The flag lives on `slideRootAttrs` (parsed verbatim) so we don't
 * need a typed model field — round-tripping is already handled by the
 * generic root-attrs path.
 */
export const setSlideHiddenHandler: CommandHandler<SetSlideHiddenPayload, PptxSnapshot> = {
  type: "pptx:set-slide-hidden",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const nextAttrs = nextRootAttrs(slide.slideRootAttrs, payload.hidden);

    // No-op short-circuit: same flag, identical attrs object.
    if (nextAttrs === slide.slideRootAttrs) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision) };
    }

    const nextSlide: Slide = { ...slide, slideRootAttrs: nextAttrs };
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
        path: ["slides", sIdx, "hidden"],
        field: "hidden",
        summary: payload.hidden ? "true" : "false",
      }),
    };
  },
};

function nextRootAttrs(
  current: Readonly<Record<string, string>>,
  hidden: boolean
): Readonly<Record<string, string>> {
  const has = current.show === "0";
  if (hidden && has) return current;
  if (!hidden && !("show" in current)) return current;

  const next = { ...current };
  if (hidden) {
    next.show = "0";
  } else {
    delete next.show;
  }
  return next;
}
