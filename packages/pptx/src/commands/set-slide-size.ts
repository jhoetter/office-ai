import type { CommandHandler } from "@officeai/core";
import type { PptxPresentation, PptxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetSlideSizePayload } from "./payloads.js";

/**
 * Resize the deck via PowerPoint's "Slide Size" picker. Mirrors the
 * `<p:sldSz cx="..." cy="..." type="..."/>` element on
 * `ppt/presentation.xml`.
 *
 * Three input shapes are accepted to match the desktop app's UX:
 *   • `preset: "widescreen"`  → 16:9, 13.333" × 7.5"
 *   • `preset: "standard"`    → 4:3, 10"     × 7.5"
 *   • `preset: "a4"`          → A4 paper (≈ 11.69" × 8.27")
 *   • `preset: "letter"`      → US Letter   (11"     × 8.5")
 *   • `preset: "custom"`      → caller supplies cx/cy in EMU directly
 *
 * The serializer (`serializePresentationXml`) already pulls slide size
 * from `root.slideSize`, so this handler simply mutates the typed
 * field; no opaque-tail rewriting is needed. Existing shapes keep
 * their EMU coordinates verbatim — agents should re-layout shapes
 * after a size change if the visual fit matters.
 */
export const setSlideSizeHandler: CommandHandler<SetSlideSizePayload, PptxSnapshot> = {
  type: "pptx:set-slide-size",
  apply(snapshot, payload) {
    const next = resolveSize(payload);
    const current = snapshot.root.slideSize;
    if (
      next.cxEmu === current.cxEmu &&
      next.cyEmu === current.cyEmu &&
      (next.type ?? undefined) === current.type
    ) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision) };
    }

    const root: PptxPresentation = { ...snapshot.root, slideSize: next };
    const evolved = evolveSnapshot(snapshot, root, { presentation: true });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: snapshot.root.id,
        path: ["presentation", "slideSize"],
        field: "slideSize",
        summary: `${next.cxEmu}×${next.cyEmu}${next.type ? ` (${next.type})` : ""}`,
      }),
    };
  },
};

function resolveSize(payload: SetSlideSizePayload): {
  cxEmu: number;
  cyEmu: number;
  type?: string;
} {
  if (payload.preset && payload.preset !== "custom") {
    const preset = SLIDE_SIZE_PRESETS[payload.preset];
    return { ...preset };
  }
  if (payload.cxEmu === undefined || payload.cyEmu === undefined) {
    throw new Error("set-slide-size: cxEmu and cyEmu are required for preset 'custom'");
  }
  if (payload.cxEmu <= 0 || payload.cyEmu <= 0) {
    throw new Error("set-slide-size: cxEmu/cyEmu must be positive EMU values");
  }
  return {
    cxEmu: payload.cxEmu,
    cyEmu: payload.cyEmu,
    ...(payload.sizeType ? { type: payload.sizeType } : {}),
  };
}

export const SLIDE_SIZE_PRESETS = {
  widescreen: { cxEmu: 12_192_000, cyEmu: 6_858_000, type: "screen16x9" as const },
  standard: { cxEmu: 9_144_000, cyEmu: 6_858_000, type: "screen4x3" as const },
  a4: { cxEmu: 10_692_000, cyEmu: 7_560_000, type: "A4" as const },
  letter: { cxEmu: 10_058_400, cyEmu: 7_772_400, type: "letter" as const },
} as const;

export type SlideSizePresetKey = keyof typeof SLIDE_SIZE_PRESETS;
