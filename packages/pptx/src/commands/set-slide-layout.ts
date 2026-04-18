/**
 * `pptx:set-slide-layout` — swap the layout a slide points to and
 * (by default) re-stamp its placeholders. Existing user content is
 * preserved by `clonePlaceholdersIntoSlide` when the placeholder idx
 * matches.
 */

import type { CommandHandler } from "@officeai/core";
import type { PptxPresentation, PptxSnapshot, SlideLayout } from "../model/types.js";
import {
  applyAddedLayout,
  clonePlaceholdersIntoSlide,
  resolveLayoutForKind,
  setSlideLayoutRel,
} from "./layout-helpers.js";
import {
  buildDiff,
  evolveSnapshot,
  findSlide,
  makeError,
  maxCNvPrId,
} from "./helpers.js";
import type { SetSlideLayoutPayload } from "./payloads.js";

export const setSlideLayoutHandler: CommandHandler<SetSlideLayoutPayload, PptxSnapshot> = {
  type: "pptx:set-slide-layout",
  apply(snapshot, payload, ctx) {
    if (!payload.layoutPartPath && !payload.layoutKind) {
      throw makeError("invalid-payload", "either layoutPartPath or layoutKind must be supplied");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);

    let layout: SlideLayout;
    let added: ReturnType<typeof applyAddedLayout> | null = null;
    let layoutPartPath: string;
    if (payload.layoutPartPath) {
      const existing = snapshot.root.layouts.get(payload.layoutPartPath);
      if (!existing) {
        throw makeError(
          "unknown-target",
          `layout part ${payload.layoutPartPath} not found in presentation.layouts`
        );
      }
      layout = existing;
      layoutPartPath = payload.layoutPartPath;
    } else {
      const resolved = resolveLayoutForKind(snapshot, payload.layoutKind!);
      layout = resolved.layout;
      layoutPartPath = resolved.layout.partPath;
      if (resolved.added) {
        added = applyAddedLayout(snapshot, resolved.added, layout);
      }
    }

    const clonePlaceholders = payload.clonePlaceholders !== false;
    const baseCNvPrId = maxCNvPrId(slide.shapes);
    const updatedSlide = {
      ...(clonePlaceholders ? clonePlaceholdersIntoSlide(slide, layout, ctx, baseCNvPrId) : slide),
      layoutPartPath,
    };

    // Slide → layout rel (re-pointed if it changed).
    const { relationships: relsAfterSlide, relsPath: slideRelsPath } = setSlideLayoutRel(
      added
        ? { ...snapshot, relationships: added.relationships }
        : snapshot,
      slide,
      layoutPartPath
    );

    const root: PptxPresentation = {
      ...snapshot.root,
      layouts: added ? added.layouts : snapshot.root.layouts,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? updatedSlide : s)),
    };

    const dirtyRels = added
      ? [...added.dirtyRels, slideRelsPath]
      : [slideRelsPath];

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        slides: [slide.partPath],
        layouts: added ? added.dirtyLayouts : [],
        relationships: dirtyRels,
        contentTypes: added?.dirtyContentTypes ?? false,
      },
      {
        relationships: relsAfterSlide,
        contentTypes: added?.contentTypes ?? snapshot.contentTypes,
      }
    );

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: slide.id,
        path: ["slides", sIdx],
        field: "layout",
        summary: `layout:${layout.kind}`,
      }),
    };
  },
};
